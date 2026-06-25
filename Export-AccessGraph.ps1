[CmdletBinding()]
param(
    [string]$DatabasePath,

    [string]$OutDir,

    [ValidateSet('None', 'ReferencedOnly', 'AllTableFields')]
    [string]$FieldNodeMode = 'ReferencedOnly',

    [switch]$DisableCodeHeuristics,

    [switch]$DisableMacroHeuristics,

    [switch]$NestInNamedFolder,

    [switch]$SkipViewerCopy
)

Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'

$AcObjectType = @{
    Table  = 0
    Query  = 1
    Form   = 2
    Report = 3
    Macro  = 4
    Module = 5
}

$script:EdgeId = 0
$script:SqlNodeCache = @{}
$script:NodeIndex = @{}
$script:EdgeIndex = @{}
$script:NameTargets = @{}
$script:DataNameTargets = @{}
$script:Warnings = New-Object 'System.Collections.Generic.List[object]'
$script:Nodes = New-Object 'System.Collections.Generic.List[object]'
$script:Edges = New-Object 'System.Collections.Generic.List[object]'
$script:KnownTableFields = @{}
$script:ProcIndex = @{}           # proc_name_lower → [List[string]] of module node IDs
$script:ProcCallRe = $null        # compiled regex for bare proc calls (built by Build-ProcIndex)
$script:ModuleCodeCache = @{}     # module_name → code text (read once, reused)

# DAO field type number → friendly name (for field node metadata)
$script:DAO_FIELD_TYPE = @{
    1 = 'Boolean';  2 = 'Byte';     3 = 'Integer';  4 = 'Long'
    5 = 'Currency'; 6 = 'Single';   7 = 'Double';   8 = 'Date/Time'
    10 = 'Text';    11 = 'OLE';     12 = 'Memo';    15 = 'GUID'
    16 = 'BigInt';  20 = 'Decimal'
}

# Built-in VBA / Access function names — excluded from cross-module call detection.
$script:VBA_BUILTIN_NAMES = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@(
        'asc','ascw','chr','chrw','format','instr','instrb','instrrev','join',
        'lcase','left','len','lenb','ltrim','mid','replace','right','space',
        'split','str','strcomp','strconv','strreverse','trim','rtrim','ucase','val','string',
        'cbool','cbyte','ccur','cdate','cdbl','cdec','cint','clng','clnglng','clngptr','csng','cstr','cvar','cverr',
        'isarray','isdate','isempty','iserror','ismissing','isnull','isnumeric','isobject','typename','vartype',
        'abs','atn','cos','exp','fix','int','log','rnd','round','sgn','sin','sqr','tan',
        'date','dateadd','datediff','datepart','dateserial','datevalue','day','formatdatetime',
        'hour','minute','month','monthname','now','second','time','timeserial','timevalue','timer',
        'weekday','weekdayname','year',
        'inputbox','msgbox',
        'curdir','dir','eof','filecopy','filedatetime','filelen','freefile','getattr','loc','lof','setattr',
        'array','erase','filter','lbound','ubound',
        'appactivate','beep','command','doevents','environ','sendkeys','shell',
        'error',
        'callbyname','createobject','getobject',
        'deletesetting','getsetting','savesetting',
        'hex','oct',
        'choose','iif','nz','partition','qbcolor','randomize','rgb',
        'davg','dcount','dfirst','dlast','dlookup','dmax','dmin','dstdev','dstdevp','dsum','dvar','dvarp',
        'codedb','currentdb','currentuser','eval','guidfromstring','hyperlinkpart','stringfromguid','syscmd'
    ),
    [System.StringComparer]::OrdinalIgnoreCase
)

function Add-WarningEntry {
    param(
        [string]$Code,
        [string]$Message,
        [hashtable]$Meta = @{}
    )

    $entry = [pscustomobject][ordered]@{
        code    = $Code
        message = $Message
        meta    = $Meta
    }

    $script:Warnings.Add($entry)
    Write-Warning $Message
}

function Ensure-Directory {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        $null = New-Item -ItemType Directory -Force -Path $Path
    }
}

function Safe-FileName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return '_'
    }

    return ($Name -replace '[\\/:*?"<>|]', '_')
}

function Get-FileHashInfo {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        return @{ hash = $null; size = 0 }
    }

    $hash = Get-FileHash -LiteralPath $Path -Algorithm SHA256
    $item = Get-Item -LiteralPath $Path
    return @{ hash = $hash.Hash.ToLowerInvariant(); size = $item.Length }
}

function Get-TextHash {
    param([string]$Text)

    $sha = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes([string]$Text)
        $hash = $sha.ComputeHash($bytes)
        return ([System.BitConverter]::ToString($hash) -replace '-', '').ToLowerInvariant()
    }
    finally {
        $sha.Dispose()
    }
}

function Get-PreviewText {
    param(
        [string]$Text,
        [int]$MaxLength = 180
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return ''
    }

    $flat = ($Text -replace '\s+', ' ').Trim()
    if ($flat.Length -le $MaxLength) {
        return $flat
    }

    return $flat.Substring(0, $MaxLength).TrimEnd() + '...'
}

function Format-MetaTitle {
    param([System.Collections.IDictionary]$Meta)

    if (-not $Meta -or $Meta.Count -eq 0) {
        return ''
    }

    $parts = New-Object 'System.Collections.Generic.List[string]'
    foreach ($key in ($Meta.Keys | Sort-Object)) {
        $value = $Meta[$key]
        if ($null -eq $value) {
            continue
        }

        if ($value -is [System.Collections.IEnumerable] -and -not ($value -is [string])) {
            $joined = ($value | ForEach-Object { $_ }) -join ', '
            $parts.Add(("{0}: {1}" -f $key, $joined))
        }
        else {
            $parts.Add(("{0}: {1}" -f $key, $value))
        }
    }

    return ($parts -join "`n")
}

function Update-NodeTitle {
    param($Node)

    $Node.title = Format-MetaTitle -Meta $Node.meta
}

function Register-ObjectNameTarget {
    param(
        [string]$Name,
        [string]$NodeId,
        [string]$Group,
        [switch]$IsDataObject
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return
    }

    if (-not $script:NameTargets.ContainsKey($Name)) {
        $script:NameTargets[$Name] = New-Object 'System.Collections.Generic.List[object]'
    }

    $script:NameTargets[$Name].Add([pscustomobject]@{
        id    = $NodeId
        group = $Group
        name  = $Name
    })

    if ($IsDataObject) {
        if (-not $script:DataNameTargets.ContainsKey($Name)) {
            $script:DataNameTargets[$Name] = New-Object 'System.Collections.Generic.List[object]'
        }

        $script:DataNameTargets[$Name].Add([pscustomobject]@{
            id    = $NodeId
            group = $Group
            name  = $Name
        })
    }
}

function Add-Node {
    param(
        [string]$Id,
        [string]$Label,
        [string]$Group,
        [hashtable]$Meta = @{}
    )

    if ($script:NodeIndex.ContainsKey($Id)) {
        $existing = $script:NodeIndex[$Id]
        foreach ($key in $Meta.Keys) {
            $existing.meta[$key] = $Meta[$key]
        }
        if ($Label) {
            $existing.label = $Label
        }
        if ($Group) {
            $existing.group = $Group
        }
        Update-NodeTitle -Node $existing
        return $existing
    }

    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) {
        $metaCopy[$key] = $Meta[$key]
    }

    $node = [pscustomobject][ordered]@{
        id    = $Id
        label = $Label
        group = $Group
        title = ''
        meta  = $metaCopy
    }

    Update-NodeTitle -Node $node
    $script:Nodes.Add($node)
    $script:NodeIndex[$Id] = $node
    return $node
}

function Add-Edge {
    param(
        [string]$From,
        [string]$To,
        [string]$Label,
        [string]$Kind,
        [string]$Arrows = 'to',
        [hashtable]$Meta = @{}
    )

    $metaJson = ConvertTo-Json -InputObject $Meta -Depth 8 -Compress
    $edgeKey = ($From, $To, $Kind, $Label, $Arrows, $metaJson) -join '|'
    if ($script:EdgeIndex.ContainsKey($edgeKey)) {
        return
    }

    $script:EdgeId += 1
    $metaCopy = [ordered]@{}
    foreach ($key in $Meta.Keys) {
        $metaCopy[$key] = $Meta[$key]
    }

    $edge = [pscustomobject][ordered]@{
        id     = ('e{0}' -f $script:EdgeId)
        from   = $From
        to     = $To
        label  = $Label
        kind   = $Kind
        arrows = $Arrows
        title  = Format-MetaTitle -Meta $metaCopy
        meta   = $metaCopy
    }

    $script:Edges.Add($edge)
    $script:EdgeIndex[$edgeKey] = $edge
}

function Save-TextObject {
    param(
        $AccessApp,
        [int]$Type,
        [string]$Name,
        [string]$Folder
    )

    Ensure-Directory -Path $Folder
    $path = Join-Path $Folder ((Safe-FileName -Name $Name) + '.txt')

    try {
        $AccessApp.SaveAsText($Type, $Name, $path)
        $hashInfo = Get-FileHashInfo -Path $path
        return [pscustomobject]@{
            path = $path
            hash = $hashInfo.hash
            size = $hashInfo.size
        }
    }
    catch {
        $errMsg = $_.Exception.Message
        $objGroup = switch ($Type) { 0 { 'table' } 1 { 'query' } 2 { 'form' } 3 { 'report' } 4 { 'macro' } 5 { 'module' } default { '' } }
        Add-WarningEntry -Code 'SaveAsTextFailed' -Message ("SaveAsText failed for object '{0}': {1}" -f $Name, $errMsg) -Meta @{ owner = $Name; group = $objGroup; name = $Name; type = $Type; folder = $Folder }
        return [pscustomobject]@{
            path = $null
            hash = $null
            size = 0
        }
    }
}

function Convert-AccessLiteral {
    param([string]$RawValue)

    if ($null -eq $RawValue) {
        return $null
    }

    $value = $RawValue.Trim()
    if ($value -eq 'Null') {
        return $null
    }

    if ($value.Length -ge 2 -and $value.StartsWith('"') -and $value.EndsWith('"')) {
        return ($value.Substring(1, $value.Length - 2) -replace '""', '"')
    }

    return $value
}

function Remove-AccessBrackets {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $Name
    }

    $trimmed = $Name.Trim()
    if ($trimmed.StartsWith('[') -and $trimmed.EndsWith(']') -and $trimmed.Length -ge 2) {
        return $trimmed.Substring(1, $trimmed.Length - 2)
    }

    return $trimmed
}

function Get-ObjectId {
    param(
        [string]$Group,
        [string]$Name
    )

    return ('{0}:{1}' -f $Group, $Name)
}

function Is-SystemOrTemporaryName {
    param([string]$Name)

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return $true
    }

    return ($Name -like 'MSys*' -or $Name -like '~*')
}

function Is-LikelySql {
    param([string]$Text)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return $false
    }

    $trimmed = $Text.Trim()
    return ($trimmed -match '^(?is)\s*(SELECT|INSERT|UPDATE|DELETE|TRANSFORM|PARAMETERS|WITH)\b')
}

function Get-AccessExportParse {
    param([string]$Path)

    $lines = Get-Content -LiteralPath $Path

    $root = [pscustomobject]@{
        Type       = 'Root'
        Properties = @{}
        Children   = New-Object 'System.Collections.Generic.List[object]'
    }

    $stack = New-Object System.Collections.Stack
    $stack.Push($root)

    $inBlob = $false
    $inCode = $false
    $codeLines = New-Object 'System.Collections.Generic.List[string]'

    foreach ($line in $lines) {
        if ($inCode) {
            $codeLines.Add($line)
            continue
        }

        if ($inBlob) {
            if ($line.Trim() -eq 'End') {
                $inBlob = $false
            }
            continue
        }

        if ($line -match '^\s*CodeBehind(Form|Report)\b') {
            $inCode = $true
            $codeLines.Add($line)
            continue
        }

        if ($line -match '^\s*Begin\s+(.+?)\s*$') {
            $typeName = $matches[1].Trim()
            $block = [pscustomobject]@{
                Type       = $typeName
                Properties = @{}
                Children   = New-Object 'System.Collections.Generic.List[object]'
            }

            $stack.Peek().Children.Add($block)
            $stack.Push($block)
            continue
        }

        # Anonymous Begin blocks (controls container / defaults block)
        if ($line -match '^\s*Begin\s*$') {
            $block = [pscustomobject]@{
                Type       = '_anonymous'
                Properties = @{}
                Children   = New-Object 'System.Collections.Generic.List[object]'
            }

            $stack.Peek().Children.Add($block)
            $stack.Push($block)
            continue
        }

        if ($line -match '^\s*End\s*$') {
            if ($stack.Count -gt 1) {
                [void]$stack.Pop()
            }
            continue
        }

        if ($line -match '^\s*([A-Za-z][A-Za-z0-9_]*)\s*=\s*(.+?)\s*$') {
            $propName = $matches[1]
            $rawValue = $matches[2]

            if ($rawValue.Trim() -eq 'Begin') {
                $stack.Peek().Properties[$propName] = '[BLOB]'
                $inBlob = $true
                continue
            }

            $stack.Peek().Properties[$propName] = Convert-AccessLiteral -RawValue $rawValue
            continue
        }
    }

    $designRoot = $null
    foreach ($child in $root.Children) {
        if ($child.Type -match '^(Form|Report)$') {
            $designRoot = $child
            break
        }
    }

    return [pscustomobject]@{
        Root       = $root
        DesignRoot = $designRoot
        Code       = ($codeLines -join [Environment]::NewLine)
        RawLines   = $lines
    }
}

function Get-FlattenedBlocks {
    param($Block)

    $results = New-Object 'System.Collections.Generic.List[object]'

    if ($null -ne $Block) {
        # Use a stack-based traversal instead of nested recursive functions
        # to avoid PowerShell parameter-binding issues with List[object]
        $stack = New-Object System.Collections.Stack
        $stack.Push($Block)
        while ($stack.Count -gt 0) {
            $current = $stack.Pop()
            if ($null -eq $current.Children) { continue }
            # Process children in reverse so they appear in original order
            $childCount = $current.Children.Count
            for ($i = $childCount - 1; $i -ge 0; $i--) {
                $child = $current.Children[$i]
                $results.Add($child)
                $stack.Push($child)
            }
        }
    }

    return $results.ToArray()
}

function Get-ControlBlocks {
    param($DesignRoot)

    $knownControlTypes = @(
        'textbox',
        'combobox',
        'listbox',
        'checkbox',
        'optionbutton',
        'togglebutton',
        'boundobjectframe',
        'attachment',
        'subform',
        'subreport',
        'customcontrol'
    )

    $blocks = Get-FlattenedBlocks -Block $DesignRoot
    return @(
        $blocks | Where-Object {
            $typeName = ([string]$_.Type).ToLowerInvariant()
            $_.Properties.ContainsKey('ControlSource') -or
            $_.Properties.ContainsKey('SourceObject') -or
            ($knownControlTypes -contains $typeName)
        }
    )
}

function Find-ReferencedDataNames {
    param(
        [string]$Text,
        [string[]]$KnownNames
    )

    $hits = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @()
    }

    foreach ($name in $KnownNames) {
        if ([string]::IsNullOrWhiteSpace($name)) {
            continue
        }

        $escaped = [regex]::Escape($name)
        $pattern = "(?is)(?<![\w])(?:\[$escaped\]|$escaped)(?![\w])"
        if ([regex]::IsMatch($Text, $pattern)) {
            [void]$hits.Add($name)
        }
    }

    return [string[]]$hits
}

function Get-TargetsByName {
    param(
        [string]$Name,
        [hashtable]$TargetTable
    )

    if ([string]::IsNullOrWhiteSpace($Name)) {
        return @()
    }

    if ($TargetTable.ContainsKey($Name)) {
        return $TargetTable[$Name].ToArray()
    }

    $unbracketed = Remove-AccessBrackets -Name $Name
    if ($TargetTable.ContainsKey($unbracketed)) {
        return $TargetTable[$unbracketed].ToArray()
    }

    return @()
}

function Ensure-SqlNode {
    param(
        [string]$SqlText,
        [string]$Origin,
        [string]$SqlFolder
    )

    $hash = Get-TextHash -Text $SqlText
    if ($script:SqlNodeCache.ContainsKey($hash)) {
        return $script:SqlNodeCache[$hash]
    }

    Ensure-Directory -Path $SqlFolder
    $path = Join-Path $SqlFolder ($hash + '.sql')
    if (-not (Test-Path -LiteralPath $path)) {
        Set-Content -LiteralPath $path -Value $SqlText -Encoding UTF8
    }

    $nodeId = 'sql:' + $hash.Substring(0, 20)
    $node = Add-Node -Id $nodeId -Label ('SQL ' + $hash.Substring(0, 8)) -Group 'sql' -Meta @{
        origin    = $Origin
        sqlHash   = $hash
        sqlPath   = $path
        sqlLength = $SqlText.Length
        preview   = Get-PreviewText -Text $SqlText
    }

    $script:SqlNodeCache[$hash] = $node
    return $node
}

function Add-SqlReferenceEdges {
    param(
        [string]$SqlText,
        [string]$FromNodeId,
        [string]$RelationKind,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    foreach ($name in (Find-ReferencedDataNames -Text $SqlText -KnownNames $KnownDataNames)) {
        foreach ($target in (Get-TargetsByName -Name $name -TargetTable $script:DataNameTargets)) {
            Add-Edge -From $FromNodeId -To $target.id -Label 'uses' -Kind $RelationKind -Arrows 'to' -Meta @{ name = $name }
        }
    }
}

function Get-FieldReferenceFromControlSource {
    param([string]$ControlSource)

    if ([string]::IsNullOrWhiteSpace($ControlSource)) {
        return $null
    }

    $trimmed = $ControlSource.Trim()
    if ($trimmed.StartsWith('=')) {
        return $null
    }

    if ($trimmed -match '[\+\-\*\/\&\(\)]') {
        return $null
    }

    $parts = $trimmed -split '\.'
    $candidate = $parts[$parts.Count - 1].Trim()
    $candidate = Remove-AccessBrackets -Name $candidate

    if ([string]::IsNullOrWhiteSpace($candidate)) {
        return $null
    }

    return $candidate
}

function Ensure-FieldNode {
    param(
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$FieldName,
        [bool]$Verified,
        [string]$DataType = $null
    )

    if ($FieldNodeMode -eq 'None') {
        return $null
    }

    $nodeId = ('field:{0}:{1}:{2}' -f $OwnerGroup, $OwnerName, $FieldName)
    $node = Add-Node -Id $nodeId -Label $FieldName -Group 'field' -Meta @{
        ownerId    = $OwnerNodeId
        ownerGroup = $OwnerGroup
        ownerName  = $OwnerName
        fieldName  = $FieldName
        verified   = $Verified
        dataType   = $DataType
    }

    Add-Edge -From $OwnerNodeId -To $nodeId -Label 'field' -Kind 'field-owner' -Arrows 'to' -Meta @{ owner = $OwnerName; field = $FieldName }
    return $node
}

function Ensure-TableFieldSet {
    param([string]$TableName)

    if (-not $script:KnownTableFields.ContainsKey($TableName)) {
        $script:KnownTableFields[$TableName] = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    }
}

function Get-ResolvedDataSourceInfo {
    param(
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$RecordSource,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    $result = [pscustomobject]@{
        mode      = 'none'
        raw       = $RecordSource
        targetIds = @()
        targetRef = $null
        sqlNodeId = $null
    }

    if ([string]::IsNullOrWhiteSpace($RecordSource)) {
        return $result
    }

    $targets = @(Get-TargetsByName -Name $RecordSource -TargetTable $script:DataNameTargets)
    if ($targets.Count -gt 0) {
        foreach ($target in $targets) {
            Add-Edge -From $OwnerNodeId -To $target.id -Label 'RecordSource' -Kind 'recordsource' -Arrows 'to' -Meta @{ recordSource = $RecordSource }
        }

        $result.mode = 'named'
        $result.targetIds = @($targets | ForEach-Object { $_.id })
        if ($targets.Count -eq 1) {
            $result.targetRef = $targets[0]
        }
        return $result
    }

    if (Is-LikelySql -Text $RecordSource) {
        $sqlNode = Ensure-SqlNode -SqlText $RecordSource -Origin ("{0}:{1}:RecordSource" -f $OwnerGroup, $OwnerName) -SqlFolder $SqlFolder
        Add-Edge -From $OwnerNodeId -To $sqlNode.id -Label 'RecordSource' -Kind 'recordsource-sql' -Arrows 'to' -Meta @{ preview = Get-PreviewText -Text $RecordSource }
        Add-SqlReferenceEdges -SqlText $RecordSource -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames

        $result.mode = 'sql'
        $result.targetIds = @($sqlNode.id)
        $result.sqlNodeId = $sqlNode.id
        $result.targetRef = [pscustomobject]@{ id = $sqlNode.id; group = 'sql'; name = $sqlNode.label }
        return $result
    }

    Add-WarningEntry -Code 'UnresolvedRecordSource' -Message ("Could not resolve RecordSource '{0}' on {1} '{2}'." -f $RecordSource, $OwnerGroup, $OwnerName) -Meta @{ owner = $OwnerName; group = $OwnerGroup; recordSource = $RecordSource }
    $result.mode = 'unresolved'
    return $result
}

function Resolve-SourceObjectTarget {
    param([string]$SourceObject)

    if ([string]::IsNullOrWhiteSpace($SourceObject)) {
        return $null
    }

    $trimmed = $SourceObject.Trim()
    if ($trimmed -match '^(?i)(Form|Report)\.(.+)$') {
        $targetGroup = $matches[1].ToLowerInvariant()
        $targetName = $matches[2]
        return [pscustomobject]@{
            group = $targetGroup
            name  = $targetName
            id    = Get-ObjectId -Group $targetGroup -Name $targetName
        }
    }

    return $null
}

function Add-FormOrReportEdgesFromExport {
    param(
        [string]$ObjectGroup,
        [string]$ObjectName,
        [string]$RawPath,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    if ([string]::IsNullOrWhiteSpace($RawPath) -or -not (Test-Path -LiteralPath $RawPath)) {
        return
    }

    $objectId = Get-ObjectId -Group $ObjectGroup -Name $ObjectName
    $parse = Get-AccessExportParse -Path $RawPath
    $designRoot = $parse.DesignRoot
    if ($null -eq $designRoot) {
        Add-WarningEntry -Code 'DesignRootMissing' -Message ("No design root found while parsing {0} '{1}'." -f $ObjectGroup, $ObjectName) -Meta @{ owner = $ObjectName; group = $ObjectGroup; path = $RawPath }
        return
    }

    $recordSource = $null
    if ($designRoot.Properties.ContainsKey('RecordSource')) {
        $recordSource = [string]$designRoot.Properties['RecordSource']
    }

    $resolvedRecordSource = Get-ResolvedDataSourceInfo -OwnerNodeId $objectId -OwnerGroup $ObjectGroup -OwnerName $ObjectName -RecordSource $recordSource -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames

    $controlBlocks = Get-ControlBlocks -DesignRoot $designRoot
    foreach ($control in $controlBlocks) {
        $controlType = [string]$control.Type
        $controlName = if ($control.Properties.ContainsKey('Name')) { [string]$control.Properties['Name'] } else { '' }

        if ($control.Properties.ContainsKey('SourceObject')) {
            $sourceObject = [string]$control.Properties['SourceObject']
            $target = Resolve-SourceObjectTarget -SourceObject $sourceObject
            if ($null -ne $target) {
                $linkMasterFields = if ($control.Properties.ContainsKey('LinkMasterFields')) { [string]$control.Properties['LinkMasterFields'] } else { $null }
                $linkChildFields = if ($control.Properties.ContainsKey('LinkChildFields')) { [string]$control.Properties['LinkChildFields'] } else { $null }

                Add-Edge -From $objectId -To $target.id -Label 'SourceObject' -Kind 'sourceobject' -Arrows 'to' -Meta @{
                    controlName      = $controlName
                    controlType      = $controlType
                    sourceObject     = $sourceObject
                    linkMasterFields = $linkMasterFields
                    linkChildFields  = $linkChildFields
                }
            }
        }

        if ($control.Properties.ContainsKey('ControlSource')) {
            $controlSource = [string]$control.Properties['ControlSource']
            $fieldName = Get-FieldReferenceFromControlSource -ControlSource $controlSource

            if ($fieldName -and $null -ne $resolvedRecordSource.targetRef) {
                $ownerRef = $resolvedRecordSource.targetRef
                $verified = $false
                $dataType = $null

                if ($ownerRef.group -eq 'table') {
                    if ($script:KnownTableFields.ContainsKey($ownerRef.name) -and $script:KnownTableFields[$ownerRef.name].Contains($fieldName)) {
                        $verified = $true
                    }
                }

                $fieldNode = Ensure-FieldNode -OwnerNodeId $ownerRef.id -OwnerGroup $ownerRef.group -OwnerName $ownerRef.name -FieldName $fieldName -Verified $verified -DataType $dataType
                if ($null -ne $fieldNode) {
                    Add-Edge -From $objectId -To $fieldNode.id -Label 'ControlSource' -Kind 'controlsource' -Arrows 'to' -Meta @{
                        controlName   = $controlName
                        controlType   = $controlType
                        controlSource = $controlSource
                    }
                }
            }
            elseif ($null -ne $resolvedRecordSource.targetRef) {
                Add-Edge -From $objectId -To $resolvedRecordSource.targetRef.id -Label 'ControlExpr' -Kind 'control-expression' -Arrows 'to' -Meta @{
                    controlName   = $controlName
                    controlType   = $controlType
                    controlSource = $controlSource
                }
            }
        }

        # RowSource (ComboBox / ListBox data binding)
        if ($control.Properties.ContainsKey('RowSource')) {
            $rsValue = [string]$control.Properties['RowSource']
            if (-not [string]::IsNullOrWhiteSpace($rsValue)) {
                $rsTargets = @(Get-TargetsByName -Name $rsValue -TargetTable $script:DataNameTargets)
                if ($rsTargets.Count -gt 0) {
                    Add-Edge -From $objectId -To $rsTargets[0].id -Label 'RowSource' -Kind 'rowsource' -Arrows 'to' -Meta @{
                        controlName = $controlName
                        controlType = $controlType
                        rowSource   = $rsValue
                    }
                }
                elseif (Is-LikelySql -Text $rsValue) {
                    $rsSqlNode = Ensure-SqlNode -SqlText $rsValue -Origin ("RowSource:{0}" -f $controlName) -SqlFolder $SqlFolder
                    Add-Edge -From $objectId -To $rsSqlNode.id -Label 'RowSource' -Kind 'rowsource' -Arrows 'to' -Meta @{
                        controlName = $controlName
                        controlType = $controlType
                    }
                    Add-SqlReferenceEdges -SqlText $rsValue -FromNodeId $rsSqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
                }
            }
        }
    }

    if (-not $DisableCodeHeuristics) {
        Add-CodeHeuristicEdges -OwnerNodeId $objectId -OwnerGroup $ObjectGroup -OwnerName $ObjectName -Text $parse.Code -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
    }
}

function Build-ProcIndex {
    <#
    .SYNOPSIS
        Index all public VBA procedures across standalone modules for cross-module
        call detection. Populates $script:ProcIndex and $script:ProcCallRe.
    #>

    $procDeclRe = '(?im)^\s*(?:Public\s+)?(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'
    $privateProcRe = '(?im)^\s*Private\s+(?:Sub|Function|Property\s+(?:Get|Let|Set))\s+(\w+)'

    foreach ($node in $script:NodeIndex.Values) {
        if ($node.group -ne 'module') { continue }
        $rawPath = $node.meta.rawPath
        if (-not $rawPath -or -not (Test-Path -LiteralPath $rawPath)) { continue }

        $code = Get-Content -LiteralPath $rawPath -Raw
        if ([string]::IsNullOrWhiteSpace($code)) { continue }
        $script:ModuleCodeCache[$node.label] = $code

        # Collect private procedure names to exclude
        $privateNames = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($m in [regex]::Matches($code, $privateProcRe)) {
            [void]$privateNames.Add($m.Groups[1].Value)
        }

        # Index all public procedures
        $nodeId = $node.id
        foreach ($m in [regex]::Matches($code, $procDeclRe)) {
            $procName = $m.Groups[1].Value
            $pnameLower = $procName.ToLowerInvariant()
            if ($privateNames.Contains($procName)) { continue }
            if ($script:VBA_BUILTIN_NAMES.Contains($pnameLower)) { continue }
            if ($pnameLower.Length -lt 2) { continue }

            if (-not $script:ProcIndex.ContainsKey($pnameLower)) {
                $script:ProcIndex[$pnameLower] = New-Object 'System.Collections.Generic.List[string]'
            }
            $script:ProcIndex[$pnameLower].Add($nodeId)
        }
    }

    # Build compiled regex for call detection
    $procNames = @($script:ProcIndex.Keys | Sort-Object { $_.Length } -Descending)
    if ($procNames.Count -gt 0) {
        $escaped = $procNames | ForEach-Object { [regex]::Escape($_) }
        $alt = $escaped -join '|'
        # Match bare calls: ProcName( — but NOT object.ProcName(
        # Also match: Call ProcName
        $script:ProcCallRe = [regex]::new(
            "(?<![\.\w])(?:$alt)\s*\(|\bCall\s+(?:$alt)\b",
            [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
        )
    }
}

function Add-CodeHeuristicEdges {
    param(
        [string]$OwnerNodeId,
        [string]$OwnerGroup,
        [string]$OwnerName,
        [string]$Text,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    if ([string]::IsNullOrWhiteSpace($Text)) {
        return
    }

    $patterns = @(
        @{ regex = '(?is)\bDoCmd\.OpenForm\s+"(?<name>(?:[^"]|"")+)"'; group = 'form'; label = 'OpenForm'; kind = 'vba-openform' },
        @{ regex = '(?is)\bDoCmd\.OpenReport\s+"(?<name>(?:[^"]|"")+)"'; group = 'report'; label = 'OpenReport'; kind = 'vba-openreport' },
        @{ regex = '(?is)\bDoCmd\.OpenQuery\s+"(?<name>(?:[^"]|"")+)"'; group = 'query'; label = 'OpenQuery'; kind = 'vba-openquery' },
        @{ regex = '(?is)\bDoCmd\.OpenTable\s+"(?<name>(?:[^"]|"")+)"'; group = 'table'; label = 'OpenTable'; kind = 'vba-opentable' },
        @{ regex = '(?is)\bCurrentDb\s*\(\s*\)\s*\.\s*QueryDefs\s*\(\s*"(?<name>(?:[^"]|"")+)"\s*\)'; group = 'query'; label = 'QueryDefs'; kind = 'vba-querydefs' },
        @{ regex = '(?is)\bDBEngine\s*\(\s*0\s*\)\s*\(\s*0\s*\)\s*\.\s*QueryDefs\s*\(\s*"(?<name>(?:[^"]|"")+)"\s*\)'; group = 'query'; label = 'QueryDefs'; kind = 'vba-querydefs' },
        @{ regex = '(?is)\bDoCmd\.RunMacro\s+"(?<name>(?:[^"]|"")+)"'; group = 'macro'; label = 'RunMacro'; kind = 'vba-runmacro' }
    )

    foreach ($pattern in $patterns) {
        $matches = [regex]::Matches($Text, $pattern.regex)
        foreach ($match in $matches) {
            $name = ($match.Groups['name'].Value -replace '""', '"')
            if ([string]::IsNullOrWhiteSpace($name)) {
                continue
            }

            $targetId = Get-ObjectId -Group $pattern.group -Name $name
            if ($script:NodeIndex.ContainsKey($targetId)) {
                Add-Edge -From $OwnerNodeId -To $targetId -Label $pattern.label -Kind $pattern.kind -Arrows 'to' -Meta @{ name = $name }
            }
        }
    }

    $sqlMatches = [regex]::Matches($Text, '(?is)\bDoCmd\.RunSQL\s+"(?<sql>(?:[^"]|"")+)"')
    foreach ($match in $sqlMatches) {
        $sqlText = ($match.Groups['sql'].Value -replace '""', '"')
        if ([string]::IsNullOrWhiteSpace($sqlText)) {
            continue
        }

        $sqlNode = Ensure-SqlNode -SqlText $sqlText -Origin ("{0}:{1}:VBA" -f $OwnerGroup, $OwnerName) -SqlFolder $SqlFolder
        Add-Edge -From $OwnerNodeId -To $sqlNode.id -Label 'RunSQL' -Kind 'vba-runsql' -Arrows 'to' -Meta @{ preview = Get-PreviewText -Text $sqlText }
        Add-SqlReferenceEdges -SqlText $sqlText -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
    }

    # VBA SourceObject assignment: Me.sfrmChild.SourceObject = "Form.frmName" or "sfrmName"
    $soMatches = [regex]::Matches($Text, '(?im)\.SourceObject\s*=\s*"(?<name>(?:[^"]|"")+)"')
    foreach ($match in $soMatches) {
        $soValue = ($match.Groups['name'].Value -replace '""', '"').Trim()
        if ([string]::IsNullOrWhiteSpace($soValue)) { continue }

        $targetId = $null
        if ($soValue -match '^(?i)(Form|Report)\.(.+)$') {
            $targetId = Get-ObjectId -Group ($matches[1].ToLowerInvariant()) -Name $matches[2]
        }
        else {
            # Bare name — try form first, then report
            $tryForm = Get-ObjectId -Group 'form' -Name $soValue
            $tryReport = Get-ObjectId -Group 'report' -Name $soValue
            if ($script:NodeIndex.ContainsKey($tryForm)) { $targetId = $tryForm }
            elseif ($script:NodeIndex.ContainsKey($tryReport)) { $targetId = $tryReport }
        }

        if ($targetId -and $script:NodeIndex.ContainsKey($targetId)) {
            Add-Edge -From $OwnerNodeId -To $targetId -Label 'SourceObject' -Kind 'vba-sourceobject' -Arrows 'to' -Meta @{ sourceObject = $soValue }
        }
    }

    # VBA type-dependency edges: Dim/As, New, qualified member access to known modules/classes
    $seenTypeEdges = @{}
    foreach ($targetName in $script:NameTargets.Keys) {
        foreach ($target in $script:NameTargets[$targetName]) {
            if ($target.group -ne 'module') { continue }
            if ($target.id -eq $OwnerNodeId) { continue }

            $escaped = [regex]::Escape($targetName)
            $typePattern = "(?im)(?:\bAs\s+$escaped\b|\bNew\s+$escaped\b|\b$escaped\s*\.)"
            if ([regex]::IsMatch($Text, $typePattern)) {
                $edgeKey = "$OwnerNodeId->$($target.id)"
                if (-not $seenTypeEdges.ContainsKey($edgeKey)) {
                    $seenTypeEdges[$edgeKey] = $true
                    Add-Edge -From $OwnerNodeId -To $target.id -Label 'uses type' -Kind 'vba-type-ref' -Arrows 'to' -Meta @{ name = $targetName }
                }
            }
        }
    }

    # VBA data-reference edges: scan string literals for table/query names
    # Catches OpenRecordset("tbl"), Execute "INSERT INTO tbl", DLookup("f","tbl",...), SQL strings, etc.
    if ($KnownDataNames.Count -gt 0) {
        $literalMatches = [regex]::Matches($Text, '"((?:[^"]|"")*)"')
        if ($literalMatches.Count -gt 0) {
            $literalText = ($literalMatches | ForEach-Object { $_.Groups[1].Value -replace '""', '"' }) -join ' '
            $seenDataEdges = @{}
            foreach ($name in (Find-ReferencedDataNames -Text $literalText -KnownNames $KnownDataNames)) {
                foreach ($target in (Get-TargetsByName -Name $name -TargetTable $script:DataNameTargets)) {
                    $edgeKey = "$OwnerNodeId->$($target.id)"
                    if (-not $seenDataEdges.ContainsKey($edgeKey)) {
                        $seenDataEdges[$edgeKey] = $true
                        Add-Edge -From $OwnerNodeId -To $target.id -Label 'uses data' -Kind 'vba-data-ref' -Arrows 'to' -Meta @{ name = $name }
                    }
                }
            }
        }
    }

    # Cross-module procedure calls (bare FuncName( or Call SubName)
    if ($null -ne $script:ProcCallRe) {
        $seenCallEdges = @{}
        foreach ($match in $script:ProcCallRe.Matches($Text)) {
            $matched = $match.Value
            # Strip leading 'Call ' if present, trailing '(' or whitespace
            $procName = ($matched -replace '(?i)^\s*Call\s+', '').TrimEnd('( ')
            $pnameLower = $procName.ToLowerInvariant()
            $targetIds = $script:ProcIndex[$pnameLower]
            if ($null -eq $targetIds) { continue }
            foreach ($tid in $targetIds) {
                if ($tid -eq $OwnerNodeId) { continue }  # skip self-edges
                $edgeKey = "$OwnerNodeId->$($tid):call:$pnameLower"
                if (-not $seenCallEdges.ContainsKey($edgeKey)) {
                    $seenCallEdges[$edgeKey] = $true
                    Add-Edge -From $OwnerNodeId -To $tid -Label 'calls' -Kind 'vba-call' -Arrows 'to' -Meta @{ procedure = $procName }
                }
            }
        }
    }
}

function Add-MacroHeuristicEdges {
    param(
        [string]$MacroName,
        [string]$RawPath,
        [string]$SqlFolder,
        [string[]]$KnownDataNames
    )

    if ($DisableMacroHeuristics) {
        return
    }

    if ([string]::IsNullOrWhiteSpace($RawPath) -or -not (Test-Path -LiteralPath $RawPath)) {
        return
    }

    $macroId = Get-ObjectId -Group 'macro' -Name $MacroName
    $lines = Get-Content -LiteralPath $RawPath

    for ($i = 0; $i -lt $lines.Count; $i += 1) {
        $line = $lines[$i]
        if ($line -notmatch '^\s*Action\s*=\s*"?(?<action>[A-Za-z0-9_]+)"?\s*$') {
            continue
        }

        $action = $matches['action']
        $argValue = $null
        for ($j = $i + 1; $j -lt [Math]::Min($i + 8, $lines.Count); $j += 1) {
            if ($lines[$j] -match '^\s*Action\s*=') {
                break
            }
            if ($lines[$j] -match '^\s*Argument\s*=\s*(.+?)\s*$') {
                $argValue = Convert-AccessLiteral -RawValue $matches[1]
                break
            }
        }

        switch -Regex ($action) {
            '^OpenForm$' {
                if ($argValue) {
                    $targetId = Get-ObjectId -Group 'form' -Name $argValue
                    if ($script:NodeIndex.ContainsKey($targetId)) {
                        Add-Edge -From $macroId -To $targetId -Label 'OpenForm' -Kind 'macro-openform' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenReport$' {
                if ($argValue) {
                    $targetId = Get-ObjectId -Group 'report' -Name $argValue
                    if ($script:NodeIndex.ContainsKey($targetId)) {
                        Add-Edge -From $macroId -To $targetId -Label 'OpenReport' -Kind 'macro-openreport' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenQuery$' {
                if ($argValue) {
                    $targetId = Get-ObjectId -Group 'query' -Name $argValue
                    if ($script:NodeIndex.ContainsKey($targetId)) {
                        Add-Edge -From $macroId -To $targetId -Label 'OpenQuery' -Kind 'macro-openquery' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^OpenTable$' {
                if ($argValue) {
                    $targetId = Get-ObjectId -Group 'table' -Name $argValue
                    if ($script:NodeIndex.ContainsKey($targetId)) {
                        Add-Edge -From $macroId -To $targetId -Label 'OpenTable' -Kind 'macro-opentable' -Arrows 'to' -Meta @{ name = $argValue }
                    }
                }
            }
            '^RunSQL$' {
                if ($argValue) {
                    $sqlNode = Ensure-SqlNode -SqlText $argValue -Origin ("macro:{0}" -f $MacroName) -SqlFolder $SqlFolder
                    Add-Edge -From $macroId -To $sqlNode.id -Label 'RunSQL' -Kind 'macro-runsql' -Arrows 'to' -Meta @{ preview = Get-PreviewText -Text $argValue }
                    Add-SqlReferenceEdges -SqlText $argValue -FromNodeId $sqlNode.id -RelationKind 'sql-reference' -SqlFolder $SqlFolder -KnownDataNames $KnownDataNames
                }
            }
        }
    }
}

function Copy-ViewerIfPresent {
    param(
        [string]$DestinationFolder,
        [string]$GraphJson,
        [switch]$Disabled
    )

    if ($Disabled) {
        return
    }

    $viewerSource = Join-Path $PSScriptRoot 'access-graph-viewer.html'
    if (Test-Path -LiteralPath $viewerSource) {
        $html = Get-Content -LiteralPath $viewerSource -Raw
        $embedTag = "<script>var EMBEDDED_GRAPH = $GraphJson;</script>"
        $html = $html -replace '<!-- EMBED_GRAPH_DATA -->', $embedTag
        Set-Content -LiteralPath (Join-Path $DestinationFolder 'index.html') -Value $html -Encoding UTF8
    }
}

function Select-AccessDatabaseFile {
    param([string]$Title = 'Select an Access database to graph')

    $filter = 'Access Databases (*.accdb;*.mdb;*.accde;*.mde)|*.accdb;*.mdb;*.accde;*.mde|All files (*.*)|*.*'

    $pickerScript = {
        param($Filter, $Title)
        Add-Type -AssemblyName System.Windows.Forms
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Filter = $Filter
        $dialog.Title = $Title
        $dialog.Multiselect = $false
        $dialog.CheckFileExists = $true
        $dialog.RestoreDirectory = $true
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            $dialog.FileName
        }
    }

    # OpenFileDialog requires an STA thread; PowerShell 7 defaults to MTA.
    if ([System.Threading.Thread]::CurrentThread.GetApartmentState() -eq [System.Threading.ApartmentState]::STA) {
        return (& $pickerScript $filter $Title)
    }

    $runspace = [runspacefactory]::CreateRunspace()
    $runspace.ApartmentState = 'STA'
    $runspace.ThreadOptions = 'ReuseThread'
    $runspace.Open()
    try {
        $ps = [powershell]::Create()
        $ps.Runspace = $runspace
        [void]$ps.AddScript($pickerScript).AddArgument($filter).AddArgument($Title)
        $result = $ps.Invoke()
        $ps.Dispose()
        if ($result -and $result.Count -gt 0) {
            return [string]$result[0]
        }
        return $null
    }
    finally {
        $runspace.Close()
        $runspace.Dispose()
    }
}

if ([string]::IsNullOrWhiteSpace($DatabasePath)) {
    Write-Host 'No database path provided - opening file picker...'
    $DatabasePath = Select-AccessDatabaseFile
    if ([string]::IsNullOrWhiteSpace($DatabasePath)) {
        Write-Host 'No database selected. Exiting.'
        return
    }
    Write-Host ('Selected database: ' + $DatabasePath)
}

$resolvedDatabasePath = Resolve-Path -LiteralPath $DatabasePath
$databaseFullPath = $resolvedDatabasePath.Path

# When -OutDir is not specified, create an output folder next to the database.
# With -NestInNamedFolder, the output is placed inside a subfolder named after
# the database (e.g. example.accdb -> <db folder>\example\access-graph-out).
if ([string]::IsNullOrWhiteSpace($OutDir)) {
    $dbParent = Split-Path -Parent $databaseFullPath
    if ($NestInNamedFolder) {
        $dbName = [System.IO.Path]::GetFileNameWithoutExtension($databaseFullPath)
        $OutDir = Join-Path (Join-Path $dbParent $dbName) 'access-graph-out'
    }
    else {
        $OutDir = Join-Path $dbParent 'access-graph-out'
    }
}

# Resolve OutDir to absolute path so COM SaveAsText receives full paths
Ensure-Directory -Path $OutDir
$OutDir = (Resolve-Path -LiteralPath $OutDir).Path

$rawDir = Join-Path $OutDir 'raw'
$folders = @(
    $OutDir,
    $rawDir,
    (Join-Path $rawDir 'tables'),
    (Join-Path $rawDir 'queries'),
    (Join-Path $rawDir 'forms'),
    (Join-Path $rawDir 'reports'),
    (Join-Path $rawDir 'macros'),
    (Join-Path $rawDir 'modules'),
    (Join-Path $rawDir 'sql')
)

foreach ($folder in $folders) {
    Ensure-Directory -Path $folder
}

$access = $null
$db = $null

try {
    Write-Progress -Activity 'Access Graph Export' -Status 'Opening database...' -PercentComplete 0
    Write-Host ('Opening database: ' + $databaseFullPath)
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($databaseFullPath)
    $db = $access.CurrentDb()

    $tableNames = New-Object 'System.Collections.Generic.List[string]'
    Write-Progress -Activity 'Access Graph Export' -Status 'Scanning tables...' -PercentComplete 5
    $tableIdx = 0
    foreach ($tableDef in $db.TableDefs) {
        $tdName = $null
        try { $tdName = [string]$tableDef.Name } catch { continue }
        if ([string]::IsNullOrWhiteSpace($tdName) -or (Is-SystemOrTemporaryName -Name $tdName)) {
            continue
        }

        $tableName = $tdName
        $tableNames.Add($tableName)
        $tableIdx++
        Write-Progress -Activity 'Access Graph Export' -Status "Scanning table: $tableName" -PercentComplete 5

        Ensure-TableFieldSet -TableName $tableName
        $fieldSet = $script:KnownTableFields[$tableName]
        try {
            foreach ($field in $tableDef.Fields) {
                [void]$fieldSet.Add([string]$field.Name)
            }
        } catch {
            Add-WarningEntry -Code 'FieldEnumFailed' -Message ("Could not enumerate fields for table '{0}': {1}" -f $tableName, $_.Exception.Message) -Meta @{ owner = $tableName; group = 'table' }
        }

        # SaveAsText does not support acTable (type 0); skip export for tables
        $rawInfo = [pscustomobject]@{ path = $null; hash = $null; size = 0 }
        $tableNodeId = Get-ObjectId -Group 'table' -Name $tableName
        $tblConnect = try { [string]$tableDef.Connect } catch { '' }
        $tblSource  = try { [string]$tableDef.SourceTableName } catch { '' }
        $tblFieldCount = try { $tableDef.Fields.Count } catch { 0 }
        Add-Node -Id $tableNodeId -Label $tableName -Group 'table' -Meta @{
            connect     = $tblConnect
            sourceTable = $tblSource
            fieldCount  = $tblFieldCount
            rawPath     = $rawInfo.path
            rawHash     = $rawInfo.hash
            rawSize     = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $tableName -NodeId $tableNodeId -Group 'table' -IsDataObject

        if ($FieldNodeMode -eq 'AllTableFields') {
            foreach ($field in $tableDef.Fields) {
                $fieldType = try { [int]$field.Type } catch { -1 }
                $fieldTypeName = if ($script:DAO_FIELD_TYPE.ContainsKey($fieldType)) { $script:DAO_FIELD_TYPE[$fieldType] } else { "Type$fieldType" }
                Ensure-FieldNode -OwnerNodeId $tableNodeId -OwnerGroup 'table' -OwnerName $tableName -FieldName ([string]$field.Name) -Verified $true -DataType $fieldTypeName | Out-Null
            }
        }
    }

    Write-Progress -Activity 'Access Graph Export' -Status 'Scanning relationships...' -PercentComplete 10
    foreach ($relation in $db.Relations) {
        if (-not $script:NodeIndex.ContainsKey((Get-ObjectId -Group 'table' -Name $relation.Table))) {
            continue
        }
        if (-not $script:NodeIndex.ContainsKey((Get-ObjectId -Group 'table' -Name $relation.ForeignTable))) {
            continue
        }

        $fieldPairs = New-Object 'System.Collections.Generic.List[string]'
        foreach ($field in $relation.Fields) {
            $fieldPairs.Add(("{0} <-> {1}" -f $field.Name, $field.ForeignName))
        }

        Add-Edge -From (Get-ObjectId -Group 'table' -Name $relation.ForeignTable) -To (Get-ObjectId -Group 'table' -Name $relation.Table) -Label 'relation' -Kind 'relation' -Arrows 'none' -Meta @{
            name   = $relation.Name
            fields = ($fieldPairs -join '; ')
        } | Out-Null
    }

    $queryNames = New-Object 'System.Collections.Generic.List[string]'
    $queryIdx = 0
    foreach ($queryDef in $db.QueryDefs) {
        $qdName = $null
        try { $qdName = [string]$queryDef.Name } catch { continue }
        if ([string]::IsNullOrWhiteSpace($qdName) -or (Is-SystemOrTemporaryName -Name $qdName)) {
            continue
        }

        $queryName = $qdName
        $queryNames.Add($queryName)
        $queryIdx++
        Write-Progress -Activity 'Access Graph Export' -Status "Exporting query ($queryIdx): $queryName" -PercentComplete 15
        $rawInfo = Save-TextObject -AccessApp $access -Type $AcObjectType.Query -Name $queryName -Folder (Join-Path $rawDir 'queries')

        $sqlText = try { [string]$queryDef.SQL } catch { '' }
        $queryNodeId = Get-ObjectId -Group 'query' -Name $queryName
        $qConnect = try { [string]$queryDef.Connect } catch { '' }
        Add-Node -Id $queryNodeId -Label $queryName -Group 'query' -Meta @{
            connect    = $qConnect
            sqlHash    = Get-TextHash -Text $sqlText
            sqlPreview = Get-PreviewText -Text $sqlText
            rawPath    = $rawInfo.path
            rawHash    = $rawInfo.hash
            rawSize    = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $queryName -NodeId $queryNodeId -Group 'query' -IsDataObject
    }

    $allForms = @($access.CurrentProject.AllForms)
    $formTotal = $allForms.Count
    $formIdx = 0
    foreach ($obj in $allForms) {
        $name = $null
        try { $name = [string]$obj.Name } catch { continue }
        $formIdx++
        $pct = [int](25 + (15 * $formIdx / [Math]::Max($formTotal, 1)))
        Write-Progress -Activity 'Access Graph Export' -Status "Exporting form ($formIdx/$formTotal): $name" -PercentComplete $pct
        $rawInfo = Save-TextObject -AccessApp $access -Type $AcObjectType.Form -Name $name -Folder (Join-Path $rawDir 'forms')
        $nodeId = Get-ObjectId -Group 'form' -Name $name
        Add-Node -Id $nodeId -Label $name -Group 'form' -Meta @{
            rawPath = $rawInfo.path
            rawHash = $rawInfo.hash
            rawSize = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $name -NodeId $nodeId -Group 'form'
    }

    $allReports = @($access.CurrentProject.AllReports)
    $reportTotal = $allReports.Count
    $reportIdx = 0
    foreach ($obj in $allReports) {
        $name = $null
        try { $name = [string]$obj.Name } catch { continue }
        $reportIdx++
        $pct = [int](40 + (10 * $reportIdx / [Math]::Max($reportTotal, 1)))
        Write-Progress -Activity 'Access Graph Export' -Status "Exporting report ($reportIdx/$reportTotal): $name" -PercentComplete $pct
        $rawInfo = Save-TextObject -AccessApp $access -Type $AcObjectType.Report -Name $name -Folder (Join-Path $rawDir 'reports')
        $nodeId = Get-ObjectId -Group 'report' -Name $name
        Add-Node -Id $nodeId -Label $name -Group 'report' -Meta @{
            rawPath = $rawInfo.path
            rawHash = $rawInfo.hash
            rawSize = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $name -NodeId $nodeId -Group 'report'
    }

    $allMacros = @($access.CurrentProject.AllMacros)
    Write-Progress -Activity 'Access Graph Export' -Status 'Exporting macros...' -PercentComplete 50
    foreach ($obj in $allMacros) {
        $name = $null
        try { $name = [string]$obj.Name } catch { continue }
        $rawInfo = Save-TextObject -AccessApp $access -Type $AcObjectType.Macro -Name $name -Folder (Join-Path $rawDir 'macros')
        $nodeId = Get-ObjectId -Group 'macro' -Name $name
        Add-Node -Id $nodeId -Label $name -Group 'macro' -Meta @{
            rawPath = $rawInfo.path
            rawHash = $rawInfo.hash
            rawSize = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $name -NodeId $nodeId -Group 'macro'
    }

    $allModules = @($access.CurrentProject.AllModules)
    Write-Progress -Activity 'Access Graph Export' -Status 'Exporting modules...' -PercentComplete 55
    foreach ($obj in $allModules) {
        $name = $null
        try { $name = [string]$obj.Name } catch { continue }
        $rawInfo = Save-TextObject -AccessApp $access -Type $AcObjectType.Module -Name $name -Folder (Join-Path $rawDir 'modules')
        $nodeId = Get-ObjectId -Group 'module' -Name $name
        Add-Node -Id $nodeId -Label $name -Group 'module' -Meta @{
            rawPath = $rawInfo.path
            rawHash = $rawInfo.hash
            rawSize = $rawInfo.size
        } | Out-Null
        Register-ObjectNameTarget -Name $name -NodeId $nodeId -Group 'module'
    }

    $knownDataNames = @($script:DataNameTargets.Keys | Sort-Object { $_.Length } -Descending)

    Write-Progress -Activity 'Access Graph Export' -Status 'Analyzing query edges...' -PercentComplete 60
    foreach ($queryDef in $db.QueryDefs) {
        if ([string]::IsNullOrWhiteSpace($queryDef.Name) -or (Is-SystemOrTemporaryName -Name $queryDef.Name)) {
            continue
        }

        $queryName = [string]$queryDef.Name
        $queryNodeId = Get-ObjectId -Group 'query' -Name $queryName
        $sqlText = [string]$queryDef.SQL

        foreach ($referencedName in (Find-ReferencedDataNames -Text $sqlText -KnownNames $knownDataNames)) {
            foreach ($target in (Get-TargetsByName -Name $referencedName -TargetTable $script:DataNameTargets)) {
                if ($target.id -eq $queryNodeId) {
                    continue
                }

                Add-Edge -From $queryNodeId -To $target.id -Label 'uses' -Kind 'query-sql-reference' -Arrows 'to' -Meta @{ name = $referencedName } | Out-Null
            }
        }
    }

    $formEdgeIdx = 0
    foreach ($form in $allForms) {
        $name = [string]$form.Name
        $formEdgeIdx++
        $pct = [int](65 + (15 * $formEdgeIdx / [Math]::Max($formTotal, 1)))
        Write-Progress -Activity 'Access Graph Export' -Status "Analyzing form edges ($formEdgeIdx/$formTotal): $name" -PercentComplete $pct
        $rawPath = $script:NodeIndex[(Get-ObjectId -Group 'form' -Name $name)].meta.rawPath
        try {
            Add-FormOrReportEdgesFromExport -ObjectGroup 'form' -ObjectName $name -RawPath $rawPath -SqlFolder (Join-Path $rawDir 'sql') -KnownDataNames $knownDataNames
        } catch {
            Add-WarningEntry -Code 'FormEdgeParseFailed' -Message ("Failed to parse form edges for '{0}': {1}" -f $name, $_.Exception.Message) -Meta @{ owner = $name; group = 'form' }
        }
    }

    Write-Progress -Activity 'Access Graph Export' -Status 'Analyzing report edges...' -PercentComplete 82
    foreach ($report in $allReports) {
        $name = [string]$report.Name
        $rawPath = $script:NodeIndex[(Get-ObjectId -Group 'report' -Name $name)].meta.rawPath
        try {
            Add-FormOrReportEdgesFromExport -ObjectGroup 'report' -ObjectName $name -RawPath $rawPath -SqlFolder (Join-Path $rawDir 'sql') -KnownDataNames $knownDataNames
        } catch {
            Add-WarningEntry -Code 'ReportEdgeParseFailed' -Message ("Failed to parse report edges for '{0}': {1}" -f $name, $_.Exception.Message) -Meta @{ owner = $name; group = 'report' }
        }
    }

    if (-not $DisableMacroHeuristics) {
        Write-Progress -Activity 'Access Graph Export' -Status 'Analyzing macro edges...' -PercentComplete 87
        foreach ($macro in $allMacros) {
            $name = [string]$macro.Name
            $rawPath = $script:NodeIndex[(Get-ObjectId -Group 'macro' -Name $name)].meta.rawPath
            Add-MacroHeuristicEdges -MacroName $name -RawPath $rawPath -SqlFolder (Join-Path $rawDir 'sql') -KnownDataNames $knownDataNames
        }
    }

    if (-not $DisableCodeHeuristics) {
        Write-Progress -Activity 'Access Graph Export' -Status 'Building procedure index...' -PercentComplete 88
        Build-ProcIndex
    }

    if (-not $DisableCodeHeuristics) {
        Write-Progress -Activity 'Access Graph Export' -Status 'Analyzing module code...' -PercentComplete 90
        foreach ($module in $allModules) {
            $name = [string]$module.Name
            $text = $script:ModuleCodeCache[$name]
            if (-not $text) {
                $rawPath = $script:NodeIndex[(Get-ObjectId -Group 'module' -Name $name)].meta.rawPath
                if ($rawPath -and (Test-Path -LiteralPath $rawPath)) {
                    $text = Get-Content -LiteralPath $rawPath -Raw
                }
            }
            if ($text) {
                Add-CodeHeuristicEdges -OwnerNodeId (Get-ObjectId -Group 'module' -Name $name) -OwnerGroup 'module' -OwnerName $name -Text $text -SqlFolder (Join-Path $rawDir 'sql') -KnownDataNames $knownDataNames
            }
        }
    }

    Write-Progress -Activity 'Access Graph Export' -Status 'Writing graph output...' -PercentComplete 95
    $graph = [pscustomobject][ordered]@{
        meta = [ordered]@{
            database      = $databaseFullPath
            generatedAt   = [DateTime]::UtcNow.ToString('o')
            fieldNodeMode = $FieldNodeMode
            stats         = [ordered]@{
                nodeCount = $script:Nodes.Count
                edgeCount = $script:Edges.Count
                tables    = @($script:Nodes | Where-Object { $_.group -eq 'table' }).Count
                queries   = @($script:Nodes | Where-Object { $_.group -eq 'query' }).Count
                forms     = @($script:Nodes | Where-Object { $_.group -eq 'form' }).Count
                reports   = @($script:Nodes | Where-Object { $_.group -eq 'report' }).Count
                macros    = @($script:Nodes | Where-Object { $_.group -eq 'macro' }).Count
                modules   = @($script:Nodes | Where-Object { $_.group -eq 'module' }).Count
                sqlNodes  = @($script:Nodes | Where-Object { $_.group -eq 'sql' }).Count
                fieldNodes = @($script:Nodes | Where-Object { $_.group -eq 'field' }).Count
                warnings  = $script:Warnings.Count
            }
            warnings      = $script:Warnings.ToArray()
        }
        nodes = $script:Nodes.ToArray()
        edges = $script:Edges.ToArray()
    }

    $graphPath = Join-Path $OutDir 'graph.json'
    $graphJson = $graph | ConvertTo-Json -Depth 25
    Set-Content -LiteralPath $graphPath -Value $graphJson -Encoding UTF8

    Copy-ViewerIfPresent -DestinationFolder $OutDir -GraphJson $graphJson -Disabled:$SkipViewerCopy

    Write-Progress -Activity 'Access Graph Export' -Completed
    Write-Host ('Graph written to: ' + $graphPath)
    Write-Host ('Nodes: {0}  Edges: {1}  Warnings: {2}' -f $script:Nodes.Count, $script:Edges.Count, $script:Warnings.Count)
}
finally {
    if ($db) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($db) } catch {}
    }
    if ($access) {
        try { $access.CloseCurrentDatabase() } catch {}
        try { $access.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) } catch {}
    }

    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
