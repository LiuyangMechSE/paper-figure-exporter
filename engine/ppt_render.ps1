<#
.SYNOPSIS
Render figure selections through installed desktop Microsoft PowerPoint.
.DESCRIPTION
Only a private copy of Source is opened. Original slides are exported as PNG
previews, then each selected set of top-level shape IDs is isolated and rendered
as a full-slide, print-quality PDF. The Python caller performs the final crop.
No OLE server is activated, no link update is requested, and Source is not saved.
PowerPoint is never quit, because Office may reuse the user's running instance.

Plan JSON: {"figures":[{"id":"s1f1","slide":1,"shape_ids":[2,3],
                       "canvas_pt":[-12,0,544,864]}],
            "known_shapes":{"1":[2,3,4],"2":[2]},
            "suppress_master_shapes":true}
known_shapes inventories every top-level source object, including intentionally
excluded captions/tables. Export fails if native PowerPoint exposes an unknown
object, so unsupported OOXML wrappers cannot silently disappear from the PDF.
canvas_pt is optional and uses original slide coordinates, in points. It permits
off-slide figure content to be rendered by enlarging the private slide canvas.
Canvas dimensions are limited to 4000 pt. The page-size XML is changed before a
fresh private copy opens, avoiding PowerPoint's automatic text resizing. Source
geometry is verified for each copy; selected figure elements translate together.
OutputDir must be empty or absent. The script writes render_status.json even
after a rendering failure; completed files remain available for diagnosis.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$Source,
    [Parameter(Mandatory=$true)][string]$Plan,
    [Parameter(Mandatory=$true)][string]$OutputDir,
    [ValidateRange(320,1600)][int]$PreviewWidth = 1280
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Initialize-ZipSupport {
    # Windows PowerShell 5.1 does not load the enum/ZipArchive assembly merely
    # because ZipFile's FileSystem assembly was loaded. Load both explicitly in
    # EVERY renderer process; no profile or previous session is required.
    Add-Type -AssemblyName System.IO.Compression -ErrorAction Stop
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop
    $archiveType = 'System.IO.Compression.ZipArchive' -as [type]
    $modeType = 'System.IO.Compression.ZipArchiveMode' -as [type]
    $fileType = 'System.IO.Compression.ZipFile' -as [type]
    if ($null -eq $archiveType -or $null -eq $modeType -or $null -eq $fileType) {
        throw 'PowerPoint export needs System.IO.Compression and System.IO.Compression.FileSystem. ZIP types did not load in this PowerShell process.'
    }
    return [ordered]@{
        compression_assembly = $archiveType.Assembly.FullName
        zipfile_assembly = $fileType.Assembly.FullName
        zip_archive_mode_loaded = $true
    }
}

function Release-ComReference($value) {
    if ($null -ne $value -and [System.Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}

function Write-Status {
    $status.updated_utc = [DateTime]::UtcNow.ToString('o')
    $status | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $statusPath -Encoding UTF8
}

function Get-Sha256([string]$Path) {
    $stream = [IO.File]::OpenRead($Path)
    $hasher = [Security.Cryptography.SHA256]::Create()
    try { return [BitConverter]::ToString($hasher.ComputeHash($stream)).Replace('-','') }
    finally { $hasher.Dispose(); $stream.Dispose() }
}

function Get-DeckGeometry($Deck) {
    $geometry = @{}
    $inventory = @()
    for ($slideNumber = 1; $slideNumber -le $Deck.Slides.Count; $slideNumber++) {
        $sourceSlide = $Deck.Slides.Item($slideNumber)
        $slideGeometry = @{}
        $slideInventory = @()
        try {
            for ($shapeNumber = 1; $shapeNumber -le $sourceSlide.Shapes.Count; $shapeNumber++) {
                $sourceShape = $sourceSlide.Shapes.Item($shapeNumber)
                try {
                    $slideGeometry[[int]$sourceShape.Id] = @{
                        Left=[double]$sourceShape.Left; Top=[double]$sourceShape.Top
                        Width=[double]$sourceShape.Width; Height=[double]$sourceShape.Height
                        Rotation=[double]$sourceShape.Rotation
                        LockAspectRatio=[int]$sourceShape.LockAspectRatio
                    }
                    $shapeText = ''
                    try { if ($sourceShape.HasTextFrame -eq -1) { $shapeText=[string]$sourceShape.TextFrame.TextRange.Text } } catch { }
                    $slideInventory += [ordered]@{
                        id=[int]$sourceShape.Id; name=[string]$sourceShape.Name
                        type=[int]$sourceShape.Type; visible=[int]$sourceShape.Visible
                        left_pt=[double]$sourceShape.Left; top_pt=[double]$sourceShape.Top
                        width_pt=[double]$sourceShape.Width; height_pt=[double]$sourceShape.Height
                        rotation_deg=[double]$sourceShape.Rotation
                        text=$shapeText
                    }
                } finally { Release-ComReference $sourceShape }
            }
        } finally { Release-ComReference $sourceSlide }
        $geometry[$slideNumber] = $slideGeometry
        $inventory += [ordered]@{slide=$slideNumber; shape_count=$slideInventory.Count; shapes=$slideInventory}
    }
    $status.shape_inventory = $inventory
    return $geometry
}

function Set-ShapeGeometry($Shape, $Geometry, [double]$Dx=0, [double]$Dy=0) {
    # The top-level geometry includes a group's bounding transform, so restoring
    # its original width/height also restores the scale of its child elements.
    $targets = @{
        Left=([double]$Geometry.Left + $Dx); Top=([double]$Geometry.Top + $Dy)
        Width=[double]$Geometry.Width; Height=[double]$Geometry.Height
        Rotation=[double]$Geometry.Rotation
    }
    $resize = [Math]::Abs([double]$Shape.Width - $targets.Width) -gt 0.001 -or [Math]::Abs([double]$Shape.Height - $targets.Height) -gt 0.001
    if ($resize) { $Shape.LockAspectRatio = 0 }
    foreach ($property in @('Rotation','Width','Height','Left','Top')) {
        if ([Math]::Abs([double]$Shape.$property - $targets[$property]) -gt 0.001) {
            try { $Shape.$property = [single]$targets[$property] }
            catch { throw "Shape $($Shape.Id) cannot set $property to $($targets[$property]): $($_.Exception.Message)" }
        }
    }
    if ($resize) { $Shape.LockAspectRatio = [int]$Geometry.LockAspectRatio }
    foreach ($property in @('Rotation','Width','Height','Left','Top')) {
        if ([Math]::Abs([double]$Shape.$property - $targets[$property]) -gt 0.02) {
            throw "PowerPoint could not preserve shape $($Shape.Id) $property after canvas change."
        }
    }
}

function Assert-DeckGeometry($Deck, $Geometry, [double]$Width, [double]$Height, $Figure=$null, [double]$Dx=0, [double]$Dy=0) {
    $actualWidth = [double]$Deck.PageSetup.SlideWidth
    $actualHeight = [double]$Deck.PageSetup.SlideHeight
    # PowerPoint quantizes fractional OOXML page dimensions when reading them.
    # This affects page whitespace, not the independently checked shape layout.
    # Permit at most 1/8 pt rounding and record actual dimensions for the caller.
    $canvasTolerance = 0.1251
    $translatedIds = @{}
    if ($null -ne $Figure) { foreach ($shapeId in @($Figure.shape_ids)) { $translatedIds[[int]$shapeId]=$true } }
    $status.last_canvas_check = [ordered]@{
        requested_size_pt=@($Width,$Height)
        actual_size_pt=@($actualWidth,$actualHeight)
        delta_pt=@(($actualWidth-$Width),($actualHeight-$Height))
        tolerance_pt=$canvasTolerance
    }
    if ([Math]::Abs($actualWidth - $Width) -gt $canvasTolerance -or [Math]::Abs($actualHeight - $Height) -gt $canvasTolerance) {
        throw "PowerPoint canvas mismatch: requested ${Width} x ${Height} pt; actual ${actualWidth} x ${actualHeight} pt (tolerance $canvasTolerance pt)."
    }
    for ($slideNumber = 1; $slideNumber -le $Deck.Slides.Count; $slideNumber++) {
        $sourceSlide = $Deck.Slides.Item($slideNumber)
        try {
            for ($shapeNumber = 1; $shapeNumber -le $sourceSlide.Shapes.Count; $shapeNumber++) {
                $sourceShape = $sourceSlide.Shapes.Item($shapeNumber)
                try {
                    $expected = $Geometry[$slideNumber][[int]$sourceShape.Id]
                    foreach ($property in @('Left','Top','Width','Height','Rotation')) {
                        $expectedValue = [double]$expected[$property]
                        if ($null -ne $Figure -and $slideNumber -eq [int]$Figure.slide -and $translatedIds.ContainsKey([int]$sourceShape.Id)) {
                            if ($property -eq 'Left') { $expectedValue += $Dx }
                            if ($property -eq 'Top') { $expectedValue += $Dy }
                        }
                        if ([Math]::Abs([double]$sourceShape.$property - $expectedValue) -gt 0.02) {
                            throw "Canvas copy changed slide $slideNumber shape $($sourceShape.Id) $property."
                        }
                    }
                }
                finally { Release-ComReference $sourceShape }
            }
        } finally { Release-ComReference $sourceSlide }
    }
}

function Set-PackageCanvas([string]$Path, [double]$Width, [double]$Height, $Figure, [double]$Dx, [double]$Dy) {
    # Never set PageSetup.SlideWidth/SlideHeight on a live presentation: Office
    # may scale text runs independently of shape geometry. A new OOXML copy
    # with only p:sldSz changed preserves all explicit shape/font coordinates.
    $package = [IO.Compression.ZipFile]::Open($Path, [IO.Compression.ZipArchiveMode]::Update)
    try {
        $entry = $package.GetEntry('ppt/presentation.xml')
        if ($null -eq $entry) { throw 'OOXML presentation.xml is missing.' }
        $reader = New-Object IO.StreamReader($entry.Open())
        try { [xml]$xml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $sizeNode = $xml.SelectSingleNode('/*[local-name()="presentation"]/*[local-name()="sldSz"]')
        if ($null -eq $sizeNode) { throw 'OOXML slide size is missing.' }
        $sizeNode.SetAttribute('cx', ([long][Math]::Round($Width * 12700)).ToString([Globalization.CultureInfo]::InvariantCulture))
        $sizeNode.SetAttribute('cy', ([long][Math]::Round($Height * 12700)).ToString([Globalization.CultureInfo]::InvariantCulture))
        $sizeNode.RemoveAttribute('type')
        $slideIds = $xml.SelectNodes('/*[local-name()="presentation"]/*[local-name()="sldIdLst"]/*[local-name()="sldId"]')
        $slideRid = $slideIds[[int]$Figure.slide-1].GetAttribute('id','http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        $entry.Delete()
        $replacement = $package.CreateEntry('ppt/presentation.xml')
        $writer = New-Object IO.StreamWriter($replacement.Open(), (New-Object Text.UTF8Encoding($false)))
        try { $writer.Write($xml.OuterXml) } finally { $writer.Dispose() }

        # Translate the selected top-level OOXML transforms before opening.
        # This also works for user-locked shapes whose COM Left/Top setters are
        # rejected. Preserve all lock flags and all child/group transforms.
        $relsEntry = $package.GetEntry('ppt/_rels/presentation.xml.rels')
        $reader = New-Object IO.StreamReader($relsEntry.Open())
        try { [xml]$relsXml = $reader.ReadToEnd() } finally { $reader.Dispose() }
        $target = $null
        foreach ($relationship in $relsXml.DocumentElement.ChildNodes) {
            if ($relationship.GetAttribute('Id') -eq $slideRid) { $target=$relationship.GetAttribute('Target'); break }
        }
        if ([string]::IsNullOrEmpty($target)) { throw "Cannot resolve source slide $($Figure.slide)." }
        if ($target.StartsWith('/')) { $slidePart=$target.TrimStart('/') }
        else { $slidePart=([Uri]::new([Uri]'https://ooxml.invalid/ppt/presentation.xml',$target)).AbsolutePath.TrimStart('/') }
        $slideEntry=$package.GetEntry($slidePart)
        if ($null -eq $slideEntry) { throw "Cannot find slide part $slidePart." }
        $reader=New-Object IO.StreamReader($slideEntry.Open())
        try { [xml]$slideXml=$reader.ReadToEnd() } finally { $reader.Dispose() }
        $wantedIds=@{}
        foreach ($shapeId in @($Figure.shape_ids)) { $wantedIds[[int]$shapeId]=$true }
        $seenIds=@{}
        $dxEmu=[long][Math]::Round($Dx*12700)
        $dyEmu=[long][Math]::Round($Dy*12700)
        $shapeNodes=$slideXml.SelectNodes('/*[local-name()="sld"]/*[local-name()="cSld"]/*[local-name()="spTree"]/*')
        foreach ($treeChild in $shapeNodes) {
            $variants=@($treeChild)
            if ($treeChild.LocalName -eq 'AlternateContent') {
                # Each compatibility branch represents the SAME top-level
                # object. Move each branch exactly once, never group children.
                $variants=@($treeChild.SelectNodes('./*[local-name()="Choice" or local-name()="Fallback"]/*'))
            }
            foreach ($shapeNode in $variants) {
                $identity=$shapeNode.SelectSingleNode('./*[starts-with(local-name(),"nv")]/*[local-name()="cNvPr"]')
                if ($null -eq $identity) { continue }
                $shapeId=[int]$identity.GetAttribute('id')
                if (-not $wantedIds.ContainsKey($shapeId)) { continue }
                $seenIds[$shapeId]=$true
                if ($dxEmu -eq 0 -and $dyEmu -eq 0) { continue }
                $offset=$shapeNode.SelectSingleNode('./*[local-name()="xfrm"]/*[local-name()="off"] | ./*[local-name()="spPr" or local-name()="grpSpPr"]/*[local-name()="xfrm"]/*[local-name()="off"]')
                if ($null -eq $offset) { throw "Cannot translate shape $shapeId on slide $($Figure.slide): no explicit top-level transform. Use an explicit group or a zero-origin canvas." }
                $offset.SetAttribute('x',([long]$offset.GetAttribute('x')+$dxEmu).ToString([Globalization.CultureInfo]::InvariantCulture))
                $offset.SetAttribute('y',([long]$offset.GetAttribute('y')+$dyEmu).ToString([Globalization.CultureInfo]::InvariantCulture))
            }
        }
        foreach ($shapeId in $wantedIds.Keys) { if (-not $seenIds.ContainsKey($shapeId)) { throw "Figure $($Figure.id): shape ID $shapeId is not a top-level object on slide $($Figure.slide)." } }
        $slideEntry.Delete()
        $replacement=$package.CreateEntry($slidePart)
        $writer=New-Object IO.StreamWriter($replacement.Open(),(New-Object Text.UTF8Encoding($false)))
        try { $writer.Write($slideXml.OuterXml) } finally { $writer.Dispose() }
    } finally { $package.Dispose() }
}

function Open-PrivateDeck($App, [string]$Path) {
    $security = $App.AutomationSecurity
    $alerts = $App.DisplayAlerts
    try {
        $App.AutomationSecurity = 3
        $App.DisplayAlerts = 1
        return $App.Presentations.Open($Path, -1, -1, 0)
    } finally {
        $App.AutomationSecurity = $security
        $App.DisplayAlerts = $alerts
    }
}

$sourcePath = (Resolve-Path -LiteralPath $Source).Path
$planPath = (Resolve-Path -LiteralPath $Plan).Path
$sourceExtension = [IO.Path]::GetExtension($sourcePath).ToLowerInvariant()
if ($sourceExtension -notin @('.pptx', '.pptm', '.ppsx', '.ppsm')) {
    throw 'The renderer accepts OOXML PowerPoint files (.pptx/.pptm/.ppsx/.ppsm). Convert legacy .ppt to .pptx first.'
}
$outputPath = [IO.Path]::GetFullPath($OutputDir)
if (Test-Path -LiteralPath $outputPath) {
    if (-not (Test-Path -LiteralPath $outputPath -PathType Container)) { throw "OutputDir is not a directory: $outputPath" }
    if (@(Get-ChildItem -LiteralPath $outputPath -Force).Count -ne 0) { throw "Refusing to overwrite nonempty OutputDir: $outputPath" }
} else {
    [void](New-Item -ItemType Directory -Path $outputPath)
}
$planObject = Get-Content -LiteralPath $planPath -Raw -Encoding UTF8 | ConvertFrom-Json
if ($null -eq $planObject.PSObject.Properties['figures']) { throw 'Plan must contain a figures array.' }
$figures = @($planObject.figures)
if ($figures.Count -gt 0 -and $null -eq $planObject.PSObject.Properties['known_shapes']) {
    throw 'Export plans must include known_shapes: a mapping from every slide number to all known top-level shape IDs. Generate a new analysis plan before exporting.'
}
$seenIds = @{}
foreach ($figure in $figures) {
    if ($null -eq $figure.PSObject.Properties['id'] -or [string]$figure.id -notmatch '^[A-Za-z0-9_-]+$') {
        throw 'Each figure id must contain only letters, numbers, underscores, or hyphens.'
    }
    if ($seenIds.ContainsKey([string]$figure.id)) { throw "Duplicate figure id: $($figure.id)" }
    $seenIds[[string]$figure.id] = $true
    if ($null -eq $figure.PSObject.Properties['slide'] -or [int]$figure.slide -lt 1) { throw 'Each figure needs a positive, one-based slide number.' }
    if ($null -eq $figure.PSObject.Properties['shape_ids'] -or @($figure.shape_ids).Count -eq 0) { throw "Empty shape selection: $($figure.id)" }
    if ($null -ne $figure.PSObject.Properties['canvas_pt']) {
        $box = @($figure.canvas_pt)
        if ($box.Count -ne 4) { throw "canvas_pt must have four coordinates: $($figure.id)" }
        foreach ($coordinate in $box) {
            if ([double]::IsNaN([double]$coordinate) -or [double]::IsInfinity([double]$coordinate)) { throw 'Canvas coordinates must be finite.' }
        }
        $canvasWidth = [double]$box[2] - [double]$box[0]
        $canvasHeight = [double]$box[3] - [double]$box[1]
        if ($canvasWidth -lt 1 -or $canvasHeight -lt 1 -or $canvasWidth -gt 4000 -or $canvasHeight -gt 4000) {
            throw "Canvas width/height must be between 1 and 4000 pt: $($figure.id)"
        }
    }
}
$suppressMaster = $true
if ($null -ne $planObject.PSObject.Properties['suppress_master_shapes']) { $suppressMaster = [bool]$planObject.suppress_master_shapes }
$statusPath = Join-Path $outputPath 'render_status.json'
$status = [ordered]@{
    status = 'started'
    source = $sourcePath
    source_sha256_before = Get-Sha256 $sourcePath
    source_sha256_after = $null
    source_unchanged = $null
    backend = 'Microsoft PowerPoint ExportAsFixedFormat, print intent'
    runtime = [ordered]@{
        powershell_executable = [Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
        powershell_version = $PSVersionTable.PSVersion.ToString()
        powershell_edition = $PSVersionTable.PSEdition
        clr_version = [Environment]::Version.ToString()
    }
    zip_support = $null
    suppress_master_shapes = $suppressMaster
    slide_width_pt = $null
    slide_height_pt = $null
    slide_count = $null
    geometry_restore_verified = $false
    last_canvas_check = $null
    shape_inventory = @()
    shape_inventory_file = 'shape_inventory.json'
    canvas_strategy = 'Fresh private OOXML copy per figure; page size set before open; no live PowerPoint resize'
    rendered = @()
    previews = @()
    warnings = @(
        'Embedded OLE objects use their saved PowerPoint preview. The exporter does not activate the object or recalculate its source.',
        'PowerPoint preserves supported vectors and fonts; bitmap-only objects remain bitmaps, and restricted font embedding may rasterize text.',
        'Animations, video playback, and interactive objects export only their static saved view.',
        'Only disk-saved source contents are exported; unsaved edits in an already open PowerPoint window are not included.'
    )
    error = $null
    error_context = $null
    current_operation = 'source_preparation'
    updated_utc = [DateTime]::UtcNow.ToString('o')
}
Write-Status

$workRoot = Join-Path ([IO.Path]::GetTempPath()) ('paper-fig-render-' + [Guid]::NewGuid().ToString('N'))
[void](New-Item -ItemType Directory -Path $workRoot)
$tempDeckPath = Join-Path $workRoot ('source' + $sourceExtension)
$application = $null
$presentation = $null
$oldAutomation = $null
$oldAlerts = $null
$settingsCaptured = $false
$failure = $null
try {
    $status.current_operation = 'initialize ZIP dependencies'
    $status.zip_support = Initialize-ZipSupport
    Write-Status
    $status.current_operation = 'prepare private source copy'
    Copy-Item -LiteralPath $sourcePath -Destination $tempDeckPath
    # Do not let an automatically updating linked OLE/chart object refresh during
    # Open. Change only update flags in the private OOXML copy, retaining cached
    # content and relationships. No external resource is fetched by this script.
    $archive = [IO.Compression.ZipFile]::Open($tempDeckPath, [IO.Compression.ZipArchiveMode]::Update)
    $changedParts = 0
    $externalLinks = 0
    try {
        foreach ($entry in @($archive.Entries)) {
            if ($entry.FullName -notmatch '\.(xml|rels)$') { continue }
            $reader = New-Object IO.StreamReader($entry.Open())
            try { $xmlText = $reader.ReadToEnd() } finally { $reader.Dispose() }
            if ($entry.FullName.EndsWith('.rels')) {
                $externalLinks += [regex]::Matches($xmlText, 'TargetMode="External"').Count
                continue
            }
            if ($xmlText -notmatch 'updateAutomatic|autoUpdate') { continue }
            [xml]$xml = $xmlText
            $changed = $false
            foreach ($node in $xml.SelectNodes('//*[@updateAutomatic]')) {
                if ($node.GetAttribute('updateAutomatic') -notin @('0','false','off')) {
                    $node.SetAttribute('updateAutomatic','0')
                    $changed = $true
                }
            }
            foreach ($node in $xml.SelectNodes('//*[local-name()="autoUpdate"]')) {
                if ($node.GetAttribute('val') -notin @('0','false','off')) {
                    $node.SetAttribute('val','0')
                    $changed = $true
                }
            }
            if ($changed) {
                $entryName = $entry.FullName
                $entry.Delete()
                $replacement = $archive.CreateEntry($entryName)
                $writer = New-Object IO.StreamWriter($replacement.Open(), (New-Object Text.UTF8Encoding($false)))
                try { $writer.Write($xml.OuterXml) } finally { $writer.Dispose() }
                $changedParts++
            }
        }
    } finally { $archive.Dispose() }
    if ($changedParts -gt 0) { $status.warnings += "Disabled saved automatic link-update flags in $changedParts part(s) of the temporary copy." }
    if ($externalLinks -gt 0) {
        $status.warnings += "Source has $externalLinks external relationship(s), including possible hyperlinks. No explicit link update is made. Verify any linked-media appearance against the saved source preview."
    }

    $application = New-Object -ComObject PowerPoint.Application
    $oldAutomation = $application.AutomationSecurity
    $oldAlerts = $application.DisplayAlerts
    $settingsCaptured = $true
    $application.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
    $application.DisplayAlerts = 1     # ppAlertsNone; do not show background conversion dialogs
    # ReadOnly=true still permits in-memory visibility changes. Untitled=true
    # also detaches the opened document from any save-to-original behavior.
    $presentation = $application.Presentations.Open($tempDeckPath, -1, -1, 0)
    # Restore the application-wide security choice promptly after opening.
    $application.AutomationSecurity = $oldAutomation
    $application.DisplayAlerts = $oldAlerts
    $width = [double]$presentation.PageSetup.SlideWidth
    $height = [double]$presentation.PageSetup.SlideHeight
    $status.slide_width_pt = $width
    $status.slide_height_pt = $height
    $status.slide_count = [int]$presentation.Slides.Count
    $previewHeight = [Math]::Max(1, [int][Math]::Round($PreviewWidth * $height / $width))
    foreach ($figure in $figures) {
        $status.current_operation = "prepare figure $($figure.id)"
        if ([int]$figure.slide -gt $status.slide_count) { throw "Slide outside source deck: $($figure.slide)" }
    }

    for ($slideIndex = 1; $slideIndex -le $status.slide_count; $slideIndex++) {
        $slide = $presentation.Slides.Item($slideIndex)
        try {
            $previewName = 'slide-{0:D3}.png' -f $slideIndex
            $previewPath = Join-Path $outputPath $previewName
            $slide.Export($previewPath, 'PNG', $PreviewWidth, $previewHeight)
            if (-not (Test-Path -LiteralPath $previewPath -PathType Leaf)) { throw "PowerPoint did not create preview $previewName" }
            $status.previews += [ordered]@{slide=$slideIndex; file=$previewName}
        } finally { Release-ComReference $slide }
    }
    Write-Status

    $originalGeometry = Get-DeckGeometry $presentation
    [ordered]@{
        source=$sourcePath; source_sha256=$status.source_sha256_before
        slide_width_pt=$width; slide_height_pt=$height; slides=$status.shape_inventory
    } | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath (Join-Path $outputPath 'shape_inventory.json') -Encoding UTF8
    Write-Status
    if ($figures.Count -gt 0) {
        $inventoryErrors=@()
        foreach ($slideInventory in $status.shape_inventory) {
            $slideKey=[string]$slideInventory.slide
            $knownProperty=$planObject.known_shapes.PSObject.Properties[$slideKey]
            if ($null -eq $knownProperty) { $inventoryErrors += "slide $slideKey has no known_shapes entry"; continue }
            $knownIds=@{}
            foreach ($shapeId in @($knownProperty.Value)) { $knownIds[[int]$shapeId]=$true }
            $unknownIds=@($slideInventory.shapes | Where-Object { -not $knownIds.ContainsKey([int]$_.id) } | ForEach-Object { [int]$_.id })
            if ($unknownIds.Count -gt 0) { $inventoryErrors += "slide $slideKey unknown shape IDs: $($unknownIds -join ', ')" }
        }
        if ($inventoryErrors.Count -gt 0) {
            throw ('Source inventory mismatch; no figures exported. ' + ($inventoryErrors -join '; ') + '. Inspect shape_inventory.json and regenerate the analysis plan.')
        }
    }
    $presentation.Saved = -1
    $presentation.Close()
    Release-ComReference $presentation
    $presentation = $null

    foreach ($figure in $figures) {
        $canvas = @(0.0, 0.0, $width, $height)
        if ($null -ne $figure.PSObject.Properties['canvas_pt']) { $canvas = @($figure.canvas_pt | ForEach-Object { [double]$_ }) }
        $figureDeckPath = Join-Path $workRoot ('figure-' + [string]$figure.id + $sourceExtension)
        Copy-Item -LiteralPath $tempDeckPath -Destination $figureDeckPath
        Set-PackageCanvas $figureDeckPath ($canvas[2]-$canvas[0]) ($canvas[3]-$canvas[1]) $figure (-$canvas[0]) (-$canvas[1])
        $presentation = Open-PrivateDeck $application $figureDeckPath
        Assert-DeckGeometry $presentation $originalGeometry ($canvas[2]-$canvas[0]) ($canvas[3]-$canvas[1]) $figure (-$canvas[0]) (-$canvas[1])
        $slide = $presentation.Slides.Item([int]$figure.slide)
        $visibility = @{}
        $wanted = @{}
        foreach ($shapeId in @($figure.shape_ids)) { $wanted[[int]$shapeId] = $true }
        $found = @{}
        $masterVisibility = $slide.DisplayMasterShapes
        $range = $null
        try {
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $shape = $slide.Shapes.Item($shapeIndex)
                try {
                    $id = [int]$shape.Id
                    $status.current_operation = "figure $($figure.id) shape $id"
                    $visibility[$id] = [int]$shape.Visible
                    if ($wanted.ContainsKey($id)) {
                        $found[$id] = $true
                    } else { $shape.Visible = 0 }
                } finally { Release-ComReference $shape }
            }
            $status.current_operation = "figure $($figure.id) validate shape ids"
            foreach ($shapeId in $wanted.Keys) {
                if (-not $found.ContainsKey($shapeId)) { throw "Figure $($figure.id): shape ID $shapeId is not a top-level shape on slide $($figure.slide)." }
            }
            if ($suppressMaster) { $slide.DisplayMasterShapes = 0 }
            $status.current_operation = "figure $($figure.id) print range"
            $presentation.PrintOptions.Ranges.ClearAll()
            $range = $presentation.PrintOptions.Ranges.Add([int]$figure.slide, [int]$figure.slide)
            $pdfName = [string]$figure.id + '.pdf'
            $pdfPath = Join-Path $outputPath $pdfName
            # PDF=2, print intent=2, no frame=0, vertical handout order=1,
            # slides=1, include hidden slides=-1, slide-range=4.
            # BitmapMissingFonts=true retains appearance when fonts cannot embed.
            $status.current_operation = "figure $($figure.id) PDF export"
            $presentation.ExportAsFixedFormat($pdfPath, 2, 2, 0, 1, 1, -1, $range, 4, '', $false, $true, $true, $true, $false)
            if (-not (Test-Path -LiteralPath $pdfPath -PathType Leaf) -or (Get-Item -LiteralPath $pdfPath).Length -eq 0) {
                throw "PowerPoint did not create a nonempty PDF for $($figure.id)."
            }
            $status.rendered += [ordered]@{
                id=[string]$figure.id; slide=[int]$figure.slide
                shape_ids=@($figure.shape_ids); canvas_pt=$canvas
                actual_canvas_size_pt=@([double]$presentation.PageSetup.SlideWidth,[double]$presentation.PageSetup.SlideHeight)
                file=$pdfName
            }
            Write-Status
        } finally {
            for ($shapeIndex = 1; $shapeIndex -le $slide.Shapes.Count; $shapeIndex++) {
                $shape = $slide.Shapes.Item($shapeIndex)
                try {
                    $id = [int]$shape.Id
                    if ($visibility.ContainsKey($id)) { $shape.Visible = $visibility[$id] }
                } finally { Release-ComReference $shape }
            }
            $slide.DisplayMasterShapes = $masterVisibility
            Assert-DeckGeometry $presentation $originalGeometry ($canvas[2]-$canvas[0]) ($canvas[3]-$canvas[1]) $figure (-$canvas[0]) (-$canvas[1])
            $status.geometry_restore_verified = $true
            Release-ComReference $range
            Release-ComReference $slide
            $presentation.Saved = -1
            $presentation.Close()
            Release-ComReference $presentation
            $presentation = $null
        }
    }
    $status.status = 'complete'
} catch {
    $failure = $_
    $status.status = 'failed'
    $status.error = $_.Exception.Message
    $status.error_context = [ordered]@{operation=$status.current_operation; line=$_.InvocationInfo.ScriptLineNumber; stack=$_.ScriptStackTrace}
} finally {
    if ($null -ne $presentation) {
        try {
            $presentation.Saved = -1 # Discard only changes to the private, untitled copy.
            $presentation.Close()
        } catch { $status.warnings += "Could not close temporary PowerPoint presentation: $($_.Exception.Message)" }
        Release-ComReference $presentation
    }
    if ($null -ne $application) {
        if ($settingsCaptured) {
            try { $application.AutomationSecurity = $oldAutomation; $application.DisplayAlerts = $oldAlerts }
            catch { $status.warnings += "Could not restore PowerPoint application settings: $($_.Exception.Message)" }
        }
        # Intentionally do not call Quit: PowerPoint may be shared with the user.
        Release-ComReference $application
    }
    try {
        $status.source_sha256_after = Get-Sha256 $sourcePath
        $status.source_unchanged = $status.source_sha256_before -eq $status.source_sha256_after
        if (-not $status.source_unchanged) {
            $status.warnings += 'The source changed during this run, possibly through concurrent user edits. Re-run against a stable saved source.'
        }
    } catch { $status.warnings += "Could not verify source hash: $($_.Exception.Message)" }
    # Delete only the exact private directory allocated above, after validating
    # its resolved parent and distinctive name. No user-source directory moves.
    try {
        $resolvedWorkRoot = [IO.Path]::GetFullPath($workRoot)
        $resolvedTempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')
        if ([IO.Path]::GetDirectoryName($resolvedWorkRoot).TrimEnd('\') -ne $resolvedTempRoot -or [IO.Path]::GetFileName($resolvedWorkRoot) -notmatch '^paper-fig-render-[0-9a-f]{32}$') {
            throw 'Temporary directory safety validation failed.'
        }
        Remove-Item -LiteralPath $resolvedWorkRoot -Recurse -Force
    } catch { $status.warnings += "Temporary copy cleanup failed: $($_.Exception.Message)" }
    Write-Status
}
if ($null -ne $failure) {
    Write-Error ("PowerPoint rendering failed. See {0}. {1}" -f $statusPath, $status.error)
    exit 1
}
Write-Output $statusPath
