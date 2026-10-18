<#
Who reads this: the AI assistant maintaining the lesson store, or a contributor.

Creates or updates the lessons database from an import file, then writes the
approved lessons to an export file. It uses the database engine's DAO objects
only. It never starts Access and never touches a running Access.

Exit codes: 0 done; 10 the database is in use and could not be opened;
20 the database engine is not available in this PowerShell (wrong bitness or
not installed); 30 any other failure (message on stderr).

Why no SQL parameters: DAO parameter objects fail under PowerShell's COM binding
for numbers, dates and Null, and break the parameters declared after a long-text
one. Short values go into the SQL text with quotes doubled; long text goes in
through a recordset, because a SQL string literal is limited to 255 characters.
#>
param(
    [Parameter(Mandatory)][string]$DatabasePath,
    [Parameter(Mandatory)][string]$Import,
    [Parameter(Mandatory)][string]$Export
)

$ErrorActionPreference = 'Stop'
$dbFailOnError = 128
$dbOpenDynaset = 2
$dbOpenSnapshot = 4
$user = $env:USERNAME
$nowLit = '#' + (Get-Date).ToString('MM/dd/yyyy HH:mm:ss', [Globalization.CultureInfo]::InvariantCulture) + '#'

try { $eng = New-Object -ComObject DAO.DBEngine.120 } catch { [Console]::Error.WriteLine("DAO engine unavailable: $($_.Exception.Message)"); exit 20 }

# A quoted SQL string literal, or Null for an empty value.
function Q($s) { if ($null -eq $s -or "$s" -eq '') { return 'Null' } else { return "'" + ("$s").Replace("'", "''") + "'" } }

function Get-Scalar($db, [string]$sql) {
    $rs = $db.OpenRecordset($sql, $dbOpenSnapshot)
    try { if ($rs.EOF) { return $null } else { return $rs.Fields.Item(0).Value } } finally { $rs.Close() }
}

# Write long text fields of the one row that $sql selects. Empty values leave the field Null.
function Set-Fields($db, [string]$sql, $values) {
    $rs = $db.OpenRecordset($sql, $dbOpenDynaset)
    try {
        $rs.Edit()
        foreach ($k in $values.Keys) {
            $v = $values[$k]
            if ($null -ne $v -and "$v" -ne '') { $rs.Fields.Item($k).Value = [string]$v } else { $rs.Fields.Item($k).Value = $null }
        }
        $rs.Update()
    } finally { $rs.Close() }
}

function Initialize-Schema($db) {
    $ddl = @(
        'CREATE TABLE tlkpMechanism (MechanismID COUNTER CONSTRAINT pkMechanism PRIMARY KEY, MechanismName TEXT(60) NOT NULL, MechanismDescription LONGTEXT, MechanismAliases TEXT(255), MechanismAlways YESNO, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'CREATE UNIQUE INDEX uxMechanismName ON tlkpMechanism (MechanismName)',
        'CREATE TABLE tblLesson (LessonID COUNTER CONSTRAINT pkLesson PRIMARY KEY, LessonSlug TEXT(100) NOT NULL, MechanismID LONG NOT NULL, LessonKind TEXT(20) NOT NULL, LessonConfidence TEXT(20) NOT NULL, LessonStatus TEXT(20) NOT NULL, LessonText LONGTEXT NOT NULL, LessonAppliesWhen LONGTEXT, LessonTriggersWhen LONGTEXT, LessonEvidence LONGTEXT, LessonEnvironment LONGTEXT, LessonSources LONGTEXT, LessonOrigin TEXT(10) NOT NULL, LessonApproval TEXT(10) NOT NULL, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'CREATE UNIQUE INDEX uxLessonSlug ON tblLesson (LessonSlug)',
        'ALTER TABLE tblLesson ADD CONSTRAINT fkLessonMechanism FOREIGN KEY (MechanismID) REFERENCES tlkpMechanism (MechanismID)',
        'CREATE TABLE tblLessonEvent (LessonEventID COUNTER CONSTRAINT pkLessonEvent PRIMARY KEY, LessonID LONG NOT NULL, EventKind TEXT(20) NOT NULL, EventBuild TEXT(100) NOT NULL, EventOrigin TEXT(10) NOT NULL, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'ALTER TABLE tblLessonEvent ADD CONSTRAINT fkLessonEventLesson FOREIGN KEY (LessonID) REFERENCES tblLesson (LessonID)'
    )
    foreach ($s in $ddl) { $db.Execute($s, $dbFailOnError) }
}

$db = $null
try {
    $payload = Get-Content -LiteralPath $Import -Raw -Encoding UTF8 | ConvertFrom-Json
    $u = Q $user

    if (Test-Path -LiteralPath $DatabasePath) {
        try { $db = $eng.OpenDatabase($DatabasePath, $false, $false) }
        catch {
            if ($_.Exception.Message -match 'already in use|exclusively|could not lock|3045|3051|3050|3260') {
                [Console]::Error.WriteLine("Database in use: $($_.Exception.Message)"); exit 10
            }
            throw
        }
    } else {
        $db = $eng.CreateDatabase($DatabasePath, ';LANGID=0x0409;CP=1252;COUNTRY=0', 128)
        Initialize-Schema $db
    }

    # Mechanisms: the library owns these rows.
    foreach ($m in @($payload.mechanisms)) {
        if (-not $m.name) { throw 'A mechanism has no name' }
        $yn = if ($m.always) { 'True' } else { 'False' }
        $byName = "SELECT * FROM tlkpMechanism WHERE MechanismName = $(Q $m.name)"
        if ($null -eq (Get-Scalar $db "SELECT MechanismID FROM tlkpMechanism WHERE MechanismName = $(Q $m.name)")) {
            $db.Execute("INSERT INTO tlkpMechanism (MechanismName, MechanismAlways, CreatedDate, CreatedBy) VALUES ($(Q $m.name), $yn, $nowLit, $u)", $dbFailOnError)
        } else {
            $db.Execute("UPDATE tlkpMechanism SET MechanismAlways = $yn, ModifiedDate = $nowLit, ModifiedBy = $u WHERE MechanismName = $(Q $m.name)", $dbFailOnError)
        }
        Set-Fields $db $byName ([ordered]@{ MechanismDescription = $m.description; MechanismAliases = $m.aliases })
    }

    # Lessons: new library ids are added as Approved. A Library row gets its text fields
    # re-imposed from the text on every run. Approval, Own rows and events set in the
    # database are never overwritten.
    foreach ($l in @($payload.lessons)) {
        $mid = Get-Scalar $db "SELECT MechanismID FROM tlkpMechanism WHERE MechanismName = $(Q $l.mechanism)"
        if ($null -eq $mid) { throw "Lesson $($l.slug) names unknown mechanism $($l.mechanism)" }
        $bySlug = "SELECT * FROM tblLesson WHERE LessonSlug = $(Q $l.slug)"
        $origin = Get-Scalar $db "SELECT LessonOrigin FROM tblLesson WHERE LessonSlug = $(Q $l.slug)"
        if ($null -eq $origin) {
            $db.Execute("INSERT INTO tblLesson (LessonSlug, MechanismID, LessonKind, LessonConfidence, LessonStatus, LessonText, LessonOrigin, LessonApproval, CreatedDate, CreatedBy) VALUES ($(Q $l.slug), $([int]$mid), $(Q $l.kind), $(Q $l.confidence), $(Q $l.status), '.', 'Library', 'Approved', $nowLit, $u)", $dbFailOnError)
            $origin = 'Library'
        }
        if ($origin -eq 'Library') {
            $db.Execute("UPDATE tblLesson SET MechanismID = $([int]$mid), LessonKind = $(Q $l.kind), LessonConfidence = $(Q $l.confidence), LessonStatus = $(Q $l.status) WHERE LessonSlug = $(Q $l.slug)", $dbFailOnError)
            Set-Fields $db $bySlug ([ordered]@{
                LessonText = $l.lesson; LessonAppliesWhen = $l.applies_when; LessonTriggersWhen = $l.triggers_when
                LessonEvidence = $l.evidence; LessonEnvironment = $l.environment; LessonSources = $l.sources
            })
        }
        $lid = [int](Get-Scalar $db "SELECT LessonID FROM tblLesson WHERE LessonSlug = $(Q $l.slug)")
        $events = @()
        foreach ($b in @($l.confirmations)) { if ($b) { $events += , @('Confirmation', $b) } }
        foreach ($b in @($l.contradictions)) { if ($b) { $events += , @('Contradiction', $b) } }
        foreach ($ev in $events) {
            if ($null -eq (Get-Scalar $db "SELECT LessonEventID FROM tblLessonEvent WHERE LessonID = $lid AND EventKind = $(Q $ev[0]) AND EventBuild = $(Q $ev[1]) AND EventOrigin = 'Library'")) {
                $db.Execute("INSERT INTO tblLessonEvent (LessonID, EventKind, EventBuild, EventOrigin, CreatedDate, CreatedBy) VALUES ($lid, $(Q $ev[0]), $(Q $ev[1]), 'Library', $nowLit, $u)", $dbFailOnError)
            }
        }
    }

    # Export: approved, not retired.
    $out = [ordered]@{ mechanisms = @(); lessons = @() }
    $rs = $db.OpenRecordset('SELECT * FROM tlkpMechanism ORDER BY MechanismName', $dbOpenSnapshot)
    while (-not $rs.EOF) {
        $out.mechanisms += [ordered]@{ name = $rs.Fields.Item('MechanismName').Value; description = $rs.Fields.Item('MechanismDescription').Value; aliases = $rs.Fields.Item('MechanismAliases').Value; always = [bool]$rs.Fields.Item('MechanismAlways').Value }
        $rs.MoveNext()
    }
    $rs.Close()
    $rs = $db.OpenRecordset("SELECT l.*, m.MechanismName FROM tblLesson AS l INNER JOIN tlkpMechanism AS m ON l.MechanismID = m.MechanismID WHERE l.LessonApproval = 'Approved' AND l.LessonStatus <> 'retired' ORDER BY l.LessonSlug", $dbOpenSnapshot)
    while (-not $rs.EOF) {
        $lid = [int]$rs.Fields.Item('LessonID').Value
        $conf = @(); $contra = @()
        $er = $db.OpenRecordset("SELECT EventKind, EventBuild FROM tblLessonEvent WHERE LessonID = $lid ORDER BY LessonEventID", $dbOpenSnapshot)
        while (-not $er.EOF) {
            if ($er.Fields.Item('EventKind').Value -eq 'Confirmation') { $conf += $er.Fields.Item('EventBuild').Value } else { $contra += $er.Fields.Item('EventBuild').Value }
            $er.MoveNext()
        }
        $er.Close()
        $out.lessons += [ordered]@{
            slug = $rs.Fields.Item('LessonSlug').Value; mechanism = $rs.Fields.Item('MechanismName').Value; kind = $rs.Fields.Item('LessonKind').Value
            confidence = $rs.Fields.Item('LessonConfidence').Value; status = $rs.Fields.Item('LessonStatus').Value; lesson = $rs.Fields.Item('LessonText').Value
            applies_when = $rs.Fields.Item('LessonAppliesWhen').Value; triggers_when = $rs.Fields.Item('LessonTriggersWhen').Value
            evidence = $rs.Fields.Item('LessonEvidence').Value; environment = $rs.Fields.Item('LessonEnvironment').Value; sources = $rs.Fields.Item('LessonSources').Value
            origin = $rs.Fields.Item('LessonOrigin').Value; confirmations = $conf; contradictions = $contra
        }
        $rs.MoveNext()
    }
    $rs.Close()
    $json = $out | ConvertTo-Json -Depth 6 -Compress
    [IO.File]::WriteAllText($Export, $json, (New-Object Text.UTF8Encoding $false))
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 30
}
finally {
    if ($db) { try { $db.Close() } catch {} }
}
