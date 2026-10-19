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

# The lookup tables: name, key column, value column, value width, and the audit columns it carries
# ('' none, 'optional' CreatedBy may be empty, 'required'). Rows come from _lookups.md.
$lookups = @(
    @{ T = 'tlkpLessonKind';       Id = 'LessonKindID';       V = 'LessonKind';       W = 20; Audit = 'optional' }
    @{ T = 'tlkpLessonConfidence'; Id = 'LessonConfidenceID'; V = 'LessonConfidence'; W = 20; Audit = 'required' }
    @{ T = 'tlkpLessonStatus';     Id = 'LessonStatusID';     V = 'LessonStatus';     W = 20; Audit = '' }
    @{ T = 'tlkpLessonOrigin';     Id = 'LessonOriginID';     V = 'LessonOrigin';     W = 10; Audit = '' }
    @{ T = 'tlkpLessonApproval';   Id = 'LessonApprovalID';   V = 'LessonApproval';   W = 10; Audit = '' }
    @{ T = 'tlkpLessonEventKind';  Id = 'EventKindID';        V = 'EventKind';        W = 20; Audit = '' }
)

function Initialize-Schema($db) {
    $ddl = @()
    foreach ($lk in $lookups) {
        $audit = switch ($lk.Audit) {
            'optional' { ', CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100), ModifiedDate DATETIME, ModifiedBy TEXT(100)' }
            'required' { ', CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100)' }
            default    { '' }
        }
        $ddl += "CREATE TABLE $($lk.T) ($($lk.Id) COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, $($lk.V) TEXT($($lk.W))$audit)"
    }
    $ddl += @(
        'CREATE TABLE tlkpMechanism (MechanismID COUNTER CONSTRAINT pkMechanism PRIMARY KEY, MechanismName TEXT(60) NOT NULL, MechanismDescription LONGTEXT, MechanismAliases TEXT(255), MechanismAlways YESNO, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'CREATE UNIQUE INDEX uxMechanismName ON tlkpMechanism (MechanismName)',
        'CREATE TABLE tblLesson (LessonID COUNTER CONSTRAINT pkLesson PRIMARY KEY, LessonSlug TEXT(100) NOT NULL, MechanismID LONG NOT NULL, LessonKindID LONG, LessonConfidenceID LONG, LessonStatusID LONG, LessonText LONGTEXT NOT NULL, LessonAppliesWhen LONGTEXT, LessonTriggersWhen LONGTEXT, LessonEvidence LONGTEXT, LessonEnvironment LONGTEXT, LessonSources LONGTEXT, LessonOriginID LONG, LessonApprovalID LONG, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'CREATE UNIQUE INDEX uxLessonSlug ON tblLesson (LessonSlug)',
        'CREATE TABLE tblLessonEvent (LessonEventID COUNTER CONSTRAINT pkLessonEvent PRIMARY KEY, LessonID LONG NOT NULL, EventKindID LONG, EventBuild TEXT(100) NOT NULL, EventOriginID LONG, CreatedDate DATETIME NOT NULL, CreatedBy TEXT(100) NOT NULL, ModifiedDate DATETIME, ModifiedBy TEXT(100))',
        'ALTER TABLE tblLesson ADD CONSTRAINT fkLessonMechanism FOREIGN KEY (MechanismID) REFERENCES tlkpMechanism (MechanismID)',
        'ALTER TABLE tblLessonEvent ADD CONSTRAINT fkLessonEventLesson FOREIGN KEY (LessonID) REFERENCES tblLesson (LessonID)',
        'ALTER TABLE tblLesson ADD CONSTRAINT tlkpLessonKindtblLesson FOREIGN KEY (LessonKindID) REFERENCES tlkpLessonKind (LessonKindID)',
        'ALTER TABLE tblLesson ADD CONSTRAINT tlkpLessonConfidencetblLesson FOREIGN KEY (LessonConfidenceID) REFERENCES tlkpLessonConfidence (LessonConfidenceID)',
        'ALTER TABLE tblLesson ADD CONSTRAINT tlkpLessonStatustblLesson FOREIGN KEY (LessonStatusID) REFERENCES tlkpLessonStatus (LessonStatusID)',
        'ALTER TABLE tblLesson ADD CONSTRAINT tlkpLessonOrigintblLesson FOREIGN KEY (LessonOriginID) REFERENCES tlkpLessonOrigin (LessonOriginID)',
        'ALTER TABLE tblLesson ADD CONSTRAINT tlkpLessonApprovaltblLesson FOREIGN KEY (LessonApprovalID) REFERENCES tlkpLessonApproval (LessonApprovalID)',
        'ALTER TABLE tblLessonEvent ADD CONSTRAINT tlkpLessonEventKindtblLessonEvent FOREIGN KEY (EventKindID) REFERENCES tlkpLessonEventKind (EventKindID)',
        'ALTER TABLE tblLessonEvent ADD CONSTRAINT tlkpLessonOrigintblLessonEvent FOREIGN KEY (EventOriginID) REFERENCES tlkpLessonOrigin (LessonOriginID)'
    )
    foreach ($s in $ddl) { $db.Execute($s, $dbFailOnError) }
    # The designer's default of 0 on these two number fields is part of the shipped design.
    $db.TableDefs.Item('tblLesson').Fields.Item('LessonConfidenceID').DefaultValue = '0'
    $db.TableDefs.Item('tblLessonEvent').Fields.Item('EventOriginID').DefaultValue = '0'
}

# The number of a lookup value. A value the lookups do not hold stops the sync.
function Get-LookupId($db, [string]$table, [string]$idCol, [string]$valCol, $value) {
    $id = Get-Scalar $db "SELECT $idCol FROM $table WHERE $valCol = $(Q $value)"
    if ($null -eq $id) { throw "'$value' is not a row of $table" }
    return [int]$id
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

    $hasKindId = $false
    foreach ($f in $db.TableDefs.Item('tblLesson').Fields) { if ($f.Name -eq 'LessonKindID') { $hasKindId = $true } }
    if (-not $hasKindId) { throw 'This lessons database is from the retired first version. Delete it; the next sync creates a new one.' }

    # Lookups: add a missing row under its own number. A row that exists is never changed or removed.
    foreach ($lk in $lookups) {
        foreach ($row in @($payload.lookups.($lk.T))) {
            if ($null -ne (Get-Scalar $db "SELECT $($lk.Id) FROM $($lk.T) WHERE $($lk.Id) = $([int]$row.id)")) { continue }
            $cols = "$($lk.Id), $($lk.V)"; $vals = "$([int]$row.id), $(Q $row.value)"
            if ($lk.Audit) { $cols += ', CreatedDate, CreatedBy'; $vals += ", $nowLit, $u" }
            $db.Execute("INSERT INTO $($lk.T) ($cols) VALUES ($vals)", $dbFailOnError)
        }
    }
    $idLibrary  = Get-LookupId $db 'tlkpLessonOrigin' 'LessonOriginID' 'LessonOrigin' 'Library'
    $idApproved = Get-LookupId $db 'tlkpLessonApproval' 'LessonApprovalID' 'LessonApproval' 'Approved'

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
        $kindId   = Get-LookupId $db 'tlkpLessonKind' 'LessonKindID' 'LessonKind' $l.kind
        $confId   = Get-LookupId $db 'tlkpLessonConfidence' 'LessonConfidenceID' 'LessonConfidence' $l.confidence
        $statusId = Get-LookupId $db 'tlkpLessonStatus' 'LessonStatusID' 'LessonStatus' $l.status
        $origin = Get-Scalar $db "SELECT LessonOriginID FROM tblLesson WHERE LessonSlug = $(Q $l.slug)"
        if ($null -eq $origin) {
            $db.Execute("INSERT INTO tblLesson (LessonSlug, MechanismID, LessonKindID, LessonConfidenceID, LessonStatusID, LessonText, LessonOriginID, LessonApprovalID, CreatedDate, CreatedBy) VALUES ($(Q $l.slug), $([int]$mid), $kindId, $confId, $statusId, '.', $idLibrary, $idApproved, $nowLit, $u)", $dbFailOnError)
            $origin = $idLibrary
        }
        if ([int]$origin -eq $idLibrary) {
            $db.Execute("UPDATE tblLesson SET MechanismID = $([int]$mid), LessonKindID = $kindId, LessonConfidenceID = $confId, LessonStatusID = $statusId WHERE LessonSlug = $(Q $l.slug)", $dbFailOnError)
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
            $evKindId = Get-LookupId $db 'tlkpLessonEventKind' 'EventKindID' 'EventKind' $ev[0]
            if ($null -eq (Get-Scalar $db "SELECT LessonEventID FROM tblLessonEvent WHERE LessonID = $lid AND EventKindID = $evKindId AND EventBuild = $(Q $ev[1]) AND EventOriginID = $idLibrary")) {
                $db.Execute("INSERT INTO tblLessonEvent (LessonID, EventKindID, EventBuild, EventOriginID, CreatedDate, CreatedBy) VALUES ($lid, $evKindId, $(Q $ev[1]), $idLibrary, $nowLit, $u)", $dbFailOnError)
            }
        }
    }

    # Export: approved, not inactive.
    $out = [ordered]@{ mechanisms = @(); lessons = @() }
    $rs = $db.OpenRecordset('SELECT * FROM tlkpMechanism ORDER BY MechanismName', $dbOpenSnapshot)
    while (-not $rs.EOF) {
        $out.mechanisms += [ordered]@{ name = $rs.Fields.Item('MechanismName').Value; description = $rs.Fields.Item('MechanismDescription').Value; aliases = $rs.Fields.Item('MechanismAliases').Value; always = [bool]$rs.Fields.Item('MechanismAlways').Value }
        $rs.MoveNext()
    }
    $rs.Close()
    $sqlExport = 'SELECT l.*, m.MechanismName, k.LessonKind, c.LessonConfidence, s.LessonStatus, o.LessonOrigin ' +
        'FROM ((((((tblLesson AS l INNER JOIN tlkpMechanism AS m ON l.MechanismID = m.MechanismID) ' +
        'INNER JOIN tlkpLessonKind AS k ON l.LessonKindID = k.LessonKindID) ' +
        'INNER JOIN tlkpLessonConfidence AS c ON l.LessonConfidenceID = c.LessonConfidenceID) ' +
        'INNER JOIN tlkpLessonStatus AS s ON l.LessonStatusID = s.LessonStatusID) ' +
        'INNER JOIN tlkpLessonOrigin AS o ON l.LessonOriginID = o.LessonOriginID) ' +
        'INNER JOIN tlkpLessonApproval AS a ON l.LessonApprovalID = a.LessonApprovalID) ' +
        "WHERE a.LessonApproval = 'Approved' AND s.LessonStatus <> 'inactive' ORDER BY l.LessonSlug"
    $rs = $db.OpenRecordset($sqlExport, $dbOpenSnapshot)
    while (-not $rs.EOF) {
        $lid = [int]$rs.Fields.Item('LessonID').Value
        $conf = @(); $contra = @()
        $er = $db.OpenRecordset("SELECT k.EventKind, e.EventBuild FROM tblLessonEvent AS e INNER JOIN tlkpLessonEventKind AS k ON e.EventKindID = k.EventKindID WHERE e.LessonID = $lid ORDER BY e.LessonEventID", $dbOpenSnapshot)
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
