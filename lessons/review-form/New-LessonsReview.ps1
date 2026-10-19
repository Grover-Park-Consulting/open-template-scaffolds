<#
Who reads this: the AI assistant building the lesson review form, or a contributor.

Creates LessonsReview.accdb in the folder that holds lessons.accdb, with the nine linked
tables and the eight named queries the form depends on. It uses the database engine's DAO
objects only; it never starts Access. The modules, forms and startup settings are added
afterwards, as BUILD.md says.

Exit codes: 0 done; 20 the database engine is not available in this PowerShell (use the
32-bit one); 30 any other failure (message on stderr), including a LessonsReview.accdb
that already exists, which is never replaced.
#>
param(
    [Parameter(Mandatory)][string]$FolderPath
)

$ErrorActionPreference = 'Stop'
$dbVersion120 = 128

$tables = 'tblLesson', 'tblLessonEvent', 'tlkpMechanism', 'tlkpLessonKind', 'tlkpLessonConfidence',
    'tlkpLessonStatus', 'tlkpLessonOrigin', 'tlkpLessonApproval', 'tlkpLessonEventKind'

$queries = [ordered]@{
    qryLesson_frm           = 'SELECT tblLesson.* FROM tblLesson ORDER BY tblLesson.LessonSlug;'
    qryLessonEvent_sfrm     = 'SELECT tblLessonEvent.* FROM tblLessonEvent ORDER BY tblLessonEvent.LessonEventID;'
    qryMechanism_cbo        = 'SELECT tlkpMechanism.MechanismID, tlkpMechanism.MechanismName FROM tlkpMechanism ORDER BY tlkpMechanism.MechanismName;'
    qryLessonKind_cbo       = 'SELECT tlkpLessonKind.LessonKindID, tlkpLessonKind.LessonKind FROM tlkpLessonKind ORDER BY tlkpLessonKind.LessonKindID;'
    qryLessonConfidence_cbo = 'SELECT tlkpLessonConfidence.LessonConfidenceID, tlkpLessonConfidence.LessonConfidence FROM tlkpLessonConfidence ORDER BY tlkpLessonConfidence.LessonConfidenceID, tlkpLessonConfidence.LessonConfidence;'
    qryLessonOrigin_cbo     = 'SELECT tlkpLessonOrigin.LessonOriginID, tlkpLessonOrigin.LessonOrigin FROM tlkpLessonOrigin ORDER BY tlkpLessonOrigin.LessonOriginID;'
    qryLessonEventKind_cbo  = 'SELECT tlkpLessonEventKind.EventKindID, tlkpLessonEventKind.EventKind FROM tlkpLessonEventKind ORDER BY tlkpLessonEventKind.EventKind;'
    # The filter box: one extra row, All, joined to the approvals. MSysObjects only supplies a table that has a row.
    qryLessonApproval_cbo   = 'SELECT Top 1 0 AS LessonApprovalID, " ALL " AS txtLessonApproval FROM MsysObjects UNION SELECT LA.LessonApprovalID, LA.LessonApproval FROM tlkpLessonApproval AS LA ORDER BY LessonApprovalID;'
}

$db = $null
try {
    try { $eng = New-Object -ComObject DAO.DBEngine.120 } catch { [Console]::Error.WriteLine("DAO engine unavailable: $($_.Exception.Message)"); exit 20 }
    $store = Join-Path $FolderPath 'lessons.accdb'
    $target = Join-Path $FolderPath 'LessonsReview.accdb'
    if (-not (Test-Path -LiteralPath $store)) { throw "The lessons database was not found: $store" }
    if (Test-Path -LiteralPath $target) { throw "LessonsReview.accdb already exists and is not replaced: $target" }

    $db = $eng.CreateDatabase($target, ';LANGID=0x0409;CP=1252;COUNTRY=0', $dbVersion120)
    foreach ($name in $tables) {
        $td = $db.CreateTableDef($name)
        $td.Connect = ";DATABASE=$store"
        $td.SourceTableName = $name
        $db.TableDefs.Append($td)
    }
    foreach ($name in $queries.Keys) {
        $qd = $db.CreateQueryDef($name, $queries[$name])
    }
    exit 0
}
catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 30
}
finally {
    if ($db) { try { $db.Close() } catch {} }
}
