<#
.SYNOPSIS
    One-time setup helper for the Open Template Scaffolds template library MCP server.

.DESCRIPTION
    Checks for Python, installs the server's dependencies, and PRINTS the exact,
    ready-to-paste configuration for your AI client with the real path already
    filled in. For Claude Code it also OFFERS to register the server for you, and
    does so only if you answer Y; nothing on your machine is changed without your
    say-so. Registering with the full path works wherever the library sits,
    including as a subfolder of another project.

    Run it from PowerShell:
        powershell -ExecutionPolicy Bypass -File setup.ps1

    Windows only (Microsoft Access is Windows-only, so the audience is too).
#>

$ErrorActionPreference = 'Stop'

function Write-Head($text) { Write-Host ""; Write-Host "=== $text ===" -ForegroundColor Cyan }

# --- Resolve paths from the script's own location (cwd-independent) ----------
$ServerDir   = $PSScriptRoot
$ServerPy    = Join-Path $ServerDir 'server.py'
$ServerPyFwd = $ServerPy -replace '\\', '/'   # forward slashes for JSON

if (-not (Test-Path $ServerPy)) {
    Write-Host "ERROR: server.py was not found next to this script ($ServerPy)." -ForegroundColor Red
    Write-Host "Run setup.ps1 from inside the library's mcp-server folder." -ForegroundColor Red
    exit 1
}

# --- 1. Check Python ---------------------------------------------------------
Write-Head "1. Checking Python"
$python = (Get-Command python -ErrorAction SilentlyContinue)
if (-not $python) {
    Write-Host "Python is not on your PATH." -ForegroundColor Red
    Write-Host "Install Python 3.10 or newer from https://www.python.org/ ," -ForegroundColor Yellow
    Write-Host "and tick 'Add Python to PATH' during the install. Then re-run setup.ps1." -ForegroundColor Yellow
    exit 1
}
$verRaw = (& python --version) 2>&1
Write-Host "Found: $verRaw  ($($python.Source))" -ForegroundColor Green

$verNum = [version](& python -c "import sys; print('{}.{}.{}'.format(*sys.version_info[:3]))")
if ($verNum -lt [version]'3.10.0') {
    Write-Host "Python $verNum is too old; the server needs 3.10 or newer." -ForegroundColor Red
    Write-Host "Install a newer Python from https://www.python.org/ and re-run." -ForegroundColor Yellow
    exit 1
}

# --- 2. Install dependencies -------------------------------------------------
Write-Head "2. Installing dependencies (mcp, pyyaml)"
$req = Join-Path $ServerDir 'requirements.txt'
try {
    & python -m pip install -r $req
    Write-Host "Dependencies installed." -ForegroundColor Green
}
catch {
    Write-Host "pip install reported a problem:" -ForegroundColor Red
    Write-Host $_ -ForegroundColor Red
    Write-Host "Fix the pip error above, then re-run setup.ps1. (The config below is still correct.)" -ForegroundColor Yellow
}

# --- 3. Print ready-to-use configuration ------------------------------------
Write-Head "3. Register the server with your AI client"

Write-Host ""
$AddCmd = "claude mcp add --scope user open-template-scaffolds -- python `"$ServerPyFwd`""
Write-Host "-- Claude Code ---------------------------------------------------------" -ForegroundColor White
Write-Host "Register the server for every project on this machine, using this copy's"
Write-Host "full path. This works wherever the library sits, even inside another project:"
Write-Host ""
Write-Host "    $AddCmd" -ForegroundColor Green
Write-Host ""
if (Get-Command claude -ErrorAction SilentlyContinue) {
    $answer = Read-Host "Register it now for Claude Code, for every project on this machine? (Y/N)"
    if ($answer -match '^[Yy]') {
        & claude mcp add --scope user open-template-scaffolds -- python "$ServerPyFwd"
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Registered. Start a new Claude Code session to pick it up." -ForegroundColor Green
        } else {
            Write-Host "Registration did not succeed (exit code $LASTEXITCODE). Run the command above yourself." -ForegroundColor Red
        }
    } else {
        Write-Host "Not registered. Run the command above whenever you are ready."
    }
} else {
    Write-Host "(The 'claude' command was not found, so run the line above once Claude Code is installed.)"
}
Write-Host ""
Write-Host "Alternative: opening this LIBRARY folder itself in Claude Code also offers the"
Write-Host "server (it ships a .mcp.json), but only when the library folder is what you open."

Write-Host ""
Write-Host "-- Claude Desktop (paste this block) ---------------------------------" -ForegroundColor White
Write-Host "Open  Settings -> Developer -> Edit Config  (this opens"
Write-Host "claude_desktop_config.json, under %APPDATA%\Claude\), and add the"
Write-Host "'open-template-scaffolds' entry inside the existing 'mcpServers' object:"
Write-Host ""
$snippet = @"
    "open-template-scaffolds": {
      "command": "python",
      "args": ["$ServerPyFwd"]
    }
"@
Write-Host $snippet -ForegroundColor Green
Write-Host ""
Write-Host "(If 'mcpServers' doesn't exist yet, wrap it:  { `"mcpServers`": { ...above... } } )"
Write-Host "Save the file and restart Claude Desktop."

Write-Host ""
Write-Host "-- Any other MCP client ----------------------------------------------" -ForegroundColor White
Write-Host "Same idea: a server named 'open-template-scaffolds', started with:"
Write-Host "    python `"$ServerPyFwd`"" -ForegroundColor Green

Write-Head "Done"
Write-Host "The path above is this copy's real location -- copy it exactly as shown." -ForegroundColor Green
