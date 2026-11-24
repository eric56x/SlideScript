$ErrorActionPreference = "Stop"
$scriptDir    = Split-Path -Parent $MyInvocation.MyCommand.Path
$rootDir      = Split-Path -Parent $scriptDir
$docsDir      = Join-Path $rootDir "docs"
$toolsDir     = Join-Path $rootDir "tools"
$releasesRoot = Join-Path $rootDir "releases"
$runtimePy    = Join-Path $rootDir "runtime\python.exe"
$pyFile       = Join-Path $scriptDir "SlideScript.py"
$runnerFile   = Join-Path $scriptDir "run_engine_fast.py"
$iconFile     = Join-Path $scriptDir "SlideScript.ico"
$caBundleFile = Join-Path $scriptDir "win_ca_bundle.pem"

$pyContent = Get-Content -LiteralPath $pyFile -Raw -Encoding UTF8
if ($pyContent -match '__version__\s*=\s*["'']([^"'']+)["'']') { $version = $Matches[1] } else { throw "Version introuvable" }

Write-Host "=== Compilation et packaging de SlideScript v$version ===" -ForegroundColor Cyan
& $runtimePy -m py_compile $pyFile
if ($LASTEXITCODE -ne 0) { throw "Erreur syntaxique dans SlideScript.py" }

$todayIso = Get-Date -Format "yyyy-MM-dd"
& $runtimePy (Join-Path $toolsDir "generate_specs_from_code.py") $pyFile (Join-Path $docsDir "specifications-fr.md") (Join-Path $docsDir "specifications-en.md") $version $todayIso
& $runtimePy (Join-Path $toolsDir "convert_md_2_html.py") (Join-Path $docsDir "specifications-fr.md") (Join-Path $docsDir "specifications-fr.html") $version "fr"
& $runtimePy (Join-Path $toolsDir "convert_md_2_html.py") (Join-Path $docsDir "specifications-en.md") (Join-Path $docsDir "specifications-en.html") $version "en"
& $runtimePy (Join-Path $toolsDir "convert_md_2_html.py") (Join-Path $docsDir "best-practices.md")    (Join-Path $docsDir "best-practices.html")    $version "en"
& $runtimePy (Join-Path $toolsDir "convert_md_2_html.py") (Join-Path $rootDir "CHANGELOG.md")          (Join-Path $docsDir "changelog.html")         $version "en"

Get-Process -Name "SlideScript" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 400

$buildTemp = Join-Path "C:\Temp" "Build_SlideScript_v$version"
if (Test-Path -LiteralPath $buildTemp) { Remove-Item -LiteralPath $buildTemp -Recurse -Force }
$distDir = Join-Path $buildTemp "dist"
$workDir = Join-Path $buildTemp "work"
$specDir = Join-Path $buildTemp "spec"
New-Item -ItemType Directory -Force -Path $distDir, $workDir, $specDir | Out-Null

$argsList = @("-m", "PyInstaller", "--clean", "--noconfirm", "--onefile", "--windowed", "--collect-all", "pptx", "--collect-all", "lxml", "--hidden-import=docx", "--hidden-import=PIL", "--hidden-import=PIL.Image", "--hidden-import=PIL.ImageChops", "--hidden-import=PIL.ImageDraw", "--hidden-import=PIL.ImageStat", "--hidden-import=PIL.ImageTk", "--name", "SlideScript", "--distpath", $distDir, "--workpath", $workDir, "--specpath", $specDir)
if (Test-Path -LiteralPath $iconFile) { $argsList += @("--icon", $iconFile, "--add-data", "$iconFile;.") }
$verParts = ($version -split '\.') + @('0','0','0','0')
$v1,$v2, $v3,$v4 = $verParts[0],$verParts[1], $verParts[2],$verParts[3]
$verFile = Join-Path $buildTemp "version_info.txt"
$verContent = @"
VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=($v1,$v2, $v3,$v4),
    prodvers=($v1,$v2, $v3,$v4),
    mask=0x3f,
    flags=0x0,
    OS=0x40004,
    fileType=0x1,
    subtype=0x0,
    date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', 'SlideScript'),
        StringStruct('FileDescription', 'SlideScript — Audio/Video Transcription, Diarization & Slide Extraction'),
        StringStruct('FileVersion', '$version'),
        StringStruct('InternalName', 'SlideScript'),
        StringStruct('LegalCopyright', 'Copyright (c) 2026 SlideScript Contributors'),
        StringStruct('OriginalFilename', 'SlideScript.exe'),
        StringStruct('ProductName', 'SlideScript'),
        StringStruct('ProductVersion', '$version')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)
"@
[IO.File]::WriteAllText($verFile, $verContent, [Text.UTF8Encoding]::new($false))
$argsList += @("--version-file", $verFile)
$argsList += $pyFile
& $runtimePy @argsList

$builtExe = Join-Path $distDir "SlideScript.exe"
if (($LASTEXITCODE -ne 0) -or (-not (Test-Path -LiteralPath $builtExe))) { throw "Echec PyInstaller" }
Start-Sleep -Seconds 2
Copy-Item -LiteralPath $builtExe -Destination (Join-Path $rootDir "SlideScript.exe") -Force

$releaseDir = Join-Path $releasesRoot "v$version"
if (Test-Path -LiteralPath $releaseDir) { Remove-Item -LiteralPath $releaseDir -Recurse -Force }
New-Item -ItemType Directory -Force -Path $releaseDir | Out-Null

$stageWin64 = Join-Path $buildTemp "stage_win64"
New-Item -ItemType Directory -Force -Path $stageWin64, "$stageWin64\hf_cache", "$stageWin64\source" | Out-Null
New-Item -ItemType File -Force -Path "$stageWin64\hf_cache\.keep" | Out-Null
foreach ($repoFolder in @(
    "models--Systran--faster-whisper-tiny",
    "models--pyannote--speaker-diarization-community-1",
    "models--pyannote--segmentation-3.0"
)) {
    $srcRepo = Join-Path $rootDir "hf_cache\$repoFolder"
    if (Test-Path -LiteralPath $srcRepo) {
        Copy-Item -LiteralPath $srcRepo -Destination "$stageWin64\hf_cache\$repoFolder" -Recurse -Force
    }
}
Copy-Item -LiteralPath $builtExe -Destination "$stageWin64\SlideScript.exe" -Force
Copy-Item -LiteralPath (Join-Path $rootDir "runtime") -Destination "$stageWin64\runtime" -Recurse -Force
Copy-Item -LiteralPath (Join-Path $rootDir "ffmpeg")  -Destination "$stageWin64\ffmpeg"  -Recurse -Force
if (Test-Path $runnerFile)   { Copy-Item $runnerFile   "$stageWin64\source\run_engine_fast.py" -Force }
if (Test-Path $caBundleFile) { Copy-Item $caBundleFile "$stageWin64\source\win_ca_bundle.pem" -Force }
Copy-Item -LiteralPath $docsDir -Destination "$stageWin64\docs" -Recurse -Force
foreach ($m in @("README.md", "CHANGELOG.md", "LICENSE")) { if (Test-Path "$rootDir\$m") { Copy-Item "$rootDir\$m" "$stageWin64\$m" -Force } }

$stageSrc = Join-Path $buildTemp "stage_source"
New-Item -ItemType Directory -Force -Path $stageSrc | Out-Null
Copy-Item -LiteralPath $scriptDir -Destination "$stageSrc\source" -Recurse -Force
Copy-Item -LiteralPath $toolsDir  -Destination "$stageSrc\tools"  -Recurse -Force
Copy-Item -LiteralPath $docsDir   -Destination "$stageSrc\docs"   -Recurse -Force
foreach ($m in @("README.md", "CHANGELOG.md", "LICENSE", "requirements.txt", ".gitignore")) { if (Test-Path "$rootDir\$m") { Copy-Item "$rootDir\$m" "$stageSrc\$m" -Force } }

$win64Zip = Join-Path $releaseDir "SlideScript-v$version-win64.zip"
$srcZip   = Join-Path $releaseDir "SlideScript-v$version-source.zip"

& $runtimePy -c "import os, time, zipfile; from pathlib import Path
def make_zip(src_dir, dst_zip):
    for attempt in range(5):
        try:
            with zipfile.ZipFile(dst_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
                for p in sorted(Path(src_dir).rglob('*')):
                    if p.is_file():
                        zf.write(p, p.relative_to(src_dir))
            return
        except PermissionError:
            time.sleep(1.5)
    raise RuntimeError(f'Impossible de compresser {src_dir}')
make_zip(r'$stageWin64', r'$win64Zip')
make_zip(r'$stageSrc', r'$srcZip')"

$hashes = foreach ($z in @($win64Zip, $srcZip)) { "$((Get-FileHash $z -Algorithm SHA256).Hash.ToLower())  $(Split-Path -Leaf $z)" }
[IO.File]::WriteAllText((Join-Path $releaseDir "SHA256SUMS.txt"), ($hashes -join "`r`n") + "`r`n", [Text.UTF8Encoding]::new($false))
Remove-Item -LiteralPath $buildTemp -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "SUCCES : SlideScript.exe v$version, docs HTML et releases prêts dans $rootDir" -ForegroundColor Green