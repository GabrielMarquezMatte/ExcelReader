$ErrorActionPreference = 'Stop'

$smoke = $PSScriptRoot
$repo = (Resolve-Path (Join-Path $smoke '..\..')).Path
$vswhere = Join-Path ${env:ProgramFiles(x86)} 'Microsoft Visual Studio\Installer\vswhere.exe'

foreach ($folder in 'feed', 'packages', 'nuget-cache', 'bin', 'obj') {
    Remove-Item -Recurse -Force (Join-Path $smoke $folder) -ErrorAction SilentlyContinue
}

dotnet pack (Join-Path $repo 'src\ExcelReader.NetFramework') -c Release -p:Version=0.0.0-smoke -o (Join-Path $smoke 'feed')
if ($LASTEXITCODE -ne 0) { throw "dotnet pack failed ($LASTEXITCODE)" }

$msbuild = & $vswhere -latest -requires Microsoft.Component.MSBuild -find 'MSBuild\**\Bin\MSBuild.exe' | Select-Object -First 1
if (-not $msbuild) { throw 'MSBuild not found through vswhere' }

$project = Join-Path $smoke 'PackageSmoke.csproj'
& $msbuild $project -t:restore -p:RestorePackagesConfig=true -v:minimal
if ($LASTEXITCODE -ne 0) { throw "restore failed ($LASTEXITCODE)" }

& $msbuild $project -p:Configuration=Release -v:minimal
if ($LASTEXITCODE -ne 0) { throw "build failed ($LASTEXITCODE)" }

$native = Join-Path $smoke 'bin\Release\x64\ExcelReader.Native.dll'
if (-not (Test-Path $native)) { throw "missing $native" }

& (Join-Path $smoke 'bin\Release\PackageSmoke.exe') (Join-Path $repo 'RealExcel.xlsx')
exit $LASTEXITCODE
