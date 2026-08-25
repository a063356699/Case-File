$ErrorActionPreference = "Stop"

$projectRoot = Split-Path -Parent $PSScriptRoot
$nodeBin = "C:\Users\user\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin"
$env:PATH = $nodeBin + ";" + $env:PATH

& "$projectRoot\node_modules\.bin\esbuild.cmd" `
  "$projectRoot\scripts\standalone-entry.tsx" `
  --bundle `
  --platform=browser `
  --format=iife `
  --loader:.xlsm=dataurl `
  --loader:.png=dataurl `
  --outfile="$projectRoot\work\standalone-bundle.js"

if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$css = Get-Content -Raw -Encoding UTF8 -LiteralPath "$projectRoot\app\globals.css"
$css = $css -replace '@import\s+"tailwindcss";?', ""
$logoData = "data:image/png;base64," + [Convert]::ToBase64String([System.IO.File]::ReadAllBytes("$projectRoot\public\taiching-logo.png"))
$css = $css.Replace("/taiching-logo.png", $logoData)
$bundle = Get-Content -Raw -Encoding UTF8 -LiteralPath "$projectRoot\work\standalone-bundle.js"
$outputDir = "$projectRoot\docs"
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null
[System.IO.File]::WriteAllText("$outputDir\.nojekyll", "")

$html = @"
<!doctype html>
<html lang="zh-Hant">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <meta name="color-scheme" content="light">
  <title>&#32317;&#34920; &#21934;&#27231;html</title>
  <style>$css</style>
</head>
<body>
  <div id="root"></div>
  <script>$bundle</script>
</body>
</html>
"@

[System.IO.File]::WriteAllText("$outputDir\index.html", $html, [System.Text.UTF8Encoding]::new($false))
Write-Host "Built docs/index.html"
