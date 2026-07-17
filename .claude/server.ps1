param($port = 8124)
$root = Split-Path -Parent $PSScriptRoot
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Listening on http://localhost:$port/"
while ($listener.IsListening) {
  $context = $listener.GetContext()
  $reqPath = $context.Request.Url.LocalPath
  if ($reqPath -eq '/') { $reqPath = '/index.html' }
  $filePath = Join-Path $root $reqPath.TrimStart('/')
  if (Test-Path $filePath -PathType Leaf) {
    $bytes = [System.IO.File]::ReadAllBytes($filePath)
    if ($filePath -like '*.html') { $context.Response.ContentType = 'text/html; charset=utf-8' }
    elseif ($filePath -like '*.js') { $context.Response.ContentType = 'application/javascript' }
    elseif ($filePath -like '*.css') { $context.Response.ContentType = 'text/css' }
    $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
  } else {
    $context.Response.StatusCode = 404
  }
  $context.Response.OutputStream.Close()
}
