# Minimal static file server for this project (no Node/Python available in this environment).
$root = Split-Path -Parent $PSScriptRoot
$port = 8130
$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
$listener.Start()
Write-Host "Serving $root on http://localhost:$port/"

while ($listener.IsListening) {
    $context = $listener.GetContext()
    $reqPath = $context.Request.Url.LocalPath
    # Comme GitHub Pages : un chemin de dossier sert son index.html
    if ($reqPath.EndsWith("/")) { $reqPath += "index.html" }
    $filePath = Join-Path $root $reqPath.TrimStart("/")
    if (Test-Path $filePath -PathType Leaf) {
        $bytes = [System.IO.File]::ReadAllBytes($filePath)
        $ext = [System.IO.Path]::GetExtension($filePath)
        $ctype = switch ($ext) {
            ".html" { "text/html; charset=utf-8" }
            ".css"  { "text/css" }
            ".js"   { "application/javascript" }
            ".xml"  { "application/xml" }
            ".txt"  { "text/plain" }
            ".jpg"  { "image/jpeg" }
            ".jpeg" { "image/jpeg" }
            ".png"  { "image/png" }
            ".svg"  { "image/svg+xml" }
            default { "application/octet-stream" }
        }
        $context.Response.ContentType = $ctype
        $context.Response.ContentLength64 = $bytes.Length
        $context.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    } else {
        $context.Response.StatusCode = 404
    }
    $context.Response.OutputStream.Close()
}
