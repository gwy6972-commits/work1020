# start_server.ps1 - Simple local web server using PowerShell & .NET HttpListener
$port = 8080
$folder = $PSScriptRoot

$listener = New-Object System.Net.HttpListener
$prefix = "http://localhost:$port/"
$listener.Prefixes.Add($prefix)

try {
    $listener.Start()
    Write-Host "==========================================================" -ForegroundColor Cyan
    Write-Host "   국방예산 신규사업 분석 대시보드 로컬 서버 시작" -ForegroundColor Green
    Write-Host "   접속 주소: $prefix" -ForegroundColor Yellow
    Write-Host "   종료하려면 터미널에서 Ctrl + C 를 누르세요." -ForegroundColor Gray
    Write-Host "==========================================================" -ForegroundColor Cyan

    # Open default browser
    Start-Process $prefix

    while ($listener.IsListening) {
        $context = $listener.GetContext()
        $request = $context.Request
        $response = $context.Response

        $localPath = $request.Url.LocalPath.TrimStart('/')
        if ([string]::IsNullOrEmpty($localPath)) {
            $localPath = "index.html"
        }

        $filePath = [System.IO.Path]::Combine($folder, $localPath.Replace('/', '\'))

        if ([System.IO.File]::Exists($filePath)) {
            $ext = [System.IO.Path]::GetExtension($filePath).ToLower()
            $mime = switch ($ext) {
                ".html" { "text/html; charset=utf-8" }
                ".css"  { "text/css; charset=utf-8" }
                ".js"   { "application/javascript; charset=utf-8" }
                ".json" { "application/json; charset=utf-8" }
                ".png"  { "image/png" }
                ".jpg"  { "image/jpeg" }
                ".svg"  { "image/svg+xml" }
                ".xlsx" { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }
                default { "application/octet-stream" }
            }
            $response.ContentType = $mime
            $bytes = [System.IO.File]::ReadAllBytes($filePath)
            $response.ContentLength64 = $bytes.Length
            $response.OutputStream.Write($bytes, 0, $bytes.Length)
        } else {
            $response.StatusCode = 404
            $buf = [System.Text.Encoding]::UTF8.GetBytes("404 Not Found")
            $response.OutputStream.Write($buf, 0, $buf.Length)
        }
        $response.OutputStream.Close()
    }
} finally {
    $listener.Stop()
}
