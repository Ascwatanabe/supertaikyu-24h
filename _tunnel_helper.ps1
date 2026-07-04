# cloudflared を起動し、外部公開URLを同じウィンドウに表示する

$logOut = "$env:TEMP\cf_tunnel_out.log"
$logErr = "$env:TEMP\cf_tunnel_err.log"

Remove-Item $logOut, $logErr -ErrorAction SilentlyContinue

Start-Process -FilePath "cloudflared" `
    -ArgumentList "tunnel", "--url", "http://localhost:8765" `
    -RedirectStandardOutput $logOut `
    -RedirectStandardError $logErr `
    -WindowStyle Hidden

Write-Host "[トンネル] URL 取得中（最大30秒）..."

$found = $false
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep -Seconds 2
    $content = (Get-Content $logOut -Raw -ErrorAction SilentlyContinue) + `
               (Get-Content $logErr -Raw -ErrorAction SilentlyContinue)
    $match = [regex]::Match($content, 'https://\S+\.trycloudflare\.com')
    if ($match.Success) {
        Write-Host ""
        Write-Host "  ============================================"
        Write-Host "  外部公開URL（チームに共有してください）:"
        Write-Host ("  " + $match.Value)
        Write-Host "  ============================================"
        Write-Host ""
        $found = $true
        break
    }
}

if (-not $found) {
    Write-Host "[トンネル] URL の取得がタイムアウトしました。cloudflared のログ: $logErr"
}
