# 監視するリポジトリのリスト
$repos = @(
    "C:\Users\sakur\Desktop\claude作業用\github\bank-passbook",
    "C:\Users\sakur\Desktop\claude作業用\github\receipt-to-csv-ver2",
    "C:\Users\sakur\Desktop\claude作業用\github\receipt-to-csv-ver2 copy"
)

Write-Host "自動プッシュ監視を開始しました..." -ForegroundColor Green
Write-Host "停止するには Ctrl+C を押してください" -ForegroundColor Yellow
Write-Host ""

# 各リポジトリにFileSystemWatcherを設定
$watchers = @()
foreach ($repo in $repos) {
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $repo
    $watcher.IncludeSubdirectories = $true
    $watcher.EnableRaisingEvents = $true
    $watcher.Filter = "*.*"
    $watchers += $watcher
    Write-Host "監視中: $repo" -ForegroundColor Cyan
}

# 最後にプッシュした時刻（連続変更をまとめるため）
$lastPush = @{}
foreach ($repo in $repos) { $lastPush[$repo] = [datetime]::MinValue }

Write-Host ""

while ($true) {
    Start-Sleep -Seconds 5

    foreach ($repo in $repos) {
        # .gitフォルダ内の変更は無視
        $changed = & git -C $repo status --short 2>$null
        if ($changed) {
            $now = [datetime]::Now
            # 前回プッシュから10秒以上経過していたらプッシュ
            if (($now - $lastPush[$repo]).TotalSeconds -ge 10) {
                $repoName = Split-Path $repo -Leaf
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $repoName に変更を検出 → プッシュ中..." -ForegroundColor Yellow

                & git -C $repo add . 2>$null
                & git -C $repo commit -m "自動コミット $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" 2>$null
                & git -c http.sslVerify=false -C $repo push 2>$null

                $lastPush[$repo] = $now
                Write-Host "[$(Get-Date -Format 'HH:mm:ss')] $repoName プッシュ完了" -ForegroundColor Green
            }
        }
    }
}
