# (로컬 전용) 환경변수 미설정 시 기본값 지정
if (-not $env:SLACK_WEBHOOK_URL -or [string]::IsNullOrWhiteSpace($env:SLACK_WEBHOOK_URL)) {
    $env:SLACK_WEBHOOK_URL = "https://hooks.slack.com/services/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"
}
# (동적 포트 기동 시 이 값은 아래 Allure 오픈 블록에서 설정됩니다)
if (-not $env:LIVE_ALLURE_URL) { $env:LIVE_ALLURE_URL = $null }

Clear-Host

# Timestamp
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Paths
$resultsDir = "allure-results\$timestamp"
$reportDir = "allure-report\$timestamp"
$zipDir = "allure-report-zips"
$logFile = "logs\test_run_$timestamp.log"
New-Item -ItemType Directory -Path "logs" -ErrorAction SilentlyContinue | Out-Null

# 🔸 절전모드 및 화면 꺼짐 방지 설정 (관리자 권한 필요)
Write-Host "`n[INFO] 절전모드 및 화면 꺼짐 방지 설정 중..."
try {
    powercfg -change -standby-timeout-ac 0
    powercfg -change -monitor-timeout-ac 0
    powercfg -change -disk-timeout-ac 0
    Write-Host "[OK] 절전모드 설정 완료"
} catch {
    Write-Host "[WARN] 관리자 권한이 필요하거나 설정 실패: $_"
}

# 🔸 ADB 연결 확인
Write-Host "`n[INFO] ADB 디바이스 연결 상태 확인 중..."
$adbOutput = & adb devices
$adbLines = $adbOutput -split "`n" | Where-Object { $_ -match "device$" -and $_ -notmatch "List of devices" }
if ($adbLines.Count -gt 0) {
    Write-Host "[OK] 연결된 ADB 디바이스:"
    $adbLines | ForEach-Object { Write-Host "  - $($_.Trim())" }
} else {
    Write-Host "[FAIL] 연결된 ADB 디바이스가 없습니다. 연결 후 다시 시도하세요."
    exit 1
}

# 🔸 Appium 서버 상태 확인 및 자동 실행
$portOpen = Test-NetConnection -ComputerName "localhost" -Port 4723 | Select-Object -ExpandProperty TcpTestSucceeded
if (-not $portOpen) {
    Write-Host "[WARN] Appium 서버가 꺼져 있습니다."
    Write-Host "[INFO] Appium 서버를 시작합니다..."
    Start-Process "cmd.exe" "/c start /B appium"
    Start-Sleep -Seconds 8
} else {
    Write-Host "[OK] Appium 서버가 실행 중입니다. (포트 4723)"
}

# Prepare directories
if (-not (Test-Path $resultsDir)) {
    New-Item -ItemType Directory -Path $resultsDir | Out-Null
}

# Run Pytest
Write-Host "`n[INFO] Pytest 테스트 실행 중..."
# === 콘솔/파이썬 출력 UTF-8 강제 ===
try {
  [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
  $OutputEncoding = [Console]::OutputEncoding
} catch {}
$env:PYTHONIOENCODING = 'utf-8'
$env:PYTHONUTF8       = '1'
try { chcp 65001 > $null } catch {}
$pytestCommand = "pytest -s -o log_cli=true test_tms_updated.py --alluredir=`"$resultsDir`""
# Invoke-Expression "$pytestCommand"   # (중복 실행 방지를 위해 주석 처리)
cmd /c $pytestCommand 2>&1 | Tee-Object -FilePath $logFile
$testExitCode = $LASTEXITCODE

# Generate Allure report
Write-Host "`n[INFO] Allure 리포트 생성 중..."
try {
    allure generate "$resultsDir" -o "$reportDir" --clean | Out-Null
    Write-Host "[OK] 리포트 생성 완료: $reportDir"
} catch {
    Write-Host "[FAIL] 리포트 생성 실패: $_"
    exit 1
}

# Compress report
if (-not (Test-Path $zipDir)) {
    New-Item -ItemType Directory -Path $zipDir | Out-Null
}
$zipPath = "$zipDir\allure_report_$timestamp.zip"
Compress-Archive -Path "$reportDir\*" -DestinationPath $zipPath -Force
Write-Host "[OK] ZIP 저장 완료: $zipPath"

# Open report
Write-Host "`n[INFO] Allure 리포트를 브라우저로 실행합니다..."

# === Allure 서버를 '빈 포트'에 기동하고, 해당 URL을 Slack에 사용 ===
function Get-FreeTcpPort {
  $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
  $l.Start()
  $p = $l.LocalEndpoint.Port
  $l.Stop()
  return $p
}

# Application(.bat/.cmd/.exe)만 대상으로, 여러 개라면 첫 번째만 사용
$allureCmdObj = Get-Command allure -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
$allureCmd    = if ($allureCmdObj) { [string]$allureCmdObj.Source } else { $null }
$allureExists = -not [string]::IsNullOrWhiteSpace($allureCmd)

# (디버그) 실제 사용되는 Allure 실행 파일 경로 표시
Write-Host "[INFO] Using Allure binary: $allureCmd"

$port = Get-FreeTcpPort
$allureArgs = "open `"$reportDir`" -h 127.0.0.1 -p $port"

# ArgumentList는 문자열 1개 또는 문자열 배열 허용.
# 환경에 따라 배열 섞임을 방지하기 위해 단일 문자열로 합쳐 안정화
if ($allureArgs -is [System.Array]) {
    $allureArgs = ($allureArgs -join ' ')
}


try {
  if (-not $allureExists) { throw "Allure CLI not found in PATH." }

  # 새 콘솔 창 없이 백그라운드로 서버 기동
  # 혹시라도 $allureCmd가 배열이면 첫 요소만 사용 (이중 안전장치)
  if ($allureCmd -is [System.Array]) { $allureCmd = [string]($allureCmd | Select-Object -First 1) }
  Start-Process -FilePath $allureCmd -ArgumentList $allureArgs -WindowStyle Hidden

  Start-Sleep -Seconds 4  # 기동 대기 (필요시 3~4로 늘려도 됨)

  $env:LIVE_ALLURE_URL = "http://127.0.0.1:$port/"
  Write-Host "[OK] Allure 서버 시작: $($env:LIVE_ALLURE_URL)"
  # 브라우저 강제 오픈 (Windows 표준 방식)
  Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"$($env:LIVE_ALLURE_URL)`""

} catch {
  Write-Warning "Allure open 실패: $($_.Exception.Message) → index.html 직접 오픈으로 대체"
  $env:LIVE_ALLURE_URL = $null
  $indexPath = (Join-Path $reportDir "index.html")
  Start-Process -FilePath "cmd.exe" -ArgumentList "/c start `"$indexPath`""

}

# Summary
# === Slack notify (보안: 환경변수 SLACK_WEBHOOK_URL 사용) ===
try {
    $webhook = $env:SLACK_WEBHOOK_URL
    if ($webhook) {   # 성공/실패 무조건 전송
        # 상태 이모지 결정
        $code = if ($testExitCode) { [int]$testExitCode } else { [int]$LASTEXITCODE }
        $statusEmoji = if ($code -eq 0) { "✅" } else { "❌" }
        $statusText  = if ($code -eq 0) { "성공" } else { "실패" }

        # Allure 링크: LIVE_ALLURE_URL이 있으면 그걸, 없으면 로컬 파일 경로
        $allureLink = if ($env:LIVE_ALLURE_URL) {
            $env:LIVE_ALLURE_URL               # 예: http://localhost:5252/
        } else {
            $p = (Resolve-Path $reportDir).Path.Replace('\','/')
            "file:///$p/index.html"            # 로컬 Allure HTML
        }

        # Allure summary 파싱(가능할 때)
        $passed = $failed = $broken = $skipped = $unknown = $total = $null
        $summaryJson = Join-Path $reportDir "widgets\summary.json"
        if (Test-Path $summaryJson) {
            try {
                $summary = Get-Content $summaryJson -Raw | ConvertFrom-Json
                $stat = $summary.stat
                $total   = $stat.total
                $passed  = $stat.passed
                $failed  = $stat.failed
                $broken  = $stat.broken
                $skipped = $stat.skipped
                $unknown = $stat.unknown
            } catch {}
        }

        # 마지막 로그 20줄 (있을 때만)
        # Python 로거가 남긴 최신 파일(UTF-8)을 우선 사용, 없으면 PS 로그로 대체
        $tail = ""
        $pyLog = Get-ChildItem -Path "logs" -Filter "test_run_*.log" -File -ErrorAction SilentlyContinue |
                Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($pyLog) {
            $tail = (Get-Content -Path $pyLog.FullName -Tail 40 -Encoding UTF8 -ErrorAction SilentlyContinue) -join "`n"
        } elseif (Test-Path $logFile) {
            $tail = (Get-Content -Path $logFile -Tail 40 -ErrorAction SilentlyContinue) -join "`n"
        }


        # 메시지 구성 (채널 표시는 정보용 텍스트)
        $lines = @()
        $lines += "$statusEmoji TMS 테스트 $statusText"
        $lines += "• Report: $allureLink"
        $lines += "• Channel: #tms-autotest-alerts"
        if ($null -ne $total) {
            $lines += "• Stats: total=$total, passed=$passed, failed=$failed, broken=$broken, skipped=$skipped, unknown=$unknown"
        }
        if ($tail) {
            $lines += '```'
            $lines += $tail
            $lines += '```'
                    }
        $payload = @{ text = ($lines -join "`n") } | ConvertTo-Json -Depth 4 -Compress
        $bytes = [Text.Encoding]::UTF8.GetBytes($payload)
        Invoke-RestMethod -Uri $webhook -Method POST -ContentType 'application/json; charset=utf-8' -Body $bytes | Out-Null
        Write-Host "[OK] Slack 알림 전송 완료"
    } else {
        Write-Host "[SKIP] SLACK_WEBHOOK_URL 미설정. 알림 생략"
    }
} catch {
    Write-Host "[WARN] Slack 알림 실패: $($_.Exception.Message)"
}

Write-Host "`n[SUMMARY] 실행 요약"
Write-Host "결과 디렉토리  : $resultsDir"
Write-Host "리포트 디렉토리: $reportDir"
Write-Host "ZIP 저장 경로  : $zipPath"
Write-Host "로그 파일      : $logFile"
