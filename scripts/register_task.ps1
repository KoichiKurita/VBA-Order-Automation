<#
.SYNOPSIS
  受注取込バッチをタスクスケジューラに登録する

.DESCRIPTION
  ・毎日指定の時刻に scripts\run_batch.vbs を実行する
  ・「ユーザーがログオンしているときのみ実行」で登録する
    （Microsoft は、ログオンしていない状態での Office の自動実行をサポートしていないため）
  ・前回の実行が終わっていなければ新しい実行は開始しない／1時間で強制終了
  ・実行できなかった時刻（PCの電源が切れていた等）があれば、起動後に実行する

.EXAMPLE
  .\scripts\register_task.ps1 -Time 07:00
#>
param(
    [string]$Time = "07:00",
    [string]$TaskName = "DailyOrderImport",
    [string]$TaskPath = "\VBA-Order-Automation\"
)

$ErrorActionPreference = "Stop"
$root   = Split-Path -Parent $PSScriptRoot
$vbs    = Join-Path $PSScriptRoot "run_batch.vbs"
$book   = Join-Path $root "OrderTool.xlsm"
if (-not (Test-Path $vbs))  { throw "run_batch.vbs が見つかりません: $vbs" }
if (-not (Test-Path $book)) { throw "OrderTool.xlsm が見つかりません: $book" }

$action    = New-ScheduledTaskAction -Execute "cscript.exe" -Argument "//nologo `"$vbs`" `"$book`"" -WorkingDirectory $root
$trigger   = New-ScheduledTaskTrigger -Daily -At $Time
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" -LogonType Interactive -RunLevel Limited
$settings  = New-ScheduledTaskSettingsSet `
                -MultipleInstances IgnoreNew `
                -ExecutionTimeLimit (New-TimeSpan -Hours 1) `
                -StartWhenAvailable `
                -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath `
    -Action $action -Trigger $trigger -Principal $principal -Settings $settings `
    -Description "受注CSVの自動取込とレポート作成（VBA-Order-Automation）" -Force | Out-Null

Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath |
    Select-Object TaskPath, TaskName, State,
        @{n = "Time"; e = { $_.Triggers[0].StartBoundary }} | Format-List
Write-Host "登録しました。すぐに試す場合： Start-ScheduledTask -TaskPath '$TaskPath' -TaskName '$TaskName'" -ForegroundColor Green
