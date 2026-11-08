#Requires -Version 5.1
<#
.SYNOPSIS
    WindowOps.ps1 的端到端演示（PowerShell 版）

.DESCRIPTION
    依次演示：
      1. 枚举当前桌面上打开的窗口（句柄 / PID / 进程 / 类名 / 位置尺寸 / 标题）
      2. 启动一个控制台窗口作为演示目标，并按标题找到它
      3. 查看窗口详细信息（不止标题）
      4. 激活窗口 + SendInput 真实键盘输入（支持中文）+ 组合键
      5. 启动“字符映射表”，枚举子控件，用 WM_SETTEXT / WM_GETTEXT 读写控件文本
      6. 鼠标点击 + 真实键盘输入到指定控件
      7. 关闭演示窗口

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\demo.ps1

.EXAMPLE
    .\demo.ps1 -KeepOpen    # 演示结束后保留窗口，方便手动观察
#>
[CmdletBinding()]
param([switch]$KeepOpen)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [Text.Encoding]::UTF8 } catch { }

$here = Split-Path -Parent $MyInvocation.MyCommand.Path
. (Join-Path $here 'WindowOps.ps1')

function Write-Section([string]$Text) {
    ''
    ('=' * 72)
    "  $Text"
    ('=' * 72)
}

function Wait-Window {
    param(
        [string]$TitlePattern,
        [string]$ProcName,
        [int]$TimeoutSeconds = 15
    )
    $deadline = (Get-Date).AddSeconds($TimeoutSeconds)
    while ((Get-Date) -lt $deadline) {
        $found = @(Get-TargetWindows -TitlePattern $TitlePattern -ProcName $ProcName) | Select-Object -First 1
        if ($found) { return $found }
        Start-Sleep -Milliseconds 200
    }
    return $null
}

$consoleProc = $null
$charmapProc = $null
$consoleWin = $null
$charmapWin = $null

try {
    # ------------------------------------------------------------------
    Write-Section '1. 枚举当前打开的窗口（包含但不限于标题）'
    $all = @(Get-TargetWindows -AllowUnfiltered)
    "共找到 $($all.Count) 个可见顶层窗口，下面列出前 8 个："
    $all | Select-Object -First 8 | ForEach-Object { ConvertTo-WindowObject $_ } |
        Format-Table -AutoSize | Out-String -Width 400

    # ------------------------------------------------------------------
    Write-Section '2. 启动演示目标：一个控制台窗口'
    $title = 'WinOps-PSDemo'
    $outFile = Join-Path $env:TEMP 'winops_ps_demo.txt'
    $outFile2 = Join-Path $env:TEMP 'winops_ps_keys.txt'
    Remove-Item -LiteralPath $outFile, $outFile2 -ErrorAction SilentlyContinue

    $consoleProc = Start-Process cmd -ArgumentList "/k title $title" -PassThru
    $consoleWin = Wait-Window -TitlePattern $title
    if (-not $consoleWin) { throw '演示控制台窗口没有出现' }

    "按标题找到窗口: [$($consoleWin.Title.Trim())]"
    "  句柄   = 0x$('{0:X}' -f $consoleWin.Handle.ToInt64())"
    "  PID    = $($consoleWin.Pid)"
    "  进程   = $($consoleWin.ProcessName)"
    "  类名   = $($consoleWin.ClassName)"

    # ------------------------------------------------------------------
    Write-Section '3. 窗口详细信息'
    ConvertTo-WindowObject $consoleWin | Format-List | Out-String -Width 300

    # ------------------------------------------------------------------
    Write-Section '4. 激活窗口 + SendInput 真实键盘输入（中文可输入）'
    $activated = [Win32Ext]::ForceForeground($consoleWin.Handle)
    "激活成功 = $activated"
    Start-Sleep -Milliseconds 300

    $command = 'echo PowerShell-SendInput-OK > "' + $outFile + '"'
    "输入命令: $command"
    [Win32Ext]::TypeText($command)
    [Win32Ext]::SendCombo('enter')
    Start-Sleep -Milliseconds 1200

    if (Test-Path -LiteralPath $outFile) {
        "控制台执行结果: $((Get-Content -LiteralPath $outFile -Raw).Trim())"
    } else {
        '未检测到结果文件（键盘输入可能没有送达）'
    }

    # ------------------------------------------------------------------
    Write-Section '5. 组合键：Ctrl+L 清屏，再用 Enter 执行命令'
    [Win32Ext]::SendCombo('ctrl+l')
    Start-Sleep -Milliseconds 200
    $command2 = 'echo Keys-OK > "' + $outFile2 + '"'
    [Win32Ext]::TypeText($command2)
    [Win32Ext]::SendCombo('enter')
    Start-Sleep -Milliseconds 1200

    if (Test-Path -LiteralPath $outFile2) {
        "组合键执行结果: $((Get-Content -LiteralPath $outFile2 -Raw).Trim())"
    } else {
        '未检测到结果文件'
    }

    # ------------------------------------------------------------------
    Write-Section '6. 读写 GUI 控件文本（字符映射表）'
    $charmapProc = Start-Process charmap -PassThru
    $charmapWin = Wait-Window -ProcName 'charmap'
    if (-not $charmapWin) { throw '字符映射表窗口没有出现' }

    "窗口: [$($charmapWin.Title)]   类名 = $($charmapWin.ClassName)"
    $kids = @([Win32Ext]::GetChildren($charmapWin.Handle))
    "共 $($kids.Count) 个子控件，前 6 个："
    $kids | Select-Object -First 6 | ForEach-Object { ConvertTo-WindowObject $_ } |
        Format-Table -AutoSize | Out-String -Width 400

    $control = Get-TargetControl $charmapWin -ClassNamePattern 'RICHEDIT*'
    "选中的文本框控件: $($control.ClassName)   句柄 = 0x$('{0:X}' -f $control.Handle.ToInt64())"

    $payload = 'Hello 世界，PowerShell 通过 WM_SETTEXT 写入'
    [Win32Ext]::SetControlText($control.Handle, $payload)
    Start-Sleep -Milliseconds 300
    "写入后 WM_GETTEXT 读回: $([Win32Ext]::GetControlText($control.Handle))"

    # ------------------------------------------------------------------
    Write-Section '7. 鼠标点击 + 真实键盘输入到控件'
    $clickX = $control.Left + [int]($control.Width / 2)
    $clickY = $control.Top + [int]($control.Height / 2)
    [void][Win32Ext]::ForceForeground($charmapWin.Handle)
    [Win32Ext]::ClickAt($clickX, $clickY, $false, $false)
    Start-Sleep -Milliseconds 200
    [Win32Ext]::SendCombo('ctrl+End')   # 把光标移到文本末尾
    Start-Sleep -Milliseconds 150

    [Win32Ext]::TypeText(' + 键盘追加')
    Start-Sleep -Milliseconds 400
    "点击并输入后读回: $([Win32Ext]::GetControlText($control.Handle))"

    # ------------------------------------------------------------------
    Write-Section '8. 演示结束'
    '演示完成，下面关闭演示窗口。'
}
finally {
    if (-not $KeepOpen) {
        if ($charmapWin) { [Win32Ext]::CloseWindow($charmapWin.Handle) }
        if ($consoleWin) {
            [void][Win32Ext]::ForceForeground($consoleWin.Handle)
            Start-Sleep -Milliseconds 200
            [Win32Ext]::TypeText('exit')
            [Win32Ext]::SendCombo('enter')
            Start-Sleep -Milliseconds 600
        }
        foreach ($proc in @($charmapProc, $consoleProc)) {
            if ($proc) {
                try {
                    if (-not $proc.HasExited) { $proc.Kill() }
                } catch { }
            }
        }
        '演示窗口已关闭。'
    } else {
        '已保留演示窗口（-KeepOpen）。'
    }
}
