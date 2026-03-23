<#
.SYNOPSIS
    Simulates typing clipboard text character by character via virtual keystrokes.

.DESCRIPTION
    Instead of pasting clipboard content directly, this tool sends each character
    as an individual keystroke using SendKeys. This is useful for applications that
    block paste operations or detect clipboard paste events.

.PARAMETER DelayMs
    Delay in milliseconds between each keystroke. Default: 30ms.

.PARAMETER StartDelaySec
    Countdown in seconds before typing starts, giving you time to focus the target window. Default: 3 seconds.

.EXAMPLE
    .\KeyboardPaster.ps1
    Types clipboard content with default settings.

.EXAMPLE
    .\KeyboardPaster.ps1 -DelayMs 50 -StartDelaySec 5
    Types with 50ms delay between keys and 5 second countdown.
#>

param(
    [ValidateRange(0, 1000)]
    [int]$DelayMs = 30,

    [ValidateRange(1, 30)]
    [int]$StartDelaySec = 3
)

Add-Type -AssemblyName System.Windows.Forms

# Read clipboard content
$clipText = Get-Clipboard -Raw
if ([string]::IsNullOrEmpty($clipText)) {
    Write-Host "Clipboard is empty or contains no text." -ForegroundColor Red
    exit 1
}

# Preview
$preview = if ($clipText.Length -gt 100) { $clipText.Substring(0, 100) + "..." } else { $clipText }
$lineCount = ($clipText -split "`n").Count
Write-Host "=== KeyboardPaster ===" -ForegroundColor Cyan
Write-Host "Characters: $($clipText.Length) | Lines: $lineCount | Delay: ${DelayMs}ms"
Write-Host "Preview: $preview" -ForegroundColor DarkGray
Write-Host ""

# Countdown
Write-Host "Typing starts in $StartDelaySec seconds — focus the target window now!" -ForegroundColor Yellow
for ($i = $StartDelaySec; $i -gt 0; $i--) {
    Write-Host " $i..." -NoNewline
    Start-Sleep -Seconds 1
}
Write-Host " Go!" -ForegroundColor Green

# Escape a character for SendKeys syntax
function ConvertTo-SendKeysChar([char]$c) {
    switch ($c) {
        '+' { return '{+}' }
        '^' { return '{^}' }
        '%' { return '{%}' }
        '~' { return '{~}' }
        '{' { return '{{}' }
        '}' { return '{}}' }
        '(' { return '{(}' }
        ')' { return '{)}' }
        default { return $c.ToString() }
    }
}

# Type each character
$typed = 0
foreach ($char in $clipText.ToCharArray()) {
    switch ($char) {
        "`r" {
            # Skip carriage return; newline is handled by `n
            continue
        }
        "`n" {
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
        }
        "`t" {
            [System.Windows.Forms.SendKeys]::SendWait("{TAB}")
        }
        default {
            $escaped = ConvertTo-SendKeysChar $char
            [System.Windows.Forms.SendKeys]::SendWait($escaped)
        }
    }

    $typed++
    if ($DelayMs -gt 0) {
        Start-Sleep -Milliseconds $DelayMs
    }
}

Write-Host ""
Write-Host "Done! Typed $typed characters." -ForegroundColor Green
