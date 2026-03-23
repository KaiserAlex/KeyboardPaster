<#
.SYNOPSIS
    Background tool that types clipboard content via Ctrl+Shift+V using simulated keystrokes.

.DESCRIPTION
    Runs in the background with a system tray icon and registers a global Ctrl+Shift+V hotkey.
    When pressed, it types the clipboard text character by character instead of pasting —
    useful for applications that block paste or detect clipboard events.

.PARAMETER DelayMs
    Delay in milliseconds between each keystroke. Default: 30ms.

.EXAMPLE
    .\KeyboardPaster.ps1
    Starts the background listener with default settings.

.EXAMPLE
    .\KeyboardPaster.ps1 -DelayMs 50
    Starts with 50ms delay between keystrokes.
#>

param(
    [ValidateRange(0, 1000)]
    [int]$DelayMs = 30
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$winFormsAsm = [System.Windows.Forms.Form].Assembly.Location
Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;
using System.Runtime.InteropServices;

public class KeyboardPasterForm : Form
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private const uint MOD_CONTROL = 0x0002;
    private const uint MOD_SHIFT = 0x0004;
    private const uint MOD_NOREPEAT = 0x4000;
    private const uint VK_V = 0x56;
    private const int HOTKEY_ID = 9001;

    public event EventHandler HotkeyPressed;

    public KeyboardPasterForm()
    {
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Minimized;
    }

    public bool RegisterPasteHotkey()
    {
        return RegisterHotKey(this.Handle, HOTKEY_ID, MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT, VK_V);
    }

    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_HOTKEY && m.WParam.ToInt32() == HOTKEY_ID)
        {
            var handler = HotkeyPressed;
            if (handler != null) handler(this, EventArgs.Empty);
        }
        base.WndProc(ref m);
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            UnregisterHotKey(this.Handle, HOTKEY_ID);
        }
        base.Dispose(disposing);
    }
}
"@ -ReferencedAssemblies $winFormsAsm

function Send-Character([char]$c) {
    switch ($c) {
        "`r" { return }
        "`n" { [System.Windows.Forms.SendKeys]::SendWait("{ENTER}"); return }
        "`t" { [System.Windows.Forms.SendKeys]::SendWait("{TAB}"); return }
        default {
            $escaped = switch ($c) {
                '+' { '{+}' }
                '^' { '{^}' }
                '%' { '{%}' }
                '~' { '{~}' }
                '{' { '{{}' }
                '}' { '{}}' }
                '(' { '{(}' }
                ')' { '{)}' }
                default { $c.ToString() }
            }
            [System.Windows.Forms.SendKeys]::SendWait($escaped)
        }
    }
}

# Store delay in script scope for event handler access
$script:TypeDelayMs = $DelayMs
$script:IsTyping = $false

# Hidden form for hotkey messages
$form = New-Object KeyboardPasterForm

# System tray icon
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = [System.Drawing.SystemIcons]::Application
$trayIcon.Text = "KeyboardPaster — Ctrl+Shift+V"
$trayIcon.Visible = $true

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$exitItem = $contextMenu.Items.Add("Exit KeyboardPaster")
$exitItem.add_Click({
    $trayIcon.Visible = $false
    [System.Windows.Forms.Application]::Exit()
})
$trayIcon.ContextMenuStrip = $contextMenu

# Register global hotkey
if (-not $form.RegisterPasteHotkey()) {
    $trayIcon.Dispose()
    Write-Host "ERROR: Could not register Ctrl+Shift+V. The hotkey may already be in use." -ForegroundColor Red
    exit 1
}

# Hotkey handler — type clipboard content
$form.add_HotkeyPressed({
    if ($script:IsTyping) { return }
    $script:IsTyping = $true

    try {
        $clipText = Get-Clipboard -Raw
        if ([string]::IsNullOrEmpty($clipText)) {
            $trayIcon.ShowBalloonTip(2000, "KeyboardPaster", "Clipboard is empty.", [System.Windows.Forms.ToolTipIcon]::Warning)
            return
        }

        # Wait for modifier keys to be released
        [System.Threading.Thread]::Sleep(300)

        foreach ($char in $clipText.ToCharArray()) {
            Send-Character $char
            if ($script:TypeDelayMs -gt 0) {
                [System.Threading.Thread]::Sleep($script:TypeDelayMs)
            }
        }

        $trayIcon.ShowBalloonTip(1500, "KeyboardPaster", "Typed $($clipText.Length) characters.", [System.Windows.Forms.ToolTipIcon]::Info)
    }
    finally {
        $script:IsTyping = $false
    }
})

# Startup
$trayIcon.ShowBalloonTip(3000, "KeyboardPaster", "Press Ctrl+Shift+V to type clipboard content.", [System.Windows.Forms.ToolTipIcon]::Info)
Write-Host "KeyboardPaster is running." -ForegroundColor Green
Write-Host "Press Ctrl+Shift+V in any window to type clipboard content." -ForegroundColor Cyan
Write-Host "Right-click the tray icon or press Ctrl+C to exit." -ForegroundColor DarkGray

try {
    [System.Windows.Forms.Application]::Run($form)
}
finally {
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    $form.Dispose()
    Write-Host "`nKeyboardPaster stopped." -ForegroundColor Yellow
}
