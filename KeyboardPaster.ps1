<#
.SYNOPSIS
    Background tool that types clipboard content via configurable hotkey (default: Ctrl+Shift+V).

.DESCRIPTION
    Runs in the background with a system tray icon. On first launch it adds itself to
    Windows autostart. Press the configured hotkey to type clipboard text character by
    character instead of pasting. Right-click the tray icon to toggle autostart or
    change the hotkey.

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
    [int]$DelayMs = 30,

    [switch]$Background
)

# ── Relaunch as hidden background process if needed ─────────────────────────
if (-not $Background) {
    $args = @('-WindowStyle', 'Hidden', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath, '-Background')
    if ($PSBoundParameters.ContainsKey('DelayMs')) { $args += '-DelayMs'; $args += $DelayMs }
    Start-Process powershell.exe -ArgumentList $args -WindowStyle Hidden
    exit
}

# ── Assemblies ──────────────────────────────────────────────────────────────
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ── C# Types ───────────────────────────────────────────────────────────────
$refs = @(
    [System.Windows.Forms.Form].Assembly.Location,
    [System.Drawing.Graphics].Assembly.Location
)

Add-Type -TypeDefinition @"
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// Hidden window that receives global hotkey messages
public class KeyboardPasterForm : Form
{
    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern bool UnregisterHotKey(IntPtr hWnd, int id);

    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 9001;

    public event EventHandler HotkeyPressed;

    public KeyboardPasterForm()
    {
        this.FormBorderStyle = FormBorderStyle.None;
        this.ShowInTaskbar = false;
        this.WindowState = FormWindowState.Minimized;
    }

    public bool RegisterHotkeyCombo(uint modifiers, uint vk)
    {
        return RegisterHotKey(this.Handle, HOTKEY_ID, modifiers, vk);
    }

    public void UnregisterHotkeyCombo()
    {
        UnregisterHotKey(this.Handle, HOTKEY_ID);
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
        if (disposing) UnregisterHotKey(this.Handle, HOTKEY_ID);
        base.Dispose(disposing);
    }
}

// Dialog for capturing a new hotkey combination
public class HotkeyDialog : Form
{
    private TextBox txtHotkey;
    private Button btnOK;

    public uint ResultMod { get; private set; }
    public uint ResultVk { get; private set; }
    public string ResultDisplay { get; private set; }
    public bool HasResult { get; private set; }

    public HotkeyDialog(string currentDisplay)
    {
        this.Text = "Change Hotkey";
        this.ClientSize = new Size(310, 145);
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.StartPosition = FormStartPosition.CenterScreen;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.TopMost = true;

        var lblCurrent = new Label
        {
            Text = "Current:  " + currentDisplay,
            Location = new Point(15, 15),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        var lblInstr = new Label
        {
            Text = "Press new key combination:",
            Location = new Point(15, 40),
            AutoSize = true,
            Font = new Font("Segoe UI", 9f)
        };
        txtHotkey = new TextBox
        {
            Location = new Point(15, 63),
            Size = new Size(280, 28),
            ReadOnly = true,
            TextAlign = HorizontalAlignment.Center,
            Font = new Font("Segoe UI", 10f, FontStyle.Bold)
        };
        txtHotkey.KeyDown += OnCaptureKey;

        btnOK = new Button
        {
            Text = "OK",
            Location = new Point(134, 105),
            Size = new Size(80, 28),
            DialogResult = DialogResult.OK,
            Enabled = false
        };
        var btnCancel = new Button
        {
            Text = "Cancel",
            Location = new Point(218, 105),
            Size = new Size(80, 28),
            DialogResult = DialogResult.Cancel
        };

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
        this.Controls.AddRange(new Control[] { lblCurrent, lblInstr, txtHotkey, btnOK, btnCancel });
    }

    private void OnCaptureKey(object sender, KeyEventArgs e)
    {
        e.SuppressKeyPress = true;

        Keys kc = e.KeyCode;
        if (kc == Keys.ControlKey || kc == Keys.ShiftKey || kc == Keys.Menu ||
            kc == Keys.LControlKey || kc == Keys.RControlKey ||
            kc == Keys.LShiftKey || kc == Keys.RShiftKey ||
            kc == Keys.LMenu || kc == Keys.RMenu)
            return;

        if (e.Modifiers == Keys.None) return;

        var parts = new List<string>();
        uint mod = 0x4000; // MOD_NOREPEAT
        if (e.Control) { parts.Add("Ctrl"); mod |= 0x0002; }
        if (e.Shift)   { parts.Add("Shift"); mod |= 0x0004; }
        if (e.Alt)     { parts.Add("Alt"); mod |= 0x0001; }
        parts.Add(kc.ToString());

        ResultMod = mod;
        ResultVk = (uint)kc;
        ResultDisplay = string.Join("+", parts);
        HasResult = true;

        txtHotkey.Text = ResultDisplay;
        btnOK.Enabled = true;
    }
}
"@ -ReferencedAssemblies $refs

# ── Settings ────────────────────────────────────────────────────────────────
$script:SettingsDir  = Join-Path $env:APPDATA "KeyboardPaster"
$script:SettingsFile = Join-Path $script:SettingsDir "settings.json"

function Get-DefaultSettings {
    @{
        AutoStart     = $true
        HotkeyMod     = 0x4006   # MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT
        HotkeyVk      = 0x56     # VK_V
        HotkeyDisplay = "Ctrl+Shift+V"
        DelayMs       = $DelayMs
    }
}

function Load-Settings {
    $defaults = Get-DefaultSettings
    if (Test-Path $script:SettingsFile) {
        try {
            $json = Get-Content $script:SettingsFile -Raw | ConvertFrom-Json
            $json.PSObject.Properties | ForEach-Object {
                if ($defaults.ContainsKey($_.Name)) { $defaults[$_.Name] = $_.Value }
            }
        } catch { }
    }
    return $defaults
}

function Save-Settings([hashtable]$s) {
    if (-not (Test-Path $script:SettingsDir)) {
        New-Item -Path $script:SettingsDir -ItemType Directory -Force | Out-Null
    }
    $s | ConvertTo-Json | Set-Content -Path $script:SettingsFile -Encoding UTF8
}

# ── Autostart (Registry) ───────────────────────────────────────────────────
$script:RegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$script:RegName = "KeyboardPaster"

function Set-Autostart([bool]$Enabled) {
    if ($Enabled) {
        $scriptPath = $PSCommandPath
        $value = "powershell.exe -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Background"
        Set-ItemProperty -Path $script:RegPath -Name $script:RegName -Value $value
    } else {
        Remove-ItemProperty -Path $script:RegPath -Name $script:RegName -ErrorAction SilentlyContinue
    }
}

function Get-AutostartEnabled {
    $null -ne (Get-ItemProperty -Path $script:RegPath -Name $script:RegName -ErrorAction SilentlyContinue)
}

# ── Character sending ──────────────────────────────────────────────────────
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

# ── Custom tray icon ───────────────────────────────────────────────────────
function New-TrayIcon {
    $bmp = New-Object System.Drawing.Bitmap(16, 16)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.Clear([System.Drawing.Color]::FromArgb(50, 120, 200))
    $font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $g.DrawString("V", $font, [System.Drawing.Brushes]::White, 1, 0)
    $font.Dispose()
    $g.Dispose()
    $icon = [System.Drawing.Icon]::FromHandle($bmp.GetHicon())
    $bmp.Dispose()
    return $icon
}

# ── Main ────────────────────────────────────────────────────────────────────
$script:Settings = Load-Settings
if ($PSBoundParameters.ContainsKey('DelayMs')) { $script:Settings.DelayMs = $DelayMs }
$script:IsTyping = $false
$script:FirstRun = -not (Test-Path $script:SettingsFile)

# Persist defaults & enable autostart on first run
if ($script:FirstRun) {
    Save-Settings $script:Settings
    Set-Autostart $true
}

# Sync autostart registry with settings
if ($script:Settings.AutoStart -and -not (Get-AutostartEnabled)) { Set-Autostart $true }
elseif (-not $script:Settings.AutoStart -and (Get-AutostartEnabled)) { Set-Autostart $false }

# Hidden form
$form = New-Object KeyboardPasterForm

# Register hotkey from settings
$hotkeyMod = [uint32]$script:Settings.HotkeyMod
$hotkeyVk  = [uint32]$script:Settings.HotkeyVk
if (-not $form.RegisterHotkeyCombo($hotkeyMod, $hotkeyVk)) {
    Write-Host "ERROR: Could not register hotkey $($script:Settings.HotkeyDisplay). It may already be in use." -ForegroundColor Red
    exit 1
}

# ── Tray icon & context menu ───────────────────────────────────────────────
$trayIcon = New-Object System.Windows.Forms.NotifyIcon
$trayIcon.Icon = New-TrayIcon
$trayIcon.Text = "KeyboardPaster`n$($script:Settings.HotkeyDisplay)"
$trayIcon.Visible = $true

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip

# Menu: Autostart toggle
$autoItem = New-Object System.Windows.Forms.ToolStripMenuItem("Autostart")
$autoItem.Checked = [bool]$script:Settings.AutoStart
$autoItem.add_Click({
    $script:Settings.AutoStart = -not $script:Settings.AutoStart
    $autoItem.Checked = $script:Settings.AutoStart
    Set-Autostart $script:Settings.AutoStart
    Save-Settings $script:Settings
})

# Menu: Change hotkey
$hotkeyItem = New-Object System.Windows.Forms.ToolStripMenuItem("Change Hotkey...")
$hotkeyItem.add_Click({
    # Unregister current hotkey so the capture dialog can use all keys
    $form.UnregisterHotkeyCombo()

    $dlg = New-Object HotkeyDialog($script:Settings.HotkeyDisplay)
    $result = $dlg.ShowDialog()

    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and $dlg.HasResult) {
        if ($form.RegisterHotkeyCombo($dlg.ResultMod, $dlg.ResultVk)) {
            $script:Settings.HotkeyMod     = [int]$dlg.ResultMod
            $script:Settings.HotkeyVk      = [int]$dlg.ResultVk
            $script:Settings.HotkeyDisplay  = $dlg.ResultDisplay
            Save-Settings $script:Settings
            $trayIcon.Text = "KeyboardPaster`n$($dlg.ResultDisplay)"
            $trayIcon.ShowBalloonTip(2000, "KeyboardPaster",
                "Hotkey changed: $($dlg.ResultDisplay)",
                [System.Windows.Forms.ToolTipIcon]::Info)
        } else {
            $form.RegisterHotkeyCombo([uint32]$script:Settings.HotkeyMod, [uint32]$script:Settings.HotkeyVk)
            $trayIcon.ShowBalloonTip(2000, "KeyboardPaster",
                "'$($dlg.ResultDisplay)' is already in use.",
                [System.Windows.Forms.ToolTipIcon]::Error)
        }
    } else {
        # Cancelled — re-register previous hotkey
        $form.RegisterHotkeyCombo([uint32]$script:Settings.HotkeyMod, [uint32]$script:Settings.HotkeyVk)
    }
    $dlg.Dispose()
})

# Menu: Separator + Exit
$separator = New-Object System.Windows.Forms.ToolStripSeparator
$exitItem  = New-Object System.Windows.Forms.ToolStripMenuItem("Exit")
$exitItem.add_Click({
    $trayIcon.Visible = $false
    [System.Windows.Forms.Application]::Exit()
})

$contextMenu.Items.AddRange(@($autoItem, $hotkeyItem, $separator, $exitItem))
$trayIcon.ContextMenuStrip = $contextMenu

# ── Hotkey handler ──────────────────────────────────────────────────────────
$form.add_HotkeyPressed({
    if ($script:IsTyping) { return }
    $script:IsTyping = $true

    try {
        $clipText = Get-Clipboard -Raw
        if ([string]::IsNullOrEmpty($clipText)) { return }

        # Wait for modifier keys to be released
        [System.Threading.Thread]::Sleep(300)

        foreach ($char in $clipText.ToCharArray()) {
            Send-Character $char
            if ($script:Settings.DelayMs -gt 0) {
                [System.Threading.Thread]::Sleep($script:Settings.DelayMs)
            }
        }
    }
    finally {
        $script:IsTyping = $false
    }
})

# ── Start ───────────────────────────────────────────────────────────────────
$startMsg = if ($script:FirstRun) {
    "First launch! Autostart enabled.`nPress $($script:Settings.HotkeyDisplay) to type clipboard."
} else {
    "Ready! Press $($script:Settings.HotkeyDisplay) to type clipboard."
}
$trayIcon.ShowBalloonTip(3000, "KeyboardPaster", $startMsg, [System.Windows.Forms.ToolTipIcon]::Info)

try {
    [System.Windows.Forms.Application]::Run($form)
}
finally {
    $trayIcon.Visible = $false
    $trayIcon.Dispose()
    $form.Dispose()
}
