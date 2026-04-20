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

function Get-HostExecutablePath {
    $procPath = (Get-Process -Id $PID -ErrorAction SilentlyContinue).Path
    if (-not [string]::IsNullOrWhiteSpace($procPath) -and (Test-Path $procPath)) {
        return $procPath
    }

    if ($PSVersionTable.PSEdition -eq 'Core') {
        return 'pwsh.exe'
    }
    return 'powershell.exe'
}

$script:HostExe = Get-HostExecutablePath

# ── Relaunch as hidden background process if needed ─────────────────────────
if (-not $Background) {
    $launchArgs = "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$PSCommandPath`" -Background"
    if ($PSBoundParameters.ContainsKey('DelayMs')) { $launchArgs += " -DelayMs $DelayMs" }
    Start-Process -FilePath $script:HostExe -ArgumentList $launchArgs -WindowStyle Hidden
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

    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);

    private const int WM_HOTKEY = 0x0312;
    private const int HOTKEY_ID = 9001;

    // Returns true if any modifier key (Ctrl, Shift, Alt) is currently held down
    public static bool AreModifiersHeld()
    {
        const int VK_SHIFT   = 0x10;
        const int VK_CONTROL = 0x11;
        const int VK_MENU    = 0x12; // Alt
        return (GetAsyncKeyState(VK_SHIFT)   & 0x8000) != 0 ||
               (GetAsyncKeyState(VK_CONTROL) & 0x8000) != 0 ||
               (GetAsyncKeyState(VK_MENU)    & 0x8000) != 0;
    }

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
        AutoStart      = $true
        HotkeyMod      = 0x4006   # MOD_CONTROL | MOD_SHIFT | MOD_NOREPEAT
        HotkeyVk       = 0x56     # VK_V
        HotkeyDisplay  = "Ctrl+Shift+V"
        DelayMs        = $DelayMs
        EnterAfterPaste = $false
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
        $value = "`"$script:HostExe`" -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$scriptPath`" -Background"
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

# Menu: Enter after Paste toggle
$enterItem = New-Object System.Windows.Forms.ToolStripMenuItem("Enter after Paste")
$enterItem.Checked = [bool]$script:Settings.EnterAfterPaste
$enterItem.add_Click({
    $script:Settings.EnterAfterPaste = -not $script:Settings.EnterAfterPaste
    $enterItem.Checked = $script:Settings.EnterAfterPaste
    Save-Settings $script:Settings
})

# Menu: Delay configuration
$delayItem = New-Object System.Windows.Forms.ToolStripMenuItem("Delay: $($script:Settings.DelayMs)ms")
$delayItem.add_Click({
    $inputForm = New-Object System.Windows.Forms.Form
    $inputForm.Text = "Keystroke Delay"
    $inputForm.ClientSize = New-Object System.Drawing.Size(280, 120)
    $inputForm.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
    $inputForm.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterScreen
    $inputForm.MaximizeBox = $false
    $inputForm.MinimizeBox = $false
    $inputForm.TopMost = $true

    $lbl = New-Object System.Windows.Forms.Label
    $lbl.Text = "Delay between keystrokes (0-1000 ms):"
    $lbl.Location = New-Object System.Drawing.Point(15, 15)
    $lbl.AutoSize = $true
    $lbl.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $nud = New-Object System.Windows.Forms.NumericUpDown
    $nud.Location = New-Object System.Drawing.Point(15, 42)
    $nud.Size = New-Object System.Drawing.Size(248, 28)
    $nud.Minimum = 0
    $nud.Maximum = 1000
    $nud.Value = $script:Settings.DelayMs
    $nud.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)

    $btnOK = New-Object System.Windows.Forms.Button
    $btnOK.Text = "OK"
    $btnOK.Location = New-Object System.Drawing.Point(113, 82)
    $btnOK.Size = New-Object System.Drawing.Size(75, 28)
    $btnOK.DialogResult = [System.Windows.Forms.DialogResult]::OK

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Location = New-Object System.Drawing.Point(192, 82)
    $btnCancel.Size = New-Object System.Drawing.Size(75, 28)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

    $inputForm.AcceptButton = $btnOK
    $inputForm.CancelButton = $btnCancel
    $inputForm.Controls.AddRange(@($lbl, $nud, $btnOK, $btnCancel))

    if ($inputForm.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $script:Settings.DelayMs = [int]$nud.Value
        Save-Settings $script:Settings
        $delayItem.Text = "Delay: $($script:Settings.DelayMs)ms"
    }
    $inputForm.Dispose()
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

$contextMenu.Items.AddRange(@($autoItem, $enterItem, $delayItem, $hotkeyItem, $separator, $exitItem))
$trayIcon.ContextMenuStrip = $contextMenu

# ── Hotkey handler ──────────────────────────────────────────────────────────
$form.add_HotkeyPressed({
    if ($script:IsTyping) { return }
    $script:IsTyping = $true

    try {
        $clipText = Get-Clipboard -Raw
        if ([string]::IsNullOrEmpty($clipText)) { return }

        # Wait until modifier keys are actually released
        $timeout = 50
        while ([KeyboardPasterForm]::AreModifiersHeld() -and $timeout -gt 0) {
            [System.Threading.Thread]::Sleep(20)
            $timeout--
        }
        [System.Threading.Thread]::Sleep(50)

        foreach ($char in $clipText.ToCharArray()) {
            Send-Character $char
            if ($script:Settings.DelayMs -gt 0) {
                [System.Threading.Thread]::Sleep($script:Settings.DelayMs)
            }
        }

        if ($script:Settings.EnterAfterPaste) {
            [System.Windows.Forms.SendKeys]::SendWait("{ENTER}")
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
