# KeyboardPaster

A lightweight PowerShell tool that runs in the background and types clipboard content character by character using simulated keystrokes — triggered by a configurable hotkey (default: **Ctrl+Shift+V**).

## Why?

Some applications block `Ctrl+V` paste operations, detect clipboard paste events, or require input to come from actual keystrokes. **KeyboardPaster** bypasses these restrictions by sending each character as a virtual key press.

## Usage

```powershell
# Start the background listener:
.\KeyboardPaster.ps1

# With custom delay between keystrokes:
.\KeyboardPaster.ps1 -DelayMs 50
```

Once running:

1. Copy text with `Ctrl+C` as usual
2. Click into the target application
3. Press **Ctrl+Shift+V** (or your custom hotkey) — the text is typed out character by character
4. A tray notification confirms completion

### Parameters

| Parameter  | Default | Range    | Description                                 |
|------------|---------|----------|---------------------------------------------|
| `-DelayMs` | `30`    | `0–1000` | Delay in milliseconds between each keystroke |

### System Tray Menu (right-click)

| Entry                | Description                                                       |
|----------------------|-------------------------------------------------------------------|
| **Autostart**        | Toggle Windows autostart on/off (enabled by default on first run) |
| **Enter after Paste**| Toggle sending an Enter keystroke after pasting is complete        |
| **Delay: Xms**       | Opens a dialog to change the keystroke delay at runtime (0–1000)  |
| **Change Hotkey…**   | Opens a dialog to capture a new keyboard shortcut                 |
| **Exit**             | Stops KeyboardPaster                                              |

### Stopping

- **Right-click** the system tray icon → *Exit*

## How It Works

1. On first launch, adds itself to Windows autostart (`HKCU\...\Run` registry)
2. Automatically relaunches as a hidden background process (no console window)
3. Registers a global hotkey via the Win32 `RegisterHotKey` API
3. Runs a hidden window with a custom system tray icon to listen for the hotkey
4. On hotkey press, reads the clipboard and sends each character via `SendKeys`
5. Handles special characters (`+`, `^`, `%`, `~`, `{`, `}`, etc.), newlines, and tabs
6. Settings (hotkey, autostart) are persisted in `%APPDATA%\KeyboardPaster\settings.json`

## Requirements

- Windows PowerShell 5.1+ or PowerShell 7+ (Windows only)
- .NET `System.Windows.Forms` assembly (included with Windows)

## License

MIT
