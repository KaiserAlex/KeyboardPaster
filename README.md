# KeyboardPaster

A lightweight PowerShell tool that runs in the background and types clipboard content character by character using simulated keystrokes — triggered by **Ctrl+Shift+V**.

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
3. Press **Ctrl+Shift+V** — the text is typed out character by character
4. A tray notification confirms completion

### Parameters

| Parameter  | Default | Range    | Description                                 |
|------------|---------|----------|---------------------------------------------|
| `-DelayMs` | `30`    | `0–1000` | Delay in milliseconds between each keystroke |

### Stopping

- **Right-click** the system tray icon → *Exit KeyboardPaster*
- Or press **Ctrl+C** in the terminal

## How It Works

1. Registers a global **Ctrl+Shift+V** hotkey via the Win32 `RegisterHotKey` API
2. Runs a hidden window with a system tray icon to listen for the hotkey
3. On hotkey press, reads the clipboard and sends each character via `SendKeys`
4. Handles special characters (`+`, `^`, `%`, `~`, `{`, `}`, etc.), newlines, and tabs
5. Provides tray balloon notifications for status feedback

## Requirements

- Windows PowerShell 5.1+ or PowerShell 7+ (Windows only)
- .NET `System.Windows.Forms` assembly (included with Windows)

## License

MIT
