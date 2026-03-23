# KeyboardPaster

A lightweight PowerShell tool that types clipboard content character by character using simulated keystrokes — instead of pasting.

## Why?

Some applications block `Ctrl+V` paste operations, detect clipboard paste events, or require input to come from actual keystrokes. **KeyboardPaster** bypasses these restrictions by sending each character as a virtual key press.

## Usage

```powershell
# Copy text to clipboard, then run:
.\KeyboardPaster.ps1

# Custom delay between keystrokes (in ms) and startup countdown (in seconds):
.\KeyboardPaster.ps1 -DelayMs 50 -StartDelaySec 5
```

### Parameters

| Parameter        | Default | Range     | Description                                      |
|------------------|---------|-----------|--------------------------------------------------|
| `-DelayMs`       | `30`    | `0–1000`  | Delay in milliseconds between each keystroke      |
| `-StartDelaySec` | `3`     | `1–30`    | Countdown before typing starts (switch to target) |

## How It Works

1. Reads text from the Windows clipboard
2. Shows a preview and character count
3. Counts down to give you time to focus the target window
4. Sends each character individually via `System.Windows.Forms.SendKeys`
5. Handles special characters, newlines, and tabs

## Requirements

- Windows PowerShell 5.1+ or PowerShell 7+ (Windows only)
- .NET `System.Windows.Forms` assembly (included with Windows)

## License

MIT
