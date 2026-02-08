# PortWatcher

PortWatcher is a lightweight Windows **system tray** tool for monitoring **TCP LISTEN ports** in real time.

It runs quietly in the background and notifies you when ports are opened or released.  
Useful for developers running local servers or debugging port conflicts.

---

## Features

- Monitor TCP LISTEN ports in real time
- Tray notifications for port open / close events
- Snapshot view of current listening ports
- Language switch: Chinese / English
- Sorting by Port / PID / Process name
- Runs as a background tray application
- Single-file executable (no .NET installation required)

---

## Usage

1. Download `PortWatcher.exe` from **GitHub Releases**
2. Double-click to run

The app runs in the **system tray** (bottom-right corner).  
If you don’t see it, click the `^` arrow to show hidden tray icons.

### Tray actions

- Right-click tray icon:
  - Scan now
  - Show snapshot
  - Change language
  - Change sorting
  - Exit
- Double-click tray icon:
  - Show snapshot

---

## For Developers

Requirements:
- Windows
- .NET SDK 8.0+

Run from source:

```bash
dotnet restore
dotnet run
```

Build executable:
```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```
---
