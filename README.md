# AutoCAD .NET 8 Event Monitor with Named Pipes IPC

A .NET 8 solution demonstrating real-time AutoCAD event monitoring using Named Pipes for inter-process communication.

## Architecture

```
┌────────────────────────────────┐
│  AutoCAD Process               │
│  ┌──────────────────────────┐  │
│  │ InteropFromAcadAddin     │  │
│  │ - Monitors AutoCAD events│  │
│  │ - Broadcasts via Pipes   │  │
│  └──────────────────────────┘  │
└────────────────────────────────┘
          │ Named Pipe
          │ (\\.\pipe\AutoCADEvents)
          ↓
┌────────────────────────────────┐
│  External .NET 8 Client        │
│  AcadDocEventsTester.exe       │
│  - Receives events             │
│  - Displays in console         │
└────────────────────────────────┘
```

## Projects

- **InteropFromAcadAddin** - AutoCAD plugin (.NET 8, NETLOAD into AutoCAD)
- **AcadDocEventsTester** - External client (.NET 8 console app)

## Quick Start

### 1. Build

```bash
dotnet build -c Debug -p:Platform=x64
```

### 2. Load Plugin in AutoCAD 2026

```
Command: NETLOAD
Browse to: InteropFromAcadAddin\bin\x64\Debug\InteropFromAcadAddin.dll
```

### 3. Run Event Monitor

```bash
AcadDocEventsTester\bin\Debug\AcadDocEventsTester.exe
```

### 4. Test

Run any commands in AutoCAD (LINE, CIRCLE, etc.) and watch events appear in the console!

## Features

- ✅ .NET 8 with modern SDK-style projects
- ✅ Named Pipes IPC (no COM registration needed)
- ✅ Real-time event streaming
- ✅ Cross-process communication
- ✅ 24+ AutoCAD document events monitored

## Events Monitored

- **Command Events**: Started, Ended, Cancelled, Failed
- **LISP Events**: Started, Ended, Cancelled  
- **Document Events**: Created, Activated, Closed, Destroyed
- **Lock Mode Events**: Changed, Vetoed

## Requirements

- AutoCAD 2025+ (.NET 8 support)
- .NET 8.0 SDK
- Visual Studio 2022 (v17.10+)
- Windows (Named Pipes)

## Key Files

- `PluginEntry.cs` - AutoCAD plugin entry point
- `EventBroadcaster.cs` - Named Pipes server
- `EventMessage.cs` - Event serialization
- `ProgramNamedPipes.cs` - Named Pipes client

## Why Named Pipes?

.NET 8 COM interop has limitations. Named Pipes provide:
- No registration needed
- Clean .NET-to-.NET communication
- Same-machine, cross-process
- Simple and reliable

## License

MIT
