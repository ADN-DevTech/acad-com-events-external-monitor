# AutoCAD Event Monitor via COM Interop

A .NET Framework project demonstrating how to expose AutoCAD document events through COM interop and consume them from an external console application.

## Overview

This solution consists of two projects:

1. **InteropFromAcadAddin** - AutoCAD add-in that exposes document events via COM
2. **AcadDocEventsTester** - Console application that subscribes to and displays AutoCAD events in real-time

## Features

- **24 Real-time AutoCAD Events** monitored:
  - Document Command Events (Start, End, Cancel, Fail, Unknown)
  - LISP Events (Start, End, Cancel)
  - Document Close Events
  - DocumentManager Events (Created, Activated, Destroyed, Lock Mode changes, etc.)
- Color-coded console output with timestamps
- COM connection point architecture bypassing ServicedComponent transparent proxy issues

## Prerequisites

- Visual Studio 2022
- AutoCAD 2024 (or adjust version in code)
- .NET Framework 4.8
- AutoCAD .NET API references

## Build Instructions

### 1. Build the Solution

```bash
# Open the solution in Visual Studio
# Build in Debug or Release mode (x64)
# Solution will output to:
# - InteropFromAcadAddin: bin\x64\Debug\InteropFromAcadAddin.dll
# - AcadDocEventsTester: bin\Debug\AcadDocEventsTester.exe
```

### 2. Register the AutoCAD Add-in DLL with RegAsm

Open **Command Prompt as Administrator** and navigate to your AutoCAD 2024 installation directory:

```cmd
cd "C:\Program Files\Autodesk\AutoCAD 2024"

regasm "<YouPath>\InteropFromAcadAddin\InteropFromAcadAddin\bin\x64\Debug\InteropFromAcadAddin.dll" /codebase /tlb
```

**Note:** Adjust paths according to your AutoCAD installation and project location.

### 3. Add Reference to Console Application (for IntelliSense)

In the **AcadDocEventsTester** project:

1. Right-click **References** → **Add Reference**
2. Click **Browse** and navigate to:
   ```
   <YouPath>\InteropFromAcadAddin\InteropFromAcadAddin\bin\x64\Debug\InteropFromAcadAddin.dll
   ```
3. Click **OK**

This provides IntelliSense for the `IDocumentEventService` and `IDocumentEventServiceEvents` interfaces.

## Running the Event Monitor

### Step 1: Load the Add-in in AutoCAD

1. Launch **AutoCAD 2024**
2. Type `NETLOAD` at the command prompt
3. Browse to and select:
   ```
   <YouPath>\InteropFromAcadAddin\InteropFromAcadAddin\bin\x64\Debug\InteropFromAcadAddin.dll
   ```
4. The add-in will register itself and start monitoring

### Step 2: Run the Console Application

1. Run **AcadDocEventsTester.exe** (from Visual Studio or directly)
2. The console will connect to the running AutoCAD instance
3. You should see:
   ```
   **AutoCAD Event Monitor**
   
   Connected to: AutoCAD 24.3
   Plugin service retrieved successfully.
   
   Service started - now monitoring AutoCAD events.
   
   Event subscription successful!
   
   --- Listening for AutoCAD events ---
   Run commands in AutoCAD to see events here.
   Press ENTER to stop monitoring...
   ```

### Step 3: Watch Events in Real-time

- Execute commands in AutoCAD (LINE, CIRCLE, MOVE, etc.)
- Open/close documents
- Switch between documents
- All events will appear in the console with color-coding and timestamps

**Example Output:**
```
[14:23:45.123]  CMD START: 'LINE' | Drawing1.dwg
[14:23:47.456]  CMD END: 'LINE' | Drawing1.dwg
[14:23:50.789]  DOC CREATED | Drawing2.dwg
[14:23:51.012]  DOC BECAME CURRENT | Drawing2.dwg
```

### Step 4: Stop Monitoring

- Press **ENTER** in the console application to disconnect and exit
- The AutoCAD add-in continues running until AutoCAD closes

## Project Structure

```
InteropFromAcadAddin/
├── InteropFromAcadAddin/           # AutoCAD Add-in Project
│   ├── PluginEntry.cs              # Main COM server + event implementation
│   ├── Properties/AssemblyInfo.cs  # COM visibility settings
│   └── scripts/
│       ├── Register-AcadAddin.bat  # RegAsm registration helper
│       └── Unregister-AcadAddin.bat
│
└── AcadDocEventsTester/            # Console Client Project
    ├── Program.cs                  # COM connection + vtable interop
    └── ComEventSink.cs             # Event sink implementation
```

## Extending Event Coverage

This implementation currently monitors **Document** and **DocumentManager** events. You can extend it to include:

- **Database Events** (ObjectAppended, ObjectModified, ObjectErased, etc.)
- **Editor Events** (PromptForSelection, etc.)
- **Layout Manager Events**
- **Plot Events**
- **Application Events**

**Reference Sample:** Check out the comprehensive `EventsWatcher` sample in the AutoCAD SDK:
```
<YouPath>\ARX2026\samples\dotNet\EventsWatcher
```

## Troubleshooting

### "Failed to retrieve plugin service"
- Ensure the DLL is registered with RegAsm (with `/codebase` flag)
- Verify the DLL is NETLOAD'ed in AutoCAD

### "Unable to cast transparent proxy..."
- This project uses direct vtable calls to bypass ServicedComponent proxy issues
- Ensure you're using the updated `Program.cs` with `FindConnectionPoint`, `Advise`, `Unadvise` methods

### COM Registration Issues
- Always run `regasm` as **Administrator**
- Use the `/codebase` flag to register the DLL location
- Verify bitness matches (x64 AutoCAD requires x64 DLL)

## Technical Details

- **COM Interface GUIDs:**
  - `IDocumentEventService`: `D1C762E4-2C22-482B-9333-82F3E906E6B4`
  - `IDocumentEventServiceEvents`: `A7E1B0A4-2E2E-4B6D-9E39-7D6C5B2C9A11`
  - `PluginEntry` CLSID: `75B37C6C-8047-41E3-8C4D-D5196EEFF368`

- **Architecture:**
  - Uses COM connection points for event marshaling
  - Direct vtable calls to bypass .NET Remoting transparent proxies
  - ServicedComponent base class for COM+ compatibility with AutoCAD

## License

This is a demonstration project for educational purposes.

## Video Demo

Watch the demonstration video showing the event monitor in action:

https://github.com/ADN-DevTech/acad-com-events-external-monitor/assets/Demo.mp4

**Or download:** [Demo.mp4](./docs/Demo.mp4)

The video demonstrates:
- Loading the AutoCAD add-in with NETLOAD
- Running the console application
- Real-time event monitoring as commands are executed
- Color-coded event output with timestamps
- Document lifecycle events (create, activate, close)

---

**Note:** Paths in this README are examples. Adjust them to match your actual project location and AutoCAD installation directory.
