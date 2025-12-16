# ⚠️ NET8 Branch - COM Interop Migration (Not Recommended)

This branch contains the initial .NET 8 migration attempt using COM interop.

## ❌ Known Issues

- **Event Communication Broken**: External clients cannot reliably receive events from the AutoCAD plugin
- **COM Activation Issues**: `GetInterfaceObject` and `Activator.CreateInstance` create separate instances
- **Singleton Pattern Fails**: Static events don't properly share state across COM-created instances
- **.NET 8 COM Limitations**: COM interop in .NET 8 has stricter requirements and behavioral changes

## ✅ Working Solution

**Use the `feature/namedpipes-ipc` branch instead!**

```bash
git checkout feature/namedpipes-ipc
```

The Named Pipes implementation provides:
- ✅ Reliable inter-process communication
- ✅ No COM registration needed
- ✅ Works perfectly with .NET 8
- ✅ All events properly broadcast
- ✅ Clean, modern architecture

## Why This Branch Exists

This branch is kept for reference to show:
1. The challenges of migrating COM interop to .NET 8
2. Why COM is not recommended for .NET 8 IPC scenarios
3. The evolution from COM to Named Pipes

## Summary

- **This branch (NET8)**: COM-based migration with communication issues ❌
- **feature/namedpipes-ipc branch**: Named Pipes IPC - fully working ✅

For production use, always use the **feature/namedpipes-ipc** branch.

