using InteropFromAcadAddin;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AcadDocEventsTester
{
    class Program
    {
        const string progId = "AutoCAD.Application.25.1";  // AutoCAD 2026

        static void Main(string[] args)
        {
            Console.WriteLine("**AutoCAD Event Monitor (.NET 8)**\n");

            // Get AutoCAD Application using custom GetActiveObject (Marshal.GetActiveObject not available in .NET 8)
            dynamic? acadApp = ComUtils.GetActiveObject(progId);
            
            if (acadApp == null)
            {
                Console.WriteLine($"ERROR: AutoCAD is not running or ProgID '{progId}' is incorrect.");
                Console.WriteLine("Please launch AutoCAD 2026 first.");
                Console.WriteLine("\nPress ENTER to exit...");
                Console.ReadLine();
                return;
            }

            Console.WriteLine($"Connected to: {acadApp.Name} {acadApp.Version}");

            // Try direct COM activation instead of GetInterfaceObject
            // This works better with .NET 8 COM hosting
            object? pluginObj = null;
            try
            {
                Type? comType = Type.GetTypeFromProgID("InteropFromAcadAddin.IDocumentEventService");
                if (comType != null)
                {
                    pluginObj = Activator.CreateInstance(comType);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Direct COM activation failed: {ex.Message}");
                Console.WriteLine("Trying GetInterfaceObject...");
                pluginObj = acadApp.GetInterfaceObject("InteropFromAcadAddin.IDocumentEventService");
            }

            if (pluginObj != null)
            {
                Console.WriteLine("Plugin service retrieved successfully.\n");

                // NOTE: Don't call Start() here - it's already called when the plugin loads in AutoCAD
                // Calling Start() from external client would try to load AutoCAD assemblies in this process
                Console.WriteLine("Service is already running in AutoCAD process.\n");

                IntPtr pUnknown = IntPtr.Zero;
                IntPtr pContainer = IntPtr.Zero;
                IntPtr pConnectionPoint = IntPtr.Zero;

                try
                {
                    // Get the IUnknown pointer
                    pUnknown = Marshal.GetIUnknownForObject(pluginObj);

                    // Query for IConnectionPointContainer
                    Guid containerGuid = typeof(IConnectionPointContainer).GUID;
                    int hr = Marshal.QueryInterface(pUnknown, ref containerGuid, out pContainer);
                    Marshal.ThrowExceptionForHR(hr);

                    // Manually call FindConnectionPoint using vtable
                    Guid eventGuid = new Guid("A7E1B0A4-2E2E-4B6D-9E39-7D6C5B2C9A11");
                    pConnectionPoint = FindConnectionPoint(pContainer, eventGuid);

                    if (pConnectionPoint != IntPtr.Zero)
                    {
                        // Advise the sink
                        ComEventSink sink = new ComEventSink();
                        int cookie = Advise(pConnectionPoint, sink);

                        Console.WriteLine("Event subscription successful!");
                        Console.WriteLine("\n--- Listening for AutoCAD events ---");
                        Console.WriteLine("Run commands in AutoCAD to see events here.");
                        Console.WriteLine("Press ENTER to stop monitoring...\n");
                        
                        Console.ReadLine();

                        // Cleanup
                        Console.WriteLine("\nStopping event monitoring...");
                        Unadvise(pConnectionPoint, cookie);
                        Console.WriteLine("Disconnected from events.");
                        
                        // NOTE: Don't call Stop() - the service keeps running in AutoCAD
                    }
                }
                finally
                {
                    if (pConnectionPoint != IntPtr.Zero) Marshal.Release(pConnectionPoint);
                    if (pContainer != IntPtr.Zero) Marshal.Release(pContainer);
                    if (pUnknown != IntPtr.Zero) Marshal.Release(pUnknown);
                }
            }
            else
            {
                Console.WriteLine("ERROR: Failed to retrieve plugin service.");
                Console.WriteLine("Make sure the AutoCAD add-in is properly loaded.");
            }

            Console.WriteLine("\nPress ENTER to exit...");
            Console.ReadLine();
        }

        private static IntPtr FindConnectionPoint(IntPtr pContainer, Guid riid)
        {
            // IConnectionPointContainer::FindConnectionPoint is at vtable offset 4 (3rd method after IUnknown)
            IntPtr vtable = Marshal.ReadIntPtr(pContainer);
            IntPtr pFindConnectionPoint = Marshal.ReadIntPtr(vtable, IntPtr.Size * 4);

            var findConnectionPoint = Marshal.GetDelegateForFunctionPointer<FindConnectionPointDelegate>(pFindConnectionPoint);
            
            IntPtr pConnectionPoint = IntPtr.Zero;
            int hr = findConnectionPoint(pContainer, ref riid, out pConnectionPoint);
            Marshal.ThrowExceptionForHR(hr);

            return pConnectionPoint;
        }

        private static int Advise(IntPtr pConnectionPoint, object sink)
        {
            // IConnectionPoint::Advise is at vtable offset 5
            IntPtr vtable = Marshal.ReadIntPtr(pConnectionPoint);
            IntPtr pAdvise = Marshal.ReadIntPtr(vtable, IntPtr.Size * 5);

            var advise = Marshal.GetDelegateForFunctionPointer<AdviseDelegate>(pAdvise);
            
            IntPtr pSink = Marshal.GetIUnknownForObject(sink);
            try
            {
                int cookie;
                int hr = advise(pConnectionPoint, pSink, out cookie);
                Marshal.ThrowExceptionForHR(hr);
                return cookie;
            }
            finally
            {
                Marshal.Release(pSink);
            }
        }

        private static void Unadvise(IntPtr pConnectionPoint, int cookie)
        {
            // IConnectionPoint::Unadvise is at vtable offset 6
            IntPtr vtable = Marshal.ReadIntPtr(pConnectionPoint);
            IntPtr pUnadvise = Marshal.ReadIntPtr(vtable, IntPtr.Size * 6);

            var unadvise = Marshal.GetDelegateForFunctionPointer<UnadviseDelegate>(pUnadvise);
            
            int hr = unadvise(pConnectionPoint, cookie);
            Marshal.ThrowExceptionForHR(hr);
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int FindConnectionPointDelegate(IntPtr pThis, ref Guid riid, out IntPtr ppCP);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int AdviseDelegate(IntPtr pThis, IntPtr pUnkSink, out int pdwCookie);

        [UnmanagedFunctionPointer(CallingConvention.StdCall)]
        private delegate int UnadviseDelegate(IntPtr pThis, int dwCookie);
    }
}
