using System;
using System.Runtime.InteropServices;

namespace AcadDocEventsTester
{
    /// <summary>
    /// Utility class for COM interop in .NET 8
    /// Provides GetActiveObject functionality removed from Marshal in .NET 8
    /// </summary>
    internal static class ComUtils
    {
        [DllImport("ole32.dll")]
        private static extern int CLSIDFromProgIDEx(
            [MarshalAs(UnmanagedType.LPWStr)] string lpszProgID,
            out Guid lpclsid);

        [DllImport("oleaut32.dll")]
        private static extern int GetActiveObject(
            [MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            IntPtr pvReserved,
            [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        /// <summary>
        /// Gets a running COM object from the Running Object Table (ROT)
        /// Replacement for Marshal.GetActiveObject which is not available in .NET 8
        /// </summary>
        /// <param name="progId">The ProgID of the COM object (e.g., "AutoCAD.Application.26.0")</param>
        /// <param name="throwOnError">If true, throws exception on error; otherwise returns null</param>
        /// <returns>The COM object if found, otherwise null</returns>
        public static object? GetActiveObject(string progId, bool throwOnError = false)
        {
            if (progId == null)
                throw new ArgumentNullException(nameof(progId));

            // Get CLSID from ProgID
            var hr = CLSIDFromProgIDEx(progId, out var clsid);
            if (hr < 0)
            {
                if (throwOnError)
                    Marshal.ThrowExceptionForHR(hr);
                return null;
            }

            // Get the active object
            hr = GetActiveObject(clsid, IntPtr.Zero, out var obj);
            if (hr < 0)
            {
                if (throwOnError)
                    Marshal.ThrowExceptionForHR(hr);
                return null;
            }

            return obj;
        }
    }
}
