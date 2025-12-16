using InteropFromAcadAddin;
using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace AcadDocEventsTester
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.None)]
    public class ComEventSink : IDocumentEventServiceEvents
    {
        // =====================================================================
        // DOCUMENT COMMAND EVENTS
        // =====================================================================
        
        public void CommandStartedEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? CMD START: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void CommandEndedEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? CMD END: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void CommandCancelledEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? CMD CANCEL: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void CommandFailedEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? CMD FAIL: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void UnknownCommandEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? UNKNOWN CMD: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        // =====================================================================
        // DOCUMENT LISP EVENTS
        // =====================================================================
        
        public void LispCancelledEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? LISP CANCEL | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void LispEndedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.DarkCyan;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? LISP END | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void LispWillStartEvent(string documentName, string firstLine)
        {
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? LISP START: {firstLine} | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        // =====================================================================
        // DOCUMENT CLOSE EVENTS
        // =====================================================================
        
        public void BeginDocumentCloseEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ?? BEGIN DOC CLOSE | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void CloseAbortedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? CLOSE ABORTED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void CloseWillStartEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ?? CLOSE WILL START | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        // =====================================================================
        // DOCUMENTMANAGER EVENTS
        // =====================================================================
        
        public void DocumentCreatedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC CREATED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentToBeActivatedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC TO BE ACTIVATED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentActivatedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC ACTIVATED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentToBeDeactivatedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC TO BE DEACTIVATED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentActivationChangedEvent(string documentName, bool activated)
        {
            Console.ForegroundColor = activated ? ConsoleColor.White : ConsoleColor.Gray;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC ACTIVATION CHANGED: {(activated ? "ACTIVE" : "INACTIVE")} | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentBecameCurrentEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Cyan;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC BECAME CURRENT | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentToBeDestroyedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC TO BE DESTROYED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentDestroyedEvent(string fileName)
        {
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC DESTROYED | {GetShortDocName(fileName)}");
            Console.ResetColor();
        }

        public void DocumentCreationStartedEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC CREATION STARTED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentCreationCancelledEvent(string documentName)
        {
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ? DOC CREATION CANCELLED | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentLockModeChangedEvent(string documentName, string previousMode, string currentMode, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ?? LOCK CHANGED: {previousMode} ? {currentMode} | Cmd: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentLockModeChangeVetoedEvent(string documentName, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.DarkYellow;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ?? LOCK VETOED | Cmd: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        public void DocumentLockModeWillChangeEvent(string documentName, string currentMode, string newMode, string commandName)
        {
            Console.ForegroundColor = ConsoleColor.DarkBlue;
            Console.WriteLine($"[{DateTime.Now:HH:mm:ss.fff}] ?? LOCK WILL CHANGE: {currentMode} ? {newMode} | Cmd: '{commandName}' | {GetShortDocName(documentName)}");
            Console.ResetColor();
        }

        // =====================================================================
        // HELPER METHODS
        // =====================================================================
        
        private string GetShortDocName(string fullPath)
        {
            if (string.IsNullOrEmpty(fullPath))
                return "<no document>";

            try
            {
                return System.IO.Path.GetFileName(fullPath);
            }
            catch
            {
                return fullPath;
            }
        }
    }
}
