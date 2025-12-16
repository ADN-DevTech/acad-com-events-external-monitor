using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections;
using System.Runtime.InteropServices;
using System.Reflection;
using Exception = System.Exception;
using System.EnterpriseServices;

// Assembly Attributes: Required for the AutoCAD runtime and COM visibility
[assembly: ExtensionApplication(typeof(InteropFromAcadAddin.PluginExtension))]
[assembly: CommandClass(typeof(InteropFromAcadAddin.PluginEntry))]
[assembly: ComVisible(true)] // Ensure assembly-level COM visibility

namespace InteropFromAcadAddin
{
    // Define the ProgId the client uses
    public static class Consts
    {
        // ProgId used by the client when looking up the service
        public const string AcadPluginProgId = "InteropFromAcadAddin.IDocumentEventService";
    }

    // =========================================================================
    // COM INTERFACES
    // =========================================================================

    // Event Interface: IDocumentEventServiceEvents
    // This defines the outgoing events the COM object fires. The client's sink 
    // (ComEventSink.cs) will implement IDispatch for this GUID.
    [ComVisible(true)]
    [Guid("A7E1B0A4-2E2E-4B6D-9E39-7D6C5B2C9A11")] // This is the GUID used in the client's Advise call
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)] // Must be IDispatch for connection points
    public interface IDocumentEventServiceEvents
    {
        // Document Command Events
        [DispId(1)]
        void CommandStartedEvent(string documentName, string commandName);
        [DispId(2)]
        void CommandEndedEvent(string documentName, string commandName);
        [DispId(3)]
        void CommandCancelledEvent(string documentName, string commandName);
        [DispId(4)]
        void CommandFailedEvent(string documentName, string commandName);
        [DispId(5)]
        void UnknownCommandEvent(string documentName, string commandName);
        
        // Document LISP Events
        [DispId(6)]
        void LispCancelledEvent(string documentName);
        [DispId(7)]
        void LispEndedEvent(string documentName);
        [DispId(8)]
        void LispWillStartEvent(string documentName, string firstLine);

        // Document Close Events
        [DispId(9)]
        void BeginDocumentCloseEvent(string documentName);
        [DispId(10)]
        void CloseAbortedEvent(string documentName);
        [DispId(11)]
        void CloseWillStartEvent(string documentName);

        // DocumentManager Events
        [DispId(20)]
        void DocumentCreatedEvent(string documentName);
        [DispId(21)]
        void DocumentToBeActivatedEvent(string documentName);
        [DispId(22)]
        void DocumentActivatedEvent(string documentName);
        [DispId(23)]
        void DocumentToBeDeactivatedEvent(string documentName);
        [DispId(24)]
        void DocumentActivationChangedEvent(string documentName, bool activated);
        [DispId(25)]
        void DocumentBecameCurrentEvent(string documentName);
        [DispId(26)]
        void DocumentToBeDestroyedEvent(string documentName);
        [DispId(27)]
        void DocumentDestroyedEvent(string fileName);
        [DispId(28)]
        void DocumentCreationStartedEvent(string documentName);
        [DispId(29)]
        void DocumentCreationCancelledEvent(string documentName);
        [DispId(30)]
        void DocumentLockModeChangedEvent(string documentName, string previousMode, string currentMode, string commandName);
        [DispId(31)]
        void DocumentLockModeChangeVetoedEvent(string documentName, string commandName);
        [DispId(32)]
        void DocumentLockModeWillChangeEvent(string documentName, string currentMode, string newMode, string commandName);
    }

    // Service Interface: IDocumentEventService
    [ComVisible(true)]
    [Guid("D1C762E4-2C22-482B-9333-82F3E906E6B4")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDocumentEventService
    {
        [DispId(101)]
        void Start();
        [DispId(102)]
        void Stop();
    }


    // =========================================================================
    // APPLICATION EXTENSION & REGISTRATION
    // =========================================================================

    public class PluginExtension : IExtensionApplication
    {
        // Now, we hold the instance here, as PluginEntry is no longer a singleton.
        // It's created once when the AutoCAD application loads the extension.
        private static PluginEntry? _service;

        public void Initialize()
        {
            // IMPORTANT: We cannot rely on the COM service client to initialize the 
            // instance, so we create it here when AutoCAD loads the add-in.
            // If the client calls Activator.CreateInstance, a separate instance will be created.
            if (_service == null)
            {
                _service = new PluginEntry();
            }

            _service.AttachToAllOpenDocuments();
            // Start listening when the add-in loads
            _service.Start();

            // Explicit self-registration for the COM component
            SelfRegisterComTypes();
        }

        public void Terminate()
        {
            _service?.Stop();
            _service?.DetachFromAllDocuments();
        }

        private static void SelfRegisterComTypes()
        {
            try
            {
                var asm = Assembly.GetExecutingAssembly();
                var reg = new RegistrationServices();

                // CRITICAL: SetCodeBase is essential to register the DLL's location.
                reg.RegisterAssembly(asm, AssemblyRegistrationFlags.SetCodeBase);
            }
            catch (Exception ex)
            {
                // This will output a message to the AutoCAD trace log if registration fails
                System.Diagnostics.Trace.WriteLine("COM Self-Registration failed: " + ex.Message);
            }
        }
    }

    // =========================================================================
    // COM SERVER CLASS
    // =========================================================================

    [ProgId(Consts.AcadPluginProgId)]
    [Guid("75B37C6C-8047-41E3-8C4D-D5196EEFF368")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    [ComSourceInterfaces(typeof(IDocumentEventServiceEvents))]
    [ComVisible(true)]
    public class PluginEntry : ServicedComponent, IDocumentEventService
    {
      
        // ======================================================================
        // COM EVENT DELEGATES - Document Command Events
        // ======================================================================
        public delegate void CommandStartedEventHandler(string documentName, string commandName);
        public event CommandStartedEventHandler CommandStartedEvent;

        public delegate void CommandEndedEventHandler(string documentName, string commandName);
        public event CommandEndedEventHandler CommandEndedEvent;

        public delegate void CommandCancelledEventHandler(string documentName, string commandName);
        public event CommandCancelledEventHandler CommandCancelledEvent;

        public delegate void CommandFailedEventHandler(string documentName, string commandName);
        public event CommandFailedEventHandler CommandFailedEvent;

        public delegate void UnknownCommandEventHandler(string documentName, string commandName);
        public event UnknownCommandEventHandler UnknownCommandEvent;

        // ======================================================================
        // COM EVENT DELEGATES - Document LISP Events
        // ======================================================================
        public delegate void LispCancelledEventHandler(string documentName);
        public event LispCancelledEventHandler LispCancelledEvent;

        public delegate void LispEndedEventHandler(string documentName);
        public event LispEndedEventHandler LispEndedEvent;

        public delegate void LispWillStartEventHandler(string documentName, string firstLine);
        public event LispWillStartEventHandler LispWillStartEvent;

        // ======================================================================
        // COM EVENT DELEGATES - Document Close Events
        // ======================================================================
        public delegate void BeginDocumentCloseEventHandler(string documentName);
        public event BeginDocumentCloseEventHandler BeginDocumentCloseEvent;

        public delegate void CloseAbortedEventHandler(string documentName);
        public event CloseAbortedEventHandler CloseAbortedEvent;

        public delegate void CloseWillStartEventHandler(string documentName);
        public event CloseWillStartEventHandler CloseWillStartEvent;

        // ======================================================================
        // COM EVENT DELEGATES - DocumentManager Events
        // ======================================================================
        public delegate void DocumentCreatedEventHandler(string documentName);
        public event DocumentCreatedEventHandler DocumentCreatedEvent;

        public delegate void DocumentToBeActivatedEventHandler(string documentName);
        public event DocumentToBeActivatedEventHandler DocumentToBeActivatedEvent;

        public delegate void DocumentActivatedEventHandler(string documentName);
        public event DocumentActivatedEventHandler DocumentActivatedEvent;

        public delegate void DocumentToBeDeactivatedEventHandler(string documentName);
        public event DocumentToBeDeactivatedEventHandler DocumentToBeDeactivatedEvent;

        public delegate void DocumentActivationChangedEventHandler(string documentName, bool activated);
        public event DocumentActivationChangedEventHandler DocumentActivationChangedEvent;

        public delegate void DocumentBecameCurrentEventHandler(string documentName);
        public event DocumentBecameCurrentEventHandler DocumentBecameCurrentEvent;

        public delegate void DocumentToBeDestroyedEventHandler(string documentName);
        public event DocumentToBeDestroyedEventHandler DocumentToBeDestroyedEvent;

        public delegate void DocumentDestroyedEventHandler(string fileName);
        public event DocumentDestroyedEventHandler DocumentDestroyedEvent;

        public delegate void DocumentCreationStartedEventHandler(string documentName);
        public event DocumentCreationStartedEventHandler DocumentCreationStartedEvent;

        public delegate void DocumentCreationCancelledEventHandler(string documentName);
        public event DocumentCreationCancelledEventHandler DocumentCreationCancelledEvent;

        public delegate void DocumentLockModeChangedEventHandler(string documentName, string previousMode, string currentMode, string commandName);
        public event DocumentLockModeChangedEventHandler DocumentLockModeChangedEvent;

        public delegate void DocumentLockModeChangeVetoedEventHandler(string documentName, string commandName);
        public event DocumentLockModeChangeVetoedEventHandler DocumentLockModeChangeVetoedEvent;

        public delegate void DocumentLockModeWillChangeEventHandler(string documentName, string currentMode, string newMode, string commandName);
        public event DocumentLockModeWillChangeEventHandler DocumentLockModeWillChangeEvent;


        private readonly Hashtable _eventHandlers = new Hashtable();
        private bool _isStarted = false;

        // Constructor MUST be public for COM activation.
        // It must NOT contain any complex AutoCAD API calls that might fail.
        private static PluginEntry _instance;
        public static PluginEntry Instance => _instance ??= new PluginEntry();

        public PluginEntry()
        {
            _instance = this;
        }

        // IDocumentEventService Implementation
        public void Start()
        {
            if (_isStarted) return;
            _isStarted = true;

            // Register DocumentManager event handlers
            Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
            Application.DocumentManager.DocumentToBeActivated += DocumentManager_DocumentToBeActivated;
            Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
            Application.DocumentManager.DocumentToBeDeactivated += DocumentManager_DocumentToBeDeactivated;
            Application.DocumentManager.DocumentActivationChanged += DocumentManager_DocumentActivationChanged;
            Application.DocumentManager.DocumentBecameCurrent += DocumentManager_DocumentBecameCurrent;
            Application.DocumentManager.DocumentToBeDestroyed += DocumentManager_DocumentToBeDestroyed;
            Application.DocumentManager.DocumentDestroyed += DocumentManager_DocumentDestroyed;
            Application.DocumentManager.DocumentCreationCanceled += DocumentManager_DocumentCreationCanceled;
            Application.DocumentManager.DocumentLockModeChanged += DocumentManager_DocumentLockModeChanged;
            Application.DocumentManager.DocumentLockModeChangeVetoed += DocumentManager_DocumentLockModeChangeVetoed;
            Application.DocumentManager.DocumentLockModeWillChange += DocumentManager_DocumentLockModeWillChange;

            // Attach to all currently open documents
            AttachToAllOpenDocuments();
        }

        public void Stop()
        {
            if (!_isStarted) return;
            _isStarted = false;

            // Unregister DocumentManager event handlers
            Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;
            Application.DocumentManager.DocumentToBeActivated -= DocumentManager_DocumentToBeActivated;
            Application.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;
            Application.DocumentManager.DocumentToBeDeactivated -= DocumentManager_DocumentToBeDeactivated;
            Application.DocumentManager.DocumentActivationChanged -= DocumentManager_DocumentActivationChanged;
            Application.DocumentManager.DocumentBecameCurrent -= DocumentManager_DocumentBecameCurrent;
            Application.DocumentManager.DocumentToBeDestroyed -= DocumentManager_DocumentToBeDestroyed;
            Application.DocumentManager.DocumentDestroyed -= DocumentManager_DocumentDestroyed;
            Application.DocumentManager.DocumentCreationCanceled -= DocumentManager_DocumentCreationCanceled;
            Application.DocumentManager.DocumentLockModeChanged -= DocumentManager_DocumentLockModeChanged;
            Application.DocumentManager.DocumentLockModeChangeVetoed -= DocumentManager_DocumentLockModeChangeVetoed;
            Application.DocumentManager.DocumentLockModeWillChange -= DocumentManager_DocumentLockModeWillChange;

            // Detach from all documents
            DetachFromAllDocuments();
        }

        public void AttachToAllOpenDocuments()
        {
            foreach (Document doc in Application.DocumentManager)
            {
                AttachToDocument(doc);
            }
        }

        public void DetachFromAllDocuments()
        {
            // Iterate over a copy of the keys since detaching modifies the collection
            foreach (Document doc in new ArrayList(_eventHandlers.Keys))
            {
                DetachFromDocument(doc);
            }
        }

        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentCreatedEvent?.Invoke(docName);
            AttachToDocument(e.Document);
        }

        // =========================================================================
        // DOCUMENTMANAGER EVENT HANDLERS
        // =========================================================================

        private void DocumentManager_DocumentToBeActivated(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentToBeActivatedEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentActivatedEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentToBeDeactivated(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentToBeDeactivatedEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentActivationChanged(object sender, DocumentActivationChangedEventArgs e)
        {
            var docName = Application.DocumentManager.MdiActiveDocument?.Name ?? string.Empty;
            DocumentActivationChangedEvent?.Invoke(docName, e.NewValue);
        }

        private void DocumentManager_DocumentBecameCurrent(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentBecameCurrentEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentToBeDestroyed(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentToBeDestroyedEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentDestroyed(object sender, DocumentDestroyedEventArgs e)
        {
            var fileName = e.FileName ?? string.Empty;
            DocumentDestroyedEvent?.Invoke(fileName);
        }

        private void DocumentManager_DocumentCreationCanceled(object sender, DocumentCollectionEventArgs e)
        {
            var docName = e.Document?.Name ?? string.Empty;
            DocumentCreationCancelledEvent?.Invoke(docName);
        }

        private void DocumentManager_DocumentLockModeChanged(object sender, DocumentLockModeChangedEventArgs e)
        {
            var docName = Application.DocumentManager.MdiActiveDocument?.Name ?? string.Empty;
            var previousMode = e.MyPreviousMode.ToString();
            var currentMode = e.MyCurrentMode.ToString();
            var commandName = e.GlobalCommandName ?? string.Empty;
            DocumentLockModeChangedEvent?.Invoke(docName, previousMode, currentMode, commandName);
        }

        private void DocumentManager_DocumentLockModeChangeVetoed(object sender, DocumentLockModeChangeVetoedEventArgs e)
        {
            var docName = Application.DocumentManager.MdiActiveDocument?.Name ?? string.Empty;
            var commandName = e.GlobalCommandName ?? string.Empty;
            DocumentLockModeChangeVetoedEvent?.Invoke(docName, commandName);
        }

        private void DocumentManager_DocumentLockModeWillChange(object sender, DocumentLockModeWillChangeEventArgs e)
        {
            var docName = Application.DocumentManager.MdiActiveDocument?.Name ?? string.Empty;
            var currentMode = e.MyCurrentMode.ToString();
            var newMode = e.MyNewMode.ToString();
            var commandName = e.GlobalCommandName ?? string.Empty;
            DocumentLockModeWillChangeEvent?.Invoke(docName, currentMode, newMode, commandName);
        }

        // =========================================================================
        // DOCUMENT ATTACH/DETACH METHODS
        // =========================================================================

        private void AttachToDocument(Document doc)
        {
            if (!_eventHandlers.ContainsKey(doc))
            {
                // Subscribe to all relevant AutoCAD document events
                // Command Events
                doc.CommandWillStart += callback_CommandWillStart;
                doc.CommandEnded += callback_CommandEnded;
                doc.CommandCancelled += callback_CommandCancelled;
                doc.CommandFailed += callback_CommandFailed;
                doc.UnknownCommand += callback_UnknownCommand;
                
                // LISP Events
                doc.LispCancelled += callback_LispCancelled;
                doc.LispEnded += callback_LispEnded;
                doc.LispWillStart += callback_LispWillStart;

                // Document Close Events
                doc.BeginDocumentClose += callback_BeginDocumentClose;
                doc.CloseAborted += callback_CloseAborted;
                doc.CloseWillStart += callback_CloseWillStart;

                _eventHandlers.Add(doc, true);
            }
        }

        private void DetachFromDocument(Document doc)
        {
            if (_eventHandlers.ContainsKey(doc))
            {
                // Unsubscribe from all relevant AutoCAD document events
                // Command Events
                doc.CommandWillStart -= callback_CommandWillStart;
                doc.CommandEnded -= callback_CommandEnded;
                doc.CommandCancelled -= callback_CommandCancelled;
                doc.CommandFailed -= callback_CommandFailed;
                doc.UnknownCommand -= callback_UnknownCommand;
                
                // LISP Events
                doc.LispCancelled -= callback_LispCancelled;
                doc.LispEnded -= callback_LispEnded;
                doc.LispWillStart -= callback_LispWillStart;

                // Document Close Events
                doc.BeginDocumentClose -= callback_BeginDocumentClose;
                doc.CloseAborted -= callback_CloseAborted;
                doc.CloseWillStart -= callback_CloseWillStart;

                _eventHandlers.Remove(doc);
            }
        }

        // =========================================================================
        // DOCUMENT EVENT HANDLERS (CALLBACKS)
        // =========================================================================

        private void callback_CommandWillStart(object sender, CommandEventArgs e)
        {
            var docName = SafeDocName(sender);
            CommandStartedEvent?.Invoke(docName, e.GlobalCommandName ?? string.Empty);
        }

        private void callback_CommandEnded(object sender, CommandEventArgs e)
        {
            var docName = SafeDocName(sender);
            CommandEndedEvent?.Invoke(docName, e.GlobalCommandName ?? string.Empty);
        }

        private void callback_CommandCancelled(object sender, CommandEventArgs e)
        {
            var docName = SafeDocName(sender);
            CommandCancelledEvent?.Invoke(docName, e.GlobalCommandName ?? string.Empty);
        }

        private void callback_CommandFailed(object sender, CommandEventArgs e)
        {
            var docName = SafeDocName(sender);
            CommandFailedEvent?.Invoke(docName, e.GlobalCommandName ?? string.Empty);
        }

        private void callback_UnknownCommand(object sender, UnknownCommandEventArgs e)
        {
            var docName = SafeDocName(sender);
            UnknownCommandEvent?.Invoke(docName, e.GlobalCommandName ?? string.Empty);
        }

        private void callback_LispCancelled(object sender, EventArgs e)
        {
            var docName = SafeDocName(sender);
            LispCancelledEvent?.Invoke(docName);
        }

        private void callback_LispEnded(object sender, EventArgs e)
        {
            var docName = SafeDocName(sender);
            LispEndedEvent?.Invoke(docName);
        }

        private void callback_LispWillStart(object sender, LispWillStartEventArgs e)
        {
            var docName = SafeDocName(sender);
            LispWillStartEvent?.Invoke(docName, e.FirstLine ?? string.Empty);
        }

        private void callback_BeginDocumentClose(object sender, DocumentBeginCloseEventArgs e)
        {
            var docName = SafeDocName(sender);
            BeginDocumentCloseEvent?.Invoke(docName);
        }

        private void callback_CloseAborted(object sender, EventArgs e)
        {
            var docName = SafeDocName(sender);
            CloseAbortedEvent?.Invoke(docName);
        }

        private void callback_CloseWillStart(object sender, EventArgs e)
        {
            var docName = SafeDocName(sender);
            CloseWillStartEvent?.Invoke(docName);
        }

        private static string SafeDocName(object sender)
        {
            try
            {
                var doc = sender as Document;
                return doc?.Name ?? string.Empty;
            }
            catch
            {
                return string.Empty;
            }
        }
    }
}