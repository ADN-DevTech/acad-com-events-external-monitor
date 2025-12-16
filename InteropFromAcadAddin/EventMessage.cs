using System;

namespace InteropFromAcadAddin
{
    /// <summary>
    /// Message structure for IPC communication via Named Pipes
    /// </summary>
    [Serializable]
    public class EventMessage
    {
        public string EventType { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public string CommandName { get; set; } = string.Empty;
        public string? AdditionalData { get; set; }
        public DateTime Timestamp { get; set; } = DateTime.Now;

        public override string ToString()
        {
            return $"[{Timestamp:HH:mm:ss.fff}] {EventType} - {DocumentName} - {CommandName}";
        }
    }
}

