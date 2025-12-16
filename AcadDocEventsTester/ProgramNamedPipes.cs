using System;
using System.IO;
using System.IO.Pipes;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace AcadDocEventsTester
{
    class ProgramNamedPipes
    {
        const string PipeName = "AcadEventsPipe";

        static async Task Main(string[] args)
        {
            Console.WriteLine("**AutoCAD Event Monitor (.NET 8) - Named Pipes**\n");
            Console.WriteLine("Connecting to AutoCAD plugin...\n");

            var cts = new CancellationTokenSource();
            Console.CancelKeyPress += (s, e) =>
            {
                e.Cancel = true;
                cts.Cancel();
            };

            try
            {
                await using var pipeClient = new NamedPipeClientStream(
                    ".",
                    PipeName,
                    PipeDirection.In);

                Console.WriteLine("Waiting for AutoCAD plugin to start...");
                await pipeClient.ConnectAsync(5000, cts.Token);

                Console.WriteLine("Connected successfully!\n");
                Console.WriteLine("--- Listening for AutoCAD events ---");
                Console.WriteLine("Run commands in AutoCAD to see events here.");
                Console.WriteLine("Press Ctrl+C to stop monitoring...\n");

                using var reader = new StreamReader(pipeClient);

                while (!cts.Token.IsCancellationRequested && pipeClient.IsConnected)
                {
                    var line = await reader.ReadLineAsync();
                    if (line == null) break;

                    try
                    {
                        var message = JsonSerializer.Deserialize<EventMessage>(line);
                        if (message != null)
                        {
                            DisplayEvent(message);
                        }
                    }
                    catch (JsonException ex)
                    {
                        Console.WriteLine($"Error parsing message: {ex.Message}");
                    }
                }
            }
            catch (TimeoutException)
            {
                Console.WriteLine("\nERROR: Connection timeout!");
                Console.WriteLine("Make sure:");
                Console.WriteLine("1. AutoCAD 2026 is running");
                Console.WriteLine("2. InteropFromAcadAddin.dll is NETLOAD'ed");
            }
            catch (OperationCanceledException)
            {
                Console.WriteLine("\n\nStopping event monitoring...");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"\nERROR: {ex.Message}");
            }

            Console.WriteLine("\nPress ENTER to exit...");
            Console.ReadLine();
        }

        static void DisplayEvent(EventMessage message)
        {
            Console.ForegroundColor = message.EventType switch
            {
                "CommandStarted" => ConsoleColor.Green,
                "CommandEnded" => ConsoleColor.Blue,
                "CommandCancelled" => ConsoleColor.Yellow,
                "CommandFailed" => ConsoleColor.Red,
                "DocumentCreated" => ConsoleColor.Cyan,
                _ => ConsoleColor.Gray
            };

            Console.WriteLine(message.ToString());
            Console.ResetColor();
        }
    }

    // Local copy of EventMessage for deserialization
    [Serializable]
    public class EventMessage
    {
        public string EventType { get; set; } = string.Empty;
        public string DocumentName { get; set; } = string.Empty;
        public string CommandName { get; set; } = string.Empty;
        public string? AdditionalData { get; set; }
        public DateTime Timestamp { get; set; }

        public override string ToString()
        {
            return $"[{Timestamp:HH:mm:ss.fff}] {EventType,-20} {DocumentName,-30} {CommandName}";
        }
    }
}

