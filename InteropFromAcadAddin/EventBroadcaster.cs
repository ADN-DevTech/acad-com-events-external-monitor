using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Pipes;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace InteropFromAcadAddin
{
    /// <summary>
    /// Named Pipe server that broadcasts AutoCAD events to external clients
    /// </summary>
    public class EventBroadcaster : IDisposable
    {
        private const string PipeName = "AcadEventsPipe";
        private readonly List<NamedPipeServerStream> _connectedClients = new();
        private readonly object _lock = new();
        private CancellationTokenSource? _cancellationTokenSource;
        private bool _isRunning;

        public void Start()
        {
            if (_isRunning) return;
            _isRunning = true;
            _cancellationTokenSource = new CancellationTokenSource();
            
            // Start listening for client connections
            Task.Run(() => ListenForClients(_cancellationTokenSource.Token));
        }

        public void Stop()
        {
            if (!_isRunning) return;
            _isRunning = false;
            _cancellationTokenSource?.Cancel();

            lock (_lock)
            {
                foreach (var client in _connectedClients)
                {
                    try
                    {
                        client.Dispose();
                    }
                    catch { }
                }
                _connectedClients.Clear();
            }
        }

        private async Task ListenForClients(CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                try
                {
                    var pipeServer = new NamedPipeServerStream(
                        PipeName,
                        PipeDirection.Out,
                        NamedPipeServerStream.MaxAllowedServerInstances,
                        PipeTransmissionMode.Message,
                        PipeOptions.Asynchronous);

                    await pipeServer.WaitForConnectionAsync(cancellationToken);

                    lock (_lock)
                    {
                        _connectedClients.Add(pipeServer);
                    }

                    System.Diagnostics.Trace.WriteLine($"Client connected. Total clients: {_connectedClients.Count}");
                }
                catch (OperationCanceledException)
                {
                    break;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Trace.WriteLine($"Error accepting client: {ex.Message}");
                    await Task.Delay(1000, cancellationToken);
                }
            }
        }

        public void BroadcastEvent(EventMessage message)
        {
            if (!_isRunning) return;

            lock (_lock)
            {
                var disconnectedClients = new List<NamedPipeServerStream>();

                foreach (var client in _connectedClients)
                {
                    try
                    {
                        if (client.IsConnected)
                        {
                            var json = JsonSerializer.Serialize(message);
                            var writer = new StreamWriter(client) { AutoFlush = true };
                            writer.WriteLine(json);
                        }
                        else
                        {
                            disconnectedClients.Add(client);
                        }
                    }
                    catch
                    {
                        disconnectedClients.Add(client);
                    }
                }

                // Remove disconnected clients
                foreach (var client in disconnectedClients)
                {
                    _connectedClients.Remove(client);
                    client.Dispose();
                }
            }
        }

        public void Dispose()
        {
            Stop();
            _cancellationTokenSource?.Dispose();
        }
    }
}

