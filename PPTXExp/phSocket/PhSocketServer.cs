using System;
using System.Net;
using System.Net.Sockets;
using System.Collections.Generic;

namespace PPTXExp.phSocket {
    public class PhSocketServer {
        private bool done = false;
        private int portNum = 9999;
        private TcpListener listener = null;
        private Dictionary<string, PhThreadClientHandler> clients = new Dictionary<string, PhThreadClientHandler>();

        public void StartListeningData() {
            TcpListener listener = new TcpListener(this.portNum);
            listener.Start();

            while (!done) {
                Console.Write("Waiting for connection...");
                TcpClient client = listener.AcceptTcpClient();

                Console.WriteLine("Connection accepted.");
                NetworkStream ns = client.GetStream();

                PhThreadClientHandler handler = new PhThreadClientHandler(client, ns);
                handler.StartClientHandler();
                clients.Add(phCommon.PhCommon.UUID(), handler);
            }

            listener.Stop();
        }

        public void StopListeningData() {
            this.done = false;
        }
    }
}
