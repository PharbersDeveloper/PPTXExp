using System;
using System.Threading;

namespace PPTXExp.phSocket {
    public class PhThreadSocketServ {
        PhSocketServer serv = new PhSocketServer();

        public void Threadproc() {
            serv.StartListeningData();
        }

        public void startListen() {
            Thread t = new Thread(new ThreadStart(this.Threadproc));
            t.Start();
            //t.Join();
        }

        public void stopListen() {
            serv.StopListeningData();
        }
    }
}
