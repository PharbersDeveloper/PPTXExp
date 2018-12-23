using System;
using System.Threading;

namespace PhPPTGen {
    class Program {
        static void Main(string[] args) {
            phSocket.PhThreadSocketServ s = new phSocket.PhThreadSocketServ();
            s.startListen();
            phCommon.PhMsgLst lst = phCommon.PhMsgLst.GetInstance();
            lst.StartChecking();
        }
    }
}