﻿using System;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.Text;

namespace PhPPTGen.phSocket {
    public class PhThreadClientHandler {
        private TcpClient client = null;
        private NetworkStream ns = null;
        private bool isRunning = true;
        private Thread t = null;
        private Byte[] bytes = new Byte[1024];

        private DateTime last_msg = DateTime.Now;

        public PhThreadClientHandler(TcpClient c, NetworkStream n) {
            this.client = c;
            this.ns = n;
        }

        public void StartClientHandler() {
            t = new Thread(new ThreadStart(this.HandleClient));
            t.Start();
            //t.Join();
        }

        public void StopClientHandler() {
            Console.WriteLine("Close Client Handler!!");
            ns.Close();
            client.Close();
            isRunning = false;
        }

        public void HandleClient() {

            while (isRunning) {

                if (!client.Connected) {
                    StopClientHandler();
                    break;
                }

                if (client.Available == 0 || !ns.CanRead) {
                    TimeSpan span = DateTime.Now - last_msg;
                    if (span.TotalMinutes > 60) {
                        client.Close();
                    }
                    Thread.Sleep(500);
                    continue;
                }

                last_msg = DateTime.Now;

                try {
                    Array.Clear(bytes, 0, 1024);
                    int nRec = ns.Read(bytes, 0, 1024);

                    if (nRec > 0) {
                        phCommon.phMsgDefine.PhMsgContent msg = new phCommon.phMsgDefine.PhMsgContent();
                        msg.msg_content = Encoding.UTF8.GetString(bytes); 
                        phCommon.PhMsgLst lst = phCommon.PhMsgLst.GetInstance();
                        lst.PushMsg(msg);
                    }

                } catch (Exception e) {
                    Console.WriteLine(e.ToString());
                }
            }

            StopClientHandler();
        }
    }
}
