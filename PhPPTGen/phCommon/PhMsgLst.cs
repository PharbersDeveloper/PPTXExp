using System;
using System.Net;
using System.Threading;
using System.Collections.Generic;

namespace PhPPTGen.phCommon {
    public class PhMsgLst {
        private Object locker = new Object();
        private List<phMsgDefine.PhMsgContent> lst = new List<phMsgDefine.PhMsgContent>();
        private bool isRunning = true;

        private static PhMsgLst instance;
        public static PhMsgLst GetInstance() {
            if (instance == null) {
                instance = new PhMsgLst();
            }
            return instance;
        }

        private PhMsgLst() {

        }

        public void StartChecking() {
            Thread t = new Thread(new ThreadStart(this.CheckingHandler));
            t.Start();
            //t.Join();
        }

        public void StopChecking() {
            this.isRunning = false;
        }

        public void CheckingHandler() {
            
            while (isRunning) {
                Thread.Sleep(500);
                lock(locker) {
                    if (lst.Count > 0) {
                        List<phMsgDefine.PhMsgContent>.Enumerator iterator = lst.GetEnumerator();
                        if (iterator.MoveNext()) {
                            phMsgDefine.PhMsgContent current = iterator.Current;
                            phModel.PhRequest req = PhCommon.Content2Object<phModel.PhRequest>(current);
                            Console.WriteLine("Current Command is :");
                            Console.WriteLine(req.command);
                            string cls = phModel.PhMsgDefine.PhCommand2Cls(req.command);
                            phCommandFactory.PhCommandFactory fct = phCommandFactory.PhCommandFactory.GetInstance();
                            fct.CreateCommandInstance(cls, req);
                            lst.Remove(current);
                        }
                    }
                }
            }

            StopChecking();
        }

        public void PushMsg(phMsgDefine.PhMsgContent msg) {
            lock(locker) {
                lst.Add(msg);
                Console.WriteLine(lst.Count);
            }
        }
    }
}
