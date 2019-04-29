using PhPPTGen.phModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace PhPPTGen.phCommon {
	public class PhRequestLst {
		private Object locker = new Object();
		private List<PhRequest> lst = new List<PhRequest>();
		private bool isRunning = true;

		private static PhRequestLst instance;
		public static PhRequestLst GetInstance() {
			if (instance == null) {
				instance = new PhRequestLst();
			}
			return instance;
		}

		private PhRequestLst() {

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
				lock (locker) {
					if (lst.Count > 0) {
						List<PhRequest>.Enumerator iterator = lst.GetEnumerator();
						if (iterator.MoveNext()) {
							PhRequest req = iterator.Current;
							Console.WriteLine("Current Command is :");
							Console.WriteLine(req.command);
							string cls = phModel.PhMsgDefine.PhCommand2Cls(req.command);
							phCommandFactory.PhCommandFactory fct = phCommandFactory.PhCommandFactory.GetInstance();
							try {
								fct.CreateCommandInstance(cls, req);
							} catch (Exception ex) {
								Console.WriteLine("Failed with error info: {0}", ex.Message);
							}

						lst.Remove(req);
						}
					}
				}
			}

			StopChecking();
		}

		public void PushMsg(PhRequest req) {
			lock (locker) {
				lst.Add(req);
				Console.WriteLine(lst.Count);
			}
		}
	}
}
