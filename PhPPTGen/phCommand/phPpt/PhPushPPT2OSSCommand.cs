using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phPpt {
	class PhPushPPT2OSSCommand: PhCommand {
		public PhPushPPT2OSSCommand() {
		}

		public override Object Exec(params Object[] parameters) {
			Console.WriteLine("PPTExp PhCommand: Create a new PPTX, with the UUID name!");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;

			/**
             * 1. go to the tmp dir
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;
			var file_path = workingPath + "\\result.pptx";

			/**
             * 2. push result.pptx file to OSS
             */

			PhOss.PhOssHandler.GetInstance().UploadPPT(file_path, jobid);

			return null;
		}
	}
}
