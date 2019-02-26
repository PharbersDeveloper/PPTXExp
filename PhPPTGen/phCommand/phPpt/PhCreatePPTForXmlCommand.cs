using PhPPTGen.phOpenxml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;


namespace PhPPTGen.phCommand.phPpt {
	class PhCreatePPTForXmlCommand: PhCommand {
		public override object Exec(params object[] parameters) {
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
			* 2. crate a result.pptx file
			*/
			PhOpenxmlPPTHandler.GetInstance().CreatePresentation(file_path);

			return null;
		}

	}
}
