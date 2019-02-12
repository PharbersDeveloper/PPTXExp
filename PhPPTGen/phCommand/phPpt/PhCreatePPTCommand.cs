using System;
using Spire.Presentation;

namespace PhPPTGen.phCommand {
    public class PhCreatePPTCommand : PhCommand {
        public PhCreatePPTCommand() {
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
             * 2. crate a result.pptx file
             */
            Presentation ppt = new Presentation();
            ppt.SaveToFile(file_path, Spire.Presentation.FileFormat.Pptx2010);

            return null;
        }
    }
}
