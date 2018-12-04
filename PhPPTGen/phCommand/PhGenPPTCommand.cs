using System;
using System.IO;

namespace PhPPTGen.phCommand {
    public class PhGenPPTCommand : PhCommand {
        public PhGenPPTCommand() {

        }

        public override Object Exec(params Object[] parameters) {
            Console.WriteLine("Execute Command: PhPPTGen Generate ppt command");
            var req = (phModel.PhRequest)parameters[0];
            var jobid = req.jobid;
            Console.WriteLine("Execte job with id");
            Console.WriteLine(jobid);

            /**
             * 1. create a tmp dir for all the data
             */
            var fct = phCommandFactory.PhCommandFactory.GetInstance();
            var tmpDir = fct.GetTmpDictionary();
            var workingPath = tmpDir + jobid;
            
            if (!Directory.Exists(workingPath)) {
                Directory.CreateDirectory(workingPath);
            } else {
                throw new Exception("Can not generate ppt working path twice");
            }

            return null;
        }
    }
}
