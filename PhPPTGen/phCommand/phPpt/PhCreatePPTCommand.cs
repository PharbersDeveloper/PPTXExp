using System;
using Spire.Xls;
using Spire.Presentation;
using System.Drawing;
using System.IO;
using Spire.Presentation.Drawing;

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
            var workingPath = tmpDir + jobid;
            var file_path = workingPath + "\\" + "result.pptx";

            /**
             * 2. crate a result.pptx file
             */
            Presentation ppt = new Presentation();
            ppt.SaveToFile(file_path, Spire.Presentation.FileFormat.Pptx2010);

            //var file_name = req.file_name;
            //Console.WriteLine("get file with name");
            //var fct = phCommandFactory.PhCommandFactory.GetInstance();
            //var pptx = fct.GetHandledPPTX(file_name);

            //if (pptx == null) {
                //Workbook book = new Workbook();
                //book.LoadFromFile(@"D:\\pptresult\data2.xls");
                //Image image = book.Worksheets[0].SaveToImage(1, 1, 5, 4);
                ////Image image = book.Worksheets[0].ToImage(1, 1, 5, 4);

                //Presentation ppt = new Presentation();
                //IImageData oleImage = ppt.Images.Append(image);
                //Rectangle rec = new Rectangle(60, 60, image.Width, image.Height);
                //using (MemoryStream ms = new MemoryStream())
                //{
                //    book.SaveToStream(ms);
                //    ms.Position = 0;
                //    Spire.Presentation.IOleObject oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", ms.ToArray(), rec);
                //    oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
                //    oleObject.ProgId = "Excel.Sheet.8";
                //}
                //ppt.SaveToFile(@"D:\\pptresult\InsertOle.pptx", Spire.Presentation.FileFormat.Pptx2010);
            //}

            //this.TestPPT();
            //this.TestWithExcel();

            return null;
        }
    }
}
