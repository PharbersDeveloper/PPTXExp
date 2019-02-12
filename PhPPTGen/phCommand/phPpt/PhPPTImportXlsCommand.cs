using System;
using System.IO;
using Spire.Xls;
using Spire.Presentation;
using System.Drawing;
using Spire.Presentation.Drawing;
using PhPPTGen.phCommand.phExcel;

namespace PhPPTGen.phCommand.phPpt {
    public class PhPPTImportXlsCommand : PhCommand {
        public PhPPTImportXlsCommand() {

        }

        public override object Exec(params object[] parameters) {
            Console.WriteLine("PPTGen PhCommand: Insert table into pptx");
            var req = (phModel.PhRequest)parameters[0];
            var jobid = req.jobid;
            var e2p = req.e2p;

            /**
             * 1. go to the work place
             */
            var fct = phCommandFactory.PhCommandFactory.GetInstance();
            var tmpDir = fct.GetTmpDictionary();
            var workingPath = tmpDir + "\\" + jobid;
			string workbookKey = jobid + e2p.name;

			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2p.name + ".xlsx";
			var xlsPath = workingPath + "\\" + e2p.name + ".xls";
			if (!File.Exists(ePath)) {
                throw new Exception("Excel name is not exists");
            }

            Workbook book = new Workbook();
			//PhUpdateXlsCommand.workbookMap.TryGetValue(workbookKey, out book);
			//PhUpdateXlsCommand.workbookMap.Remove(workbookKey);
			book.LoadFromFile(ePath);
			book.SaveToFile(xlsPath);
			book.LoadFromFile(xlsPath);
            Worksheet sheet = book.Worksheets[0];
            var col = sheet.Columns.Length;
            var row = sheet.Rows.Length;
			sheet.Rows[1].Cells[col - 1].Style.HorizontalAlignment = HorizontalAlignType.Left;

			/**
             * 3. put the excel into pptx
             */
			var ppt_path = workingPath + "\\" + "result.pptx";

            Presentation ppt = new Presentation();
            ppt.LoadFromFile(ppt_path);

            while (ppt.Slides.Count <= e2p.slider) {
                ppt.Slides.Append();
            }

            Image image = book.Worksheets[0].SaveToImage(1, 1, row, col);
            IImageData oleImage = ppt.Images.Append(image);
            Rectangle rec = new Rectangle(0, 0, 0, 0);
            if (e2p.pos.Length == 4){
                rec = new Rectangle(e2p.pos[0], e2p.pos[1], image.Width * e2p.pos[2] / 100, image.Height * e2p.pos[3] / 100);
            } else{
                rec = new Rectangle(e2p.pos[0], e2p.pos[1], image.Width, image.Height);
            }
            
            using (MemoryStream ms = new MemoryStream()) {
                book.SaveToStream(ms);
                ms.Position = 0;
                Spire.Presentation.IOleObject oleObject = ppt.Slides[e2p.slider].Shapes.AppendOleObject("excel", ms.ToArray(), rec);
                oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
                oleObject.ProgId = "Excel.Sheet.8";

			}
            ppt.SaveToFile(ppt_path, Spire.Presentation.FileFormat.Pptx2010);
			//book.SaveToFile(ePath);
			return null;
        }
    }
}
