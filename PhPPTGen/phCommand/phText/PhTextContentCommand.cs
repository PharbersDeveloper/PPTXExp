using System;
using System.Drawing;
using Spire.Presentation;

namespace PhPPTGen.phCommand.phText {
    public class PhTextContentCommand : PhCommand {
        public PhTextContentCommand() {

        }

        public override object Exec(params object[] parameters) {
            Console.WriteLine("PPTGen phCommand: generate text shape for ppt");
            var req = (phModel.PhRequest)parameters[0];
            var jobid = req.jobid;
            var text = req.text;

            /**
             * 1. go to the tmp dir
             */
            var fct = phCommandFactory.PhCommandFactory.GetInstance();
            var tmpDir = fct.GetTmpDictionary();
            var workingPath = tmpDir + "\\" + jobid;
            var file_path = workingPath + "\\result.pptx";

            /**
             * 2. set text shapes in the slider
             */
            Presentation presentation = new Presentation();
            presentation.LoadFromFile(file_path);

            if (presentation.Slides.Count <= text.slider) {
                presentation.Slides.Append();
            }

            //append new shape
            IAutoShape shape = presentation.Slides[text.slider].Shapes.AppendShape(ShapeType.Rectangle, new RectangleF(50, 70, 450, 150));
            shape.Fill.FillType = Spire.Presentation.Drawing.FillFormatType.None;
            shape.ShapeStyle.LineColor.Color = Color.White;

            shape.TextFrame.Paragraphs[0].Alignment = TextAlignmentType.Left;
            shape.TextFrame.Paragraphs[0].Indent = 50;
            shape.TextFrame.Paragraphs[0].LineSpacing = 150;
            shape.TextFrame.Text = text.content;

            shape.TextFrame.Paragraphs[0].TextRanges[0].LatinFont = new TextFont("Arial Rounded MT Bold");
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.FillType = Spire.Presentation.Drawing.FillFormatType.Solid;
            shape.TextFrame.Paragraphs[0].TextRanges[0].Fill.SolidColor.Color = Color.Black;

            //save the document
            presentation.SaveToFile(file_path, FileFormat.Pptx2010);

            return null;
        }
    }
}
