using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PhPPTGen.phCommand.phPpt;
using PhPPTGen.phOpenxml;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Drawing = DocumentFormat.OpenXml.Drawing;

namespace PhPPTGen.phCommand.phChart {
	class PhPPTImportChartCommand : PhCommand {

		public override object Exec(params object[] parameters) {
			Console.WriteLine("PPTGen PhCommand: Insert table into pptx");
			var req = (phModel.PhRequest)parameters[0];
			var jobid = req.jobid;
			var e2c = req.e2c;

			//清除格式信息
			PhExcelFormatConfig.GetInstans().OneExcelOver();

			/**
             * 1. go to the work place
             */
			var fct = phCommandFactory.PhCommandFactory.GetInstance();
			var tmpDir = fct.GetTmpDictionary();
			var workingPath = tmpDir + "\\" + jobid;
			string workbookKey = jobid + e2c.name;


			/**
             * 2. get workbook 
             */
			var ePath = workingPath + "\\" + e2c.name + ".xlsx";
			if (!File.Exists(ePath)) {
				throw new Exception("Excel name is not exists");
			}



			/**
             * 3. put the excel into pptx
             */
			var ppt_path = workingPath + "\\" + "result.pptx";
			

			using (PresentationDocument myPresDoc = PresentationDocument.Open(ppt_path, true)) {
				PresentationPart pptPart = myPresDoc.PresentationPart;
				if (pptPart.Presentation.SlideIdList.Count() < e2c.slider) {
					PhOpenxmlPPTHandler.GetInstance().InsertNewSlide(myPresDoc, e2c.slider, "");
				}

				var sld = pptPart.SlideParts.Last();
				// Get the p:cSld element.Shape tree of a slide (spTree) is a child element of cSld
				// because all slide types may contain a shape tree. Hence we get the CSld and get the ShapeTree.
				CommonSlideData comSlddata = sld.Slide.CommonSlideData;
				//p:spTree Element
				ShapeTree shapeTree = comSlddata.ShapeTree;
				//This element specifies the existence of a graphics frame.                
				// p:graphicFrame Element
				GraphicFrame graphicFrame = new GraphicFrame();

				//This element specifies all non-visual properties for a graphic frame
				//p:nvGraphicFramePr element
				NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new NonVisualGraphicFrameProperties();

				//This element specifies non-visual canvas properties. Currently I have hardcoded the id and the name.
				// Feel free to choose a unqiue id and the name.
				// p:cNvPr Element
				NonVisualDrawingProperties nonVisualDrawingProperties = new NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Chart 3" };

				//This element specifies the non-visual drawing properties for a graphic frame. These non-visual properties are properties that the
				//generating application would utilize when rendering the slide surface.
				//p:cNvGraphicFramePr
				NonVisualGraphicFrameDrawingProperties nonVisualGraphicFrameDrawingProperties = new NonVisualGraphicFrameDrawingProperties();

				//p:nvPr Element
				ApplicationNonVisualDrawingProperties applicationnonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

				nonVisualGraphicFrameProperties.Append(nonVisualDrawingProperties);
				nonVisualGraphicFrameProperties.Append(nonVisualGraphicFrameDrawingProperties);
				nonVisualGraphicFrameProperties.Append(applicationnonVisualDrawingProperties);

				//This element specifies the transform to be applied to the corresponding graphic frame. (2D Transform for Graphic Frame)
				//p:xfrm Element
				Transform transform = new Transform();

				//Notice that I have hardcoded the dimensions of the chart I will be importing.
				//Feel free to choose any dimension that works best for your document
				//a:off Element.Specifies the location of the bounding box of an object
				Drawing.Offset offset = new Drawing.Offset() { X = e2c.pos[0] * 12709L, Y = e2c.pos[1] * 12709L };
				//a:ext Element. Specifies the size of the bounding box enclosing the referenced object
				Drawing.Extents extents = new Drawing.Extents() { Cx = e2c.pos[2] * 14081L, Cy = e2c.pos[3] * 11430L };

				// Add position element to transform element.
				transform.Append(offset);
				transform.Append(extents);

				// Add Transform and the non-visual properties to GraphicFrame element.
				graphicFrame.Append(nonVisualGraphicFrameProperties);
				graphicFrame.Append(transform);

				// Open SpreadSheet
				using (SpreadsheetDocument mySpreadsheet = SpreadsheetDocument.Open(ePath, true)) {
					//Get all the appropriate parts
					WorkbookPart workbookPart = mySpreadsheet.WorkbookPart;

					//生成chart在excel中
					PhExcelHandler.GetInstance().InsertChartIntoExcel(workbookPart, e2c.chartType);
					//WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById("rId1");
					WorksheetPart worksheetPart = mySpreadsheet.WorkbookPart.WorksheetParts.ElementAt(0);
					DrawingsPart drawingPart = worksheetPart.DrawingsPart;
					//ChartPart chartPart = (ChartPart)drawingPart.GetPartById("rId1");
					ChartPart chartPart = mySpreadsheet.WorkbookPart.WorksheetParts.ElementAt(0).DrawingsPart.ChartParts.
										ElementAt(0) as ChartPart;

					//Add a Chart Part to the Slide and get the relationship
					ChartPart importedChartPart = sld.AddPart<ChartPart>(chartPart);
					string relId = sld.GetIdOfPart(importedChartPart);

					//This element describes a single graphical object frame for a spreadsheet which contains a graphical object.
					DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame frame = drawingPart.WorksheetDrawing.Descendants<DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame>().First();
					string chartName = frame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name;

					//Clone this node so we can add it to my slide
					Drawing.Graphic clonedGraphic = (Drawing.Graphic)frame.Graphic.CloneNode(true);
					Drawing.Charts.ChartReference c = clonedGraphic.GraphicData.GetFirstChild<Drawing.Charts.ChartReference>();
					c.Id = relId;

					//Add it
					graphicFrame.Append(clonedGraphic);
					shapeTree.Append(graphicFrame);
					myPresDoc.Close();
				}
			}

				return null;
		}
	}
}
