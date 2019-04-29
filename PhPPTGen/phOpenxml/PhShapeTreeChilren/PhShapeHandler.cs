using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using Newtonsoft.Json.Linq;
using PhPPTGen.phOpenxml.phTextHandler;


namespace PhPPTGen.phOpenxml.PhShapeTreeChilren {
	class PhShapeHandler : PhBaseHandler {
		//todo: 还需要进一步分离
		protected override OpenXmlElement AppendDefaultElement(JToken format, params object[] paras) {
			var textHandlerMap = new Dictionary<string, PhRunTextHandler>() {
				{ "run", new PhRunTextHandler() }, {"field",new PhFieldTextHandler() }
			};
			P.Shape shape = new P.Shape {
				NonVisualShapeProperties = new P.NonVisualShapeProperties(
					new P.NonVisualDrawingProperties() { Id = (uint)format["index"], Name = (string)format["name"] },
					new P.NonVisualShapeDrawingProperties() { TextBox = true },
					new ApplicationNonVisualDrawingProperties()),
				ShapeProperties = new P.ShapeProperties(
					new A.Transform2D(new A.Offset() { X = (long)(((double)format["x"]) / 0.00000278), Y = (long)(((double)format["y"]) / 0.00000278) },
					new A.Extents() { Cx = (long)(((double)format["cx"]) / 0.00000278), Cy = (long)(((double)format["cy"]) / 0.00000278) }),
					new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
					new A.SolidFill(
						new A.RgbColorModelHex(
							new A.Alpha() { Val = int.Parse((string)format["alpha"]) }
						) { Val = (string)format["backColor"] }
					),
					getLineType((string)format["lineType"] ?? "line", new A.Outline(
						new A.SolidFill(
							new A.RgbColorModelHex(
								new A.Alpha() { Val = int.Parse((string)format["lineAlpha"] ?? "0") }
							) { Val = (string)format["lineColor"] ?? "000000" }
						)
					) { Width = int.Parse((string)format["lineWidth"] ?? "19050") }				
				)),
				TextBody = new P.TextBody(new A.BodyProperties() { Rotation = int.Parse((string)format["rotation"] ?? "0"), Anchor = A.TextAnchoringTypeValues.Center },
					new A.ListStyle())
			};
			foreach (JToken paragraphFormat in (JArray)format["content"]) {
				A.Paragraph paragraph = new Paragraph(
					new A.ParagraphProperties() {
						Alignment = (A.TextAlignmentTypeValues)Enum.Parse(typeof(A.TextAlignmentTypeValues), (string)paragraphFormat["alignment"])
					});
				foreach (JToken runFormat in (JArray)paragraphFormat["run"]) {

					paragraph.Append(textHandlerMap[(string)runFormat["type"]].CreateRun(runFormat));
				}
				shape.TextBody.Append(paragraph);
			}
			return shape;
		}

		private A.Outline getLineType(string typeName, A.Outline outline) {

			switch (typeName) {
				case "dash": 
					outline.Append(new A.PresetDash() { Val = A.PresetLineDashValues.Dash });
				break;
				case "line":
				break;
			}
			return outline;
		}
	}
}
