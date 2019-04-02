using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Newtonsoft.Json.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.IO;

namespace PhPPTGen.phOpenxml.PhShapeTreeChilren {
	class PhPictureHandler : PhBaseHandler {
		protected override OpenXmlElement AppendDefaultElement(JToken format, params object[] paras) {
			var sld = (SlidePart)paras[0];
			var path = PhConfigHandler.GetInstance().path;
			ImagePart imagePart1 = sld.AddNewPart<ImagePart>((string)format["type"], (string)format["id"]);
			using (var data = new FileStream(path + (string)format["path"], FileMode.Open, FileAccess.ReadWrite)) {
				imagePart1.FeedData(data);
			}


			P.Picture picture1 = new P.Picture();

			P.NonVisualPictureProperties nonVisualPictureProperties1 = new P.NonVisualPictureProperties();

			P.NonVisualDrawingProperties nonVisualDrawingProperties2 = new P.NonVisualDrawingProperties() {
				Id = (uint)format["index"],
				Name = (string)format["name"]
			};

			A.NonVisualDrawingPropertiesExtensionList nonVisualDrawingPropertiesExtensionList1 = new A.NonVisualDrawingPropertiesExtensionList();

			A.NonVisualDrawingPropertiesExtension nonVisualDrawingPropertiesExtension1 = new A.NonVisualDrawingPropertiesExtension() { Uri = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}" };

			OpenXmlUnknownElement openXmlUnknownElement1 = OpenXmlUnknownElement.CreateOpenXmlUnknownElement("<a16:creationId xmlns:a16=\"http://schemas.microsoft.com/office/drawing/2014/main\" id=\"{27DB384F-45F7-47C0-BB49-BD5AD74BC2EC}\" />");

			nonVisualDrawingPropertiesExtension1.Append(openXmlUnknownElement1);

			nonVisualDrawingPropertiesExtensionList1.Append(nonVisualDrawingPropertiesExtension1);

			nonVisualDrawingProperties2.Append(nonVisualDrawingPropertiesExtensionList1);

			P.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new P.NonVisualPictureDrawingProperties();
			A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

			nonVisualPictureDrawingProperties1.Append(pictureLocks1);
			ApplicationNonVisualDrawingProperties applicationNonVisualDrawingProperties2 = new ApplicationNonVisualDrawingProperties();

			nonVisualPictureProperties1.Append(nonVisualDrawingProperties2);
			nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);
			nonVisualPictureProperties1.Append(applicationNonVisualDrawingProperties2);

			P.BlipFill blipFill1 = new P.BlipFill();

			A.Blip blip1 = new A.Blip() { Embed = (string)format["id"] };

			A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

			A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

			A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
			useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

			blipExtension1.Append(useLocalDpi1);

			blipExtensionList1.Append(blipExtension1);

			blip1.Append(blipExtensionList1);

			A.Stretch stretch1 = new A.Stretch();
			A.FillRectangle fillRectangle1 = new A.FillRectangle();

			stretch1.Append(fillRectangle1);

			blipFill1.Append(blip1);
			blipFill1.Append(stretch1);

			P.ShapeProperties shapeProperties1 = new P.ShapeProperties();

			A.Transform2D transform2D1 = new A.Transform2D();
			A.Offset offset2 = new A.Offset() { X = (long)(((double)format["x"]) / 0.00000278), Y = (long)(((double)format["y"]) / 0.00000278) };
			A.Extents extents2 = new A.Extents() { Cx = (long)(((double)format["cx"]) / 0.00000278), Cy = (long)(((double)format["cy"]) / 0.00000278) };

			transform2D1.Append(offset2);
			transform2D1.Append(extents2);

			A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
			A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

			presetGeometry1.Append(adjustValueList1);

			shapeProperties1.Append(transform2D1);
			shapeProperties1.Append(presetGeometry1);

			picture1.Append(nonVisualPictureProperties1);
			picture1.Append(blipFill1);
			picture1.Append(shapeProperties1);
			return picture1;
		}
	}
}
