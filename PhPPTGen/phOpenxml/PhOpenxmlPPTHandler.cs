using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P = DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using System.Text.RegularExpressions;
using Newtonsoft.Json.Linq;
using System.IO;
using Newtonsoft.Json;
using PhPPTGen.phOpenxml.phTextHandler;

namespace PhPPTGen.phOpenxml {
	class PhOpenxmlPPTHandler {

		private static readonly PhOpenxmlPPTHandler Instance = new PhOpenxmlPPTHandler();
		private readonly Dictionary<string, JToken> FormatMap;
		private readonly Dictionary<string, PhRunTextHandler> textHandlerMap;

		public static PhOpenxmlPPTHandler GetInstance() {
			return Instance;
		}

		public void CreatePresentation(string filepath) {
			// Create a presentation at a specified file path. The presentation document type is pptx, by default.
			PresentationDocument presentationDoc = PresentationDocument.Create(filepath, PresentationDocumentType.Presentation);
			PresentationPart presentationPart = presentationDoc.AddPresentationPart();
			presentationPart.Presentation = new Presentation();

			CreatePresentationParts(presentationPart);

			//Close the presentation handle
			presentationDoc.Close();
		}

		// Insert the specified slide into the presentation at the specified position.
		public void InsertNewSlide(PresentationDocument presentationDocument, int position, string slideTitle, string type = "defult") {

			if (presentationDocument == null) {
				throw new ArgumentNullException("presentationDocument");
			}


			if (slideTitle == null) {
				throw new ArgumentNullException("slideTitle");
			}

			PresentationPart presentationPart = presentationDocument.PresentationPart;

			// Verify that the presentation is not empty.
			if (presentationPart == null) {
				throw new InvalidOperationException("The presentation document is empty.");
			}

			// Declare and instantiate a new slide.
			Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
			//uint drawingObjectId = 1;

			// Construct the slide content.            
			// Specify the non-visual properties of the new slide.
			P.NonVisualGroupShapeProperties nonVisualProperties = slide.CommonSlideData.ShapeTree.AppendChild(new P.NonVisualGroupShapeProperties());
			nonVisualProperties.NonVisualDrawingProperties = new P.NonVisualDrawingProperties() { Id = 1, Name = "" };
			nonVisualProperties.NonVisualGroupShapeDrawingProperties = new P.NonVisualGroupShapeDrawingProperties();
			nonVisualProperties.ApplicationNonVisualDrawingProperties = new ApplicationNonVisualDrawingProperties();

			// 加入预设shape
			slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());



			JToken typeFormat = null;
			using (StreamReader reader = File.OpenText(PhConfigHandler.GetInstance().path + PhConfigHandler.GetInstance().configMap["pptType"].Value<string>())) {
				typeFormat = JToken.ReadFrom(new JsonTextReader(reader));
			}
			foreach (JToken format in (JArray)typeFormat[type]) {
				P.Shape shape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
				shape.NonVisualShapeProperties = new P.NonVisualShapeProperties
				(new P.NonVisualDrawingProperties() { Id = (uint)format["index"], Name = (string)format["name"] },
				new P.NonVisualShapeDrawingProperties() { TextBox = true },
				new ApplicationNonVisualDrawingProperties(
					//new PlaceholderShape() {
					//Type = (PlaceholderValues)Enum.Parse(typeof(PlaceholderValues), (string)format["placeholderType"]) }
					));
				shape.ShapeProperties = new P.ShapeProperties(
					new A.Transform2D(new A.Offset() { X = (long)(((double)format["x"]) / 0.00000278), Y = (long)(((double)format["y"]) / 0.00000278) }, 
					new A.Extents() { Cx = (long)(((double)format["cx"])/0.00000278), Cy = (long)(((double)format["cy"]) / 0.00000278) }),
					new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
				new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() {
						Val = int.Parse((string)format["alpha"])
					}) { Val = (string)format["backColor"] }));
				shape.TextBody = new P.TextBody(new A.BodyProperties(),
						new A.ListStyle());
				foreach (JToken paragraphFormat in (JArray)format["content"]) {
					A.Paragraph paragraph = new Paragraph(
						new A.ParagraphProperties() {
							Alignment = (A.TextAlignmentTypeValues)Enum.Parse(typeof(A.TextAlignmentTypeValues), (string)paragraphFormat["alignment"])
						});
					foreach(JToken runFormat in (JArray)paragraphFormat["run"]) {
						
						paragraph.Append(textHandlerMap[(string)runFormat["type"]].CreateRun(runFormat));
					}
					shape.TextBody.Append(paragraph);
				}
			}



			// Create the slide part for the new slide.
			SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
			slidePart.AddNewPart<VmlDrawingPart>("rId2");

			// Save the new slide part.
			slide.Save(slidePart);

			// Modify the slide ID list in the presentation part.
			// The slide ID list should not be null.
			SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

			// Find the highest slide ID in the current list.
			uint maxSlideId = 1;
			SlideId prevSlideId = null;

			foreach (SlideId slideId in slideIdList.ChildElements) {
				if (slideId.Id > maxSlideId) {
					maxSlideId = slideId.Id;
				}

				position--;
				if (position == 0) {
					prevSlideId = slideId;
				}

			}

			maxSlideId++;

			// Get the ID of the previous slide.
			SlidePart lastSlidePart;

			if (prevSlideId != null) {
				lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
			} else {
				lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
			}

			// Use the same slide layout as that of the previous slide.
			if (null != lastSlidePart.SlideLayoutPart) {
				slidePart.AddPart(lastSlidePart.SlideLayoutPart);
			}

			// Insert the new slide into the slide list after the previous slide.
			SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
			newSlideId.Id = maxSlideId;
			newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

			// Save the modified presentation.
			presentationPart.Presentation.Save();
		}

		public void InsertText(PresentationDocument presentationDocument, int position, string content, int[] pos) {

			PresentationPart pptPart = presentationDocument.PresentationPart;
			if (pptPart.Presentation.SlideIdList.Count() < position) {
				PhOpenxmlPPTHandler.GetInstance().InsertNewSlide(presentationDocument, position, "");
			}

			var slide = pptPart.SlideParts.Last().Slide;
;
			uint drawingObjectId = (uint)slide.CommonSlideData.ShapeTree.ChildElements.Count();
			// Declare and instantiate the body shape of the new slide.
			P.Shape bodyShape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
			drawingObjectId++;

			// Specify the required shape properties for the body shape.
			bodyShape.NonVisualShapeProperties = new P.NonVisualShapeProperties(new P.NonVisualDrawingProperties() { Id = drawingObjectId, Name = "Content Placeholder" },
					new P.NonVisualShapeDrawingProperties(new A.ShapeLocks() { NoGrouping = true }),
					new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Index = 1 }));
			bodyShape.ShapeProperties = new P.ShapeProperties(new A.Transform2D(new A.Offset() { X = pos[0], Y = pos[1] }, new A.Extents() { Cx = pos[2], Cy = pos[3] }));

			// Specify the text of the body shape.
			P.TextBody textBody = new P.TextBody(new A.BodyProperties(),
					new A.ListStyle());
			foreach (Match m in new Regex(@"(?<=(#{#))[\s\S]*?(?=(#}#))").Matches(content)) {
				string paragraphContent = m.Value;
				A.Paragraph paragraph = new Paragraph();
				JToken paragraphCss = FormatMap["pptParagraphFormat"][new Regex(@"(?<=(#P#))[\s\S]*").Match(paragraphContent).Value];

				//正则读取content里面的段落格式代号，根据代号在json中取得具体格式
				paragraph.Append(new A.ParagraphProperties() {
					Alignment = (A.TextAlignmentTypeValues)Enum.Parse(typeof(A.TextAlignmentTypeValues), (string)paragraphCss["Alignment"]) });

				foreach (Match match in new Regex(@"(?<=(#\[#))[\s\S]*?(?=(#\]#))").Matches(new Regex(@"[\s\S]*(?=(#P#))").Match(paragraphContent).Value)) {
					string runContent = match.Value;
					A.Text text = new A.Text() { Text = new Regex(@"[\s\S]*(?=(#C#))").Match(runContent).Value };
					JToken runCss = FormatMap["pptFontFormat"][new Regex(@"(?<=(#C#))[\s\S]*").Match(runContent).Value];
					//正则读取content里面的字段格式代号，根据代号在json中取得具体格式
					A.SolidFill solidFill = new A.SolidFill(new A.RgbColorModelHex() { Val = new HexBinaryValue((string)runCss["Color"]) });
					A.RunProperties runProperties = new A.RunProperties(
						solidFill,
						new A.LatinFont() { Typeface = (string)runCss["Font"] }, new A.ComplexScriptFont() { Typeface = (string)runCss["Font"] }
						) { Language = "en-US", AlternativeLanguage = "zh-CN",
						FontSize = int.Parse((string)runCss["FontSize"]) * 100,
						Bold = Boolean.Parse((string)runCss["Bold"]), Dirty = false };
					paragraph.Append(new A.Run(runProperties, text));
					
				}
				textBody.Append(paragraph);
			}
			bodyShape.TextBody = textBody;
		}

		public void InsertExcel(PresentationDocument presentationDocument, string excelPath, string emfPath, int[] pos) {
			PresentationPart pptPart = presentationDocument.PresentationPart;

			var sld = pptPart.SlideParts.Last();

			VmlDrawingPart vmlDrawingPart1 = sld.GetPartsOfType<VmlDrawingPart>().First();
			//GenerateVmlDrawingPart1Content(vmlDrawingPart1);
			string imagePartId = "imgId" + sld.GetPartsCountOfType<ImagePart>();
			ImagePart imagePart = vmlDrawingPart1.AddNewPart<ImagePart>("image/x-emf", imagePartId);
			using(var data = new FileStream(emfPath, FileMode.Open, FileAccess.ReadWrite)) {
				imagePart.FeedData(data);
			}

			sld.AddPart(imagePart, imagePartId);

			string embId = "embId" + sld.GetPartsCountOfType<EmbeddedPackagePart>();
			EmbeddedPackagePart embeddedPackagePart = sld.AddNewPart<EmbeddedPackagePart>(
				"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", embId);
			embeddedPackagePart.FeedData(new FileStream(excelPath, FileMode.Open, FileAccess.ReadWrite));
			ShapeTree shapeTree = sld.Slide.CommonSlideData.ShapeTree;
			P.GraphicFrame graphicFrame = new P.GraphicFrame();
			shapeTree.Append(graphicFrame);
			P.NonVisualGraphicFrameProperties nonVisualGraphicFrameProperties = new P.NonVisualGraphicFrameProperties(
				new P.NonVisualDrawingProperties() {
					Id = (UInt32Value)(uint)shapeTree.ChildElements.Count(),
					Name = "Chart" + shapeTree.ChildElements.Count()
				}, new P.NonVisualGraphicFrameDrawingProperties(), new ApplicationNonVisualDrawingProperties());

			Transform transform1 = new Transform(new A.Offset() { X = pos[0], Y = pos[1] }, new A.Extents() { Cx = pos[2], Cy = pos[3] });
			A.Graphic graphic = new A.Graphic();
			A.GraphicData graphicData = new A.GraphicData() { Uri = "http://schemas.openxmlformats.org/presentationml/2006/ole" };
			AlternateContent alternateContent = new AlternateContent();
			alternateContent.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
			AlternateContentChoice alternateContentChoice = new AlternateContentChoice(new OleObject(new OleObjectEmbed()) {
				Name = "Worksheet",
				Id = embId,
				ImageWidth = pos[2],
				ImageHeight = pos[3],
				ProgId = "Excel.Sheet.12"
			}) { Requires = "v" };
			alternateContentChoice.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
			alternateContent.Append(alternateContentChoice);

			AlternateContentFallback alternateContentFallback = new AlternateContentFallback();
			OleObject oleObject = new OleObject(new OleObjectEmbed()) {
				Name = "Worksheet",
				Id = embId,
				ImageWidth = pos[2],
				ImageHeight = pos[3],
				ProgId = "Excel.Sheet.12"
			};

			P.Picture picture = new P.Picture();

			P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
				//动态id
				new P.NonVisualDrawingProperties() { Id = (UInt32Value)0U, Name = "" },
				new P.NonVisualPictureDrawingProperties(), new ApplicationNonVisualDrawingProperties());


			P.BlipFill blipFill = new P.BlipFill(
				new A.Blip() { Embed = imagePartId },
				new A.Stretch(new A.FillRectangle()));

			P.ShapeProperties shapeProperties = new P.ShapeProperties();

			A.Transform2D transform2D = new A.Transform2D(
				new A.Offset() { X = pos[0], Y = pos[1] },
				new A.Extents() { Cx = pos[2], Cy = pos[3] });

			A.PresetGeometry presetGeometry = new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle };

			shapeProperties.Append(transform2D);
			shapeProperties.Append(presetGeometry);

			picture.Append(nonVisualPictureProperties);
			picture.Append(blipFill);
			picture.Append(shapeProperties);

			oleObject.Append(picture);
			alternateContentFallback.Append(oleObject);

			alternateContent.Append(alternateContentFallback);


			graphicData.Append(alternateContent);
			graphic.Append(graphicData);
			graphicFrame.Append(nonVisualGraphicFrameProperties, transform1, graphic);
			sld.Slide.Save();
		}

		private void CreatePresentationParts(PresentationPart presentationPart) {
			SlideMasterIdList slideMasterIdList1 = new SlideMasterIdList(new SlideMasterId() { Id = (UInt32Value)2147483648U, RelationshipId = "rId1" });
			SlideIdList slideIdList1 = new SlideIdList(new SlideId() { Id = (UInt32Value)256U, RelationshipId = "rId2" });
			SlideSize slideSize1 = new SlideSize() { Cx = 9144000, Cy = 6858000, Type = SlideSizeValues.Screen4x3 };
			NotesSize notesSize1 = new NotesSize() { Cx = 6858000, Cy = 9144000 };
			DefaultTextStyle defaultTextStyle1 = new DefaultTextStyle();

			presentationPart.Presentation.Append(slideMasterIdList1, slideIdList1, slideSize1, notesSize1, defaultTextStyle1);

			SlidePart slidePart1;
			SlideLayoutPart slideLayoutPart1;
			SlideMasterPart slideMasterPart1;
			ThemePart themePart1;


			slidePart1 = CreateSlidePart(presentationPart);
			slideLayoutPart1 = CreateSlideLayoutPart(slidePart1);
			slideMasterPart1 = CreateSlideMasterPart(slideLayoutPart1);
			themePart1 = CreateTheme(slideMasterPart1);

			slideMasterPart1.AddPart(slideLayoutPart1, "rId1");
			presentationPart.AddPart(slideMasterPart1, "rId1");
			presentationPart.AddPart(themePart1, "rId5");
			//slideIdList1.RemoveChild(slideIdList1.GetFirstChild<SlideId>());
		}

		private SlidePart CreateSlidePart(PresentationPart presentationPart) {
			SlidePart slidePart = presentationPart.AddNewPart<SlidePart>("rId2");
			slidePart.Slide = new Slide(
					new CommonSlideData(
						new ShapeTree(
							new P.NonVisualGroupShapeProperties(
								new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
								new P.NonVisualGroupShapeDrawingProperties(),
								new ApplicationNonVisualDrawingProperties()),
							new GroupShapeProperties(new TransformGroup()))),
					new ColorMapOverride(new MasterColorMapping()));
			slidePart.AddNewPart<VmlDrawingPart>("rId2");
			// 加入预设title shape
			var slide = slidePart.Slide;
			//slide.CommonSlideData.ShapeTree.AppendChild(new GroupShapeProperties());

			JToken typeFormat = null;
			using (StreamReader reader = File.OpenText(PhConfigHandler.GetInstance().path + PhConfigHandler.GetInstance().configMap["pptType"].Value<string>())) {
				typeFormat = JToken.ReadFrom(new JsonTextReader(reader));
			}
			foreach (JToken format in (JArray)typeFormat["title"]) {
				P.Shape shape = slide.CommonSlideData.ShapeTree.AppendChild(new P.Shape());
				shape.NonVisualShapeProperties = new P.NonVisualShapeProperties
				(new P.NonVisualDrawingProperties() { Id = (uint)format["index"], Name = (string)format["name"] },
				new P.NonVisualShapeDrawingProperties() { TextBox = true },
				new ApplicationNonVisualDrawingProperties(
					//new PlaceholderShape() {
					//Type = (PlaceholderValues)Enum.Parse(typeof(PlaceholderValues), (string)format["placeholderType"]) }
					));
				shape.ShapeProperties = new P.ShapeProperties(
					new A.Transform2D(new A.Offset() { X = (long)(((double)format["x"]) / 0.00000278), Y = (long)(((double)format["y"]) / 0.00000278) },
					new A.Extents() { Cx = (long)(((double)format["cx"]) / 0.00000278), Cy = (long)(((double)format["cy"]) / 0.00000278) }),
					new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle },
				new A.SolidFill(new A.RgbColorModelHex(new A.Alpha() {
					Val = int.Parse((string)format["alpha"])
				}) { Val = (string)format["backColor"] }));
				shape.TextBody = new P.TextBody(new A.BodyProperties(),
						new A.ListStyle());
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
			}
			return slidePart;
		}

		//官方给的祖传代码
		private SlideLayoutPart CreateSlideLayoutPart(SlidePart slidePart1) {
			SlideLayoutPart slideLayoutPart1 = slidePart1.AddNewPart<SlideLayoutPart>("rId1");
			SlideLayout slideLayout = new SlideLayout(
			new CommonSlideData(new ShapeTree(
			  new P.NonVisualGroupShapeProperties(
			  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
			  new P.NonVisualGroupShapeDrawingProperties(),
			  new ApplicationNonVisualDrawingProperties()),
			  new GroupShapeProperties(new TransformGroup()))),
			new ColorMapOverride(new MasterColorMapping()));
			slideLayoutPart1.SlideLayout = slideLayout;
			return slideLayoutPart1;
		}

		//官方给的祖传代码
		private SlideMasterPart CreateSlideMasterPart(SlideLayoutPart slideLayoutPart1) {
			SlideMasterPart slideMasterPart1 = slideLayoutPart1.AddNewPart<SlideMasterPart>("rId1");
			SlideMaster slideMaster = new SlideMaster(
			new CommonSlideData(new ShapeTree(
			  new P.NonVisualGroupShapeProperties(
			  new P.NonVisualDrawingProperties() { Id = (UInt32Value)1U, Name = "" },
			  new P.NonVisualGroupShapeDrawingProperties(),
			  new ApplicationNonVisualDrawingProperties()),
			  new GroupShapeProperties(new TransformGroup()))),
			new P.ColorMap() { Background1 = A.ColorSchemeIndexValues.Light1, Text1 = A.ColorSchemeIndexValues.Dark1, Background2 = A.ColorSchemeIndexValues.Light2, Text2 = A.ColorSchemeIndexValues.Dark2, Accent1 = A.ColorSchemeIndexValues.Accent1, Accent2 = A.ColorSchemeIndexValues.Accent2, Accent3 = A.ColorSchemeIndexValues.Accent3, Accent4 = A.ColorSchemeIndexValues.Accent4, Accent5 = A.ColorSchemeIndexValues.Accent5, Accent6 = A.ColorSchemeIndexValues.Accent6, Hyperlink = A.ColorSchemeIndexValues.Hyperlink, FollowedHyperlink = A.ColorSchemeIndexValues.FollowedHyperlink },
			new SlideLayoutIdList(new SlideLayoutId() { Id = (UInt32Value)2147483649U, RelationshipId = "rId1" }),
			new TextStyles(new TitleStyle(), new BodyStyle(), new OtherStyle()));
			slideMasterPart1.SlideMaster = slideMaster;

			return slideMasterPart1;
		}

		//官方给的祖传代码
		private ThemePart CreateTheme(SlideMasterPart slideMasterPart1) {

			ThemePart themePart1 = slideMasterPart1.AddNewPart<ThemePart>("rId5");
			A.Theme theme1 = new A.Theme() { Name = "Office Theme" };

			A.ThemeElements themeElements1 = new A.ThemeElements(
			new A.ColorScheme(
			  new A.Dark1Color(new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" }),
			  new A.Light1Color(new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" }),
			  new A.Dark2Color(new A.RgbColorModelHex() { Val = "1F497D" }),
			  new A.Light2Color(new A.RgbColorModelHex() { Val = "EEECE1" }),
			  new A.Accent1Color(new A.RgbColorModelHex() { Val = "4F81BD" }),
			  new A.Accent2Color(new A.RgbColorModelHex() { Val = "C0504D" }),
			  new A.Accent3Color(new A.RgbColorModelHex() { Val = "9BBB59" }),
			  new A.Accent4Color(new A.RgbColorModelHex() { Val = "8064A2" }),
			  new A.Accent5Color(new A.RgbColorModelHex() { Val = "4BACC6" }),
			  new A.Accent6Color(new A.RgbColorModelHex() { Val = "F79646" }),
			  new A.Hyperlink(new A.RgbColorModelHex() { Val = "0000FF" }),
			  new A.FollowedHyperlinkColor(new A.RgbColorModelHex() { Val = "800080" })) { Name = "Office" },
			  new A.FontScheme(
			  new A.MajorFont(
			  new A.LatinFont() { Typeface = "Calibri" },
			  new A.EastAsianFont() { Typeface = "" },
			  new A.ComplexScriptFont() { Typeface = "" }),
			  new A.MinorFont(
			  new A.LatinFont() { Typeface = "Calibri" },
			  new A.EastAsianFont() { Typeface = "" },
			  new A.ComplexScriptFont() { Typeface = "" })) { Name = "Office" },
			  new A.FormatScheme(
			  new A.FillStyleList(
			  new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
			  new A.GradientFill(
				new A.GradientStopList(
				new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 50000 },
				  new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
				new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 37000 },
				 new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 35000 },
				new A.GradientStop(new A.SchemeColor(new A.Tint() { Val = 15000 },
				 new A.SaturationModulation() { Val = 350000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 100000 }
				),
				new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
			  new A.NoFill(),
			  new A.PatternFill(),
			  new A.GroupFill()),
			  new A.LineStyleList(
			  new A.Outline(
				new A.SolidFill(
				new A.SchemeColor(
				  new A.Shade() { Val = 95000 },
				  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
				new A.PresetDash() { Val = A.PresetLineDashValues.Solid }) {
				  Width = 9525,
				  CapType = A.LineCapValues.Flat,
				  CompoundLineType = A.CompoundLineValues.Single,
				  Alignment = A.PenAlignmentValues.Center
			  },
			  new A.Outline(
				new A.SolidFill(
				new A.SchemeColor(
				  new A.Shade() { Val = 95000 },
				  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
				new A.PresetDash() { Val = A.PresetLineDashValues.Solid }) {
				  Width = 9525,
				  CapType = A.LineCapValues.Flat,
				  CompoundLineType = A.CompoundLineValues.Single,
				  Alignment = A.PenAlignmentValues.Center
			  },
			  new A.Outline(
				new A.SolidFill(
				new A.SchemeColor(
				  new A.Shade() { Val = 95000 },
				  new A.SaturationModulation() { Val = 105000 }) { Val = A.SchemeColorValues.PhColor }),
				new A.PresetDash() { Val = A.PresetLineDashValues.Solid }) {
				  Width = 9525,
				  CapType = A.LineCapValues.Flat,
				  CompoundLineType = A.CompoundLineValues.Single,
				  Alignment = A.PenAlignmentValues.Center
			  }),
			  new A.EffectStyleList(
			  new A.EffectStyle(
				new A.EffectList(
				new A.OuterShadow(
				  new A.RgbColorModelHex(
				  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
			  new A.EffectStyle(
				new A.EffectList(
				new A.OuterShadow(
				  new A.RgbColorModelHex(
				  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false })),
			  new A.EffectStyle(
				new A.EffectList(
				new A.OuterShadow(
				  new A.RgbColorModelHex(
				  new A.Alpha() { Val = 38000 }) { Val = "000000" }) { BlurRadius = 40000L, Distance = 20000L, Direction = 5400000, RotateWithShape = false }))),
			  new A.BackgroundFillStyleList(
			  new A.SolidFill(new A.SchemeColor() { Val = A.SchemeColorValues.PhColor }),
			  new A.GradientFill(
				new A.GradientStopList(
				new A.GradientStop(
				  new A.SchemeColor(new A.Tint() { Val = 50000 },
					new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
				new A.GradientStop(
				  new A.SchemeColor(new A.Tint() { Val = 50000 },
					new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
				new A.GradientStop(
				  new A.SchemeColor(new A.Tint() { Val = 50000 },
					new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
				new A.LinearGradientFill() { Angle = 16200000, Scaled = true }),
			  new A.GradientFill(
				new A.GradientStopList(
				new A.GradientStop(
				  new A.SchemeColor(new A.Tint() { Val = 50000 },
					new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 },
				new A.GradientStop(
				  new A.SchemeColor(new A.Tint() { Val = 50000 },
					new A.SaturationModulation() { Val = 300000 }) { Val = A.SchemeColorValues.PhColor }) { Position = 0 }),
				new A.LinearGradientFill() { Angle = 16200000, Scaled = true }))) { Name = "Office" });

			theme1.Append(themeElements1);
			theme1.Append(new A.ObjectDefaults());
			theme1.Append(new A.ExtraColorSchemeList());

			themePart1.Theme = theme1;
			return themePart1;

		}

		private PhOpenxmlPPTHandler() {
			FormatMap = new Dictionary<string, JToken>();
			foreach (JToken jToken in PhConfigHandler.GetInstance().configMap["pptFormat"]) {
				using (StreamReader reader = File.OpenText(PhConfigHandler.GetInstance().path + jToken.First().Value<string>())) {
					FormatMap.Add(((JProperty)jToken).Name, JToken.ReadFrom(new JsonTextReader(reader)));
				}
			}
			textHandlerMap = new Dictionary<string, PhRunTextHandler>() {
				{ "run", new PhRunTextHandler() }, {"field",new PhFieldTextHandler() }
			};
		}

	}
}
