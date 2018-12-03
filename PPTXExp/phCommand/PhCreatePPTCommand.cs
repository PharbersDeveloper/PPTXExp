using System;
using Independentsoft.Office;
using Independentsoft.Office.Presentation;
using Independentsoft.Office.Presentation.Drawing;
using Independentsoft.Office.Spreadsheet;

using Independentsoft.Office.Word;
using Independentsoft.Office.Word.Drawing;
using Independentsoft.Office.Charts;
using Independentsoft.Office.Drawing;

namespace PPTXExp.phCommand {
    public class PhCreatePPTCommand : PhCommand {
        public PhCreatePPTCommand() {
        }

        CommonSlideData GetLayoutCommonSlideData() {
            //GroupShape shapeTree = new GroupShape();
            Independentsoft.Office.Presentation.Drawing.GroupShape shapeTree = new Independentsoft.Office.Presentation.Drawing.GroupShape();
            shapeTree.ID = "1";
            shapeTree.Name = "layout";

            shapeTree.ShapeProperties.TransformGroup2D = new Independentsoft.Office.Drawing.TransformGroup2D();

            CommonSlideData commonSlideData = new CommonSlideData();
            commonSlideData.ShapeTree = shapeTree;

            return commonSlideData;
        }

        //void TestPPT() {
        //    Presentation presentation = new Presentation();

        //    GroupShape shapeTree = new GroupShape();
        //    shapeTree.ID = "1";
        //    shapeTree.Name = "ShapeTree";
        //    shapeTree.ShapeProperties.TransformGroup2D = new Independentsoft.Office.Drawing.TransformGroup2D();

        //    Independentsoft.Office.Drawing.TextRun run1 = new Independentsoft.Office.Drawing.TextRun("Hello World");

        //    Independentsoft.Office.Drawing.TextParagraph paragraph1 = new Independentsoft.Office.Drawing.TextParagraph();
        //    paragraph1.Alignment = Independentsoft.Office.Drawing.TextAlignmentType.Center;
        //    paragraph1.Content.Add(run1);

        //    ShapeTextBody textBody = new ShapeTextBody();
        //    textBody.Paragraphs.Add(paragraph1);
        //    textBody.TextBodyProperties.AutoFit = true;

        //    Shape shape1 = new Shape();
        //    shape1.ID = "2";
        //    shape1.Name = "Title 1";
        //    shape1.Locking = new Independentsoft.Office.Drawing.ShapeLocking();
        //    shape1.Locking.DisallowGrouping = true;
        //    shape1.Placeholder = new Placeholder(PlaceholderType.CenteredTitle);
        //    shape1.TextBody = textBody;

        //    Unit offsetX = new Unit(2, UnitType.Inch);
        //    Unit offsetY = new Unit(1, UnitType.Inch);

        //    Unit width = new Unit(6, UnitType.Inch);
        //    Unit height = new Unit(2, UnitType.Inch);

        //    Independentsoft.Office.Drawing.RgbHexColor color = new Independentsoft.Office.Drawing.RgbHexColor();
        //    color.Value = "FF0000"; //red

        //    Independentsoft.Office.Drawing.SolidFill fill = new Independentsoft.Office.Drawing.SolidFill();
        //    fill.ColorChoice = color;

        //    Independentsoft.Office.Drawing.Outline outline = new Independentsoft.Office.Drawing.Outline();
        //    outline.LineWidth = new Unit(3, UnitType.Pixel);
        //    outline.SolidFill = fill;

        //    shape1.ShapeProperties.Transform2D = new Independentsoft.Office.Drawing.Transform2D();
        //    shape1.ShapeProperties.Transform2D.Offset = new Independentsoft.Office.Drawing.Offset(offsetX, offsetY);
        //    shape1.ShapeProperties.Transform2D.Extents = new Independentsoft.Office.Drawing.Extents(width, height);
        //    shape1.ShapeProperties.PresetGeometry = new Independentsoft.Office.Drawing.PresetGeometry(Independentsoft.Office.Drawing.ShapeType.Rectangle);
        //    shape1.ShapeProperties.Outline = outline;
        //    shape1.IsTextBox = true;

        //    shapeTree.Elements.Add(shape1);

        //    CommonSlideData commonSlideData = new CommonSlideData();
        //    commonSlideData.ShapeTree = shapeTree;

        //    SlideLayout layout1 = new SlideLayout();
        //    layout1.CommonSlideData = GetLayoutCommonSlideData();

        //    Slide slide1 = new Slide();
        //    slide1.CommonSlideData = commonSlideData;
        //    slide1.Layout = layout1;

        //    SlideMaster master1 = new SlideMaster();
        //    master1.CommonSlideData = GetLayoutCommonSlideData();
        //    master1.Layouts.Add(layout1);

        //    master1.TextStyles = new SlideMasterTextStyles();
        //    master1.TextStyles.TitleStyle = new SlideMasterTitleTextStyle();
        //    master1.TextStyles.TitleStyle.ListLevel1TextStyle.DefaultTextRunProperties.FontSize = 44;

        //    presentation.Slides.Add(slide1);
        //    presentation.SlideMasters.Add(master1);

        //    presentation.SlideSize = new SlideSize(9144000, 6858000, SlideSizeType.Screen4x3);
        //    presentation.NotesSlideSize = new NotesSlideSize(6858000, 9144000);
        //    presentation.Save("/Users/alfredyang/Desktop/output.pptx", true);
        //}

        void TestWithExcel() {

            // Create Excel Work book
            Workbook book = new Workbook();
            Worksheet sheet1 = new Worksheet();

            //set cell values
            sheet1["A2"] = new Cell("Category 1");
            sheet1["A3"] = new Cell("Category 2");
            sheet1["A4"] = new Cell("Category 3");
            sheet1["A5"] = new Cell("Category 4");

            sheet1["B1"] = new Cell("Series 1");
            sheet1["B2"] = new Cell(4.3);
            sheet1["B3"] = new Cell(2.5);
            sheet1["B4"] = new Cell(3.5);
            sheet1["B5"] = new Cell(4.5);

            sheet1["C1"] = new Cell("Series 2");
            sheet1["C2"] = new Cell(2.4);
            sheet1["C3"] = new Cell(4.4);
            sheet1["C4"] = new Cell(1.8);
            sheet1["C5"] = new Cell(2.8);

            sheet1["D1"] = new Cell("Series 3");
            sheet1["D2"] = new Cell(2);
            sheet1["D3"] = new Cell(2);
            sheet1["D4"] = new Cell(3);
            sheet1["D5"] = new Cell(5);

            book.Sheets.Add(sheet1);

            //create Bar Chart
            BarChart barChart = new BarChart();
            barChart.Direction = BarDirection.Column;
            barChart.Grouping = BarGrouping.Clustered;

            BarChartSerie serie1 = new BarChartSerie();
            serie1.Index = 0;
            serie1.Order = 0;

            serie1.SeriesText = new SeriesText();
            serie1.SeriesText.StringReference = new StringReference();
            serie1.SeriesText.StringReference.Formula = "Sheet1!$B$1";

            serie1.SeriesText.StringReference.StringCache = new StringCache();

            StringPoint seriesTextPoint11 = new StringPoint(0, "Series 1");
            serie1.SeriesText.StringReference.StringCache.StringPoints.Add(seriesTextPoint11);

            serie1.CategoryAxis = new CategoryAxis();
            serie1.CategoryAxis.StringReference = new StringReference();
            serie1.CategoryAxis.StringReference.Formula = "Sheet1!$A$2:$A$5";

            serie1.CategoryAxis.StringReference.StringCache = new StringCache();

            StringPoint categoryAxisPoint11 = new StringPoint(0, "Category 1");
            StringPoint categoryAxisPoint12 = new StringPoint(1, "Category 2");
            StringPoint categoryAxisPoint13 = new StringPoint(2, "Category 3");
            StringPoint categoryAxisPoint14 = new StringPoint(3, "Category 4");

            serie1.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint11);
            serie1.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint12);
            serie1.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint13);
            serie1.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint14);

            serie1.Values = new Values();
            serie1.Values.NumberReference = new NumberReference();
            serie1.Values.NumberReference.Formula = "Sheet1!$B$2:$B$5";

            serie1.Values.NumberReference.NumberCache = new NumberCache();
            serie1.Values.NumberReference.NumberCache.Format = "General";

            NumericPoint valuesPoint11 = new NumericPoint(0, "4.3");
            NumericPoint valuesPoint12 = new NumericPoint(1, "2.5");
            NumericPoint valuesPoint13 = new NumericPoint(2, "3.5");
            NumericPoint valuesPoint14 = new NumericPoint(3, "4.5");

            serie1.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint11);
            serie1.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint12);
            serie1.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint13);
            serie1.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint14);

            BarChartSerie serie2 = new BarChartSerie();
            serie2.Index = 1;
            serie2.Order = 1;

            serie2.SeriesText = new SeriesText();
            serie2.SeriesText.StringReference = new StringReference();
            serie2.SeriesText.StringReference.Formula = "Sheet1!$C$1";

            serie2.SeriesText.StringReference.StringCache = new StringCache();

            StringPoint seriesTextPoint21 = new StringPoint(0, "Series 2");
            serie2.SeriesText.StringReference.StringCache.StringPoints.Add(seriesTextPoint21);

            serie2.CategoryAxis = new CategoryAxis();
            serie2.CategoryAxis.StringReference = new StringReference();
            serie2.CategoryAxis.StringReference.Formula = "Sheet1!$A$2:$A$5";

            serie2.CategoryAxis.StringReference.StringCache = new StringCache();

            StringPoint categoryAxisPoint21 = new StringPoint(0, "Category 1");
            StringPoint categoryAxisPoint22 = new StringPoint(1, "Category 2");
            StringPoint categoryAxisPoint23 = new StringPoint(2, "Category 3");
            StringPoint categoryAxisPoint24 = new StringPoint(3, "Category 4");

            serie2.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint21);
            serie2.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint22);
            serie2.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint23);
            serie2.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint24);

            serie2.Values = new Values();
            serie2.Values.NumberReference = new NumberReference();
            serie2.Values.NumberReference.Formula = "Sheet1!$C$2:$C$5";

            serie2.Values.NumberReference.NumberCache = new NumberCache();
            serie2.Values.NumberReference.NumberCache.Format = "General";

            NumericPoint valuesPoint21 = new NumericPoint(0, "2.4");
            NumericPoint valuesPoint22 = new NumericPoint(1, "4.4");
            NumericPoint valuesPoint23 = new NumericPoint(2, "1.8");
            NumericPoint valuesPoint24 = new NumericPoint(3, "2.8");

            serie2.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint21);
            serie2.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint22);
            serie2.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint23);
            serie2.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint24);

            BarChartSerie serie3 = new BarChartSerie();
            serie3.Index = 2;
            serie3.Order = 2;

            serie3.SeriesText = new SeriesText();
            serie3.SeriesText.StringReference = new StringReference();
            serie3.SeriesText.StringReference.Formula = "Sheet1!$D$1";

            serie3.SeriesText.StringReference.StringCache = new StringCache();

            StringPoint seriesTextPoint31 = new StringPoint(0, "Series 3");
            serie3.SeriesText.StringReference.StringCache.StringPoints.Add(seriesTextPoint31);

            serie3.CategoryAxis = new CategoryAxis();
            serie3.CategoryAxis.StringReference = new StringReference();
            serie3.CategoryAxis.StringReference.Formula = "Sheet1!$A$2:$A$5";

            serie3.CategoryAxis.StringReference.StringCache = new StringCache();

            StringPoint categoryAxisPoint31 = new StringPoint(0, "Category 1");
            StringPoint categoryAxisPoint32 = new StringPoint(1, "Category 2");
            StringPoint categoryAxisPoint33 = new StringPoint(2, "Category 3");
            StringPoint categoryAxisPoint34 = new StringPoint(3, "Category 4");

            serie3.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint31);
            serie3.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint32);
            serie3.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint33);
            serie3.CategoryAxis.StringReference.StringCache.StringPoints.Add(categoryAxisPoint34);

            serie3.Values = new Values();
            serie3.Values.NumberReference = new NumberReference();
            serie3.Values.NumberReference.Formula = "Sheet1!$D$2:$D$5";

            serie3.Values.NumberReference.NumberCache = new NumberCache();
            serie3.Values.NumberReference.NumberCache.Format = "General";

            NumericPoint valuesPoint31 = new NumericPoint(0, "2");
            NumericPoint valuesPoint32 = new NumericPoint(1, "2");
            NumericPoint valuesPoint33 = new NumericPoint(2, "3");
            NumericPoint valuesPoint34 = new NumericPoint(3, "5");

            serie3.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint31);
            serie3.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint32);
            serie3.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint33);
            serie3.Values.NumberReference.NumberCache.NumericPoints.Add(valuesPoint34);

            barChart.Series.Add(serie1);
            barChart.Series.Add(serie2);
            barChart.Series.Add(serie3);

            barChart.FirstAxisID = 1;
            barChart.SecondAxisID = 2;

            //create ChartSpace object and add barChart to the chartSpace
            ChartSpace chartSpace = new ChartSpace();
            chartSpace.ExternalData = book.GetBytes();
            chartSpace.ExternalDataFileName = "Workbook1.xlsx";
            chartSpace.PlotArea = new PlotArea();
            chartSpace.PlotArea.Layout = new Layout();
            chartSpace.PlotArea.Charts.Add(barChart);

            ChartCategoryAxis chartCategoryAxis = new ChartCategoryAxis();
            chartCategoryAxis.ID = 1;
            chartCategoryAxis.Scaling = new Scaling();
            chartCategoryAxis.Scaling.Orientation = Independentsoft.Office.Charts.Orientation.MinimumToMaximum;
            chartCategoryAxis.Position = AxisPosition.Bottom;
            chartCategoryAxis.TickLabelPosition = TickLabelPosition.NextTo;
            chartCategoryAxis.CrossingAxisID = 2;
            chartCategoryAxis.Crosses = Crosses.Zero;
            chartCategoryAxis.Auto = true;
            chartCategoryAxis.LabelAlignment = LabelAlignment.Center;
            chartCategoryAxis.LabelOffset = 100;

            ValueAxis valueAxis = new ValueAxis();
            valueAxis.ID = 2;
            valueAxis.Scaling = new Scaling();
            valueAxis.Scaling.Orientation = Independentsoft.Office.Charts.Orientation.MinimumToMaximum;
            valueAxis.Position = AxisPosition.Left;
            valueAxis.MajorGridlines = new MajorGridlines();
            valueAxis.NumberFormat = new NumberFormat();
            valueAxis.NumberFormat.Format = "General";
            valueAxis.NumberFormat.IsSourceLinked = true;
            valueAxis.TickLabelPosition = TickLabelPosition.NextTo;
            valueAxis.CrossingAxisID = 1;
            valueAxis.Crosses = Crosses.Zero;
            valueAxis.CrossBetween = CrossBetween.Between;

            chartSpace.PlotArea.Axes.Add(chartCategoryAxis);
            chartSpace.PlotArea.Axes.Add(valueAxis);

            Legend legend = new Legend();
            legend.Position = LegendPosition.Right;
            legend.Layout = new Layout();

            chartSpace.Legend = legend;
            chartSpace.PlotVisibleOnly = true;

            Unit chartWidth = new Unit(640, UnitType.Pixel);
            Unit chartHeight = new Unit(480, UnitType.Pixel);

            Offset offset = new Offset(0, 0);
            Extents extents = new Extents(chartWidth, chartHeight);

            Inline inline = new Inline(chartSpace);
            inline.Size = new DrawingObjectSize(chartWidth, chartHeight);
            inline.ID = "1";
            inline.Name = "Chart 1";
            inline.Description = "Chart description";

            DrawingObject drawingObject = new DrawingObject(inline);

            ///////////////////////////////////////////////////

            Presentation presentation = new Presentation();

            Independentsoft.Office.Presentation.Drawing.GroupShape shapeTree = new Independentsoft.Office.Presentation.Drawing.GroupShape();
            shapeTree.ID = "1";
            shapeTree.Name = "ShapeTree";
            shapeTree.ShapeProperties.TransformGroup2D = new Independentsoft.Office.Drawing.TransformGroup2D();

            Independentsoft.Office.Presentation.Drawing.GraphicFrame graphicFrame = new Independentsoft.Office.Presentation.Drawing.GraphicFrame();
            graphicFrame.ID = "3";
            graphicFrame.Name = "Chart 1";
            graphicFrame.GraphicObject = chartSpace;
            graphicFrame.Locking = new GraphicFrameLocking();

            graphicFrame.Transform2D = new Independentsoft.Office.Presentation.Drawing.Transform2D();
            graphicFrame.Transform2D.Extents = new Extents(6197600, 4165600);
            graphicFrame.Transform2D.Offset = new Offset(1473200, 1346200);

            shapeTree.Elements.Add(graphicFrame);

            CommonSlideData commonSlideData = new CommonSlideData();
            commonSlideData.ShapeTree = shapeTree;

            SlideLayout layout1 = new SlideLayout();
            layout1.CommonSlideData = GetLayoutCommonSlideData();

            Slide slide1 = new Slide();
            slide1.CommonSlideData = commonSlideData;
            slide1.Layout = layout1;

            SlideMaster master1 = new SlideMaster();
            master1.CommonSlideData = GetLayoutCommonSlideData();
            master1.Layouts.Add(layout1);

            master1.TextStyles = new SlideMasterTextStyles();
            master1.TextStyles.TitleStyle = new SlideMasterTitleTextStyle();
            master1.TextStyles.TitleStyle.ListLevel1TextStyle.DefaultTextRunProperties.FontSize = 44;

            presentation.Slides.Add(slide1);
            presentation.SlideMasters.Add(master1);

            presentation.SlideSize = new SlideSize(9144000, 6858000, SlideSizeType.Screen4x3);
            presentation.NotesSlideSize = new NotesSlideSize(6858000, 9144000);

            presentation.Save("/Users/alfredyang/Desktop/output.pptx", true);
        }

        public override Object Exec(params Object[] parameters) {

            // TODO: Windows 才能使用，我去
            Console.WriteLine("PPTExp PhCommand: Create a new PPTX, with the UUID name!");
            var req = (phModel.PhRequest)parameters[0];
            var file_name = req.file_name;
            Console.WriteLine("get file with name");
            //var fct = phCommandFactory.PhCommandFactory.GetInstance();
            //var pptx = fct.GetHandledPPTX(file_name);

            //if (pptx == null) {
            //Workbook book = new Workbook();
            //book.LoadFromFile(@"~/Desktop/data.xlsx");
            //Image image = book.Worksheets[0].ToImage(1, 1, 5, 4);


            //Presentation ppt = new Presentation();
            //IImageData oleImage = ppt.Images.Append(image);
            //Rectangle rec = new Rectangle(60, 60, image.Width, image.Height);
            //using (MemoryStream ms = new MemoryStream()) {
            //    book.SaveToStream(ms);
            //    ms.Position = 0;
            //    Spire.Presentation.IOleObject oleObject = ppt.Slides[0].Shapes.AppendOleObject("excel", ms.ToArray(), rec);
            //    oleObject.SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImage;
            //    oleObject.ProgId = "Excel.Sheet.8";
            //}
            //ppt.SaveToFile("~/Desktop/InsertOle.ppt", Spire.Presentation.FileFormat.Pptx2007);
            //}

            //this.TestPPT();
            this.TestWithExcel();

            return null;
        }
    }
}
