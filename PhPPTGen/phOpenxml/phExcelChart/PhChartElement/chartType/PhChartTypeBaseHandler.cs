using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace PhPPTGen.phOpenxml.phExcelChart.PhChartElement {
	abstract class PhChartTypeBaseHandler: PhBaseElementHandler {

		protected C.SeriesText CreateSeriesText(string value, string formulaValue) {
			C.SeriesText seriesText = new C.SeriesText();
			seriesText.Append(CreateStringReference(new List<string> { value }, formulaValue));
			return seriesText;
		}

		protected C.StringReference CreateStringReference(List<string> values, string formulaValue) {
			C.StringReference stringReference = new C.StringReference();
			C.Formula formula = new C.Formula {
				//excel上的位置	
				Text = formulaValue // "Sheet1!$B$1:$D$1";
			};

			C.StringCache stringCache = new C.StringCache();
			stringCache.Append(new C.PointCount() { Val = new UInt32Value((uint)values.Count) });
			foreach (string value in values) {
				C.StringPoint stringPoint = new C.StringPoint() { Index = (UInt32Value)(uint)values.ToList().IndexOf(value) };
				C.NumericValue numericValue = new C.NumericValue {
					Text = value
				};
				stringPoint.Append(numericValue);
				stringCache.Append(stringPoint);
			}

			stringReference.Append(formula);
			stringReference.Append(stringCache);

			return stringReference;
		}

		protected C.NumberReference CreateNumberReference(List<string> values, string formulaValue, string numFormat) {

			C.NumberReference numberReference = new C.NumberReference();
			C.Formula formula = new C.Formula {
				Text = formulaValue // "Sheet1!$B$2:$D$2"
			};

			C.NumberingCache numberingCache = new C.NumberingCache();

			numberingCache.Append(new C.FormatCode { Text = numFormat });
			numberingCache.Append(new C.PointCount() { Val = (UInt32Value)(uint)values.Count });
			uint index = 0;
			foreach (string value in values) {
				C.NumericPoint numericPoint = new C.NumericPoint() { Index = index };
				C.NumericValue numericValue = new C.NumericValue { Text = value };
				numericPoint.Append(numericValue);
				numberingCache.Append(numericPoint);
				index ++;
			}
			numberReference.Append(formula);
			numberReference.Append(numberingCache);

			return numberReference;
		}
	}
}
