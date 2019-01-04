using Newtonsoft.Json;
using PhPPTGen.phModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen.phCommand.phExcel.css {
    class Css {
        private static Dictionary<string, PhExcelCss> cssMap = new Dictionary<string, PhExcelCss>();
        public static void init() {
            if (cssMap.Count == 0) {
                string json = @"{
  'row_title_common' : {
                    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin'],
    'width' : '40',
    'horizontalAlignType' : 'Left'
  },
	'row_title_common1' : {
	'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
	'fontSize' : '9',
	'fontName' : 'Tahoma',
	'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin'],
	'width' : '30',
	'horizontalAlignType' : 'Left'
	},
  'row_title_chart' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'width' : '26.75',
    'horizontalAlignType' : 'Left'
  },
  'col_title_common' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FF0000',
    'cellBorders' : ['top#Thin', 'bottom#Thin']
  },
	'col_title_common1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FF0000',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
	'height' : '50'
  },
  'col_title_chart' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin']
  },
  'row_title_chart2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'cellColor' : '#0070C0',
    'verticalAlignType' : 'Top',
    'width' : '13'
  },
  'row_1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#D9D9D9',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
  },
  'row_2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FF0000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#E2EFDA',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
  },
  'row_3' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFF2CC',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
  },
  'row_4' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#D9E1F2',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
  },
  'row_5' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75',
    'cellBordersColor' : '#CDFFFF'
  },
  'row_6' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75',
    'cellBordersColor' : '#000000'
  },
  'row_7' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FFFF00',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
  },
  'row_8' : {
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '14'
  },
  'col_common' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '12'
  },
  'col_common1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#002060',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '15'
  },
  'col_common2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#002060',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '21'
  },
  'col_common3' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#002060',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '10'
  },
  'col_common4' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#002060',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '8.11'
  },
  'col_chart' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin', 'right#Thin'],
    'width' : '7'
  },
  'col_chart2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#EDEDED',
    'cellBorders' : ['left#Thin', 'right#Thin'],
    'width' : '11'
  },
  'timeline_1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'height' : '11.75'
  },
  'timeline_2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FF0000',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'height' : '11.75'
  },
  'timeline_3' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#0070C0',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'height' : '14'
  }
}
";
                cssMap = JsonConvert.DeserializeObject <Dictionary<string, PhExcelCss>>(json);
            }
        }
        public static PhExcelCss getCss(string cssName) {
            PhExcelCss css = new PhExcelCss();

            if (cssMap.TryGetValue(cssName, out css)) {
                return css;
            }
            return new PhExcelCss();
        }
    }
}
