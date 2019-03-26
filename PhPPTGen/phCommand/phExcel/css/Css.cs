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
		private static Dictionary<string, PhExcelCssForOpenxml> openxmlCssMap = new Dictionary<string, PhExcelCssForOpenxml>();
		public static void init() {
            if (cssMap.Count == 0) {
                string json = @"
  {
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
    'row_title_chart1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'width' : '18',
    'horizontalAlignType' : 'Left'
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
  'row_title_chart3' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'width' : '16',
    'horizontalAlignType' : 'Left'
  },
  'row_title_chart4' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'width' : '15.23',
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
    'height' : '25'
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
    'col_title_chart1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '8',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin']
  },
  'col_title_common2' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FF0000',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '34.2'
  },
  'col_title_rank' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FF0000',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'left#Thin']
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
  'row_9' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#D9D9D9',
    'cellBorders' : ['top#Thin', 'bottom#Thin'],
    'height' : '11.75'
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
  'col_common5' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '9'
  },
  'col_common6' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin'],
    'width' : '7.5'
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
    'col_chart1' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '6',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : [],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['left#Thin', 'right#Thin'],
    'width' : '5'
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
  },
    'timeline_4' : {
    'factory' : 'PhPPTGen.phCommand.phExcel.PhSetXlsCssBaseCommand',
    'fontSize' : '8',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'fontStyle' : ['bold'],
    'cellColor' : '#FFFFFF',
    'cellBorders' : ['top#Thin', 'bottom#Thin', 'right#Thin', 'left#Thin'],
    'height' : '11.75'
  }
}
";

				string xmlJson = @"{
  'timeline_1' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
	'timeline_2' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FF0000',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'timeline_3' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#0070C0',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '14'
  },
    'timeline_4' : {
    'fontSize' : '8',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'col_title_common' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FF0000',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0'
  },
  'col_title_common1' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FF0000',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '25'
  },
  'col_title_chart' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0'
  },
  'col_title_chart1' : {
    'fontSize' : '8',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0'
  },
  'col_title_common2' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FF0000',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0',
    'height' : '34.2'
  },
  'col_title_rank' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FF0000',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0'
  },
  'col_common' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'width' : '12'
  },
  'col_common1' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
	  'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'cellColor' : '#FFFFFF',
    'width' : '15'
  },
  'col_common2' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'width' : '21'
  },
  'col_common3' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
	  'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'cellColor' : '#FFFFFF',
    'width' : '10'
  },
  'col_common4' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'width' : '8.11'
  },
  'col_common5' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'width' : '9'
  },
  'col_common6' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'None000000',
    'width' : '7.5'
  },
  'col_chart' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '7'
  },
   'col_chart1' : {
    'fontSize' : '6',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'cellColor' : '#FFFFFF',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '5'
  },
  'col_chart2' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'cellColor' : '#EDEDED',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '11'
  },
  'row_title_common' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'width' : '40',
	  'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0',
    'horizontalAlignType' : 'Left'
  },
  'row_title_common1' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'None000000',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '30',
    'horizontalAlignType' : 'Left'
  },
  'row_title_chart' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '26.75',
    'horizontalAlignType' : 'Left'
  },
   'row_title_chart1' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '18',
    'horizontalAlignType' : 'Left'
  },
  'row_title_chart2' : {
    'fontSize' : '9',
    'fontColor' : '#FFFFFF',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'cellColor' : '#0070C0',
    'verticalAlignType' : 'Top',
    'width' : '13'
  },
  'row_title_chart3' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '16',
    'horizontalAlignType' : 'Left'
  },
  'row_title_chart4' : {
    'fontSize' : '9',
    'fontName' : 'Tahoma',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'leftBorder' : 'ThinF0F0F0',
    'rightBorder' : 'ThinF0F0F0',
    'width' : '15.23',
    'horizontalAlignType' : 'Left'
  },
  'row_1' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#D9D9D9',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_2' : {
    'fontSize' : '9',
    'fontColor' : '#FF0000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#E2EFDA',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_3' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#FFF2CC',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_4' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#D9E1F2',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_5' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_6' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'false',
    'cellColor' : '#FFFFFF',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_7' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#FFFF00',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  },
  'row_8' : {
    'border' : 'ThinF0F0F0ThinF0F0F0ThinF0F0F0ThinF0F0F0',
    'height' : '14'
  },
  'row_9' : {
    'fontSize' : '9',
    'fontColor' : '#000000',
    'fontName' : 'Tahoma',
    'bold': 'true',
    'cellColor' : '#D9D9D9',
    'topBorder' : 'ThinF0F0F0',
    'bottomBorder' : 'ThinF0F0F0',
    'height' : '11.75'
  }
}";
				openxmlCssMap = JsonConvert.DeserializeObject<Dictionary<string, PhExcelCssForOpenxml>>(xmlJson);
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

		public static PhExcelCssForOpenxml getOpenxmlCss(string cssName) {
			PhExcelCssForOpenxml css = new PhExcelCssForOpenxml();

			if (openxmlCssMap.TryGetValue(cssName, out css)) {
				return css;
			}
			return new PhExcelCssForOpenxml();
		}
	}
}
