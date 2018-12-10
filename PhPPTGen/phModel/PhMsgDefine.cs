
namespace PhPPTGen.phModel {
    public class PhMsgDefine {
        public static string Ph_Msg_GenPPT = "GenPPT";
        public static string Ph_Msg_ExcelPush = "ExcelPush";
        public static string Ph_Msg_Excel2PPT = "Excel2PPT";
        public static string Ph_Msg_ExportPPT = "ExportPPT";
        public static string Ph_Msg_PushText = "PushText";
        public static string Ph_Msg_Excel2Chart = "Excel2Chart";

        public static string PhCommand2Cls(string cmd) {
            if (cmd == Ph_Msg_GenPPT) {
                return "PhPPTGen.phCommand.PhGenPPTCommand";
            } else if (cmd == Ph_Msg_ExcelPush) {
                return "PhPPTGen.phCommand.phExcel.PhUpdateXlsCommand";
            } else if (cmd == Ph_Msg_Excel2PPT) {
                return "PhPPTGen.phCommand.phPpt.PhPPTImportXlsCommand";
            } else if (cmd == Ph_Msg_PushText) {
                return "PhPPTGen.phCommand.phText.PhTextContentCommand";
            } else if (cmd == Ph_Msg_ExportPPT) {
                return "PhPPTGen.phCommand.PhGenPPTCommand";
            } else if (cmd == Ph_Msg_Excel2Chart) {
                return "PhPPTGen.phCommand.phChart.PhChartContentCommand";
            } else {
                throw new System.Exception("Can not handler message");
            }
        }
    }
}
