
namespace PhPPTGen.phModel {
    public class PhRequest {
        public string id { get; set; }
        public string jobid { get; set; }
        public string command { get; set; }

        // Excel 
        public PhExcelPush push { get; set; }
        public PhExcel2PPT e2p { get; set; }
		public PhExcel2Chart e2c { get; set; }
        public PhExportPPT exp { get; set; }

        // Shape

        // Text
        public PhTextSetContent text { get; set; }
    }
}
