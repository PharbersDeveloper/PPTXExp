
namespace PhPPTGen.phModel {
    public class PhRequest {
        public string id { get; set; }
        public PhProcessStep step { get; set; }
        public PhExcelPush push { get; set; }
        public PhExcel2PPT e2p { get; set; }
        public PhExportPPT exp { get; set; }
    }
}
