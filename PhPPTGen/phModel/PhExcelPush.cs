
namespace PhPPTGen.phModel {
    public class PhExcelPush {
        public string id { get; set; }
        public string name { get; set; }
        public string cell { get; set; }
        public string css { get; set; }     // 格式信息
        public string cate { get; set; }  // One of the Cell type, (Number Or String)
        public string value { get; set; } // 全部通过string转化为应有的cate
    }
}
