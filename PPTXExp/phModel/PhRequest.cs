using System;
namespace PPTXExp.phModel {
    public class PhRequest {
        //public PhRequest() {
        //}
        public string id { get; set; }
        public string res { get; set; }
        public string command { get; set; }
        public string file_name { get; set; }
        public int slider_idx { get; set; }
        public int table_idx { get; set; }
    }
}
