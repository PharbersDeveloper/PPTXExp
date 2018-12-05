using System;
using Newtonsoft.Json;
using JsonApiSerializer;

namespace PhPPTGen.phCommon {
    public class PhCommon {
        public static string UUID() {
            return Guid.NewGuid().ToString();
        }

        public static T[] Content2ObjectLst<T> (phMsgDefine.PhMsgContent content) {
            var json = content.msg_content.Trim();
            T[] lst = JsonConvert.DeserializeObject<T[]>(json, new JsonApiSerializerSettings());
            System.Console.WriteLine(lst);
            return lst;
        }

        public static T Content2Object<T> (phMsgDefine.PhMsgContent content) {
            var json = content.msg_content.Replace("\0", "");
            var last = json.LastIndexOf("}");
            var start = json.IndexOf("{");
            var length = last - start + 1;
            json = json.Substring(start, length);
            T obj = JsonConvert.DeserializeObject<T>(json, new JsonApiSerializerSettings());
            System.Console.WriteLine(obj);
            return obj;
        }
    }
}
