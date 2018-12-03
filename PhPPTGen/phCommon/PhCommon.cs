using System;
using System.Collections.Generic;

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
            json = json.Substring(json.IndexOf("{"),  json.LastIndexOf("}"));
            T obj = JsonConvert.DeserializeObject<T>(json, new JsonApiSerializerSettings());
            System.Console.WriteLine(obj);
            return obj;
        }
    }
}
