using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen {
	class PhConfigHandler {
		public readonly string path;
		public readonly Dictionary<string, JToken> configMap;
		private static readonly PhConfigHandler instance = new PhConfigHandler();

		public static PhConfigHandler GetInstance() {
			return instance;
		}

		private PhConfigHandler() {
			path = @"..\..\resources\";
			using (StreamReader reader = File.OpenText(path + "PhConfig.json")) {
				//JToken.ReadFrom(new JsonTextReader(reader));
				configMap = (Dictionary<string, JToken>)new JsonSerializer().Deserialize(reader, typeof(Dictionary<string, JToken>));
			}
		}
	}
}
