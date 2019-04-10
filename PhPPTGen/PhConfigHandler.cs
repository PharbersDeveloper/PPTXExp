using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PhPPTGen {
	public class PhConfigHandler {
		public string path;
		private Dictionary<string, JToken> configMap { get; set;  }
		private static readonly PhConfigHandler instance = new PhConfigHandler();

		public static PhConfigHandler GetInstance() {
			return instance;
		}

		public Dictionary<string, JToken> GetConfigMap() {
			return configMap;
		}

		public void SetPath(string path) {
			this.path = path;
		}

		public void init() {
			using (StreamReader reader = File.OpenText(path + "PhConfig.json")) {
				//JToken.ReadFrom(new JsonTextReader(reader));
				configMap = (Dictionary<string, JToken>)new JsonSerializer().Deserialize(reader, typeof(Dictionary<string, JToken>));
			}
		}

		private PhConfigHandler() {
			
		}
	}
}
