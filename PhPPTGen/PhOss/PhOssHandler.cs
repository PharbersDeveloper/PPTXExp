using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using Aliyun.OSS;
using Aliyun.OSS.Common;

namespace PhPPTGen.PhOss {
	public class PhOssHandler {
		readonly string accessKeyId = "LTAIEoXgk4DOHDGi";
		readonly string accessKeySecret = "x75sK6191dPGiu9wBMtKE6YcBBh8EI";
		readonly string endpoint = "oss-cn-beijing.aliyuncs.com";
		private static PhOssHandler _instance = new PhOssHandler();

		public static PhOssHandler GetInstance() {
			return _instance;
		}

		public string DownloadFile(string downloadFilename) {
			var bucketName = "pharbers-max-bi";

			var localFilename = "<yourLocalFilename>";

			// 创建OSSClient实例。
			var client = new OssClient(endpoint, accessKeyId, accessKeySecret);
			try {
				// 下载文件。
				var result = client.GetObject(bucketName, downloadFilename);
				using (var requestStream = result.Content) {
					using (var fs = File.Open(downloadFilename, FileMode.OpenOrCreate)) {
						int length = 4 * 1024;
						var buf = new byte[length];
						do {
							length = requestStream.Read(buf, 0, length);
							fs.Write(buf, 0, length);
						} while (length != 0);
					}
				}
				Console.WriteLine("Get object succeeded");
			} catch (OssException ex) {
				Console.WriteLine("Failed with error code: {0}; Error info: {1}. \nRequestID:{2}\tHostID:{3}",
					ex.ErrorCode, ex.Message, ex.RequestId, ex.HostId);
			} catch (Exception ex) {
				Console.WriteLine("Failed with error info: {0}", ex.Message);
			}
			return localFilename;
		}

		public void UploadPPT(string PPTPath, string PPTName) {
			var bucketName = "pharbers-max-bi";

			// 创建OSSClient实例。
			var client = new OssClient(endpoint, accessKeyId, accessKeySecret);
			//try {
				// 上传文件。
				var result = client.PutObject(bucketName, PPTName, PPTPath);
				Console.WriteLine("Put object succeeded, ETag: {0} ", result.ETag);
			//} catch (Exception ex) {
			//	Console.WriteLine("Put object failed, {0}", ex.Message);
			//}
		}

		private PhOssHandler() {

		}
	}
}
