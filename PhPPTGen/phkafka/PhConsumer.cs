using System;
using System.Threading;
using Confluent.Kafka;
using System.Threading.Tasks;
using PhPPTGen.phModel;
using Newtonsoft.Json;
using JsonApiSerializer;

namespace PhPPTGen.phkafka {
	class PhConsumer {
		//public static void Main(string[] args) {
		//	PushMsg("test1", "test");
		//	PushMsg("test2", "test");
		//	var conf = new ConsumerConfig {
		//		GroupId = "console-consumer-25919",
		//		BootstrapServers = "59.110.31.50:9093",
		//		// Note: The AutoOffsetReset property determines the start offset in the event
		//		// there are not yet any committed offsets for the consumer group for the
		//		// topic/partitions of interest. By default, offsets are committed
		//		// automatically, so in this example, consumption will only start from the
		//		// earliest message in the topic 'my-topic' the first time you run the program.
		//		AutoOffsetReset = AutoOffsetReset.Earliest
		//	};

		//	using (var c = new ConsumerBuilder<Ignore, string>(conf).Build()) {
		//		c.Subscribe("test");

		//		CancellationTokenSource cts = new CancellationTokenSource();
		//		Console.CancelKeyPress += (_, e) => {
		//			e.Cancel = true; // prevent the process from terminating.
		//			cts.Cancel();
		//		};

		//		try {
		//			while (true) {
		//				try {
		//					var cr = c.Consume(cts.Token);
		//					Console.WriteLine($"Consumed message '{cr.Value}' at: '{cr.TopicPartitionOffset}'.");
		//				} catch (ConsumeException e) {
		//					Console.WriteLine($"Error occured: {e.Error.Reason}");
		//				}
		//			}
		//		} catch (OperationCanceledException) {
		//			// Ensure the consumer leaves the group cleanly and final offsets are committed.
		//			c.Close();
		//		}
		//	}
		//}

		public static async Task PushMsg(string msg, string topic) {
			var config = new ProducerConfig {
				BootstrapServers = "59.110.31.50:9092" 
			};

			// If serializers are not specified, default serializers from
			// `Confluent.Kafka.Serializers` will be automatically used where
			// available. Note: by default strings are encoded as UTF8.
			using (var p = new ProducerBuilder<Null, string>(config).Build()) {
				try {
					var dr = await p.ProduceAsync(topic, new Message<Null, string> { Value = msg });
					Console.WriteLine($"Delivered '{dr.Value}' to '{dr.TopicPartitionOffset}'");
				} catch (ProduceException<Null, string> e) {
					Console.WriteLine($"Delivery failed: {e.Error.Reason}");
				}
			}
		}

		public static void PullMsg(string topic) {
			var conf = new ConsumerConfig {
				GroupId = "console-consumer-25919",
				BootstrapServers = "59.110.31.50:9093",
				// Note: The AutoOffsetReset property determines the start offset in the event
				// there are not yet any committed offsets for the consumer group for the
				// topic/partitions of interest. By default, offsets are committed
				// automatically, so in this example, consumption will only start from the
				// earliest message in the topic 'my-topic' the first time you run the program.
				AutoOffsetReset = AutoOffsetReset.Earliest
			};

			using (var c = new ConsumerBuilder<Ignore, string>(conf).Build()) {
				c.Subscribe(topic);

				CancellationTokenSource cts = new CancellationTokenSource();
				Console.CancelKeyPress += (_, e) => {
					e.Cancel = true; // prevent the process from terminating.
					cts.Cancel();
				};

				try {
					while (true) {
						try {
							var cr = c.Consume(cts.Token);
							Console.WriteLine($"Consumed message '{cr.Value}' at: '{cr.TopicPartitionOffset}'.");
							var msg = JsonConvert.DeserializeObject<PhRequest>(cr.Value, new JsonApiSerializerSettings());
							phCommon.PhRequestLst.GetInstance().PushMsg(msg);
						} catch (ConsumeException e) {
							Console.WriteLine($"Error occured: {e.Error.Reason}");
						}
					}
				} catch (OperationCanceledException) {
					// Ensure the consumer leaves the group cleanly and final offsets are committed.
					c.Close();
				}
			}
		}
	}
}
