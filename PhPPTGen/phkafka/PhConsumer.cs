using System;
using System.Threading;
using Confluent.Kafka;
using System.Threading.Tasks;
using PhPPTGen.phModel;
using Newtonsoft.Json;
using JsonApiSerializer;
using System.Collections.Generic;
using Confluent.SchemaRegistry;
using Avro.Generic;
using Confluent.SchemaRegistry.Serdes;
using Confluent.Kafka.SyncOverAsync;
using System.IO;

namespace PhPPTGen.phkafka {
	public class PhConsumer {
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

		readonly Dictionary<string, string> config = new Dictionary<string, string>() {
			{ "group.id", "console-consumer-25919" },
			{ "bootstrap.servers", "59.110.31.50:9093" },
			{ "schema.registry.url", "59.110.31.50:8081" },
			{ "auto.offset.reset", "beginning" }
		};

		private static readonly PhConsumer Instance = new PhConsumer();

		private PhConsumer() { }

		public static PhConsumer GetInstance() {
			return Instance;
		}

		public async Task PushMsg(string msg, string topic, string schema) {

			using (var schemaRegistry = new CachedSchemaRegistryClient(new SchemaRegistryConfig { SchemaRegistryUrl = config["schema.registry.url"] }))
			using (var producer = new ProducerBuilder<Null, GenericRecord>(new ProducerConfig { BootstrapServers = config["bootstrap.servers"] })
					.SetValueSerializer(new AvroSerializer<GenericRecord>(schemaRegistry))
					.Build()) {
				var resultSchema = (Avro.RecordSchema)Avro.Schema
					.Parse(File.ReadAllText($"{PhConfigHandler.GetInstance().path}msg/{schema}.asvc"));
				var record = new GenericRecord(resultSchema);
				record.Add("data", msg);
				record.Add("id", Guid.NewGuid().ToString());

				try {
					await producer
					.ProduceAsync(topic, new Message<Null, GenericRecord> { Value = record })
					.ContinueWith(task => Console.WriteLine(
						task.IsFaulted
							? $"error producing message: {task.Exception.Message}"
							: $"produced to: {task.Result.TopicPartitionOffset}"));

					producer.Flush(TimeSpan.FromSeconds(30));
					//Console.WriteLine($"Delivered '{dr.Value}' to '{dr.TopicPartitionOffset}'");
				} catch (ProduceException<Null, string> e) {
					Console.WriteLine($"Delivery failed: {e.Error.Reason}");
				}
			}
		}

		public void PullMsg(string topic) {

			var conf = new ConsumerConfig {
				GroupId = "console-consumer-25919",
				BootstrapServers = "59.110.31.50:9093",
				AutoOffsetReset = AutoOffsetReset.Earliest
			};
			using (var schemaRegistry = new CachedSchemaRegistryClient(new SchemaRegistryConfig { SchemaRegistryUrl = "59.110.31.50:8081" }))
			using (var c = new ConsumerBuilder<Null, GenericRecord>(conf)
				.SetValueDeserializer(new AvroDeserializer<GenericRecord>(schemaRegistry).AsSyncOverAsync())
				.Build()) {
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
							var msg = JsonConvert.DeserializeObject<PhRequest>(cr.Value["data"].ToString(), new JsonApiSerializerSettings());
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
