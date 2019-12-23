﻿using Avro.Generic;
using Confluent.Kafka;
using Confluent.Kafka.SyncOverAsync;
using Confluent.SchemaRegistry.Serdes;
using Confluent.SchemaRegistry;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;


namespace AvroBlogExample {
	/// <summary>
	///     Complete source for the examples programs presented in the blog post:
	///     https://www.confluent.io/blog/decoupling-systems-with-apache-kafka-schema-registry-and-avro/
	/// </summary>
	class Program {
		async static Task ProduceGeneric(string bootstrapServers, string schemaRegistryUrl) {
			using (var schemaRegistry = new CachedSchemaRegistryClient(new SchemaRegistryConfig { SchemaRegistryUrl = schemaRegistryUrl }))
			using (var producer =
				new ProducerBuilder<Null, GenericRecord>(new ProducerConfig { BootstrapServers = bootstrapServers })
					.SetValueSerializer(new AvroSerializer<GenericRecord>(schemaRegistry))
					.Build()) {
				var logLevelSchema = (Avro.EnumSchema)Avro.Schema.Parse(
					File.ReadAllText("LogLevel.asvc"));

				var logMessageSchema = (Avro.RecordSchema)Avro.Schema
					.Parse(File.ReadAllText("LogMessage.V1.asvc")
						.Replace(
							"MessageTypes.LogLevel",
							File.ReadAllText("LogLevel.asvc")));

				var record = new GenericRecord(logMessageSchema);
				record.Add("IP", "127.0.0.1");
				record.Add("Message", "a test log message");
				record.Add("Severity", new GenericEnum(logLevelSchema, "Error"));
				await producer
					.ProduceAsync("log-messages", new Message<Null, GenericRecord> { Value = record })
					.ContinueWith(task => Console.WriteLine(
						task.IsFaulted
							? $"error producing message: {task.Exception.Message}"
							: $"produced to: {task.Result.TopicPartitionOffset}"));

				producer.Flush(TimeSpan.FromSeconds(30));
			}
		}

		//async static Task ProduceSpecific(string bootstrapServers, string schemaRegistryUrl) {
		//	using (var schemaRegistry = new CachedSchemaRegistryClient(new SchemaRegistryConfig { SchemaRegistryUrl = schemaRegistryUrl }))
		//	using (var producer =
		//		new ProducerBuilder<Null, MessageTypes.LogMessage>(new ProducerConfig { BootstrapServers = bootstrapServers })
		//			.SetValueSerializer(new AvroSerializer<MessageTypes.LogMessage>(schemaRegistry))
		//			.Build()) {
		//		await producer.ProduceAsync("log-messages",
		//			new Message<Null, MessageTypes.LogMessage> {
		//				Value = new MessageTypes.LogMessage {
		//					IP = "192.168.0.1",
		//					Message = "a test message 2",
		//					Severity = MessageTypes.LogLevel.Info,
		//					Tags = new Dictionary<string, string> { { "location", "CA" } }
		//				}
		//			});

		//		producer.Flush(TimeSpan.FromSeconds(30));
		//	}
		//}

		static void ConsumeSpecific(string bootstrapServers, string schemaRegistryUrl) {
			CancellationTokenSource cts = new CancellationTokenSource();
			Console.CancelKeyPress += (_, e) => {
				e.Cancel = true; // prevent the process from terminating.
				cts.Cancel();
			};

			var consumerConfig = new ConsumerConfig {
				GroupId = Guid.NewGuid().ToString(),
				BootstrapServers = bootstrapServers,
				AutoOffsetReset = AutoOffsetReset.Earliest
			};

			using (var schemaRegistry = new CachedSchemaRegistryClient(new SchemaRegistryConfig { SchemaRegistryUrl = schemaRegistryUrl }))
			using (var consumer =
				new ConsumerBuilder<Null, GenericRecord>(consumerConfig)
					.SetValueDeserializer(new AvroDeserializer<GenericRecord>(schemaRegistry).AsSyncOverAsync())
					.Build()) {
				consumer.Subscribe("log-messages");

				try {
					while (true) {
						try {
							var consumeResult = consumer.Consume(cts.Token);

							Console.WriteLine(
								consumeResult.Message.Timestamp.UtcDateTime.ToString("yyyy-MM-dd HH:mm:ss")
								+ $": [{consumeResult.Value}] {consumeResult.Value}");
						} catch (ConsumeException e) {
							Console.WriteLine($"an error occured: {e.Error.Reason}");
						}
					}
				} catch (OperationCanceledException) {
					// commit final offsets and leave the group.
					consumer.Close();
				}
			}
		}

		private static void PrintUsage()
			=> Console.WriteLine("Usage: .. <generic-produce|specific-produce|consume> <bootstrap-servers> <schema-registry-url>");

		//static async Task Main(string[] args) {
		//	if (args.Length != 3) {
		//		PrintUsage();
		//		return;
		//	}

		//	var mode = args[0];
		//	var bootstrapServers = args[1];
		//	var schemaRegistryUrl = args[2];

		//	switch (mode) {
		//		case "generic-produce":
		//		await ProduceGeneric(bootstrapServers, schemaRegistryUrl);
		//		break;
		//		case "specific-produce":
		//		//await ProduceSpecific(bootstrapServers, schemaRegistryUrl);
		//		break;
		//		case "consume":
		//		ConsumeSpecific(bootstrapServers, schemaRegistryUrl);
		//		break;
		//		default:
		//		PrintUsage();
		//		break;
		//	}
		//}
	}
}