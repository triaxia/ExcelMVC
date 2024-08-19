using System.Text;
using RabbitMQ.Client;

namespace RabbitMQSender
{
    public static class Program
    {
        public async static Task Main(string[] args)
        {
            var factory = new ConnectionFactory { HostName = "localhost" };
            using var connection = await factory.CreateConnectionAsync();
            using var channel = await connection.CreateChannelAsync();

            await channel.QueueDeclareAsync(queue: "hello",
                     durable: false,
                     exclusive: false,
                     autoDelete: false,
                     arguments: null);

            while (true)
            {
                var message = $"{DateTime.Now:O}";
                var body = Encoding.UTF8.GetBytes(message);
                await channel.BasicPublishAsync(exchange: string.Empty,
                                     routingKey: "hello",
                                     //basicProperties: null,
                                     body: body);
                Console.WriteLine($"Sent {message}");
                await Task.Delay(1000);
            }
        }
    }
}
