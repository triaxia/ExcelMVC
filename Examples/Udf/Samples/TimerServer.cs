using ExcelMvc.Rtd;
using System;
using System.Collections.Concurrent;
using System.Threading;
using ExcelMvc.Diagnostics;

namespace Samples
{
    public class Topic
    {
        public string[] args;
        public DateTime value;
    }
    public class TimerServer : IRtdServerImpl
    {
        public event EventHandler<EventArgs> Updated;
        public readonly ConcurrentDictionary<int, Topic> Topics
            = new ConcurrentDictionary<int, Topic>();

        private Timer Timer { get; set; }

        public int Start()
        {
            Timer = new Timer(TimerElapsed, null, 1000, 1000);
            Messages.Instance.AddInfoLine("Started");
            return 1;
        }

        public void Stop()
        {
            Messages.Instance.AddInfoLine("Stopped");
            Timer.Dispose();
            Topics.Clear();
        }

        public int Heartbeat()
        {
            return 1;
        }

        public object Connect(int topicId, string[] args)
        {
            Messages.Instance.AddInfoLine($"{topicId} connected");
            Topics[topicId] = new Topic { args = args, value = DateTime.Now };
            return Format(Topics[topicId]);
        }

        public void Disconnect(int topicId)
        {
            Messages.Instance.AddInfoLine($"{topicId} disconnected");
            Topics.TryRemove(topicId, out var _);
        }

        public object[,] GetTopicValues()
        {
            var snapshot = Topics.ToArray();
            var values = new object[2, snapshot.Length];
            for (int i = 0; i < snapshot.Length; i++)
            {
                values[0, i] = snapshot[i].Key;
                values[1, i] = Format(snapshot[i].Value);
            }
            return values;
        }

        private void TimerElapsed(object _)
        {
            Messages.Instance.AddInfoLine("Time Ticked");
            var now = DateTime.Now;
            foreach (var pair in Topics.ToArray())
                pair.Value.value = DateTime.Now;
            Updated?.Invoke(this, EventArgs.Empty);
        }

        private static string Format(Topic topic)
            => $"{topic.value:O}{string.Join(",", topic.args)}";
    }
}
