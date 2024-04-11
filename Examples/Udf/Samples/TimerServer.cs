using ExcelMvc.Rtd;
using System;
using System.Collections.Concurrent;
//using System.Threading;

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
            
        private System.Timers.Timer Timer { get; }
        public TimerServer()
        {
            Timer = new System.Timers.Timer(5000);
            Timer.Elapsed += Timer_Elapsed;
            Timer.Start();
        }

        public object Connect(int topicId, string[] args)
        {
            ExcelMvc.Diagnostics.Messages.Instance.AddInfoLine($"{topicId} connected");
            return Topics[topicId] = new Topic { args = args, value = DateTime.Now };
        }

        public void Disconnect(int topicId)
        {
            ExcelMvc.Diagnostics.Messages.Instance.AddInfoLine($"{topicId} disconnected");
            Topics.TryRemove(topicId, out var _);
        }

        public object[,] GetTopicValues()
        {
            var snapshot = Topics.ToArray();
            var values = new object[2, snapshot.Length];
            for (int i = 0; i < snapshot.Length; i++)
            {
                values[0, i] = snapshot[i].Key;
                values[1, i] = $"{snapshot[i].Value.value:O}{string.Join(",", snapshot[i].Value.args)}";
            }
            return values;
        }

        public int Heartbeat()
        {
            return 1;
        }

        public int Start()
        {
            ExcelMvc.Diagnostics.Messages.Instance.AddInfoLine("Started");
            return 1;
        }

        public void Stop()
        {
            ExcelMvc.Diagnostics.Messages.Instance.AddInfoLine("Stopped");
            Timer.Stop();
            Topics.Clear();
        }

        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            ExcelMvc.Diagnostics.Messages.Instance.AddInfoLine("Time Ticked");
            var now = DateTime.Now;
            foreach (var pair in Topics)
                pair.Value.value = DateTime.Now;
            Updated?.Invoke(this, EventArgs.Empty);
        }
    }
}
