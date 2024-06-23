using System;
using System.Collections.Concurrent;
using System.Threading;
using Function.Interfaces;

namespace FunctionLibrary
{
    public class TimerServer : IRtdServerImpl
    {
        public event EventHandler<RtdServerUpdatedEventArgs> Updated;
        public readonly ConcurrentDictionary<int, RtdTopic> Topics
            = new ConcurrentDictionary<int, RtdTopic>();

        private Timer Timer { get; set; }

        public int Start()
        {
            Timer = new Timer(TimerElapsed, null, 1000, 1000);
            FunctionHost.Instance.RaisePosted(this, new MessageEventArgs("Started"));
            return 1;
        }

        public void Stop()
        {
            FunctionHost.Instance.RaisePosted(this, new MessageEventArgs("Stopped"));
            Timer.Dispose();
            Topics.Clear();
        }

        public int Heartbeat()
        {
            return 1;
        }

        public object Connect(int topicId, string[] args)
        {
            FunctionHost.Instance.RaisePosted(this, new MessageEventArgs($"{topicId} connected"));
            Topics[topicId] = new RtdTopic(args, DateTime.Now);
            return Format(Topics[topicId]);
        }

        public void Disconnect(int topicId)
        {
            FunctionHost.Instance.RaisePosted(this, new MessageEventArgs($"{topicId} disconnected"));
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
            FunctionHost.Instance.RaisePosted(this, new MessageEventArgs("Ticked"));
            var now = DateTime.Now;
            foreach (var pair in Topics.ToArray())
                pair.Value.Value = now;
            Updated?.Invoke(this, new RtdServerUpdatedEventArgs(this, Topics.Values));
        }
        private static string Format(RtdTopic topic) => $"{topic}";
    }
}
