using Function.Definitions;
using System;
using System.Collections.Concurrent;
using System.Threading;

namespace ExcelMvc.Rtd
{
    public class RtdServerImplTest : IRtdServerImpl
    {
        public event EventHandler<EventArgs> Updated;
        public readonly ConcurrentDictionary<int, DateTime> Topics = new ConcurrentDictionary<int, DateTime>();
        private Timer Timer { get; }
        public RtdServerImplTest()
        {
            Timer = new Timer(OnTimer, null, 5000, 5000);
        }

        public object Connect(int topicId, string[] args)
        {
            return Topics[topicId] = DateTime.Now;
        }

        public void Disconnect(int topicId)
        {
            Topics.TryRemove(topicId, out var _);
        }

        public object[,] GetTopicValues()
        {
            var snapshot = Topics.ToArray();
            var values = new object[2, snapshot.Length];
            for (int i = 0; i < snapshot.Length; i++)
            {
                values[0, i] = snapshot[i].Key;
                values[1, i] = snapshot[i].Value;
            }
            return values;
        }

        public int Heartbeat()
        {
            return 1;
        }

        public int Start()
        {
            return 1;
        }

        public void Terminate()
        {
            Timer.Change(0, Timeout.Infinite);
            Topics.Clear();
        }

        private void OnTimer(object state)
        {
            var now = DateTime.Now;
            foreach (var key in Topics.Keys)
                Topics[key] = now;
            Updated?.Invoke(this, EventArgs.Empty);
        }
    }
}
