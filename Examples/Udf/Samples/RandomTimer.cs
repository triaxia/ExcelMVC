﻿using Addin.Interfaces;
using System;
using System.Collections.Concurrent;
using System.Threading;

namespace Samples
{
    public class RandomTimer : IRtdServerImpl
    {
        public class Topic
        {
            public string[] args;
            public double value;
        }
        public event EventHandler<EventArgs> Updated;
        public readonly ConcurrentDictionary<int, Topic> Topics
            = new ConcurrentDictionary<int, Topic>();

        private Timer Timer { get; }
        public RandomTimer()
        {
            Timer = new Timer(OnTimer, null, 5000, 5000);
        }

        public object Connect(int topicId, string[] args)
        {
            return Topics[topicId] = new Topic { args = args, value = new Random().NextDouble() };
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
                values[1, i] = $"{snapshot[i].Value.value}{string.Join(",", snapshot[i].Value.args)}";
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

        public void Stop()
        {
            Timer.Change(0, Timeout.Infinite);
            Topics.Clear();
        }

        private void OnTimer(object state)
        {
            var now = DateTime.Now;
            foreach (var pair in Topics)
                pair.Value.value = new Random().NextDouble();
            Updated?.Invoke(this, EventArgs.Empty);
        }
    }
}