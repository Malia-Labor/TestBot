using Discord;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BasicBot
{
    public class ScheduleHandler
    {
        private System.Threading.Timer _timer;
        private DateTime _target;
        private int _waitTime;
        private double _targetHours = 16;
        private double _targetMinutes = 30;

        public void LoadScheduler()
        {
            // Target time to perform task
            _target = DateTime.Today.AddHours(_targetHours).AddMinutes(_targetMinutes);
            // If target time is already passed for today, set target for tomorrow
            if(_target < DateTime.Now)
                _target = DateTime.Today.AddHours(_targetHours).AddMinutes(_targetMinutes).AddDays(1);
            // Get milliseconds between target time and now
            _waitTime = (int)((_target - DateTime.Now).TotalMilliseconds);
            // Create timer
            _timer = new System.Threading.Timer(DoSomething, null, _waitTime, _waitTime);
        }

        void DoSomething(object state)
        {
            // change target date to next day
            _target = DateTime.Today.AddHours(_targetHours).AddMinutes(_targetMinutes).AddDays(1);
            // Change wait time for timer
            _waitTime = (int)((_target - DateTime.Now).TotalMilliseconds);
            // Do something on a timer...
            Console.WriteLine(_target);
        }
    }
}
