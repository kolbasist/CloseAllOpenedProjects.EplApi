using Microsoft.Win32;
using System;
using System.Windows.Forms;
using MyEplanAPI.Service;

namespace Warden
{
    internal class TimersManager
    {
        private Timer _dailyTimer;
        private Timer _initialTimer;
        private Timer _lockTimer;
        private ProjectCloser _projectCloser;

        public TimersManager()
        {
            SystemEvents.SessionSwitch += SessionSwitched;

            _projectCloser = new ProjectCloser();

            StartInitialTimer(false);
        }

        private void SessionSwitched(object sender, SessionSwitchEventArgs e)
        {
            Logger.SendMSGToEplanLog("[TimerManager.SessionSwitched] Invoked");

            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                Logger.SendMSGToEplanLog("[TimerManager.SessionSwitched] Screen Locked");
                _lockTimer = new Timer();
                StartTimer(_lockTimer, (int)TimeSpan.FromMinutes(15).TotalMilliseconds, false, "Lock Timer");
            }
            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                if (_lockTimer != null)
                {
                    _lockTimer.Stop();
                    _lockTimer.Dispose();
                    _lockTimer = null;
                }
            }
        }

        private void TimerElapsed(object sender, EventArgs e)
        {
            _projectCloser.CloseAll();
            Logger.SendMSGToEplanLog($"[ClosingManager.TimerElapsed] Event was invoked at {DateTime.Now}");
        }

        private void StartDailyTimer(object sender, EventArgs e)
        {
            _dailyTimer = new Timer();
            StartTimer(_dailyTimer, (int)TimeSpan.FromHours(24).TotalMilliseconds, true, "Daily Timer");
        }

        private void StartTimer(Timer timer, int interval, bool isAutoresetRequired, string timerName)
        {
            timer.Interval = interval;
            timer.Tick += TimerElapsed;
            timer.Start();
            Logger.SendMSGToEplanLog($"[ClosingManager.StartTimer] {timerName} will go off at {DateTime.Now.AddMilliseconds(interval)}.");
        }

        private void StartInitialTimer(bool isDebugging = false)
        {
            _initialTimer = new Timer();

            if (isDebugging == false)
            {
                DateTime now = DateTime.Now;
                DateTime nextRunTime = new DateTime(now.Year, now.Month, now.Day, 2, 0, 0);
                Logger.SendMSGToEplanLog($"[ClosingManager.StartDailyTimer] Daily closing timer was set as {nextRunTime} at {DateTime.Now}");

                if (nextRunTime < now)
                {
                    nextRunTime = nextRunTime.AddDays(1);
                }

                int interval = (int)(nextRunTime - now).TotalMilliseconds;
                StartTimer(_initialTimer, interval, false, "Initial Timer");
            }
            else
            {
                StartTimer(_initialTimer, (int)TimeSpan.FromMinutes(2).TotalMilliseconds, false, "Initial Timer");
            }

            _initialTimer.Tick += StartDailyTimer;
            _initialTimer.Tick += StopInitialTimer;
        }

        private void StopInitialTimer(object sender, EventArgs e)
        {
            if (_initialTimer != null)
            {
                _initialTimer.Stop();
                _initialTimer.Tick -= StopInitialTimer;
                Logger.SendMSGToEplanLog("Initial Timer disposed.");
                _initialTimer.Dispose();
                _initialTimer = null;
            }
        }
    }
}
