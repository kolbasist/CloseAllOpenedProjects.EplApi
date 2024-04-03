using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Base;
using Eplan.EplApi.EServices;
using Microsoft.Win32;
using System;
using System.Timers;
using CENTEC.EplAPI.Service;

namespace CENTEC.EplAddin.TestEplanAPI
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

            StartInitialTimer();

            _initialTimer.Elapsed += StartDailyTimer;
            _initialTimer.Elapsed += StopInitialTimer;

            _projectCloser = new ProjectCloser();
        }

        private void SessionSwitched(object sender, SessionSwitchEventArgs e)
        {
            Logger.SendMSGToEplanLog("[TimerManager.SesseionSwitched] Invoked");

            if (e.Reason == SessionSwitchReason.SessionLock)
            {
                Logger.SendMSGToEplanLog("[TimerManager.SesseionSwitched] Screen Locked");
                _lockTimer = new Timer();
                StartTimer(_lockTimer, TimeSpan.FromMinutes(1).TotalMilliseconds, false, "Lock Timer");
            }

            else if (e.Reason == SessionSwitchReason.SessionUnlock)
            {
                _lockTimer.Dispose();
                _lockTimer.Stop();
            }
        }

        private void TimerElapsed(object sender, ElapsedEventArgs e)
        {
            _projectCloser.CloseAll();
            Logger.SendMSGToEplanLog($"[ClosingManager.TimerElapsed] Event was invoked at {DateTime.Now}");
        }

        private void StartDailyTimer(object sender, ElapsedEventArgs e)
        {
            _dailyTimer = new Timer();
            StartTimer(_dailyTimer, TimeSpan.FromHours(24).TotalMilliseconds, true, "Daily Timer");
        }

        private void StartTimer(Timer timer, double interval, bool isAutoresetRequired, string timerName)
        {
            timer.Close();
            timer.AutoReset = isAutoresetRequired;
            timer.Interval = interval;
            timer.Start();
            timer.Elapsed += TimerElapsed;
            Logger.SendMSGToEplanLog($"[ClosingManager.StartTimer] {timerName} has succesfully will go of at {DateTime.Now.AddMilliseconds(interval)}.");
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

                double interval = (nextRunTime - now).TotalMilliseconds;
                StartTimer(_initialTimer, interval, false, "initial Timer");
            }
            else
            {
                StartTimer(_initialTimer, TimeSpan.FromMinutes(2).TotalMilliseconds, false, "Initial Timer");
            }
        }

        private void StopInitialTimer(object sender, ElapsedEventArgs e)
        {
            _initialTimer.Stop();
            _initialTimer.Elapsed -= StopInitialTimer;
            Logger.SendMSGToEplanLog("Initial Timer disposed.");
        }
    }
}
