//
// AMGD
//

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Timers;

namespace MarketplacePublisher
{
    public partial class MarketplacePublisher : ServiceBase
    {
        public MarketplacePublisher()
        {
            InitializeComponent();

            try
            {
                if (!System.Diagnostics.EventLog.SourceExists("MarketplacePublisher-Source"))
                {
                    System.Diagnostics.EventLog.CreateEventSource(
                        "MarketplacePublisher-Source", "MarketplacePublisher-Log");
                }
                eventLog1.Source = "MarketplacePublisher-Source";
                eventLog1.Log = "MarketplacePublisher-Log";

                eventLog1.WriteEntry("About to create the timer...");
                TimerPublisher ltp = new TimerPublisher(eventLog1);
            }
            catch (Exception pe)
            {
                Console.WriteLine("Error: " + pe.ToString());
            }

        } // MarketplacePublisher

        protected override void OnStart(string[] args)
        {
            eventLog1.WriteEntry("Starting the service MarketplacePublisher - OnStart");
        } // OnStart

        protected override void OnStop()
        {
            eventLog1.WriteEntry("Stopping the service MarketplacePublisher - OnStop");
        } // OnStop

        protected override void OnContinue()
        {
            eventLog1.WriteEntry("Continuing the service MarketplacePublisher - OnContinue");
        } // OnContinue

    } // partial class MarketplacePublisher

    public class TimerPublisher
    {
        //private static System.Timers.Timer aTimer;
        private static EventLog _event_log = null;

        public TimerPublisher(EventLog pevent_log)
        {
            // Normally, the timer is declared at the class level,
            // so that it stays in scope as long as it is needed.
            // If the timer is declared in a long-running method,  
            // KeepAlive must be used to prevent the JIT compiler 
            // from allowing aggressive garbage collection to occur 
            // before the method ends. You can experiment with this
            // by commenting out the class-level declaration and 
            // uncommenting the declaration below; then uncomment
            // the GC.KeepAlive(aTimer) at the end of the method.
            System.Timers.Timer aTimer;

            _event_log = pevent_log;

            // Create a timer with a ten second interval.
            aTimer = new System.Timers.Timer(10000);

            // Hook up the Elapsed event for the timer.
            aTimer.Elapsed += new ElapsedEventHandler(OnTimedEvent);

            // Set the Interval to 5 minutes (300,000 milliseconds 1000 millisecs * 60 sec * 5 mins).
            // aTimer.Interval = 300000;
            aTimer.Interval = 5000; // 5 secs
            aTimer.Enabled = true;

            // If the timer is declared in a long-running method, use
            // KeepAlive to prevent garbage collection from occurring
            // before the method ends.
            GC.KeepAlive(aTimer);
        }

        // Specify what you want to happen when the Elapsed event is raised.
        private static void OnTimedEvent(object source, ElapsedEventArgs e)
        {
            _event_log.WriteEntry("The Elapsed event was raised at " + e.SignalTime);
        }
    } // TimerPublisher

} // namespace MarketplacePublisher
