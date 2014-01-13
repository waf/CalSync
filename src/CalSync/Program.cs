﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Configuration;
using WPF = System.Windows;
using Microsoft.Office.Interop.Outlook;
using CalSync.Synchronize;
using CalSync.Infrastructure;

namespace CalSync
{
    class Program
    {
        public static readonly Application Outlook = new Application();
        private const string EmailSubject = "CalSync Synchronization Message";
        private const string OutlookRuleName = "CalSync Folder Rule";
        private const string OutlookFolderName = "CalSync Messages";

        static void Main(string[] args)
        {
            // parse config from App.config
            Config cfg = Config.Read();
            if(!cfg.ValidConfiguration)
            {
                WPF.MessageBox.Show("Error reading configuration. Please ensure CalSync.config contains valid values.");
                return;
            }

            // create required outlook components, if necessary
            var calendar = Outlook.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            Setup setup = new Setup(calendar, OutlookRuleName, OutlookFolderName, EmailSubject);
            if(!setup.IsSetupComplete)
            {
                setup.Install();
                WPF.MessageBox.Show("CalSync setup completed successfully.");
                return;
            }

            // send sync message to remote inbox
            var rangeStart = DateTime.Now.Date;
            var rangeEnd = rangeStart.AddDays(cfg.SyncRangeDays);
            Sender.SendSynchronizationMessage(calendar, rangeStart, rangeEnd, cfg.TargetEmailAddress, EmailSubject);

            // read events from the local sync folder, and add them to the calendar
            Receiver.ProcessReceivedMessages(setup.SyncFolder, calendar);
        }
    }
}