using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CalSync.Infrastructure
{
    public class Setup
    {
        private MAPIFolder calendar;
        private string ruleName;
        private string folderName;
        private string emailSubject;

        public Setup(MAPIFolder calendar, String ruleName, String folderName, String emailSubject)
        {
            this.calendar = calendar;
            this.ruleName = ruleName;
            this.folderName = folderName;
            this.emailSubject = emailSubject;
        }

        public bool IsSetupComplete
        {
            get
            {
                return SyncFolder != null;
            }
        }

        public MAPIFolder SyncFolder
        {
            get
            {
                return calendar.Folders.Cast<MAPIFolder>().SingleOrDefault(f => f.Name == this.folderName);
            }
        }

        public void Install()
        {
            var folder = CreateSyncFolder();
            CreateSyncFolderRule(folder);
        }

        private MAPIFolder CreateSyncFolder()
        {
            return calendar.Folders.Add(this.folderName, OlDefaultFolders.olFolderInbox);
        }

        private void CreateSyncFolderRule(MAPIFolder targetFolder)
        {
            var allRules = Program.Outlook.Session.DefaultStore.GetRules();
            Rule rule = allRules.Cast<Rule>().SingleOrDefault(r => r.Name == this.ruleName);
            if (rule == null)
            {
                Rule textRule = allRules.Create(this.ruleName, OlRuleType.olRuleReceive);
                textRule.Conditions.Subject.Text = new[] { this.emailSubject };
                textRule.Conditions.Subject.Enabled = true;
                textRule.Actions.MoveToFolder.Folder = targetFolder;
                textRule.Actions.MoveToFolder.Enabled = true;
                textRule.Exceptions.Subject.Text = new[] { "RE: ", "FW: " };
                textRule.Exceptions.Subject.Enabled = true;

                allRules.Save(true);
            }
        }
    }
}
