using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.IO;
using DDay.iCal;
using DDay.iCal.Serialization.iCalendar;

namespace CalSync.Synchronize
{
    class Sender
    {
        /// <summary>
        /// Send a message to an email address with an ical attachment of calendar events for a given date range
        /// </summary>
        /// <param name="calendar">An outlook calendar with events</param>
        /// <param name="remoteEmailAddress">The email address to send to</param>
        /// <param name="startDate">the beginning of the date range</param>
        /// <param name="endDate">the end of the date range</param>
        internal static void SendSynchronizationMessage(MAPIFolder calendar, DateTime startDate, DateTime endDate, String remoteEmailAddress, String emailSubject)
        {
            // send sync message to remote email address
            var calendarEvents = ReadOutlookCalendarEvents(calendar, startDate, endDate);
            var syncMessage = CreateSyncEmailMessage(calendarEvents);
            if (syncMessage != null)
            {
                syncMessage.To = remoteEmailAddress;
                syncMessage.Subject = emailSubject;
                syncMessage.Body = String.Format("Synchronization Message for CalSync. {0} Events sent at UTC Time: {1}", calendarEvents.Count(), DateTime.UtcNow);
                syncMessage.Send();
            }

            // delete the sent message so we don't clutter the local user's account.
            var deleteRule = String.Format("[Subject] = '{0}'", emailSubject);
            Program.Outlook.Session.GetDefaultFolder(OlDefaultFolders.olFolderSentMail).Items.Restrict(deleteRule).Cast<MailItem>().ToList().ForEach(deleted => deleted.Delete());
            Program.Outlook.Session.GetDefaultFolder(OlDefaultFolders.olFolderDeletedItems).Items.Restrict(deleteRule).Cast<MailItem>().ToList().ForEach(deleted => deleted.Delete());
        }

        /// <summary>
        /// Return events from an outlook calendar for a given date range
        /// </summary>
        /// <param name="calendar">the calendar to read</param>
        /// <param name="startDate">date range start</param>
        /// <param name="endDate">date range end</param>
        /// <returns>a list of events</returns>
        private static IEnumerable<IEvent> ReadOutlookCalendarEvents(MAPIFolder calendar, DateTime startDate, DateTime endDate)
        {
            // this implementation is a hack, but it's worth it not to get bogged down in the nasty
            // Outlook COM api (recurring appointments are particularly ugly). We export an iCal using 
            // Outlook's exporter (which handles the details for us), then read the iCal back using our
            // iCal processing library.

            // Set the properties for the export. We export the subject so we can filter on it later. 
            // we don't include the subject in the end result.
            CalendarSharing exporter = calendar.GetCalendarExporter();
            exporter.CalendarDetail = OlCalendarDetail.olFreeBusyAndSubject;
            exporter.IncludeAttachments = false;
            exporter.IncludePrivateDetails = false;
            exporter.RestrictToWorkingHours = false;
            exporter.StartDate = startDate;
            exporter.EndDate = endDate;

            // export to an ical file
            var tempFile = Path.GetTempFileName();
            exporter.SaveAsICal(tempFile);

            // ..and read it back in
            IICalendarCollection calendars = iCalendar.LoadFromFile(tempFile);
            var events = calendars.SelectMany(c => c.Events);
            File.Delete(tempFile);

            return events;
        }

        /// <summary>
        /// Create an email message with an ical attachment of calendar events
        /// </summary>
        /// <returns>a list of calendar events</returns>
        /// <returns>The synchronization email message</returns>
        private static MailItem CreateSyncEmailMessage(IEnumerable<IEvent> events)
        {
            var eventsToSync = events.Where(e => !(e.IsAllDay || e.Summary == "Busy"));

            MailItem mail = null;
            if (eventsToSync.Count() != 0)
            {
                // create an iCal file, blanking out each event's Subject to a generic 'Busy' message
                iCalendar target = new iCalendar();
                foreach (var appt in eventsToSync)
                {
                    appt.Summary = "Busy"; 
                    target.Events.Add(appt);
                }
                var serializer = new iCalendarSerializer();
                var tempFile = Path.GetTempFileName() + ".ics";
                serializer.Serialize(target, tempFile);

                // create the email message with the ical file as an attachment
                mail = Program.Outlook.CreateItem(OlItemType.olMailItem) as MailItem;
                mail.Attachments.Add(tempFile, OlAttachmentType.olByValue);
                File.Delete(tempFile);
            }

            return mail;
        }
    }
}
