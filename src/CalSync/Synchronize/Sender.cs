using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
            string outputFile = Path.GetTempPath() + "calsync.ics";
            // send sync message to remote email address
            var calendarEvents = ReadOutlookCalendarEvents(calendar, startDate, endDate);
            var eventsToSync = FilterCalendarEvents(calendarEvents);
            var syncMessage = CreateSynchronizationMessage(eventsToSync, outputFile);
            if (syncMessage != null)
            {
                syncMessage.To = remoteEmailAddress;
                syncMessage.Subject = emailSubject;
                syncMessage.Body = String.Format("Synchronization Message for CalSync. {0} Events sent at UTC Time: {1}", calendarEvents.Count(), DateTime.UtcNow);
                syncMessage.Send();
            }

            // delete the sent/deleted messages so we don't clutter the local user's account.
            var deleteRule = String.Format("[Subject] = '{0}'", emailSubject);
            DeleteMessages(OlDefaultFolders.olFolderSentMail, deleteRule);
            DeleteMessages(OlDefaultFolders.olFolderDeletedItems, deleteRule);
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
            var events = ReadCalendarExport(tempFile);
            File.Delete(tempFile);

            return events;
        }

        /// <summary>
        /// We don't want to sync every single calendar event. Given a list of calendar events, return the ones we care about.
        /// </summary>
        /// <param name="calendarEvents">list of all calendar events</param>
        /// <returns>filtered list of events</returns>
        private static IEnumerable<IEvent> FilterCalendarEvents(IEnumerable<IEvent> calendarEvents)
        {
            return calendarEvents.Where(e =>
                !(e.IsAllDay ||
                  e.Summary == "Busy" ||
                  e.Properties["X-MICROSOFT-CDO-BUSYSTATUS"].Value.ToString() == "FREE"));
        }

        /// <summary>
        /// Create an email message with an ical attachment of calendar events
        /// </summary>
        /// <param name="outputFile">the filename to read</param>
        /// <returns>The synchronization email message</returns>
        private static MailItem CreateSynchronizationMessage(IEnumerable<IEvent> events, String outputFile)
        {
            MailItem mail = null;
            if (events.Count() != 0)
            {
                // create an iCal file, blanking out each event's Subject to a generic 'Busy' message
                iCalendar target = new iCalendar();
                foreach (var e in events)
                {
                    e.Summary = "Busy"; 
                    target.Events.Add(e);
                }

                // get the previous run's output. only bother creating a synchronization message if we have new information
                var oldCalendar = ReadCalendarExport(outputFile);
                if (!oldCalendar.SequenceEqual(target.Events))
                {
                    var serializer = new iCalendarSerializer();
                    serializer.Serialize(target, outputFile);

                    // create the email message with the ical file as an attachment
                    mail = Program.Outlook.CreateItem(OlItemType.olMailItem) as MailItem;
                    mail.Attachments.Add(outputFile, OlAttachmentType.olByValue);
                }
            }

            return mail;
        }

        /// <summary>
        /// Read and parse an ICal file
        /// </summary>
        /// <param name="icalFile"></param>
        /// <returns></returns>
        private static IEnumerable<IEvent> ReadCalendarExport(String icalFile)
        {
            return File.Exists(icalFile) ? iCalendar.LoadFromFile(icalFile).SelectMany(c => c.Events) : new List<IEvent>();
        }

        /// <summary>
        /// Delete email messages that match a given rule
        /// </summary>
        /// <param name="folder">The folder in which to delete messages</param>
        /// <param name="deleteRule">The rule specifying which messages to delete</param>
        private static void DeleteMessages(OlDefaultFolders folder, string deleteRule)
        {
            Program.Outlook.Session.GetDefaultFolder(folder).Items.Restrict(deleteRule).Cast<MailItem>().ToList().ForEach(deleted => deleted.Delete());
        }
    }
}
