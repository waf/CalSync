using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iCal = DDay.iCal;

namespace CalSync.Synchronize
{
    class Receiver
    {
        private const String CalendarItemCategory = "[Calendar Sync]";

        /// <summary>
        /// Given a folder of email messages with ical attachments, add the appointments from the ical attachments
        /// to the provided calendar.
        /// </summary>
        /// <param name="messageFolder">the folder that contains the email messages</param>
        /// <param name="calendar">the calendar to add appointments to</param>
        internal static void ProcessReceivedMessages(MAPIFolder messageFolder, MAPIFolder calendar)
        {
            // read emails from outlook
            var pendingMessages = messageFolder.Items.OfType<MailItem>();

            // determine the calendar items we need to add
            var newItems = ConvertEmailMessagesToCalendarItems(pendingMessages);
            if (!newItems.Any())
                return;
            var existingItems = GetSynchronizedCalendarItems(calendar, newItems.Min(e => e.Start), newItems.Max(e => e.End));
            var itemsToAdd = newItems.Except(existingItems, new AppointmentItemComparer());

            // save these items to the outlook calendar
            itemsToAdd.ToList().ForEach(evt => evt.Save());

            // cleanup processed messages
            pendingMessages.ToList().ForEach(msg => msg.Delete());
        }

        /// <summary>
        /// Given a list of email messages with ical attachments, return a list of outlook calendar appointments
        /// that represent the appointments in the ical attachments
        /// </summary>
        /// <param name="messages">The email messages</param>
        /// <returns>the outlook calendar items</returns>
        private static IEnumerable<AppointmentItem> ConvertEmailMessagesToCalendarItems(IEnumerable<MailItem> messages)
        {
            return messages
                // get email attachments
                .SelectMany(msg => msg.Attachments.OfType<Attachment>())
                // parse attachments to ical
                .SelectMany(attachment =>
                {
                    var attachmentData = attachment.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x37010102");
                    string data = Encoding.Unicode.GetString(attachmentData);
                    using (var memoryStream = new MemoryStream(attachmentData))
                    {
                        return iCal.iCalendar.LoadFromStream(memoryStream);
                    }
                })
                // create outlook events for each event in the ical
                .SelectMany(cal => cal.GetOccurrences(DateTime.MinValue, DateTime.MaxValue))
                .Select(occurrence =>
                {
                    AppointmentItem appt = Program.Outlook.CreateItem(OlItemType.olAppointmentItem);
                    appt.Categories = CalendarItemCategory;
                    appt.Subject = "Busy";
                    appt.Start = occurrence.Period.StartTime.Value;
                    appt.End = occurrence.Period.EndTime.Value;
                    appt.ReminderSet = false;
                    return appt;
                });
        }

        /// <summary>
        /// Retrieve all outlook calendar items created by CalSync for the provided date range.
        /// </summary>
        /// <param name="calendar">The calendar to read</param>
        /// <param name="minDate">The range's begin date</param>
        /// <param name="maxDate">The range's end date</param>
        /// <returns>the outlook calendar items</returns>
        private static IEnumerable<AppointmentItem> GetSynchronizedCalendarItems(MAPIFolder calendar, DateTime minDate, DateTime maxDate)
        {
            string query = String.Format("[Categories] = '{0}' AND ([Start] >= '{1:g}') AND ([End] <= '{2:g}')", CalendarItemCategory, minDate, maxDate);
            var existingEvents = calendar.Items.Restrict(query).OfType<AppointmentItem>();

            return existingEvents;
        }

    }
}
