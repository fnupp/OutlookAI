using System;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookAI
{
    /// <summary>
    /// Thread-safe POCO containing email data extracted from Outlook MailItem.
    /// Used to pass email information across thread boundaries without COM objects.
    /// </summary>
    public class EmailData
    {
        public string EntryID { get; set; }
        public string Subject { get; set; }
        public string SenderName { get; set; }
        public string SenderEmailAddress { get; set; }
        public string Body { get; set; }
        public DateTime ReceivedTime { get; set; }

        /// <summary>
        /// Extracts data from a MailItem synchronously.
        /// Must be called on the STA thread that created the MailItem.
        /// </summary>
        public static EmailData FromMailItem(Outlook.MailItem mailItem)
        {
            if (mailItem == null)
                throw new ArgumentNullException(nameof(mailItem));

            return new EmailData
            {
                EntryID = mailItem.EntryID,
                Subject = mailItem.Subject ?? "",
                SenderName = mailItem.SenderName ?? "",
                SenderEmailAddress = mailItem.SenderEmailAddress ?? "",
                Body = mailItem.Body ?? "",
                ReceivedTime = mailItem.ReceivedTime
            };
        }
    }
}
