using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
namespace OutlookAI
{
    public partial class OutlookAIRibbon
    {

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // Ui Labels aktualisiern
            UpdateRibbonLabels();
        }

        private void UpdateRibbonLabels()
        {
            button1.Label = ThisAddIn.userdata.Titel1;
            button2.Label = ThisAddIn.userdata.Titel2;
            button3.Label = ThisAddIn.userdata.Titel3;
            button4.Label = ThisAddIn.userdata.Titel4;
            button5.Label = Resources.Settings;
            btnSummary1.Label = ThisAddIn.userdata.SummaryTitel1;
            btnSummary2.Label = ThisAddIn.userdata.SummaryTitel2;

            button1.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Titel1);
            button2.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Titel2);
            button3.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Titel3);
            button4.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Titel4);
            btnSummary1.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Summary1);
            btnSummary2.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.Summary2);
        }

        private async void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            MailItem mail = GetMail();
            await Reply(mail, ThisAddIn.userdata.Prompt1);
        }
        private async void Button2_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = GetMail();
            await Reply(mail, ThisAddIn.userdata.Prompt2);
        }
        private async void Button3_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = GetMail();
            await Reply(mail, ThisAddIn.userdata.Prompt3);
        }
        private async void Button4_Click(object sender, RibbonControlEventArgs e)
        {
            MailItem mail = GetMail();
            InputBox inputBox = new InputBox(ThisAddIn.userdata.Prompt4, "Textinput");
            if (inputBox.ShowDialog() == DialogResult.OK)
            {
                await Reply(mail, ThisAddIn.userdata.Prompt4 + "\n" + inputBox.InputText);
            }
        }
        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            PromptBox p = new PromptBox();
            p.ShowDialog();
            UpdateRibbonLabels();
        }

        private async Task Reply(MailItem mail, string prompt)
        {
            if (mail == null) return;
            try
            {
                string finalPrompt = $"{prompt} \n Hier die zu beantwortende Email:\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}";
                string response = await ThisAddIn.GetLLMResponse(finalPrompt);

                var reply = mail.ReplyAll();
                response = response.Replace("\r\n", "<br>").Replace("\n", "<br>");
                reply.HTMLBody = "<br>" + response + "<br><br>" + reply.HTMLBody;
                reply.Display();
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Fehler: " + ex.Message);
            }
        }



        private MailItem GetMail()
        {
            //Entweder ist der aktuelle Kontext eine Mail - falls das nicht so ist nimm alle Mails aus dem Explorer

            MailItem mail = null;
            var outlookApp = Globals.ThisAddIn.Application;

            try
            {
                var ctx = (Inspector)this.Context;
                mail = ctx.CurrentItem as MailItem;
            }
            catch (System.Exception)
            {//ignore & fallback
            }
            if (mail == null)
            {
                var selection = outlookApp.ActiveExplorer().Selection;
                if (selection.Count > 0 && selection[1] is MailItem selectedmail) { mail = selectedmail; }
            }
            return mail;
        }

        private List<MailItem> GetMails()
        {
            List<MailItem> mails = new List<MailItem>();
            MailItem mail = null;
            var outlookApp = Globals.ThisAddIn.Application;

            try
            {
                var ctx = (Inspector)this.Context;
                mail = ctx.CurrentItem as MailItem;
                mails.Add(mail);
            }
            catch (System.Exception)
            {//ignore & fallback
            }
            if (mail == null)
            {
                var selection = outlookApp.ActiveExplorer().Selection;
                foreach (var item in selection)
                {
                    if (item is MailItem selectedmail)
                    {
                        mails.Add(selectedmail);
                    }
                }
            }
            return mails;
        }




        private async void Summary_Click(object sender, RibbonControlEventArgs e)
        {
            var mails = GetMails();
            await Summarize(mails, ThisAddIn.userdata.Summary1);
        }
        private async void BtnSummary2_Click(object sender, RibbonControlEventArgs e)
        {
            var mails = GetMails();
            await Summarize(mails, ThisAddIn.userdata.Summary2);

        }

        private async Task Summarize(List<MailItem> mails, string prompt = "")
        {
            if (mails == null || mails.Count == 0) return;
            try
            {
                List<Task<string>> responses = new List<Task<string>>();
                var msgs = new List<Task>();

                foreach (var mail in mails)
                {

                    string finalPrompt = $"{prompt} \r\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}";
                    var response = ThisAddIn.GetLLMResponse(finalPrompt);
                    responses.Add(response);
                    msgs.Add(response.ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                        {
                            MessageBox.Show("Fehler: " + t.Exception.InnerException.Message);
                        }
                        else
                        {
                            MessageBox.Show(t.Result);
                        }
                    }));
                }
                await Task.WhenAll(responses);
                await Task.WhenAll(msgs);

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Fehler: " + ex.Message);
            }
        }

        private void Group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            PromptBox p = new PromptBox();
            p.ShowDialog();
            UpdateRibbonLabels();
        }



        public void ExportCalendarToJson(string outputPath)
        {

            Items calendarItems = null;

            try
            {
                // Zugriff auf den Standardkalender
                Microsoft.Office.Interop.Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");
                MAPIFolder calendarFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

                // Alle Kalendereinträge abrufen
                calendarItems = calendarFolder.Items;
                calendarItems.IncludeRecurrences = true;

                List<CalendarEntry> entries = new List<CalendarEntry>();
                foreach (object item2 in calendarItems)
                {
                    if (item2 is AppointmentItem item)
                    {

                        // Nur zukünftige Einträge berücksichtigen
                        if (!item.IsRecurring)
                        {
                            if (item.Start < DateTime.Now.AddDays(-30)) continue;
                            if (item.Start > DateTime.Now.AddDays(60)) continue;
                        }

                        // Kalendereintrag-Daten sammeln
                        CalendarEntry entry = new CalendarEntry
                        {
                            Subject = item.Subject,
                            Body = item.Body,
                            Start = item.Start,
                            End = item.End,
                            IsRecurring = item.IsRecurring,
                            RecurrencePattern = item.IsRecurring ? GetRecurrencePattern(item) : null,
                            Location = item.Location
                        };

                        entries.Add(entry);
                        if (item2 != null) Marshal.ReleaseComObject(item2);

                    }
                }
                // JSON-Datei erstellen
                string json = JsonConvert.SerializeObject(entries, Formatting.Indented);
                File.WriteAllText(outputPath, json);
                System.Windows.Forms.MessageBox.Show("Kalendereinträge erfolgreich exportiert!");

            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Fehler beim Exportieren: " + ex.Message);
            }

            finally
            {
                // COM-Objekte freigeben
                if (calendarItems != null) Marshal.ReleaseComObject(calendarItems);

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private RecurrencePatternData GetRecurrencePattern(AppointmentItem item)
        {
            try
            {
                RecurrencePattern pattern = item.GetRecurrencePattern();
                return new RecurrencePatternData
                {
                    RecurrenceType = pattern.RecurrenceType.ToString(),
                    Interval = pattern.Interval,
                    PatternStartDate = pattern.PatternStartDate,
                    PatternEndDate = pattern.PatternEndDate,
                    Occurrences = pattern.Occurrences
                };
            }
            catch
            {
                return null; // Falls kein gültiges RecurrencePattern vorhanden ist
            }
        }


        public void ImportCalendarFromJson(string inputPath)
        {
            AppointmentItem appointment = null;
            try
            {
                // JSON-Datei einlesen
//                if (!File.Exists(inputPath))
 //               {
 //                   System.Windows.Forms.MessageBox.Show("Die angegebene JSON-Datei wurde nicht gefunden.");
 //                   return;
 //               }
                FileDialog fileDialog = new OpenFileDialog();
                fileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
                fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var dr = fileDialog.ShowDialog();
                if (dr != DialogResult.OK) return;

                string json = File.ReadAllText(fileDialog.FileName);
                List<CalendarEntry> entries = JsonConvert.DeserializeObject<List<CalendarEntry>>(json);

                // Zugriff auf den Kalender "Test"
                Microsoft.Office.Interop.Outlook.Application outlookApp = Globals.ThisAddIn.Application;
                NameSpace outlookNamespace = outlookApp.GetNamespace("MAPI");

                MAPIFolder calendarFolder = outlookNamespace.PickFolder();
                //MAPIFolder calendarFolder = GetOrCreateCalendarFolder(outlookNamespace, "Test");

                // Kalendereinträge hinzufügen
                foreach (var entry in entries)
                {
                    appointment = (AppointmentItem)calendarFolder.Items.Add(OlItemType.olAppointmentItem);
                    appointment.Subject = entry.Subject;
                    appointment.Body = entry.Body;
                    appointment.Start = entry.Start;
                    appointment.End = entry.End;
                    appointment.End = entry.End;
                    appointment.Location = entry.Location;

                    if (entry.IsRecurring && entry.RecurrencePattern != null)
                    {
                        RecurrencePattern pattern = appointment.GetRecurrencePattern();
                        pattern.RecurrenceType = (OlRecurrenceType)Enum.Parse(typeof(OlRecurrenceType), entry.RecurrencePattern.RecurrenceType);
                        pattern.Interval = entry.RecurrencePattern.Interval;
                        pattern.PatternStartDate = entry.RecurrencePattern.PatternStartDate;
                        pattern.PatternEndDate = entry.RecurrencePattern.PatternEndDate < new DateTime(2030, 1, 1) ? entry.RecurrencePattern.PatternEndDate : new DateTime(2030, 1, 1);
                        pattern.Occurrences = entry.RecurrencePattern.Occurrences;
                    }

                    appointment.Save();
                    if (appointment != null) Marshal.ReleaseComObject(appointment);
                }

                System.Windows.Forms.MessageBox.Show("Kalendereinträge erfolgreich importiert!");
            }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Fehler beim Importieren: " + ex.Message);
            }
            finally
            {
                // COM-Objekte freigeben
                if (appointment != null) Marshal.ReleaseComObject(appointment);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private MAPIFolder GetOrCreateCalendarFolder(NameSpace outlookNamespace, string folderName)
        {
            try
            {
                // Prüfen, ob der Kalender "Test" existiert
                MAPIFolder calendarFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar).Parent.Folders[folderName];
                return calendarFolder;
            }
            catch
            {
                // Kalender "Test" erstellen, falls er nicht existiert
                MAPIFolder defaultCalendar = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
                MAPIFolder parentFolder = defaultCalendar.Parent as MAPIFolder;
                return parentFolder.Folders.Add(folderName, OlDefaultFolders.olFolderCalendar);
            }
        }

        private async void Button7_Click(object sender, RibbonControlEventArgs e)
        {
            //ExportCalendarToJson(outputPath: @"C:\Users\Public\Documents\calendar_entries.json");
            ImportCalendarFromJson(inputPath: @"C:\Users\Public\Documents\calendar_entries.json");
        }
    }

    public class CalendarEntry
    {
        public string Subject { get; set; }
        public string Body { get; set; }
        public DateTime Start { get; set; }
        public DateTime End { get; set; }
        public bool IsRecurring { get; set; }
        public RecurrencePatternData RecurrencePattern { get; set; }
        public string Location { get; set; }
    }

    public class RecurrencePatternData
    {
        public string RecurrenceType { get; set; }
        public int Interval { get; set; }
        public DateTime PatternStartDate { get; set; }
        public DateTime PatternEndDate { get; set; }
        public int Occurrences { get; set; }
    }


}
