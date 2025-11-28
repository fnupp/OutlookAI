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
using System.Drawing.Printing;
using System.ComponentModel.Design.Serialization;
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
            using (ErrorLogger.BeginCorrelation("RibbonReply-Button1"))
            {
                ErrorLogger.LogInfo("Reply button 1 clicked");

                MailItem mail = GetMail();
                InputBox inputBox = new InputBox(ThisAddIn.userdata.Prompt4, "Textinput");
                if (inputBox.ShowDialog() == DialogResult.OK)
                {
                    string prompt = ThisAddIn.userdata.Prompt1.Replace("§§Input§§", inputBox.InputText);
                    await Reply(mail, prompt);
                }

                ErrorLogger.LogInfo("Reply operation completed");
            }
        }
        private async void Button2_Click(object sender, RibbonControlEventArgs e)
        {
            using (ErrorLogger.BeginCorrelation("RibbonReply-Button2"))
            {
                ErrorLogger.LogInfo("Reply button 2 clicked");

                MailItem mail = GetMail();
                InputBox inputBox = new InputBox(ThisAddIn.userdata.Prompt4, "Textinput");
                if (inputBox.ShowDialog() == DialogResult.OK)
                {
                    string prompt = ThisAddIn.userdata.Prompt2.Replace("§§Input§§", inputBox.InputText);
                    await Reply(mail, prompt);
                }

                ErrorLogger.LogInfo("Reply operation completed");
            }
        }
        private async void Button3_Click(object sender, RibbonControlEventArgs e)
        {
            using (ErrorLogger.BeginCorrelation("RibbonReply-Button3"))
            {
                ErrorLogger.LogInfo("Reply button 3 clicked");

                MailItem mail = GetMail();
                InputBox inputBox = new InputBox(ThisAddIn.userdata.Prompt4, "Textinput");
                if (inputBox.ShowDialog() == DialogResult.OK)
                {
                    string prompt = ThisAddIn.userdata.Prompt3.Replace("§§Input§§", inputBox.InputText);
                    await Reply(mail, prompt);
                }

                ErrorLogger.LogInfo("Reply operation completed");
            }
        }
        private async void Button4_Click(object sender, RibbonControlEventArgs e)
        {
            using (ErrorLogger.BeginCorrelation("RibbonReply-Button4"))
            {
                ErrorLogger.LogInfo("Reply button 4 clicked");

                MailItem mail = GetMail();
                InputBox inputBox = new InputBox(ThisAddIn.userdata.Prompt4, "Textinput");
                if (inputBox.ShowDialog() == DialogResult.OK)
                {
                    string prompt = ThisAddIn.userdata.Prompt4.Replace("§§Input§§", inputBox.InputText);
                    await Reply(mail, prompt);
                }

                ErrorLogger.LogInfo("Reply operation completed");
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
            if (mail == null)
            {
                ErrorLogger.LogWarning("Reply called with null mail item");
                return;
            }

            try
            {
                ErrorLogger.LogInfo($"Generating reply for email: {mail.Subject}");

                string finalPrompt = $"{prompt} \n Hier die zu beantwortende Email:\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}";
                string response = await ThisAddIn.GetLLMResponse(finalPrompt);

                var reply = mail.ReplyAll();
                response = response.Replace("\r\n", "<br>").Replace("\n", "<br>");
                reply.HTMLBody = "<br>" + response + "<br><br>" + reply.HTMLBody;
                reply.Display();

                ErrorLogger.LogInfo("Reply generated successfully");
            }
            catch (System.Exception ex)
            {
                ErrorLogger.LogError("Failed to generate reply", ex);
                System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.ErrorGeneric + ex.Message);
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
            using (ErrorLogger.BeginCorrelation("RibbonSummary-Single"))
            {
                ErrorLogger.LogInfo("Summary button clicked");

                var mails = GetMails();
                if (mails != null && mails.Count > 0)
                {
                    if (ThisAddIn.userdata.SummaryMultiple1 ?? true)
                        await SummarizeMultiple(mails, ThisAddIn.userdata.Summary1);
                    else
                        await Summarize(mails, ThisAddIn.userdata.Summary1);
                }
                else
                {
                    ErrorLogger.LogWarning("No emails selected for summary");
                }

                ErrorLogger.LogInfo("Summary operation completed");
            }
        }
        private async void Summary2_Click(object sender, RibbonControlEventArgs e)
        {
            using (ErrorLogger.BeginCorrelation("RibbonSummary-Multiple"))
            {
                ErrorLogger.LogInfo("Summary 2 button clicked");

                var mails = GetMails();
                if (mails != null && mails.Count > 0)
                {
                    if (ThisAddIn.userdata.SummaryMultiple2 ?? true)
                        await SummarizeMultiple(mails, ThisAddIn.userdata.Summary2);
                    else
                        await Summarize(mails, ThisAddIn.userdata.Summary2);
                }
                else
                {
                    ErrorLogger.LogWarning("No emails selected for summary");
                }

                ErrorLogger.LogInfo("Summary operation completed");
            }
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
                            MessageBox.Show(OutlookAI.Resources.ErrorGeneric + t.Exception.InnerException.Message);
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
                System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.ErrorGeneric + ex.Message);
            }
        }
        private async Task SummarizeMultiple(List<MailItem> mails, string prompt = "")
        {
            if (mails == null || mails.Count == 0) return;
            try
            {
                List<Task<string>> responses = new List<Task<string>>();
                var msgs = new List<Task>();

                StringBuilder sb = new StringBuilder(prompt);
                foreach (var mail in mails)
                {
                    sb.AppendLine($"$--$\r\nAbsender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}");
                }

                var response = ThisAddIn.GetLLMResponse(sb.ToString());
                await response;
                if (response.IsFaulted)
                {
                    MessageBox.Show(OutlookAI.Resources.ErrorGeneric + response.Exception.InnerException.Message);
                }
                else
                {
                    MessageBox.Show(response.Result);
                }
        }
            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.ErrorGeneric + ex.Message);
            }
}

private void Group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
{
    PromptBox p = new PromptBox();
    p.ShowDialog();
    UpdateRibbonLabels();
}
public void ExportCalendarToJson()
{

    Items calendarItems = null;

    try
    {

        FileDialog fileDialog = new SaveFileDialog();
        fileDialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*";
        fileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        var dr = fileDialog.ShowDialog();
        if (dr != DialogResult.OK) return;

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
                    if (item.Start < DateTime.Now.AddDays(-20)) continue;
                    if (item.Start > DateTime.Now.AddDays(20)) continue;
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
        File.WriteAllText(fileDialog.FileName, json);
        System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.CalendarExportSuccess);

    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.CalendarExportError + ex.Message);
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


public void ImportCalendarFromJson()
{
    AppointmentItem appointment = null;
    try
    {
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
        if (calendarFolder == null) return;
        //MAPIFolder calendarFolder = GetOrCreateCalendarFolder(outlookNamespace, "Test");

        var oldcalendarItems = calendarFolder.Items;
        oldcalendarItems.IncludeRecurrences = true;

        foreach (object item2 in oldcalendarItems)
        {
            if (item2 is AppointmentItem item)
            {
                if (item.Categories == "Synced")
                {
                    item.Delete();
                }
            }
        }

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
            appointment.Categories = "Synced";

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

        System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.CalendarImportSuccess);
    }
    catch (System.Exception ex)
    {
        System.Windows.Forms.MessageBox.Show(OutlookAI.Resources.CalendarImportError + ex.Message);
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

private void ExportSync_Click(object sender, RibbonControlEventArgs e)
{
    ExportCalendarToJson();
    // ImportCalendarFromJson(); // inputPath: @"C:\Users\Public\Documents\calendar_entries.json");
}

private void Import_Click(object sender, RibbonControlEventArgs e)
{
    ImportCalendarFromJson();
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
