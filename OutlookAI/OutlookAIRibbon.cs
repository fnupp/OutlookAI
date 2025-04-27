using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace OutlookAI
{
    public partial class OutlookAIRibbon
    {
       
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            FileInfo fi = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"));
            if (!fi.Exists)
            {
                // initiale Befüllung
                UserData data = new UserData
                {
                    Prompt1 = "Schreibe mir für die folgende E - Mail drei Antwortmöglichkeiten:\r\n1.Zustimmende Antwort: Verwende einen freundlichen, professionellen Ton und füge mögliche nächste Schritte hinzu.\r\n2.Ablehnende Antwort: Erkläre die Gründe für die Ablehnung und gib eventuell Alternativen an.\r\n3.Nachfragende Antwort: Stelle klare Fragen zu den Punkten, die unklar sind, um weitere Informationen zu erhalten.\r\n\r\nNutze als Sprache der Antwort die Sprache der E - Mail. Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe knapp und verwende Absätze, um die Argumentation zu gliedern.",
                    Prompt2 = "Schreibe mir für die folgende E - Mail eine ToDoListeSchreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern.",
                    Prompt3 = "Schreibe mir für die folgende Email eine Antwort micht 3 Rückfragen \nNutze als Sprache der Antwort die Sprache der Email. Erzeuge keinen Emailfooter oder Betreff. \n Schreibe ausführlich und in einem informellen Stil.",
                    Prompt4 = "Schreibe mir für die folgende E - Mail eine Antwort und nimm Bezug auf diese EmailNutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern. Berücksichtige im besonderen die folgenden Punkte:",
                    Titel1 = "3 Antworten",
                    Titel2 = "ToDo",
                    Titel3 = "Rückfragen",
                    Titel4 = "Custom",
                    OpenAIAPIActive = false,
                    OpenAIAPIKey = "",
                    OpenAIAPIModel = "gpt-4o-mini",
                    OpenAIAPIUrl = "https://api.openai.com/v1/chat/completions",
                    OllamaActive = false,
                    OllamaUrl = "",
                    ComposePrompt1 = "Formuliere diese Email professioneller. \r\n - Erzeuge keinen Betreff oder Signatur.\r\n - Behalte die Anrede (Du, sie) bei\r\n",
                    ComposePrompt2 = "Überarbeite diese E-Mail so, dass sie klarer strukturiert und leichter verständlich ist, ohne den Inhalt zu verändern. Erzeuge keinen Betreff oder Signatur. Behalte die Anrede (Du, sie) bei:",
                    ComposePrompt3 = "Mach diese E-Mail kürzer und persönlicher, als würdest du einem guten Kollegen oder einer Bekannten schreiben:\r\n - Erzeuge keinen Betreff oder Signatur.\r\n - Behalte die Anrede (Du, sie) bei",
                    ComposeTitle1 = "Professioneller",
                    ComposeTitle2 = "Klarer ",
                    ComposeTitle3 = "Informeller",
                    ProxyActive = false,
                    SummaryTitel1 = "Zusammenfassung 1",
                    SummaryTitel2 = "Zusammenfassung 2",
                    Summary1 = "Fasse die folgende Eamil zusammen, zähle die zentralen Aussagen und Informationen auf und beschreibe den Ton der Email.\r\n\r\n\r\n   E-Mail analysieren:\r\n        Lies die Ursprungs-E-Mail, um den Kontext, das Anliegen und den Ton des Absenders zu verstehen.\r\n\r\n    Inhaltlichen Input einbauen:\r\n        Verwende ausschließlich die bereitgestellten inhaltlichen Informationen und Aussagen als Basis für die Antwort.\r\n        ",
                    Summary2 = "Fasse die folgende Eamil zusammen.Lies die Ursprungs-E-Mail, um den Kontext, das Anliegen und den Ton des Absenders zu verstehen.",
                };
                string json = JsonConvert.SerializeObject(data);
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            }

            //lade Settings
            string jsonData = File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"));
            UserData loadedData = JsonConvert.DeserializeObject<UserData>(jsonData);
            ThisAddIn.userdata = loadedData;

            // Ui Labels aktualisiern
            UpdateRibbonLabels();
        }

        private void UpdateRibbonLabels()
        {
            button1.Label = ThisAddIn.userdata.Titel1;
            button2.Label = ThisAddIn.userdata.Titel2;
            button3.Label = ThisAddIn.userdata.Titel3;
            button4.Label = ThisAddIn.userdata.Titel4;
            button5.Label = "Einstellungen";
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
            PromptBox p = new PromptBox(ThisAddIn.userdata);
            p.ShowDialog();
            UpdateRibbonLabels();
        }

        private async Task Reply(MailItem mail, string prompt)
        {
            if (mail == null) return;
            try
            {
                string response;
                string finalPrompt = $"{prompt} \n Hier die zu beantwortende Email:\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}";

//                System.Windows.Forms.MessageBox.Show(finalPrompt);
                if (ThisAddIn.userdata.OllamaActive)
                {
                    response = await ThisAddIn.GetChatOllamaResponse(finalPrompt);
                }
                else if (ThisAddIn.userdata.OpenAIAPIActive)
                {
                    response = await ThisAddIn.GetChatGPTResponse(finalPrompt);
                }
                else
                {
                    response = "No active LLM. Active in Settings.";
                }

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

        private async Task Summarize(List<MailItem> mails, string prompt ="")
        {
            if (mails == null || mails.Count == 0) return;
            try
            {
                List<Task<string>> responses = new List<Task<string>>();
                var msgs = new List<Task>();
                
                foreach (var mail in mails)
                {

                    Task<string> response;
                    string finalPrompt = $"{prompt} \r\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}";

                    //                System.Windows.Forms.MessageBox.Show(finalPrompt);
                    if (ThisAddIn.userdata.OllamaActive)
                    {
                        response = ThisAddIn.GetChatOllamaResponse(finalPrompt);
                    }
                    else if (ThisAddIn.userdata.OpenAIAPIActive)
                    {
                        response = ThisAddIn.GetChatGPTResponse(finalPrompt);
                    }
                    else
                    {
                        MessageBox.Show("No active LLM. Activate in Settings.");
                        return;
                    }
                    responses.Add(response);

                    msgs.Add(response.ContinueWith(t =>
                    {
                        if (t.IsFaulted)
                        {
                            MessageBox.Show("Fehler: " + t.Exception.InnerException.Message);
                        }
                        else
                        {
                            string result = t.Result;
                            MessageBox.Show(result);
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
    }

}
