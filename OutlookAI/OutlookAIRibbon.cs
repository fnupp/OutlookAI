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

        private async Task Summarize(List<MailItem> mails, string prompt ="")
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
    }

}
