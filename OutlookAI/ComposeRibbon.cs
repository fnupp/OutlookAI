using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OutlookAI
{
    public partial class ComposeRibbon
    {


       
        private void ComposeRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            UpdateRibbonLabels();
        }

        private void UpdateRibbonLabels()
        {
            btnCompose1.Label = ThisAddIn.userdata.ComposeTitle1;
            btnCompose2.Label = ThisAddIn.userdata.ComposeTitle2;
            btnCompose3.Label = ThisAddIn.userdata.ComposeTitle3;

            btnCompose1.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.ComposeTitle1);
            btnCompose2.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.ComposeTitle2);
            btnCompose3.Visible = !string.IsNullOrEmpty(ThisAddIn.userdata.ComposeTitle3);
        }

        /*

       private string originalBody;

       public void TrackChanges(MailItem mailItem)
       {
           // Speichern des ursprünglichen Inhalts
           originalBody = mailItem.Body;
       }

       public string GetNewText(MailItem mailItem)
       {
           // Vergleichen des aktuellen Inhalts mit dem ursprünglichen
           string currentBody = mailItem.HTMLBody;

           // Prüfen, ob der ursprüngliche Text im aktuellen Text enthalten ist
           int originalTextIndex = currentBody.IndexOf(originalBody, StringComparison.Ordinal);
           if (originalTextIndex > 0)
           {
               // Der neue Text liegt vor dem ursprünglichen Text
               return currentBody.Substring(0, originalTextIndex).Trim();
           }
           return string.Empty;
       }
*/

        private async void BtnCompose_Click(object sender, RibbonControlEventArgs e)
        {
            var ctx = (Inspector)this.Context;
            var mail = ctx.CurrentItem as MailItem;
            string selectedText = GetSelectedText(ctx);
            if (string.IsNullOrWhiteSpace(selectedText))
            { 
                return;
            }
            await ThisAddIn.GetLLMResponse(ThisAddIn.userdata.ComposePrompt1 + " \r\n" + selectedText).ContinueWith(UpdateMail(mail));
        }

        private async void BtnCompose2_Click(object sender, RibbonControlEventArgs e)
        {

            var ctx = (Inspector)this.Context;
            var mail = ctx.CurrentItem as MailItem;

            string selectedText = GetSelectedText(ctx);

            await ThisAddIn.GetLLMResponse(ThisAddIn.userdata.ComposePrompt2 + " \r\n" + selectedText).ContinueWith(UpdateMail(mail));
        }

        private async void BtnCompose3_Click(object sender, RibbonControlEventArgs e)
        {

            var ctx = (Inspector)this.Context;
            var mail = ctx.CurrentItem as MailItem;

            string selectedText = GetSelectedText(ctx);

            await ThisAddIn.GetLLMResponse(ThisAddIn.userdata.ComposePrompt3 + " \r\n" + selectedText).ContinueWith(UpdateMail(mail));
        }


        private static Action<System.Threading.Tasks.Task<string>> UpdateMail(MailItem mail)
        {
            return task =>
            {
                if (!task.IsFaulted)
                {
                    string response = task.Result;
                    response = response.Replace("\r\n", "<br>").Replace("\n", "<br>");
                    mail.HTMLBody = response + "\n\n" + mail.HTMLBody;
                    mail.Display();
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Fehler bei der Verarbeitung: " + task.Exception?.Message);
                }
            };
        }

        private static string GetSelectedText(Inspector ctx)
        {
            dynamic wordEditor = ctx.WordEditor;
            var selection = wordEditor.Application.Selection;
            if (selection == null)// || selection.Type != Microsoft.Office.Interop.Word.WdSelectionType.wdSelectionNormal)
            {
                System.Windows.Forms.MessageBox.Show("Bitte wählen Sie den Text aus, den Sie umformulieren möchten.");
                return "";
            }

            return selection.Text;
        }

        private void Group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            PromptBox p = new PromptBox();
            p.ShowDialog();
            UpdateRibbonLabels();

        }
    }
}
