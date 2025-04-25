using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
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
        UserData _userdata;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            FileInfo fi = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"));
            if (!fi.Exists)
            {
                // Speichern
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
                    ApiKey = "",

                };

                string json = JsonConvert.SerializeObject(data);
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            }

            string jsonData = File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"));
            UserData loadedData = JsonConvert.DeserializeObject<UserData>(jsonData);
            _userdata = loadedData;

            this.button1.Label = _userdata.Titel1;
            this.button2.Label = _userdata.Titel2;
            this.button3.Label = _userdata.Titel3;
            this.button4.Label = _userdata.Titel4;
            this.button5.Label = "Einstellungen";
        }

        private async void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            MailItem mail = GetMail();
            await Reply(mail, _userdata.Prompt1);
        }
        private async void Button2_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = GetMail();
            await Reply(mail, _userdata.Prompt2);
        }
        private async void Button3_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = GetMail();
            await Reply(mail, _userdata.Prompt3);
        }
        private async void Button4_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = GetMail();
            InputBox inputBox = new InputBox(_userdata.Prompt4, "Textinput");
            string userInput = string.Empty;
            if (inputBox.ShowDialog() == DialogResult.OK)
            {
                userInput = inputBox.InputText;
                //MessageBox.Show("Eingegebener Text: " + userInput);
            }
            await Reply(mail, _userdata.Prompt4 + "\n" + userInput);
        }
        private void Button5_Click(object sender, RibbonControlEventArgs e)
        {
            PromptBox p = new PromptBox(_userdata);
            p.ShowDialog();
        }

        private async Task Reply(MailItem mail, string prompt)
        {
            if (mail == null) return;
            try
            {
                string response = await GetChatGPTResponse($"{prompt} \n Hier die zu beantwortende Email:\n Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}");
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

        private async Task<string> GetChatGPTResponse(string userInput)
        {
            // Ersetzen Sie "YOUR_API_KEY" durch Ihren tatsächlichen API-Schlüssel
            string apiKey = _userdata.ApiKey; //                
            string apiUrl = "https://api.openai.com/v1/chat/completions";

            using (HttpClient client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                // Erstellen des Anforderungskörpers
                var requestBody = new
                {
                    model = "gpt-4o-mini", // Das GPT-Modell, das Sie verwenden möchten
                    messages = new[]
                    {
                        new { role = "user", content = userInput }
                    }
                };

                string jsonRequestBody = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonRequestBody, Encoding.UTF8, "application/json");

                // Senden der POST-Anfrage
                HttpResponseMessage response = await client.PostAsync(apiUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    dynamic jsonResponseParsed = JsonConvert.DeserializeObject(jsonResponse);
                    return jsonResponseParsed.choices[0].message.content.ToString();
                }
                else
                {
                    throw new System.Exception($"Fehler bei der Anfrage an ChatGPT: {response.StatusCode}\n{await response.Content.ReadAsStringAsync()}");
                }
            }
        }

    }
}
