using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Tab;
using System.Xml.Linq;



using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Security.Policy;
using System.Windows.Forms;

namespace OutlookAI
{
    public partial class Ribbon1
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
                    Prompt2 = "Schreibe mir für die folgende E - Mail eine ToDoListe (Bitte berücksichtige: [kurze Zusammenfassung]) drei Antwortmöglichkeiten:Zustimmende Antwort: Verwende einen freundlichen, professionellen Ton und füge mögliche nächste Schritte hinzu.Ablehnende Antwort: Erkläre die Gründe für die Ablehnung und gib eventuell Alternativen an.Nachfragende Antwort: Stelle klare Fragen zu den Punkten, die unklar sind, um weitere Informationen zu erhalten.Nutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern.",
                    Prompt3 = "Schreibe mir für die folgende Email 3 Antwortmöglichkeiten:\n-Zustimmende Antwort\n-Ablehnede Antwort\n-Nachfragende Antwort\n\nNutze als Sprache der Antwort die Sprache der Email. Erzeuge keinen Emailfooter oder Betreff. \n Schreibe ausführlich und in einem informellen Stil.",
                    Prompt4 = "Schreibe mir für die folgende E - Mail eine Antwort und nimm Bezug auf diese EmailNutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern. Berücksichtige im besonderen die folgenden Punkte:"




                    /*
                    Prompt1 = "Schreibe mir für die folgende E - Mail drei Antwortmöglichkeiten:"
                            + "Zustimmende Antwort: Verwende einen freundlichen, professionellen Ton und füge mögliche nächste Schritte hinzu."
                            + "Ablehnende Antwort: Erkläre die Gründe für die Ablehnung und gib eventuell Alternativen an."
                            + "Ablehnende Antwort: Erkläre die Gründe für die Ablehnung und gib eventuell Alternativen an."
                            + "Nachfragende Antwort: Stelle klare Fragen zu den Punkten, die unklar sind, um weitere Informationen zu erhalten."
                            + "Nutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe knapp und verwende Absätze, um die Argumentation zu gliedern.",
                    


                    Prompt2 = "Schreibe mir für die folgende E - Mail eine ToDoListe (Bitte berücksichtige: [kurze Zusammenfassung]) drei Antwortmöglichkeiten:"
                            + "Zustimmende Antwort: Verwende einen freundlichen, professionellen Ton und füge mögliche nächste Schritte hinzu."
                            + "Ablehnende Antwort: Erkläre die Gründe für die Ablehnung und gib eventuell Alternativen an."
                            + "Nachfragende Antwort: Stelle klare Fragen zu den Punkten, die unklar sind, um weitere Informationen zu erhalten."
                            + "Nutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern.",

                    Prompt3 = "Schreibe mir für die folgende Email 3 Antwortmöglichkeiten:\n"
                        + "-Zustimmende Antwort\n"
                        + "-Ablehnede Antwort\n"
                        + "-Nachfragende Antwort\n\n"
                        + "Nutze als Sprache der Antwort die Sprache der Email. Erzeuge keinen Emailfooter oder Betreff. \n Schreibe ausführlich und in einem informellen Stil.",

                    Prompt4 = "Schreibe mir für die folgende E - Mail eine Antwort und nimm Bezug auf diese Email"
                            + "Nutze als Sprache der Antwort die Sprache der E - Mail.Erzeuge keinen E-Mail - Fußzeilen oder Betreff. Schreibe ausführlich und verwende Absätze, um die Argumentation zu gliedern. Berücksichtige im besonderen die folgenden Punkte:",

                    */
                }
            ;
                string json = JsonConvert.SerializeObject(data);
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            }
            // Abrufen
            string jsonData = File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"));
            UserData loadedData = JsonConvert.DeserializeObject<UserData>(jsonData);
            _userdata = loadedData;
        }

        private async void button1_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = getMail();
            await Reply(mail, _userdata.Prompt1);
        }
        private async void button2_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = getMail();
            await Reply(mail, _userdata.Prompt2);
        }
        private async void button3_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = getMail();
            await Reply(mail, _userdata.Prompt3);
        }
        private async void button4_Click(object sender, RibbonControlEventArgs e)
        {

            MailItem mail = getMail();
            InputBox inputBox = new InputBox("Bitte geben Sie Ihren Text ein:", "Textinput");
            string userInput = string.Empty;
            if (inputBox.ShowDialog() == DialogResult.OK)
            {
                userInput = inputBox.InputText;
                //MessageBox.Show("Eingegebener Text: " + userInput);
            }
            await Reply(mail, _userdata.Prompt4 + "\n" + userInput);
        }

        private async Task Reply(MailItem mail, string prompt)
        {
            try
            {
                string response = await GetChatGPTResponse(prompt + "\n"
                    + $"Absender: {mail.Sender.Name}\nBetreff: {mail.Subject}\nInhalt: {mail.Body}");
                //                System.Windows.Forms.MessageBox.Show(response);
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

        private MailItem getMail()
        {
            MailItem mail = null;
            var outlookApp = Globals.ThisAddIn.Application;

            try
            {
                var ctx = (Inspector)this.Context;
                mail = ctx.CurrentItem as MailItem;
            }
            catch (System.Exception)
            { }
            if (mail == null)
            {
                var selection = outlookApp.ActiveExplorer().Selection;
                if (selection.Count > 0)
                {
                    if (selection[1] is MailItem)
                    {
                        mail = (MailItem)selection[1];
                    }
                }
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

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            PromptBox p = new PromptBox(_userdata);
            p.ShowDialog();
        }
    }
}
