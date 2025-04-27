using Newtonsoft.Json;
using System.IO;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Net;

namespace OutlookAI
{
    public partial class ThisAddIn
    {

        internal static UserData userdata;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
            InitSettingsFile();

            //lade Settings
            string jsonData = File.ReadAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI", "OutlookAI.json"));
            UserData loadedData = JsonConvert.DeserializeObject<UserData>(jsonData);
            ThisAddIn.userdata = loadedData;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }


        public  static async Task<string> GetLLMResponse(string prompt)
        {
            string response;
            if (ThisAddIn.userdata.OllamaActive)
            {
                response = await ThisAddIn.GetChatOllamaResponse(prompt);
            }
            else if (ThisAddIn.userdata.OpenAIAPIActive)
            {
                response = await ThisAddIn.GetChatGPTResponse(prompt);
            }
            else
            {
                response = "No active LLM. Active in Settings.";
            }

            return response;
        }

        private static async Task<string> GetChatOllamaResponse(string prompt)
        {
            //var ollamaUrl = "http://localhost:11434/api/generate";
            //var model = "llama3"; 



            using (var client = CreateHttpClient())
            {
                var requestBody = new
                {
                    model = ThisAddIn.userdata.Ollamamodel,
                    prompt,
                    stream = false
                };

                var json = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                var response = await client.PostAsync(ThisAddIn.userdata.OllamaUrl, content);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync();
                    dynamic jsonResponseParsed = JsonConvert.DeserializeObject(jsonResponse);
                    return jsonResponseParsed.response.ToString();
                }
                else
                {
                    throw new System.Exception($"Fehler bei der Anfrage an oLLAMA: {response.StatusCode}\n{await response.Content.ReadAsStringAsync()}");
                }
            }
        }

        private static async Task<string> GetChatGPTResponse(string userInput)
        {
            using (HttpClient client = CreateHttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", ThisAddIn.userdata.OpenAIAPIKey);

                var requestBody = new
                {
                    model = ThisAddIn.userdata.OpenAIAPIModel,  //"gpt-4o-mini", 
                    messages = new[]
                    {
                        new { role = "user", content = userInput }
                    }
                };

                string jsonRequestBody = JsonConvert.SerializeObject(requestBody);
                var content = new StringContent(jsonRequestBody, Encoding.UTF8, "application/json");

                HttpResponseMessage response = await client.PostAsync(ThisAddIn.userdata.OpenAIAPIUrl, content);

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


        public static HttpClient CreateHttpClient()
        {
            if (ThisAddIn.userdata.ProxyActive)
            {
                var proxy = new WebProxy(ThisAddIn.userdata.ProxyUrl)
                {
                    Credentials = new NetworkCredential(ThisAddIn.userdata.ProxyUsername, ThisAddIn.userdata.ProxyPassword)
                };

                var handler = new HttpClientHandler
                {
                    Proxy = proxy,
                    UseProxy = true
                };
                return new HttpClient(handler);
            }
            return new HttpClient();
        }

        private static void InitSettingsFile()
        {
            FileInfo fi = new FileInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI", "OutlookAI.json"));
            if (!fi.Directory.Exists)
                fi.Directory.Create();
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
                File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"OutlookAI", "OutlookAI.json"), json);
            }
        }

        #region Von VSTO generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
