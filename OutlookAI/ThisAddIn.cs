using Newtonsoft.Json;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace OutlookAI
{
    public partial class ThisAddIn
    {

        internal static UserData userdata;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Hinweis: Outlook löst dieses Ereignis nicht mehr aus. Wenn Code vorhanden ist, der 
            //    muss ausgeführt werden, wenn Outlook heruntergefahren wird. Weitere Informationen finden Sie unter https://go.microsoft.com/fwlink/?LinkId=506785.
        }

        public static async Task<string> GetChatOllamaResponse(string prompt)
        {
            //var ollamaUrl = "http://localhost:11434/api/generate";
            //var model = "llama3"; 



            using (var client = new HttpClient())
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

        public static async Task<string> GetChatGPTResponse(string userInput)
        {
            using (HttpClient client = new HttpClient())
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
