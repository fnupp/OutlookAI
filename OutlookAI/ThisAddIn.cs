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
        private static readonly object _httpClientLock = new object();
        private static Lazy<HttpClient> _httpClient = new Lazy<HttpClient>(() => CreateHttpClientInternal());
        private static Lazy<HttpClient> _httpClientWithProxy = new Lazy<HttpClient>(() => CreateHttpClientInternal(useProxy: true));

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
            var client = GetHttpClient();
            var requestBody = new
            {
                model = ThisAddIn.userdata.Ollamamodel,
                prompt,
                stream = false
            };

            var json = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(json, Encoding.UTF8, "application/json");
            var ollamaUrl = ThisAddIn.userdata.OllamaUrl;
            if (!ThisAddIn.userdata.OllamaUrl.EndsWith("/"))
                ollamaUrl += "/";
            ollamaUrl += "api/generate";
            var response = await client.PostAsync(ollamaUrl, content).ConfigureAwait(false);

            if (response.IsSuccessStatusCode)
            {
                string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                dynamic jsonResponseParsed = JsonConvert.DeserializeObject(jsonResponse);
                return jsonResponseParsed.response.ToString();
            }
            else
            {
                throw new System.Exception($"{OutlookAI.Resources.ErrorcallingOllama}: {response.StatusCode}\n{await response.Content.ReadAsStringAsync().ConfigureAwait(false)}");
            }
        }
        private static async Task<string> GetChatGPTResponse(string userInput)
        {
            var client = GetHttpClient();

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

            // Create request message with authorization header
            using (var request = new HttpRequestMessage(HttpMethod.Post, ThisAddIn.userdata.OpenAIAPIUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", ThisAddIn.userdata.OpenAIAPIKey);
                request.Content = content;

                HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(false);

                if (response.IsSuccessStatusCode)
                {
                    string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    dynamic jsonResponseParsed = JsonConvert.DeserializeObject(jsonResponse);
                    return jsonResponseParsed.choices[0].message.content.ToString();
                }
                else
                {
                    throw new System.Exception($"{OutlookAI.Resources.ErrorcallingOpenai}: {response.StatusCode}\n{await response.Content.ReadAsStringAsync().ConfigureAwait(false)}");
                }
            }
        }


        /// <summary>
        /// Gets the appropriate HttpClient instance based on proxy settings.
        /// Uses static instances to avoid socket exhaustion.
        /// </summary>
        public static HttpClient GetHttpClient()
        {
            lock (_httpClientLock)
            {
                if (ThisAddIn.userdata.ProxyActive)
                {
                    return _httpClientWithProxy.Value;
                }
                return _httpClient.Value;
            }
        }

        /// <summary>
        /// Creates a new HttpClient instance with optional proxy configuration.
        /// This method is called lazily only once per configuration.
        /// </summary>
        private static HttpClient CreateHttpClientInternal(bool useProxy = false)
        {
            if (useProxy)
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
                return new HttpClient(handler)
                {
                    Timeout = TimeSpan.FromMinutes(5) // Reasonable timeout for LLM calls
                };
            }
            return new HttpClient()
            {
                Timeout = TimeSpan.FromMinutes(5) // Reasonable timeout for LLM calls
            };
        }

        /// <summary>
        /// Invalidates and disposes existing HttpClient instances.
        /// Call this method when proxy settings or other HTTP configuration changes.
        /// New instances will be created on the next GetHttpClient() call.
        /// </summary>
        public static void InvalidateHttpClients()
        {
            lock (_httpClientLock)
            {
                // Dispose existing clients if they were initialized
                if (_httpClient.IsValueCreated)
                {
                    _httpClient.Value?.Dispose();
                }
                if (_httpClientWithProxy.IsValueCreated)
                {
                    _httpClientWithProxy.Value?.Dispose();
                }

                // Recreate lazy instances with fresh factory methods
                _httpClient = new Lazy<HttpClient>(() => CreateHttpClientInternal());
                _httpClientWithProxy = new Lazy<HttpClient>(() => CreateHttpClientInternal(useProxy: true));
            }
        }

        [Obsolete("Use GetHttpClient() instead. This method is kept for backward compatibility.")]
        public static HttpClient CreateHttpClient()
        {
            return GetHttpClient();
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
                    Prompt1 = OutlookAI.Resources.Prompt1,
                    Prompt2 = OutlookAI.Resources.Prompt2,
                    Prompt3 = OutlookAI.Resources.Prompt3,
                    Prompt4 = OutlookAI.Resources.Prompt4,
                    Titel1 = OutlookAI.Resources.Title1,
                    Titel2 = OutlookAI.Resources.Title2,
                    Titel3 = OutlookAI.Resources.Title3,
                    Titel4 = OutlookAI.Resources.Title4,
                    OpenAIAPIActive = false,
                    OpenAIAPIKey = "",
                    OpenAIAPIModel = OutlookAI.Resources.OpenAiDefaultModel,
                    OpenAIAPIUrl = "https://api.openai.com/v1/chat/completions",
                    OllamaActive = false,
                    OllamaUrl = "http://localhost:11434",
                    ComposePrompt1 = OutlookAI.Resources.ComposePrompt1,
                    ComposePrompt2 = OutlookAI.Resources.ComposePrompt2,
                    ComposePrompt3 = OutlookAI.Resources.ComposePrompt3,
                    ComposeTitle1 = OutlookAI.Resources.ComposePromptTitle1,
                    ComposeTitle2 = OutlookAI.Resources.ComposePromptTitle2,
                    ComposeTitle3 = OutlookAI.Resources.ComposePromptTitle3,
                    ProxyActive = false,
                    SummaryTitel1 = OutlookAI.Resources.SummarizeTitle1,
                    SummaryTitel2 = OutlookAI.Resources.SummarizeTitle2,
                    Summary1 = OutlookAI.Resources.SummarizePrompt1,
                    Summary2 = OutlookAI.Resources.SummarizePrompt2,
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
