using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Deployment.Application;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace OutlookAI
{
    public partial class PromptBox : Form
    {
        public PromptBox()
        {
            InitializeComponent();


            userDataBindingSource.DataSource = ThisAddIn.userdata;
            if (ApplicationDeployment.IsNetworkDeployed)
                labelVersion.Text = ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString();
            else
                labelVersion.Text = "Nicht veröffentlicht";

        }


        private void OK_Click(object sender, EventArgs e)
        {
            string json = JsonConvert.SerializeObject(userDataBindingSource.DataSource);
            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            this.Close();
        }

        private async void Button2_Click(object sender, EventArgs e)
        {
            await GetOllamaModels();
        }


        public async Task<List<ModelInfo>> GetOllamaModels()
        {
            var ollamaUrl = "http://localhost:11434/api/tags";
            try
            {
                HttpClient httpClient = ThisAddIn.CreateHttpClient();

                var response = await httpClient.GetAsync(ollamaUrl);

                if (!response.IsSuccessStatusCode)
                {
                    throw new System.Exception($"Fehler bei der Anfrage an oLLAMA: {response.StatusCode}\n{await response.Content.ReadAsStringAsync()}");
                }

                string jsonResponse = await response.Content.ReadAsStringAsync();
                ModelListResponse modelListResponse = JsonConvert.DeserializeObject<ModelListResponse>(jsonResponse);


                if (modelListResponse?.Models != null)
                {
                    if (ThisAddIn.userdata.OllamaModels == null)
                        ThisAddIn.userdata.OllamaModels = new List<string>();
                    else
                        ThisAddIn.userdata.OllamaModels.Clear();
                    ThisAddIn.userdata.OllamaModels.AddRange(modelListResponse.Models.ConvertAll(m => m.Name));
                }
                return modelListResponse.Models;
            }
            catch (System.Exception ex)
            {
                throw new System.Exception($"Fehler bei der Anfrage an oLLAMA: {ex.Message}");
            }
        }


        private void ListBox1_Click(object sender, EventArgs e)
        {
            textBox2.Text = listBox1.SelectedItem.ToString();
        }

    }

    public class ModelListResponse
    {
        public List<ModelInfo> Models { get; set; }
    }

    public class ModelInfo
    {
        public string Name { get; set; }
        public string Mod { get; set; }
        public string Size { get; set; }

        public string CreatedAt { get; set; }
    }
}
