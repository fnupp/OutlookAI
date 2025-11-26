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
                labelVersion.Text = OutlookAI.Resources.NichtVeröffentlicht;


        }

        private void InitializeEmailMonitoringTab()
        {
            // Load mailboxes
            LoadAvailableMailboxes();

            // Load categories
            LoadEmailCategories();
        }


        private void LoadAvailableMailboxes()
        {
            checkedListBoxMailboxes.Items.Clear();

            try
            {
                var outlookApp = Globals.ThisAddIn.Application;
                foreach (Microsoft.Office.Interop.Outlook.Store store in outlookApp.Session.Stores)
                {
                    string storeName = store.DisplayName;
                    checkedListBoxMailboxes.Items.Add(storeName);

                    // Check if this mailbox is monitored
                    if (ThisAddIn.userdata.MonitoredMailboxes != null &&
                        ThisAddIn.userdata.MonitoredMailboxes.Contains(storeName))
                    {
                        int index = checkedListBoxMailboxes.Items.IndexOf(storeName);
                        checkedListBoxMailboxes.SetItemChecked(index, true);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading mailboxes: {ex.Message}");
            }
        }

        private void LoadEmailCategories()
        {
            listBoxCategories.DataSource = null;
            listBoxCategories.DisplayMember = "CategoryName";

            if (ThisAddIn.userdata.EmailCategories == null)
                ThisAddIn.userdata.EmailCategories = new List<EmailCategory>();

            listBoxCategories.DataSource = ThisAddIn.userdata.EmailCategories;
        }

        private void buttonRefreshMailboxes_Click(object sender, EventArgs e)
        {
            LoadAvailableMailboxes();
        }

        private void buttonAddCategory_Click(object sender, EventArgs e)
        {
            var editorForm = new CategoryEditorForm();
            if (editorForm.ShowDialog() == DialogResult.OK)
            {
                ThisAddIn.userdata.EmailCategories.Add(editorForm.Category);
                LoadEmailCategories();
            }
        }

        private void buttonEditCategory_Click(object sender, EventArgs e)
        {
            if (listBoxCategories.SelectedItem is EmailCategory selectedCategory)
            {
                var editorForm = new CategoryEditorForm(selectedCategory);
                if (editorForm.ShowDialog() == DialogResult.OK)
                {
                    // Update the category in the list
                    int index = ThisAddIn.userdata.EmailCategories.IndexOf(selectedCategory);
                    if (index >= 0)
                    {
                        ThisAddIn.userdata.EmailCategories[index] = editorForm.Category;
                        LoadEmailCategories();
                    }
                }
            }
            else
            {
                MessageBox.Show("Please select a category to edit.");
            }
        }

        private void buttonDeleteCategory_Click(object sender, EventArgs e)
        {
            if (listBoxCategories.SelectedItem is EmailCategory selectedCategory)
            {
                var result = MessageBox.Show(
                    $"Are you sure you want to delete the category '{selectedCategory.CategoryName}'?",
                    "Confirm Delete",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    ThisAddIn.userdata.EmailCategories.Remove(selectedCategory);
                    LoadEmailCategories();
                }
            }
            else
            {
                MessageBox.Show("Please select a category to delete.");
            }
        }

        private void checkedListBoxMailboxes_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            // Update MonitoredMailboxes list when checkboxes change
            // Note: Use BeginInvoke because ItemCheck fires before the check state changes
            this.BeginInvoke(new Action(() =>
            {
                if (ThisAddIn.userdata.MonitoredMailboxes == null)
                    ThisAddIn.userdata.MonitoredMailboxes = new List<string>();
                else
                    ThisAddIn.userdata.MonitoredMailboxes.Clear();

                foreach (var item in checkedListBoxMailboxes.CheckedItems)
                {
                    ThisAddIn.userdata.MonitoredMailboxes.Add(item.ToString());
                }
            }));
        }


        private void OK_Click(object sender, EventArgs e)
        {
                string json = JsonConvert.SerializeObject(userDataBindingSource.DataSource);
            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI", "OutlookAI.json"), json);

            // Invalidate HttpClient instances to apply new proxy/connection settings
            ThisAddIn.InvalidateHttpClients();

            // Restart email monitoring to apply new settings
            Globals.ThisAddIn.RestartEmailMonitoring();

            this.Close();
        }

        private async void Button2_Click(object sender, EventArgs e)
        {
            await GetOllamaModels();
        }


        public async Task<List<ModelInfo>> GetOllamaModels()
        {

            var ollamaUrl = ThisAddIn.userdata.OllamaUrl;
            if (!ThisAddIn.userdata.OllamaUrl.EndsWith("/"))
                ollamaUrl += "/";
            ollamaUrl += "api/tags";
            try
            {
                HttpClient httpClient = ThisAddIn.GetHttpClient();

                var response = await httpClient.GetAsync(ollamaUrl).ConfigureAwait(false);

                if (!response.IsSuccessStatusCode)
                {
                    throw new System.Exception($"{OutlookAI.Resources.ErrorcallingOllama}: {response.StatusCode}\n{await response.Content.ReadAsStringAsync().ConfigureAwait(false)}");
                }

                string jsonResponse = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
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
                throw new System.Exception($"{OutlookAI.Resources.ErrorcallingOllama}: {ex.Message}");
            }
        }

        private void ListBox1_Click(object sender, EventArgs e)
        {
            textBox2.Text = listBox1.SelectedItem.ToString();
        }

        private void label27_Click(object sender, EventArgs e)
        {

        }

        private void PromptBox_Load(object sender, EventArgs e)
        {
            InitializeEmailMonitoringTab();

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
