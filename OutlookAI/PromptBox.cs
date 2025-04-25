using Newtonsoft.Json;
using System;
using System.IO;
using System.Windows.Forms;

namespace OutlookAI
{
    public partial class PromptBox : Form
    {
        public PromptBox()
        {
            InitializeComponent();
        }

        private readonly UserData userData;
        public PromptBox(UserData ud)
        {
            this.userData = ud;
            InitializeComponent();
            P1.Text = ud.Prompt1;
            P2.Text = ud.Prompt2;
            P3.Text = ud.Prompt3;
            P4.Text = ud.Prompt4;
            T1.Text = ud.Titel1;
            T2.Text = ud.Titel2;
            T3.Text = ud.Titel3;
            T4.Text = ud.Titel4;
            textBoxApiKey.Text = ud.ApiKey;
        }



        private void Button1_Click(object sender, EventArgs e)
        {

            userData.Prompt1 = P1.Text;
            userData.Prompt2 = P2.Text;
            userData.Prompt3 = P3.Text;
            userData.Prompt4 = P4.Text;
            userData.ApiKey = textBoxApiKey.Text;
            userData.Titel1 = T1.Text;
            userData.Titel2 = T2.Text;
            userData.Titel3 = T3.Text;
            userData.Titel4 = T4.Text;

            string json = JsonConvert.SerializeObject(userData);
            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            this.Close();
        }

    }
}
