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
            UserData ud = new UserData();
        }
        UserData ud;
        public PromptBox(UserData ud)
        {
            this.ud = ud;
            InitializeComponent();
            P1.Text = ud.Prompt1;
            P2.Text = ud.Prompt2;
            P3.Text = ud.Prompt3;
            P4.Text = ud.Prompt4;
            textBoxApiKey.Text = ud.ApiKey;
        }



        private void button1_Click(object sender, EventArgs e)
        {

            ud.Prompt1 = P1.Text;
            ud.Prompt2 = P2.Text;
            ud.Prompt3 = P3.Text;
            ud.Prompt4 = P4.Text;
            ud.ApiKey = textBoxApiKey.Text;
            string json = JsonConvert.SerializeObject(ud);
            File.WriteAllText(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "OutlookAI.json"), json);
            this.Close();
        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }
    }
}
