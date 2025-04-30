using System.Windows.Forms;

namespace OutlookAI
{
    public partial class InputBox : Form
    {
        public InputBox()
        {
            InitializeComponent();
        }
        public string InputText { get; private set; }
        public InputBox(string prompt, string title):this()
        {
            this.Text = title;
            textBoxPrompt.Text = prompt;
        }
    }
}
