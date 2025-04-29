namespace OutlookAI
{
    partial class OutlookAIRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public OutlookAIRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Verwendete Ressourcen bereinigen.
        /// </summary>
        /// <param name="disposing">"true", wenn verwaltete Ressourcen gelöscht werden sollen, andernfalls "false".</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Vom Komponenten-Designer generierter Code

        /// <summary>
        /// Erforderliche Methode für die Designerunterstützung.
        /// Der Inhalt der Methode darf nicht mit dem Code-Editor geändert werden.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDialogLauncher ribbonDialogLauncherImpl1 = this.Factory.CreateRibbonDialogLauncher();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.button1 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.separator2 = this.Factory.CreateRibbonSeparator();
            this.btnSummary1 = this.Factory.CreateRibbonButton();
            this.btnSummary2 = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.button5 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button7 = this.Factory.CreateRibbonButton();
            this.Button_Export = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.DialogLauncher = ribbonDialogLauncherImpl1;
            this.group1.Items.Add(this.button1);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.separator2);
            this.group1.Items.Add(this.btnSummary1);
            this.group1.Items.Add(this.btnSummary2);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.button5);
            this.group1.Label = "OutlookAI";
            this.group1.Name = "group1";
            this.group1.DialogLauncherClick += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Group1_DialogLauncherClick);
            // 
            // button1
            // 
            this.button1.Label = "3 Antworten kurz";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button1_Click);
            // 
            // button2
            // 
            this.button2.Label = "3 Antworten lang";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "button3";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button3_Click);
            // 
            // button4
            // 
            this.button4.Label = "Input";
            this.button4.Name = "button4";
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button4_Click);
            // 
            // separator2
            // 
            this.separator2.Name = "separator2";
            // 
            // btnSummary1
            // 
            this.btnSummary1.Label = "Zusammenfassung 1";
            this.btnSummary1.Name = "btnSummary1";
            this.btnSummary1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Summary_Click);
            // 
            // btnSummary2
            // 
            this.btnSummary2.Label = "Zusammenfassung";
            this.btnSummary2.Name = "btnSummary2";
            this.btnSummary2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BtnSummary2_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // button5
            // 
            this.button5.Image = global::OutlookAI.Properties.Resources._211751_gear_icon_1_;
            this.button5.Label = "Settings";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Button5_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.Button_Export);
            this.group2.Items.Add(this.button7);
            this.group2.Label = "Calendar Sync";
            this.group2.Name = "group2";
            // 
            // button7
            // 
            this.button7.Label = "Import Transfer File";
            this.button7.Name = "button7";
            this.button7.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.Import_Click);
            // 
            // Button_Export
            // 
            this.Button_Export.Label = "Export Transfer File";
            this.Button_Export.Name = "Button_Export";
            this.Button_Export.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ExportSync_Click);
            // 
            // button6
            // 
            this.button6.Label = "Zusammenfassung 1";
            this.button6.Name = "button6";
            // 
            // OutlookAIRibbon
            // 
            this.Name = "OutlookAIRibbon";
            this.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Read";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSummary1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSummary2;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton Button_Export;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
    }

    partial class ThisRibbonCollection
    {
        internal OutlookAIRibbon Ribbon1
        {
            get { return this.GetRibbon<OutlookAIRibbon>(); }
        }
    }
}
