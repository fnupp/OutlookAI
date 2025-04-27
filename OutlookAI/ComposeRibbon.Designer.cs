namespace OutlookAI
{
    partial class ComposeRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Erforderliche Designervariable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ComposeRibbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnCompose1 = this.Factory.CreateRibbonButton();
            this.btnCompose2 = this.Factory.CreateRibbonButton();
            this.btnCompose3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnCompose1);
            this.group1.Items.Add(this.btnCompose2);
            this.group1.Items.Add(this.btnCompose3);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // btnCompose1
            // 
            this.btnCompose1.Label = "btnCompose1";
            this.btnCompose1.Name = "btnCompose1";
            this.btnCompose1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompose_Click);
            // 
            // btnCompose2
            // 
            this.btnCompose2.Label = "btnCompose2";
            this.btnCompose2.Name = "btnCompose2";
            this.btnCompose2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompose2_Click);
            // 
            // btnCompose3
            // 
            this.btnCompose3.Label = "btnCompose3";
            this.btnCompose3.Name = "btnCompose3";
            this.btnCompose3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCompose3_Click);
            // 
            // ComposeRibbon
            // 
            this.Name = "ComposeRibbon";
            this.RibbonType = "Microsoft.Outlook.Mail.Compose";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ComposeRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompose1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompose2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCompose3;
    }

    partial class ThisRibbonCollection
    {
        internal ComposeRibbon ComposeRibbon
        {
            get { return this.GetRibbon<ComposeRibbon>(); }
        }
    }
}
