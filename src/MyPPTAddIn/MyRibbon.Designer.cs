namespace MyPPTAddIn
{
    partial class MyRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public MyRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.tabDZ = this.Factory.CreateRibbonTab();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.colorDialog = new System.Windows.Forms.ColorDialog();
            this.btnColorPickup = this.Factory.CreateRibbonButton();
            this.btnColorPanel = this.Factory.CreateRibbonButton();
            this.btnAddFooter = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button11 = this.Factory.CreateRibbonButton();
            this.button12 = this.Factory.CreateRibbonButton();
            this.gallery1 = this.Factory.CreateRibbonGallery();
            this.tab1.SuspendLayout();
            this.tabDZ.SuspendLayout();
            this.group3.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabHome";
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabHome";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // tabDZ
            // 
            this.tabDZ.Groups.Add(this.group3);
            this.tabDZ.Groups.Add(this.group2);
            this.tabDZ.Groups.Add(this.group4);
            this.tabDZ.Label = "大众";
            this.tabDZ.Name = "tabDZ";
            this.tabDZ.Position = this.Factory.RibbonPosition.AfterOfficeId("TabHome");
            // 
            // group3
            // 
            this.group3.Items.Add(this.btnColorPickup);
            this.group3.Items.Add(this.btnColorPanel);
            this.group3.Label = "颜色";
            this.group3.Name = "group3";
            // 
            // group2
            // 
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button6);
            this.group2.Items.Add(this.button7);
            this.group2.Items.Add(this.button8);
            this.group2.Items.Add(this.button9);
            this.group2.Items.Add(this.btnAddFooter);
            this.group2.Label = "标签";
            this.group2.Name = "group2";
            // 
            // group4
            // 
            this.group4.Items.Add(this.button10);
            this.group4.Items.Add(this.button11);
            this.group4.Items.Add(this.button12);
            this.group4.Items.Add(this.gallery1);
            this.group4.Label = "插入";
            this.group4.Name = "group4";
            // 
            // btnColorPickup
            // 
            this.btnColorPickup.Image = global::MyPPTAddIn.Properties.Resources.Color_Dropper_01;
            this.btnColorPickup.Label = "拾色器";
            this.btnColorPickup.Name = "btnColorPickup";
            this.btnColorPickup.ShowImage = true;
            this.btnColorPickup.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnColorPickup_Click);
            // 
            // btnColorPanel
            // 
            this.btnColorPanel.Image = global::MyPPTAddIn.Properties.Resources.color_scheme_01;
            this.btnColorPanel.Label = "颜色面板";
            this.btnColorPanel.Name = "btnColorPanel";
            this.btnColorPanel.ShowImage = true;
            this.btnColorPanel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnColorPanel_Click);
            // 
            // btnAddFooter
            // 
            this.btnAddFooter.Image = global::MyPPTAddIn.Properties.Resources.Text_Footer_01;
            this.btnAddFooter.Label = "添加页脚";
            this.btnAddFooter.Name = "btnAddFooter";
            this.btnAddFooter.ShowImage = true;
            this.btnAddFooter.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddFooter_Click);
            // 
            // button4
            // 
            this.button4.Image = global::MyPPTAddIn.Properties.Resources.Comments_01;
            this.button4.Label = "添加Comment";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            // 
            // button5
            // 
            this.button5.Image = global::MyPPTAddIn.Properties.Resources.next_01;
            this.button5.Label = "下一个Comment";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            // 
            // button6
            // 
            this.button6.Image = global::MyPPTAddIn.Properties.Resources.delete_text_message_01;
            this.button6.Label = "删除所有Comments";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button7
            // 
            this.button7.Image = global::MyPPTAddIn.Properties.Resources.add_01;
            this.button7.Label = "添加\"Draft\"标签";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            // 
            // button8
            // 
            this.button8.Image = global::MyPPTAddIn.Properties.Resources.add_02;
            this.button8.Label = "添加\"Backup\"标签";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // button9
            // 
            this.button9.Image = global::MyPPTAddIn.Properties.Resources.add_03;
            this.button9.Label = "添加主题标签";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // button10
            // 
            this.button10.Image = global::MyPPTAddIn.Properties.Resources.table_01;
            this.button10.Label = "分栏文本";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // button11
            // 
            this.button11.Image = global::MyPPTAddIn.Properties.Resources.text_01;
            this.button11.Label = "本页结论框";
            this.button11.Name = "button11";
            this.button11.ShowImage = true;
            // 
            // button12
            // 
            this.button12.Image = global::MyPPTAddIn.Properties.Resources.comments_02;
            this.button12.Label = "备注";
            this.button12.Name = "button12";
            this.button12.ShowImage = true;
            // 
            // gallery1
            // 
            this.gallery1.Image = global::MyPPTAddIn.Properties.Resources.right_01;
            this.gallery1.ItemImageSize = new System.Drawing.Size(18, 12);
            ribbonDropDownItemImpl1.Image = global::MyPPTAddIn.Properties.Resources.arrow_right_01;
            ribbonDropDownItemImpl2.Image = global::MyPPTAddIn.Properties.Resources.arrow_triangle_02;
            this.gallery1.Items.Add(ribbonDropDownItemImpl1);
            this.gallery1.Items.Add(ribbonDropDownItemImpl2);
            this.gallery1.Label = "箭头";
            this.gallery1.Name = "gallery1";
            this.gallery1.ShowImage = true;
            this.gallery1.ShowItemLabel = false;
            // 
            // MyRibbon
            // 
            this.Name = "MyRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabDZ);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.MyRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabDZ.ResumeLayout(false);
            this.tabDZ.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        public Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabDZ;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddFooter;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnColorPickup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnColorPanel;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button12;
        internal Microsoft.Office.Tools.Ribbon.RibbonGallery gallery1;
        private System.Windows.Forms.ColorDialog colorDialog;
    }

    partial class ThisRibbonCollection
    {
        internal MyRibbon MyRibbon
        {
            get { return this.GetRibbon<MyRibbon>(); }
        }
    }
}
