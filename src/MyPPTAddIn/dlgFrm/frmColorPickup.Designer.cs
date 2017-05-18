namespace MyPPTAddIn.dlgFrm
{
    partial class frmColorPickup
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.timerSelectColor = new System.Windows.Forms.Timer(this.components);
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel4 = new System.Windows.Forms.Panel();
            this.txtHex = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtRGB = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPoint = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.panelColor = new System.Windows.Forms.Panel();
            this.lblDescribe = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panelImg = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.btnSave = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel4.SuspendLayout();
            this.panelColor.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panelImg.SuspendLayout();
            this.SuspendLayout();
            // 
            // timerSelectColor
            // 
            this.timerSelectColor.Interval = 300;
            this.timerSelectColor.Tick += new System.EventHandler(this.timerSelectColor_Tick);
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.panel4);
            this.panel1.Controls.Add(this.panelColor);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(375, 133);
            this.panel1.TabIndex = 0;
            // 
            // panel4
            // 
            this.panel4.Controls.Add(this.txtHex);
            this.panel4.Controls.Add(this.label3);
            this.panel4.Controls.Add(this.txtRGB);
            this.panel4.Controls.Add(this.label2);
            this.panel4.Controls.Add(this.txtPoint);
            this.panel4.Controls.Add(this.label1);
            this.panel4.Dock = System.Windows.Forms.DockStyle.Left;
            this.panel4.Location = new System.Drawing.Point(0, 0);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(155, 133);
            this.panel4.TabIndex = 0;
            // 
            // txtHex
            // 
            this.txtHex.Location = new System.Drawing.Point(41, 98);
            this.txtHex.Name = "txtHex";
            this.txtHex.ReadOnly = true;
            this.txtHex.Size = new System.Drawing.Size(100, 21);
            this.txtHex.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 102);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(23, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "HEX";
            // 
            // txtRGB
            // 
            this.txtRGB.Location = new System.Drawing.Point(41, 62);
            this.txtRGB.Name = "txtRGB";
            this.txtRGB.ReadOnly = true;
            this.txtRGB.Size = new System.Drawing.Size(100, 21);
            this.txtRGB.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 66);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(23, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "RGB";
            // 
            // txtPoint
            // 
            this.txtPoint.Location = new System.Drawing.Point(41, 26);
            this.txtPoint.Name = "txtPoint";
            this.txtPoint.ReadOnly = true;
            this.txtPoint.Size = new System.Drawing.Size(100, 21);
            this.txtPoint.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(23, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "X-Y";
            // 
            // panelColor
            // 
            this.panelColor.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelColor.Controls.Add(this.lblDescribe);
            this.panelColor.Location = new System.Drawing.Point(171, 20);
            this.panelColor.Name = "panelColor";
            this.panelColor.Size = new System.Drawing.Size(180, 96);
            this.panelColor.TabIndex = 1;
            // 
            // lblDescribe
            // 
            this.lblDescribe.AutoSize = true;
            this.lblDescribe.BackColor = System.Drawing.Color.Transparent;
            this.lblDescribe.ForeColor = System.Drawing.Color.Black;
            this.lblDescribe.Location = new System.Drawing.Point(49, 42);
            this.lblDescribe.Name = "lblDescribe";
            this.lblDescribe.Size = new System.Drawing.Size(83, 12);
            this.lblDescribe.TabIndex = 0;
            this.lblDescribe.Text = "按ESC选定颜色";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.panelImg);
            this.panel2.Location = new System.Drawing.Point(0, 133);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(375, 163);
            this.panel2.TabIndex = 1;
            // 
            // panelImg
            // 
            this.panelImg.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelImg.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom;
            this.panelImg.Controls.Add(this.label5);
            this.panelImg.Location = new System.Drawing.Point(37, 31);
            this.panelImg.Name = "panelImg";
            this.panelImg.Size = new System.Drawing.Size(300, 100);
            this.panelImg.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.BackColor = System.Drawing.Color.Transparent;
            this.label5.Font = new System.Drawing.Font("宋体", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.label5.Location = new System.Drawing.Point(140, 40);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(21, 21);
            this.label5.TabIndex = 0;
            this.label5.Text = "+";
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(150, 314);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "确定";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmColorPickup
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(375, 349);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmColorPickup";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "取色器";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmColorPickup_FormClosing);
            this.Load += new System.EventHandler(this.frmColorPickup_Load);
            this.panel1.ResumeLayout(false);
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.panelColor.ResumeLayout(false);
            this.panelColor.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panelImg.ResumeLayout(false);
            this.panelImg.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timerSelectColor;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.TextBox txtHex;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtRGB;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtPoint;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Panel panelColor;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panelImg;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblDescribe;
        private System.Windows.Forms.Button btnSave;
    }
}