namespace cs_ppx
{
    partial class control_ref_guide
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.RefInfo = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.RefreshInfo = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.ListOfGT = new System.Windows.Forms.ComboBox();
            this.ListOfDTR = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.IsGT = new System.Windows.Forms.RadioButton();
            this.IsDatum = new System.Windows.Forms.RadioButton();
            this.AddTol = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // RefInfo
            // 
            this.RefInfo.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.RefInfo.Location = new System.Drawing.Point(6, 19);
            this.RefInfo.Multiline = true;
            this.RefInfo.Name = "RefInfo";
            this.RefInfo.Size = new System.Drawing.Size(303, 300);
            this.RefInfo.TabIndex = 0;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.Controls.Add(this.RefreshInfo);
            this.groupBox1.Controls.Add(this.RefInfo);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(315, 368);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Reference Information";
            // 
            // RefreshInfo
            // 
            this.RefreshInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RefreshInfo.Location = new System.Drawing.Point(114, 325);
            this.RefreshInfo.Name = "RefreshInfo";
            this.RefreshInfo.Size = new System.Drawing.Size(76, 37);
            this.RefreshInfo.TabIndex = 1;
            this.RefreshInfo.Text = "Refresh";
            this.RefreshInfo.UseVisualStyleBackColor = true;
            this.RefreshInfo.Click += new System.EventHandler(this.RefreshInfo_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox2.Controls.Add(this.ListOfGT);
            this.groupBox2.Controls.Add(this.ListOfDTR);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.IsGT);
            this.groupBox2.Controls.Add(this.IsDatum);
            this.groupBox2.Controls.Add(this.AddTol);
            this.groupBox2.Location = new System.Drawing.Point(3, 377);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(315, 138);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Reference Setup";
            // 
            // ListOfGT
            // 
            this.ListOfGT.FormattingEnabled = true;
            this.ListOfGT.Items.AddRange(new object[] {
            "straightness",
            "flatness",
            "circularity",
            "cylindricity",
            "parallelism",
            "angularity",
            "perpendicularity",
            "location",
            "concentricity",
            "runout",
            "total runout",
            "line profile",
            "surface profile",
            "symmetry"});
            this.ListOfGT.Location = new System.Drawing.Point(78, 50);
            this.ListOfGT.Name = "ListOfGT";
            this.ListOfGT.Size = new System.Drawing.Size(69, 21);
            this.ListOfGT.TabIndex = 10;
            // 
            // ListOfDTR
            // 
            this.ListOfDTR.FormattingEnabled = true;
            this.ListOfDTR.Location = new System.Drawing.Point(192, 50);
            this.ListOfDTR.Name = "ListOfDTR";
            this.ListOfDTR.Size = new System.Drawing.Size(72, 21);
            this.ListOfDTR.TabIndex = 9;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(156, 55);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(30, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "DTR";
            // 
            // IsGT
            // 
            this.IsGT.AutoSize = true;
            this.IsGT.Location = new System.Drawing.Point(32, 51);
            this.IsGT.Name = "IsGT";
            this.IsGT.Size = new System.Drawing.Size(40, 17);
            this.IsGT.TabIndex = 6;
            this.IsGT.TabStop = true;
            this.IsGT.Text = "GT";
            this.IsGT.UseVisualStyleBackColor = true;
            // 
            // IsDatum
            // 
            this.IsDatum.AutoSize = true;
            this.IsDatum.Location = new System.Drawing.Point(32, 28);
            this.IsDatum.Name = "IsDatum";
            this.IsDatum.Size = new System.Drawing.Size(56, 17);
            this.IsDatum.TabIndex = 5;
            this.IsDatum.TabStop = true;
            this.IsDatum.Text = "Datum";
            this.IsDatum.UseVisualStyleBackColor = true;
            // 
            // AddTol
            // 
            this.AddTol.Location = new System.Drawing.Point(114, 93);
            this.AddTol.Name = "AddTol";
            this.AddTol.Size = new System.Drawing.Size(75, 35);
            this.AddTol.TabIndex = 4;
            this.AddTol.Text = "Add";
            this.AddTol.UseVisualStyleBackColor = true;
            this.AddTol.Click += new System.EventHandler(this.AddTol_Click);
            // 
            // control_ref_guide
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "control_ref_guide";
            this.Size = new System.Drawing.Size(321, 518);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TextBox RefInfo;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button RefreshInfo;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox ListOfDTR;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.RadioButton IsGT;
        private System.Windows.Forms.RadioButton IsDatum;
        private System.Windows.Forms.Button AddTol;
        private System.Windows.Forms.ComboBox ListOfGT;
    }
}
