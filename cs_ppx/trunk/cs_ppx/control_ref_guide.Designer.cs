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
            this.AddTol = new System.Windows.Forms.Button();
            this.RefInput = new System.Windows.Forms.TextBox();
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
            this.RefInfo.Size = new System.Drawing.Size(303, 293);
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
            this.groupBox1.Size = new System.Drawing.Size(315, 361);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Reference Information";
            // 
            // RefreshInfo
            // 
            this.RefreshInfo.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.RefreshInfo.Location = new System.Drawing.Point(114, 318);
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
            this.groupBox2.Controls.Add(this.RefInput);
            this.groupBox2.Controls.Add(this.AddTol);
            this.groupBox2.Location = new System.Drawing.Point(3, 370);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(315, 326);
            this.groupBox2.TabIndex = 2;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Reference Setup";
            // 
            // AddTol
            // 
            this.AddTol.Location = new System.Drawing.Point(114, 285);
            this.AddTol.Name = "AddTol";
            this.AddTol.Size = new System.Drawing.Size(75, 35);
            this.AddTol.TabIndex = 4;
            this.AddTol.Text = "Add";
            this.AddTol.UseVisualStyleBackColor = true;
            this.AddTol.Click += new System.EventHandler(this.AddTol_Click);
            // 
            // RefInput
            // 
            this.RefInput.AcceptsReturn = true;
            this.RefInput.AcceptsTab = true;
            this.RefInput.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.RefInput.Location = new System.Drawing.Point(6, 19);
            this.RefInput.Multiline = true;
            this.RefInput.Name = "RefInput";
            this.RefInput.Size = new System.Drawing.Size(303, 260);
            this.RefInput.TabIndex = 5;
            // 
            // control_ref_guide
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "control_ref_guide";
            this.Size = new System.Drawing.Size(321, 699);
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
        private System.Windows.Forms.Button AddTol;
        private System.Windows.Forms.TextBox RefInput;
    }
}
