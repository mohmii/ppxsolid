﻿namespace cs_ppx
{
    
    partial class control_pp_details
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
            this.label1 = new System.Windows.Forms.Label();
            this.MP_details = new System.Windows.Forms.DataGridView();
            this.PlaneName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Distance = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.PlaneScore = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.AddRemark = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ObjectPtr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.MachiningTree = new HiddenCbTreeView.MixedCheckBoxesTreeView();
            ((System.ComponentModel.ISupportInitialize)(this.MP_details)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Machining Plane Details";
            // 
            // MP_details
            // 
            this.MP_details.AllowUserToAddRows = false;
            this.MP_details.AllowUserToDeleteRows = false;
            this.MP_details.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MP_details.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.MP_details.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.PlaneName,
            this.Distance,
            this.PlaneScore,
            this.AddRemark,
            this.ObjectPtr});
            this.MP_details.Location = new System.Drawing.Point(3, 16);
            this.MP_details.Name = "MP_details";
            this.MP_details.ReadOnly = true;
            this.MP_details.RowHeadersVisible = false;
            this.MP_details.Size = new System.Drawing.Size(304, 212);
            this.MP_details.TabIndex = 2;
            this.MP_details.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.MP_details_CellContentClick);
            // 
            // PlaneName
            // 
            this.PlaneName.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.PlaneName.HeaderText = "Name";
            this.PlaneName.Name = "PlaneName";
            this.PlaneName.ReadOnly = true;
            this.PlaneName.ToolTipText = "The plane name";
            this.PlaneName.Width = 60;
            // 
            // Distance
            // 
            this.Distance.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Distance.HeaderText = "D (mm)";
            this.Distance.Name = "Distance";
            this.Distance.ReadOnly = true;
            this.Distance.ToolTipText = "The distance from the virtual centroid";
            this.Distance.Width = 65;
            // 
            // PlaneScore
            // 
            this.PlaneScore.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.PlaneScore.HeaderText = "Score";
            this.PlaneScore.Name = "PlaneScore";
            this.PlaneScore.ReadOnly = true;
            this.PlaneScore.ToolTipText = "The score based on the number of intersection with other planes";
            this.PlaneScore.Width = 60;
            // 
            // AddRemark
            // 
            this.AddRemark.HeaderText = "Remark";
            this.AddRemark.Name = "AddRemark";
            this.AddRemark.ReadOnly = true;
            this.AddRemark.ToolTipText = "Additional remark for plane characteristic";
            // 
            // ObjectPtr
            // 
            this.ObjectPtr.HeaderText = "Object";
            this.ObjectPtr.Name = "ObjectPtr";
            this.ObjectPtr.ReadOnly = true;
            this.ObjectPtr.Visible = false;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 231);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(115, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Machining Plan Details";
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button1.Location = new System.Drawing.Point(211, 537);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(96, 35);
            this.button1.TabIndex = 5;
            this.button1.Text = "Generate";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click_1);
            // 
            // MachiningTree
            // 
            this.MachiningTree.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MachiningTree.Location = new System.Drawing.Point(3, 247);
            this.MachiningTree.Name = "MachiningTree";
            this.MachiningTree.Size = new System.Drawing.Size(304, 284);
            this.MachiningTree.TabIndex = 6;
            this.MachiningTree.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView1_AfterCheck);
            this.MachiningTree.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.MachiningTree_NodeMouseClick);
            // 
            // control_pp_details
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.MachiningTree);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.MP_details);
            this.Controls.Add(this.label1);
            this.Name = "control_pp_details";
            this.Size = new System.Drawing.Size(310, 575);
            this.Load += new System.EventHandler(this.UserControl1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.MP_details)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView MP_details;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DataGridViewTextBoxColumn PlaneName;
        private System.Windows.Forms.DataGridViewTextBoxColumn Distance;
        private System.Windows.Forms.DataGridViewTextBoxColumn PlaneScore;
        private System.Windows.Forms.DataGridViewTextBoxColumn AddRemark;
        private System.Windows.Forms.DataGridViewTextBoxColumn ObjectPtr;
        private System.Windows.Forms.Button button1;
        private HiddenCbTreeView.MixedCheckBoxesTreeView MachiningTree;
    }
}
