using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Diagnostics;
using HiddenCbTreeView;

namespace cs_ppx
{
    [ComVisible(true)]
    [ProgId("TaskPane_PP_Details")]

    public partial class control_pp_details : UserControl
    {
        
        public ISldWorks SwApp;
        public Component2[] CompName;
        
        public control_pp_details()
        {
            InitializeComponent();
        }

        public void getSwApp(ref ISldWorks iSwApp)
        {   
            SwApp = iSwApp;
        }

        public void getCompName(ref Component2[] CompNameIn)
        {
            CompName = CompNameIn;
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            
        }

        public void RegisterToTree(List<MachiningPlan> MachiningPlanCollection)
        {
            this.MachiningTree.CheckBoxes = true;
            this.MachiningTree.Nodes.Clear();

            int index = 0;

            foreach (MachiningPlan Plan in MachiningPlanCollection)
            {
                index++;

                TreeNode PlanDetails = new TreeNode(index.ToString());

                HiddenCheckBoxTreeNode Time = new HiddenCheckBoxTreeNode("Time: " + Plan.MachiningTime.ToString());
                PlanDetails.Nodes.Add(Time);

                HiddenCheckBoxTreeNode Cost = new HiddenCheckBoxTreeNode("Cost: " + Plan.MachiningCost.ToString());
                PlanDetails.Nodes.Add(Cost);

                HiddenCheckBoxTreeNode Setups = new HiddenCheckBoxTreeNode("Setups: " + Plan.NumberOfSetups.ToString());
                PlanDetails.Nodes.Add(Setups);

                HiddenCheckBoxTreeNode TAD = new HiddenCheckBoxTreeNode("TADs: " + Plan.NumberOfTADchanges.ToString());
                PlanDetails.Nodes.Add(TAD);

                HiddenCheckBoxTreeNode Tool = new HiddenCheckBoxTreeNode("Tools: " + Plan.NumberOfTool.ToString());
                PlanDetails.Nodes.Add(Tool);

                HiddenCheckBoxTreeNode ProcessPlan = new HiddenCheckBoxTreeNode("Sequence");

                foreach (MachiningProcess MP in Plan.MachiningProceses)
                {
                    
                    HiddenCheckBoxTreeNode ProcessDetails = new HiddenCheckBoxTreeNode(MP.MachiningReference.name);

                    HiddenCheckBoxTreeNode CuttingTool = new HiddenCheckBoxTreeNode(MP.cuttingTool);
                    ProcessDetails.Nodes.Add(CuttingTool);

                    HiddenCheckBoxTreeNode ToolPath = new HiddenCheckBoxTreeNode(MP.toolPath);
                    ProcessDetails.Nodes.Add(ToolPath);

                    //TreeNode TAD = new TreeNode();
                    //TAD.Nodes.Add(MP.SelectedTAD.X.ToString());
                    //TAD.Nodes.Add(MP.SelectedTAD.Y.ToString());
                    //TAD.Nodes.Add(MP.SelectedTAD.Z.ToString());

                    //ProcessDetails.Nodes.Add(TAD);
                    //ProcessDetails.Nodes.Add(MP.TRV.Volume.ToString());
                    //ProcessDetails.Nodes.Add(MP.VisibilityCone.ToString());

                    ProcessPlan.Nodes.Add(ProcessDetails);
                }

                PlanDetails.Nodes.Add(ProcessPlan);

                this.MachiningTree.Nodes.Add(PlanDetails);
            }
        }

        public void RegisterToPlaneTable(List<AddedReferencePlane> SelectedPlanes)
        {

            this.MP_details.Columns[1].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.MP_details.Columns[2].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.MP_details.Columns[3].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;

            foreach (AddedReferencePlane TmpPlane in SelectedPlanes)
            {
                this.MP_details.Rows.Add(TmpPlane.name, TmpPlane.DistanceFromCentroid, TmpPlane.Score, TmpPlane.Remark, TmpPlane);
            }
        }

        private void MP_details_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            AddedReferencePlane TmpPlane;
            object value = MP_details.Rows[e.RowIndex].Cells[4].Value;
            if (value is DBNull) { return; }

            TmpPlane = (AddedReferencePlane)value;

            ShowThePlane(TmpPlane);

        }

        private void ShowThePlane(AddedReferencePlane RefPlane)
        {
            String Name = RefPlane.name;
            
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            AssemblyDoc AssyDoc = (AssemblyDoc)Doc;

            if (CompName == null) { return; }

            Name = Name + "@" + CompName[0].Name2 + "@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());

            Doc.Extension.SelectByID2(Name, "PLANE", 0, 0, 0, false, 0, null, 0);

        }

        List<String> CheckedNodes = new List<String>();

        private void treeView1_AfterCheck(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Checked)
            {
                CheckedNodes.Add(e.Node.FullPath.ToString());
            }
            else
            {
                CheckedNodes.Remove(e.Node.FullPath.ToString());
            }
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if (CheckedNodes.Count > 0)
            {
                foreach (string NodeIndex in CheckedNodes)
                {
                    SwAddin.GenerateMP(Convert.ToInt32(NodeIndex) - 1);
                }
            }
        }

        private void MachiningTree_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            //SwApp.SendMsgToUser("I got this when the node is click and the value is " + e.Node.FullPath.ToString());

            if (e.Node.FullPath.ToString().Count() == 1)
            {
                SwAddin.ShowTheMP(Convert.ToInt32(e.Node.FullPath.ToString()) - 1);
            }
        }

    }
}
