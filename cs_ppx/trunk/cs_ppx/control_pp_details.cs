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
using System.Collections.Generic;
using System.Diagnostics;

namespace cs_ppx
{
    [ComVisible(true)]
    [ProgId("cs_ppx.TaskPane_PP_Details")]

    public partial class control_pp_details : UserControl
    {
        
        public ISldWorks SwApp;
        public Component2[] CompName;
        
        
        public control_pp_details()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SwApp.SendMsgToUser("Success");
        }

        public void getSwApp(ref ISldWorks SwAppIn)
        {
            SwApp = SwAppIn;
        }

        public void getCompName(ref Component2[] CompNameIn)
        {
            CompName = CompNameIn;
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {
            
        }

        public void RegisterToPlaneTable(List<AddedReferencePlane> SelectedPlanes)
        {

            this.MP_details.Columns[1].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            this.MP_details.Columns[2].DefaultCellStyle.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;

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

        //traverse component and get the components name
        private void GetCompName(Component2 components, ref Component2[] compName)
        {
            object[] childComp = (object[])components.GetChildren();
            int i = 0;

            Component2[] Components = new Component2[2];

            //get the names
            while (i < childComp.Length)
            {
                Components[i] = (Component2)childComp[i];
                i++;
            }

            if (Components[0].Name2.Contains("rm"))
            {
                compName[0] = Components[0];
                compName[1] = Components[1];
            }
            else
            {
                compName[0] = Components[1];
                compName[1] = Components[0];
            }

        }

    }
}
