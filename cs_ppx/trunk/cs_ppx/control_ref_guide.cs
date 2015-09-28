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
    [ProgId("TaskPane_Ref_Guide")]

    public partial class control_ref_guide : UserControl
    {
        public ISldWorks SwApp;

        public control_ref_guide()
        {
            InitializeComponent();
        }

        public void getSwApp(ref ISldWorks iSwApp)
        {
            SwApp = iSwApp;
        }

        private void RefreshInfo_Click(object sender, EventArgs e)
        {
            SwApp.SendMsgToUser("the reference info is successfully refreshed");

            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SelectedObject = SwSelMgr.GetSelectedObject5(1);
            Entity SwEntity = (Entity)SelectedObject;

            //Create the definition
            //AttributeDef SwAttDef = SwApp.DefineAttribute("open_face");
            //Boolean RetVal = SwAttDef.AddParameter("status", (int)swParamType_e.swParamTypeDouble, 1, 0);

            AttributeDef SwAttDef = SwApp.DefineAttribute("added_ref_tol");
            //Boolean RetVal = SwAttDef.AddParameter("bb", (int)swParamType_e.swParamTypeDouble, 1, 0);

            Boolean RetVal = SwAttDef.Register();

            SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);

            int i = 0;
            while (SwAttribute == null && i < 300)
            {
                SwAttribute = SwEntity.FindAttribute(SwAttDef, i);
                i++;
            }

            if (SwAttribute == null)
            {
                SwApp.SendMsgToUser("attribute was not found");
            }
            else
            {
                //Parameter OpenParam = SwAttribute.GetParameter("status");
                Parameter OpenParam = SwAttribute.GetParameter("bb");

                String Value = OpenParam.GetDoubleValue().ToString();

                if (Value.Equals("1"))
                {
                    SwApp.SendMsgToUser("this is a open face");
                }
                else
                {
                    SwApp.SendMsgToUser("this is not a open face");
                }

            }

        }

        private int NewTolCounter;

        //set the tolerance to the coressponding entity
        private void AddTol_Click(object sender, EventArgs e)
        {
            //SwApp.SendMsgToUser("the new reference is successfully added");

            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SelectedObject = SwSelMgr.GetSelectedObject5(1);
            Entity SwEntity = (Entity)SelectedObject;

            //create the definition
            AttributeDef SwAttDef = SwApp.DefineAttribute("added_ref_tol");
            string AddedParamType = GetParamType();
            string AddedParamValue1 = GetParamValue1();
            string AddedParamValue2 = GetParamValue2();

            Boolean RetVal = SwAttDef.AddParameter(AddedParamType, (int)swParamType_e.swParamTypeString, 1, 0);
            RetVal = SwAttDef.AddParameter(AddedParamValue1, (int)swParamType_e.swParamTypeString, 1, 0);
            RetVal = SwAttDef.AddParameter(AddedParamValue2, (int)swParamType_e.swParamTypeString, 1, 0);

            RetVal = SwAttDef.Register();

            SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);
            NewTolCounter++;

            SwAttribute = SwAttDef.CreateInstance5(DocumentModel, SwEntity, "TolNumber" + NewTolCounter.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);

            if (SwAttribute != null)
            {
                SwApp.SendMsgToUser("The attribute has been added to the selected entity");
            }

        }

        private string GetParamType()
        {
            if (this.IsDatum.Checked == true)
            {
                return "set_as_datum";
            }

            if (this.IsGT.Checked == true)
            {
                return "added_gt";
            }

            return "";
        }

        private string GetParamValue1()
        {
            if (this.ListOfGT.SelectedItem != "")
            {
                return this.ListOfGT.SelectedItem.ToString();
            }

            return "";
        }

        private string GetParamValue2()
        {
            if (this.ListOfDTR.SelectedItem != "")
            {
                return (this.ListOfDTR.SelectedItem.ToString());
            }

            return "";
        }

    }
}
