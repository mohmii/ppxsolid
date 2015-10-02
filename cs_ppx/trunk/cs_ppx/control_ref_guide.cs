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
            //SwApp.SendMsgToUser("the reference info is successfully refreshed");

            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SwObj = (Object) SwSelMgr.GetSelectedObject6(1, -1);
            Entity SwEntity = (Entity) SwObj;
            Boolean RetVal = false;

            //Create the definition            
            AttributeDef SwAttDef = SwApp.DefineAttribute("ppx_tolerance" + SwEntity.ModelName);
            RetVal = SwAttDef.AddParameter("datum", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("flatness", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("parallelism", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("angularity", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("perpendicularity", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("location", (int)swParamType_e.swParamTypeString, 0.0, 0);

            Parameter TolParam = null;
            string TolValue = "";
            string[] RefInfo = new string[6]; 
            
            SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);
                        
            if (SwEntity != null)
            {
                int i = 0;
                while (SwAttribute == null && i < 3000)
                {
                    SwAttribute = SwEntity.FindAttribute(SwAttDef, i);
                    i++;
                }
            }

            if (SwAttribute == null)
            {
                SwApp.SendMsgToUser("attribute was not found");

            }
            else
            {
                TolParam = SwAttribute.GetParameter("datum");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("datum " + TolValue + System.Environment.NewLine); }
                
                TolParam = SwAttribute.GetParameter("flatness");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("flatness " + TolValue + System.Environment.NewLine); }

                TolParam = SwAttribute.GetParameter("parallelism");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("parallelism " + TolValue + System.Environment.NewLine); }

                TolParam = SwAttribute.GetParameter("angularity");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("angularity " + TolValue + System.Environment.NewLine); }

                TolParam = SwAttribute.GetParameter("perpendicularity");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("perpendicularity " + TolValue + System.Environment.NewLine); }

                TolParam = SwAttribute.GetParameter("location");
                if (TolParam != null) { TolValue = TolParam.GetStringValue(); this.RefInfo.AppendText("location " + TolValue + System.Environment.NewLine); }
                
            }

        }

        private int NewTolCounter;
        private int AttDefCounter;

        //set the tolerance to the coressponding entity
        private void AddTol_Click(object sender, EventArgs e)
        {
            //SwApp.SendMsgToUser("the new reference is successfully added");

            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SelectedObject = SwSelMgr.GetSelectedObject5(1);
            Entity SwEntity = (Entity)SelectedObject;

            //SwApp.SendMsgToUser(SwEntity.ModelName);
            List<string> TolName = new List<string>();
            List<string> TolValue = new List<string>();

            string[] ToleranceInput = this.RefInput.Lines;
            
            //filter the duplicate tolerance, if similar to the similar tolerance value's place
            foreach (string ThisTol in ToleranceInput)
            {
                string[] ReadTolerance = ThisTol.Split(' ');

                if (ReadTolerance.Count() == 2)
                {
                    Boolean SimilarTol = TolName.Contains(ReadTolerance.FirstOrDefault());

                    if (SimilarTol == true)
                    {
                        int TolIndex = TolName.IndexOf(ReadTolerance.First());
                        TolValue[TolIndex] = TolValue[TolIndex] + " " + ReadTolerance.Last();
                    }
                    else
                    {
                        TolName.Add(ReadTolerance.First());
                        TolValue.Add(ReadTolerance.Last());
                    }
                }
            
            }

            Boolean RetVal = false;
            SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);
            
            //create the definition
            AttributeDef SwAttDef = SwApp.DefineAttribute("ppx_tolerance" + SwEntity.ModelName);
            RetVal = SwAttDef.AddParameter("datum", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("flatness", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("parallelism", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("angularity", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("perpendicularity", (int)swParamType_e.swParamTypeString, 0.0, 0);
            RetVal = SwAttDef.AddParameter("location", (int)swParamType_e.swParamTypeString, 0.0, 0);

            RetVal = SwAttDef.Register();
            if (RetVal == true) { NewTolCounter++; }

            //SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);
            SwAttribute = SwAttDef.CreateInstance5(DocumentModel, SwEntity, "TolNumber" + NewTolCounter.ToString(), 0, (int)swInConfigurationOpts_e.swThisConfiguration);

            if (SwAttribute != null)
            { 
                SolidWorks.Interop.sldworks.Attribute SwAttributeCheck = default(SolidWorks.Interop.sldworks.Attribute);
                
                int i = 0;
                while (SwAttributeCheck == null && i < 3000)
                {
                    SwAttributeCheck = SwEntity.FindAttribute(SwAttDef, i);
                    i++;
                }
                
                int IndexTol = 0;
                foreach (string ThisTol in TolName)
                {   
                    Parameter ParamType = (Parameter)SwAttributeCheck.GetParameter(ThisTol);

                    if (ParamType != null)
                    {   
                        RetVal = ParamType.SetStringValue2(TolValue[IndexTol], (int)swInConfigurationOpts_e.swAllConfiguration, "");
                        string CurrentValue = ParamType.GetStringValue();
                    }

                    IndexTol++;
                }

            }

            if (SwAttribute != null)
            {
                SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);

                int i = 0;
                while (SwAttribute == null && i < 3000)
                {
                    SwAttribute = SwEntity.FindAttribute(SwAttDef, i);
                    i++;
                }

                if (SwAttribute != null)
                {
                    SwApp.SendMsgToUser("The attribute has been added to the selected entity");
                }
            }

        }

        //get the tolerance name
        private string _GetName(int OrderNum)
        {
            switch (OrderNum)
            { 
                case 0:
                    return "datum";
                case 1:    
                    return "flatness";
                case 2:
                    return "parallelism";
                case 3:
                    return "angularity";
                case 4:
                    return "perpendicularity";
                case 5:
                    return "location";
    
            }

            return "";

        }

        

    }
}
