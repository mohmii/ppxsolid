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

        public void getSwApp(ISldWorks iSwApp)
        {
            SwApp = iSwApp;
        }

        private void RefreshInfo_Click(object sender, EventArgs e)
        {
            SwApp.SendMsgToUser("the reference info is successfully refreshed");
        }

        private void AddTol_Click(object sender, EventArgs e)
        {
            SwApp.SendMsgToUser("the new reference is successfully added");
        }


    }
}
