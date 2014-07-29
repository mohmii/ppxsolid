using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
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

        private void UserControl1_Load(object sender, EventArgs e)
        {
        
        }

    }
}
