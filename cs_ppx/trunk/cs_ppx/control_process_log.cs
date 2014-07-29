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
    [ProgId("cs_ppx.TaskPane_Process_Log")]

    public partial class control_process_log : UserControl
    {
        public control_process_log()
        {
            InitializeComponent();
        }

        private void control_process_log_Load(object sender, EventArgs e)
        {

        }
    }
}
