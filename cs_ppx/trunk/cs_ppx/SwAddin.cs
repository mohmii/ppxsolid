using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Reflection;
using System.Windows.Media.Media3D;
using System.Windows;
using System.IO;
using System.Diagnostics;
using System.Linq;

using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;
using SolidWorks.Interop.sldcostingapi;
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;

namespace cs_ppx
{
    /// <summary>
    /// Summary description for cs_ppx.
    /// </summary>
    [Guid("8c7d635f-735d-49a0-9259-08b4a9d926f3"), ComVisible(true)]
    [SwAddin(
        Description = "this is a thesis project",
        Title = "cs_ppx",
        LoadAtStartup = true
        )]

    //public partial class SwAddin : ISwAddin
    public partial class SwAddin : ISwAddin
        {
        #region Local Variables
        static ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        //variable for PP Details taskpane
        TaskpaneView PPDetails_TaskPaneView;
        static control_pp_details PPDetails_TaskPaneHost;

        //variable for Process Log taskpane
        TaskpaneView ProcessLog_TaskPaneView;
        private static control_process_log ProcessLog_TaskPaneHost;

        //variable for Reference Guide
        TaskpaneView RefGuide_TaskPaneView;
        static control_ref_guide RefGuide_TaskPaneHost;

        public const int mainCmdGroupID = 20;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;

        //just added (for ribbon registration)
        public const int mainItemID3 = 2; 
        public const int mainItemID4 = 3;
        public const int mainItemID5 = 4;
        public const int mainItemID6 = 5;
        public const int mainItemID7 = 6;
        public const int mainItemID8 = 7;
        public const int mainItemID9 = 8;
        public const int mainItemID10 = 9;
        public const int mainItemID11 = 10;
        public const int mainItemID12 = 11;
        public const int mainItemID13 = 12;
        public const int mainItemID14 = 13;
        public const int mainItemID15 = 14;//
        public const int mainItemID16 = 15;
        public const int mainItemID17 = 16;
        
        public const int flyoutGroupID = 91;
                

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;
        #endregion


        // Public Properties
        public static ISldWorks SwApp
        {
            get { return iSwApp; }
        }
        public ICommandManager CmdMgr
        {
            get { return iCmdMgr; }
        }

        public Hashtable OpenDocs
        {
            get { return openDocs; }
        }

        #endregion

        #region SolidWorks Registration
        [ComRegisterFunctionAttribute]
        public static void RegisterFunction(Type t)
        {
            #region Get Custom Attribute: SwAddinAttribute
            SwAddinAttribute SWattr = null;
            Type type = typeof(SwAddin);

            foreach (System.Attribute attr in type.GetCustomAttributes(false))
            {
                if (attr is SwAddinAttribute)
                {
                    SWattr = attr as SwAddinAttribute;
                    break;
                }
            }

            #endregion

            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                Microsoft.Win32.RegistryKey addinkey = hklm.CreateSubKey(keyname);
                addinkey.SetValue(null, 0);

                addinkey.SetValue("Description", SWattr.Description);
                addinkey.SetValue("Title", SWattr.Title);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                addinkey = hkcu.CreateSubKey(keyname);
                addinkey.SetValue(null, Convert.ToInt32(SWattr.LoadAtStartup), Microsoft.Win32.RegistryValueKind.DWord);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem registering this dll: SWattr is null. \n\"" + nl.Message + "\"");
                System.Windows.Forms.MessageBox.Show("There was a problem registering this dll: SWattr is null.\n\"" + nl.Message + "\"");
            }

            catch (System.Exception e)
            {
                Console.WriteLine(e.Message);

                System.Windows.Forms.MessageBox.Show("There was a problem registering the function: \n\"" + e.Message + "\"");
            }
        }

        [ComUnregisterFunctionAttribute]
        public static void UnregisterFunction(Type t)
        {
            try
            {
                Microsoft.Win32.RegistryKey hklm = Microsoft.Win32.Registry.LocalMachine;
                Microsoft.Win32.RegistryKey hkcu = Microsoft.Win32.Registry.CurrentUser;

                string keyname = "SOFTWARE\\SolidWorks\\Addins\\{" + t.GUID.ToString() + "}";
                hklm.DeleteSubKey(keyname);

                keyname = "Software\\SolidWorks\\AddInsStartup\\{" + t.GUID.ToString() + "}";
                hkcu.DeleteSubKey(keyname);
            }
            catch (System.NullReferenceException nl)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + nl.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + nl.Message + "\"");
            }
            catch (System.Exception e)
            {
                Console.WriteLine("There was a problem unregistering this dll: " + e.Message);
                System.Windows.Forms.MessageBox.Show("There was a problem unregistering this dll: \n\"" + e.Message + "\"");
            }
        }

        #endregion

        #region ISwAddin Implementation
        public SwAddin()
        {
        }

        public bool ConnectToSW(object ThisSW, int cookie)
        {
            iSwApp = (ISldWorks)ThisSW;
            addinID = cookie;

            //set instance to keep all open models
            allModels = new List<ModelDoc2>();

            //Setup callbacks
            iSwApp.SetAddinCallbackInfo(0, this, addinID);

            #region Setup the Command Manager
            iCmdMgr = iSwApp.GetCommandManager(cookie);
            AddCommandMgr();
            #endregion

            #region Setup the Event Handlers
            SwEventPtr = (SolidWorks.Interop.sldworks.SldWorks)iSwApp;
            openDocs = new Hashtable();
            AttachEventHandlers();
            #endregion

            #region Setup Sample Property Manager
            AddPMP();
            #endregion

            AddTaskPane_PPDetails();
            AddTaskPane_ProcessLog();
            AddTaskPane_RefGuide();

            //iSwApp.SendMsgToUser("ppxsolid is loaded");

            return true;
        }

        public bool DisconnectFromSW()
        {
            RemoveTaskPane_PPDetails();
            RemoveTaskPane_ProcessLog();
            RemoveTaskPane_RefGuide();
            RemoveCommandMgr();
            RemovePMP();
            DetachEventHandlers();

            System.Runtime.InteropServices.Marshal.ReleaseComObject(iCmdMgr);
            iCmdMgr = null;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(iSwApp);
            iSwApp = null;
            //The addin _must_ call GC.Collect() here in order to retrieve all managed code pointers 
            GC.Collect();
            GC.WaitForPendingFinalizers();

            GC.Collect();
            GC.WaitForPendingFinalizers();

            //iSwApp.SendMsgToUser("ppxsolid is unloaded");

            return true;
        }
        #endregion

        #region UI Methods

        public void AddTaskPane_PPDetails()
        {
            PPDetails_TaskPaneView = SwApp.CreateTaskpaneView2("", "PP Details");
            PPDetails_TaskPaneHost = PPDetails_TaskPaneView.AddControl("TaskPane_PP_Details", "");
            PPDetails_TaskPaneHost.getSwApp(ref iSwApp);
            //PPDetails_TaskPaneHost.getCompName(ref compName);
        }

        public void RemoveTaskPane_PPDetails()
        {
            try
            {
                PPDetails_TaskPaneHost = null;
                PPDetails_TaskPaneView.DeleteView();
                Marshal.ReleaseComObject(PPDetails_TaskPaneView);
                PPDetails_TaskPaneView = null;
            }
            catch (Exception EX)
            {
                iSwApp.SendMsgToUser("Can't remove process plane task pane." + "\r\" Exception message :" + EX.Message);
            }
            
        }

        public void AddTaskPane_ProcessLog()
        {
            ProcessLog_TaskPaneView = SwApp.CreateTaskpaneView2("", "Process Log");
            ProcessLog_TaskPaneHost = ProcessLog_TaskPaneView.AddControl("cs_ppx.TaskPane_Process_Log", "");            
        }

        public void RemoveTaskPane_ProcessLog()
        {
            try
            {
                ProcessLog_TaskPaneHost = null;
                ProcessLog_TaskPaneView.DeleteView();
                Marshal.ReleaseComObject(ProcessLog_TaskPaneView);
                ProcessLog_TaskPaneView = null;
            }
            catch (Exception EX)
            {
                iSwApp.SendMsgToUser("Can't remove process log task pane." + "\r\" Exception message :" + EX.Message);
            }

        }

        public void AddTaskPane_RefGuide()
        {
            RefGuide_TaskPaneView = SwApp.CreateTaskpaneView2("", "Reference Information");
            RefGuide_TaskPaneHost = RefGuide_TaskPaneView.AddControl("TaskPane_Ref_Guide", "");
            RefGuide_TaskPaneHost.getSwApp(ref iSwApp);
        }

        public void RemoveTaskPane_RefGuide()
        {
            try
            {
                RefGuide_TaskPaneHost = null;
                RefGuide_TaskPaneView.DeleteView();
                Marshal.ReleaseComObject(RefGuide_TaskPaneView);
                RefGuide_TaskPaneView = null;
            }
            catch (Exception EX)
            {
                iSwApp.SendMsgToUser("Can't remove reference information task pane." + "\r\" Exception message :" + EX.Message);
            }

        }
        
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;

            //set the index for the button on the ribbon
            int cmdIndex0, cmdIndex1, cmdIndex2, cmdIndex3, cmdIndex4, cmdIndex5, cmdIndex6, cmdIndex7, cmdIndex8, cmdIndex9, cmdIndex10, cmdIndex11, cmdIndex12, cmdIndex13, cmdIndex14,cmdIndex15,cmdIndex16;
            string Title = "ppx", ToolTip = "Flexible Process Planning";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[16] { mainItemID1, mainItemID2, mainItemID3, mainItemID4, mainItemID5, mainItemID6, mainItemID7, mainItemID8, mainItemID9, mainItemID10, 
                mainItemID11, mainItemID12, mainItemID13, mainItemID14, mainItemID15 ,mainItemID16};

            if (getDataResult)
            {
                if (!CompareIDs((int[])registryIDs, knownIDs)) //if the IDs don't match, reset the commandGroup
                {
                    ignorePrevious = true;
                }
            }

            cmdGroup = iCmdMgr.CreateCommandGroup2(mainCmdGroupID, Title, ToolTip, "", -1, ignorePrevious, ref cmdGroupErr);
            cmdGroup.LargeIconList = iBmp.CreateFileFromResourceBitmap("cs_ppx.ToolbarLarge.bmp", thisAssembly);
            cmdGroup.SmallIconList = iBmp.CreateFileFromResourceBitmap("cs_ppx.ToolbarSmall.bmp", thisAssembly);
            cmdGroup.LargeMainIcon = iBmp.CreateFileFromResourceBitmap("cs_ppx.MainIconLarge.bmp", thisAssembly);
            cmdGroup.SmallMainIcon = iBmp.CreateFileFromResourceBitmap("cs_ppx.MainIconSmall.bmp", thisAssembly);

            //set the button properties here

            int menuToolbarOption = (int)(swCommandItemType_e.swMenuItem | swCommandItemType_e.swToolbarItem);
            //cmdIndex0 = cmdGroup.AddCommandItem2("Load Samples", -1, "Load all raw model and product samples", "Load Samples", 2, "LoadSamples", "", mainItemID1, menuToolbarOption);
            cmdIndex0 = cmdGroup.AddCommandItem2("Read Tolerance", -1, "Read tolerance from product", "Read Tolerance", 2, "ReadTolerance", "", mainItemID1, menuToolbarOption);
            //cmdIndex1 = cmdGroup.AddCommandItem2("Select Sample", -1, "Select sample available", "Select Sample", 2, "ShowPMP", "EnablePMP", mainItemID2, menuToolbarOption);
            //cmdIndex2 = cmdGroup.AddCommandItem2("Match Body", -1, "Matching material and product", "Match Body", 2, "MatchBody", "", mainItemID3, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2("Main TRV", -1, "Generate TRV from raw model and product", "Main TRV", 2, "MainTRV", "", mainItemID4, menuToolbarOption);
            cmdIndex4 = cmdGroup.AddCommandItem2("Plane Generator", -1, "Generate all planes", "Plane Generator", 2, "PlaneGenerator", "", mainItemID5, menuToolbarOption);
            cmdIndex5 = cmdGroup.AddCommandItem2("Plane Calculator", -1, "Calculating relationship matrix", "Plane Calculator", 2, "PlaneCalculator", "", mainItemID6, menuToolbarOption);
            cmdIndex6 = cmdGroup.AddCommandItem2("TRV Feature", -1, "Generate TRV features", "TRV Feature", 2, "TRVFeature", "", mainItemID7, menuToolbarOption);
            //cmdIndex7 = cmdGroup.AddCommandItem2("Machinable Space", -1, "Machinable space calculation", "Machinable Space", 2, "MachinableSpace", "", mainItemID8, menuToolbarOption);
            //cmdIndex8 = cmdGroup.AddCommandItem2("TRV network", -1, "TRV network calculation", "TRV Network", 2, "TRVNetwork", "", mainItemID9, menuToolbarOption);
            //cmdIndex9 = cmdGroup.AddCommandItem2("Calculate Setup", -1, "Setup calculation", "Calculate Setup", 2, "SetupCalculator", "", mainItemID10, menuToolbarOption);
            cmdIndex10 = cmdGroup.AddCommandItem2("Set Open Face", -1, "Defining a open face", "Set Open Face", 2, "SetOpenFace", "", mainItemID11, menuToolbarOption);
            cmdIndex11 = cmdGroup.AddCommandItem2("Read Open Face", -1, "Reading a open face", "Read Open Face", 2, "ReadOpenFace", "", mainItemID12, menuToolbarOption);
            cmdIndex12 = cmdGroup.AddCommandItem2("Analyze Open Faces", -1, "Analyzing open faces", "Analyze Open Faces", 2, "AnalyzeOpenFace", "", mainItemID13, menuToolbarOption);
            //cmdIndex13 = cmdGroup.AddCommandItem2("Cost Analysis", -1, "Analyzing the cost", "Analyze Cost", 2, "CostAnalysis", "", mainItemID14, menuToolbarOption);
            //cmdIndex14 = cmdGroup.AddCommandItem2("Test Bar", -1, "Testing the ribbon toolbar", "Test Bar", 2, "TestBar", "", mainItemID15, menuToolbarOption);
            

            //cmdIndex14 = cmdGroup.AddCommandItem2("Plane CalculatorZ", -1, "Check Topological Data", "Plane CalculatorZ", 2, "PlaneCalculatorZ", "", mainItemID15, menuToolbarOption);
            cmdIndex14 = cmdGroup.AddCommandItem2("Plane CalculatorZ", -1, "Check Topological Data", "Plane CalculatorZ", 2, "CheckLeveling", "", mainItemID15, menuToolbarOption);

            cmdIndex15 = cmdGroup.AddCommandItem2("CheckFeature", -1, "CheckFeature", "CheckFeature", 2, "CheckFeature", "", mainItemID16, menuToolbarOption);
            cmdIndex16 = cmdGroup.AddCommandItem2("Plane CalculatorXY", -1, "Check Topological Data", "Plane CalculatorXY", 2, "PlaneCalculatorXY", "", mainItemID17, menuToolbarOption);
            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
            
            bool bResult;
            
            //FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              //cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");


            //flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            //flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


            foreach (int type in docTypes)
            {
                CommandTab cmdTab;

                cmdTab = iCmdMgr.GetCommandTab(type, Title);

                if (cmdTab != null & !getDataResult | ignorePrevious)//if tab exists, but we have ignored the registry info (or changed command group ID), re-create the tab.  Otherwise the ids won't matchup and the tab will be blank
                {
                    bool res = iCmdMgr.RemoveCommandTab(cmdTab);
                    cmdTab = null;
                }

                //if cmdTab is null, must be first load (possibly after reset), add the commands to the tabs
                if (cmdTab == null)
                {
                    cmdTab = iCmdMgr.AddCommandTab(type, Title);

                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();

                    int[] cmdIDs = new int[7];
                    int[] TextType = new int[7];

                    //load samples button
                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex0);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    ////show pmp button
                    //cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);
                    //TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    ////match body button
                    //cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex2);
                    //TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    //main trv button
                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex3);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    //plane generator button
                    cmdIDs[4] = cmdGroup.get_CommandID(cmdIndex4);
                    TextType[4] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    //plane calculator button
                    cmdIDs[5] = cmdGroup.get_CommandID(cmdIndex5);
                    TextType[5] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    //TRV generator button
                    cmdIDs[6] = cmdGroup.get_CommandID(cmdIndex6);
                    TextType[6] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);

                    ////add another group
                    //CommandTabBox cmdBox2 = cmdTab.AddCommandTabBox();
                    //cmdIDs = new int[3];
                    //TextType = new int[3];

                    ////Machinable space calculation button
                    //cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex7);
                    //TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    ////TRV network calculation buttion
                    //cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex8);
                    //TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                                        
                    ////Setup calculation button
                    //cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex9);
                    //TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    //bResult = cmdBox2.AddCommands(cmdIDs, TextType);
                    //cmdTab.AddSeparator(cmdBox2, cmdGroup.ToolbarId);

                    //add another group
                    CommandTabBox cmdBox3 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[3];
                    TextType = new int[3];

                    //Set open face button
                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex10);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex11);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex12);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    bResult = cmdBox3.AddCommands(cmdIDs, TextType);
                    cmdTab.AddSeparator(cmdBox3, cmdGroup.ToolbarId);

                    //add another group
                    CommandTabBox cmdBox4 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[4];
                    TextType = new int[4];

                    ////Cost analysis
                    //cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex13);
                    //TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex14);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex15);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[3] = cmdGroup.get_CommandID(cmdIndex16);
                    TextType[3] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox4.AddCommands(cmdIDs, TextType);

                    //bResult = cmdBox4.AddCommands(cmdIDs, TextType);
                    //cmdTab.AddSeparator(cmdBox4, cmdGroup.ToolbarId);

                    /*
                    CommandTabBox cmdBox1 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[1];
                    TextType = new int[1];

                    cmdIDs[0] = flyGroup.CmdID;
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow | (int)swCommandTabButtonFlyoutStyle_e.swCommandTabButton_ActionFlyout;

                    bResult = cmdBox1.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox1, cmdIDs[0]);
                     */

                }

            }
            thisAssembly = null;

        }

        public void RemoveCommandMgr()
        {
            iBmp.Dispose();

            iCmdMgr.RemoveCommandGroup(mainCmdGroupID);
            iCmdMgr.RemoveFlyoutGroup(flyoutGroupID);
        }

        public bool CompareIDs(int[] storedIDs, int[] addinIDs)
        {
            List<int> storedList = new List<int>(storedIDs);
            List<int> addinList = new List<int>(addinIDs);

            addinList.Sort();
            storedList.Sort();

            if (addinList.Count != storedList.Count)
            {
                return false;
            }
            else
            {

                for (int i = 0; i < addinList.Count; i++)
                {
                    if (addinList[i] != storedList[i])
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public Boolean AddPMP()
        {
            ppage = new UserPMPage(this);
            return true;
        }

        public Boolean RemovePMP()
        {
            ppage = null;
            return true;
        }

        #endregion

        #region UI Callbacks
        public void CreateCube()
        {
            //make sure we have a part open
            string partTemplate = iSwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart);
            if ((partTemplate != null) && (partTemplate != ""))
            {
                IModelDoc2 modDoc = (IModelDoc2)iSwApp.NewDocument(partTemplate, (int)swDwgPaperSizes_e.swDwgPaperA2size, 0.0, 0.0);

                modDoc.InsertSketch2(true);
                modDoc.SketchRectangle(0, 0, 0, .1, .1, .1, false);
                //Extrude the sketch
                IFeatureManager featMan = modDoc.FeatureManager;
                featMan.FeatureExtrusion(true,
                    false, false,
                    (int)swEndConditions_e.swEndCondBlind, (int)swEndConditions_e.swEndCondBlind,
                    0.1, 0.0,
                    false, false,
                    false, false,
                    0.0, 0.0,
                    false, false,
                    false, false,
                    true,
                    false, false);
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("There is no part template available. Please check your options and make sure there is a part template selected, or select a new part template.");
            }
        }

        //load samples of raw material and product to solidworks
        public void LoadSamples()
        {             
            //open raw material type 100x100x100, 200x100x100, 200x150x100
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_100x100x100.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_120x70x65_mori.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);

            
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_200x100x100.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_200x150x100.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_50x25x30_3dparallelogram.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_51x28x22_diamond.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_55x40x28_roof.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_60x50x22_lem21.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_80x65x70_pyramid.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\rm_80x70x55_4sidepocket.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);

            //open the product
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_100x100x100_2.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_100x100x100_3.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_mori.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);


            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_3dparallelogram.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_4sidepocket.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_diamond.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_lem21.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_pyramid.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
            iSwApp.OpenDoc6("D:\\iis documents\\Research\\cad_solidworks\\illustrations\\sample 0\\prod_roof.sldprt",
                (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", (int)swFileLoadError_e.swGenericError,
                (int)swFileLoadWarning_e.swFileLoadWarning_AlreadyOpen);
                               
        }
        
        //match the raw material and prodyct in assembly model
        public void MatchBody()
        {

            //create the assembly template
            string assyTemplate = SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateAssembly);
            
            if ((assyTemplate != null) && (assyTemplate != ""))
            {

                int longStatus = 0;

                //create new asssembly document
                ModelDoc2 assyDoc = (ModelDoc2)SwApp.NewDocument(assyTemplate, 0, 0.0, 0.0);
                
                SwApp.ActivateDoc3(assyDoc.GetTitle(), false, 0, ref longStatus);
                assyDoc = (ModelDoc2)SwApp.ActiveDoc;

                AssemblyDoc assyModel = (AssemblyDoc)assyDoc;          
                
                if (assyModel != null)
                {
                    ModelDoc2[] availModel = getModels(rpList, modelList);
                    //AttributeDef attDef = null;

                    //markBBfaces(ref availModel[0], ref attDef);

                    addComponents(assyModel, availModel);

                    assyDoc.ViewZoomtofit2();

                }

            }                        
        }

        //additional class related to body matching
        #region MatchBody
        
        //contains user selection list for raw material and product
        public string[] rpList {get; set;}
        public AttributeDef BBDef = null;
        public AttributeDef NFDef = null;

        //keep all open models
        private List<ModelDoc2> allModels;
        public List<ModelDoc2> modelList
        {
            get { return allModels; }

            set
            { 
                foreach (ModelDoc2 path in value)
                {
                    if (allModels.Contains(path).Equals(false))
                    allModels.Add(path);
                }                 
            }
        }
        
        //add neccessary components into the assembly document
        public void addComponents(AssemblyDoc assyModel, ModelDoc2[] models)
        {
            string[] compNames = new string[1];
            double[] compXforms = new double[16];
            double[] arrayA = new double[16];
            double[] arrayB = new double[16];
            double[] arrayC = new double[16];
            string[] compCoordSysNames = new string[1];
            object vcompNames;
            object vcompXforms;
            object vcompCoordSysNames;
            object[] vcomponents = new object[1];
            
            string pathName = models[0].GetPathName();
            
                        
            // x-axis components of rotation
            compXforms[0] = 1.0;
            compXforms[1] = 0.0;
            compXforms[2] = 0.0;
            // y-axis components of rotation
            compXforms[3] = 0.0;
            compXforms[4] = 1.0;
            compXforms[5] = 0.0;
            // z-axis components of rotation
            compXforms[6] = 0.0;
            compXforms[7] = 0.0;
            compXforms[8] = 1.0;

            // Add a translation vector to the transform (zero translation) 
            compXforms[9] = 0.0;
            compXforms[10] = 0.0;
            compXforms[11] = 0.0;

            // Add a scaling factor to the transform
            compXforms[12] = 0.0;

            // The last three elements in the transformation matrix are unused

            // Register the coordinate system name for the component 
            compCoordSysNames[0] = "Coordinate System1";
            
            vcompXforms = compXforms;
                        
            //vcompXforms = Nothing   //also achieves zero rotation and translation of the component
            vcompCoordSysNames = compCoordSysNames;

            //make sure the assembly doc is activated
            assyModel = (AssemblyDoc)SwApp.ActiveDoc;

            
            //insert the raw
            compNames[0] = models[0].GetPathName();
            vcompNames = compNames;
            vcomponents = assyModel.AddComponents3((vcompNames), (vcompXforms), (vcompCoordSysNames));

            //insert the product
            compNames[0] = models[1].GetPathName();
            vcompNames = compNames;
            vcomponents = assyModel.AddComponents3((vcompNames), (vcompXforms), (vcompCoordSysNames));
            
        }

        //mark all the faces coincide with the bounding box
        private void markBBfaces(Component2 SwComp, ref AttributeDef attributDefinition)
        {
            bool bRet;
            Array BodyArray = null;
            Body2 TmpBody = null;
            Face2 SwFace = null;
            int Errors = 0;
            int Warnings = 0;
            int i = 1;

            String CompPathName = SwComp.GetPathName();
            iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisible
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

            //create the BB attribute definition
            AttributeDef attDef = SwApp.DefineAttribute("bb");
            attributDefinition = attDef;
            bRet = attDef.AddParameter("bb", (int)swParamType_e.swParamTypeDouble, 1, 0);
            bRet = attDef.Register();
            
            PartDoc swPartDoc = (PartDoc) CompDocumentModel;
            BodyArray = (Array)swPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);

            if (BodyArray.Length == 1)
            {
                
                TmpBody = (Body2) BodyArray.GetValue(0);

                SwFace = (Face2)TmpBody.GetFirstFace();

                while (SwFace !=null)
                {
                    
                    SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);
                    swAtt = attDef.CreateInstance5(CompDocumentModel, SwFace, "bb_face" + i.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);

                    SwFace = (Face2)SwFace.GetNextFace();
                    i++;
                }
            }

            iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be visble

        }

        //get the needed models
        static ModelDoc2[] getModels(string[] selections, List<ModelDoc2> allModels)
        {
            ModelDoc2[] tmpModels = new ModelDoc2[2];
            int index = 0;

            foreach (string sel in selections)
            {
                foreach (ModelDoc2 dbModel in allModels)
                {
                    if (dbModel.GetPathName().Contains(sel))
                    {
                        if (index < tmpModels.Length)
                            tmpModels.SetValue(dbModel, index);
                        index++;
                    }
                }
            }

            return tmpModels;
        }

        #endregion

        //raw material and product alignment process
        public void MainTRV()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            int docType = (int)Doc.GetType();
            bool boolStatus;            
            
            if (Doc.GetType() == 2)
            {
                AssemblyDoc assyModel = (AssemblyDoc)Doc;

                if (assyModel.GetComponentCount(false) != 0)
                {
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    //test was made
                    //check the initial existing component name
                    if (compName == null)
                    {
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);
                        PPDetails_TaskPaneHost.getCompName(ref compName);
                    }
                                        
                    boolStatus = compName[0].Select2(true, 0);

                    ModelDoc2 ModDoc2 = (ModelDoc2)compName[0].GetModelDoc2();
                    markBBfaces(compName[0], ref BBDef);

                    Body2 TmpBody = (Body2)compName[0].GetBody();
                    RawMaterialBody = (Body2)TmpBody.Copy();
                    RawMaterialFaces = (object[])RawMaterialBody.GetFaces();
                    
                    assyModel = (AssemblyDoc) Doc;

                    assyModel.EditPart();
                    
                    Doc.ClearSelection2(true);
                    boolStatus = compName[1].Select2(true, 0);
                    boolStatus = compName[0].IsFixed();
                    
                    assyModel = (AssemblyDoc)Doc;

                    assyModel.InsertCavity4(0, 0, 0, true, 1, -1);

                    assyModel.EditAssembly();

                    boolStatus = compName[0].Select2(true, 0);
                    Boolean retVal = false;
                    retVal = SetAttribute(compName[0], ref NFDef);

                    boolStatus = compName[1].Select2(true, 0);
                    boolStatus = assyModel.SetComponentTransparent(true);
                                        
                    //Doc.ClearSelection2(true);

                    //change the material color to red
                    changeColor(ref compName[0]);

                    Doc.GraphicsRedraw2();
                    Doc.ForceRebuild3(true);
                    
                    ProcessLog_TaskPaneHost.LogProcess("Generate main TRV");
                    PPDetails_TaskPaneHost.LogProcess("Generate main TRV");

                }
            }
            else
            {
                iSwApp.SendMsgToUser("There is no model document");
            }

        }

        //tools for maintrv
        # region MainTRV

        //set temporary body and face
        public static Body2 RawMaterialBody;
        public static object[] RawMaterialFaces;
        
        //traverse component and get the components name
        public static void GetCompName(Component2 components, ref Component2[] compName)
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

            if (Components[0].Name2.Contains("rawm_"))
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

        public static Component2[] compName;

        //change color of a the removal volume
        static bool changeColor(ref Component2 componentName)
        {
            if (componentName != null)
            {
                double[] variant = componentName.GetMaterialPropertyValues2((int)swInConfigurationOpts_e.swAllConfiguration, "");
                variant[0] = 1;
                variant[1] = 0;
                variant[2] = 0;
                variant[3] = 1;
                variant[4] = 1;
                variant[5] = 0.8;
                variant[6] = 0.3125;
                variant[7] = 0;
                variant[8] = 0;

                componentName.SetMaterialPropertyValues2(variant, (int)swInConfigurationOpts_e.swAllConfiguration, "");

                return true;
            }
            else
            {
                return false;
            }
        }

        //get random string
        public static string GetRandomString()
        {
            string path = Path.GetRandomFileName();
            path = path.Replace(".", "");
            return path;
        }

        //set an attribute to the the new created faces
        public bool SetAttribute(Component2 SwComp, ref AttributeDef NewFace)
        {
            Array BodyArray;
            Body2 TmpBody;
            Face2 SwFace;
            Entity SwEntity;
            int Errors = 0;
            int Warnings = 0;
            Boolean bRet = false;
            int faceIndex = 1;

            //create the BB attribute definition
            AttributeDef attDef = null;
            //AttributeDef attDef = SwApp.DefineAttribute("newface");
            //NewFace = attDef;
            //bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);
            //bRet = attDef.Register();

            String CompPathName = SwComp.GetPathName();
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

            BodyArray = null;
            TmpBody = null;

            PartDoc PartDocument = (PartDoc)CompDocumentModel;
            BodyArray = (Array)PartDocument.GetBodies2((int)swBodyType_e.swSolidBody, true);

            if (BodyArray.Length == 1)
            {
                TmpBody = (Body2)BodyArray.GetValue(0);

                SwFace = null;
                SwEntity = null;
                object[] AllFaces = (object[])TmpBody.GetFaces();

                //bRet = SetNewFaceAtt(CompDocumentModel, AllFaces, ref NewFace);

                SwFace = (Face2)TmpBody.GetFirstFace();
                while (SwFace != null)
                {
                    SwEntity = (Entity)SwFace;
                    SwEntity = SwEntity.GetSafeEntity();

                    //set the attribute definition
                    attDef = SwApp.DefineAttribute("newface" + SwEntity.ModelName);
                    NewFace = attDef;
                    bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);
                    bRet = attDef.Register();

                    SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);

                    if (bRet == true)
                    {
                        swAtt = attDef.CreateInstance5(CompDocumentModel, SwFace, TmpBody.Name + "_nf" + faceIndex.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);
                        faceIndex++;
                    }

                    SwFace = (Face2)SwFace.GetNextFace();
                }

            }
            else if (BodyArray.Length > 1)
            {
                for (int m = 0; m < BodyArray.Length; m++)
                {
                    TmpBody = (Body2)BodyArray.GetValue(m);

                    SwFace = null;
                    SwEntity = null;
                    object[] AllFaces = (object[])TmpBody.GetFaces();

                    //bRet = SetNewFaceAtt(CompDocumentModel, AllFaces, ref NewFace);

                    //SwFace = (Face2)TmpBody.GetFirstFace();
                    //while (SwFace != null)
                    foreach (Face2 ThisFace in AllFaces)
                    {
                        //SwEntity = (Entity)SwFace;
                        SwEntity = (Entity)ThisFace;
                        SwEntity = SwEntity.GetSafeEntity();

                        //set the attribute definition
                        attDef = SwApp.DefineAttribute("newface" + SwEntity.ModelName);
                        NewFace = attDef;
                        bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);
                        bRet = attDef.Register();

                        SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);

                        if (bRet == true)
                        {
                            swAtt = attDef.CreateInstance5(CompDocumentModel, ThisFace, TmpBody.Name + "_nf" + faceIndex.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);
                            faceIndex++;
                        }

                        //SwFace = (Face2)SwFace.GetNextFace();
                    }
                }

            }
           
            //iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            //iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be visble

            return true;
           
        }

        //Set the attribute "nf" on face that is newly created
        public bool SetNewFaceAtt(ModelDoc2 ThisPartDoc, object[] ThisObjFaces, ref AttributeDef NewFace)
        {
            Entity SwEntity = null;
            AttributeDef attDef = null;
            Boolean bRet = false;
            int faceIndex = 1;

            try
            {
                foreach (Face2 ThisFace in ThisObjFaces)
                {   
                    SwEntity = (Entity)ThisFace;
                    SwEntity = SwEntity.GetSafeEntity();

                    //set the attribute definition
                    attDef = SwApp.DefineAttribute("newface" + SwEntity.ModelName);
                    bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);
                    bRet = attDef.Register();

                    SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);

                    if (bRet == true)
                    {
                        NewFace = attDef;
                        swAtt = attDef.CreateInstance5(ThisPartDoc, ThisFace, "new_face" + faceIndex.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);
                        faceIndex++;
                    }
                }
            }
            catch (Exception Ex)
            {
                ProcessLog_TaskPaneHost.LogProcess("[ERROR] look up on SetNewFaceAtt");
                PPDetails_TaskPaneHost.LogProcess("[ERROR] look up on SetNewFaceAtt");
                
                return false;

            }

            return false;
        }

        //check number of Main TRV
        public static int GetNumberOfBodies()
        {



            return 0;
        }

        //analyzes the body's LWD (get the largest LW area and lowest D)
        public static int CheckSimpleLWD(Tessellation TessData)
        {
            
            return 0;
        }


        #endregion

        //generate the plane needed for TRV generation
        public void PlaneGenerator()
        {            
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            
            int docType = (int)Doc.GetType();
            bool boolStatus;

            if (Doc.GetType() == 2)
            {
                AssemblyDoc assyModel = (AssemblyDoc)Doc;
                
                if (assyModel.GetComponentCount(false) != 0)
                {
                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    if (compName == null)
                    {
                        //define the raw material and product
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);
                        PPDetails_TaskPaneHost.getCompName(ref compName);
                    }

                    //get faces from product
                    object bodyEnts = (object) compName[1].GetBody();
                    Body2 ProdBody = (Body2)bodyEnts;
                    object[] ProdFaces = (object[])ProdBody.GetFaces();

                    //get the part document of the raw material
                    ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                    PartDoc swPartDoc = (PartDoc)compModDoc;

                    SelectionMgr SwSelMgr = (SelectionMgr)compModDoc.SelectionManager;
                    SelectData SwSelData = (SelectData)SwSelMgr.CreateSelectData();

                    //prepare for face traversal
                    string entName = null;
                    string refName = null;
                    int countFace = 0;
                    RefPlane swRefPlane = null;
                    Entity swEnt = null;
                    Object TessFaceArray = null;
                    Double[] TessVerticesArray = null;
                    PatternConfig PatternState = null;
                    AttributeDef attDef = null;
                    Boolean bRet = false;
                    int Errors = 0;
                    int Warnings = 0;
                    bool PlanarState = false;
                    
                    AddedReferencePlane tmpInitPlane = null;
                    AddedReferenceProfile tmpInitProfile = null;
                    AddedReferenceSurface tmpInitSurface = null;
                    InitialRefPlanes = new List<AddedReferencePlane>(); //set the instance for the first time for saving all reference planes
                    InitialRefProfiles = new List<AddedReferenceProfile>(); //set the instance for the first time for saving all reference profiles
                    InitialRefSurfaces = new List<AddedReferenceSurface>(); //set the instance for the first time for saving all reference surface
                    InitialTRVBodies = new List<RemovalBody>(); //set the instance for the first time for saving all initial bodies
                    RemovalBody ThisInitialTRV = null;
                                     
                    //select the raw material for editing
                    boolStatus = compName[0].Select2(true, 0);

                    if (boolStatus == true)
                    {   
                        Array AllBodies = (Array)swPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                     
                        foreach (Body2 ThisBody in AllBodies)
                        {
                            //set the new instance for body and thisbody
                            ThisInitialTRV = new RemovalBody();
                            ThisInitialTRV.BodyObj = ThisBody;
                            ThisInitialTRV.Name = ThisBody.Name;

                            //calculate the virtual centroid
                            bRet = CalculateMidPoint(Doc, compName[0], ThisBody, ref ThisInitialTRV);

                            //calculate the TAD (listing all available TAD and the recommendation based on LWD
                            bRet = CalculateTAD(ThisBody, ref ThisInitialTRV, ProdFaces);

                            object[] AllFaces = ThisBody.GetFaces();
                                            
                            //foreach (Face2 tmpFace in swFaces)
                            foreach (Face2 tmpFace in AllFaces)
                            {
                                TessFaceArray = (Object)tmpFace.GetTessTriangles(true);
                                TessVerticesArray = GetMaxMin(TessFaceArray);
                                swEnt = (Entity)tmpFace;
                                swEnt = swEnt.GetSafeEntity();

                                entName = swPartDoc.GetEntityName(tmpFace);
                                PatternState = new PatternConfig();

                                //set the attribute definition
                                attDef = SwApp.DefineAttribute("newface" + swEnt.ModelName);
                                bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);

                                SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);

                                int n = 0; // arbitrary counter

                                while (swAtt == null && n < 300)
                                {
                                    swAtt = swEnt.FindAttribute(attDef, n);
                                    n++;
                                }

                                ////select the face without name
                                //if (entName == "")
                                if (swAtt != null)
                                {

                                    //set the reference plane instance
                                    tmpInitPlane = new AddedReferencePlane();
                                    PlanarState = IsPlanar(tmpFace);

                                    if (PlanarState == true)
                                    {   
                                        //bRet = swEnt.Select2(false, 0);
                                        bRet = swEnt.Select4(false, SwSelData);

                                        //create coincide reference plane to it
                                        //swRefPlane = (RefPlane) Doc.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0);
                                        swRefPlane = (RefPlane) compModDoc.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0);
                                        tmpInitPlane.IsPlanar = true;
                                    }
                                    else
                                    {
                                        //swRefPlane = SetupRefPlane(tmpFace, Doc, TessVerticesArray);
                                        swRefPlane = SetupRefPlane(tmpFace, Doc, compName[0], TessVerticesArray);
                                        tmpInitPlane.IsPlanar = false;
                                    }

                                    if (swRefPlane != null)
                                    {
                                        countFace += 1;

                                        //set reference name
                                        refName = "REF_PLANE" + countFace.ToString();
                                        swPartDoc.SetEntityName(swRefPlane, refName);

                                        //set the reference plane, included with name, coincided face, and its pointer
                                        tmpInitPlane.name = refName;
                                        tmpInitPlane.AttachedFace = tmpFace;
                                        tmpInitPlane.ReferencePlane = swRefPlane;
                                        tmpInitPlane.MaxMinValue = TessVerticesArray;
                                        tmpInitPlane.BodyOwner = ThisBody.Name;
                                        tmpInitPlane.AttachedBody = ThisInitialTRV;
                                        tmpInitPlane.SafeDirection = true;

                                        //set the surface if it is a freeform or cylinder surface
                                        if (PlanarState == false)
                                        {
                                            bRet = swEnt.Select4(false, SwSelData);
                                            compModDoc.InsertOffsetSurface(0.0, false);

                                            Feature ThisFeature = (Feature)compModDoc.Extension.GetLastFeatureAdded();
                                            object[] TmpFaces = (object[])ThisFeature.GetFaces();
                                            Face2 ThisFace; 

                                            if (TmpFaces.Count() == 1)
                                            {
                                                ThisFace = (Face2)TmpFaces.First();
                                                tmpInitSurface = new AddedReferenceSurface();

                                                tmpInitSurface.SurfacePlane = ThisFace.GetSurface();
                                                tmpInitSurface.SurfaceFeature = ThisFeature;
                                                tmpInitSurface.SurfaceName = ThisFeature.Name + "@" + compName[0].Name2 + "@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());
                                                tmpInitSurface.BodyOwner = ThisBody.Name;
                                                tmpInitSurface.AttachedBody = ThisInitialTRV;

                                                //add to the surface collection
                                                InitialRefSurfaces.Add(tmpInitSurface);
                                                tmpInitPlane.ReferenceSurface = tmpInitSurface;

                                            }
                                            
                                        }
                                        
                                        //Check the face configurations
                                        tmpInitPlane.Patterns = CollectPatterns(tmpFace, IsPlanar(tmpFace));

                                        //Check the face location on bounding box
                                        //tmpInitPlane.isOnBB = CheckFaceOnBB(tmpFace);
                                        tmpInitPlane.isOnBB = CheckFaceOnBB(tmpInitPlane);

                                        //set the MaxMinValue
                                        SetMaxMin(TessVerticesArray, ref MaxMinValue);

                                        //check the profile on this face add it to InitialRefProfiles if exist
                                        List<Edge> EdgeProfiles = new List<Edge>();
                                        List<AddedReferenceProfile> ConnectedProfile = new List<AddedReferenceProfile>();
                                        bRet = FindRefProfile(tmpFace, ref EdgeProfiles);

                                        if (bRet == true)
                                        {
                                            foreach (Edge ThisEdge in EdgeProfiles)
                                            { 
                                                tmpInitProfile = new AddedReferenceProfile();
                                                string CurveType = "";

                                                tmpInitProfile.name = refName;
                                                tmpInitProfile.EdgeProfile = ThisEdge;
                                                tmpInitProfile.CurveParam = GetCurveParam(ThisEdge, ref CurveType, refName);
                                                tmpInitProfile.CurveType = CurveType;
                                                tmpInitProfile.AttachedFace = tmpFace;
                                                tmpInitProfile.ReferencePlane = swRefPlane;
                                                tmpInitProfile.BodyOwner = ThisBody.Name;
                                                tmpInitProfile.AttachedBody = ThisInitialTRV;
                                                
                                                InitialRefProfiles.Add(tmpInitProfile); //add the profile to initial profile list
                                                
                                                ConnectedProfile.Add(tmpInitProfile); //collect all the profile that is correspond to this profile
                                                
                                            }

                                            tmpInitPlane.Profiles = ConnectedProfile; //add the collected profile to the corresponding plane

                                        }

                                        //add the reference plane
                                        InitialRefPlanes.Add(tmpInitPlane);

                                    }

                                }
                            }

                            //add this initial TRV body into collection of initial TRV bodies
                            InitialTRVBodies.Add(ThisInitialTRV);

                        }

                    }

                    //store number of reference plane
                    addedRefPlane = countFace;

                    //close editing the raw material
                    //assyModel.EditAssembly();

                    boolStatus = compName[0].Select2(true, 0);
                    
                }
            }

            //Doc.ClearSelection2(true);

            ProcessLog_TaskPaneHost.LogProcess("Generate all coincident planes");
            PPDetails_TaskPaneHost.LogProcess("Generate all coincident planes");
        }

        //tools for plane generator
        #region PlaneGenerator

        //save the number of generated reference plane
        public int addedRefPlane { get; set; }

        //save all initial body
        public List<RemovalBody> InitialTRVBodies;

        //save all generated reference plane
        public List<AddedReferencePlane> InitialRefPlanes;

        //save all generated edge profile
        public List<AddedReferenceProfile> InitialRefProfiles;

        //save all generated reference surface
        public List<AddedReferenceSurface> InitialRefSurfaces;

        //save all circle pattern
        public List<CircularPattern> InitialCirclePattern;

        //save all noncircle pattern
        public List<NonCircularPattern> InitialNonCirclePattern;

        //get the normal array of doubles
        public static Object _GetNormalArray(Face2 ThisFace)
        {
            if (ThisFace.Normal != null)
            {
                return (Object)ThisFace.Normal;
            }

            return null;
        }

        //check whether if its a planar or non-planar face
        public static bool IsPlanar(Face2 ThisFace)
        {
            Double[] ThisFaceNormal = (Double[]) _GetNormalArray(ThisFace);

            //this face is non-planar if only if all normal equal to 0
            if (isEqual(ThisFaceNormal[0], 0) && isEqual(ThisFaceNormal[1], 0) && isEqual(ThisFaceNormal[2], 0))
            {
                return false;
            }
                        
            return true;
        }

        //check whether if its a flat or inclined face
        public bool IsFlat(Face2 ThisFace)
        { 
            if (IsPlanar(ThisFace))
            {
                Double[] ThisFaceNormal = (Double[])_GetNormalArray(ThisFace);

                //this face is flat if only if the z normal equal to 1
                if (isEqual(ThisFaceNormal[2], 1))
                {
                    return true;
                }
            
            }

            return false;
        }

        //set the reference plane for non-planar and inclined surface
        public RefPlane SetupRefPlane(Face2 ThisFace, ModelDoc2 ThisModelDoc, Double[] ThisTessArray)
        {
            Object BoxFaceArray = null;
            Object TessFaceArray = null;
            Double[] BoxVerticesArray = null;
            Double[] TessVerticesArray = null;
            //Double DepthZ;
            Double[] TmpPoint;
            ModelDocExtension SwExt = null;
            FeatureManager SwFM = null;
            Boolean SelectionStatus = false;
            RefPlane RefPlaneInstance = null;
            SketchPoint SkPoint1, SkPoint2, SkPoint3 = null;

            try
            {
                SwExt = ThisModelDoc.Extension;
                SwFM = ThisModelDoc.FeatureManager;

                BoxFaceArray = (Object)ThisFace.GetBox();
                BoxVerticesArray = (Double[])BoxFaceArray;

                //TessFaceArray = (Object)ThisFace.GetTessTriangles(true);
                //TessVerticesArray = GetMaxMin(TessFaceArray);

                ////get the lowest Z level
                //if (BoxVerticesArray[2] < BoxVerticesArray[5]) { DepthZ = BoxVerticesArray[2]; }
                //else { DepthZ = BoxVerticesArray[5]; }

                //DepthZ = TessVerticesArray[5];


                //draw the 1st corner points
                ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[0], ThisTessArray[1], ThisTessArray[5] };
                SkPoint1 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                //ThisModelDoc.SketchManager.Insert3DSketch(true);

                //draw the 2nd corner points
                //ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[0], ThisTessArray[4], ThisTessArray[5] };
                SkPoint2 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                //ThisModelDoc.SketchManager.Insert3DSketch(true);

                //draw the 3rd corner points
                //ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[3], ThisTessArray[4], ThisTessArray[5] };
                SkPoint3 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                ThisModelDoc.SketchManager.Insert3DSketch(true);

                //select the constraint and insert the reference plane
                SkPoint1.Select2(true, 0);
                SkPoint2.Select2(true, 1);
                SkPoint3.Select2(true, 2);
                
                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[0], ThisTessArray[1], ThisTessArray[5], true, 0, null, 1);
                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[0], ThisTessArray[4], ThisTessArray[5], true, 1, null, 1);
                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[3], ThisTessArray[4], ThisTessArray[5], true, 2, null, 1);

                RefPlaneInstance = (RefPlane)SwFM.InsertRefPlane(4, 0, 4, 0, 4, 0);

                if (RefPlaneInstance != null) { return RefPlaneInstance; }

            }

            catch (Exception Ex)
            {
                iSwApp.SendMsgToUser(Ex.Message);
            }
 
            return null;
 
        }

        //OVERLOAD set the reference plane for non-planar and inclined surface
        public RefPlane SetupRefPlane(Face2 ThisFace, ModelDoc2 ThisModelDoc, Component2 ThisComp, Double[] ThisTessArray)
        {
            Object BoxFaceArray = null;
            Object TessFaceArray = null;
            Double[] BoxVerticesArray = null;
            Double[] TessVerticesArray = null;
            //Double DepthZ;
            Double[] TmpPoint;
            ModelDocExtension SwExt = null;
            FeatureManager SwFM = null;
            Boolean SelectionStatus = false;
            RefPlane RefPlaneInstance = null;
            SketchPoint SkPoint1, SkPoint2, SkPoint3 = null;

            try
            {
                AssemblyDoc ThisAssyDoc = (AssemblyDoc)ThisModelDoc;

                ThisComp.Select2(true, 0);

                ThisAssyDoc.EditPart();

                SwExt = ThisModelDoc.Extension;
                SwFM = ThisModelDoc.FeatureManager;

                BoxFaceArray = (Object)ThisFace.GetBox();
                BoxVerticesArray = (Double[])BoxFaceArray;

                //TessFaceArray = (Object)ThisFace.GetTessTriangles(true);
                //TessVerticesArray = GetMaxMin(TessFaceArray);

                ////get the lowest Z level
                //if (BoxVerticesArray[2] < BoxVerticesArray[5]) { DepthZ = BoxVerticesArray[2]; }
                //else { DepthZ = BoxVerticesArray[5]; }

                //DepthZ = TessVerticesArray[5];


                //draw the 1st corner points
                ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[0], ThisTessArray[1], ThisTessArray[5] };
                SkPoint1 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                //ThisModelDoc.SketchManager.Insert3DSketch(true);

                //draw the 2nd corner points
                //ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[0], ThisTessArray[4], ThisTessArray[5] };
                SkPoint2 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                //ThisModelDoc.SketchManager.Insert3DSketch(true);

                //draw the 3rd corner points
                //ThisModelDoc.SketchManager.Insert3DSketch(true);
                TmpPoint = new Double[] { ThisTessArray[3], ThisTessArray[4], ThisTessArray[5] };
                SkPoint3 = (SketchPoint)ThisModelDoc.SketchManager.CreatePoint(TmpPoint[0], TmpPoint[1], TmpPoint[2]);
                ThisModelDoc.SketchManager.Insert3DSketch(true);

                //select the constraint and insert the reference plane
                SkPoint1.Select2(true, 0);
                SkPoint2.Select2(true, 1);
                SkPoint3.Select2(true, 2);

                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[0], ThisTessArray[1], ThisTessArray[5], true, 0, null, 1);
                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[0], ThisTessArray[4], ThisTessArray[5], true, 1, null, 1);
                //SelectionStatus = SwExt.SelectByID2("", "EXTSKETCHPOINT", ThisTessArray[3], ThisTessArray[4], ThisTessArray[5], true, 2, null, 1);

                RefPlaneInstance = (RefPlane)SwFM.InsertRefPlane(4, 0, 4, 0, 4, 0);

                ThisAssyDoc.EditAssembly();

                if (RefPlaneInstance != null) { return RefPlaneInstance; }

            }

            catch (Exception Ex)
            {
                iSwApp.SendMsgToUser(Ex.Message);

            }

            return null;

        }

        //get the maximum value of four values
        public static Single GetMax(Single Val1, Single Val2, Single Val3, Single Val4)
        { 
            Single MaxValue = Val1;
            
            if (Val2 > MaxValue) { MaxValue = Val2; }

            if (Val3 > MaxValue) { MaxValue = Val3; }

            if (Val4 > MaxValue) { MaxValue = Val4; }

            return MaxValue;
        }

        //get the minimum value of four values
        public static Single GetMin(Single Val1, Single Val2, Single Val3, Single Val4)
        { 
            Single MinValue = Val1;

            if (Val2 < MinValue) { MinValue = Val2; }

            if (Val3 < MinValue) { MinValue = Val3; }

            if (Val4 < MinValue) { MinValue = Val4; }

            return MinValue;
        }

        //get the maximum and minimum from a tesselation data (bounding box)
        public static Double[] GetMaxMin(Object TessData)
        {   
            Single[] TessArray;
            Single X_Max, X_Min, Y_Max, Y_Min, Z_Max, Z_Min;
                        
            X_Max = Single.MinValue;
            Y_Max = Single.MinValue;
            Z_Max = Single.MinValue;
            X_Min = Single.MaxValue;
            Y_Min = Single.MaxValue;
            Z_Min = Single.MaxValue;
                        
            //get the tesselation data and compare them with the current MaxMin value
            TessArray = (Single[])TessData;

            for(int i = 0; i < (TessArray.Count()/9); i++)
            {
                X_Max = GetMax(X_Max, TessArray[9 * i + 0], TessArray[9 * i + 3], TessArray[9 * i + 6]);
                X_Min = GetMin(X_Min, TessArray[9 * i + 0], TessArray[9 * i + 3], TessArray[9 * i + 6]);

                Y_Max = GetMax(Y_Max, TessArray[9 * i + 1], TessArray[9 * i + 4], TessArray[9 * i + 7]);
                Y_Min = GetMin(Y_Min, TessArray[9 * i + 1], TessArray[9 * i + 4], TessArray[9 * i + 7]);

                Z_Max = GetMax(Z_Max, TessArray[9 * i + 2], TessArray[9 * i + 5], TessArray[9 * i + 8]);
                Z_Min = GetMin(Z_Min, TessArray[9 * i + 2], TessArray[9 * i + 5], TessArray[9 * i + 8]);

            }
            
            return new Double[6] {
                Convert.ToDouble(X_Max), 
                Convert.ToDouble(Y_Max), 
                Convert.ToDouble(Z_Max), 
                Convert.ToDouble(X_Min), 
                Convert.ToDouble(Y_Min), 
                Convert.ToDouble(Z_Min) };

        }

        //update the MaxMin value by comparing tesselated face data
        public static bool SetMaxMin(Double[] TessVertices, ref Double[] ThisCurrentMaxMin)
        {
            Double[] TmpValue;
            Double X_Max, X_Min, Y_Max, Y_Min, Z_Max, Z_Min;

            if (ThisCurrentMaxMin == null)
            {
                ThisCurrentMaxMin = new Double[6];
                X_Max = Double.MinValue;
                Y_Max = Double.MinValue;
                Z_Max = Double.MinValue;
                X_Min = Double.MaxValue;
                Y_Min = Double.MaxValue;
                Z_Min = Double.MaxValue;
            }
            else
            {   
                X_Max = ThisCurrentMaxMin[0];
                Y_Max = ThisCurrentMaxMin[1];
                Z_Max = ThisCurrentMaxMin[2];
                X_Min = ThisCurrentMaxMin[3];
                Y_Min = ThisCurrentMaxMin[4];
                Z_Min = ThisCurrentMaxMin[5];
            }

            if (TessVertices[0] > X_Max) { X_Max = TessVertices[0]; }
            if (TessVertices[1] > Y_Max) { Y_Max = TessVertices[1]; }
            if (TessVertices[2] > Z_Max) { Z_Max = TessVertices[2]; }
            if (TessVertices[3] < X_Min) { X_Min = TessVertices[3]; }
            if (TessVertices[4] < Y_Min) { Y_Min = TessVertices[4]; }
            if (TessVertices[5] < Z_Min) { Z_Min = TessVertices[5]; }

            TmpValue = new Double[] { X_Max, Y_Max, Z_Max, X_Min, Y_Min, Z_Min };
            ThisCurrentMaxMin = TmpValue;
            
            return true;
        }

        //OVERLOAD update the MaxMin value by comparing tesselated face data
        public static bool SetMaxMin(Double[] TessVertices, Double[] ThisCurrentMaxMin)
        {
            Double[] TmpValue;
            Double X_Max, X_Min, Y_Max, Y_Min, Z_Max, Z_Min;

            if (ThisCurrentMaxMin == null)
            {
                ThisCurrentMaxMin = new Double[6];
                X_Max = Double.MinValue;
                Y_Max = Double.MinValue;
                Z_Max = Double.MinValue;
                X_Min = Double.MaxValue;
                Y_Min = Double.MaxValue;
                Z_Min = Double.MaxValue;
            }
            else
            {
                X_Max = ThisCurrentMaxMin[0];
                Y_Max = ThisCurrentMaxMin[1];
                Z_Max = ThisCurrentMaxMin[2];
                X_Min = ThisCurrentMaxMin[3];
                Y_Min = ThisCurrentMaxMin[4];
                Z_Min = ThisCurrentMaxMin[5];
            }

            if (TessVertices[0] > X_Max) { X_Max = TessVertices[0]; }
            if (TessVertices[1] > Y_Max) { Y_Max = TessVertices[1]; }
            if (TessVertices[2] > Z_Max) { Z_Max = TessVertices[2]; }
            if (TessVertices[3] < X_Min) { X_Min = TessVertices[3]; }
            if (TessVertices[4] < Y_Min) { Y_Min = TessVertices[4]; }
            if (TessVertices[5] < Z_Min) { Z_Min = TessVertices[5]; }

            TmpValue = new Double[] { X_Max, Y_Max, Z_Max, X_Min, Y_Min, Z_Min };
            ThisCurrentMaxMin = TmpValue;

            return true;
        }

        //collect circular and non circular patterns
        public static PatternConfig CollectPatterns(Face2 ThisFace, Boolean IsPlanar)
        {

            PatternConfig ThisFacePattern = new PatternConfig();
            Object[] objLoops = null;
            Object[] vEdgeArr = null;
            objLoops = (Object[])ThisFace.GetLoops();
            Curve ThisCurce = null;

            foreach (Loop2 tmpLoop in objLoops)
            {
                if (tmpLoop.IsOuter() == true)
                {                    
                    if (tmpLoop.GetEdgeCount() != 0)
                    {
                        vEdgeArr = (Object[])tmpLoop.GetEdges();

                        foreach (Edge tmpEdge in vEdgeArr)
                        {
                            ThisCurce = tmpEdge.GetCurve();

                            if (ThisCurce.IsLine()) { ThisFacePattern.Line++; }
                            else if (ThisCurce.IsCircle()) { ThisFacePattern.Circle++; }
                            else { ThisFacePattern.Curve++; }

                        }

                        break;
                    }
                    
                    //process this loop
                    //return isLoopOnThePlane(tmpLoop, SelectedPlane);
                }
            }


            //set the pattern composition
            if (ThisFacePattern.Line > 0 && ThisFacePattern.Circle == 0 && ThisFacePattern.Curve == 0) { ThisFacePattern.Type = "LINE"; }
            else if (ThisFacePattern.Line == 0 && ThisFacePattern.Circle > 0 && ThisFacePattern.Curve == 0) { ThisFacePattern.Type = "CIRCLE"; }
            else if (ThisFacePattern.Line == 0 && ThisFacePattern.Circle == 0 && ThisFacePattern.Curve > 0) { ThisFacePattern.Type = "CURVE"; }
            else { ThisFacePattern.Type = "MIX"; }

            return ThisFacePattern;
        }

        //get the pattern type based on the edge configuration
        public static String GetPatternType(PatternConfig ThisPatternConfig)
        {
            if (ThisPatternConfig.Line > 0 && ThisPatternConfig.Circle == 0 && ThisPatternConfig.Curve == 0) { return "LINE"; }
            else if (ThisPatternConfig.Line == 0 && ThisPatternConfig.Circle > 0 && ThisPatternConfig.Curve == 0) { return "CIRCLE"; }
            else if (ThisPatternConfig.Line == 0 && ThisPatternConfig.Circle == 0 && ThisPatternConfig.Curve > 0) { return "CURVE"; }
            else { return "MIX"; }
        }

        //check the location of a face whether if it is located on bounding box or not
        private Boolean CheckFaceOnBB(Face2 ThisFace, out Face2 SimilarFace)
        {
            if (IsPlanar(ThisFace) == true)
            {
                foreach (Face2 ThisRawFace in RawMaterialFaces)
                {
                    if (ThisFace.IsCoincident(ThisRawFace, 0.001) == 0 ||
                        ThisFace.IsCoincident(ThisRawFace, 0.001) == 1)
                    {
                        SimilarFace = ThisRawFace;
                        return true;
                    }
                    else
                    {
                        if (CheckFaceSim(ThisFace, ThisRawFace) == true)
                        {
                            SimilarFace = ThisRawFace;
                            return true;
                        }
                    }

                }
            }

            SimilarFace = null;
            return false;
        }

        
        //OVERLOAD check the location of a face whether if it is located on bounding box or not
        private Boolean CheckFaceOnBB(AddedReferencePlane ThisPlane)
        {
            Face2 FaceOnRefPlane = (Face2)ThisPlane.AttachedFace;

            if (ThisPlane.IsPlanar == true)
            {
                foreach (Face2 ThisRawFace in RawMaterialFaces)
                {
                    if (FaceOnRefPlane.IsCoincident(ThisRawFace, 0.001) == 0 ||
                        FaceOnRefPlane.IsCoincident(ThisRawFace, 0.001) == 1)
                    {
                        return true;
                    }
                    else
                    {
                        if (CheckFacePlaneSim(ThisPlane, ThisRawFace) == true)
                        {
                            return true;
                        }
                    }

                }
            }
            
            return false;
        }

        //check similarity between two faces based on normal direction and their coincident
        private bool CheckFacePlaneSim(AddedReferencePlane ThisPlane, Face2 FaceB)
        {
            Face2 FaceOnRefPlane = (Face2)ThisPlane.AttachedFace;
            MathUtility ThisMathUtility = (MathUtility) SwApp.GetMathUtility();
            double[] NormA = (double[])FaceOnRefPlane.Normal;
            double[] NormB = (double[])FaceB.Normal;
            bool NormalSim;
            bool LevelSim;

            //check if parallel
            NormalSim = checkNormalSim(NormA, NormB, ThisMathUtility);

            if (NormalSim == true)
            {
                //check if coincident
                LevelSim = CheckFaceLevel(ThisPlane, NormA, FaceB, ThisMathUtility);

                if (LevelSim == true)
                {
                    return true;
                }
            }

            return false;
        }

        //OVERLOAD check similarity between two faces based on normal direction and their coincident
        private static bool CheckFaceSim(Face2 FaceA, Face2 FaceB)
        {
            
            MathUtility ThisMathUtility = (MathUtility)SwApp.GetMathUtility();
            double[] NormA = (double[])FaceA.Normal;
            double[] NormB = (double[])FaceB.Normal;
            bool NormalSim;
            bool LevelSim;

            //check if parallel
            NormalSim = checkNormalSim(NormA, NormB, ThisMathUtility);

            if (NormalSim == true)
            {
                
                //check curve type on face B
                object[] ObjLoopsfaceB = (object[])FaceB.GetLoops();
                object[] ThisLoopEdgeB;
                string CurveB = "";
                int CurveTypeB;
                Curve CurveBInstance;
                CurveParamData ParamDataB;
                double[] CurvePointsB = null;
                MathPoint ThisPointB;
                Edge ThisEdgeB;

                foreach (Loop2 ThisLoop in ObjLoopsfaceB)
                {
                    if (ThisLoop.IsOuter())
                    {
                        ThisLoopEdgeB = (object[])ThisLoop.GetEdges();

                        ThisEdgeB = (Edge)ThisLoopEdgeB.FirstOrDefault();
                        CurveBInstance = (Curve) ThisEdgeB.GetCurve();

                        ParamDataB = (CurveParamData)ThisEdgeB.GetCurveParams3();
                        CurvePointsB = (double[])ParamDataB.StartPoint;
                        CurveTypeB = ParamDataB.CurveType;

                        break;
                           
                    }
                }

                //check curve type on face A
                object[] ObjLoopsfaceA = (object[])FaceA.GetLoops();
                object[] ThisLoopEdgeA;
                string CurveA = "";
                int CurveTypeA;
                Curve CurveAInstance;
                CurveParamData ParamDataA;
                double[] CurvePointsA = null;
                MathPoint ThisPointA;
                Edge ThisEdgeA;

                foreach (Loop2 ThisLoop in ObjLoopsfaceA)
                {
                    if (ThisLoop.IsOuter())
                    {
                        ThisLoopEdgeA = (object[])ThisLoop.GetEdges();

                        ThisEdgeA = (Edge)ThisLoopEdgeA.FirstOrDefault();
                        CurveAInstance = (Curve)ThisEdgeA.GetCurve();

                        ParamDataA = (CurveParamData)ThisEdgeA.GetCurveParams3();
                        CurvePointsA = (double[])ParamDataA.StartPoint;
                        CurveTypeA = ParamDataA.CurveType;

                        break;
                            
                    }
                }

                //create the point
                ThisPointA = ThisMathUtility.CreatePoint(new double[3] { CurvePointsA[0], CurvePointsA[1], CurvePointsA[2] });
                ThisPointB = ThisMathUtility.CreatePoint(new double[3] { CurvePointsB[0], CurvePointsB[1], CurvePointsB[2] });

                LevelSim = CheckFaceLevel(ThisPointA, NormA, ThisPointB, ThisMathUtility);
                

                if (LevelSim == true)
                {
                    return true;
                }
            }

            return false;
        }

        //check the level between two faces
        private bool CheckFaceLevel(AddedReferencePlane ThisPlane, object PlaneNormal, Face2 FaceB, MathUtility ThisUtils)
        {   
            MathPoint firstPoint = getPointOnPlane(ThisPlane);
            MathPoint secondPoint = GetPointOnFace(FaceB, ThisUtils);
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint);

            MathVector firstVector = ThisUtils.CreateVector(PlaneNormal);

            Double distance = Math.Abs(pointsVector.Dot(firstVector) / firstVector.GetLength());

            //check the distance, if equal with 0, then it should be on the same level
            return isEqual(distance, 0);

        }

        //OVERLOAD check the level between two faces
        private bool CheckFaceLevel(Face2 ThisFace, object PlaneNormal, Face2 FaceB, MathUtility ThisUtils)
        {
            MathPoint firstPoint = GetPointOnFace(ThisFace, ThisUtils);
            MathPoint secondPoint = GetPointOnFace(FaceB, ThisUtils);
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint);

            MathVector firstVector = ThisUtils.CreateVector(PlaneNormal);

            Double distance = Math.Abs(pointsVector.Dot(firstVector) / firstVector.GetLength());

            //check the distance, if equal with 0, then it should be on the same level
            return isEqual(distance, 0);

        }

        //OVERLOAD check the level between two faces with curveparam points
        private static bool CheckFaceLevel(MathPoint ThisPointA, object PlaneNormal, MathPoint ThisPointB, MathUtility ThisUtils)
        {
            //MathPoint firstPoint = GetPointOnFace(ThisFace, ThisUtils);
            MathPoint firstPoint = ThisPointA;
            MathPoint secondPoint = ThisPointB;
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint);

            MathVector firstVector = ThisUtils.CreateVector(PlaneNormal);

            Double distance = Math.Abs(pointsVector.Dot(firstVector) / firstVector.GetLength());

            //check the distance, if equal with 0, then it should be on the same level
            return isEqual(distance, 0);

        }

        //get random point on face (only for planar and non bcurve and cyclinder edge
        private MathPoint GetPointOnFace(Face2 ThisFace, MathUtility ThisUtils)
        {
            object[] ObjEdges = (object[])ThisFace.GetEdges();
            Edge ThisEdge;
            Vertex ObjVertex;
            double[] Vertex;

            ThisEdge = (Edge)ObjEdges.Last();
            ObjVertex = ThisEdge.GetStartVertex();
            Vertex = (double[])ObjVertex.GetPoint();
            
            return ThisUtils.CreatePoint(Vertex);
        }
    

        //get the reference of profile on ThisFace
        private Boolean FindRefProfile(Face2 ThisFace, ref List<Edge> PutIntoThisEdge)
        {
            object[] EdgesOnLoop = null;
            
            if (ThisFace != null)
            {
                object[] LoopsOnFace = (object[])ThisFace.GetLoops();

                if (LoopsOnFace.Count() == 1)
                {
                    Loop2 ThisLoop = (Loop2) LoopsOnFace.First();
                    EdgesOnLoop = (object[])ThisLoop.GetEdges();

                    if (EdgesOnLoop.Count() == 1)
                    {
                        PutIntoThisEdge.Add((Edge)EdgesOnLoop.First());
                    }

                }
                else
                {
                    foreach (Loop2 ThisLoop in LoopsOnFace)
                    {
                        EdgesOnLoop = (object[])ThisLoop.GetEdges();

                        if (ThisLoop.IsOuter() == false)
                        {
                            if (EdgesOnLoop.Count() == 1)
                            {
                                PutIntoThisEdge.Add((Edge)EdgesOnLoop.First());
                            }

                        }
                        
                    }

                    

                }

                if (PutIntoThisEdge.Count() >= 1)
                {
                    return true;
                }

            }

            return false;
        }

        //get the circle parameter
        private double[] GetCurveParam(Edge ThisEdge, ref string ThisCurveType, string RefName)
        { 
            Curve ThisEdgeCurve = (Curve) ThisEdge.GetCurve();
            
            double[] CurveParam = null;
            
            //check first if its a circle
            if (ThisEdgeCurve.IsCircle())
            {
                ThisCurveType = "circle";
                CurveParam = (double[])ThisEdgeCurve.CircleParams;

                ProcessLog_TaskPaneHost.LogProcess("there is a circle on " + RefName);
            }
            //check if its an ellipse
            else if (ThisEdgeCurve.IsEllipse())
            {
                ThisCurveType = "ellipse";
                CurveParam = (double[])ThisEdgeCurve.GetEllipseParams();
                ProcessLog_TaskPaneHost.LogProcess("there is an ellipse on " + RefName);
            }
            //check if its a Bcurve
            else if (ThisEdgeCurve.IsBcurve())
            {
                ThisCurveType = "bcurve";

                CurveParam = GetCurveMaxTess(ThisEdgeCurve);

                ProcessLog_TaskPaneHost.LogProcess("there is a bcurve on " + RefName);
            }

            return CurveParam;
        }

        //get max tesselation point from bcurve
        public double[] GetCurveMaxTess(Curve ThisCurve)
        { 
            double[] CurveParam = null;
            double StartPoint = 0.0;
            double EndPoint = 0.0;
            double[] ObjStartPoint = null;
            double[] ObjEndPoint = null;
            double[] CurveTessPts = null;
            Single[] CurveTessPtsSingle = null;
            bool ClosedStatus = false;
            bool PeriodicType = false;
            bool RetVal = false;
            int i = 0;

            RetVal = ThisCurve.GetEndParams(out StartPoint, out EndPoint, out ClosedStatus, out PeriodicType);

            if (RetVal == true)
            {
                ObjStartPoint = (double[]) ThisCurve.Evaluate2(StartPoint, 7);
                ObjEndPoint = (double[])ThisCurve.Evaluate2(EndPoint, 7);
                CurveTessPts = (double[])ThisCurve.GetTessPts(0.001, 0.001, ObjStartPoint, ObjEndPoint);

                CurveTessPtsSingle = new Single[CurveTessPts.Count()];

                i = 0;

                foreach (double ThisValue in CurveTessPts)
                {
                    CurveTessPtsSingle[i] = Convert.ToSingle(ThisValue);
                    i++;
                }

                CurveParam = GetMaxMin(CurveTessPtsSingle);

            }


            return CurveParam;
        }

        //calculate the TAD (listing all TAD and the recommentation based on LWD
        public bool CalculateTAD(Body2 ThisBody, ref RemovalBody ThisTRV, object[] ProdFaces)
        {
            object[] ThisBodyFaces = (object[])ThisBody.GetFaces();
            List<Face2> ClosedFaces = new List<Face2>();
            bool Retval = false;
            double similarity;
            MathUtility SwMathUtility = (MathUtility)SwApp.GetMathUtility();
            MathVector VectorValue;
            List<MathVector> ListVectorValue = new List<MathVector>();
            List<double[]> ListOfDoubles = new List<double[]>();
            List<double[]> ClosedFaceNorms = new List<double[]>();
            List<double> ListOfAreas = new List<double>();

            //find the TAD (open face) from this body
            foreach (Face2 BodyFace in ThisBodyFaces)
            {   
                foreach (Face2 ProdFace in ProdFaces)
                {
                    int sim = BodyFace.IsCoincident(ProdFace, 0.001);
 
                    if (sim == 1 || sim == 0)
                    {
                        Retval = true;
                        ClosedFaces.Add(BodyFace); //collect the coincident face (closed)
                        break;
                    }
                }

                if (Retval == false && BodyFace.Normal != null)
                {
                    VectorValue = SwMathUtility.CreateVector(BodyFace.Normal);

                    if (ListVectorValue.Count == 0)
                    {
                        ListVectorValue.Add(VectorValue);
                        ListOfDoubles.Add((double[])BodyFace.Normal);
                        ListOfAreas.Add(GetAreaOfTess((object)BodyFace.GetTessTriangles(true)));
                    }
                    else
                    {
                        Retval = false;

                        foreach (MathVector ThisVector in ListVectorValue)
                        {
                            similarity = VectorValue.Dot(ThisVector) / (VectorValue.GetLength() * ThisVector.GetLength());

                            if (isEqual(similarity, 1) == true)
                            {
                                Retval = true;
                                break;
                            }

                        }

                        if (Retval == false)
                        {
                            ListVectorValue.Add(VectorValue);
                            ListOfDoubles.Add((double[])BodyFace.Normal);
                            ListOfAreas.Add(GetAreaOfTess((object)BodyFace.GetTessTriangles(true)));
                        }

                    }
                }

                //reset the bool to false
                Retval = false;
            }

            //add the TAD (open faces) to the removal body
            ThisTRV.ListOfTAD = GetTADForm(ListOfDoubles);

            if (ThisTRV.ListOfTAD.Count() == 6)
            {
                int i = 0;
                int SmallestArea = 0;
                double PreviousArea = 1000000000000;

                foreach (double ThisArea in ListOfAreas)
                {
                    if (ThisArea < PreviousArea)
                    {
                        PreviousArea = ThisArea;
                        SmallestArea = i;
                    }
                    i++;
                }

                ThisTRV.ClosedTAD = GetTADForm(ListOfDoubles[SmallestArea]);

            }
            else
            {
                //find the unaccessible TAD (closed face)
                if (ClosedFaces.Count != 0)
                {
                    //collect the face normal
                    foreach (Face2 ThisFace in ClosedFaces)
                    {
                        if (ThisFace.Normal != null)
                        {
                            ClosedFaceNorms.Add((double[])ThisFace.Normal);
                        }
                    }

                    ListOfDoubles = new List<double[]>();

                    foreach (double[] ClosedNorm in ClosedFaceNorms)
                    {
                        VectorValue = SwMathUtility.CreateVector(ClosedNorm);

                        Retval = false;

                        foreach (MathVector ThisVector in ListVectorValue)
                        {
                            similarity = VectorValue.Dot(ThisVector) / (VectorValue.GetLength() * ThisVector.GetLength());

                            if (isEqual(similarity, 1) == true)
                            {
                                Retval = true;
                                break;
                            }

                        }

                        if (Retval == false)
                        {
                            ListVectorValue.Add(VectorValue);
                            ListOfDoubles.Add(ClosedNorm);
                        }

                    }

                }

                //add the unaccesible TAD (closed faces) into removal body
                ThisTRV.ClosedTAD = GetTADForm(ListOfDoubles);

                //remove any empty TAD
                ThisTRV.ClosedTAD = RemoveEmptyTAD(ThisTRV.ClosedTAD);
            }

            //get the recommended TAD based on existing closedTAD and its LWD
            ThisTRV.RecommendedTAD = GetRecommendedTAD(ref ThisTRV);

            //set the virtual point of the body
            ThisTRV.VirtualPoint = GetBodyVirtualPoint(ThisTRV);

            return true;
        }

        //get the lowest virtual point based on the direction of the recommended TAD
        public double[] GetBodyVirtualPoint(RemovalBody ThisBody)
        {
            double[] TmpVector = new double[3] { ThisBody.RecommendedTAD.X, ThisBody.RecommendedTAD.Y, ThisBody.RecommendedTAD.Z };
            double[] MaxMin = ThisBody.MaxMin;
            double[] MidPoint = ThisBody.MidPoint;
            double[] VirtualPoint = null;

            if (isEqual(TmpVector[0], 1) == true) // X+
            {
                VirtualPoint = new double[3] {MaxMin[3], MidPoint[1], MidPoint[2] };
            }
            else if (isEqual(TmpVector[0], -1) == true) // X-
            {
                VirtualPoint = new double[3] { MaxMin[0], MidPoint[1], MidPoint[2] }; 
            }
            else if (isEqual(TmpVector[1], 1) == true) // Y+
            {
                VirtualPoint = new double[3] { MidPoint[0], MaxMin[4], MidPoint[2] }; 
            }
            else if (isEqual(TmpVector[1], -1) == true) // Y-
            {
                VirtualPoint = new double[3] { MidPoint[0], MaxMin[1], MidPoint[2] }; 
            }
            else if (isEqual(TmpVector[2], 1) == true) // Z+
            {
                VirtualPoint = new double[3] { MidPoint[0], MidPoint[1], MaxMin[5] }; 
            }
            else if (isEqual(TmpVector[2], -1) == true) // Z-
            {
                VirtualPoint = new double[3] { MidPoint[0], MidPoint[1], MaxMin[2] };
            }
            
            return VirtualPoint;
        }

        //get the area of the corresponding tesselation of face
        public double GetAreaOfTess(object ThisFaceTess)
        {
            Double[] ThisFaceMaxMin = GetMaxMin(ThisFaceTess);
            
            //area of YZ
            if (isEqual(ThisFaceMaxMin[0], ThisFaceMaxMin[3]))
            {
                return Math.Abs((ThisFaceMaxMin[1] - ThisFaceMaxMin[4]) * (ThisFaceMaxMin[2] - ThisFaceMaxMin[5]));
            }

            //area of XZ
            if (isEqual(ThisFaceMaxMin[1], ThisFaceMaxMin[4]))
            {
                return Math.Abs((ThisFaceMaxMin[0] - ThisFaceMaxMin[3]) * (ThisFaceMaxMin[2] - ThisFaceMaxMin[5]));
            }

            //area of XY
            if (isEqual(ThisFaceMaxMin[2], ThisFaceMaxMin[5]))
            {
                return Math.Abs((ThisFaceMaxMin[0] - ThisFaceMaxMin[3]) * (ThisFaceMaxMin[1] - ThisFaceMaxMin[4]));
            }

            return 0.0;
        }

        //remove empty TAD
        public List<TAD> RemoveEmptyTAD(List<TAD> ThisTADs)
        {
            List<int> EmptyIndex = new List<int>();

            //get the index which is empty
            for (int i = 0; i < ThisTADs.Count; i++)
            {
                if (ThisTADs[i].IsEmpty())
                {
                    EmptyIndex.Add(i);
                }
            }

            //sort from biggest to smallest value
            var SortedIndex = from Index in EmptyIndex
                              orderby Index descending
                              select Index;

            //remove the empty TAD space
            foreach (int Index in SortedIndex)
            {
                ThisTADs.RemoveAt(Index);
            }

            return ThisTADs;
        }

        //set the TAD format
        public List<TAD> GetTADForm(List<double[]> ThisDoubles)
        {
            List<TAD> ListOfTAD = new List<TAD>();
            TAD ThisTAD;
            
            //add the TAD (open faces)
            if (ThisDoubles.Count != 0)
            {
                ListOfTAD = new List<TAD>();

                foreach (double[] TmpTAD in ThisDoubles)
                {
                    ThisTAD = new TAD();
                    ThisTAD.X = TmpTAD[0];
                    ThisTAD.Y = TmpTAD[1];
                    ThisTAD.Z = TmpTAD[2];

                    ListOfTAD.Add(ThisTAD);
                }
            }

            return ListOfTAD;
        }

        //OVERLOAD set the TAD format from single TAD
        public List<TAD> GetTADForm(double[] ThisDoubles)
        {
            List<TAD> ListOfTAD = new List<TAD>();
            TAD ThisTAD = new TAD();
            ThisTAD.X = ThisDoubles[0];
            ThisTAD.Y = ThisDoubles[1];
            ThisTAD.Z = ThisDoubles[2];
            
            ListOfTAD.Add(ThisTAD);
            
            return ListOfTAD;
        }



        //evaluate recommended TAD based on LWD from existing closedTAD
        public TAD GetRecommendedTAD(ref RemovalBody ThisBody)
        {
            double[] ThisTADdouble = null;
            TAD ThisTAD = null;
            List<TAD> OpenTAD = ThisBody.ListOfTAD;
            List<TAD> ClosedTAD = ThisBody.ClosedTAD;
            List<TAD> BestTADs = new List<TAD>();
            List<double> AreaOfBestTAD = new List<double>();
            double[] MaxMin;
            double[] TmpVector;
            MathVector TADVector;
            List<MathVector> BasicTAD;
            MathUtility SwMathUtility = (MathUtility)SwApp.GetMathUtility();
            double similarity;
            double MaxArea = 0.0;
            double TmpArea = 0.0;
            
            foreach (TAD TmpClosedTAD in ClosedTAD)
            {
                //inverse the vector direction
                TmpClosedTAD.Scale(-1);

                foreach (TAD TmpOpenTAD in OpenTAD)
                {
                    if (TmpClosedTAD.IsNear(TmpOpenTAD, SwMathUtility))
                    {
                        BestTADs.Add(TmpOpenTAD);
                        break;
                    }
                }
            }

            MaxMin = (double[])ThisBody.MaxMin;

            if (BestTADs.Count == 1)
            {
                TmpVector = new double[3] {BestTADs.First().X, BestTADs.First().Y, BestTADs.First().Z};

                if (isEqual(Math.Abs(TmpVector[0]), 1) == true)
                {
                    TmpArea = Math.Abs((MaxMin[1] - MaxMin[4]) * (MaxMin[2] - MaxMin[5]));
                }
                else if (isEqual(Math.Abs(TmpVector[1]), 1) == true)
                {
                    TmpArea = Math.Abs((MaxMin[0] - MaxMin[3]) * (MaxMin[2] - MaxMin[5]));
                }
                else if (isEqual(Math.Abs(TmpVector[2]), 1) == true)
                {
                    TmpArea = Math.Abs((MaxMin[0] - MaxMin[3]) * (MaxMin[1] - MaxMin[4]));
                }

                ThisBody.BestTADArea = TmpArea * 1000000;

                return BestTADs.First();
            }
            else
            {
                BasicTAD = new List<MathVector>();
                BasicTAD = InitBasicTAD();

                foreach (TAD TmpTAD in BestTADs)
                {
                    TADVector = SwMathUtility.CreateVector(new double[3] {TmpTAD.X, TmpTAD.Y, TmpTAD.Z});
                    
                    foreach (MathVector ThisVector in BasicTAD)
                    {
                        similarity = TADVector.Dot(ThisVector) / (TADVector.GetLength() * ThisVector.GetLength());

                        if (similarity > Math.Cos(45))
                        {
                            TmpVector = (double[])ThisVector.ArrayData;

                            if (isEqual(Math.Abs(TmpVector[0]), 1) == true) 
                            {
                                TmpArea = Math.Abs((MaxMin[1] - MaxMin[4]) * (MaxMin[2] - MaxMin[5]));
                            }
                            else if (isEqual(Math.Abs(TmpVector[1]), 1) == true)
                            {
                                TmpArea = Math.Abs((MaxMin[0] - MaxMin[3]) * (MaxMin[2] - MaxMin[5]));
                            }
                            else if (isEqual(Math.Abs(TmpVector[2]), 1) == true)
                            {
                                TmpArea = Math.Abs((MaxMin[0] - MaxMin[3]) * (MaxMin[1] - MaxMin[4]));
                            }

                            if (TmpArea > MaxArea)
                            {
                                MaxArea = TmpArea;
                                ThisTADdouble = TmpVector;
                            }

                        }

                    }

                }
            }

            if (ThisTADdouble != null)
            {
                //save the area of the recommended TAD
                ThisBody.BestTADArea = MaxArea * 1000000;

                ThisTAD = new TAD();
                ThisTAD.X = ThisTADdouble[0];
                ThisTAD.Y = ThisTADdouble[1];
                ThisTAD.Z = ThisTADdouble[2];

            }

            return ThisTAD;
        }

        //initialize basic direction of TAD
        public List<MathVector> InitBasicTAD()
        {
            MathUtility SwMathUtility = (MathUtility)SwApp.GetMathUtility();
            List<MathVector> BasicTAD = new List<MathVector>();
            double[] ThisTAD;

            ThisTAD = new double[3] {1, 0, 0}; // X+
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));

            ThisTAD = new double[3] {-1, 0, 0}; // X-
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));

            ThisTAD = new double[3] {0, 1, 0}; // Y+
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));

            ThisTAD = new double[3] {0, -1, 0}; // Y-
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));
            
            ThisTAD = new double[3] {0, 0, 1}; // Z+
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));

            ThisTAD = new double[3] {0, 0, -1}; // Z-
            BasicTAD.Add(SwMathUtility.CreateVector(ThisTAD));

            return BasicTAD;
        }

        //evaluate the similarity of TAD
        public bool IsTADsimilar(TAD firstTAD, TAD SecondTAD)
        {
            if ((isEqual(firstTAD.X, SecondTAD.X) == true) && 
                (isEqual(firstTAD.Y, SecondTAD.Y) == true) &&
                (isEqual(firstTAD.Z, SecondTAD.Z) == true))
            {
                return true;
            }
            
            return false;
        }


        //calculate the centroid of TRV body
        public bool CalculateMidPoint(ModelDoc2 ThisRootModDoc, Component2 ThisCompBody, Body2 ThisBody, ref RemovalBody ThisTRV)
        {
            MathUtility swMathUtil = (MathUtility)SwApp.GetMathUtility();
            SelectionMgr swSelMgr = (SelectionMgr)ThisRootModDoc.SelectionManager;
            MathTransform swXform = (MathTransform)ThisCompBody.Transform2;

            object[] AllFaces = (object[])ThisBody.GetFaces();
            Single[] TessFaceArray;
            double[] MaxMin;
            var TessList = new List<Single>();
            double[] MidPoint = new double[3];
            MathPoint PointXYZ = null;
            
            foreach (Face2 tmpFace in AllFaces)
            {
                TessFaceArray = (Single[])tmpFace.GetTessTriangles(true);
                TessList.AddRange(TessFaceArray);
            }

            TessFaceArray = TessList.ToArray();
            MaxMin = GetMaxMin(TessFaceArray);

            MidPoint[0] = (MaxMin[0] + MaxMin[3]) / 2;
            MidPoint[1] = (MaxMin[1] + MaxMin[4]) / 2;
            MidPoint[2] = (MaxMin[2] + MaxMin[5]) / 2;

            PointXYZ = swMathUtil.CreatePoint(MidPoint);
            PointXYZ = PointXYZ.MultiplyTransform(swXform);

            ThisTRV.MaxMin = MaxMin;
            ThisTRV.MidPoint = new double[3] { PointXYZ.ArrayData[0], PointXYZ.ArrayData[1], PointXYZ.ArrayData[2] };

            return true;
        }

        public Double[] MaxMinValue;


        #endregion 

        public void ReadTolerance()
        {
            iSwApp.SendMsgToUser("the is the call from read tolerance");

            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            
            MathUtility swMathUtils = (MathUtility)SwApp.GetMathUtility();
            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            bool RegStatus = false;
            bool RetVal = false;
            int Errors = 0;
            int Warnings = 0;

            List<string> FaceFound = null;
            
            AssemblyDoc assyModel = null;
            Component2[] compName = null;

            if (Doc.GetType() == 2)
            { 
                assyModel = (AssemblyDoc)Doc;
                if (assyModel.GetComponentCount(false) != 0)
                {
                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    if (compName == null)
                    {
                        //define the raw material and product
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);
                        PPDetails_TaskPaneHost.getCompName(ref compName);
                    }

                    EntityWithTolerance = new List<Entity>();
                    EntityTolerance = new List<PlaneTolerance>();

                    //select the raw material for editing
                    boolStatus = compName[1].Select2(true, 0);

                    //get the tolerance from product
                    if (boolStatus == true)
                    {
                        String CompPathName = compName[1].GetPathName();
                        iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble
                        ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                            (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document
                        
                        ModelDocExtension ThisDocExtension = (ModelDocExtension)CompDocumentModel.Extension;
                        SelectionMgr ThisSelMgr = (SelectionMgr)CompDocumentModel.SelectionManager;
                        SolidWorks.Interop.sldworks.Attribute ThisAtt = default(SolidWorks.Interop.sldworks.Attribute);

                        bool SelectStatus = false;
                        
                        int i = 1;
                        string AttName = "TolNumber" + i.ToString();
                        SelectStatus = ThisDocExtension.SelectByID2(AttName, "ATTRIBUTE", 0, 0, 0, false, 0, null, 0);

                        while (SelectStatus == true)
                        {
                            Feature ThisFeat = (Feature)ThisSelMgr.GetSelectedObject6(1, -1);
                            ThisAtt = (SolidWorks.Interop.sldworks.Attribute)ThisFeat.GetSpecificFeature2();

                            if (ThisAtt != null)
                            {
                                RetVal = GetTolerance(ThisAtt, ref EntityTolerance);
                                if (RetVal == true) { EntityWithTolerance.Add(ThisAtt.GetEntity2()); }
                            }

                            i++;
                            AttName = "TolNumber" + i.ToString();
                            SelectStatus = ThisDocExtension.SelectByID2(AttName, "ATTRIBUTE", 0, 0, 0, false, 0, null, 0);
                        }

                        //set the product's tolerance availability to be true
                        ProductToleranceExist = true;

                    }

                    FaceFound = new List<string>();

                    //copy the tolerance to added reference plane list
                    if (ProductToleranceExist == true)
                    {
                        for (int i = 0; i < EntityWithTolerance.Count(); i++)
                        {
                            int EntityType = EntityWithTolerance[i].GetType();

                            if (EntityType == 2) // Type 2 means FACE type
                            {
                                Face2 ThisFace = (Face2)EntityWithTolerance[i];

                                foreach (AddedReferencePlane ThisRefPlane in InitialRefPlanes)
                                {
                                    if (ThisRefPlane.Tolerance == null)
                                    {
                                        int FaceSimilarity = ThisFace.IsCoincident(ThisRefPlane.AttachedFace, 0.001);

                                        if (FaceSimilarity == 0 || FaceSimilarity == 1)
                                        {
                                            ThisRefPlane.Tolerance = EntityTolerance[i];
                                            FaceFound.Add(ThisRefPlane.name);
                                        }
                                        else
                                        {
                                            Face2 ThisRawFace;

                                            if (CheckFaceOnBB(ThisFace, out ThisRawFace) == true && ThisRefPlane.IsPlanar == true)
                                            {
                                                if (CheckFaceSim(ThisRawFace, ThisRefPlane.AttachedFace) == true)
                                                {
                                                    ThisRefPlane.Tolerance = EntityTolerance[i];
                                                    FaceFound.Add(ThisRefPlane.name);
                                                }
                                            }
                                        }
                                    }

                                }
                            }

                            if (EntityType == 1) // Type 1 means EDGE type
                            {
                                Edge ThisEdge = (Edge)EntityWithTolerance[i];
                                Curve ThisCurve = (Curve)ThisEdge.GetCurve();
                                double[] ThisCurveParam = null;

                                if (ThisCurve.IsCircle())
                                {
                                    ThisCurveParam = (double[])ThisCurve.CircleParams;

                                    foreach (AddedReferenceProfile ThisRefProfile in InitialRefProfiles)
                                    {
                                        if (ThisRefProfile.CurveType.Equals("circle") && ThisRefProfile.CurveParam != null && ThisCurveParam != null)
                                        {
                                            RetVal = CircleSimilarity(ThisRefProfile.CurveParam, ThisCurveParam);

                                            if (RetVal == true)
                                            {
                                                ThisRefProfile.Tolerance = EntityTolerance[i];
                                                ThisRefProfile.ReferenceSurface = GetReferenceSurface(ThisRefProfile);
                                                ThisRefProfile.Main = true;
                                                break;
                                            }
                                        }

                                    }
                                }
                                else if (ThisCurve.IsEllipse())
                                { 
                                    //add the code here for others
                                    ThisCurveParam = (double[])ThisCurve.GetEllipseParams();

                                    foreach (AddedReferenceProfile ThisRefProfile in InitialRefProfiles)
                                    {
                                        if (ThisRefProfile.CurveType.Equals("ellipse") && ThisRefProfile.CurveParam != null && ThisCurveParam != null)
                                        {
                                            RetVal = EllipseSimilarity(ThisRefProfile.CurveParam, ThisCurveParam);

                                            if (RetVal == true)
                                            {
                                                ThisRefProfile.Tolerance = EntityTolerance[i];
                                                ThisRefProfile.ReferenceSurface = GetReferenceSurface(ThisRefProfile);
                                                ThisRefProfile.Main = true;
                                                break;
                                            }
                                        }

                                    }


                                }
                                else if (ThisCurve.IsBcurve())
                                {
                                    ThisCurveParam = (double[])GetCurveMaxTess(ThisCurve);

                                    foreach (AddedReferenceProfile ThisRefProfile in InitialRefProfiles)
                                    {
                                        if (ThisRefProfile.CurveType.Equals("bcurve") && ThisRefProfile.CurveParam != null && ThisCurveParam != null)
                                        {
                                            RetVal = BCurveSimilarity(ThisRefProfile.CurveParam, ThisCurveParam);

                                            if (RetVal == true)
                                            {
                                                ThisRefProfile.Tolerance = EntityTolerance[i];
                                                ThisRefProfile.ReferenceSurface = GetReferenceSurface(ThisRefProfile);
                                                ThisRefProfile.Main = true;
                                                break;
                                            }
                                        }

                                    }
                                   

                                }

                                
                            }

                        }
                    }


                }
            }
        }

        #region ReadTolerance

        //keep the sign to notify that tolerance is available for use (will be true if exist)
        public bool ProductToleranceExist = false;

        //list of entity and the corresponding tolerance from product
        public List<Entity> EntityWithTolerance = null;
        public List<PlaneTolerance> EntityTolerance = null;
        
        //get the reference surface from the collection of reference surface
        public AddedReferenceSurface GetReferenceSurface(AddedReferenceProfile ThisProfile)
        {
            AddedReferenceSurface ThisRefSurface = new AddedReferenceSurface();
            Edge ThisProfileEdge;
            object[] ObjTmpFace;
            object[] ThisProfileObjFace;
            Face2 TmpFace;

            //get the edge and all adjacent faces from this profile
            ThisProfileEdge = (Edge)ThisProfile.EdgeProfile;
            ThisProfileObjFace = (object[])ThisProfileEdge.GetTwoAdjacentFaces2();

            //find this face whether is coincident with the face in initiail surface's feature
            foreach (AddedReferenceSurface TmpSurface in InitialRefSurfaces)
            {
                if (ThisProfile.BodyOwner.Equals(TmpSurface.BodyOwner))
                {
                    ObjTmpFace = (object[]) TmpSurface.SurfaceFeature.GetFaces();

                    if (ObjTmpFace.Count() == 1)
                    {
                        TmpFace = (Face2) ObjTmpFace.First();

                        foreach (Face2 FacesOnProfile in ThisProfileObjFace)
                        { 
                            if(TmpFace.IsCoincident(FacesOnProfile, 0.001) == 0 ||
                                TmpFace.IsCoincident(FacesOnProfile, 0.001) == 1)
                            {
                                return TmpSurface;
                            }
                        }
                    }
                }
            }

            return null;
        }

        //get the tolerance value from attribute
        public bool GetTolerance(SolidWorks.Interop.sldworks.Attribute ThisAttribute, ref List<PlaneTolerance> ToleranceOnFaces)
        {
            PlaneTolerance TmpTolerance = new PlaneTolerance();
            Parameter TolParam = null;
            string TolValue = "";

            //get all the tolerance value
            TolParam = ThisAttribute.GetParameter("datum");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue();
                TmpTolerance.datum = TolValue.Replace("datum ", "");  
            }

            TolParam = ThisAttribute.GetParameter("flatness");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue(); 
                TmpTolerance.flatness = TolValue.Replace("flatness ", ""); 
            }

            TolParam = ThisAttribute.GetParameter("parallelism");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue();
                TmpTolerance.parallelism = TolValue.Replace("parallelism ", ""); 
            }

            TolParam = ThisAttribute.GetParameter("angularity");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue();
                TmpTolerance.angularity = TolValue.Replace("angularity ", ""); 
            }

            TolParam = ThisAttribute.GetParameter("perpendicularity");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue();
                TmpTolerance.perpendicularity = TolValue.Replace("perpendicularity ", ""); 
            }

            TolParam = ThisAttribute.GetParameter("location");
            if (TolParam != null) 
            { 
                TolValue = TolParam.GetStringValue();
                TmpTolerance.location = TolValue.Replace("location ", ""); 
            }

            //add the tolerance to the list
            ToleranceOnFaces.Add(TmpTolerance);

            return true;
        }

        //check the similarity of two circles
        public bool CircleSimilarity(object ObjCircleA, double[] CircleB)
        {
            double[] CircleA = (double[])ObjCircleA;

            if (CircleA.Count() == 7 && CircleB.Count() == 7)
            {
                if ((isEqual(CircleA[0], CircleB[0]) == true) && (isEqual(CircleA[1], CircleB[1]) == true) && (isEqual(CircleA[2], CircleB[2]) == true) && (isEqual(CircleA[6], CircleB[6]) == true))
                {
                    return true;
                }

            }

            return false;
        }

        //check the similarity of two ellipse
        public bool EllipseSimilarity(object ObjEllipseA, double[] EllipseB)
        {
            double[] EllipseA = (double[])ObjEllipseA;
            bool Status = false;
            int i = 0;

            if (EllipseA.Count() == 11 && EllipseB.Count() == 11)
            {
                Status = true;

                while (Status == true && i <= 10)
                {   
                    Status = isEqual(EllipseA[i], EllipseB[i]);
                    i++;
                }

                return Status;
            }

            return false;
        }

        //check the similarity of two Bcurves
        public bool BCurveSimilarity(object ObjBCurveA, double[] BCurveB)
        {
            double[] BCurveA = (double[])ObjBCurveA;
            bool Status = false;
            int i = 0;

            if (BCurveA.Count() == 6 && BCurveB.Count() == 6)
            {
                Status = true;

                while (Status == true && i <= 5)
                {
                    Status = isEqual(BCurveA[i], BCurveB[i]);
                    i++;
                }

                return Status;
            }

            return false;
        }

        #endregion

        //plane Calculator
        public void PlaneCalculator()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            MathUtility swMathUtils = (MathUtility)SwApp.GetMathUtility();
            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            bool RegStatus = false;
            
            AssemblyDoc assyModel = null;
            Component2[] compName = null;

            if (Doc.GetType() == 2)
            {
                assyModel = (AssemblyDoc)Doc;
                if (assyModel.GetComponentCount(false) != 0)
                {
                    try
                    {
                        //get the components
                        ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                        Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                        Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                        if (compName == null)
                        {
                            //define the raw material and product
                            compName = new Component2[2];
                            GetCompName(compInAssembly, ref compName);
                            PPDetails_TaskPaneHost.getCompName(ref compName);
                        }
                        
                        //get the real centroid
                        Body2 mainBody = compName[0].GetBody();
                        MassProperty bodyProperties = Doc.Extension.CreateMassProperty();
                        bodyProperties.AddBodies(mainBody);
                        bodyProperties.UseSystemUnits = false;

                        if (InitialTRVBodies.Count != 0)
                        {
                            Doc.SketchManager.Insert3DSketch(true);

                            foreach (RemovalBody ThisBody in InitialTRVBodies)
                            {
                                centroid = ThisBody.VirtualPoint;
                                SketchPoint skPoint1 = (SketchPoint)Doc.SketchManager.CreatePoint(centroid[0], centroid[1], centroid[2]);
                            }

                            Doc.SketchManager.Insert3DSketch(true);
                        }

                        //get the part document of the raw material
                        ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                        PartDoc swPartDoc = (PartDoc)compModDoc;

                        SelectionMgr SwSelMgr = (SelectionMgr)compModDoc.SelectionManager;
                        SelectData SwSelData = (SelectData)SwSelMgr.CreateSelectData();

                        //get the plane feature
                        boolStatus = getPlanes(swPartDoc, ref InitialRefPlanes);

                        if (boolStatus == true)
                        {
                            //get the plane normal
                            boolStatus = getPlaneNormal(swMathUtils, ref InitialRefPlanes);

                            if (boolStatus == true)
                            {
                                //set the distance
                                boolStatus = setDistance(Doc, swMathUtils, ref InitialRefPlanes);

                                //find the centroid position refer to the normal direction
                                boolStatus = findCPost(ref InitialRefPlanes);

                                //find the outer reference plane
                                boolStatus = FindTheOuter(ref InitialRefPlanes);

                                //analyze the inner plane
                                boolStatus = AnalyzeDirection(ref InitialRefPlanes);

                                //analyze the tolerance if exist
                                if (ProductToleranceExist == true)
                                {
                                    boolStatus = AnalyzeDatum();

                                    boolStatus = AnalyzeProfile(Doc, compModDoc, assyModel, compName[0], Doc.SelectionManager, ref SwSelData);

                                    if (boolStatus == true)
                                    {
                                        //try to extend the surface
                                        

                                    }


                                }

                                List<int> removeId = new List<int>();

                                //store the plane feature, rank, and normal to planeList
                                //registerPlane(planeValue, planeNames, planeNormal, distance, ref removeId, swMathUtils);

                                RegStatus = registerPlane(InitialRefPlanes, ref removeId, swMathUtils);

                                ProcessLog_TaskPaneHost.LogProcess("Calculate reference planes");
                                PPDetails_TaskPaneHost.LogProcess("Calculate reference planes");

                                if ((RegStatus == true) && (removeId.Count > 0))
                                {
                                    suppressFeature(Doc, assyModel, compName[0], InitialRefPlanes, removeId);

                                }

                            }
                        }

                        if (SelectedRefPlanes.Count != 0)
                        {
                            //analyze intersection
                            boolStatus = planeIntersection(SelectedRefPlanes);

                            if (boolStatus == true)
                            {
                                PPDetails_TaskPaneHost.RegisterToPlaneTable(SelectedRefPlanes);
                                ProcessLog_TaskPaneHost.LogProcess("Add selected planes to the table");
                                PPDetails_TaskPaneHost.LogProcess("Add selected planes to the table");

                            }
                        }

                    }

                    finally
                    {
                        
                    
                    }

                }
            }

        }

        //tools for plane calculator
        #region PlaneCalculator

        public object virtualCentroid { get; set; }

        public Double[] centroid { get; set; }

        //get midpoint between two points
        public double[] getMidPoint(object boxVertices)
        {
            double[] cornerPoint = (double[])boxVertices;
            double[] midPoint = new double[3];

            midPoint[0] = (cornerPoint[0] + cornerPoint[3]) / 2;
            midPoint[1] = (cornerPoint[1] + cornerPoint[4]) / 2;

            if (cornerPoint[2] < cornerPoint[5])
            {
                midPoint[2] = cornerPoint[2];
            }
            else
            {
                midPoint[2] = cornerPoint[5];
            }

            return midPoint;
        }

        //get the plane name
        public bool getPlanes(PartDoc swPart, ref List<AddedReferencePlane> RefPlanes)
        {   
            Feature tmpFeature = null;
            string refName = null;
            
            //get the reference plane feature
            for (int i = 1; i <= RefPlanes.Count(); i++)
            {
                refName = "REF_PLANE" + i.ToString();
                tmpFeature = swPart.FeatureByName(refName);

                if (tmpFeature != null)
                {
                    RefPlanes[i-1].CorrespondFeature = tmpFeature;
                }
            }

            return true;
        }
                
        //get and store the planes normal
        public bool getPlaneNormal(MathUtility swMathUtils, List<Feature> refPlanes, ref List<object> planeNorms)
        {
            MathVector swNormalVector = null;
            MathVector swNormalPlane = null;
            string planeName = null;
            RefPlane swRefPlane = null;
            MathTransform planeTransform = null;
            object worldOrientation = null;
            
            //set the initial canonical orientation in world coordinates
            double[] initOri = new double[3];
            initOri[0] = 0;
            initOri[1] = 0;
            initOri[2] = 1;

            planeNorms = new List<object>();

            //get plane normal
            foreach (Feature tmpFeature in refPlanes)
            {
                planeName = tmpFeature.Name;
                swRefPlane = tmpFeature.GetSpecificFeature();
                worldOrientation = initOri;
                planeTransform = swRefPlane.Transform;
                swNormalVector = swMathUtils.CreateVector(worldOrientation);
                swNormalPlane = swNormalVector.MultiplyTransform(planeTransform);

                //change the normal direction
                swNormalPlane = swNormalPlane.Scale(-1);

                planeNorms.Add(swNormalPlane.ArrayData);
            }

            return true;

        }

        //OVERLOAD getPlaneNormal
        public bool getPlaneNormal(MathUtility swMathUtils, ref List<AddedReferencePlane> refPlanes)
        {
            MathVector swNormalVector = null;
            MathVector swNormalPlane = null;
            
            RefPlane swRefPlane = null;
            MathTransform planeTransform = null;
            object worldOrientation = null;

            //set the initial canonical orientation in world coordinates
            double[] initOri = new double[3];
            initOri[0] = 0;
            initOri[1] = 0;
            initOri[2] = 1;

            //get plane normal
            for (int i = 0; i < refPlanes.Count(); i++ )
            {
                swRefPlane = refPlanes[i].ReferencePlane;
                worldOrientation = initOri;
                planeTransform = swRefPlane.Transform;
                swNormalVector = swMathUtils.CreateVector(worldOrientation);
                swNormalPlane = swNormalVector.MultiplyTransform(planeTransform);

                //change the normal direction
                swNormalPlane = swNormalPlane.Scale(-1);

                refPlanes[i].PlaneNormal = swNormalPlane.ArrayData;
                refPlanes[i].NormalOrientation = _getNormalOrientation(refPlanes[i].PlaneNormal);
            }

            return true;

        }

        //set the name for each normal axis orientation
        private string _getNormalOrientation(Object ThisNormalObject)
        {               
            double[] TmpNormal = { };

            if (ThisNormalObject != null)
            {
                TmpNormal = (double[])ThisNormalObject;

                if (isEqual(TmpNormal[0], 1) && isEqual(TmpNormal[1], 0) && isEqual(TmpNormal[2], 0)) { return "x-plus"; }
                if (isEqual(TmpNormal[0],-1) && isEqual(TmpNormal[1], 0) && isEqual(TmpNormal[2], 0)) { return "x-negative"; }
                if (isEqual(TmpNormal[0], 0) && isEqual(TmpNormal[1], 1) && isEqual(TmpNormal[2], 0)) { return "y-plus"; }
                if (isEqual(TmpNormal[0], 0) && isEqual(TmpNormal[1],-1) && isEqual(TmpNormal[2], 0)) { return "y-negative"; }
                if (isEqual(TmpNormal[0], 0) && isEqual(TmpNormal[1], 0) && isEqual(TmpNormal[2], 1)) { return "z-plus"; }
                if (isEqual(TmpNormal[0], 0) && isEqual(TmpNormal[1], 0) && isEqual(TmpNormal[2],-1)) { return "z-negative"; }

                if (isEqual(TmpNormal[0], 0) && isEqual(TmpNormal[1], 0) && isEqual(TmpNormal[2], 0)) { return "inclined"; }

            }
            
            return "";
        }

        //intersect all the plane using based on their normal
        public bool planeIntersection(List<object> planeNormal, ref List<int> planeMatrixValue)
        {
            //double[] normVectorA, normVectorB = null;
            //Vector normVectorA, normVectorB;
            Vector3D crossProduct;
            Vector3D normVectorA, normVectorB;
            
            planeMatrixValue = new List<int>();            
            int trueIntersection;

            foreach (object VectorA in planeNormal)
            {
                normVectorA = getVector(VectorA);
                trueIntersection = 0;

                foreach (object VectorB in planeNormal)
                {
                    normVectorB = getVector(VectorB);

                    crossProduct = new Vector3D();
                    crossProduct = Vector3D.CrossProduct(normVectorA, normVectorB);

                    if (isIntersect(crossProduct) == true)
                    {
                        trueIntersection += 1;
                    }
                }

                planeMatrixValue.Add(trueIntersection);
            }

            return true;
        }

        //OVERLOAD intersect all the plane using based on their normal
        public bool planeIntersection(List<AddedReferencePlane> SelectedPlanesIn)
        {
            //double[] normVectorA, normVectorB = null;
            //Vector normVectorA, normVectorB;
            Vector3D crossProduct;
            Vector3D normVectorA, normVectorB;

            int trueIntersection;

            foreach (AddedReferencePlane TmpPlaneA in SelectedPlanesIn)
            {
                normVectorA = getVector(TmpPlaneA.PlaneNormal);
                trueIntersection = 0;
                TmpPlaneA.PlaneIntersection = new List<string>();

                foreach (AddedReferencePlane TmpPlaneB in SelectedPlanesIn)
                {
                    if (TmpPlaneA.BodyOwner.Equals(TmpPlaneB.BodyOwner))
                    {
                        normVectorB = getVector(TmpPlaneB.PlaneNormal);

                        crossProduct = new Vector3D();
                        crossProduct = Vector3D.CrossProduct(normVectorA, normVectorB);

                        if (isIntersect(crossProduct) == true)
                        {
                            trueIntersection += 1;
                            TmpPlaneA.PlaneIntersection.Add(TmpPlaneB.name);
                        }
                    }
                    
                }

                TmpPlaneA.Score = trueIntersection;

            }

            return true;
        }

        //get vector
        static Vector3D getVector(object vectorObject)
        {
            double[] vectorDouble = {};
            Vector3D vector;

            vectorDouble = (double[])vectorObject;

            vector = new Vector3D(vectorDouble[0], vectorDouble[1], vectorDouble[2]);

            return vector;
        }

        //compute the distance between the plane and the centroid
        public void setDistance(ModelDoc2 Doc, MathUtility swMathUtils, List<Feature> planeFeatures, List<object> planeNormal,
            Double[] centroid, ref List<double> distance)
        {
            distance = new List<double>();
            double tmpDistance;

            if (planeFeatures != null)
            {
                for (int i = 0; i < planeFeatures.Count; i++)
                {
                    tmpDistance = new double();
                    Face2 tmpFace = getAttachedFace(planeFeatures[i]);

                    //tmpDistance = checkDistance(Doc, planeFeatures[i], planeNormal[i], centroid, swMathUtils);

                    //tmpDistance = checkDistance(Doc, tmpFace, planeNormal[i], centroid, swMathUtils);
                    distance.Add(tmpDistance);
                }
            }

        }

        //OVERLOAD setDistance
        public Boolean setDistance(ModelDoc2 Doc, MathUtility swMathUtils, ref List<AddedReferencePlane> RefPlanes)
        {
            double TmpDistance;
            RefPlane TmpRefPlane;
            object TmpNormal;
            double[] centroid;
            
            if (RefPlanes.Count() > 0)
            {
                for (int i = 0; i < RefPlanes.Count; i++)
                {
                    TmpDistance = new double();
                    TmpRefPlane = RefPlanes[i].ReferencePlane;
                    TmpNormal = RefPlanes[i].PlaneNormal;
                    centroid = RefPlanes[i].AttachedBody.VirtualPoint;

                    TmpDistance = checkDistance(Doc, TmpRefPlane, TmpNormal, centroid, swMathUtils);
                    RefPlanes[i].DistanceFromCentroid = SetZeroTolerance(TmpDistance * 1000);
                    
                }
            }

            return true;

        }

        //remove the trailing zeros
        public Double SetZeroTolerance(Double ThisValue)
        { 
            if (isEqual(ThisValue, 0.000)) { return 0.000; }

            return Math.Round(ThisValue, 3);

        }

        //get face that corresponds to refplane
        public Face2 getAttachedFace(Feature feature)
        {
            if (InitialRefPlanes.Count != 0)
            {
                foreach (AddedReferencePlane CheckRefPlane in InitialRefPlanes)
                {
                    if (CheckRefPlane.name.Equals(feature.Name))
                    {
                        return CheckRefPlane.AttachedFace;
                    }
                }
            }

            return null;
        }

        //check distance between plane and point
        public double checkDistance(ModelDoc2 Doc, RefPlane Plane, object planeNormal, double[] point2Check, MathUtility swMathUtils)
        {
            //reference: http://mathworld.wolfram.com/Point-PlaneDistance.html
            
            MathPoint firstPoint = getPointOnPlane(Plane); //select a random point from reference plane corners
            Double[] TmpFirstPoint = (Double[])firstPoint.ArrayData;
            MathPoint secondPoint = swMathUtils.CreatePoint(point2Check);
            Double[] TmpSecondPoint = (Double[])secondPoint.ArrayData;
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint); //first vector
            Double[] TmpPointsVector = (Double[])pointsVector.ArrayData;
            MathVector pointsVector2 = (MathVector)pointsVector.Scale(-1);
            Double[] TmpPointsVector2 = (Double[])pointsVector2.ArrayData;
            MathVector vectorFromPlane = swMathUtils.CreateVector(planeNormal); //plane vector, aka second vector
            Double[] TmpVectorFromPlane = (Double[])vectorFromPlane.ArrayData;
            Double distance = Math.Abs(vectorFromPlane.Dot(pointsVector2)) / vectorFromPlane.GetLength(); //first dot second and divided by second magnitude
            
            //Double distance = getDistance(Plane);
            
            return distance;
        }

        //OVERLOAD check distance between plane and point
        public double checkDistance(RefPlane Plane, object planeNormal, double[] point2Check, MathUtility swMathUtils)
        {
            //reference: http://mathworld.wolfram.com/Point-PlaneDistance.html

            MathPoint firstPoint = getPointOnPlane(Plane); //select a random point from reference plane corners
            Double[] TmpFirstPoint = (Double[])firstPoint.ArrayData;
            MathPoint secondPoint = swMathUtils.CreatePoint(point2Check);
            Double[] TmpSecondPoint = (Double[])secondPoint.ArrayData;
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint); //first vector
            Double[] TmpPointsVector = (Double[])pointsVector.ArrayData;
            MathVector pointsVector2 = (MathVector)pointsVector.Scale(-1);
            Double[] TmpPointsVector2 = (Double[])pointsVector2.ArrayData;
            MathVector vectorFromPlane = swMathUtils.CreateVector(planeNormal); //plane vector, aka second vector
            Double[] TmpVectorFromPlane = (Double[])vectorFromPlane.ArrayData;
            Double distance = Math.Abs(vectorFromPlane.Dot(pointsVector2)) / vectorFromPlane.GetLength(); //first dot second and divided by second magnitude

            //Double distance = getDistance(Plane);

            return distance;
        }

        //temporary pre-set distance
        public Double getDistance(Feature Plane)
        {
            switch (Plane.Name.ToLower())
            {
                case "#plane1": return 40.0;
                    
                case "#plane2": return 15.0;
                    
                case "#plane3": return 40.0;
                    
                case "#plane4": return 20.0;
                    
                case "#plane5": return 40.0;
                    
                case "#plane6": return 0.0;
                    
                case "#plane7": return 10.0;
                    
                case "#plane8": return 10.0;
                    
                case "#plane9": return 22.5;
                    
            }
                
            return 0.0;
        }

        //add all plane into one list
        public bool registerPlane(List<AddedReferencePlane> ListOfPlanes, List<int> planeRank, List<Feature> planeFeatures, List<object> planeNorms, List<double> distance, ref List<int> removeId,
            MathUtility mathUtils)
        {
            AddedReferencePlane tmpSetPlane;
            bool retSim = false;

            //set the instance for the list that will keep the reference planes
            planeList = new List<_planeProperties>();

            for (int i = 0; i<ListOfPlanes.Count; i++)
            { 
                tmpSetPlane = new AddedReferencePlane();
                tmpSetPlane = ListOfPlanes[i];

                /*
                tmpSetPlane.name = planeFeatures[i].Name;
                tmpSetPlane.rankNum = planeRank[i];
                tmpSetPlane.featureObj = planeFeatures[i];
                tmpSetPlane.planeNormal = planeNorms[i];
                tmpSetPlane.distance = distance[i];
                tmpSetPlane.mark = 0;
                 * */

                //check the similarity (parallel and coincident)
                //if (planeList.Count != 0)
                //{
                    retSim = checkPlaneSim(ListOfPlanes, tmpSetPlane, mathUtils);
                //}

                //add the plane that has to be removed
                if (retSim == true)
                {
                    removeId.Add(i);
                }

                //add only unsimilar plane
                //if (planeList.Count == 0 || retSim == false)
                //{
                //    planeList.Add(tmpSetPlane);
                //}
            }

            return true;
        }

        //OVERLOAD add all plane into one list
        public bool registerPlane(List<AddedReferencePlane> ListOfPlanes, ref List<int> removeId, MathUtility mathUtils)
        {
            AddedReferencePlane tmpSetPlane;
            bool retSim;

            //set the instance for the list that will keep the reference planes
            SelectedRefPlanes = new List<AddedReferencePlane>();　　　//表に出力される参照面
            planeList = new List<_planeProperties>();

            for (int i = 0; i < ListOfPlanes.Count; i++)
            {
                tmpSetPlane = new AddedReferencePlane();
                tmpSetPlane = ListOfPlanes[i];
                retSim = false;

                //process only if the refplane is not on the bounding box
                if (tmpSetPlane.isOnBB == true)
                {
                    if (CheckBBPlaneLocation(tmpSetPlane, mathUtils) == true && tmpSetPlane.SafeDirection == true)
                    {
                        SelectedRefPlanes.Add(tmpSetPlane);
                    }
                    removeId.Add(i);
                }
                else
                {                    
                    //add the reference plane for the first time if the list still empty
                    if (SelectedRefPlanes.Count == 0)
                    {
                        //retSim = checkPlaneSim(ListOfPlanes, tmpSetPlane, mathUtils);

                        //if (retSim == true)
                        //{
                        //    SelectedRefPlanes.Add(tmpSetPlane);
                        //}
                        //else
                        //{
                            //if (tmpSetPlane.CPost == false)
                            if (tmpSetPlane.SafeDirection == true)
                            {
                                SelectedRefPlanes.Add(tmpSetPlane);
                            }
                            else
                            {
                                removeId.Add(i); //add the plane that need to be removed
                            }
                        //}
                    }
                    else
                    {
                        //check the similarity (parallel and coincident)
                        retSim = checkPlaneSim(SelectedRefPlanes, tmpSetPlane, mathUtils);

                        if (retSim == true)
                        {
                            removeId.Add(i); //add the plane that need to be removed
                        }
                        else
                        {
                            //if (tmpSetPlane.CPost == false)
                            if (tmpSetPlane.SafeDirection == true)
                            {
                                SelectedRefPlanes.Add(tmpSetPlane);
                            }
                            else
                            {
                                //retSim = checkPlaneSim(ListOfPlanes, removeId, tmpSetPlane, mathUtils);

                                ////retSim = checkPlaneSim(ListOfPlanes, tmpSetPlane, mathUtils);

                                //if (retSim == true)
                                //{
                                //    SelectedRefPlanes.Add(tmpSetPlane);
                                //}
                                //else
                                //{
                                //    if (tmpSetPlane.Patterns.Type == "MIX")
                                //    {
                                //        if (tmpSetPlane.Patterns.Line <= 2)
                                //        {
                                //            SelectedRefPlanes.Add(tmpSetPlane);
                                //        }
                                //    }

                                //    else
                                //    {
                                        removeId.Add(i); //add the plane that needs to be removed
                                //    }
                                //}
                            }
                        }

                    }
                    
                }

            }

            return true;
        }

        //check plane on bounding box if coincident with virtual point
        public bool CheckBBPlaneLocation(AddedReferencePlane ThisPlane, MathUtility ThisUtils)
        {
            double[] centroid = ThisPlane.AttachedBody.VirtualPoint;
            double distance = checkDistance(ThisPlane.ReferencePlane, ThisPlane.PlaneNormal, centroid, ThisUtils);

            if (isEqual(distance, 0.0) == true)
            {
                return true;
            }
            
            return false;
        }

        //check plane similarity
        public bool checkPlaneSim(List<_planeProperties> planeProperties, _planeProperties plane2Check, MathUtility mathUtils)
        {
            bool normalSim = false;
            bool similarity = false;
            bool sameLevel = false;

            foreach (_planeProperties tmpPlane in planeProperties)
            {
                //check if parallel
                normalSim = checkNormalSim(tmpPlane.planeNormal, plane2Check.planeNormal, mathUtils);

                if (normalSim == true)
                {
                    //check if coincident
                    sameLevel = checkPlaneLevel(tmpPlane.featureObj, tmpPlane.planeNormal, plane2Check.featureObj, mathUtils);

                    if (sameLevel == true)
                    {
                        similarity = true;
                        break;
                    }
                }
                
            }

            //if similar, parallel and coincident, then it should return true
            return similarity;
        }

        //OVERLOAD plane similarity
        public bool checkPlaneSim(List<AddedReferencePlane> ListOfPlanes, AddedReferencePlane plane2Check, MathUtility mathUtils)
        {
            bool normalSim = false;
            bool similarity = false;
            bool sameLevel = false;
            bool NameCheck = false;
            bool BodyCheck = false;

            foreach (AddedReferencePlane tmpPlane in ListOfPlanes)
            {
                NameCheck = plane2Check.name.ToString().Equals(tmpPlane.name.ToString());

                if (NameCheck == false)
                {
                    BodyCheck = plane2Check.AttachedBody.Name.Equals(tmpPlane.AttachedBody.Name);

                    if (BodyCheck == true)
                    {

                        //check if parallel
                        normalSim = checkNormalSim(tmpPlane.PlaneNormal, plane2Check.PlaneNormal, mathUtils);

                        if (normalSim == true)
                        {
                            //check if coincident
                            sameLevel = checkPlaneLevel(tmpPlane.CorrespondFeature, tmpPlane.PlaneNormal, plane2Check.CorrespondFeature, mathUtils);

                            if (sameLevel == true)
                            {
                                similarity = true;
                                break;
                            }
                        }
                    }
                }
            }

            //if similar, parallel and coincident, then it should return true
            return similarity;
        }

        //OVERLOAD plane similarity to avoid deletion because of CPPost true
        public bool checkPlaneSim(List<AddedReferencePlane> ListOfPlanes, List<int> RemoveList, AddedReferencePlane plane2Check, MathUtility mathUtils)
        {
            bool normalSim = false;
            bool similarity = false;
            bool sameLevel = false;
            bool NameCheck = false;
            bool BodyCheck = false;
            AddedReferencePlane tmpPlane;

            foreach (int RemoveIndex in RemoveList)
            {
                tmpPlane = new AddedReferencePlane();
                tmpPlane = ListOfPlanes[RemoveIndex];

                NameCheck = plane2Check.name.ToString().Equals(tmpPlane.name.ToString());

                if (NameCheck == false)
                {
                    BodyCheck = plane2Check.AttachedBody.Name.Equals(tmpPlane.AttachedBody.Name);

                    if (BodyCheck == true)
                    {
                        //check if parallel
                        normalSim = checkNormalSim(tmpPlane.PlaneNormal, plane2Check.PlaneNormal, mathUtils);

                        if (normalSim == true)
                        {
                            //check if coincident
                            sameLevel = checkPlaneLevel(tmpPlane.CorrespondFeature, tmpPlane.PlaneNormal, plane2Check.CorrespondFeature, mathUtils);

                            if (sameLevel == true)
                            {
                                similarity = true;
                                break;
                            }
                        }
                    }
                }
            }

            //if similar, parallel and coincident, then it should return true
            return similarity;
        }

        //check normal direction similarity
        public static bool checkNormalSim(object firstN, object secondN, MathUtility mathUtils)
        {
            double[] firstDouble = null;
            double[] secondDouble = null;
            MathVector firstVector = mathUtils.CreateVector(firstN);
            MathVector secondVector = mathUtils.CreateVector(secondN);

            firstDouble = (double[])firstN;
            secondDouble = (double[])secondN;

            MathVector firstNormalize = firstVector.Normalise();
            MathVector secondNormalize = secondVector.Normalise();

            double dotProduct = firstNormalize.Dot(secondNormalize);

            //check the dotproduct, if equal to 1, means two vectors are in the same direction.

            //if (isEqual(dotProduct, 1) == true | isEqual(dotProduct, -1) == true) { return true; }

            if (isEqual(dotProduct, 1) == true) { return true; }

            return false;

        }

        //check plane level similarity
        public bool checkPlaneLevel(Feature firstFeat, Object firsN, Feature secondFeat, MathUtility mathUtils)
        {
            
            MathPoint firstPoint = getPointOnPlane(firstFeat);
            MathPoint secondPoint = getPointOnPlane(secondFeat);
            MathVector pointsVector = (MathVector)firstPoint.Subtract(secondPoint);

            MathVector firstVector = mathUtils.CreateVector(firsN);

            Double distance = Math.Abs(pointsVector.Dot(firstVector) / firstVector.GetLength());

            //check the distance, if equal with 0, then it should be on the same level
            return isEqual(distance, 0);

        }

        //delete feature from document
        public void suppressFeature(ModelDoc2 Doc, AssemblyDoc assyDoc, Component2 compName, List<AddedReferencePlane> ListOfPlanes, List<int> removeId)
        {   
            Feature tmpFeature = null;
            bool boolStatus;

            if (assyDoc != null)
            {
                //get the part document of the raw material
                ModelDoc2 compModDoc = (ModelDoc2)compName.GetModelDoc2();
                PartDoc swPartDoc = (PartDoc)compModDoc;

                //select the raw material for editing
                boolStatus = compName.Select2(true, 0);
                if (boolStatus == true)
                {
                    assyDoc.EditPart();
                    FeatureManager featMgr = (FeatureManager)Doc.FeatureManager;

                    foreach (int i in removeId)
                    {
                        tmpFeature = swPartDoc.FeatureByName(ListOfPlanes[i].name);
                        tmpFeature.Select2(true, 0);
                        Doc.EditDelete();
                        
                        tmpFeature.SetSuppression2((int)swFeatureSuppressionAction_e.swSuppressFeature, (int)swInConfigurationOpts_e.swThisConfiguration, null);

                        featMgr.UpdateFeatureTree();
                        
                    }
                    
                    assyDoc.EditAssembly();
                }
            }

        }


        //check whether the cross product is zero
        static bool isIntersect(Vector3D vector)
        {
            if (isEqual(vector.X, 0) == true && isEqual(vector.Y,0) == true && isEqual(vector.Z,0) == true)
            {
                return false;
            }

            return true;
        }

        //checking the double is zero to avoid trailing zeros
        static bool isEqual(double value1, double value2)
        {
            return (Math.Abs(Math.Round((value1 - value2), 4)) <= 0.0000001);
            
        }

        //find the centroid position for each reference planes
        public Boolean findCPost(ref List<AddedReferencePlane> AllReferencePlanes)
        {
            Boolean Location = false;
            double[] centroid;
            
            foreach (AddedReferencePlane TmpReferencePlane in AllReferencePlanes)
            {
                centroid = TmpReferencePlane.AttachedBody.VirtualPoint;
                Location = checkCentroidLocation(centroid, TmpReferencePlane, TmpReferencePlane.PlaneNormal);
                TmpReferencePlane.CPost = Location;
            }

            return true;
        }

        //Check centroid location
        public bool checkCentroidLocation(Double[] Centroid, AddedReferencePlane ThisRefPlane, object objPlaneNormal)
        {
            MathVector planeNormal, tmpMathVector = null;
            MathPoint pointOnPlane, pointOnBox = null;
            Double returnLocation = 0;
            MathUtility swMath = (MathUtility)SwApp.GetMathUtility();// = new MathUtility();

            //get random point on the plane
            pointOnPlane = getPointOnPlane(ThisRefPlane);
            double[] TmpPointOnBox = (double[])pointOnPlane.ArrayData;

            //get the centroid of the body box
            pointOnBox = swMath.CreatePoint(Centroid);
            double[] TmpPointBox = (double[])pointOnBox.ArrayData;

            //create vector of plane normal
            planeNormal = swMath.CreateVector(objPlaneNormal);
            double[] TmpPlaneNormal = (double[])planeNormal.ArrayData;

            //find the dot product of normal and the subtracted value from point on plane with point on box
            tmpMathVector = (MathVector)pointOnBox.Subtract(pointOnPlane);
            returnLocation = planeNormal.Dot(tmpMathVector);

            if (returnLocation > 0)
            {
                return true; //the centroid is on the same side with with the plane normal direction
            }
            else
            {
                return false; //the centroid is not on the same side with with the plane normal direction
            }

        }

        //check if point location is on a plane or not
        public static bool CheckPointLocation(Vertex ThisPoint, RefPlane ThisRefPlane, object objPlaneNormal)
        {
            MathVector planeNormal, tmpMathVector = null;
            MathPoint pointOnPlane, pointOnBox = null;
            Double returnLocation = 0;
            MathUtility swMath = (MathUtility)SwApp.GetMathUtility();// = new MathUtility();

            //get random point on the plane
            pointOnPlane = getPointOnPlane(ThisRefPlane);
            double[] TmpPointOnBox = (double[])pointOnPlane.ArrayData;

            //get the centroid of the body box
            pointOnBox = swMath.CreatePoint(ThisPoint.GetPoint());
            double[] TmpPointBox = (double[])pointOnBox.ArrayData;

            //create vector of plane normal
            planeNormal = swMath.CreateVector(objPlaneNormal);
            double[] TmpPlaneNormal = (double[])planeNormal.ArrayData;

            //find the dot product of normal and the subtracted value from point on plane with point on box
            tmpMathVector = (MathVector)pointOnBox.Subtract(pointOnPlane);
            returnLocation = planeNormal.Dot(tmpMathVector);

            if (isEqual(returnLocation, 0) == true)
            {
                return true; //the centroid is on the same side with with the plane normal direction
            }
            else
            {
                return false; //the centroid is not on the same side with with the plane normal direction
            }

        }

        //OVERLOAD check if point location is on a plane or not
        public static bool CheckPointLocation(double[] ThisPoint, RefPlane ThisRefPlane, object objPlaneNormal)
        {
            MathVector planeNormal, tmpMathVector = null;
            MathPoint pointOnPlane, pointOnBox = null;
            Double returnLocation = 0;
            MathUtility swMath = (MathUtility)SwApp.GetMathUtility();// = new MathUtility();

            //get random point on the plane
            pointOnPlane = getPointOnPlane(ThisRefPlane);
            double[] TmpPointOnBox = (double[])pointOnPlane.ArrayData;

            //get the centroid of the body box
            pointOnBox = swMath.CreatePoint(ThisPoint);
            double[] TmpPointBox = (double[])pointOnBox.ArrayData;

            //create vector of plane normal
            planeNormal = swMath.CreateVector(objPlaneNormal);
            double[] TmpPlaneNormal = (double[])planeNormal.ArrayData;

            //find the dot product of normal and the subtracted value from point on plane with point on box
            tmpMathVector = (MathVector)pointOnBox.Subtract(pointOnPlane);
            returnLocation = planeNormal.Dot(tmpMathVector);

            if (isEqual(returnLocation, 0) == true)
            {
                return true; //the centroid is on the same side with with the plane normal direction
            }
            else
            {
                return false; //the centroid is not on the same side with with the plane normal direction
            }

        }

        public static Boolean FindTheOuter(ref List<AddedReferencePlane> ListOfReferencePlanes)
        {
            
            //find from the x-plus
            FindOrientation(ref ListOfReferencePlanes, "x-plus");
            
            //find from the x-negative
            FindOrientation(ref ListOfReferencePlanes, "x-negative");           

            //find from the y-plus
            FindOrientation(ref ListOfReferencePlanes, "y-plus");

            //find from the y-negative
            FindOrientation(ref ListOfReferencePlanes, "y-negative");

            //find from the z-plus
            FindOrientation(ref ListOfReferencePlanes, "z-plus");

            //find from the z-negative
            FindOrientation(ref ListOfReferencePlanes, "z-negative");                        

            return true;
        }

        //analyze the x, y, z direction
        private static bool AnalyzeDirection(ref List<AddedReferencePlane> ListOfRefPlanes)
        {
            bool boolStatus = false;

            boolStatus = AnalyzeDirection(ref ListOfRefPlanes, "x");
            boolStatus = AnalyzeDirection(ref ListOfRefPlanes, "y");
            boolStatus = AnalyzeDirection(ref ListOfRefPlanes, "z");

            if (boolStatus == false) { return false; }

            return true;
        }

        //analyze particular direction
        private static bool AnalyzeDirection(ref List<AddedReferencePlane> ListOfRefPlanes, string ThisAxis)
        {   
            int RefAxis = WhichAxis(ThisAxis);

            //query the reference plane by the lowest order of plane (ascending)
            var PlaneOrderByAscending = from Plane in ListOfRefPlanes
                                        where (Plane.NormalOrientation.Contains(ThisAxis) && Plane.isOnBB == false)
                                        orderby Plane.Location(RefAxis) ascending
                                        select Plane;

            //do the comparison of planes
            if (PlaneOrderByAscending.Count() > 1)
            {
                foreach (AddedReferencePlane PlaneA in PlaneOrderByAscending)
                {
                    if (PlaneA.NormalOrientation.Contains("plus"))
                    {
                        foreach (AddedReferencePlane PlaneB in PlaneOrderByAscending)
                        {
                            if ((PlaneB.NormalOrientation.Contains("negative")) && 
                                (PlaneA.Location(RefAxis) < PlaneB.Location(RefAxis)))
                            {
                                if (IsOverlapped(PlaneA, PlaneB, RefAxis) == true)
                                {
                                    PlaneA.SafeDirection = false;
                                    PlaneB.SafeDirection = false;
                                }
                            
                            }
                        }
                    }
                }
            }

            //query the reference plane by the lowest order of plane (descending)
            var PlaneOrderByDescending = from Plane in ListOfRefPlanes
                                         where (Plane.NormalOrientation.Contains(ThisAxis) && Plane.isOnBB == false)
                                         orderby Plane.Location(RefAxis) descending
                                         select Plane;

            //do the comparison of planes
            if (PlaneOrderByDescending.Count() > 1)
            {
                foreach (AddedReferencePlane PlaneA in PlaneOrderByDescending)
                {
                    if (PlaneA.NormalOrientation.Contains("negative"))
                    {
                        foreach (AddedReferencePlane PlaneB in PlaneOrderByDescending)
                        {
                            if ((PlaneB.NormalOrientation.Contains("plus")) &&
                                (PlaneA.Location(RefAxis) > PlaneB.Location(RefAxis)))
                            {
                                if (IsOverlapped(PlaneA, PlaneB, RefAxis) == true)
                                {
                                    PlaneA.SafeDirection = false;
                                    PlaneB.SafeDirection = false;
                                }

                            }
                        }
                    }
                }
            }
            
            return true;
        }
        
        //set the list of TRV
        public static List<int> MainAxisLeveling;

        //check the number of trv by using planes which the order of x, y, z axes
        private bool CheckTRVbodies(ModelDoc2 ThisDoc, AssemblyDoc ThisAssyDoc, out int SelectedAxis)
        {
            //MainAxisLeveling = new List<int>() { 0, 0, 0 };

            bool PlanesRetrieval = false;
            List<AddedReferencePlane> SelectedPlanes = null;
            int NumberOfBodies = -1;
            Array BodyArray;
            SelectedAxis = -1;

            //open the component for editing
            ThisAssyDoc.EditPart();

            //do the comparison for x, y, z
            for (int i=0; i<3; i++)
            {
                if (MainAxisLeveling[i] != 1)
                {
                    SelectedPlanes = new List<AddedReferencePlane>();

                    //get the planes
                    switch(i)
                    {                    
                        case 0:
                            PlanesRetrieval = GetParallelPlanes(InitialRefPlanes, "x",ref SelectedPlanes);
                            break;
                        case 1:
                            PlanesRetrieval = GetParallelPlanes(InitialRefPlanes, "y", ref SelectedPlanes);
                            break;
                        case 2:
                            PlanesRetrieval = GetParallelPlanes(InitialRefPlanes, "z", ref SelectedPlanes);
                            break;
                    }

                    if (PlanesRetrieval == true)
                    {   
                        //select all planes
                        foreach (AddedReferencePlane ThisPlane in SelectedPlanes)
                        {
                            ThisPlane.CorrespondFeature.Select2(true, 0);
                        }
                        
                        //split and collect the body
                        BodyArray = (Array)ThisDoc.FeatureManager.PreSplitBody();

                        //compare the cutting result with previous cut
                        if (BodyArray.Length < NumberOfBodies)
                        {
                            NumberOfBodies = BodyArray.Length;
                            SelectedAxis = i;
                        }
                        
                    }
                }
            }

            if (SelectedAxis != -1) 
            {
                MainAxisLeveling[SelectedAxis] = 1;
                return true; 
            }

            return false;
        }

        //get planes based on the required axis
        private static bool GetParallelPlanes(List<AddedReferencePlane> ListOfRefPlanes, string ThisAxis, ref List<AddedReferencePlane> ThisPlaneCollection)
        {
            int RefAxis = WhichAxis(ThisAxis);
            AddedReferencePlane PreviousAddedPlane = new AddedReferencePlane();

            //query the reference plane by the lowest order of plane (ascending)
            var PlaneOrderByAscending = from Plane in ListOfRefPlanes
                                        where (Plane.NormalOrientation.Contains(ThisAxis) && Plane.isOnBB == false)
                                        orderby Plane.Location(RefAxis) ascending
                                        select Plane;

            //do the comparison of planes
            if (PlaneOrderByAscending.Count() != 0)
            {
                foreach (AddedReferencePlane ThisPlane in PlaneOrderByAscending)
                {
                    if (PreviousAddedPlane == null || (isEqual(PreviousAddedPlane.Location(RefAxis), ThisPlane.Location(RefAxis)) == false))
                    {
                        ThisPlaneCollection.Add(ThisPlane);
                        PreviousAddedPlane = ThisPlane;
                    }
                }
            }

            if (ThisPlaneCollection.Count() != 0)
            {
                return true;
            }
           
            ThisPlaneCollection = null;
            return false;

        }

        //check whether two reference planes are overlapped (PlaneA overlaps PlaneB)
        public static bool IsOverlapped(AddedReferencePlane PlaneA, AddedReferencePlane PlaneB, int RefAxis)
        {
            bool MaxPoint = false;
            bool MinPoint = false;

            switch (RefAxis)
            { 
                case 0: //evaluate the (y,z) point
                    MaxPoint = CompareMaxPoint(PlaneA.GetCornerValue(1), PlaneA.GetCornerValue(2),
                        PlaneB.GetCornerValue(1), PlaneB.GetCornerValue(2));
                    MinPoint = CompareMinPoint(PlaneA.GetCornerValue(4), PlaneA.GetCornerValue(5),
                        PlaneB.GetCornerValue(4), PlaneB.GetCornerValue(5));

                    break;

                case 1: //evaluate the (x,z) point
                    MaxPoint = CompareMaxPoint(PlaneA.GetCornerValue(0), PlaneA.GetCornerValue(2),
                        PlaneB.GetCornerValue(0), PlaneB.GetCornerValue(2));
                    MinPoint = CompareMinPoint(PlaneA.GetCornerValue(3), PlaneA.GetCornerValue(5),
                        PlaneB.GetCornerValue(3), PlaneB.GetCornerValue(5));

                    break;

                case 2: //evaluate the (x,y) point
                    MaxPoint = CompareMaxPoint(PlaneA.GetCornerValue(0), PlaneA.GetCornerValue(1),
                        PlaneB.GetCornerValue(0), PlaneB.GetCornerValue(1));
                    MinPoint = CompareMinPoint(PlaneA.GetCornerValue(3), PlaneA.GetCornerValue(4),
                        PlaneB.GetCornerValue(3), PlaneB.GetCornerValue(4));

                    break;
            }

            if (MaxPoint == true && MinPoint == true) { return true; }

            return false;
        }

        //compare min point between two points
        public static bool CompareMinPoint(Double MinA, Double MinB, Double MinA2, Double MinB2)
        {   
            if ((MinA >= MinA2) || (MinB >= MinB2)) { return true; }

            return false;
        }

        //compare max point between two points
        public static bool CompareMaxPoint(Double MaxA, Double MaxB, Double MaxA2, Double MaxB2)
        {
            if ((MaxA <= MaxA2) || (MaxB <= MaxB2)) { return true; }

            return false;
        }

        //return the int name of and axis 0:x, 1:y, 2:z
        private static int WhichAxis(string ThisAxis)
        {
            if (ThisAxis.Contains("x")) return 0;
            if (ThisAxis.Contains("y")) return 1;
            if (ThisAxis.Contains("z")) return 2;

            return -1;
        }


        //find similar reference plane on each axis orientation
        private static void FindOrientation(ref List<AddedReferencePlane> ThisRefList, string ThisOrientation)
        {
            int ThisAxis = WhichAxis(ThisOrientation);

            //var PlaneQuery = from Plane in ThisRefList
            //                 where Plane.NormalOrientation == ThisOrientation && Plane.isOnBB == false
            //                 orderby Plane.DistanceFromCentroid descending
            //                 select Plane.name;

            var PlaneQuery = from Plane in ThisRefList
                             where Plane.NormalOrientation.Contains(ThisOrientation) && Plane.isOnBB == false
                             orderby Plane.Location(ThisAxis) descending
                             select Plane.name;

            if (PlaneQuery.Count() != 0)
            {
                for (int i = 0; i < ThisRefList.Count(); i++)
                {
                    if (ThisRefList[i].name.Equals(PlaneQuery.First()))
                    {
                        ThisRefList[i].isOuter = true;
                        break;
                    }

                }
            }
        }

        //analyze the datum
        private bool AnalyzeDatum()
        {
            PPX_Datum ThisDatum;
            List<PPX_Datum> TmpDatumList = new List<PPX_Datum>();
            List<int> AvailableIndex = new List<int>();
            
            //list all datum from plane
            foreach (AddedReferencePlane ThisRefPlane in InitialRefPlanes)
            {
                if (ThisRefPlane.Tolerance != null)
                {
                    if (ThisRefPlane.Tolerance.datum != "")
                    {
                        ThisDatum = new PPX_Datum();
                        ThisDatum.Name = ThisRefPlane.name;
                        ThisDatum.Index = Convert.ToInt32(ThisRefPlane.GetToleranceValue(1));
                        ThisDatum.Type = "plane";
                        ThisDatum.BodyOwner = ThisRefPlane.BodyOwner;

                        TmpDatumList.Add(ThisDatum);
                    }
                }
            }

            //list all datum from profile
            foreach (AddedReferenceProfile ThisRefProfile in InitialRefProfiles)
            {
                if (ThisRefProfile.Tolerance != null)
                {
                    if (ThisRefProfile.Tolerance.datum != "")
                    {
                        ThisDatum = new PPX_Datum();
                        ThisDatum.Name = ThisRefProfile.name;
                        ThisDatum.Index = Convert.ToInt32(ThisRefProfile.GetToleranceValue(1));
                        ThisDatum.Type = "profile";
                        ThisDatum.BodyOwner = ThisRefProfile.BodyOwner;

                        TmpDatumList.Add(ThisDatum);
                    }
                }
            }

            //sort the datum from the smallest index
            List<PPX_Datum> SortedDatum = TmpDatumList.OrderBy(Plane => Plane.Index).ToList();
            AddedReferenceProfile TmpProfile;

            //list all datum from plane with profile
            foreach (AddedReferencePlane ThisRefPlane in InitialRefPlanes)
            {   
                if (ThisRefPlane.Profiles != null)
                {
                    if (ThisRefPlane.Profiles.Count() == 1)
                    {
                        TmpProfile = ThisRefPlane.Profiles.First();

                        if (TmpProfile.Main == true)
                        {
                            var Items = from item in SortedDatum
                                        where (item.Name == ThisRefPlane.name)
                                        select item;

                            if (Items.Count() == 0)
                            {
                                ThisDatum = new PPX_Datum();
                                ThisDatum.Name = ThisRefPlane.name;
                                ThisDatum.Index = SortedDatum.Last().Index + 1;
                                ThisDatum.Type = "plane";
                                ThisDatum.BodyOwner = ThisRefPlane.BodyOwner;

                                SortedDatum.Add(ThisDatum);
                            }
                        }
                    }
                }
            }

            //set new instance of datum collection and copy all sorted datum
            RefDatum = new List<PPX_Datum>();
            CopyDatum(ref RefDatum, SortedDatum);
            
            return true;
        }

        //copy all the datum to a new place
        private void CopyDatum(ref List<PPX_Datum> Target, List<PPX_Datum> Source)
        {
            PPX_Datum NewDatum;

            foreach (PPX_Datum ThisDatum in Source)
            { 
                NewDatum = new PPX_Datum();
                NewDatum.Name = ThisDatum.Name;
                NewDatum.Index = ThisDatum.Index;
                NewDatum.Type = ThisDatum.Type;
                NewDatum.BodyOwner = ThisDatum.BodyOwner;

                Target.Add(NewDatum);
            }
        }

        //analyze the profile surface for perpendicularity
        private bool AnalyzeProfile(ModelDoc2 MasterDoc, ModelDoc2 ThisCompDoc, AssemblyDoc ThisAssyDoc, 
            Component2 ThisComp, SelectionMgr ThisComSelMgr, ref SelectData ThisSelectData)
        {
            string ReqDatum;
            string [] ProfileTol;
            Face ThisTargetFace = null;
            Edge ThisProfileEdge = null;
            Edge ThisSurfaceEdge;
            object[] SurfaceEdges;
            List<string> DatumTarget = new List<string>();
            bool SelectStatus = false;
            
            List<AddedReferencePlane> RefPlaneTargets = new List<AddedReferencePlane>();
            List<AddedReferenceProfile> RefProfile = new List<AddedReferenceProfile>();
            bool StatusFound = false;

            //find the surface which will be used to extend the profile
            foreach (AddedReferenceProfile ThisProfile in InitialRefProfiles)
            {
                if (ThisProfile.Tolerance != null)
                {
                    //check the datum requirement
                    ReqDatum = ThisProfile.Tolerance.perpendicularity;

                    //only check for main profile
                    if (ReqDatum != "" && ThisProfile.Main == true)
                    {
                        ProfileTol = ReqDatum.Split(' ');

                        //get the required datum and collect its name
                        if (ProfileTol.Count() == 1)
                        {
                            DatumTarget.Add(ProfileTol.First());
                            foreach (PPX_Datum ThisDatum in RefDatum)
                            {
                                StatusFound = false;

                                if (ThisDatum.BodyOwner.Equals(ThisProfile.BodyOwner) && ThisDatum.Name != ThisProfile.name)
                                {
                                    if (ThisDatum.Index == Convert.ToInt32(ProfileTol.First()))
                                    {
                                        if (ThisDatum.Type == "plane")
                                        {
                                            //find the refplane
                                            foreach (AddedReferencePlane ThisPlane in InitialRefPlanes)
                                            {
                                                if (ThisPlane.name.Equals(ThisDatum.Name))
                                                {
                                                    //add the refplane as targets 
                                                    //and also collect the corresponding surface from profile

                                                    RefPlaneTargets.Add(ThisPlane);
                                                    RefProfile.Add(ThisProfile);
                                                    StatusFound = true;
                                                    break;
                                                    ;
                                                }
                                            }

                                        }
                                    }
                                }

                                if (StatusFound == true) { break; }
                            }
                        }

                    }
                }
            }

            SelectStatus = ThisComp.Select2(true, 0);

            ThisAssyDoc.EditPart();

            Feature ThisFeature;
            object[] ObjBodies;
            Body2 SelectedBody = null;
            
            //select the surface body

            for (int i = 1; i < RefProfile.Count(); i++)
            {
                ThisFeature = ThisCompDoc.FirstFeature();
                while (ThisFeature != null)
                {
                    if (ThisFeature.GetTypeName2() == "SurfaceBodyFolder")
                    {
                        SelectedBody = null;
                        BodyFolder ThisBodyFolder = ThisFeature.GetSpecificFeature2();
                        ObjBodies = (object[])ThisBodyFolder.GetBodies();
                        foreach (Body2 ThisBody in ObjBodies)
                        {
                            if (RefProfile[i].ReferenceSurface.SurfaceName.Contains(ThisBody.Name))
                            {
                                SelectStatus = MasterDoc.Extension.SelectByID2(RefProfile[i].ReferenceSurface.SurfaceName, "SURFACEBODY", 0, 0, 0, true, 0, null, 0);
                                SelectedBody = ThisComSelMgr.GetSelectedObject5(1);
                                //SelectStatus = ThisBody.Select2(true, ThisSelectData);
                                //int DisplayStatus = ThisBody.Display3(ThisComp, 150, (int)swTempBodySelectOptions_e.swTempBodySelectable);
                                //SelectedBody = ThisBody;
                                break;
                            }
                        }
                    }

                    if (SelectedBody != null) { break; }
                    ThisFeature = ThisFeature.GetNextFeature();
                }


                //select the edge for extension

                object[] ObjEdgesBody = (object[])SelectedBody.GetEdges();

                foreach (Edge ThisEdge in ObjEdgesBody)
                {
                    Curve ThisEdgeCurve = (Curve)ThisEdge.GetCurve();

                    if (RefProfile[i].CurveType == "circle")
                    {
                        if (CircleSimilarity(RefProfile[i].CurveParam, (double[])ThisEdgeCurve.CircleParams) == true)
                        {
                            ThisProfileEdge = ThisEdge;
                            break;
                        }
                    }

                    if (RefProfile[i].CurveType == "ellipse")
                    {
                        if (EllipseSimilarity(RefProfile[i].CurveParam, (double[])ThisEdgeCurve.GetEllipseParams()) == true)
                        {
                            ThisProfileEdge = ThisEdge;
                            break;
                        }
                    }

                    if (RefProfile[i].CurveType == "bcurve")
                    {
                        double[] ThisCurveParam = (double[])GetCurveMaxTess(ThisEdgeCurve);

                        if (BCurveSimilarity(RefProfile[i].CurveParam, ThisCurveParam) == true)
                        {
                            ThisProfileEdge = ThisEdge;
                            break;
                        }
                    }

                }

                //get the similar edge on the corresponding solid body

                Entity ThisProfileEntity = (Entity)ThisProfileEdge;
                ThisProfileEntity = ThisProfileEntity.GetSafeEntity();
                
                ThisFeature = ThisCompDoc.FirstFeature();
                while (ThisFeature != null)
                {
                    if (ThisFeature.GetTypeName2() == "SolidBodyFolder")
                    {   
                        BodyFolder ThisBodyFolder = ThisFeature.GetSpecificFeature2();
                        ObjBodies = (object[])ThisBodyFolder.GetBodies();
                        foreach (Body2 ThisBody in ObjBodies)
                        {
                            if (RefProfile[i].BodyOwner.Contains(ThisBody.Name))
                            {
                                StatusFound = false;
                                object[] ObjFacesFromBody = (object[])ThisBody.GetFaces();
                                foreach (Face2 ThisFace in ObjFacesFromBody)
                                {
                                    if (CheckFaceSim(ThisFace, RefPlaneTargets[i].AttachedFace) == true)
                                    {
                                        ThisTargetFace = (Face) ThisFace;
                                        StatusFound = true;
                                        break;
                                    }
                                }
                            }

                            if (StatusFound == true) { break; }

                        }
                    }

                    if (StatusFound == true) { break; }
                    ThisFeature = ThisFeature.GetNextFeature();
                }

                //create a surface_offset from the target face

                Entity ThisTargetFaceEntity = (Entity) ThisTargetFace;
                ThisTargetFaceEntity = ThisTargetFaceEntity.GetSafeEntity();
                SelectStatus = ThisTargetFaceEntity.Select4(false, ThisSelectData);

                ThisCompDoc.InsertOffsetSurface(0.0, false);
                Feature NewFeature = (Feature)ThisCompDoc.Extension.GetLastFeatureAdded();
                object[] ObjFaces = (object[])NewFeature.GetFaces();
                Face2 NewFace = (Face2)ObjFaces.FirstOrDefault();
                ThisTargetFaceEntity = (Entity)NewFace;
                ThisTargetFaceEntity = ThisTargetFaceEntity.GetSafeEntity();
                

                SelectStatus = MasterDoc.Extension.SelectByID2(NewFeature.Name + "@" + Path.GetFileNameWithoutExtension(ThisCompDoc.GetPathName())+ "@"+ Path.GetFileNameWithoutExtension(MasterDoc.GetPathName()), "SURFACEBODY", 0, 0, 0, false, 0, null, 0);
                SelectStatus = ThisTargetFaceEntity.Select4(false, ThisSelectData);
                SelectStatus = ThisProfileEntity.Select4(true, ThisSelectData);

                object[] ThisEdges = new object[1] { ThisProfileEdge };
                //Body2 NewSurfaceBody = SelectedBody.ExtendSurface(ThisEdges, false, 2, 0, null, NewFace);

                //MasterDoc.InsertExtendSurface(false, 2, 0);

                //try to extend exceed the selected surface
                ThisCompDoc.InsertExtendSurface(false, 0, 0.030);

                //

            }


            return true;
        }

        //strorage for all datum
        public static List<PPX_Datum> RefDatum;
    
        //storage for all generated planes *OLD ONE
        public static List<_planeProperties> planeList;

        //save only the selected reference planes *NEW ONE
        public static List<AddedReferencePlane> SelectedRefPlanes;

        public static List<int> Datum;

        #endregion

        //TRV feature
        public void TRVfeature()
        {
            ProcessLog_TaskPaneHost.LogProcess("Calculate machining processes");
            PPDetails_TaskPaneHost.LogProcess("Calculate machining processes");

            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;

            MainView = Path.GetFileNameWithoutExtension(Doc.GetPathName());
            
            int docType = (int)Doc.GetType();
            bool boolStatus = false;

            if (Doc.GetType() == 2)
            {
                AssemblyDoc assyModel = (AssemblyDoc)Doc;
                if (assyModel.GetComponentCount(false) != 0)
                {
                    List <_planeProperties> samplePlane = new List<_planeProperties>();
                    samplePlane = planeList;
                    int numPlane = samplePlane.Count;

                    FeatureManager docFeatureMgr = Doc.FeatureManager;
                    object featArr = docFeatureMgr.GetFeatures(false);
                    List<Feature> PlaneFeatures = null;

                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    if (compName == null) 
                    {
                        //define the raw material and product
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);
                        PPDetails_TaskPaneHost.getCompName(ref compName);
                    }
                    
                    boolStatus = compName[0].Select2(true, 0);

                    TraverseComponentFeatures(compName[0], ref PlaneFeatures);

                    /*PlaneFeatures = new List<Feature>();
                    //collect the feature of reference planes
                    foreach (AddedReferencePlane TmpPlane in SelectedRefPlanes)
                    {
                        PlaneFeatures.Add(TmpPlane.CorrespondFeature);
                    }
                     * */

                    //initMachiningPlan();

                    AddedReferencePlane ParentPlane = null;
                    List<AddedReferencePlane> ParentList = null;
                    List<AddedReferencePlane> RelativeList = null;
                    RemovalList = new List<RemovedBody>();

                    //int Errors = 0;
                    //int Warnings = 0;

                    //String CompPathName = compName[0].GetPathName();
                    //iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
                    //ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                    //    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

                    //calculate the total volume of TRV
                    Double MainVolume = CalMainVolume(compName[0]);

                    traversePlanes(Doc, assyModel, compName[0], PlaneFeatures, ParentPlane, ParentList, RelativeList, ref RemovalList);
                    //traversePlanes(Doc, assyModel, compName[0], CompDocumentModel, PlaneFeatures, ParentPlane, ParentList, RemovalList);

                    ////Close the parallel opened document and make the document to be visible again
                    //iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
                    //iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);

                    if (MachiningPlanList != null)
                    {
                        iSwApp.SendMsgToUser("Number of process: " + MachiningPlanList.Count.ToString());

                        boolStatus = compName[0].Select2(true, 0);
                        assyModel.EditPart();

                        Double RemovedVolume = CalRemovedVolume(MachiningPlanList);
                        Double RemovedPercentage = Math.Round(((RemovedVolume / MainVolume) * 100),2);

                        PPDetails_TaskPaneHost.RegisterToTree(MachiningPlanList);
                        ProcessLog_TaskPaneHost.LogProcess("Add machining plans to the tree");
                        ProcessLog_TaskPaneHost.LogProcess("Only " + RemovedPercentage.ToString() + 
                            " % (" + RemovedVolume.ToString() + "/" + MainVolume.ToString() + ") has been removed");
                        
                        PPDetails_TaskPaneHost.LogProcess("Add machining plans to the tree");
                        PPDetails_TaskPaneHost.LogProcess("Only " + RemovedPercentage.ToString() +
                            " % (" + RemovedVolume.ToString() + "/" + MainVolume.ToString() + ") has been removed");
                    }

                    
                }
            }

            //iSwApp.SendMsgToUser("this is from TRV feature");
        }

        //tools for calculating TRV feature
        #region TRV Feature

        public static List<MachiningPlan> MachiningPlanList;

        public static List<RemovedBody> RemovalList;

        public static List<AddedReferencePlane> PlaneListByScore;

        public static List<AddedReferencePlane> PlaneListByDistance;

        //calculate main volume
        public static double CalMainVolume(Component2 ThisComp)
        {
            
            //check multi TRV
            Array InitialBodyList = null;
            Object InitialBodyInfo = null;
            InitialBodyList = (Array)ThisComp.GetBodies3((int)swBodyType_e.swSolidBody, out InitialBodyInfo);

            Double MaterialProperty = iSwApp.GetUserPreferenceDoubleValue((int)swUserPreferenceDoubleValue_e.swMaterialPropertyDensity);
            Double[] MassProperty = null;

            //calculate the total volume of TRV
            Double MainVolume = 0.0;

            foreach (Body2 ThisBody in InitialBodyList)
            {
                MassProperty = (Double[])ThisBody.GetMassProperties(MaterialProperty);
                MainVolume = MainVolume + MassProperty[3];
            }

            return Math.Round(MainVolume * 1000000000, 2);
        }

        //calculate the volume which has been removed
        public static double CalRemovedVolume(List<MachiningPlan> ThisMPList)
        {
            Double RemovedVolume = 0.0;

            if (ThisMPList.Count() != 0)
            {
                foreach (MachiningPlan ThisBody in ThisMPList)
                {
                    if (ThisBody.MachiningVolume > RemovedVolume)
                    {
                        RemovedVolume = ThisBody.MachiningVolume;
                    }
                }
            }
            
            return Math.Round(RemovedVolume, 2);
        }

        //initiate the traverse with the first feature
        public static void TraverseComponentFeatures(Component2 swComp, ref List<Feature> planeList)
        {
            Feature swFeat;

            swFeat = (Feature)swComp.FirstFeature();
            planeList = new List<Feature>();

            TraverseFeatures(swFeat, ref planeList);

        }
        
        //traverse to get previously created planes
        public static void TraverseFeatures(Feature swFeat, ref List<Feature> planeList)
        {

            while ((swFeat != null))
            {

                if (swFeat.Name.Contains("REF_PLANE") && swFeat.GetTypeName2().Equals("RefPlane"))
                {
                    planeList.Add(swFeat);
                }
                
                swFeat = (Feature)swFeat.GetNextFeature();

            }

        }

        //copy the AddedReferencePlane
        public void CopyReference(ref List<AddedReferencePlane> TargetList, List<AddedReferencePlane> SourceList)
        {
            AddedReferencePlane TargetItem;
            foreach (AddedReferencePlane Source in SourceList)
            {
                TargetItem = new AddedReferencePlane();
                TargetItem.name = Source.name;
                TargetItem.ReferencePlane = Source.ReferencePlane;
                TargetItem.AttachedFace = Source.AttachedFace;
                TargetItem.CorrespondFeature = Source.CorrespondFeature;
                TargetItem.RankByDistance = Source.RankByDistance;
                TargetItem.DistanceFromCentroid = Source.DistanceFromCentroid;
                TargetItem.PlaneIntersection = Source.PlaneIntersection;
                TargetItem.Score = Source.Score;
                TargetItem.PlaneNormal = Source.PlaneNormal;
                TargetItem.CPost = Source.CPost;
                TargetItem.MarkingOpt = Source.MarkingOpt;
                TargetItem.Remark = Source.Remark;
                TargetItem.Possibility = Source.Possibility;
                TargetItem.MaxMinValue = Source.MaxMinValue;
                TargetItem.IsPlanar = Source.IsPlanar;
                TargetItem.isOuter = Source.isOuter;
                TargetItem.isOnBB = Source.isOnBB;
                TargetItem.NormalOrientation = Source.NormalOrientation;
                TargetItem.Patterns = Source.Patterns;
                TargetItem.BodyOwner = Source.BodyOwner;
                TargetItem.AttachedBody = Source.AttachedBody;
                TargetItem.Tolerance = Source.Tolerance;
                TargetItem.ReferenceSurface = Source.ReferenceSurface;
                TargetItem.SafeDirection = Source.SafeDirection;
                
                TargetList.Add(TargetItem);
            }
        }

        //traverse planes
        public void traversePlanes(ModelDoc2 Doc, AssemblyDoc assyModel, Component2 swComp, List<Feature> featureList, 
            AddedReferencePlane ParentPlane, List<AddedReferencePlane> ListOfParentPlanes, List<AddedReferencePlane> ListOfRelativePlanes,
            ref List<RemovedBody> PreviousRemoval)
        {
            Feature SelectedFeature = null;
            bool boolStatus;
            int index = 0;
            int SplitCounter = 0;
            int Errors = 0;
            int Warnings = 0;
            
            List<Feature> ListOfParentFeature = null;

            //sort the selected Reference plane by its score
            List<AddedReferencePlane> SortedListA = 
                SelectedRefPlanes.OrderByDescending(Plane => Plane.Score).ThenByDescending(Plane => Plane.DistanceFromCentroid).ToList();

            PlaneListByScore = new List<AddedReferencePlane>();
            CopyReference(ref PlaneListByScore, SortedListA);

            //sort the selected Reference plane by its score
            List<AddedReferencePlane> SortedListB =
                SelectedRefPlanes.OrderByDescending(Plane => Plane.DistanceFromCentroid).ToList();

            PlaneListByDistance = new List<AddedReferencePlane>();
            CopyReference(ref PlaneListByDistance, SortedListB);

            //sort the selected Reference plane by its axis
            List<AddedReferencePlane> SortedListC =
                SelectedRefPlanes.OrderByDescending(Plane => Plane.DistanceFromCentroid).ToList();
            
            //set parent and child status to manage the iteration
            Boolean ChildStatus;
            if (ParentPlane == null)
            {
                SelectedFeature = getThePlane(ref PlaneListByScore, PlaneListByDistance, featureList, 
                    ref index, ListOfParentPlanes, ListOfRelativePlanes);
                ChildStatus = false;
            }
            else
            {
                ChildStatus = true;
                SelectedFeature = getSelectedPlane(ParentPlane, featureList);
                index = getPlaneIndex(ParentPlane, PlaneListByScore);
                ListOfParentFeature = new List<Feature>();
            }

            //make the loaded document to be invisble
            String CompPathName = swComp.GetPathName();
            iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document
            
            while (SelectedFeature != null)
            {
                assyModel.EditPart();

                boolStatus = SelectedFeature.Select2(true, 0);
                //int longstatus = 0;
                
                //split and collect the body
                List<int> DeleteThisBody = null;
                Array bodyArray = null;
                bodyArray = (Array)Doc.FeatureManager.PreSplitBody();
                
                //check if there are body 
                if (bodyArray == null)
                {
                    if (ChildStatus == true)
                    {
                        ChildStatus = false;
                    }
                    else
                    {
                        setMarkOnPlane(ref PlaneListByScore, index, "EMPTYSPLIT"); //mark the plane because no bodies was collected
                        //set the continnuity to be false
                        PlaneListByScore[index].Possibility = false;
                    }
                }
                else
                {  
                    string[] bodyNames = new string[bodyArray.Length]; //set the name
                    
                    Feature SplitFeature = null;
                    Feature DeleteFeature = null;

                    SplitFeature = SplitAndSaveBody(Doc, swComp, SelectedFeature, bodyArray, ref bodyNames);

                    if (SplitFeature != null)
                    {
                        try
                        {
                            Object BodiesInfo = null;
                            
                            //List<double> Volume = new List<double>();
                            //List<string> BodyToDelete = null;

                            List<RemovalBody> RemoveThisBody = null;
                            List<AddedReferencePlane> RelativePlane = null;

                            DeleteThisBody = new List<int>();
                            SplitCounter++; //add the split counter after successful splitting process

                            bodyArray = null;
                            bodyArray = (Array) swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                            //select body that needs to be deleted from the model
                            //bool Status = SelectBodyToDelete(CompDocumentModel, bodyArray, ref BodyToDelete, index, ref Volume);

                            bool Status = SelectBodyToDelete(CompDocumentModel, bodyArray, index, ref RemoveThisBody, ref RelativePlane);

                            if (Status == true)
                            {
                                //iSwApp.SendMsgToUser("Feasible body exist, ready to be registered in the removal sequence");

                                if (ChildStatus == true)
                                {
                                    ChildStatus = false;

                                    //if (BodyToDelete.Count == 1)
                                    if (RemoveThisBody.Count == 1)
                                    {
                                        //prepare the deletion of the splitted body
                                        BodiesInfo = null;
                                        Body2 TmpBody = null;
                                        int DeleteIndex = 0;
                                        bodyArray = null;

                                        //bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                        //for (int j = 0; j < bodyArray.Length; j++)
                                        //{
                                        //    TmpBody = (Body2)bodyArray.GetValue(j);

                                        //    //if (TmpBody.Name.Equals(BodyToDelete[0]))
                                        //    if (TmpBody.Name.Equals(RemoveThisBody[0].Name))
                                        //    {
                                        //        DeleteIndex = j;
                                        //        break;
                                        //    }
                                        //}

                                        SelectData TmpSelectData = null;
                                        //bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);
                                        bool SelectionStatus = RemoveThisBody[0].BodyObj.Select2(true, TmpSelectData);
                                        //SwApp.SendMsgToUser("Removed shape by " + SelectedFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());
                                        DeleteFeature = (Feature)Doc.FeatureManager.InsertDeleteBody();

                                        //add the parent plane and its feature
                                        ListOfParentPlanes.Add(ParentPlane);
                                        ListOfParentFeature.Add(SplitFeature);
                                        ListOfParentFeature.Add(DeleteFeature);

                                        //add the relative planes
                                        if (RelativePlane.Count() != 0) { ListOfRelativePlanes.AddRange(RelativePlane); }

                                        //register the split action
                                        RemovedBody Removal = new RemovedBody();
                                        Removal.ParentPlane = ParentPlane;
                                        Removal.Removal = ListOfParentFeature;
                                        Removal.TRV = RemoveThisBody[0];
                                        

                                        if (PreviousRemoval != null) { PreviousRemoval.Add(Removal); }

                                    }
                                    else
                                    {
                                        //iSwApp.SendMsgToUser("Body > 1 section still underdevelopment");

                                        List<RemovalBody> SortedRemovalBody = RemoveThisBody.OrderByDescending(plane => plane.Volume).ToList();

                                        BodiesInfo = null;
                                        Body2 TmpBody = null;
                                        int DeleteIndex = 0;
                                        bodyArray = null;
                                                                               
                                        //bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                        //for (int j = 0; j < bodyArray.Length; j++)
                                        //{
                                        //    TmpBody = (Body2)bodyArray.GetValue(j);

                                        //    if (TmpBody.Name.Equals(SortedRemovalBody[0].Name))
                                        //    {
                                        //        DeleteIndex = j;
                                        //        break;
                                        //    }
                                        //}

                                        SelectData TmpSelectData = null;
                                        //bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);
                                        bool SelectionStatus = SortedRemovalBody[0].BodyObj.Select2(true, TmpSelectData);
                                        //SwApp.SendMsgToUser("Removed shape by " + SelectedFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());
                                        DeleteFeature = (Feature)Doc.FeatureManager.InsertDeleteBody();

                                        ListOfParentPlanes.Add(ParentPlane);
                                        ListOfParentFeature.Add(SplitFeature);
                                        ListOfParentFeature.Add(DeleteFeature);

                                        //register the split action
                                        RemovedBody Removal = new RemovedBody();
                                        Removal.ParentPlane = ParentPlane;
                                        Removal.Removal = ListOfParentFeature;
                                        Removal.TRV = SortedRemovalBody[0];

                                        if (PreviousRemoval != null) { PreviousRemoval.Add(Removal); }
                                        
                                    }

                                }
                                else
                                {
                                    
                                        //set the status of possibiliity to be true
                                        PlaneListByScore[index].Possibility = true;

                                        //set 0 mark that indicates this plane can be set with removal volume
                                        setMarkOnPlane(ref PlaneListByScore, index, "PROCESSOK");

                                    //delete the split feature
                                    ModelDoc2 DocumentModel = (ModelDoc2)swComp.GetModelDoc2();
                                    Feature LastFeature = (Feature)swComp.FeatureByName(SplitFeature.Name);
                                    bool SStatus = LastFeature.Select2(true, 3);
                                    Doc.EditDelete();
                                    
                                }
                            }
                            else
                            {
                                if (ChildStatus == true)
                                {
                                    ChildStatus = false;
                                }
                                else
                                {
                                    //set 1 on this plane that indicates nothing can be set to this plane
                                    setMarkOnPlane(ref PlaneListByScore, index, "NOBODYFOUND");

                                    //set the continuity to be false
                                    PlaneListByScore[index].Possibility = false;

                                    //delete the split feature
                                    ModelDoc2 DocumentModel = (ModelDoc2)swComp.GetModelDoc2();
                                    Feature LastFeature = (Feature)swComp.FeatureByName(SplitFeature.Name);
                                    bool SStatus = LastFeature.Select2(true, 3);
                                    Doc.EditDelete();
                                }

                            }

                        }

                        catch (Exception ex)
                        {
                            iSwApp.SendMsgToUser("Body evaluation failed" + "\r\"Exception message: " + ex.Message);
                        }
                    }

                    else
                    {
                        if (ChildStatus == true)
                        {
                            ChildStatus = false;
                        }
                        else
                        {
                            //set 2 on this plane that indicates nothing can be found
                            setMarkOnPlane(ref PlaneListByScore, index, "NOBODYSET");

                            //set the continuity to be false
                            PlaneListByScore[index].Possibility = false;
                        }
                    }
                }

                Doc = ((ModelDoc2)(SwApp.ActiveDoc));
                Doc.ClearSelection2(true);

                SelectedFeature = getThePlane(ref PlaneListByScore, PlaneListByDistance , featureList, 
                    ref index, ListOfParentPlanes, ListOfRelativePlanes);

            }

            //Close the parallel opened document and make the document to be visible again
            iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);

            //collect only the true possibilities
            List<AddedReferencePlane> PlaneListByPosibility = new List<AddedReferencePlane>();
            PlaneListByPosibility = 
                PlaneListByScore.Where(Plane => Plane.Possibility == true).
                OrderByDescending(Plane => Plane.Score).ThenByDescending(Plane => Plane.DistanceFromCentroid).ToList();

            if (PlaneListByPosibility.Count != 0)
            {
                //elaborate each possible plane
                foreach (AddedReferencePlane PossiblePlane in PlaneListByPosibility)
                {
                    if (PreviousRemoval == null) { PreviousRemoval = new List<RemovedBody>(); }
                    if (ListOfParentPlanes == null) { ListOfParentPlanes = new List<AddedReferencePlane>(); }
                    if (ListOfRelativePlanes == null) { ListOfRelativePlanes = new List<AddedReferencePlane>(); }
                    
                    traversePlanes(Doc, assyModel, swComp, featureList, PossiblePlane, ListOfParentPlanes, ListOfRelativePlanes, ref PreviousRemoval);

                    //roll back previous process before continuing the elaboration on next possible plane
                    if (PreviousRemoval.Count > 0)
                    {
                        List<RemovedBody> RemovedByPossiblePlane =
                            PreviousRemoval.Where(plane => plane.ParentPlane.name.Equals(PossiblePlane.name)).ToList();

                        bool RollBackStatus = RollBackProcess(Doc, swComp, RemovedByPossiblePlane);

                        if (RollBackStatus == true)
                        {
                            //remove the visited possible plane from removal list
                            int RemoveThis = PreviousRemoval.FindIndex(plane => plane.ParentPlane.name.Equals(PossiblePlane.name));
                            PreviousRemoval.RemoveAt(RemoveThis);

                            //remove the visited possible plane from parent list
                            RemoveThis = ListOfParentPlanes.FindIndex(plane => plane.name.Equals(PossiblePlane.name));
                            ListOfParentPlanes.RemoveAt(RemoveThis);
                        }
                    }
                }
            }
            else
            {   
                bool AssignedStatus = SetAsMachiningPlan(PreviousRemoval);

                if (AssignedStatus == true)
                {
                    ProcessLog_TaskPaneHost.LogProcess("Found " + MachiningPlanList.Count.ToString() + " machining plans");
                    PPDetails_TaskPaneHost.LogProcess("Found " + MachiningPlanList.Count.ToString() + " machining plans");
                }
            }
 
        }
               

        //Rollback previous removal process
        public static bool RollBackProcess(ModelDoc2 ThisDocument, Component2 ThisComponent, List<RemovedBody> ThisRemovalList)
        {
            try
            {
                if (ThisRemovalList.Count > 0)
                {
                    //access from the last index
                    for (int i = ThisRemovalList.Count - 1; i >= 0; i--)
                    {
                        for (int j = ThisRemovalList[i].Removal.Count - 1; j >= 0; j--)
                        {
                            //delete the split feature
                            ModelDoc2 DocumentModel = (ModelDoc2)ThisComponent.GetModelDoc2();
                            Feature DeleteThisFeature = (Feature)ThisComponent.FeatureByName(ThisRemovalList[i].Removal[j].Name);
                            bool SStatus = DeleteThisFeature.Select2(true, 3);
                            ThisDocument.EditDelete();
                        }
                    }
                }
                else { return false; }
            }
            catch { return false; }

            return true;
        }

        public static string MainView;

        //show the generated MP
        public static void ShowTheMP(int MPIndex)
        {   
            if (MachiningPlanList[MPIndex].ViewName != null) 
            {
                iSwApp.ActivateDoc(MachiningPlanList[MPIndex].ViewName);
            }
        }
        
        //Generate button click
        public static bool GenerateMP(int MPIndex)
        {
            if (MachiningPlanList[MPIndex].ViewName != null)
            {
                //iSwApp.CloseDoc(MachiningPlanList[MPIndex].ViewName);
                return false; 
            }

            if (MainView == "") { return false; }

            ModelDoc2 Doc = (ModelDoc2) SwApp.ActivateDoc(MainView);
            
            int docType = (int)Doc.GetType();
            bool boolStatus = false;

            if (Doc.GetType() == 2)
            {
                AssemblyDoc assyModel = (AssemblyDoc)Doc;
                if (assyModel.GetComponentCount(false) != 0)
                {
                    List<_planeProperties> samplePlane = new List<_planeProperties>();
                    samplePlane = planeList;
                    int numPlane = samplePlane.Count;

                    FeatureManager docFeatureMgr = Doc.FeatureManager;
                    object featArr = docFeatureMgr.GetFeatures(false);
                    List<Feature> PlaneFeatures = new List<Feature>();

                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    if (compName == null)
                    {
                        //define the raw material and product
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);
                        PPDetails_TaskPaneHost.getCompName(ref compName);
                    }

                    boolStatus = compName[0].Select2(true, 0);

                    TraverseComponentFeatures(compName[0], ref PlaneFeatures);
                    List<string> PathList = null;

                    //create the plan directory
                    string PlanDirectory = Path.GetDirectoryName(Doc.GetPathName()) + "\\Plan" + (MPIndex +1).ToString();
                    System.IO.Directory.CreateDirectory(PlanDirectory);

                    //check the generated MP
                    if (MPGenerator(Doc, assyModel, compName[0], PlaneFeatures, MachiningPlanList[MPIndex], MPIndex, PlanDirectory, out PathList) == true)
                    {   
                        ModelDoc2 NewDoc = (ModelDoc2)SwApp.NewDocument("G:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\english\\Tutorial\\assem.asmdot", 0, 0, 0);
                        
                        string[] StringNames = new string[PathList.Count + 1];
                        string[] CoordinateName = new string[PathList.Count + 1];
                        StringNames[0] = compName[1].GetPathName();
                        CoordinateName[0] = "";

                        for (int i = 1; i <= PathList.Count(); i++)
                        {
                            StringNames[i] = PathList[i-1];
                            CoordinateName[i] = "";
                        }

                        AssemblyDoc NewAssy = (AssemblyDoc)NewDoc;
                        Object vcomponents = NewAssy.AddComponents3((StringNames), null, (CoordinateName));

                        string MPPath = PlanDirectory + "\\Machining Plan " + (MPIndex+1).ToString() + ".sldasm" ;
                        bool SaveStatus = NewDoc.Extension.SaveAs(MPPath , (int) swSaveAsVersion_e.swSaveAsCurrentVersion,
                            (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);

                        if (SaveStatus == true) { MachiningPlanList[MPIndex].FullPath = MPPath; }

                        //set the openfaces
                        Boolean OpenFaceStatus = FindOpenFaces(NewDoc, NewAssy);

                        //NewDoc.Save();

                        for (int i = 0; i < PathList.Count(); i++)
                        {
                            iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(PathList[i]));
                        }

                        //save the document name
                        MachiningPlanList[MPIndex].ViewName = Path.GetFileNameWithoutExtension(NewDoc.GetPathName());
                        ProcessLog_TaskPaneHost.LogProcess("Generating Machining Plan " + (MPIndex + 1).ToString());
                        PPDetails_TaskPaneHost.LogProcess("Generating Machining Plan " + (MPIndex + 1).ToString());

                        NewDoc = SwApp.ActivateDoc(Path.GetFileNameWithoutExtension(MPPath));
                        NewDoc.ViewZoomtofit2();

                    }
                    else
                    {
                        System.IO.Directory.Delete(PlanDirectory);
                        return false;
                    }
                }
                else { return false; }

            }
            else { return false; }
            
            return true;
        }

        //Generate the selected machining plane *Trigger by the "Generate" button in the process task pane
        public static bool MPGenerator(ModelDoc2 Doc, AssemblyDoc assyModel, Component2 swComp, List<Feature> featureList, MachiningPlan ThisMP, int ThisMPIndex, 
            string ThisMPDir, out List<string> ThisPathList)
        {
            Feature SelectedFeature = null;
            bool boolStatus;
            int index = 0;
            int SequenceNUM = 1;
            int SplitCounter = 0;
            int Errors = 0;
            int Warnings = 0;

            ThisPathList = new List<string>();
            List<AddedReferencePlane> ListOfRemovalPlanes = new List<AddedReferencePlane>();
            List<Feature> ListOfRemovalFeature = null;
            List<RemovedBody> PreviousRemoval = new List<RemovedBody>();

            List<MachiningProcess> ProcessCollection = new List<MachiningProcess>();
            
            ProcessCollection = ThisMP.MachiningProceses;

            //make the loaded document to be invisble
            String CompPathName = swComp.GetPathName();
            iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART);
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

            foreach (MachiningProcess SelectedProcess in ProcessCollection)
            {
                //
                SelectedFeature = getSelectedPlane(SelectedProcess.MachiningReference, featureList);
                string featurename = SelectedFeature.GetTypeName2();


                index = getPlaneIndex(SelectedProcess.MachiningReference, PlaneListByScore);
            
                assyModel.EditPart();

                boolStatus = SelectedFeature.Select2(true, 0);
                
                //split and collect the body
                List<int> DeleteThisBody = null;
                Array bodyArray = null;
                bodyArray = (Array)Doc.FeatureManager.PreSplitBody();

                //check if there are body 
                if (bodyArray == null) 
                { 
                    iSwApp.SendMsgToUser("Nothing can be collected by using " + SelectedFeature.Name);
                    return false;
                }
                else
                {
                    string[] bodyNames = new string[bodyArray.Length]; //set the name

                    Feature SplitFeature = null;
                    Feature DeleteFeature = null;

                    SplitFeature = SplitAndSaveBody(Doc, swComp, SelectedFeature, bodyArray, ref bodyNames);

                    if (SplitFeature != null)
                    {
                        try
                        {
                            List<double> Volume = new List<double>();
                            Object BodiesInfo = null;
                            List<string> BodyToDelete = null;

                            List<RemovalBody> RemoveThisBody = null;

                            DeleteThisBody = new List<int>();
                            SplitCounter++; //add the split counter after successful splitting process

                            bodyArray = null;
                            bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                            //select body that needs to be deleted from the model
                            //bool Status = SelectBodyToDelete(CompDocumentModel, bodyArray, ref BodyToDelete, index, ref Volume, ref RemoveThisBody);
                            bool Status = SelectBodyToDelete(CompDocumentModel, bodyArray, index, ref RemoveThisBody);

                            if (Status == true)
                            {
                                //iSwApp.SendMsgToUser("Feasible body exist, ready to be registered in the removal sequence");

                                    //if (BodyToDelete.Count == 1)
                                    if (RemoveThisBody.Count == 1)
                                    {
                                        //prepare the deletion of the splitted body
                                        BodiesInfo = null;
                                        Body2 TmpBody = null;
                                        int DeleteIndex = 0;
                                        bodyArray = null;

                                        PartDoc TmpPartDoc = (PartDoc)CompDocumentModel;
                                        bodyArray = (Array)TmpPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                                        //bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                        for (int j = 0; j < bodyArray.Length; j++)
                                        {
                                            TmpBody = (Body2)bodyArray.GetValue(j);

                                            //if (TmpBody.Name.Equals(BodyToDelete[0]))
                                            if (TmpBody.Name.Equals(RemoveThisBody[0].Name))
                                            {
                                                DeleteIndex = j;
                                                break;
                                            }
                                        }

                                        SelectData TmpSelectData = null;
                                        bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);
                                        string ActualTRVPath = ThisMPDir + "\\0" + SequenceNUM.ToString() + "_ARV_Plan" + (ThisMPIndex + 1).ToString() + "_SEQ" + SequenceNUM.ToString() + "_" + SelectedFeature.Name + ".sldprt";
                                        //Insert the Body into new part document (this is the TRUE TRV
                                        bool SaveStatus = ((PartDoc) CompDocumentModel).SaveToFile3(ActualTRVPath, 1, 1, false, "", out Errors, out Warnings);

                                        //create TRV from TRUE TRV by tesselating all faces in the TRUE TRV body
                                        string TRVPath = ThisMPDir + "\\0" + SequenceNUM.ToString() + "_PRV_Plan" + (ThisMPIndex + 1).ToString() + "_SEQ" + SequenceNUM.ToString() + "_" + SelectedFeature.Name + ".sldprt";
                                        bool CreationStatus = CreateTRV(TRVPath, TmpBody);

                                        if (SaveStatus == true)
                                        {
                                            ThisPathList.Add(ActualTRVPath); 
                                            SequenceNUM++;
                                            if (CreationStatus == true)
                                            {
                                                ThisPathList.Add(TRVPath);
                                            }
                                        }

                                        //break the reference
                                        ModelDoc2 TmpDoc = (ModelDoc2)SwApp.ActivateDoc2(Path.GetFileNameWithoutExtension(ActualTRVPath), true, 0);
                                        ModelDocExtension TmpDocEx = (ModelDocExtension) TmpDoc.Extension;
                                        TmpDocEx.BreakAllExternalFileReferences2(false);
                                        TmpDoc.Save2(true);

                                        //activate the assembly file
                                        iSwApp.ActivateDoc(Path.GetFileNameWithoutExtension(Doc.GetPathName()));

                                        //SwApp.SendMsgToUser("Removed shape by " + SelectedFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());
                                        DeleteFeature = (Feature)CompDocumentModel.FeatureManager.InsertDeleteBody();
                                        ListOfRemovalPlanes.Add(SelectedProcess.MachiningReference);

                                        ListOfRemovalFeature = new List<Feature>();
                                        ListOfRemovalFeature.Add(SplitFeature);
                                        ListOfRemovalFeature.Add(DeleteFeature);

                                        //register the split action
                                        RemovedBody Removal = new RemovedBody();
                                        Removal.ParentPlane = SelectedProcess.MachiningReference;
                                        Removal.Removal = ListOfRemovalFeature;

                                        if (PreviousRemoval != null) { PreviousRemoval.Add(Removal); }

                                    }
                                    else
                                    {
                                        //iSwApp.SendMsgToUser("Body > 1 section still underdevelopment");

                                        List<RemovalBody> SortedRemovalBody = RemoveThisBody.OrderByDescending(plane => plane.Volume).ToList();

                                        BodiesInfo = null;
                                        Body2 TmpBody = null;
                                        int DeleteIndex = 0;
                                        bodyArray = null;

                                        PartDoc TmpPartDoc = (PartDoc)CompDocumentModel;
    
                                        bodyArray = (Array)TmpPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                                        //bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                        for (int j = 0; j < bodyArray.Length; j++)
                                        {
                                            TmpBody = (Body2)bodyArray.GetValue(j);

                                            //if (TmpBody.Name.Equals(BodyToDelete[0]))
                                            if (TmpBody.Name.Equals(SortedRemovalBody[0].Name))
                                            {
                                                DeleteIndex = j;
                                                break;
                                            }
                                        }

                                        SelectData TmpSelectData = null;
                                        bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);
                                        string ActualTRVPath = ThisMPDir + "\\0" + SequenceNUM.ToString() + "_ARV_Plan" + (ThisMPIndex + 1).ToString() + "_SEQ" + SequenceNUM.ToString() + "_" + SelectedFeature.Name + ".sldprt";
                                        //Insert the Body into new part document (this is the TRUE TRV
                                        bool SaveStatus = ((PartDoc)CompDocumentModel).SaveToFile3(ActualTRVPath, 1, 1, false, "", out Errors, out Warnings);

                                        //create TRV from TRUE TRV by tesselating all faces in the TRUE TRV body
                                        string TRVPath = ThisMPDir + "\\0" + SequenceNUM.ToString() + "_PRV_Plan" + (ThisMPIndex + 1).ToString() + "_SEQ" + SequenceNUM.ToString() + "_" + SelectedFeature.Name + ".sldprt";
                                        bool CreationStatus = CreateTRV(TRVPath, TmpBody);

                                        if (SaveStatus == true)
                                        {
                                            ThisPathList.Add(ActualTRVPath);
                                            SequenceNUM++;
                                            if (CreationStatus == true)
                                            {
                                                ThisPathList.Add(TRVPath);
                                            }
                                        }

                                        //break the reference
                                        ModelDoc2 TmpDoc = (ModelDoc2)SwApp.ActivateDoc2(Path.GetFileNameWithoutExtension(ActualTRVPath), true, 0);
                                        ModelDocExtension TmpDocEx = (ModelDocExtension)TmpDoc.Extension;
                                        TmpDocEx.BreakAllExternalFileReferences2(false);
                                        TmpDoc.Save2(true);

                                        //activate the assembly file
                                        iSwApp.ActivateDoc(Path.GetFileNameWithoutExtension(Doc.GetPathName()));

                                        //SwApp.SendMsgToUser("Removed shape by " + SelectedFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());
                                        DeleteFeature = (Feature)CompDocumentModel.FeatureManager.InsertDeleteBody();
                                        ListOfRemovalPlanes.Add(SelectedProcess.MachiningReference);

                                        ListOfRemovalFeature = new List<Feature>();
                                        ListOfRemovalFeature.Add(SplitFeature);
                                        ListOfRemovalFeature.Add(DeleteFeature);

                                        //register the split action
                                        RemovedBody Removal = new RemovedBody();
                                        Removal.ParentPlane = SelectedProcess.MachiningReference;
                                        Removal.Removal = ListOfRemovalFeature;

                                        if (PreviousRemoval != null) { PreviousRemoval.Add(Removal); }

                                        
                                    }
                            }
                            else
                            {
                                iSwApp.SendMsgToUser("No body can be determined by " + SelectedFeature.Name);

                                //delete the split feature
                                ModelDoc2 DocumentModel = (ModelDoc2)swComp.GetModelDoc2();
                                Feature LastFeature = (Feature)swComp.FeatureByName(SplitFeature.Name);
                                bool SStatus = LastFeature.Select2(true, 3);
                                Doc.EditDelete();

                                return false;
                            }

                        }

                        catch (Exception ex)
                        {
                            iSwApp.SendMsgToUser("Body evaluation was error" + "\r\"Exception message: " + ex.Message);
                            return false;
                        }
                    }

                    else
                    {
                        iSwApp.SendMsgToUser("Split error by using " + SelectedFeature.Name);
                        return false;
                    }
                }

                Doc = ((ModelDoc2)(SwApp.ActiveDoc));
                Doc.ClearSelection2(true);
            }

            //Close the parallel opened document and make the document to be visble again
            iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART);

            //roll back previous process before continuing the elaboration on next possible plane
            if (PreviousRemoval.Count > 0)
            {
                bool RollBackStatus = RollBackProcess(Doc, swComp, PreviousRemoval);

                if (RollBackStatus == true)
                {
                    PreviousRemoval.Clear();
                }
            }
            else { return false; }

            //assyModel.EditAssembly();
            return true;
        }

        //Create TRV from ActualTRV body
        public static Boolean CreateTRV(String ThisPath, Body2 ThisBody)
        {
            Object[] ThisBodyFaces = (Object[])ThisBody.GetFaces();
            Object TessFaceArray = null;
            Double[] TessVerticesArray = null;
            Double[] ThisBodyMaxMin = null;
            SketchPoint SketchObj1, SketchObj2, SketchObj3 = null;
            Object CornerPoints = null;
            int Errors = 0;
            int Warnings = 0;
            //Boolean SelectionStatus;
            RefPlane ThisRefPlane = null;            
            MathTransform RefPlaneTransform;
            MathPoint ThisVertex = null;
            MathUtility MathUtils = SwApp.GetMathUtility();
            
            List<Double> NewPoint= new List<Double>();


            //get the maxmin vertices of the body by tesselation
            foreach (Face2 ThisFace in ThisBodyFaces)
            {
                TessFaceArray = (Object)ThisFace.GetTessTriangles(true);
                TessVerticesArray = GetMaxMin(TessFaceArray);
                
                Boolean UpdateStatus = SetMaxMin(TessVerticesArray, ref ThisBodyMaxMin);
                
            }

            ModelDoc2 ThisDoc = (ModelDoc2)SwApp.NewDocument("G:\\Program Files\\SolidWorks Corp\\SOLIDWORKS\\lang\\english\\Tutorial\\part.prtdot", 0, 0, 0);            
            bool SaveStatus = ThisDoc.Extension.SaveAs(ThisPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, 0, 0);

            ThisDoc = SwApp.ActivateDoc3(Path.GetFileNameWithoutExtension(ThisPath), false, 2, ref Errors);
            
            //ThisDoc.ActiveView();

            SketchManager ThisSkManager = (SketchManager) ThisDoc.SketchManager;
            
            ThisSkManager.Insert3DSketch(true);
            SketchObj1 = ThisSkManager.CreatePoint(ThisBodyMaxMin[0], ThisBodyMaxMin[1], ThisBodyMaxMin[5]);
            SketchObj2 = ThisSkManager.CreatePoint(ThisBodyMaxMin[0], ThisBodyMaxMin[4], ThisBodyMaxMin[5]);
            SketchObj3 = ThisSkManager.CreatePoint(ThisBodyMaxMin[3], ThisBodyMaxMin[4], ThisBodyMaxMin[5]);
            ThisSkManager.Insert3DSketch(true);

            //select the constraint and insert the reference plane
            //SelectionStatus = ThisDoc.Extension.SelectByID2("", "EXTSKETCHPOINT", ThisBodyMaxMin[0], ThisBodyMaxMin[1], ThisBodyMaxMin[5], true, 0, null, 0);
            //SelectionStatus = ThisDoc.Extension.SelectByID2("", "EXTSKETCHPOINT", ThisBodyMaxMin[0], ThisBodyMaxMin[4], ThisBodyMaxMin[5], true, 1, null, 0);
            //SelectionStatus = ThisDoc.Extension.SelectByID2("", "EXTSKETCHPOINT", ThisBodyMaxMin[3], ThisBodyMaxMin[4], ThisBodyMaxMin[5], true, 2, null, 0);

            SketchObj1.Select2(true, 0);
            SketchObj2.Select2(true, 1);
            SketchObj3.Select2(true, 2);

            ThisRefPlane = (RefPlane)ThisDoc.FeatureManager.InsertRefPlane(4, 0, 4, 0, 4, 0);
            
            //set reference name
            PartDoc ThisPartDoc = (PartDoc)ThisDoc;
            ThisPartDoc.SetEntityName(ThisRefPlane, "REF_PLANE");

            ThisDoc.Extension.SelectByID2("REF_PLANE", "PLANE", 0, 0, 0, false, 0, null, 0);

            //get the inverse of reference plane transform
            RefPlaneTransform = ThisRefPlane.Transform.Inverse();
                        
            ThisVertex = MathUtils.CreatePoint(new Double[] {ThisBodyMaxMin[0], ThisBodyMaxMin[1], ThisBodyMaxMin[5]});
            ThisVertex = ThisVertex.MultiplyTransform(RefPlaneTransform);
            Double[] FirstCorner = (Double[])ThisVertex.ArrayData;

            ThisVertex = MathUtils.CreatePoint(new Double[] { ThisBodyMaxMin[0], ThisBodyMaxMin[4], ThisBodyMaxMin[5] });
            ThisVertex = ThisVertex.MultiplyTransform(RefPlaneTransform);
            Double[] SecondCorner = (Double[])ThisVertex.ArrayData;

            ThisVertex = MathUtils.CreatePoint(new Double[] { ThisBodyMaxMin[3], ThisBodyMaxMin[4], ThisBodyMaxMin[5] });
            ThisVertex = ThisVertex.MultiplyTransform(RefPlaneTransform);
            Double[] ThirdCorner = (Double[])ThisVertex.ArrayData;

            ThisSkManager.InsertSketch(true);
            /*
            SketchObj = ThisSkManager.Create3PointCornerRectangle(FirstCorner[0], FirstCorner[1], FirstCorner[2],
                SecondCorner[0], SecondCorner[1], SecondCorner[2],
                ThirdCorner[0], ThirdCorner[1], ThirdCorner[2]);
             * */

            CornerPoints = ThisSkManager.CreateCornerRectangle(FirstCorner[0], FirstCorner[1], FirstCorner[2],
                ThirdCorner[0], ThirdCorner[1], ThirdCorner[2]);
            
            Feature TRVFeature = ThisDoc.FeatureManager.FeatureExtrusion3(true, false, true, 0, 0, Math.Round(ThisBodyMaxMin[2] - ThisBodyMaxMin[5], 3) , 0.01,
                false, false, false, false, 0, 0, false, false, false, false, true, false, false, 0, 0, false);

            SaveStatus = ThisDoc.Save3(1, ref Errors, ref Warnings);

            return true;
        }

        //analyze open faces in the active assembly document
        public void AnalyzeOpenFace()
        {
            iSwApp.SendMsgToUser("start analyzing open faces...");
            
            ModelDoc2 Doc = (ModelDoc2) SwApp.ActiveDoc;
            
            int docType = (int)Doc.GetType();
            
            if (Doc.GetType() == 2)
            {
                AssemblyDoc AssyModel = (AssemblyDoc)Doc;
                Boolean FindStatus = FindOpenFaces(Doc, AssyModel);
            }
        }

        //find the open faces in each component
        public static bool FindOpenFaces(ModelDoc2 ThisModelDoc, AssemblyDoc ThisAssyDoc)
        {

            if (ThisAssyDoc.GetComponentCount(true) == 0) { return false; }
            else
            {
                ProcessLog_TaskPaneHost.LogProcess("Analysing open faces of current assembly model");
                PPDetails_TaskPaneHost.LogProcess("Analysing open faces of current assembly model");

                //get the components
                ConfigurationManager configDocManager = (ConfigurationManager)ThisModelDoc.ConfigurationManager;
                Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);
                Object[] childComp = (Object[])compInAssembly.GetChildren();

                //classify between the workpiece and trvs
                List<Component2> ThisComponents = new List<Component2>();
                List<Component2> ThisComponentsSorted = new List<Component2>();
                List<Component2> TRVComponents = new List<Component2>();
                List<Component2> TRVComponentsSorted = new List<Component2>();
                Component2 RootComponent = null;

                for (int i = 0; i < ThisAssyDoc.GetComponentCount(true); i++)
                {
                    Component2 TmpComponent = (Component2)childComp[i];

                    if (i == 0)
                    {
                        RootComponent = TmpComponent;
                        RootComponent.Select2(true, 0);
                        ThisAssyDoc.SetComponentTransparent(true);
                    }
                    else if (TmpComponent.Name.Contains("_ARV"))
                    {
                        ThisComponents.Add(TmpComponent);
                    }
                    else
                    {
                        TRVComponents.Add(TmpComponent);
                        TmpComponent.Select2(true, 0);
                        ThisAssyDoc.SetComponentTransparent(true);
                    }
                }

                //order it by its name (may be triggering problem if the components more than 10)
                ThisComponentsSorted = ThisComponents.OrderBy(c => c.Name2).ToList();

                //get the product/workpiece (main component) faces
                Array CompBodyArray = (Array)RootComponent.GetBodies2((int)swBodyType_e.swSolidBody);
                Body2 MainCompBody = (Body2)CompBodyArray.GetValue(0);
                Object[] MainCompFaces = (Object[])MainCompBody.GetFaces();

                //Object[] TrueTRVFaces = null;

                List<Object[]> TrueTRVFaces = new List<object[]>();

                //set the open face mark as "90"
                SelectionMgr ThisSelMgr = (SelectionMgr)ThisModelDoc.SelectionManager;
                SelectData ThisSelectData = ThisSelMgr.CreateSelectData();

                //get the color template
                Double[] PropValues = ThisModelDoc.MaterialPropertyValues;
                PropValues[0] = 1; //R
                PropValues[1] = 0; //G
                PropValues[2] = 0; //B

                for (int i = 0; i < ThisComponentsSorted.Count(); i++)
                {
                    //select and open the component for editing
                    ThisComponentsSorted[i].Select2(true, 0);
                    ThisAssyDoc.EditPart();

                    //ModelDoc2 RefModel = ThisComponents[i].GetModelDoc2();
                    //Double[] PropValues = RefModel.MaterialPropertyValues;

                    //Compare with product's faces
                    CompBodyArray = (Array)ThisComponentsSorted[i].GetBodies2((int)swBodyType_e.swSolidBody);
                    Body2 CompBody = (Body2)CompBodyArray.GetValue(0);
                    Object[] CompFaces = (Object[])CompBody.GetFaces();
                    TrueTRVFaces.Add(CompFaces);

                    //First, check the face with the main product's faces
                    foreach (Face2 CheckThisFace in CompFaces)
                    {

                        //only planar plane will proceed, non-planar will be colored red directly
                        if (IsPlanar(CheckThisFace) == true)
                        {
                            foreach (Face2 FaceReference in MainCompFaces)
                            {
                                //check only with planar face
                                if (IsPlanar(FaceReference) == true)
                                {
                                    if (IsOverlapped(CheckThisFace, FaceReference, ThisComponentsSorted[i], RootComponent) == true)
                                    {

                                        //change the face color to red and mark it with "90"
                                        CheckThisFace.MaterialPropertyValues = PropValues;
                                        ThisSelectData.Mark = 90;

                                    }
                                    else
                                    {

                                        //change the face color to green and mark it with "91"
                                        ThisSelectData.Mark = 91;

                                    }
                                    
                                    Entity ThisEntity = (Entity)CheckThisFace;
                                    Boolean MarkingStatus = ThisEntity.Select4(true, ThisSelectData);
                                }

                            }

                            //Second, check the face againts another TRV's faces
                            for (int j = i + 1; j < ThisComponentsSorted.Count(); j++)
                            {
                                CompBodyArray = (Array)ThisComponentsSorted[j].GetBodies2((int)swBodyType_e.swSolidBody);
                                Body2 OtherCompBody = (Body2)CompBodyArray.GetValue(0);
                                Object[] OtherCompFaces = (Object[])OtherCompBody.GetFaces();

                                foreach (Face2 OtherFaceReference in OtherCompFaces)
                                {
                                    //check only with planar face
                                    if (IsPlanar(OtherFaceReference) == true)
                                    {
                                        if (IsOverlapped(CheckThisFace, OtherFaceReference, ThisComponentsSorted[i], ThisComponentsSorted[j]) == true)
                                        {

                                            //change the face color to red and mark it with "90"
                                            CheckThisFace.MaterialPropertyValues = PropValues;
                                            ThisSelectData.Mark = 90;

                                        }
                                        else
                                        {

                                            //change the face color to green and mark it with "91"
                                            ThisSelectData.Mark = 91;

                                        }

                                        Entity ThisEntity = (Entity)CheckThisFace;
                                        Boolean MarkingStatus = ThisEntity.Select4(true, ThisSelectData);
                                    }
                                }
                            }
                        }
                        else
                        {
                            //change the face color to red and mark it with "90"
                            CheckThisFace.MaterialPropertyValues = PropValues;
                            ThisSelectData.Mark = 90;

                            Entity ThisEntity = (Entity)CheckThisFace;
                            Boolean MarkingStatus = ThisEntity.Select4(true, ThisSelectData);
                        }

                    }

                    ThisAssyDoc.EditAssembly();
                }

                //set the color of TRV based on TRUE TRV faces
                //order it by its name (may be triggering problem if the components more than 10)
                TRVComponentsSorted = TRVComponents.OrderBy(c => c.Name2).ToList();

                MathUtility MathUtils = SwApp.GetMathUtility();

                for (int i = 0; i < TRVComponents.Count(); i++)
                {
                    //select and open the component for editing
                    TRVComponentsSorted[i].Select2(true, 0);
                    ThisAssyDoc.EditPart();

                    //ModelDoc2 RefModel = ThisComponents[i].GetModelDoc2();
                    //Double[] PropValues = RefModel.MaterialPropertyValues;

                    //Compare with product's faces
                    CompBodyArray = (Array)TRVComponentsSorted[i].GetBodies2((int)swBodyType_e.swSolidBody);
                    Body2 CompBody = (Body2)CompBodyArray.GetValue(0);
                    Object[] CompFaces = (Object[])CompBody.GetFaces();

                    //get the product/workpiece (main component) faces
                    //CompBodyArray = (Array)ThisComponentsSorted[i].GetBodies2((int)swBodyType_e.swSolidBody);
                    //Body2 TrueTRVBody = (Body2)CompBodyArray.GetValue(0);
                    //Object[] TrueTRVFaces = (Object[])TrueTRVBody.GetFaces();

                    foreach (Face2 CheckThisFace in CompFaces)
                    {
                        foreach (Face2 WithThisFace in TrueTRVFaces[i])
                        { 
                            //if (IsSameDirection(CheckThisFace.Normal, WithThisFace.Normal) == true)
                            //{
                            if (IsPlanar(WithThisFace) == true)
                            {
                                if (CheckThisFace.IsCoincident(WithThisFace, 0.001) == 0)
                                {
                                    //change the face color to red and mark it with "90"
                                    CheckThisFace.MaterialPropertyValues = WithThisFace.MaterialPropertyValues;

                                    break;
                                }
                            }
                        }
                    }

                    ThisAssyDoc.EditAssembly();

                }

                    
            }

            return true;
        }

        //check vector sign
        public static Boolean IsSameDirection(Object FirstN, Object SecondN)
        {
            Double[] N1 = (Double[])FirstN;
            Double[] N2 = (Double[])SecondN;

            if (N1[0] == N2[0] && N1[1] == N2[1] && N1[2] == N2[2]) { return true; }
            
            return false;
        }

        //check overlapping faces
        public static Boolean IsOverlapped(Face2 FaceA, Face2 FaceB, Component2 CompA, Component2 CompB)
        {
            
            ModelDoc2 RefModelA = CompA.GetModelDoc2();
            Face2 CorrespondFaceA = RefModelA.Extension.GetCorresponding(FaceA);
            Body2 BodyA = CorrespondFaceA.CreateSheetBody();
            BodyA.ApplyTransform(CompA.Transform2);

            ModelDoc2 RefModelB = CompB.GetModelDoc2();
            Face2 CorrespondFaceB = RefModelB.Extension.GetCorresponding(FaceB);
            Body2 BodyB = CorrespondFaceB.CreateSheetBody();
            BodyB.ApplyTransform(CompB.Transform2);

            int Errors;
            Object[] ReturnedBodies;

            ReturnedBodies = BodyA.Operations2((int)swBodyOperationType_e.SWBODYINTERSECT, BodyB, out Errors);

            if (ReturnedBodies == null) { return false; }

            return true;
        }

        public static int ToolD = 10;
        public static int tDC = 5;

        //set the machining plan
        public Boolean SetAsMachiningPlan(List<RemovedBody> RemovalSequence)
        {
            if (RemovalSequence != null)
            {
                if (MachiningPlanList == null) { MachiningPlanList = new List<MachiningPlan>(); }
                List<MachiningProcess> MachiningSequences = new List<MachiningProcess>();
                Double MachiningTime = 0.0;
                Double MachiningVolume = 0.0;
                List<Object> Normals = new List<Object>();

                foreach (RemovedBody Sequence in RemovalSequence)
                {
                    MachiningProcess ThisProcess = new MachiningProcess();
                    ThisProcess.MachiningReference = Sequence.ParentPlane;
                    ThisProcess.TRV = Sequence.TRV;
                    ThisProcess.cuttingTool = ToolD.ToString();
                    ThisProcess.SelectedTAD = GetTAD(Sequence.ParentPlane.PlaneNormal);
                    ThisProcess.MachiningTime = Math.Round(CalMT(Sequence.TRV.Volume),2);
                    ThisProcess.MachiningVolume = Sequence.TRV.Volume * 1000000000;
                    MachiningSequences.Add(ThisProcess);
                    
                    MachiningTime = MachiningTime + ThisProcess.MachiningTime;
                    MachiningVolume = MachiningVolume + ThisProcess.MachiningVolume;
                    Normals.Add(Sequence.ParentPlane.PlaneNormal);
                }

                MachiningPlan ThisMachiningPlan = new MachiningPlan();
                ThisMachiningPlan.MachiningProceses = MachiningSequences;
            
                ThisMachiningPlan.NumberOfSetups = 1;
                ThisMachiningPlan.SetupNormal = "+Z";
                ThisMachiningPlan.NumberOfTADchanges = CalChange(Normals);
                ThisMachiningPlan.MachiningTime = MachiningTime + (ThisMachiningPlan.NumberOfTADchanges * tDC);
                ThisMachiningPlan.MachiningVolume = Math.Round(MachiningVolume, 2);
                MachiningPlanList.Add(ThisMachiningPlan);

                return true;

            }

            return false;

        }

        public Double GetCEdge(int ThisTool)
        { 
            switch (ThisTool)
            {
                case 4:
                    return 2;
                    
                case 8:
                    return 2;
                    
                case 10:
                    return 4;
            }

            return 0;
        }

        public Double CalMT(Double ThisVolume)
        {
            Double FeedSpeed = (1000 * 1) / (3.14 * ToolD * GetCEdge(ToolD));
            Double AxialD = 0.25 * ToolD;
            Double RadialD = 0.1 * ToolD;

            return (ThisVolume * 1000000000) / (AxialD * RadialD * FeedSpeed);

        }

        public TAD GetTAD(Object ThisNormal)
        {
            Double[] Normal = (Double[])ThisNormal;

            TAD SelectedTAD = new TAD();

            SelectedTAD.X = Normal[0];
            SelectedTAD.Y = Normal[1];
            SelectedTAD.Z = Normal[2];

            return SelectedTAD;

        }

        public int CalChange(List<Object> ThisNormals)
        {
            Double[] FirstNormal = null;
            Double[] SecondNormal = null;
            int[] AddNormal = new int[3];
            
            int Changes = 0;

            for (int i = 1; i < ThisNormals.Count(); i++)
            {                
                FirstNormal = (Double[])ThisNormals[i - 1];
                SecondNormal = (Double[])ThisNormals[i];
                AddNormal[0] = (int) (FirstNormal[0] - SecondNormal[0]); 
                AddNormal[1] = (int) (FirstNormal[1] - SecondNormal[1]);
                AddNormal[2] = (int) (FirstNormal[2] - SecondNormal[2]);
                
                if ((AddNormal[0] != 0) | (AddNormal[1] != 0) | (AddNormal[2] != 0)) { Changes++; }

            }

            return Changes;
        }

        //select body to delete after splitting *wby opening the document first
        public bool SelectBodyToDelete(Array BodyArray, string[] BodyNames, int index, ref List<int> BodyToDelete, ref List<double> VolumeSize)
        {
            bool retVal = false;
            
            for (int i = 0; i < BodyArray.Length; i++)
            {
                int Errors = 0;
                int Warnings = 0;

                SwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble

                ModelDoc2 ModDoc = (ModelDoc2)SwApp.OpenDoc6(BodyNames[i], (int)swDocumentTypes_e.swDocPART,
                    (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

                if (Errors == 0)
                {
                    iSwApp.SendMsgToUser("Loaded document: " + BodyNames[i]);

                    //get the current body in the document
                    PartDoc PartModDoc = (PartDoc)ModDoc;
                    Object[] ArrayOfBody = (Object[])PartModDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);

                    if (ArrayOfBody.Length == 1)
                    {
                        //get Centroid and Volume for the current body
                        Double[] Centroid = null;
                        
                        bool CVStatus = getCentroidAndVolume(ModDoc, ref Centroid, ref VolumeSize);

                        if (CVStatus == true)
                        {

                            Body2 BodyToCheck = (Body2)ArrayOfBody[0];

                            //check the feasibility of the splitted body (current)
                            bool bodyStatus = false;
                            //bodyStatus = setBodyToPlane(ModDoc, BodyToCheck, Centroid, PlaneListByScore[index]);

                            if (bodyStatus == true)
                            {
                                iSwApp.SendMsgToUser("This body is feasible (convex)");

                                bool SelecBody = BodyToCheck.Select(true, 0);

                                //check this body in the document tree and add collect the body pointer to the FeasibleBodies
                                BodyToDelete.Add(i+1);

                                //set this body to the current plane
                            }

                            else
                            {
                                iSwApp.SendMsgToUser("This body is not feasible (concave)");

                                //keep the index of the body

                            }
                        }
                    }
                    else
                    {
                        iSwApp.SendMsgToUser("Number of Bodies: " + ArrayOfBody.Length.ToString());
                    }

                    string ViewName = ModDoc.GetPathName();

                    iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(ViewName));

                }
            }

            if (BodyToDelete.Count != 0)
            {
                retVal = true;
            }
            
            return retVal;
        }

        //OVERLOAD selectbodytodelete
        //public static bool SelectBodyToDelete(ModelDoc2 CompDocumentModel, Array BodyArray, ref List<string> BodyToDelete, int index, ref List<double> VolumeSize, ref List<RemovalBody> RemoveThisBodies)
        public static bool SelectBodyToDelete(ModelDoc2 CompDocumentModel, Array BodyArray, int index, 
            ref List<RemovalBody> RemoveThisBodies, ref List<AddedReferencePlane> RelativePlanes)
        {
            
            List<string> BodyToDelete = new List<string>();
            List<double> VolumeSize = new List<double>();

            RemovalBody TmpRemovalBody = null;
            Boolean IsConvex;
            Boolean BodyLocation;
            int CorrectBody = 0;

            List<RemovalBody> TmpRemovalBodies = new List<RemovalBody>();
            RemoveThisBodies = new List<RemovalBody>();
            RelativePlanes = new List<AddedReferencePlane>();

            
            if (BodyArray.Length == 0)
            {
                return false;
            }

            else
            {   
                for (int i = 0; i < BodyArray.Length; i++)
                {
                    TmpRemovalBody = null;

                    bool bodyStatus = false;

                    Body2 BodyToCheck = (Body2)BodyArray.GetValue(i);

                    //iSwApp.SendMsgToUser("Check bodies Loaded document: " + BodyToCheck.Name);
                    Double[] Centroid = null;

                    bool CVStatus = getCentroidAndVolume(BodyToCheck, ref Centroid, ref VolumeSize);

                    if (CVStatus == true)
                    {

                        //get the pointerfor the body
                        PartDoc TmpPartDoc = (PartDoc)CompDocumentModel;
                        Array TmpBodyArray = (Array)TmpPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                        Body2 TmpBodyToCheck = null;

                        for (int j = 0; j < TmpBodyArray.Length; j++)
                        { 
                            TmpBodyToCheck = (Body2)TmpBodyArray.GetValue(j);

                            if (BodyToCheck.Name.Equals(TmpBodyToCheck.Name))
                            {
                                break;
                            }
                        }
                        
                        //check the feasibility of the splitted body (current)
                        IsConvex = false;
                        BodyLocation = false;

                        bodyStatus = setBodyToPlane(CompDocumentModel, TmpBodyToCheck, Centroid, PlaneListByScore[index], ref IsConvex, ref BodyLocation);

                        //count the body which is located on the correct side
                        if (BodyLocation == true) { CorrectBody++; }

                        if (bodyStatus == true)
                        {
                            //iSwApp.SendMsgToUser("This body is feasible (convex)");
                            //SelectData TmpSelectData = null;
                            //bool SelectionStatus = TmpBodyToCheck.Select2(true, TmpSelectData);

                            //check this body in the document tree and add collect the body pointer to the FeasibleBodies
                            
                            //BodyToDelete.Add(BodyToCheck.Name);

                            //get the faces of the BOdyToCheck
                            Object[] FaceObj = (Object[])TmpBodyToCheck.GetFaces();

                            //check the other reference planes which are located on the trv body
                            for (int IndexToFilter = 0; IndexToFilter < PlaneListByScore.Count(); IndexToFilter++)
                            {
                                if (PlaneListByScore[IndexToFilter].Remark == 0
                                    && PlaneListByScore[IndexToFilter].name != PlaneListByScore[index].name)
                                {
                                    //check whether if any ref planes are located on the BodyToCheck
                                    foreach (Face2 ThisFace in FaceObj)
                                    {
                                        Object ThisFaceTess = (Object)ThisFace.GetTessTriangles(true);
                                        Double[] ThisFaceMaxMin = GetMaxMin(ThisFaceTess);

                                        if (IsTessEqual((Double[])PlaneListByScore[IndexToFilter].MaxMinValue, ThisFaceMaxMin) == true)
                                        {
                                            setMarkOnPlane(ref PlaneListByScore, IndexToFilter, "PARENTRELATIVE");
                                            RelativePlanes.Add(PlaneListByScore[IndexToFilter]);
                                        }
                                    }

                                }
                            }


                            TmpRemovalBody = new RemovalBody();
                            TmpRemovalBody.Name = BodyToCheck.Name;
                            TmpRemovalBody.Volume = VolumeSize[i];
                            TmpRemovalBody.BodyObj = BodyToCheck;
                            TmpRemovalBody.IsConvex = IsConvex;
                            TmpRemovalBody.Location = BodyLocation;

                            TmpRemovalBodies.Add(TmpRemovalBody);

                            //set this body to the current plane
                        }

                        else
                        {
                            //iSwApp.SendMsgToUser("This body is not feasible (concave)");

                            //keep the index of the body

                        }
                    }

                    //if (TmpRemovalBody != null) { RemoveThisBodies.Add(TmpRemovalBody); }

                }

                //if (TmpRemovalBodies.Count() == CorrectBody)
                //{
                //    foreach (RemovalBody ThisBody in TmpRemovalBodies)
                //    {
                //        RemoveThisBodies.Add(ThisBody);
                //    }
                //}
                if (TmpRemovalBodies.Count() != 0)
                {
                    List<RemovalBody> ThisBodyResults = new List<RemovalBody>();

                    if (CheckValidBodies(TmpRemovalBodies, true, true, out ThisBodyResults) == true ||
                        CheckValidBodies(TmpRemovalBodies, true, false, out ThisBodyResults) == true)
                    {
                        RemoveThisBodies.AddRange(ThisBodyResults);
                    }
                }
                else
                {
                    TmpRemovalBodies.Clear();
                }

            }

            if (RemoveThisBodies.Count == 0) { return false; }

            return true;

            //if (BodyToDelete.Count == 0) { return false; }
            //else 
            //{
            //    for (int i = 0; i < BodyToDelete.Count(); i++)
            //    {
            //        TmpRemovalBody = new RemovalBody();
            //        TmpRemovalBody.Name = BodyToDelete[i];
            //        TmpRemovalBody.Volume = VolumeSize[i];
            //        TmpRemovalBody.BodyObj = (Body2)BodyArray.GetValue(i);

            //        RemoveThisBodies.Add(TmpRemovalBody);
            //    }

            //    return true; 
            //}

        }

        //OVERLOAD selectbodytodelete
        //public static bool SelectBodyToDelete(ModelDoc2 CompDocumentModel, Array BodyArray, ref List<string> BodyToDelete, int index, ref List<double> VolumeSize, ref List<RemovalBody> RemoveThisBodies)
        public static bool SelectBodyToDelete(ModelDoc2 CompDocumentModel, Array BodyArray, int index,
            ref List<RemovalBody> RemoveThisBodies)
        {

            List<string> BodyToDelete = new List<string>();
            List<double> VolumeSize = new List<double>();

            RemovalBody TmpRemovalBody = null;
            Boolean IsConvex;
            Boolean BodyLocation;
            int CorrectBody = 0;

            List<RemovalBody> TmpRemovalBodies = new List<RemovalBody>();
            RemoveThisBodies = new List<RemovalBody>();
            

            if (BodyArray.Length == 0)
            {
                return false;
            }

            else
            {
                for (int i = 0; i < BodyArray.Length; i++)
                {
                    TmpRemovalBody = null;

                    bool bodyStatus = false;

                    Body2 BodyToCheck = (Body2)BodyArray.GetValue(i);

                    //iSwApp.SendMsgToUser("Check bodies Loaded document: " + BodyToCheck.Name);
                    Double[] Centroid = null;

                    bool CVStatus = getCentroidAndVolume(BodyToCheck, ref Centroid, ref VolumeSize);

                    if (CVStatus == true)
                    {

                        //get the pointerfor the body
                        PartDoc TmpPartDoc = (PartDoc)CompDocumentModel;
                        Array TmpBodyArray = (Array)TmpPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                        Body2 TmpBodyToCheck = null;

                        for (int j = 0; j < TmpBodyArray.Length; j++)
                        {
                            TmpBodyToCheck = (Body2)TmpBodyArray.GetValue(j);

                            if (BodyToCheck.Name.Equals(TmpBodyToCheck.Name))
                            {
                                break;
                            }
                        }

                        //check the feasibility of the splitted body (current)
                        IsConvex = false;
                        BodyLocation = false;

                        bodyStatus = setBodyToPlane(CompDocumentModel, TmpBodyToCheck, Centroid, PlaneListByScore[index], ref IsConvex, ref BodyLocation);

                        //count the body which is located on the correct side
                        if (BodyLocation == true) { CorrectBody++; }

                        if (bodyStatus == true)
                        {
                            //iSwApp.SendMsgToUser("This body is feasible (convex)");
                            //SelectData TmpSelectData = null;
                            //bool SelectionStatus = TmpBodyToCheck.Select2(true, TmpSelectData);

                            //check this body in the document tree and add collect the body pointer to the FeasibleBodies

                            //BodyToDelete.Add(BodyToCheck.Name);

                            TmpRemovalBody = new RemovalBody();
                            TmpRemovalBody.Name = BodyToCheck.Name;
                            TmpRemovalBody.Volume = VolumeSize[i];
                            TmpRemovalBody.BodyObj = BodyToCheck;

                            TmpRemovalBodies.Add(TmpRemovalBody);

                            //set this body to the current plane
                        }

                        else
                        {
                            //iSwApp.SendMsgToUser("This body is not feasible (concave)");

                            //keep the index of the body

                        }
                    }

                    //if (TmpRemovalBody != null) { RemoveThisBodies.Add(TmpRemovalBody); }

                }

                if (TmpRemovalBodies.Count() == CorrectBody)
                {
                    foreach (RemovalBody ThisBody in TmpRemovalBodies)
                    {
                        RemoveThisBodies.Add(ThisBody);
                    }
                }
                else
                {
                    TmpRemovalBodies.Clear();
                }

            }

            if (RemoveThisBodies.Count == 0) { return false; }

            return true;

            //if (BodyToDelete.Count == 0) { return false; }
            //else 
            //{
            //    for (int i = 0; i < BodyToDelete.Count(); i++)
            //    {
            //        TmpRemovalBody = new RemovalBody();
            //        TmpRemovalBody.Name = BodyToDelete[i];
            //        TmpRemovalBody.Volume = VolumeSize[i];
            //        TmpRemovalBody.BodyObj = (Body2)BodyArray.GetValue(i);

            //        RemoveThisBodies.Add(TmpRemovalBody);
            //    }

            //    return true; 
            //}

        }
        
        //check valid bodies
        private static bool CheckValidBodies(List<RemovalBody> ThisBodiesList, bool Convexity, bool Location, 
            out List<RemovalBody> BodyResults)
        {
            BodyResults = ThisBodiesList.Where(Parent => Parent.IsConvex.Equals(Convexity) && Parent.Location.Equals(Location)).ToList();

            if (BodyResults.Count != 0) { return true; }

            return false;
        }

        //check similarity between two array of doubles
        public static bool IsTessEqual(Double[] TessA, Double[] TessB)
        {
            for (int i = 0; i < TessA.Count(); i++)
            {
                if (isEqual(TessA[i], TessB[i]) == false)
                {
                    return false;
                }
            }

            return true;
        }

        //split and save bodies
        public static Feature SplitAndSaveBody(ModelDoc2 DocumentModel, Component2 SwComp, Feature SelectedFeature, Array BodyArray, ref string[] BodyNames)
        { 
            Body2[] bodyCandidate = new Body2[BodyArray.Length];
            Vertex[] bodyOrigins = new Vertex[BodyArray.Length];
            
            //get the component model path to set as the file name
            string savePath = getSavePath(SwComp.GetPathName());
                                        
            //collect all the body which is created from the selected plane
            for (int i = 0; i < BodyArray.Length; i++)
            {
                bodyCandidate[i] = BodyArray.GetValue(i) as Body2;
                //BodyNames[i] = savePath + SelectedFeature.Name.ToString() + "-" + i + ".sldprt";
                BodyNames[i] = null;
                bodyOrigins[i] = null;
            }
                    
            return (Feature)DocumentModel.FeatureManager.PostSplitBody(bodyCandidate, false, bodyOrigins, BodyNames);
        }

        //set and return the path for saving the body collections
        public static string getSavePath(string modelPath)
        { 
            string modelDirectory = "";
            string modelFileName = "";

            modelDirectory = Path.GetDirectoryName(modelPath);
            modelFileName = Path.GetFileNameWithoutExtension(modelPath);

            return modelDirectory + "\\" + modelFileName;
        }

        //get the centroid of the current body
        public bool getCentroidAndVolume(ModelDoc2 DocumentModel, ref double[] Centroid, ref List<double> Volume)
        {
            int status = 0;

            Double[] MassProperty = null;
            
            MassProperty = (Double[])DocumentModel.Extension.GetMassProperties(0, ref status);

            if (status == 0)
            {

                Double dimension = 1; //change it to mmm by times it with 1000 (current is in m)

                //get the centroid
                Centroid = new Double[3];
                Centroid[0] = MassProperty[0] * dimension; 
                Centroid[1] = MassProperty[1] * dimension;
                Centroid[2] = MassProperty[2] * dimension;

                //get the volume
                Volume.Add(MassProperty[3]);

                return true;
            }
            
            return false;
        }

        //OVERLOAD
        public static bool getCentroidAndVolume(Body2 BodyToCheck, ref double[] Centroid, ref List<double> Volume)
        {
            int status = 0;

            Double MaterialProperty = iSwApp.GetUserPreferenceDoubleValue((int) swUserPreferenceDoubleValue_e.swMaterialPropertyDensity);            

            Double[] MassProperty = null;

            MassProperty = (Double[])BodyToCheck.GetMassProperties(MaterialProperty);

            if (status == 0)
            {

                Double dimension = 1; //change it to mmm by times it with 1000 (current is in m)

                //get the centroid
                Centroid = new Double[3];
                Centroid[0] = MassProperty[0] * dimension;
                Centroid[1] = MassProperty[1] * dimension;
                Centroid[2] = MassProperty[2] * dimension;

                //get the volume
                Volume.Add(MassProperty[3]);

                return true;
            }

            

            return false;
        }
    
        //switch path
        public string getSplitPath(string name, int index)
        { 
            switch (name)
            {
                case "#plane1":
                    return "Split1[2]@rm_120x70x65_mori-1@tes_moritrv";
                case "#plane2":
                    if (index == 0)
                    {
                        return "Split3[2]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                    else
                    {
                        return "Split4[2]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                case "#plane3":
                    return "";
                case "#plane4":
                    return "";
                case "#plane5":
                    if (index == 2)
                    {
                        return "Split3[2]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                    else
                    {
                        return "Split2[2]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                case "#plane6":
                    return "";
                case "#plane7":
                    return "";
                case "#plane8":
                    return "";
                case "#plane9":
                    if (index == 0)
                    {
                        return "Split4[1]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                    else if (index == 1)
                    {
                        return "Split3[1]@rm_120x70x65_mori-1@tes_moritrv";
                    }
                    else
                    {
                        return "Split2[1]@rm_120x70x65_mori-1@tes_moritrv";
                    }
            }

            return "";
        }

        /*
        //initialize machining plane... example
        public void initMachiningPlan()
        {
            machiningPlanList = new List<Object>();
            List<machiningPlan> machiningCandidate = new List<machiningPlan> ();

            machiningPlan tmpMP = null;

            //first plan
            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[0];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);
            
            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[3];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, -1, 0 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[1];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[7];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[2];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            machiningPlanList.Add((machiningPlan[])machiningCandidate.ToArray());

            
            //Second plan
            machiningCandidate = new List<machiningPlan>();

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[0];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[3];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, -1, 0 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[7];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[1];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[2];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            machiningPlanList.Add((machiningPlan[])machiningCandidate.ToArray());

            //Third Plan
            machiningCandidate = new List<machiningPlan>();

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[0];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[7];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[3];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, -1, 0 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[1];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Zigzag";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            tmpMP = new machiningPlan();
            tmpMP.planeProperties = planeList[2];
            tmpMP.cuttingTool = "Flat EndMill";
            tmpMP.toolPath = "Countoring";
            tmpMP.TAD = new Double[3] { 0, 0, 1 };
            machiningCandidate.Add(tmpMP);

            machiningPlanList.Add((machiningPlan[])machiningCandidate.ToArray());

        }
         * */


        //get and return the plane
        public Feature getByMP(AddedReferencePlane plane, List<Feature> _featureList)
        {
            if (plane != null)
            {
                return getSelectedPlane(plane, _featureList);
            }

            return null;
        }

        //get and return the plane
        public Feature getThePlane(ref List<AddedReferencePlane> PlaneListIn, List<Feature> _featureList, ref int index, List<AddedReferencePlane> ParentList)
        {
            //select the index that has no marking option
            index = 0;
            while (index != -1)
            {
                if (PlaneListIn[index].MarkingOpt == 0) 
                {
                    if (ParentList == null) {break;}
                    else if (IsWithin(PlaneListIn[index], ParentList) == false)
                    {
                        break; 
                    }
                }

                index++;
                if (index == PlaneListIn.Count) { index=-1; }
            }

            if (index != -1)
            {
                PlaneListIn[index].MarkingOpt = -1;
                return getSelectedPlane(PlaneListIn[index], _featureList);
            }

            return null;
        }

        //OVERLOAD get and return the plane
        public Feature getThePlane(ref List<AddedReferencePlane> PlaneListIn, List<AddedReferencePlane> PlaneListDist, 
            List<Feature> _featureList, ref int index, List<AddedReferencePlane> ParentList, List<AddedReferencePlane> RelativeList)
        {
            //select the index that has no marking option
            index = 0;
            while (index != -1)
            {
                if (PlaneListIn[index].MarkingOpt == 0)
                {
                    if (ParentList == null)
                    {
                        ////if (IsOuter(PlaneListIn[index], PlaneListDist) == true)
                        if (PlaneListIn[index].isOuter == true || PlaneListIn[index].CPost == true)
                        {   
                            break;
                        }
                    }
                    else if (IsWithin(PlaneListIn[index], ParentList) == false && IsWithin(PlaneListIn[index], RelativeList) == false)
                    {
                        break;
                    }
                }

                index++;
                if (index == PlaneListIn.Count) { index = -1; }
            }

            if (index != -1)
            {
                PlaneListIn[index].MarkingOpt = -1;
                return getSelectedPlane(PlaneListIn[index], _featureList);
            }

            return null;
        }

        //check the distance of reference plane
        public Boolean IsOuter(AddedReferencePlane ThisRefPlane, List<AddedReferencePlane> ListOfRefPlane)
        {
            int ThisRefIndex;
            int index = 0;

            ThisRefIndex = ListOfRefPlane.IndexOf(ThisRefPlane);

            foreach (AddedReferencePlane CheckThisRef in ListOfRefPlane)
            {
                if (CheckThisRef.name == ThisRefPlane.name)
                {
                    ThisRefIndex = index;

                    break;
                }

                index++;
            }

            if (ListOfRefPlane[ThisRefIndex].DistanceFromCentroid == ListOfRefPlane[0].DistanceFromCentroid)
            {
                return true;
            }

            return false;
        }

        //check if it is a parent or not
        public Boolean IsWithin(AddedReferencePlane CheckThisPlane, List<AddedReferencePlane> ListReference)
        {
            int Check = ListReference.Where(Parent => Parent.name.Equals(CheckThisPlane.name)).Count();

            if (Check == 0) { return false; }

            return true;
        }

        //get the similar plane and pass the pointer to selected feature
        public static Feature getSelectedPlane(AddedReferencePlane listOfPlane, List<Feature> listOfModelPlane)
        {

            foreach (Feature CheckedPlane in listOfModelPlane)
            {
                if (CheckedPlane.Name == listOfPlane.CorrespondFeature.Name)
                {
                    return CheckedPlane;
                }
            }

            return null;
        }

        //get and return the plane index (this assumes that the planes are already ordered by its score)
        public static int getPlaneIndex(AddedReferencePlane ParentPlane, List<AddedReferencePlane> PlaneListIn)
        {   
            for (int index = 0; index < PlaneListIn.Count(); index++)
            {
                if (PlaneListIn[index].name.Equals(ParentPlane.name)) { return index; }
            }

            return -1;
        }

        //mark the plane according the message code
        public static bool setMarkOnPlane(ref List<AddedReferencePlane> planeList, int index, string message)
        {
            switch (message)
            {
                //if the process runs ok, mark the plane
                //as 0 to avoid the system select it again    
                case "PROCESSOK":
                    planeList[index].Remark = 0;

                    //renewIntersection(ref planeList, index);
                    break;

                //if no body found
                case "NOBODYFOUND":
                    planeList[index].Remark = 1;
                    break;

                //if no body can be set to the corresponding plane
                case "NOBODYSET":
                    planeList[index].Remark = 2;
                    break;

                //split body process has failed
                case "EMPTYSPLIT":
                    planeList[index].Remark = 3;
                    break;

                //safe normal critearia is not achieved
                case "UNSAFE":
                    planeList[index].Remark = 4;
                    break;

                case "PARENTRELATIVE":
                    planeList[index].Remark = 5;
                    break;
            }
            
            return true;
        }

        public void renewIntersection(ref List<_planeProperties> planeList, int removeID)
        {
            planeList.RemoveAt(removeID);
            
            List<Object> planeNormal = new List<Object>();
            List<int> planeValue = null;

            foreach (_planeProperties tmpPlane in planeList)
            {
                planeNormal.Add(tmpPlane.planeNormal);
            }

            planeIntersection(planeNormal, ref planeValue);

            for (int i = 0; i < planeList.Count; i++)
            {
                planeList[0].rankByDistance = planeValue[i];
            }

        }

        // the reference plane belonging body
        //public bool setBodyToPlane(ModelDoc2 Doc, Body2[] bodyList, _planeProperties selectedPlane, ref List<int> selectedId)
        //{
        //    List<Body2> tmpBodyList = new List<Body2>();
        //    selectedId = new List<int>();

        //    bool returnValue = false;
            
        //    try
        //    {
        //        int index = 0;

        //        foreach (Body2 tmpBody in bodyList)
        //        {
        //            if (isBodyConvex(Doc, tmpBody) == true)
        //            {
        //                if (checkBodyLocation(tmpBody, selectedPlane.featureObj, selectedPlane.planeNormal) == true)
        //                {
        //                    tmpBodyList.Add(tmpBody);
        //                    selectedId.Add(index);
        //                }

        //            }

        //            index++;
                    
        //        }
        //    }

        //    finally
        //    {
        //        if (tmpBodyList.Count > 0)
        //        {
        //            //tmpBodyList.CopyTo(selectedPlane.bodyList);
        //            selectedPlane.bodyList = tmpBodyList;
        //            returnValue = true;
        //        }
        //    }

        //    return returnValue;
        //}



        //OVERLOAD setBodyToPlane
        public static bool setBodyToPlane(ModelDoc2 DocumentModel, Body2 BodyToCheck, Double[] Centroid, AddedReferencePlane SelectedPlane, 
            ref Boolean Convexity, ref Boolean Location)
        {

            if (isBodyConvex(DocumentModel, BodyToCheck, SelectedPlane) == true)
            {
                Convexity = true;
            }
            
            if (checkBodyLocation(BodyToCheck, Centroid, SelectedPlane.CorrespondFeature, SelectedPlane.PlaneNormal) == true)
            {
                Location = true;
            }

            //first case to accomodate the regular case, if two condition (convex and correct location) are ok
            if ((Convexity == true) && (Location== true))
            {
                return true;
            }
            
            //second case to accomodate the safe plane but has wrong direction of normal
            if ((Convexity == true) && (Location == false) && (SelectedPlane.CPost == true))
            {
                return true;
            }

            return false;
        }

        //check whether if the body convex or concave, return true if convex
        public static bool isBodyConvex(ModelDoc2 Doc, Body2 body2Check, AddedReferencePlane SelectedPlane)
        {
            //please add here to define whether the body is convex or not

            //if (CheckTopologicalData(Doc, body2Check) == true)
            //{
            //    return true;
            //}

            //else
            //    return false;

            object[] objFaces = null;

            objFaces = (object[])body2Check.GetFaces();
            List<Face2> filteredFaces = new List<Face2>();

            if (filterFaces(Doc, objFaces, SelectedPlane) == true)
            {
                foreach (Face2 tmpFace in objFaces)
                {
                    if (isFaceConvex(tmpFace) == false)
                    {
                        return false;
                    }
                }
            }
            else { return false; }

            return true;
        }

        //filter the face
        public static bool filterFaces(ModelDoc2 Doc, object[] faceList, AddedReferencePlane SelectedPlane)
        {   
            List<Face2> tmpFaces = new List<Face2>();
            List<bool> FacesLocation = new List<bool>();

            foreach (Face2 faceA in faceList)
            {   
                FacesLocation.Add(isFaceCollinear(faceA, SelectedPlane));
            }

            //get all face that is true (collinear with the reference plane)
            List<bool> CollinearFace = FacesLocation.Where(value => value.Equals(true)).ToList();

            if (CollinearFace.Count == 1) { return true; }
            //if (CollinearFace.Count > 1) { iSwApp.SendMsgToUser("More than two faces are collinear with this reference plane"); }

            return false;
        }

        //check the location of the face
        public static bool isFaceCollinear(Face2 Face2Check, AddedReferencePlane SelectedPlane)
        {
            Object[] objLoops = null;
            objLoops = (Object[])Face2Check.GetLoops();

            foreach (Loop2 tmpLoop in objLoops)
            {
                if (tmpLoop.IsOuter() == true)
                {
                    //process this loop
                    return isLoopOnThePlane(tmpLoop, SelectedPlane);
                }
            }

            return false;
        }

        //check the location of the face loop
        public static bool isLoopOnThePlane(Loop2 loop2Check, AddedReferencePlane SelectedPlane)
        {
            Object[] objVertices = null;
            List<Vertex> verticesList = new List<Vertex>();
            
            objVertices = (Object[])loop2Check.GetVertices();

            foreach (Vertex tmpVertex in objVertices)
            {
                if (CheckPointLocation(tmpVertex, SelectedPlane.ReferencePlane, SelectedPlane.PlaneNormal) == false)
                {
                    return false;
                }
            }

            return true;
        }
    
        //check whether if the face convex or concave
        public static bool isFaceConvex(Face2 face2Check)
        {
            Object[] objLoops = null;
            objLoops = (Object[])face2Check.GetLoops();

            if (IsPlanar(face2Check) == false) { return true; } //it meant for non planar face

            foreach (Loop2 tmpLoop in objLoops)
            {
                if (tmpLoop.IsOuter() == true)
                {
                    //process this loop
                    return isLoopConvex(tmpLoop, face2Check);
                }
            }

            return false;
        }

        //check whether if the loop is convex or concave
        public static bool isLoopConvex(Loop2 loop2Check, Face2 face2Check)
        {
            Object[] objVertices = null;
            List<Vertex> verticesList = new List<Vertex>();
            Vertex[] verticesArray = null;
            double[] zCrossProduct = null;
            Object[] vEdgeArr = null;
            PatternConfig ThisFacePattern = new PatternConfig();
            Curve ThisCurve = null;

            vEdgeArr = (Object[])loop2Check.GetEdges();

            foreach (Edge tmpEdge in vEdgeArr)
            {
                ThisCurve = tmpEdge.GetCurve();

                if (ThisCurve.IsLine()) { ThisFacePattern.Line++; }
                else if (ThisCurve.IsCircle()) { ThisFacePattern.Circle++; }
                else { ThisFacePattern.Curve++; }

            }

            if (GetPatternType(ThisFacePattern) == "MIX")
            {
                if (ThisFacePattern.Line != 0 && ThisFacePattern.Line > 4) { return false; }
                else { return true; }
            }

            if (GetPatternType(ThisFacePattern) == "CURVE") 
            { 
                return true; 
            }

            if (GetPatternType(ThisFacePattern) == "CIRCLE")
            {
                return true;
            }
                        
            objVertices = (Object[])loop2Check.GetVertices();

            foreach (Vertex tmpVertex in objVertices)
            {
                verticesList.Add(tmpVertex);
            }

            //add additional vertices
            verticesList.Add(objVertices[0] as Vertex);
            verticesList.Add(objVertices[1] as Vertex);

            verticesArray = new Vertex[verticesList.Count];
            verticesList.CopyTo(verticesArray);

            if (checkTriplets(ref zCrossProduct, verticesArray, face2Check) == true)
            {
                return checkSign(zCrossProduct);
            }
            
            return false;
        }

        //check triplets
        public static bool checkTriplets(ref double[] UVCrossProduct, Vertex[] vertexList, Face2 face2Check)
        {
            Object objPoint1, objPoint2, objPoint3 = null;
            Double[] point1, point2, point3 = null;
            //Vertex tmpVertex1, tmpVertex2, tmpVertex3 = null;
            double dx1, dx2, dy1, dy2;
            double du1, du2, dv1, dv2;

            //crossProduct = new double[vertexList.Length - 2];
            UVCrossProduct = new double[vertexList.Length - 2];
            
            //iterate to check each triplets
            for (int i = 0; i < (vertexList.Length - 2); i++)
            {
                //tmpVertex1 = (Vertex)vertexList[i];
                //tmpVertex2 = (Vertex)vertexList[i + 1];
                //tmpVertex3 = (Vertex)vertexList[i + 2];

                objPoint1 = vertexList[i].GetPoint();
                objPoint2 = vertexList[i + 1].GetPoint();
                objPoint3 = vertexList[i + 2].GetPoint();

                point1 = (double[])objPoint1;
                point2 = (double[])objPoint2;
                point3 = (double[])objPoint3;

                Double[] UVPoint1 = (Double[])face2Check.ReverseEvaluate(point1[0], point1[1], point1[2]);
                Double[] UVPoint2 = (Double[])face2Check.ReverseEvaluate(point2[0], point2[1], point2[2]);
                Double[] UVPoint3 = (Double[])face2Check.ReverseEvaluate(point3[0], point3[1], point3[2]);

                dx1 = point2[0] - point1[0];
                dx2 = point3[0] - point2[0];
                dy1 = point2[1] - point1[1];
                dy2 = point3[1] - point2[1];

                du1 = UVPoint2[0] - UVPoint1[0];
                du2 = UVPoint3[0] - UVPoint2[0];
                dv1 = UVPoint2[1] - UVPoint1[1];
                dv2 = UVPoint3[1] - UVPoint2[1];

                //crossProduct[i] = (dx1 * dy2) - (dx2 * dy1);

                UVCrossProduct[i] = (du1 * dv2) - (du2 * dv1);
            }

            return true;
        }

        //check the sign of vertex
        public static bool checkSign(double[] crossProduct)
        {
            int PositiveNum = 0;
            int NegativeNum = 0;
            bool returnValue = false;
            
            //check number of positive and negative value
            if (crossProduct != null)
            {   
                for (int i = 0; i < crossProduct.Length; i++)
                {
                    if (crossProduct[i] != 0)
                    {
                        if (crossProduct[i] > 0)
                        {
                            PositiveNum++;
                        }
                        else
                        {
                            NegativeNum++;
                        }
                    }
                    
                }

                //if either positive or negative leads to zero it means the sign never changes                
                if (PositiveNum == 0 | NegativeNum==0) 
                {
                    returnValue = true;
                }
            }
            return returnValue;
        }

        //compare two vertices, return true if equal
        public bool compareVertices(Object vertex1, Object vertex2)    
        {
            double[] tmpVertex1 = null;
            double[] tmpVertex2 = null;
            bool returnValue = false;

            tmpVertex1 = (double[])vertex1;
            tmpVertex2 = (double[])vertex2;

            if (isEqual(tmpVertex1[0], tmpVertex2[0]) == true) { returnValue = true; }
            else { returnValue = false; }

            if (isEqual(tmpVertex1[1], tmpVertex2[1]) == true) { returnValue = true; }
            else { returnValue = false; }
            
            if (isEqual(tmpVertex1[1], tmpVertex2[1]) == true) { returnValue = true; }
            else { returnValue = false; }

            return returnValue;
        }

        //check if the feature is in the normal direction
        public bool checkBodyLocation(Body2 body2Check, Feature featurePlane, object objPlaneNormal)
        {
            double[] boxCentroid = null;
            object bodyVertices = null;
            MathVector planeNormal, tmpMathVector = null;
            MathPoint pointOnPlane, pointOnBox = null;
            Double returnLocation = 0;
            MathUtility swMath = (MathUtility)SwApp.GetMathUtility();// = new MathUtility();

            //get the point on plane
            pointOnPlane = getPointOnPlane(featurePlane);
            
            //get the centroid of the body box
            bodyVertices = (Object) body2Check.GetBodyBox();            
            boxCentroid = getMidPoint(bodyVertices);
            pointOnBox = swMath.CreatePoint(boxCentroid);
            
            //get plane normal
            //planeNormal = (MathVector) sw objPlaneNormal;
            planeNormal = swMath.CreateVector(objPlaneNormal);

            //find the dot product of normal and the subtracted value from point on plane with point on box
            tmpMathVector = (MathVector) pointOnPlane.Subtract(pointOnBox);
            //tmpMathPoint = tmpPoint as MathPoint;
            //pointSubstration = (MathVector)tmpMathPoint.ConvertToVector();
            //returnLocation = planeNormal.Dot(pointSubstration);
            returnLocation = planeNormal.Dot(tmpMathVector);

            if (returnLocation > 0)
            { 
                return true;
            }
            else
            {
                return false;
            }

        }

        //OVERLOAD CheckBodyLocation
        public static bool checkBodyLocation(Body2 body2Check, Double[] Centroid, Feature featurePlane, object objPlaneNormal)
        {
            MathVector planeNormal, tmpMathVector = null;
            MathPoint pointOnPlane, pointOnBox = null;
            Double returnLocation = 0;
            MathUtility swMath = (MathUtility)SwApp.GetMathUtility();// = new MathUtility();

            //get the point on plane
            pointOnPlane = getPointOnPlane(featurePlane);
            double[] TmpPointOnBox = (double[])pointOnPlane.ArrayData;

            //get the centroid of the body box
            pointOnBox = swMath.CreatePoint(Centroid);
            double[] TmpPointBox = (double[])pointOnBox.ArrayData;

            //get plane normal
            //planeNormal = (MathVector) sw objPlaneNormal;
            planeNormal = swMath.CreateVector(objPlaneNormal);
            double[] TmpPlaneNormal = (double[])planeNormal.ArrayData;

            //find the dot product of normal and the subtracted value from point on plane with point on box
            //tmpMathVector = (MathVector)pointOnPlane.Subtract(pointOnBox);
            tmpMathVector = (MathVector)pointOnBox.Subtract(pointOnPlane);
            //tmpMathPoint = tmpPoint as MathPoint;
            //pointSubstration = (MathVector)tmpMathPoint.ConvertToVector();
            //returnLocation = planeNormal.Dot(pointSubstration);
            returnLocation = planeNormal.Dot(tmpMathVector);

            if (returnLocation > 0)
            {
                return true;
            }
            else
            {
                return false;
            }

        }

        //get random point on a reference plane
        public static MathPoint getPointOnPlane(Feature tmpFeature)
        {
            MathPoint tmpPoint = null;
            Object featDefinition = null;
            RefPlane featurePlane = null;
            Object[] arrayOfCorners = null;
            Random rnd = new Random();
            string name = tmpFeature.Name;

            featDefinition = (Object)tmpFeature.GetDefinition();
            featurePlane = (RefPlane)tmpFeature.GetSpecificFeature2();//featDefinition;
            arrayOfCorners = (Object[]) featurePlane.CornerPoints;
            tmpPoint = (MathPoint)arrayOfCorners[rnd.Next(0, 3)];

            

            List<double[]> listDouble = new List<double[]>();
            double[] pointDouble = (double[])tmpPoint.ArrayData;

            foreach (Object tmpObject in arrayOfCorners)
            {
                MathPoint pointObs = (MathPoint)tmpObject;
                listDouble.Add((double[])pointObs.ArrayData);
            }

            return tmpPoint;
        }

        //OVERLOAD getPointOnPlane
        public static MathPoint getPointOnPlane(RefPlane tmpRefPlane)
        {
            MathPoint tmpPoint = null;
            Object[] arrayOfCorners = null;
            Random rnd = new Random();
            
            arrayOfCorners = (Object[])tmpRefPlane.CornerPoints;
            tmpPoint = (MathPoint)arrayOfCorners[rnd.Next(0, 3)];

            return tmpPoint;
        }

        //OVERLOAD getPointOnPlane
        public static MathPoint getPointOnPlane(AddedReferencePlane ThisRefPlane)
        {
            Double[] MaxMinValue = null;
            Random rnd = new Random();

            MathUtility MathUtils;

            MathUtils = SwApp.GetMathUtility();
            MaxMinValue = (Double[])ThisRefPlane.MaxMinValue;           
            
            return MathUtils.CreatePoint(new Double[] {MaxMinValue[0], MaxMinValue[1], MaxMinValue[2]});
        }
        
        //set the machining attribute
        public static bool MachiningAttribute()
        {


            return true;
        }

        #endregion

        //Machinable Space
        public void MachinableSpace()
        {

            iSwApp.SendMsgToUser("got from machinable space");

            //set the table region
            TableRegion();

            //insert all table region into table assembly


        }

        //tools for machinable space calculation
        #region MachinableSpace
        
        //create VC requirement
        public void VirtualConeAnalysis()
        { 
            ModelDoc2 Doc = (ModelDoc2) SwApp.ActiveDoc;
            
            int docType = (int)Doc.GetType();

            if (Doc.GetType() == 2)
            {
                AssemblyDoc ThisAssyDoc = (AssemblyDoc)Doc;

                if (ThisAssyDoc.GetComponentCount(true) == 0) { return; }
                else
                {
                    ProcessLog_TaskPaneHost.LogProcess("Analysing open faces of current assembly model");
                    PPDetails_TaskPaneHost.LogProcess("Analysing open faces of current assembly model");

                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);
                    Object[] childComp = (Object[])compInAssembly.GetChildren();

                    //classify between the workpiece and trvs
                    List<Component2> ThisComponents = new List<Component2>();
                    Component2 RootComponent = null;
                    for (int i = 0; i < ThisAssyDoc.GetComponentCount(true); i++)
                    {
                        Component2 TmpComponent = (Component2)childComp[i];
                        if (TmpComponent.Name.Contains("workpiece"))
                        {
                            RootComponent = TmpComponent;
                            RootComponent.Select2(true, 0);
                            ThisAssyDoc.SetComponentTransparent(true);
                        }
                        else
                        {
                            ThisComponents.Add(TmpComponent);
                        }
                    }

                    //order it by its name (may be triggering problem if the components more than 10)
                    List<Component2> ThisComponentsSorted = ThisComponents.OrderBy(c => c.Name2).ToList();

                    //get the product/workpiece (main component) faces
                    Array CompBodyArray = (Array)RootComponent.GetBodies2((int)swBodyType_e.swSolidBody);
                    Body2 MainCompBody = (Body2)CompBodyArray.GetValue(0);
                    Object[] MainCompFaces = (Object[])MainCompBody.GetFaces();
                }


            }
        }

        //create the table region for several B-axis magnitude
        public void TableRegion()
        { 

        }

        //insert the table region to same assembly file
        public void InsertTableRegion()
        { 

        }

        #endregion

        //generate the TRV network
        public void TRVNetwork()
        {
            iSwApp.SendMsgToUser("got from TRV Network");
        }

        //tools for calculating TRV network
        #region TRVNetwork



        #endregion

        //Setup Calculator
        public void SetupCalculator()
        {
            iSwApp.SendMsgToUser("got from Setup calculator");
            ModelDoc2 swDoc = null;
            PartDoc swPart = null;
            DrawingDoc swDrawing = null;
            AssemblyDoc swAssembly = null;
            bool boolstatus = false;
            int longstatus = 0;
            int longwarnings = 0;
            object[] ObjBodies;
            Body2 SelectedBody = null;
            bool SelectStatus;

            swDoc = (ModelDoc2)SwApp.ActiveDoc;
            
            swAssembly = ((AssemblyDoc)(swDoc));

            //get the components
            ConfigurationManager configDocManager = (ConfigurationManager) swDoc.ConfigurationManager;
            Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
            Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

            if (compName == null)
            {
                //define the raw material and product
                compName = new Component2[2];
                GetCompName(compInAssembly, ref compName);
                PPDetails_TaskPaneHost.getCompName(ref compName);
            }

            //get the part document of the raw material
            ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
            PartDoc swPartDoc = (PartDoc)compModDoc;

            SelectionMgr SwSelMgr = (SelectionMgr)compModDoc.SelectionManager;
            SelectData SwSelData = (SelectData)SwSelMgr.CreateSelectData();

            SelectStatus = compName[0].Select2(true, 0);

            swAssembly.EditPart();

            Feature ThisFeature = compModDoc.FirstFeature();
            

            while (ThisFeature != null)
            {
                if (ThisFeature.GetTypeName2() == "SurfaceBodyFolder")
                {
                    SelectedBody = null;
                    BodyFolder ThisBodyFolder = ThisFeature.GetSpecificFeature2();
                    ObjBodies = (object[])ThisBodyFolder.GetBodies();
                    //foreach (Body2 ThisBody in ObjBodies)
                    //{
                    foreach (AddedReferenceSurface ThisSurface in InitialRefSurfaces)
                    {
                        //if (InitialRefSurfaces[0].SurfaceName.Contains(ThisBody.Name))
                        //{

                        //SelectStatus = ThisBody.Select2(true, SwSelData);
                        //int DisplayStatus = ThisBody.Display3(compModDoc, 150, (int)swTempBodySelectOptions_e.swTempBodySelectable);
                        //SelectStatus = swDoc.Extension.SelectByID2(InitialRefSurfaces[0].SurfaceName, "SURFACEBODY", 0, 0, 0, true, 0, null, 0);
                        //SelectStatus = ThisSurface.SurfaceFeature.Select2(true, 0);
                        SelectStatus = swDoc.Extension.SelectByID2(ThisSurface.SurfaceName, "SURFACEBODY", 0, 0, 0, true, 0, null, 0);
                        //SelectedBody = ThisBody;
                        //break;
                        // }
                        //}
                    }

                    break;
                }


                ThisFeature = ThisFeature.GetNextFeature();
            }

            //SelectStatus = SelectedBody.Select2(true, SwSelData);

            //SelectStatus = swDoc.Extension.SelectByID2(SelectedBody.Name, "SURFACEBODY", 0, 0, 0, true, 0, null, 0);

        }

        //tools for calculating setup position
        #region SetupCalculator

        #endregion

        //marking the openface
        public void SetOpenFace()
        {
            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SelectedObject = SwSelMgr.GetSelectedObject5(1);
            Entity SwEntity = (Entity)SelectedObject;

            //create the definition
            AttributeDef SwAttDef = SwApp.DefineAttribute("open_face");
            Boolean RetVal = SwAttDef.AddParameter("status", (int)swParamType_e.swParamTypeDouble, 1, 0);
            
            RetVal = SwAttDef.Register();

            SolidWorks.Interop.sldworks.Attribute SwAttribute = default(SolidWorks.Interop.sldworks.Attribute);
            OpenFaceCounter++;

            SwAttribute = SwAttDef.CreateInstance5(DocumentModel, SwEntity, "open_face" + OpenFaceCounter.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);

            if (SwAttribute != null)
            {
                iSwApp.SendMsgToUser("face got name");
            }
            
            //iSwApp.SendMsgToUser("got from SetOpenFace");
        }

        //reading the open face
        public void ReadOpenFace()
        {
            ModelDoc2 DocumentModel = SwApp.ActiveDoc;
            SelectionMgr SwSelMgr = DocumentModel.SelectionManager;
            Object SelectedObject = SwSelMgr.GetSelectedObject5(1);
            Entity SwEntity = (Entity)SelectedObject;

            //Create the definition
            //AttributeDef SwAttDef = SwApp.DefineAttribute("open_face");
            //Boolean RetVal = SwAttDef.AddParameter("status", (int)swParamType_e.swParamTypeDouble, 1, 0);

            AttributeDef SwAttDef = SwApp.DefineAttribute("open_face");
            //Boolean RetVal = SwAttDef.AddParameter("status", (int)swParamType_e.swParamTypeDouble, 1, 0);

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
                iSwApp.SendMsgToUser("attribute was not found");
            }
            else
            {
                //Parameter OpenParam = SwAttribute.GetParameter("status");
                Parameter OpenParam = SwAttribute.GetParameter("status");

                String Value = OpenParam.GetDoubleValue().ToString();

                if (Value.Equals("1"))
                {
                    iSwApp.SendMsgToUser("this is a open face");
                }
                else
                {
                    iSwApp.SendMsgToUser("this is not a open face");
                }

            }

        }

        //cost analysis
        public void CostAnalysis()
        {            
            ModelDoc2 SwDoc = null;

            SwDoc = (ModelDoc2)SwApp.ActiveDoc;
            ModelDocExtension SwDocExt = (ModelDocExtension)SwDoc.Extension;

            // Get Costing templates names
            Debug.Print("Costing template folders:");
            String machiningCostingTemplatePathName = SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsCostingTemplates);
            Debug.Print("    Name of Costing template folder: " + machiningCostingTemplatePathName);
            String machiningCostingReportTemplateFolderName = SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swFileLocationsCostingReportTemplateFolder);
            Debug.Print("    Name of Costing report template folder: " + machiningCostingReportTemplateFolderName);
            Debug.Print("");
                        
            CostManager swCosting = (CostManager)SwDocExt.GetCostingManager();
            swCosting.WaitForUIUpdate();

            // Get the number of templates
            int nbrMachiningCostingTemplate = swCosting.GetTemplateCount((int)swcCostingType_e.swcCostingType_Machining);
            int nbrCommonCostingTemplate = swCosting.GetTemplateCount((int)swcCostingType_e.swcCostingType_Common);

            // Get names of templates
            Object[] machiningCostingTemplates = (Object[])swCosting.GetTemplatePathnames((int)swcCostingType_e.swcCostingType_Machining);
            Object[] commonCostingTemplates = (Object[])swCosting.GetTemplatePathnames((int)swcCostingType_e.swcCostingType_Common);

            Array.Resize(ref machiningCostingTemplates, nbrMachiningCostingTemplate + 1);
            Array.Resize(ref commonCostingTemplates, nbrCommonCostingTemplate + 1);

            Debug.Print("Costing templates:");

            // Print names of templates to Immediate window
            for (int i = 0; i <= (nbrMachiningCostingTemplate - 1); i++)
            {
                Debug.Print("    Name of machining Costing template: " + machiningCostingTemplates[i]);
            }

            Debug.Print("");

            for (int i = 0; i <= (nbrCommonCostingTemplate - 1); i++)
            {
                Debug.Print("    Name of common Costing template: " + commonCostingTemplates[i]);
            }

            Debug.Print("");

            // Get Costing part
            Object swCostingModel = (Object)swCosting.CostingModel;
            CostPart swCostingPart = (CostPart)swCostingModel;

            // Create common Costing analysis
            CostAnalysis swCostingAnalysis = (CostAnalysis)swCostingPart.CreateCostAnalysis("c:\\program files\\solidworks corp\\solidworks\\lang\\english\\costing templates\\multibodytemplate_default(englishstandard).sldctc");

            // Get common Costing analysis data
            Debug.Print("Common Costing analysis data:");
            Debug.Print("    Template name:  " + swCostingAnalysis.CostingTemplateName);
            Debug.Print("    Currency code: " + swCostingAnalysis.CurrencyCode);
            Debug.Print("    Currency name: " + swCostingAnalysis.CurrencyName);
            Debug.Print("    Currency separator: " + swCostingAnalysis.CurrencySeparator);
            Debug.Print("    Total manufacturing cost: " + swCostingAnalysis.GetManufacturingCost());
            Debug.Print("    Material costs: " + swCostingAnalysis.GetMaterialCost());
            Debug.Print("    Setup cost: " + swCostingAnalysis.GetSetupCost());
            Debug.Print("    Total cost to charge: " + swCostingAnalysis.GetTotalCostToCharge());
            Debug.Print("    Total cost to manufacture: " + swCostingAnalysis.GetTotalCostToManufacture());
            Debug.Print("    Lot size: " + swCostingAnalysis.LotSize);
            Debug.Print("    Total quantity: " + swCostingAnalysis.TotalQuantity);
            Debug.Print("");

            Boolean isBody = false;

            // Get Costing bodies
            int nbrCostingBodies = swCostingPart.GetBodyCount();
            if ((nbrCostingBodies > 0))
            {
                Debug.Print("Costing bodies:");
                Debug.Print("    Number of Costing bodies: " + nbrCostingBodies);
                Object[] costingBodies = (Object[])swCostingPart.GetBodies();
                for (int i = 0; i <= (nbrCostingBodies - 1); i++)
                {
                    CostBody swCostingBody = (CostBody)costingBodies[i];
                    String costingBodyName = swCostingBody.GetName();
                    Debug.Print("    Name of Costing body: " + costingBodyName);
                    // Make sure body is machining body
                    if ((swCostingBody.GetBodyType() == (int)swcBodyType_e.swcBodyType_Machined))
                    {
                        isBody = true;
                        // Determine analysis status of Costing body
                        switch ((int)swCostingBody.BodyStatus)
                        {
                            case (int)swcBodyStatus_e.swcBodyStatus_NotAnalysed:
                                // Create Costing analysis
                                swCostingAnalysis = (CostAnalysis)swCostingBody.CreateCostAnalysis("c:\\program files\\solidworks corp\\solidworks\\lang\\english\\costing templates\\machiningtemplate_default(englishstandard).sldcts");
                                Debug.Print("    Creating machining Costing analysis for: " + swCostingBody.GetName());
                                break;
                            case (int)swcBodyStatus_e.swcBodyStatus_Analysed:
                                // Get Costing analysis
                                swCostingAnalysis = (CostAnalysis)swCostingBody.GetCostAnalysis();
                                Debug.Print("    Getting machining Costing analysis for: " + swCostingBody.GetName());
                                break;
                            case (int)swcBodyStatus_e.swcBodyStatus_Excluded:
                                // Body excluded from Costing analysis
                                Debug.Print("    Excluded from machining Costing analysis: " + swCostingBody.GetName());
                                break;
                            case (int)swcBodyStatus_e.swcBodyStatus_AssignedCustomCost:
                                // Body has an assigned custom Cost
                                Debug.Print("    Custom cost assigned: " + swCostingBody.GetName());
                                break;
                        }

                        Debug.Print("");

                    }
                }
            }

            if (!isBody)
            {
                Debug.Print("");
                Debug.Print("No bodies in part! Exiting macro.");
                return;
            }

            // Get machining Costing Analysis data
            CostAnalysisMachining swCostingMachining = (CostAnalysisMachining)swCostingAnalysis.GetSpecificAnalysis();
            Debug.Print("Machining Costing analysis: ");
            Debug.Print("    Current material: " + swCostingMachining.CurrentMaterial);
            Debug.Print("    Current material class: " + swCostingMachining.CurrentMaterialClass);
            Debug.Print("    Current plate thickness: " + swCostingMachining.CurrentPlateThickness);

            Debug.Print("");

            // Get Costing features
            Debug.Print("Costing features:");
            CostFeature swCostingFeat = (CostFeature)swCostingAnalysis.GetFirstCostFeature();
            while ((swCostingFeat != null))
            {
                Debug.Print("    Feature: " + swCostingFeat.Name);
                Debug.Print("      Type: " + swCostingFeat.GetType());
                Debug.Print("        Setup related: " + swCostingFeat.IsSetup);
                Debug.Print("        Overridden: " + swCostingFeat.IsOverridden);
                Debug.Print("        Combined cost: " + swCostingFeat.CombinedCost);
                Debug.Print("        Combined time: " + swCostingFeat.CombinedTime);

                CostFeature swCostingSubFeat = swCostingFeat.GetFirstSubFeature();
                while ((swCostingSubFeat != null))
                {
                    Debug.Print("      Subfeature: " + swCostingSubFeat.Name);
                    Debug.Print("        Type: " + swCostingSubFeat.GetType());
                    Debug.Print("          Setup related: " + swCostingSubFeat.IsSetup);
                    Debug.Print("          Overridden: " + swCostingSubFeat.IsOverridden);
                    Debug.Print("          Combined cost: " + swCostingSubFeat.CombinedCost);
                    Debug.Print("          Combined time: " + swCostingSubFeat.CombinedTime);
                    CostFeature swCostingNextSubFeat = (CostFeature)swCostingSubFeat.GetNextFeature();
                    swCostingSubFeat = null;
                    swCostingSubFeat = (CostFeature)swCostingNextSubFeat;
                    swCostingNextSubFeat = null;
                }
                CostFeature swCostingNextFeat = swCostingFeat.GetNextFeature();
                swCostingFeat = null;
                swCostingFeat = (CostFeature)swCostingNextFeat;
                swCostingNextFeat = null;
            }




            if (swCosting == null) { iSwApp.SendMsgToUser("CostManager is NULL"); }
            else { iSwApp.SendMsgToUser("CostManager is not NULL"); }
        }

        public static double Getsize(double[] NVector)
        {
            double size = Math.Sqrt(NVector[0] * NVector[0] + NVector[1] * NVector[1] + NVector[2] * NVector[2]);
            return size;

        }
        
        //平面-平面の稜線の確認
        public static bool CheckTopologicalData(ModelDoc2 Doc, Body2 body2Check)
        {
            Face2[] AttachFace = new Face2[2];　
            double[] NormalA = new double[3];
            double[] NormalB = new double[3];
            double[] EdgeVector = new double[3];
            double[] StartPoint = new double[3];
            double[] MidPoint = new double[3];
            double[] EndPoint = new double[3];
            Vertex SVertex;
            Vertex SVertex2;
            Vertex EVertex;
            Vertex EVertex2;
            Vertex Vertex;
            int i, j;
            double[,] Point = new double[4, 3];
            double[] GPoint = new double[5];
            bool checkA;
            long EdgeCount;
            Edge[] EdgeMatrixA = new Edge[100];
            Edge[] EdgeMatrixB = new Edge[100];
            double[] Closest = new double[5];
            double[] RVector = new double[3];
            double[] Vector = new double[3];
            int[] check = new int[3];
            int[] check1 = new int[3];
            int[] check2 = new int[3];
            bool Face = true;
            double Angle;



            //var res;

            //ModelDoc2 swDoc = iSwApp.IActiveDoc2;

            
            //Body2 Body = swPartDoc.Body();

            int EdgeNum = body2Check.GetEdgeCount();
            Edge[] AllEdge = new Edge[EdgeNum];
            int[] CheckResult = new int[EdgeNum];
            body2Check.GetEdges().CopyTo(AllEdge, 0);
            Edge tmpEdge = AllEdge[0];

            for (int k = 0; k < EdgeNum; k++)
            {
                tmpEdge = AllEdge[k];
                tmpEdge.GetTwoAdjacentFaces2().CopyTo(AttachFace,0);

                EdgeCount = AttachFace[0].GetEdgeCount();
                AttachFace[0].GetEdges().CopyTo(EdgeMatrixA,0);
                NormalA = AttachFace[0].Normal;
                double size = Getsize(NormalA);
                for (int l = 0; l < 3; l++)
                {
                    NormalA[l] = NormalA[l] / size;

                }

                EdgeCount = AttachFace[1].GetEdgeCount();
                AttachFace[1].GetEdges().CopyTo(EdgeMatrixB,0);
                AttachFace[1].Normal.CopyTo(NormalB,0);
                size = Getsize(NormalB);
                for (int l = 0; l < 3; l++)
                {
                    NormalB[l] = NormalB[l] / size;

                }

                SVertex = tmpEdge.GetStartVertex();
                EVertex = tmpEdge.GetEndVertex();
                SVertex.GetPoint().CopyTo(StartPoint,0);
                EVertex.GetPoint().CopyTo(EndPoint,0);
                for (int l = 0; l < 3; l++)
                {
                    MidPoint[l] = (StartPoint[l] + EndPoint[l]) / 2;
                }

                j = 0;

                i = 0;
                do
                {
                    if (EdgeMatrixA[i] == tmpEdge)
                    {
                        i++;
                        continue;
                    }
                    SVertex2 = EdgeMatrixA[i].GetStartVertex();
                    EVertex2 = EdgeMatrixA[i].GetEndVertex();
                    if (EVertex2 != EVertex && SVertex2 != EVertex && EVertex2 != SVertex && SVertex2 != SVertex)
                    {
                        i++;
                        continue;
                    }

                    if (SVertex2 == SVertex || SVertex2 == EVertex)
                    {
                        Vertex = EVertex2;
                    }
                    else
                    {
                        Vertex = SVertex2;
                    }
                    //Point[j, 0] = Vertex.GetPoint();
                    double[] tmpArray = new double[3];
                    tmpArray = Vertex.GetPoint().Clone();
                    Point[j, 0] = tmpArray[0];
                    Point[j, 1] = tmpArray[1];
                    Point[j, 2] = tmpArray[2];
                    j++;
                    i++;


                } while (j < 2);


                i = 0;
                do
                {
                    if (EdgeMatrixB[i] == tmpEdge)
                    {
                        i++;
                        continue;
                    }
                    SVertex2 = EdgeMatrixB[i].GetStartVertex();
                    EVertex2 = EdgeMatrixB[i].GetEndVertex();
                    if (EVertex2 != EVertex && SVertex2 != EVertex && EVertex2 != SVertex && SVertex2 != SVertex)
                    {
                        i++;
                        continue;
                    }

                    if (SVertex2 == SVertex || SVertex2 == EVertex)
                    {
                        Vertex = EVertex2;
                    }
                    else
                    {
                        Vertex = SVertex2;
                    }
                    //Point[j, 0] = Vertex.GetPoint();
                    double[] tmpArray = new double[3];
                    tmpArray = Vertex.GetPoint().Clone();
                    Point[j, 0] = tmpArray[0];
                    Point[j, 1] = tmpArray[1];
                    Point[j, 2] = tmpArray[2];
                    j++;
                    i++;


                } while (j < 4);

                do
                {
                    j = j % 4;

                    GPoint[0] = (EndPoint[0] + StartPoint[0] + Point[j, 0]) / 3;
                    GPoint[1] = (EndPoint[1] + StartPoint[1] + Point[j, 1]) / 3;
                    GPoint[2] = (EndPoint[2] + StartPoint[2] + Point[j, 2]) / 3;

                    if (j == 0 || j == 1) Closest = AttachFace[0].GetClosestPointOn(GPoint[0], GPoint[1], GPoint[2]);
                    if (j == 2 || j == 3) Closest = AttachFace[1].GetClosestPointOn(GPoint[0], GPoint[1], GPoint[2]);
                    check[0] = (int)(Closest[0] * 10000 - GPoint[0] * 10000);
                    check[1] = (int)(Closest[1] * 10000 - GPoint[1] * 10000);
                    check[2] = (int)(Closest[2] * 10000 - GPoint[2] * 10000);
                    if (check[0] == 0 && check[1] == 0 && check[2] == 0)
                    {
                        checkA = false;
                        if (j == 0 || j == 1)
                        {
                            Face = true;
                        }
                        else
                        {
                            Face = false;
                        };
                        RVector[0] = MidPoint[0] - GPoint[0];
                        RVector[1] = MidPoint[1] - GPoint[1];
                        RVector[2] = MidPoint[2] - GPoint[2];
                        break;
                    }
                    else
                    {
                        checkA = true;
                        Point[j,0] = GPoint[0];
                        Point[j,1] = GPoint[1];
                        Point[j,2] = GPoint[2];
                    };
                    j++;
                } while (checkA);
                size = Getsize(RVector);
                for (int l = 0; l < 3; l++)
                {
                    RVector[l] = RVector[l] / size;

                }
                if (Face)
                    Angle = RVector[0] * NormalB[0] + RVector[1] * NormalB[1] + RVector[2] * NormalB[2];
                else
                    Angle = RVector[0] * NormalA[0] + RVector[1] * NormalA[1] + RVector[2] * NormalA[2];

                //結果を返す（山折り：1　谷折り：-1　その他：2）
                if (Angle > 0.0)
                {
                    CheckResult[k] = 1;	//山折り
                }
                else if (Angle < 0.0)
                {
                    CheckResult[k] = -1;	//谷折り
                }
                else
                {
                    CheckResult[k] = 2;
                }
            }
            for (int K = 0; K < EdgeNum; K++)
            {
                if (CheckResult[K] == -1)
                {
                    //iSwApp.SendMsgToUser2("この形状は凹形状です", 2, 4);
                    break;
                }
                if (K == EdgeNum - 1 && CheckResult[K] == 1)
                {
                    //iSwApp.SendMsgToUser2("この形状は凸形状です", 2, 4);
                    return true;
                    
                }
            }
            return false;
        }

        public int OpenFaceCounter;

        public void ShowPMP()
        {
            if (ppage != null)
                ppage.Show();
        }

        public int EnablePMP()
        {
            if (iSwApp.ActiveDoc != null)
                return 1;
            else
                return 0;
        }

        public void FlyoutCallback()
        {
            FlyoutGroup flyGroup = iCmdMgr.GetFlyoutGroup(flyoutGroupID);
            flyGroup.RemoveAllCommandItems();

            flyGroup.AddCommandItem(System.DateTime.Now.ToLongTimeString(), "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

        }

        public int FlyoutEnable()
        {
            return 1;
        }

        public void FlyoutCommandItem1()
        {
            iSwApp.SendMsgToUser("Flyout command 1");
        }

        public int FlyoutEnableCommandItem1()
        {
            return 1;
        }
        #endregion

        #region Event Methods
        public bool AttachEventHandlers()
        {
            AttachSwEvents();
            //Listen for events on all currently open docs
            AttachEventsToAllDocuments();
            return true;
        }

        private bool AttachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify += new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 += new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 += new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify += new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify += new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }
        }        

        private bool DetachSwEvents()
        {
            try
            {
                SwEventPtr.ActiveDocChangeNotify -= new DSldWorksEvents_ActiveDocChangeNotifyEventHandler(OnDocChange);
                SwEventPtr.DocumentLoadNotify2 -= new DSldWorksEvents_DocumentLoadNotify2EventHandler(OnDocLoad);
                SwEventPtr.FileNewNotify2 -= new DSldWorksEvents_FileNewNotify2EventHandler(OnFileNew);
                SwEventPtr.ActiveModelDocChangeNotify -= new DSldWorksEvents_ActiveModelDocChangeNotifyEventHandler(OnModelChange);
                SwEventPtr.FileOpenPostNotify -= new DSldWorksEvents_FileOpenPostNotifyEventHandler(FileOpenPostNotify);
                return true;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return false;
            }

        }

        public void AttachEventsToAllDocuments()
        {
            ModelDoc2 modDoc = (ModelDoc2)iSwApp.GetFirstDocument();
            while (modDoc != null)
            {
                if (!openDocs.Contains(modDoc))
                {
                    AttachModelDocEventHandler(modDoc);
                }
                else if (openDocs.Contains(modDoc))
                {
                    bool connected = false;
                    DocumentEventHandler docHandler = (DocumentEventHandler)openDocs[modDoc];
                    if (docHandler != null)
                    {
                        connected = docHandler.ConnectModelViews();
                    }
                }

                modDoc = (ModelDoc2)modDoc.GetNext();
            }
        }

        public bool AttachModelDocEventHandler(ModelDoc2 modDoc)
        {
            if (modDoc == null)
                return false;

            DocumentEventHandler docHandler = null;

            if (!openDocs.Contains(modDoc))
            {
                switch (modDoc.GetType())
                {
                    case (int)swDocumentTypes_e.swDocPART:
                        {
                            docHandler = new PartEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocASSEMBLY:
                        {
                            docHandler = new AssemblyEventHandler(modDoc, this);
                            break;
                        }
                    case (int)swDocumentTypes_e.swDocDRAWING:
                        {
                            docHandler = new DrawingEventHandler(modDoc, this);
                            break;
                        }
                    default:
                        {
                            return false; //Unsupported document type
                        }
                }
                docHandler.AttachEventHandlers();
                openDocs.Add(modDoc, docHandler);
            }
            return true;
        }

        public bool DetachModelEventHandler(ModelDoc2 modDoc)
        {
            DocumentEventHandler docHandler;
            docHandler = (DocumentEventHandler)openDocs[modDoc];
            openDocs.Remove(modDoc);
            modDoc = null;
            docHandler = null;
            return true;
        }

        public bool DetachEventHandlers()
        {
            DetachSwEvents();

            //Close events on all currently open docs
            DocumentEventHandler docHandler;
            int numKeys = openDocs.Count;
            object[] keys = new Object[numKeys];

            //Remove all document event handlers
            openDocs.Keys.CopyTo(keys, 0);
            foreach (ModelDoc2 key in keys)
            {
                docHandler = (DocumentEventHandler)openDocs[key];
                docHandler.DetachEventHandlers(); //This also removes the pair from the hash
                docHandler = null;
            }
            return true;
        }
        //選択したフィーチャの名前とタイプを調べる
        public void CheckFeature()
        {
            ModelDoc2 swModDoc = iSwApp.ActiveDoc;
            SelectionMgr swSelMgr = swModDoc.SelectionManager;
            Feature swFeat = swSelMgr.GetSelectedObject6(1, -1);

            if (swFeat == null)
            {
                iSwApp.SendMsgToUser2("Feature Selection Failed", (int)swMessageBoxBtn_e.swMbOk, (int)swMessageBoxBtn_e.swMbOk);
            }
            else
            {
                if (swFeat.GetTypeName2() == "ProfileFeature" || swFeat.GetTypeName2() == "3DProfileFeature")
                {
                    Sketch swSketch = (Sketch)swFeat.GetSpecificFeature2();
                    iSwApp.SendMsgToUser2("Name: " + swFeat.Name + System.Environment.NewLine +
                                          "Type: " + swFeat.GetTypeName2() + System.Environment.NewLine +
                                          "SketchRegCount: " + swSketch.GetSketchRegionCount().ToString(), (int)swMessageBoxBtn_e.swMbOk, (int)swMessageBoxBtn_e.swMbOk);
                }
                else
                {
                    iSwApp.SendMsgToUser2("Name: " + swFeat.Name + System.Environment.NewLine +
                                          "Type: " + swFeat.GetTypeName2(), (int)swMessageBoxBtn_e.swMbOk, (int)swMessageBoxBtn_e.swMbOk);
                }
            }
        }
        #endregion

        #region Event Handlers
        //Events
        public int OnDocChange()
        {
            return 0;
        }

        public int OnDocLoad(string docTitle, string docPath)
        {
            return 0;
        }

        int FileOpenPostNotify(string FileName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnFileNew(object newDoc, int docType, string templateName)
        {
            AttachEventsToAllDocuments();
            return 0;
        }

        public int OnModelChange()
        {
            return 0;
        }

        #endregion

        #region base

        #endregion

    }

    
    // class for saving properties for each generated reference plane
    public class _planeProperties
    {
        public string name { get; set; } //name of the feature plane

        public int rankByDistance { get; set; } //rank number of plane

        public double distance {get; set;} //distance from centroid

        public Feature featureObj { get; set; } //associated feature pointer

        public object planeNormal { get; set; } //plane normal

        public int mark { get; set; } //marking properties

        public List<Body2> bodyList { get; set; } //associated body which is generated by the reference plane

        public double[] centroid { get; set; } //saving the centroid
    }

    //class for parent and child relations
    public class RemovedBody
    {
        public AddedReferencePlane ParentPlane { get; set; }
        public List<Feature> Removal { get; set; }
        public RemovalBody TRV { get; set; }
    }

    //class for removal body
    public class RemovalBody
    {
        public Body2 BodyObj { get; set; } //keep pointer to body object

        public double[] BodyCentroid { get; set; } //keep the body centroid information

        public List<TAD> ListOfTAD { get; set; } //keep all the candidate of TAD

        public double Volume { get; set; } //keep the volume of the body

        public string Name { get; set; } //keep the split name

        public bool IsConvex { get; set; } //keep the status of convexity

        public bool Location { get; set; } //keep the location status

        public List<TAD> ClosedTAD { get; set; } //keep the unaccessible TAD (closed face)

        public TAD RecommendedTAD { get; set; } //keep the recommended TAD based on the LWD analysis

        public double BestTADArea { get; set; } //keep the area of recommended TAD

        public double[] MaxMin { get; set; } //keep the maximum and minimum point of body from tesselation process

        public double[] MidPoint { get; set; } //keep the middle point of the bodybox

        public double[] VirtualPoint { get; set; } //keep the lowest point based on the direction of the recommended TAD
    }

    //class for plane and removal volume body relation
    public class PlaneWithBody
    {
        public string name { get; set; } //keep the name
        
        public AddedReferencePlane SelectedReferencePlane { get; set; } //keep pointer to selected AddedReferencePlane

        public RemovalBody TMpRemovalBody { get; set; } //keep all bodies that related to AddedReferencePlane

        public TAD SelectedTAD { get; set; } //keep the selected TAD
        
    }

    //class for keeping the datums
    public class PPX_Datum
    {
        public string Name { get; set; } //the name that keeping the datum
        public int Index { get; set; } //the index
        public string Type { get; set; } // type of the datum (plane or profile)
        public string BodyOwner { get; set; } //keep the owner of the datum
    }

    //class for saving addedReferencePlane
    public class AddedReferencePlane
    {
        public string name { get; set; } //keep the name

        public RefPlane ReferencePlane { get; set; } //keep the pointer to the original plane

        public AddedReferenceSurface ReferenceSurface { get; set;} //keep the pointer to the reference surface

        public Face2 AttachedFace { get; set; } //keep the pointer to the corresponding face

        public Feature CorrespondFeature { get; set; } //keep the pointer to the corresponding feature

        public int RankByDistance { get; set; } //keep the rank by the distance

        public Double DistanceFromCentroid { get; set; } //keep the distance from centroid

        public List<string> PlaneIntersection { get; set; } //keep the intersection with other planes

        public int Score { get; set; } //keep the score of intersection

        public object PlaneNormal { get; set; } //keep the plane normal

        public string NormalOrientation { get; set; } //keep the plane normal orientation

        public Boolean isOnBB { get; set; } //keep the information to indicate whether if the reference plane is located on the side of the bounding box or not

        public Boolean isOuter { get; set; } //keep the information to indicate whether if the reference plane is located on the most outer surface or not

        public Boolean CPost { get; set; } //keep the centroid position refer to the normal direction

        public Boolean SafeDirection { get; set; } //keep the direction position, true: safe, false: unsafe

        public int MarkingOpt { get; set; } //keep the marking options that is used for iterating the plane

        public int Remark { get; set; } //keep additional remark for the plane

        public Boolean Possibility { get; set; } //keep the status for possibility to add as a plane candidate

        public Object MaxMinValue { get; set; } //keep the MaxMin value that has been define by tesselating corresponding face

        public Double Location(int Axis) //keep the location of the plane based on its axis requirement, 0:x, 1:y, 2:z
        {
            Double[] ThisBoundary = (Double[])MaxMinValue;

            if (Axis == 0)
            {
                return ThisBoundary[0];
            }

            if (Axis == 1)
            {
                return ThisBoundary[1];
            }

            if (Axis == 2)
            {
                return ThisBoundary[2];
            }

            return 0.0;
        }

        public Double GetCornerValue(int Index) //get the maxmin value based on its requirement
        {
            Double[] ThisMaxMin = (Double[]) MaxMinValue;
            
            switch(Index)
            {
                case 0: return ThisMaxMin[0];
                case 1: return ThisMaxMin[1];
                case 2: return ThisMaxMin[2];
                case 3: return ThisMaxMin[3];
                case 4: return ThisMaxMin[4];
                case 5: return ThisMaxMin[5];
            }
                
            return 0.0;
        }

        public Boolean IsPlanar { get; set; } //keep the type of the reference plane (true: planar, false: non planar)

        public PatternConfig Patterns { get; set; } //keep the pattern (line, circle, arc) configuration

        public string BodyOwner { get; set; } //keep the name of the owner

        public RemovalBody AttachedBody { get; set; } //keep the body pointer

        public PlaneTolerance Tolerance { get; set; } // keep the plane tolerance

        public string GetToleranceValue(int ThisTol)
        {
            switch (ThisTol)
            { 
                case 1:
                    return Tolerance.datum;
                    
                case 2:
                    return Tolerance.flatness;
                    
                case 3:
                    return Tolerance.parallelism;
                    
                case 4:
                    return Tolerance.angularity;
                    
                case 5:
                    return Tolerance.perpendicularity;
                    
                case 6:
                    return Tolerance.location;
                    
            }

            return "";
        } //function to get tolerance value

        public List<AddedReferenceProfile> Profiles { get; set; } //keep the profile on this plane

    }

    //Class for saving AddedReferenceProfile
    public class AddedReferenceProfile
    {
        public string name { get; set; } //keep the name

        public RefPlane ReferencePlane { get; set; } //keep the pointer to the original plane

        public AddedReferenceSurface ReferenceSurface { get; set;} //keep the pointer to the reference surface

        public Face2 AttachedFace { get; set; } //keep the pointer to the corresponding face

        public Feature CorrespondFeature { get; set; } //keep the pointer to the corresponding feature

        public Edge EdgeProfile { get; set; } // keep the pointer to the corresponding edge profile

        public string CurveType { get; set; } // keep the type of curve

        public object CurveParam { get; set; } // keep the pointer to the corresponding curve profile data

        public Boolean Main { get; set; } // keep the status whether this entity profile is the main profile, in initial profile there are two similar profiles.

        public int RankByDistance { get; set; } //keep the rank by the distance

        public Double DistanceFromCentroid { get; set; } //keep the distance from centroid

        public List<string> PlaneIntersection { get; set; } //keep the intersection with other planes

        public int Score { get; set; } //keep the score of intersection

        public object PlaneNormal { get; set; } //keep the plane normal

        public string NormalOrientation { get; set; } //keep the plane normal orientation

        public Boolean isOnBB { get; set; } //keep the information to indicate whether if the reference plane is located on the side of the bounding box or not

        public Boolean isOuter { get; set; } //keep the information to indicate whether if the reference plane is located on the most outer surface or not

        public Boolean CPost { get; set; } //keep the centroid position refer to the normal direction
        
        public int MarkingOpt { get; set; } //keep the marking options that is used for iterating the plane

        public int Remark { get; set; } //keep additional remark for the plane

        public Boolean Possibility { get; set; } //keep the status for possibility to add as a plane candidate

        public Object MaxMinValue { get; set; } //keep the MaxMin value that has been define by tesselating corresponding face

        public Boolean IsPlanar { get; set; } //keep the type of the reference plane (true: planar, false: non planar)

        public PatternConfig Patterns { get; set; } //keep the pattern (line, circle, arc) configuration

        //追加部分
        public string BodyOwner { get; set; } //keep the name of the owner
        public AddedReferencePlane clone() 
        {
            AddedReferencePlane cloneplane = new AddedReferencePlane();
            cloneplane.name = this.name;
            cloneplane.ReferencePlane = this.ReferencePlane;
            cloneplane.AttachedFace = this.AttachedFace;
            cloneplane.CorrespondFeature = this.CorrespondFeature;
            cloneplane.RankByDistance = this.RankByDistance;
            cloneplane.DistanceFromCentroid = this.DistanceFromCentroid;
            cloneplane.PlaneIntersection = this.PlaneIntersection;
            cloneplane.Score = this.Score;
            cloneplane.PlaneNormal = this.PlaneNormal;
            cloneplane.NormalOrientation = this.NormalOrientation;
            cloneplane.isOnBB = this.isOnBB;
            cloneplane.isOuter = this.isOuter;
            cloneplane.CPost = this.CPost;
            cloneplane.MarkingOpt = this.MarkingOpt;
            cloneplane.Remark = this.Remark;
            cloneplane.Possibility = this.Possibility;
            cloneplane.MaxMinValue = this.MaxMinValue;
            cloneplane.IsPlanar = this.IsPlanar;
            cloneplane.Patterns = this.Patterns;

            return cloneplane;


        }



        public RemovalBody AttachedBody { get; set; } //keep the body pointer

        public PlaneTolerance Tolerance { get; set; } // keep the plane tolerance

        public string GetToleranceValue(int ThisTol)
        {
            switch (ThisTol)
            {
                case 1:
                    return Tolerance.datum;

                case 2:
                    return Tolerance.flatness;

                case 3:
                    return Tolerance.parallelism;

                case 4:
                    return Tolerance.angularity;

                case 5:
                    return Tolerance.perpendicularity;

                case 6:
                    return Tolerance.location;

            }

            return "";
        } //function to get tolerance value
    }

    public class AddedReferenceSurface
    {
        public string SurfaceName { get; set; } //keep the feature name of Surface  in the feature manager 

        public Surface SurfacePlane { get; set; } //keep the pointer to the surface that is created only for freeform surface (include cylinder)

        public Feature SurfaceFeature { get; set; } // keep the feature of the surface

        public string BodyOwner { get; set; } //keep the name of the owner

        public RemovalBody AttachedBody { get; set; } //keep the body pointer
    }

    //class for saving the machining plan
    public class MachiningProcess
    { 
        public AddedReferencePlane MachiningReference {get;set;} //keep the reference plane for machining

        public string cuttingTool { get; set; } //keep the cutting tool type information

        public string toolPath { get; set; } //keep the type of suggested tool path

        public TAD SelectedTAD { get; set; } //keep the tool access direction for machining

        public RemovalBody TRV { get; set; } //keep the volume of the removal body

        public double[] VisibilityCone { get; set; } //keep the visibility cone information

        public Double MachiningTime { get; set; } //keep the machining time for TRV

        public Double MachiningVolume { get; set; } //keep the machining volume of TRV

    }
    
    //class for TAD
    public class TAD
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }

        public bool IsEmpty()
        {
            if ((isEqual(X, 0.0) == true) &&
                (isEqual(Y, 0.0) == true) &&
                (isEqual(Z, 0.0) == true))
            {
                return true;
            }

            return false;
        }

        public bool IsSimilar(TAD ThisTAD)
        {
            if ((isEqual(X, ThisTAD.X) == true) &&
                (isEqual(Y, ThisTAD.Y) == true) &&
                (isEqual(Z, ThisTAD.Z) == true))
            {
                return true;
            }

            return false;
        }

        public bool IsNear(TAD ThisTAD, MathUtility ThisUtils)
        {   
            MathVector VectorA, VectorB;
            double DistValue;

            VectorA = ThisUtils.CreateVector(new double[3] { X, Y, Z });
            VectorB = ThisUtils.CreateVector(new double[3] { ThisTAD.X, ThisTAD.Y, ThisTAD.Z });

            DistValue = VectorA.Dot(VectorB) / (VectorA.GetLength() * VectorB.GetLength());

            if (DistValue > Math.Cos(45))
            {
                return true;
            }

            return false;

        }

        public void Scale(double ThisValue)
        {
            this.X = this.X * ThisValue;
            this.Y = this.Y * ThisValue;
            this.Z = this.Z * ThisValue;

            return;
        }

        private bool isEqual(double value1, double value2)
        {
            return (Math.Abs(Math.Round((value1 - value2), 4)) <= 0.001);

        }
    }

    //class for candidate of machining plan
    public class MachiningPlan
    {
        public string ViewName { get;set; } //keep the view name on solidwork window

        public string FullPath { get; set; } //keep the path where is being saved

        public List<MachiningProcess> MachiningProceses { get; set; } //keep the pointer of selected machining process

        public double MachiningCost { get; set; } //keep the machining cost

        public double MachiningTime { get; set; } //keep the machining time

        public double MachiningVolume { get; set; } //keep the machining volume

        public int NumberOfTADchanges { get; set; } //keep the number of TAD changes

        public int NumberOfTool { get; set; } //keep the number of needed tools

        public int NumberOfSetups { get; set; } //keep the number of setups

        public string SetupNormal { get; set; } //keep the normal of setup

    }

    //class for keeping the circle pattern
    public class CircularPattern
    {
        public Entity CircleEntity { get; set; } //keep the circle entity

        public Face2 AttachedFace { get; set; } //keep the face that owns the circle entity

        public AddedReferencePlane AttachedRefPlane { get; set; } //keep the reference plane that owns the circle entity

    }

    //class for keeping the ellips (non circle) pattern
    public class NonCircularPattern
    {
        public Entity NonCircleEntity { get; set; } //keep the non circle entity

        public Face2 AttachedFace { get; set; } //keep the face that owns the non circle entity

        public AddedReferencePlane AttachedRefPlane { get; set; } //keep the reference plane that owns the circle entity

    }

    //class for pattern configuration
    public class PatternConfig
    {
        public string Type { get; set; } //mix, line, circle, or curve only.

        public int Line { get; set; }

        public int Circle { get; set; }

        public int Curve { get; set; }
    }

    //class for keeping the plane's tolerance
    public class PlaneTolerance
    { 
        public string datum { get;set; }
        public string flatness { get;set; }
        public string parallelism { get;set; }
        public string angularity { get;set; }
        public string perpendicularity { get;set; }
        public string location { get;set; }
    }



}
