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
using SolidWorksTools;
using SolidWorksTools.File;
using System.Collections.Generic;
using System.Diagnostics;


namespace cs_ppx
{
    /// <summary>
    /// Summary description for cs_ppx.
    /// </summary>
    [Guid("8c7d635f-735d-49a0-9259-08b4a9d926f3"), ComVisible(true)]
    [SwAddin(
        Description = "cs_ppx description",
        Title = "cs_ppx",
        LoadAtStartup = true
        )]
    public class SwAddin : ISwAddin
    {
        #region Local Variables
        ISldWorks iSwApp = null;
        ICommandManager iCmdMgr = null;
        int addinID = 0;
        BitmapHandler iBmp;

        public const int mainCmdGroupID = 20;
        public const int mainItemID1 = 0;
        public const int mainItemID2 = 1;

        //just added (for tab registration)
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
        
        public const int flyoutGroupID = 91;

        

        #region Event Handler Variables
        Hashtable openDocs = new Hashtable();
        SolidWorks.Interop.sldworks.SldWorks SwEventPtr = null;
        #endregion

        #region Property Manager Variables
        UserPMPage ppage = null;
        #endregion


        // Public Properties
        public ISldWorks SwApp
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

            return true;
        }

        public bool DisconnectFromSW()
        {
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

            return true;
        }
        #endregion

        #region UI Methods
        public void AddCommandMgr()
        {
            ICommandGroup cmdGroup;
            if (iBmp == null)
                iBmp = new BitmapHandler();
            Assembly thisAssembly;
            int cmdIndex0, cmdIndex1, cmdIndex2, cmdIndex3, cmdIndex4, cmdIndex5, cmdIndex6, cmdIndex7, cmdIndex8, cmdIndex9, cmdIndex10, cmdIndex11;
            string Title = "ppx", ToolTip = "flexible process planning";


            int[] docTypes = new int[]{(int)swDocumentTypes_e.swDocASSEMBLY,
                                       (int)swDocumentTypes_e.swDocDRAWING,
                                       (int)swDocumentTypes_e.swDocPART};

            thisAssembly = System.Reflection.Assembly.GetAssembly(this.GetType());


            int cmdGroupErr = 0;
            bool ignorePrevious = false;

            object registryIDs;
            //get the ID information stored in the registry
            bool getDataResult = iCmdMgr.GetGroupDataFromRegistry(mainCmdGroupID, out registryIDs);

            int[] knownIDs = new int[12] { mainItemID1, mainItemID2, mainItemID3, mainItemID4, mainItemID5, mainItemID6, mainItemID7, mainItemID8, mainItemID9, mainItemID10, mainItemID11, mainItemID12 };

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
            cmdIndex0 = cmdGroup.AddCommandItem2("Load Samples", -1, "Load all raw model and product samples", "Load Samples", 2, "LoadSamples", "", mainItemID1, menuToolbarOption);
            cmdIndex1 = cmdGroup.AddCommandItem2("Select Sample", -1, "Select sample available", "Select Sample", 2, "ShowPMP", "EnablePMP", mainItemID2, menuToolbarOption);
            cmdIndex2 = cmdGroup.AddCommandItem2("Match Body", -1, "Matching material and product", "Match Body", 2, "MatchBody", "", mainItemID3, menuToolbarOption);
            cmdIndex3 = cmdGroup.AddCommandItem2("Main TRV", -1, "Generate TRV from raw model and product", "Main TRV", 2, "MainTRV", "", mainItemID4, menuToolbarOption);
            cmdIndex4 = cmdGroup.AddCommandItem2("Plane Generator", -1, "Generate all planes", "Plane Generator", 2, "PlaneGenerator", "", mainItemID5, menuToolbarOption);
            cmdIndex5 = cmdGroup.AddCommandItem2("Plane Calculator", -1, "Calculating relationship matrix", "Plane Calculator", 2, "PlaneCalculator", "", mainItemID6, menuToolbarOption);
            cmdIndex6 = cmdGroup.AddCommandItem2("TRV Feature", -1, "Generate TRV features", "TRV Feature", 2, "TRVFeature", "", mainItemID7, menuToolbarOption);
            cmdIndex7 = cmdGroup.AddCommandItem2("Machinable Space", -1, "Machinable space calculation", "Machinable Space", 2, "MachinableSpace", "", mainItemID8, menuToolbarOption);
            cmdIndex8 = cmdGroup.AddCommandItem2("TRV network", -1, "TRV network calculation", "TRV Network", 2, "TRVNetwork", "", mainItemID9, menuToolbarOption);
            cmdIndex9 = cmdGroup.AddCommandItem2("Calculate Setup", -1, "Setup calculation", "Calculate Setup", 2, "SetupCalculator", "", mainItemID10, menuToolbarOption);
            cmdIndex10 = cmdGroup.AddCommandItem2("Set Open Face", -1, "Defining a open face", "Set Open Face", 2, "SetOpenFace", "", mainItemID11, menuToolbarOption);
            cmdIndex11 = cmdGroup.AddCommandItem2("Read Open Face", -1, "Reading a open face", "Read Open Face", 2, "ReadOpenFace", "", mainItemID12, menuToolbarOption);
            

            cmdGroup.HasToolbar = true;
            cmdGroup.HasMenu = true;
            cmdGroup.Activate();
            
            bool bResult;
            
            FlyoutGroup flyGroup = iCmdMgr.CreateFlyoutGroup(flyoutGroupID, "Dynamic Flyout", "Flyout Tooltip", "Flyout Hint",
              cmdGroup.SmallMainIcon, cmdGroup.LargeMainIcon, cmdGroup.SmallIconList, cmdGroup.LargeIconList, "FlyoutCallback", "FlyoutEnable");


            flyGroup.AddCommandItem("FlyoutCommand 1", "test", 0, "FlyoutCommandItem1", "FlyoutEnableCommandItem1");

            flyGroup.FlyoutType = (int)swCommandFlyoutStyle_e.swCommandFlyoutStyle_Simple;


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

                    //show pmp button
                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex1);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    //match body button
                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex2);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
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

                    //add another group
                    CommandTabBox cmdBox2 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[3];
                    TextType = new int[3];

                    //Machinable space calculation button
                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex7);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    //TRV network calculation buttion
                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex8);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                                        
                    //Setup calculation button
                    cmdIDs[2] = cmdGroup.get_CommandID(cmdIndex9);
                    TextType[2] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    bResult = cmdBox2.AddCommands(cmdIDs, TextType);
                    cmdTab.AddSeparator(cmdBox2, cmdGroup.ToolbarId);

                    //add another group
                    CommandTabBox cmdBox3 = cmdTab.AddCommandTabBox();
                    cmdIDs = new int[2];
                    TextType = new int[2];

                    //Set open face button
                    cmdIDs[0] = cmdGroup.get_CommandID(cmdIndex10);
                    TextType[0] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;

                    cmdIDs[1] = cmdGroup.get_CommandID(cmdIndex11);
                    TextType[1] = (int)swCommandTabButtonTextDisplay_e.swCommandTabButton_TextBelow;
                    
                    bResult = cmdBox3.AddCommands(cmdIDs, TextType);

                    cmdTab.AddSeparator(cmdBox3, cmdGroup.ToolbarId);
                    
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
            iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble
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
                    
                    Component2[] compName = new Component2[2];
                    
                    GetCompName(compInAssembly, ref compName);
                                        
                    boolStatus = compName[0].Select2(true, 0);

                    ModelDoc2 ModDoc2 = (ModelDoc2)compName[0].GetModelDoc2();
                    markBBfaces(compName[0], ref BBDef);

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
                                        
                    Doc.ClearSelection2(true);

                    //change the material color to red
                    changeColor(ref compName[0]);

                    Doc.GraphicsRedraw2();
                    Doc.ForceRebuild3(false);
                                        
                    Array BodyArray = (Array) compName[0].GetBodies2((int)swBodyType_e.swSolidBody);

                    Body2 TmpBody = (Body2)BodyArray.GetValue(0);

                    Entity TmpEntity = null;
                    Face2 TmpFace = (Face2)TmpBody.GetFirstFace();

                    SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);
                    
                    int i = 0;

                    while (swAtt == null && TmpFace != null)
                    {
                       TmpEntity = (Entity)TmpFace;

                        while (swAtt == null && i < 300)
                        {
                            swAtt = TmpEntity.FindAttribute(NFDef, i);
                            i++;
                        }

                        TmpFace = (Face2)TmpFace.GetNextFace();
                        i = 0;
                    }
                    if (swAtt != null)
                    {
                        TmpEntity.Select(true);
                    }
                    

                }
            }
            else
            {
                iSwApp.SendMsgToUser("There is no model document");
            }

        }

        //tools for maintrv
        # region MainTRV
        
        //traverse component and get the components name
        public void GetCompName(Component2 components, ref Component2[] compName)
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

        public Component2[] compName;

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

        //set an attribute to the the new created faces
        public bool SetAttribute(Component2 SwComp, ref AttributeDef NewFace)
        {
            Array BodyArray;
            Object BodiesInfo;
            Body2 TmpBody;
            Face2 SwFace;
            Entity SwEntity;
            int Errors = 0;
            int Warnings = 0;
            Boolean bRet = false;
            int faceIndex = 1;

            //create the BB attribute definition
            AttributeDef attDef = SwApp.DefineAttribute("newface");
            NewFace = attDef;
            bRet = attDef.AddParameter("newface", (int)swParamType_e.swParamTypeDouble, 1, 0);
            bRet = attDef.Register();

            String CompPathName = SwComp.GetPathName();
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

            BodyArray = null;
            BodiesInfo = null;
            TmpBody = null;

            PartDoc PartDocument = (PartDoc)CompDocumentModel;
            BodyArray = (Array)PartDocument.GetBodies2((int)swBodyType_e.swSolidBody, true);

            if (BodyArray.Length == 1)
            {
                TmpBody = (Body2)BodyArray.GetValue(0);

                SwFace = null;
                SwEntity = null;

                SwFace = (Face2)TmpBody.GetFirstFace();
                while (SwFace != null)
                {
                    SwEntity = (Entity)SwFace;

                    SolidWorks.Interop.sldworks.Attribute swAtt = default(SolidWorks.Interop.sldworks.Attribute);

                    int i = 0;

                    while (swAtt == null && i < 300)
                    {
                        swAtt = SwEntity.FindAttribute(BBDef, i);
                        i++;
                    }

                    if (swAtt == null)
                    {
                        swAtt = attDef.CreateInstance5(CompDocumentModel, SwFace, "new_face" + faceIndex.ToString(), 0, (int)swInConfigurationOpts_e.swAllConfiguration);
                        faceIndex++;
                    }

                    SwFace = (Face2)SwFace.GetNextFace();
                }

            }
            else if (BodyArray.Length > 1)
            {


            }
           
            //iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            //iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be visble

            return true;
           
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
                    
                    //define the raw material and product
                    Component2[] compName = new Component2[2];
                    GetCompName(compInAssembly, ref compName);

                    //get faces from raw material (at this stage is main trv)
                    object bodyEnts = (object) compName[0].GetBody();
                    Body2 swBody = (Body2)bodyEnts;
                    object[] swFaces = (object[])swBody.GetFaces();

                    //get the part document of the raw material
                    ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                    PartDoc swPartDoc = (PartDoc)compModDoc;

                    //prepare for face traversal
                    string entName = null;
                    string refName = null;
                    int countFace = 0;
                    RefPlane swRefPlane = null;
                    Entity swEnt = null;

                    AddedReferencePlane tmpInitPlane = null;
                    InitialRefPlane = new List<AddedReferencePlane>(); //set the instance for the first time for sacing all the reference planes
                                     
                    //select the raw material for editing
                    boolStatus = compName[0].Select2(true, 0);

                    if (boolStatus == true)
                    {
                        assyModel.EditPart();

                        Array AllBodies = (Array)swPartDoc.GetBodies2((int)swBodyType_e.swSolidBody, true);
                                                                        
                        foreach (Face2 tmpFace in swFaces)
                        {
                            entName = swPartDoc.GetEntityName(tmpFace);
                            
                            //select the face without name
                            if (entName == "")
                            {
                                swEnt = (Entity)tmpFace;
                                swEnt.Select2(true, 0);

                                //create coincide reference plane to it
                                swRefPlane = (RefPlane)Doc.FeatureManager.InsertRefPlane(4, 0, 0, 0, 0, 0);
                                countFace += 1;

                                //set reference name
                                refName = "REF_PLANE" + countFace.ToString();
                                swPartDoc.SetEntityName(swRefPlane, refName);

                                //set the reference plane, included with name, coincided face, and its pointer
                                tmpInitPlane = new AddedReferencePlane();
                                tmpInitPlane.name = refName;
                                tmpInitPlane.AttachedFace = tmpFace;
                                tmpInitPlane.ReferencePlane = swRefPlane;

                                //add the reference plane
                                InitialRefPlane.Add(tmpInitPlane);

                            }
                        }
                    }

                    //store number of reference plane
                    addedRefPlane = countFace;

                    //close editing the raw material
                    assyModel.EditAssembly();

                    boolStatus = compName[0].Select2(true, 0);
                    
                }
            }

            Doc.ClearSelection2(true);
        }

        //tools for plane generator
        #region PlaneGenerator

        //save the number of generated reference plane
        public int addedRefPlane { get; set; }

        //save all generate reference plane
        public List<AddedReferencePlane> InitialRefPlane;
        
        #endregion 

        //plane Calculator
        public void PlaneCalculator()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            MathUtility swMathUtils = (MathUtility)SwApp.GetMathUtility();
            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            
            AssemblyDoc assyModel = null;
            Component2[] compName = null;

            if (Doc.GetType() == 2)
            {
                assyModel = (AssemblyDoc)Doc;
                if (assyModel.GetComponentCount(false) != 0)
                {
                    //initiate plane properties
                    List<Feature> planeNames = null;
                    List<object> planeNormal = null;
                    List<int> planeValue = null;
                    List<double> distance = null;

                    try
                    {
                        //get the components
                        ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                        Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                        Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                        //define the raw material and product
                        compName = new Component2[2];
                        GetCompName(compInAssembly, ref compName);

                        //get the virtual centroid
                        object boxVertices = compName[0].GetBox(false, false);
                        virtualCentroid = getMidPoint(boxVertices);
                        
                        //get the real centroid
                        Body2 mainBody = compName[0].GetBody();
                        MassProperty bodyProperties = Doc.Extension.CreateMassProperty();
                        bodyProperties.AddBodies(mainBody);
                        bodyProperties.UseSystemUnits = false;
                        centroid = (double[]) bodyProperties.CenterOfMass;
                        
                        //draw the virtual centroid
                        Doc.SketchManager.Insert3DSketch(true);
                        double[] tmpPoint = (double[])virtualCentroid;
                        SketchPoint skPoint = (SketchPoint)Doc.SketchManager.CreatePoint(tmpPoint[0], tmpPoint[1], tmpPoint[2]);
                        Doc.SketchManager.Insert3DSketch(true);

                        //get the part document of the raw material
                        ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                        PartDoc swPartDoc = (PartDoc)compModDoc;

                        //get the plane feature
                        boolStatus = getPlanes(swPartDoc, ref InitialRefPlane);

                        if (boolStatus == true)
                        {
                            //get the plane normal
                            boolStatus = getPlaneNormal(swMathUtils, ref InitialRefPlane);
                            if (boolStatus == true)
                            {
                                //analyze intersection
                                boolStatus = planeIntersection(planeNormal, ref planeValue); // <=================== DO THIS PART

                                //set the distance
                                setDistance(Doc, swMathUtils, ref InitialRefPlane, tmpPoint);
                            }
                        }

                        if (boolStatus == true)
                        {
                            List<int> removeId = new List<int>();

                            //store the plane feature, rank, and normal to planeList
                            //registerPlane(planeValue, planeNames, planeNormal, distance, ref removeId, swMathUtils);

                            registerPlane(InitialRefPlane, ref removeId, swMathUtils);

                            if (removeId.Count > 0)
                            {
                                suppressFeature(Doc, assyModel, compName[0], planeNames, removeId);
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
                    RefPlanes[i].CorrespondFeature = tmpFeature;
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
            }

            return true;

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

                    if (tmpDistance != null)
                    {
                        distance.Add(tmpDistance);
                    }
                }
            }

        }

        //OVERLOAD setDistance
        public void setDistance(ModelDoc2 Doc, MathUtility swMathUtils, ref List<AddedReferencePlane> RefPlanes, Double[] centroid)
        {
            double TmpDistance;

            if (RefPlanes.Count() > 0)
            {
                for (int i = 0; i < RefPlanes.Count; i++)
                {
                    TmpDistance = new double();
                    RefPlane TmpRefPlane = RefPlanes[i].ReferencePlane;
                    Object TmpNormal = RefPlanes[i].PlaneNormal;

                    TmpDistance = checkDistance(Doc, TmpRefPlane, TmpNormal, centroid, swMathUtils);

                    //tmpDistance = checkDistance(Doc, tmpFace, planeNormal[i], centroid, swMathUtils);

                    if (TmpDistance != null)
                    {
                        RefPlanes[i].DistanceFromCentroid = TmpDistance;
                        
                    }
                }
            }

        }

        //get face that corresponds to refplane
        public Face2 getAttachedFace(Feature feature)
        {
            if (InitialRefPlane.Count != 0)
            {
                foreach (AddedReferencePlane CheckRefPlane in InitialRefPlane)
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
            bool retSim = false;

            planeList = new List<_planeProperties>();

            for (int i = 0; i < ListOfPlanes.Count; i++)
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

            foreach (AddedReferencePlane tmpPlane in ListOfPlanes)
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

            //if similar, parallel and coincident, then it should return true
            return similarity;
        }

        //check normal direction similarity
        public bool checkNormalSim(object firstN, object secondN, MathUtility mathUtils)
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
            return isEqual(dotProduct, 1);

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
        public void suppressFeature(ModelDoc2 Doc, AssemblyDoc assyDoc, Component2 compName, List<Feature> feature2Del, List<int> removeId)
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
                        tmpFeature = swPartDoc.FeatureByName(feature2Del[i].Name);
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
            return (Math.Abs(Math.Round((value1 - value2), 4)) < 0.0001);
            
        }
            
        //storage for all generated plane
        public List<_planeProperties> planeList;
         

        #endregion

        //TRV feature
        public void TRVfeature()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            
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
                    List<Feature> featureList = null;

                    //get the components
                    ConfigurationManager configDocManager = (ConfigurationManager)Doc.ConfigurationManager;
                    Configuration configDoc = (Configuration)configDocManager.ActiveConfiguration;
                    Component2 compInAssembly = (Component2)configDoc.GetRootComponent3(true);

                    //define the raw material and product
                    Component2[] compName = new Component2[2];
                    GetCompName(compInAssembly, ref compName);
                    
                    boolStatus = compName[0].Select2(true, 0);

                    TraverseComponentFeatures(compName[0], ref featureList);

                    //initMachiningPlan();

                    List<Body2> RVList = new List<Body2>();
                                       
                    traversePlanes(Doc, assyModel, compName[0], featureList, ref RVList);

                    iSwApp.SendMsgToUser("Number of process: " + RVList.Count.ToString());

                    boolStatus = compName[0].Select2(true, 0);

                }
            }

            //iSwApp.SendMsgToUser("this is from TRV feature");
        }

        //tools for calculating TRV feature
        #region TRV Feature

        //initiate the traverse with the first feature
        public void TraverseComponentFeatures(Component2 swComp, ref List<Feature> planeList)
        {
            Feature swFeat;

            swFeat = (Feature)swComp.FirstFeature();
            planeList = new List<Feature>();

            TraverseFeatures(swFeat, ref planeList);

        }
        
        //traverse to get previously created planes
        public void TraverseFeatures(Feature swFeat, ref List<Feature> planeList)
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

        //traverse planes
        public void traversePlanes(ModelDoc2 Doc, AssemblyDoc assyModel, Component2 swComp, List< Feature> featureList, ref List<Body2> VolumeList)
        {
            Feature selFeature = null;
            bool boolStatus;
            int index = 0;
            int SplitCounter = 0;
            int Errors = 0;
            int Warnings = 0;

            
            String CompPathName = swComp.GetPathName();
            iSwApp.DocumentVisible(false, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be invisble
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document
            
            
            //setMarkOnPlane(ref planeList, 5, 4);
            //setMarkOnPlane(ref planeList, 6, 4);

            //int selectedMP = 1;
            int position = 0;

            //selFeature = getByMP(getNextPlan(selectedMP, position), featureList);

            selFeature = getThePlane(planeList, featureList, ref index);

            //selFeature = featureList[4];
            //index = 3;

            while (selFeature != null)
            {
                assyModel.EditPart();
               
                boolStatus = selFeature.Select2(true, 0);
                //int longstatus = 0;
                
                //split and collect the body
                List<int> DeleteThisBody = null;
                Array bodyArray = null;
                bodyArray = (Array)Doc.FeatureManager.PreSplitBody();
                
                //check if there are body 
                if (bodyArray == null)
                {
                    setMarkOnPlane(ref planeList, index, 1); //mark the plane because no bodies was collected
                }
                else
                {  
                    string[] bodyNames = new string[bodyArray.Length]; //set the name
                    
                    Feature myFeature = null;
                    
                    myFeature = SplitAndSaveBody(Doc, swComp, selFeature, bodyArray, ref bodyNames);

                    if (myFeature != null)
                    {
                        try
                        {
                            String SelectedBody = null;
                            List<double> Volume = new List<double>();
                            Object BodiesInfo = null;
                            List<string> BodyToDelete = null;

                            DeleteThisBody = new List<int>();
                            SplitCounter++; //add the split counter after successful splitting process

                            bodyArray = null;
                            bodyArray = (Array) swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                            //select body that needs to be deleted from the model
                            bool Status = SelectBodyToDelete(CompDocumentModel, bodyArray, ref BodyToDelete, index, ref Volume);

                            if (Status == true)
                            {
                                iSwApp.SendMsgToUser("Feasible body exist, ready to be registered in the removal sequence");
                                

                                if (BodyToDelete.Count > 1)
                                {
                                    //iSwApp.SendMsgToUser("Body > 1 still underdevelopment");

                                    foreach (String BodyName in BodyToDelete)
                                    {
                                        BodiesInfo = null;
                                        Body2 TmpBody = null;
                                        int DeleteIndex = 0;
                                        bodyArray = null;

                                        bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                        for (int j = 0; j < bodyArray.Length; j++)
                                        {
                                            TmpBody = (Body2)bodyArray.GetValue(j);

                                            if (TmpBody.Name.Equals(BodyName))
                                            {
                                                DeleteIndex = j;
                                                break;
                                            }
                                        }

                                        SelectData TmpSelectData = null;

                                        bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);

                                        VolumeList.Add(TmpBody);

                                        SwApp.SendMsgToUser("Removed shape by " + selFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());

                                        myFeature = (Feature)Doc.FeatureManager.InsertDeleteBody();
                                    }

                                }
                                else
                                {
                                    
                                    //prepare the deletion of the splitted body
                                    BodiesInfo = null;
                                    Body2 TmpBody = null;
                                    int DeleteIndex = 0;
                                    bodyArray = null;

                                    bodyArray = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                                    for (int j = 0; j < bodyArray.Length; j++)
                                    {
                                        TmpBody = (Body2)bodyArray.GetValue(j);

                                        if (TmpBody.Name.Equals(BodyToDelete[0]))
                                        {
                                            DeleteIndex = j;
                                            break;
                                        }
                                    }

                                    SelectData TmpSelectData = null;

                                    bool SelectionStatus = TmpBody.Select2(true, TmpSelectData);

                                    VolumeList.Add(TmpBody);

                                    SwApp.SendMsgToUser("Removed shape by " + selFeature.Name.ToString() + "\r\" Volume:  " + Volume[DeleteIndex].ToString());

                                    myFeature = (Feature)Doc.FeatureManager.InsertDeleteBody();

                                }

                                //set 0 mark that indicates this plane can be set with removal volume
                                setMarkOnPlane(ref planeList, index, 0);

                            }
                            else
                            {
                                //set 3 on this plane that indicates nothing can be set to this plane
                                setMarkOnPlane(ref planeList, index, 3);

                                ModelDoc2 DocumentModel = (ModelDoc2) swComp.GetModelDoc2();

                                string ComponentName = swComp.Name2;

                                Feature LastFeature = (Feature) swComp.FeatureByName("Split" + SplitCounter.ToString());

                                string FeatureName = LastFeature.Name;

                                bool SStatus = LastFeature.Select2(true, 3);

                                Doc.EditDelete();

                                //myFeature = (Feature)Doc.FeatureManager.InsertDeleteBody();

                                //SplitCounter--;

                            }

                        }

                        catch (Exception ex)
                        {
                            iSwApp.SendMsgToUser("exception message: " + ex.Message);
                        }

                        finally
                        {
                            //Object BodiesInfo = null;
                            //Array NewBodies = (Array)swComp.GetBodies3((int)swBodyType_e.swSolidBody, out BodiesInfo);

                        }

                    }

                    else
                    { 
                        //set 2 on this plane that indicates nothing can be found
                        setMarkOnPlane(ref planeList, index, 2);

                    }

                    position++;

                }

                Doc = ((ModelDoc2)(SwApp.ActiveDoc));
                Doc.ClearSelection2(true);

                selFeature = getThePlane(planeList, featureList, ref index);

                //selFeature = null;

            }

            //Doc.ClearSelection2(true);

            iSwApp.CloseDoc(Path.GetFileNameWithoutExtension(CompPathName));
            iSwApp.DocumentVisible(true, (int)swDocumentTypes_e.swDocPART); //make the loaded document to be visble
             
        }

        //getNextPlane
        public _planeProperties getNextPlan(int index, int position)
        {
            Object[] tmpListMP = (Object[]) machiningPlanList[index];

            if (position < tmpListMP.Length)
            {
                //machiningPlan MP = (machiningPlan)tmpListMP[position];
                //return  MP.planeProperties;
            }

            return null;

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
                            bodyStatus = setBodyToPlane(ModDoc, BodyToCheck, Centroid, planeList[index]);

                            if (bodyStatus == true)
                            {
                                iSwApp.SendMsgToUser("This body is feasible (convex)");

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
        public bool SelectBodyToDelete(ModelDoc2 CompDocumentModel, Array BodyArray, ref List<string> BodyToDelete, int index, ref List<double> VolumeSize)
        {
            
            BodyToDelete = new List<string>();

            if (BodyArray.Length == 0)
            {
                return false;
            }

            else
            {   
                for (int i = 0; i < BodyArray.Length; i++)
                {
                    bool bodyStatus = false;

                    Body2 BodyToCheck = (Body2)BodyArray.GetValue(i);

                    iSwApp.SendMsgToUser("Check bodies Loaded document: " + BodyToCheck.Name);
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

                        bodyStatus = setBodyToPlane(CompDocumentModel, TmpBodyToCheck, Centroid, planeList[index]);

                        if (bodyStatus == true)
                        {
                            iSwApp.SendMsgToUser("This body is feasible (convex)");

                            //check this body in the document tree and add collect the body pointer to the FeasibleBodies
                            BodyToDelete.Add(BodyToCheck.Name);

                            //set this body to the current plane
                        }

                        else
                        {
                            iSwApp.SendMsgToUser("This body is not feasible (concave)");

                            //keep the index of the body

                        }
                    }
                }
            }

            if (BodyToDelete.Count == 0) { return false; }
            else { return true; }

        }
        
        //split and save bodies
        public Feature SplitAndSaveBody(ModelDoc2 DocumentModel, Component2 SwComp, Feature SelectedFeature, Array BodyArray, ref string[] BodyNames)
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
        public string getSavePath(string modelPath)
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
        public bool getCentroidAndVolume(Body2 BodyToCheck, ref double[] Centroid, ref List<double> Volume)
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

        public List<Object> machiningPlanList;

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
        public Feature getByMP(_planeProperties plane, List<Feature> _featureList)
        {
            if (plane != null)
            {
                return getSelectedPlane(plane, _featureList);
            }

            return null;
        }

        //get and return the plane
        public Feature getThePlane(List<_planeProperties> _planeList, List<Feature> _featureList, ref int index)
        {
            //bool boolResult = false;
            //Feature nextPlane = null;
            Feature selectedPlane = null;

            int selectedIndex;
            selectedIndex = getPlaneIndex(_planeList);

            if (selectedIndex != -1)
            {
                selectedPlane = getSelectedPlane(_planeList[selectedIndex], _featureList);
                index = selectedIndex;
            }

            //nextPlane = selectPlanes(_planeList, _featurelist, ref index);

            return selectedPlane;

        }

        //get and return the plane index
        public int getPlaneIndex(List<_planeProperties> _planeList)
        {
            int selectedIndex = -1;
            int tmpRank = 0;
            int rankSimilarity = 0;
            int distSimilarity = 0;
            double tmpDistance = 0.0;

            foreach (_planeProperties tmpPlane in _planeList)
            {
                if (tmpPlane.mark == 0)
                {
                    if (tmpPlane.rankByDistance == tmpRank)
                    {
                        rankSimilarity++;
                    }
                    else if (tmpPlane.rankByDistance > tmpRank)
                    {
                        selectedIndex = _planeList.IndexOf(tmpPlane);
                        tmpRank = tmpPlane.rankByDistance;
                        rankSimilarity = 1;
                    }
                    
                }
            }

            if (rankSimilarity == 0)
            {
                return selectedIndex;
            }
            else
            {
                foreach (_planeProperties tmpPlane in _planeList)
                { 
                    if (tmpPlane.mark == 0)
                    {
                        if (isEqual(tmpPlane.distance, tmpDistance) == true)
                        {
                            distSimilarity++;
                        }
                        else if (tmpPlane.distance > tmpDistance)
                        {
                            selectedIndex = _planeList.IndexOf(tmpPlane);
                            tmpDistance = tmpPlane.distance;
                            distSimilarity = 1;
                        }
                    }
                }
            }

            return selectedIndex;

        }

        //get the similar plane and pass the pointer to selected feature
        public Feature getSelectedPlane(_planeProperties listOfPlane, List<Feature> listOfModelPlane)
        {
            Feature tmpFeature = null;

            for (int i = 0; i < listOfModelPlane.Count; i++)
            {   
                //tmpFeature = listOfModelPlane[i];

                if (listOfModelPlane[i].Name == listOfPlane.featureObj.Name)
                {
                    tmpFeature = listOfModelPlane[i];
                    
                }
            }

            return tmpFeature;
        }

        //mark the plane according the message code
        public bool setMarkOnPlane(ref List<_planeProperties> planeList, int index, int message)
        {
            switch (message)
            {
                //if the process runs ok, mark the plane
                //as 1 to avoid the system select it again    
                case 0:
                    planeList[index].mark = 1;

                    renewIntersection(ref planeList, index);
                    break;

                //if no body found
                case 1:
                    planeList[index].mark = 2;
                    break;

                //if no body can be set to the corresponding plane
                case 2:
                    planeList[index].mark = 3;
                    break;

                //split body process has failed
                case 3:
                    planeList[index].mark = 4;
                    break;

                //safe normal critearia is not achieved
                case 4:
                    planeList[index].mark = 5;
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
        public bool setBodyToPlane(ModelDoc2 Doc, Body2[] bodyList, _planeProperties selectedPlane, ref List<int> selectedId)
        {
            List<Body2> tmpBodyList = new List<Body2>();
            selectedId = new List<int>();

            bool returnValue = false;
            
            try
            {
                int index = 0;

                foreach (Body2 tmpBody in bodyList)
                {
                    if (isBodyConvex(Doc, tmpBody) == true)
                    {
                        if (checkBodyLocation(tmpBody, selectedPlane.featureObj, selectedPlane.planeNormal) == true)
                        {
                            tmpBodyList.Add(tmpBody);
                            selectedId.Add(index);
                        }

                    }

                    index++;
                    
                }
            }

            finally
            {
                if (tmpBodyList.Count > 0)
                {
                    //tmpBodyList.CopyTo(selectedPlane.bodyList);
                    selectedPlane.bodyList = tmpBodyList;
                    returnValue = true;
                }
            }

            return returnValue;
        }

        //OVERLOAD setBodyToPlane
        public bool setBodyToPlane(ModelDoc2 DocumentModel, Body2 BodyToCheck, Double[] Centroid, _planeProperties SelectedPlane)
        {
            bool Status = false;

            if (isBodyConvex(DocumentModel, BodyToCheck) == true)
            {
                if (checkBodyLocation(BodyToCheck, Centroid, SelectedPlane.featureObj, SelectedPlane.planeNormal) == true)
                {
                    Status = true;
                }

            }

            return Status;
        }

        //check whether if the body convex or concave, return true if convex
        public bool isBodyConvex(ModelDoc2 Doc, Body2 body2Check)
        {
            object[] objFaces = null;
            //Face2[] faceList = null;
            bool retVal = true;

            objFaces = (object[])body2Check.GetFaces();

            List<Face2> filteredFaces = new List<Face2>();

            filteredFaces = filterFaces(Doc, objFaces);

            //faceList = (Face2[])objFaces;

            foreach (Face2 tmpFace in objFaces)
            {
                retVal = isFaceConvex(tmpFace);

                if (retVal == false)
                {
                    break;
                }
            }

            /*
            for (int i = 0; i < faceList.Length; i++)
            {
                retVal = isFaceConvex(faceList[i]);

                if (retVal == false)
                {
                    break;
                }
            }
             */ 

            return retVal;
        }

        //filter the face
        public List<Face2> filterFaces(ModelDoc2 Doc, object[] faceList)
        {
            object pointA, pointB;
            List<Face2> returnFaces = new List<Face2>();
            List<Face2> tmpFaces = new List<Face2>();
            List<bool> boolFaces = new List<bool>();

            List<int> equality = new List<int>();
            
            int index = 0;
            int similarity = 0;
            
            foreach (Face2 faceA in faceList)
            {
                bool tmpBool = false;

                tmpBool = isFaceConvex(faceA);

                if (tmpBool == true)
                {
                    tmpFaces.Add(faceA);
                }

                boolFaces.Add(tmpBool);
            }

            foreach (Face2 faceA in faceList)
            {
                if (boolFaces[index] == false)
                {
                    similarity = 0;

                    foreach (Face2 faceB in tmpFaces)
                    {   
                        Double distance = Doc.ClosestDistance(faceA, faceB, out pointA, out pointB);

                        if (isEqual(distance, 0))
                        {
                            similarity++;
                        }
                    }

                    if (similarity == 0)
                    {
                        tmpFaces.Add(faceA);
                    }
                }
            
                index++;
            }
            
            return tmpFaces;
        }

        //check whether if the face convex or concave
        public bool isFaceConvex(Face2 face2Check)
        {
            Object[] objLoops = null;
            //Loop2[] loopList = null;

            bool retVal = false;

            objLoops = (Object[])face2Check.GetLoops();

            //loopList = (Loop2[])objLoops;

            foreach (Loop2 tmpLoop in objLoops)
            {
                if (tmpLoop.IsOuter() == true)
                {
                    //process this loop
                    retVal = isLoopConvex(tmpLoop);
                }
            }

            /*
            for (int i = 0; i < loopList.Length; i++)
            {
                if (loopList[i].IsOuter() == true)
                { 
                    //process this loop
                    retVal = isLoopConvex(loopList[i]);
                }
            }
             */ 

            return retVal;

        }

        //check whether if the loop is convex or concave
        public bool isLoopConvex(Loop2 loop2Check)
        {
            Object[] objVertices = null;
            List<Vertex> verticesList = new List<Vertex>();
            Vertex[] verticesArray = null;
            double[] zCrossProduct = null;
            int length = 0;

            bool returnValue = false;
            
            objVertices = (Object[])loop2Check.GetVertices();

            //verticesList = (Vertex[])objVertices;
            //length = verticesList.Length;

            foreach (Vertex tmpVertex in objVertices)
            {
                verticesList.Add(tmpVertex);
            }

            verticesList.Add(objVertices[0] as Vertex);
            verticesList.Add(objVertices[1] as Vertex);

            verticesArray = new Vertex[verticesList.Count];
            
            verticesList.CopyTo(verticesArray);

            length = verticesArray.Length;

            /* 
            //add first two value at the last of the array
            for (int i = 0; i < 2; i++)
            {
                //verticesList[length + i] = verticesList[i];
                objVertices[length + i] = objVertices[i];
            }
             * */

            if (checkTriplets(ref zCrossProduct, verticesArray) == true)
            {
                returnValue = checkSign(zCrossProduct);
            }
            
            return returnValue;
        }

        //check triplets
        public bool checkTriplets(ref double[] crossProduct, Vertex[] vertexList)
        {
            Object objPoint1, objPoint2, objPoint3 = null;
            Double[] point1, point2, point3 = null;
            //Vertex tmpVertex1, tmpVertex2, tmpVertex3 = null;
            double dx1, dx2, dy1, dy2;

            crossProduct = new double[vertexList.Length - 2];
            
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

                dx1 = point2[0] - point1[0];
                dx2 = point3[0] - point2[0];
                dy1 = point2[1] - point1[1];
                dy2 = point3[1] - point2[1];

                crossProduct[i] = (dx1 * dy2) - (dx2 * dy1);
            }

            return true;
        }

        //check the sign of vertex
        public bool checkSign(double[] crossProduct)
        {
            int totalNum = 0;
            bool returnValue = false;
            
            //check first value
            if (crossProduct != null)
            {   
                for (int i = 0; i < crossProduct.Length; i++)
                {
                    if (isEqual(crossProduct[i], Math.Abs(crossProduct[i])))
                    {
                        totalNum++;
                    }
                    else
                    {
                        totalNum--;
                    }
                    
                }

                //if the total number is equal with the number of point, then
                //sign is never change and this means this loop is convex.
                if (Math.Abs(totalNum) == crossProduct.Length)
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
        public bool checkBodyLocation(Body2 body2Check, Double[] Centroid, Feature featurePlane, object objPlaneNormal)
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
        public MathPoint getPointOnPlane(Feature tmpFeature)
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
        public MathPoint getPointOnPlane(RefPlane tmpRefPlane)
        {
            MathPoint tmpPoint = null;
            Object[] arrayOfCorners = null;
            Random rnd = new Random();
            
            arrayOfCorners = (Object[])tmpRefPlane.CornerPoints;
            tmpPoint = (MathPoint)arrayOfCorners[rnd.Next(0, 3)];

            return tmpPoint;
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
            AttributeDef SwAttDef = SwApp.DefineAttribute("open_face");
            Boolean RetVal = SwAttDef.AddParameter("status", (int)swParamType_e.swParamTypeDouble, 1, 0);
            RetVal = SwAttDef.Register();
            
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

    //class for removal body
    public class RemovalBody
    {
        public Body2 BodyObj { get; set; } //keep pointer to body object

        public double[] BodyCentroid { get; set; } //keep the body centroid information

        public List<TAD> ListOfTAD { get; set; } //keep all the candidate of TAD
    }

    //class for plane and removal volume body relation
    public class PlaneWithBody
    {
        public string name { get; set; } //keep the name
        
        public AddedReferencePlane SelectedReferencePlane { get; set; } //keep pointer to selected AddedReferencePlane

        public RemovalBody TMpRemovalBody { get; set; } //keep all bodies that related to AddedReferencePlane

        public TAD SelectedTAD { get; set; } //keep the selected TAD
        
    }

    //class for saving addedReferencePlane
    public class AddedReferencePlane
    {
        public string name { get; set; } //keep the name

        public RefPlane ReferencePlane { get; set; } //keep the pointer to the original plane

        public Face2 AttachedFace { get; set; } //keep the pointer to the corresponding face

        public Feature CorrespondFeature { get; set; } //keep the pointer to the corresponding feature

        public int RankByDistance { get; set; } //keep the rank by the distance

        public Double DistanceFromCentroid { get; set; } //keep the distance from centroid

        public object PlaneNormal { get; set; } //keep the plane normal
        
        public int MarkingOpt { get; set; } //keep the marking options that is used for iterating the plane

        public int Remark { get; set; } //keep additional remark for the plane

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

    }

    //class for TAD
    public class TAD
    {
        public double X { get; set; }
        public double Y { get; set; }
        public double Z { get; set; }
    }

    //class for candidate of machining plan
    public class MachiningPlan
    {
        public List<MachiningProcess> MachiningProceses { get; set; } //keep the pointer of selected machining process

        public double MachiningCost { get; set; } //keep the machining cost

        public double MachiningTime { get; set; } //keep the machining time

        public int NumberOfTADchanges { get; set; } //keep the number of TAD changes

        public int NumberOfTool { get; set; } //keep the number of needed tools

        public int NumberOfSetups { get; set; } //keep the number of setups
    }

}
