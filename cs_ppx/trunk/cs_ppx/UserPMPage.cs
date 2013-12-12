using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swpublished;
using SolidWorks.Interop.swconst;

namespace cs_ppx
{
    public class UserPMPage
    {
        //Local Objects
        IPropertyManagerPage2 swPropertyPage = null;
        PMPHandler handler = null;
        ISldWorks iSwApp = null;
        SwAddin userAddin = null;

        #region Property Manager Page Controls
        //Groups
        IPropertyManagerPageGroup group1;
        IPropertyManagerPageGroup group2;
        IPropertyManagerPageGroup group3;

        //Controls
        IPropertyManagerPageTextbox textbox1;
        IPropertyManagerPageCheckbox checkbox1;
        IPropertyManagerPageOption option1;
        IPropertyManagerPageOption option2;
        IPropertyManagerPageOption option3;
        IPropertyManagerPageListbox list1;

        //set new control for list of rm and list of product
        public IPropertyManagerPageListbox rList, pList;

        IPropertyManagerPageSelectionbox selection1;
        IPropertyManagerPageNumberbox num1;
        IPropertyManagerPageCombobox combo1;
                
        //Control IDs
        public const int group1ID = 0;
        public const int group2ID = 1;
        public const int group3ID = 11;

        public const int textbox1ID = 2;
        public const int checkbox1ID = 3;
        public const int option1ID = 4;
        public const int option2ID = 5;
        public const int option3ID = 6;
        public const int list1ID = 7;

        public const int selection1ID = 8;
        public const int num1ID = 9;
        public const int combo1ID = 10;

        public const int rListID = 11;
        public const int pListID = 12;
        
        #endregion

        public UserPMPage(SwAddin addin)
        {
            userAddin = addin;
            if (userAddin != null)
            {
                iSwApp = (ISldWorks)userAddin.SwApp;
                CreatePropertyManagerPage();
            }
            else
            {
                System.Windows.Forms.MessageBox.Show("SwAddin not set.");
            }
        }


        protected void CreatePropertyManagerPage()
        {
            int errors = -1;
            int options = (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_OkayButton |
                (int)swPropertyManagerPageOptions_e.swPropertyManagerOptions_CancelButton;

            handler = new PMPHandler(userAddin, this);
            swPropertyPage = (IPropertyManagerPage2)iSwApp.CreatePropertyManagerPage("Sample PMP", options, handler, ref errors);
            if (swPropertyPage != null && errors == (int)swPropertyManagerPageStatus_e.swPropertyManagerPage_Okay)
            {
                try
                {
                    AddControls();
                }
                catch (Exception e)
                {
                    iSwApp.SendMsgToUser2(e.Message, 0, 0);
                }
            }
        }


        //Controls are displayed on the page top to bottom in the order 
        //in which they are added to the object.
        protected void AddControls()
        {
            
            short controlType = -1;
            short align = -1;
            int options = -1;


            //Add the groups
            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group1 = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group1ID, "Sample Group 1", options);

            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Checkbox |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group2 = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group2ID, "Sample Group 2", options);

            options = (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Expanded |
                      (int)swAddGroupBoxOptions_e.swGroupBoxOptions_Visible;

            group3 = (IPropertyManagerPageGroup)swPropertyPage.AddGroupBox(group3ID, "sample group 3", options);


            //Add the controls to group1

            //textbox1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Textbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            textbox1 = (IPropertyManagerPageTextbox)group1.AddControl(textbox1ID, controlType, "Type Here", align, options, "This is an example textbox");

            //checkbox1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Checkbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            checkbox1 = (IPropertyManagerPageCheckbox)group1.AddControl(checkbox1ID, controlType, "Sample Checkbox", align, options, "This is a sample checkbox");

            //option1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            option1 = (IPropertyManagerPageOption)group1.AddControl(option1ID, controlType, "Option1", align, options, "Radio Buttons");

            //option2
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            option2 = (IPropertyManagerPageOption)group1.AddControl(option2ID, controlType, "Option2", align, options, "Radio Buttons");

            //option3
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Option;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            option3 = (IPropertyManagerPageOption)group1.AddControl(option3ID, controlType, "Option3", align, options, "Radio Buttons");

            //list1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Listbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            list1 = (IPropertyManagerPageListbox)group1.AddControl(list1ID, controlType, "Sample Listbox", align, options, "List of selectable items");
            if (list1 != null)
            {
                string[] items = { "One Fish", "Two Fish", "Red Fish", "Blue Fish" };
                list1.Height = 50;
                list1.AddItems(items);
            }

            //Add controls to group2
            //selection1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Selectionbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            selection1 = (IPropertyManagerPageSelectionbox)group2.AddControl(selection1ID, controlType, "Sample Selection", align, options, "Displays features selected in main view");
            if (selection1 != null)
            {
                int[] filter = { (int)swSelectType_e.swSelEDGES, (int)swSelectType_e.swSelVERTICES };
                selection1.Height = 40;
                selection1.SetSelectionFilters(filter);
            }

            //num1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Numberbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            num1 = (IPropertyManagerPageNumberbox)group2.AddControl(num1ID, controlType, "Sample Numberbox", align, options, "Allows for numerical input");
            if (num1 != null)
            {
                num1.Value = 50.0;
                num1.SetRange((int)swNumberboxUnitType_e.swNumberBox_UnitlessDouble, 0.0, 100.0, 0.01, true);
            }

            //combo1
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Combobox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            combo1 = (IPropertyManagerPageCombobox)group2.AddControl(combo1ID, controlType, "Sample Combobox", align, options, "Combo list");
            if (combo1 != null)
            {
                string[] items = { "One Fish", "Two Fish", "Red Fish", "Blue Fish" };
                combo1.AddItems(items);
                combo1.Height = 50;

            }

            //raw material list
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Listbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            rList = (IPropertyManagerPageListbox)group3.AddControl(rListID, controlType, "Raw Material Model", align, options, "List of selectable raw material model");
            if (rList != null)
            {
                string[] items = {}; // = { "rm1", "rm2", "rm3", "rm4" };
                rList.Height = 150;
                rList.AddItems(items);
                
            }

            //product list
            controlType = (int)swPropertyManagerPageControlType_e.swControlType_Listbox;
            align = (int)swPropertyManagerPageControlLeftAlign_e.swControlAlign_LeftEdge;
            options = (int)swAddControlOptions_e.swControlOptions_Enabled |
                      (int)swAddControlOptions_e.swControlOptions_Visible;

            pList = (IPropertyManagerPageListbox)group3.AddControl(pListID, controlType, "Product Model", align, options, "List of selectable product model");
            if (pList != null)
            {
                string[] items = { }; //{ "pr1", "pr2", "pr3", "pr4" };
                pList.Height = 150;
                pList.AddItems(items);
            }
                        
        }

        public void Show()
        {
            if (swPropertyPage != null)
            {
                GetDocList();
                swPropertyPage.Show();
            }
        }
        
        //get the open documents list
        public void GetDocList()
        {
            ModelDoc2 model;
            int DocNum = (int)iSwApp.GetDocumentCount();
            object[] models = (object[])iSwApp.GetDocuments();
            List<ModelDoc2> modelList = new List<ModelDoc2>();
            string tmpPath;
            
            rList.Clear();
            pList.Clear();

            for (int index = 0; index < DocNum; index++)
            {
                model = models[index] as ModelDoc2;
                tmpPath = model.GetPathName();
                
                if (tmpPath != "")
                {
                    modelList.Add(model);
                    
                    string fileName = (string)Path.GetFileNameWithoutExtension(tmpPath);

                    if (fileName.Contains("rm"))
                    {
                        rList.AddItems(fileName);
                    }
                    else
                    {
                        pList.AddItems(fileName);
                    }
                }
            }

            userAddin.modelList = modelList;
                                    
        }

    }
}
