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
    public partial class SwAddin : ISwAddin
    {
        public static bool SplitandDelete(int MPIndex)
        {
            if (MachiningPlanList[MPIndex].ViewName != null)
            {
                //iSwApp.CloseDoc(MachiningPlanList[MPIndex].ViewName);
                return false;
            }

            if (MainView == "") { return false; }

            ModelDoc2 Doc = (ModelDoc2)SwApp.ActivateDoc(MainView);

            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            Feature SelectedFeature = null;
            int index = 0;
            int SequenceNUM = 1;
            int SplitCounter = 0;
            int Errors = 0;
            int Warnings = 0;

            List<int> DeleteThisBody = null;
            List<AddedReferencePlane> ListOfRemovalPlanes = new List<AddedReferencePlane>();
            List<Feature> ListOfRemovalFeature = null;
            List<RemovedBody> PreviousRemoval = new List<RemovedBody>();

            List<MachiningProcess> ProcessCollection = new List<MachiningProcess>();
            String CompPathName = compName[0].GetPathName();
            ModelDoc2 CompDocumentModel = (ModelDoc2)SwApp.OpenDoc6(CompPathName, (int)swDocumentTypes_e.swDocPART,
                (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref Errors, ref Warnings); //load the document

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

                    ModelDoc2 TmpCompDoc = compName[0].GetModelDoc2();
                    

                    TraverseComponentFeatures(compName[0], ref PlaneFeatures);
                    List<string> PathList = null;

                    //create the plan directory
                    string PlanDirectory = Path.GetDirectoryName(Doc.GetPathName()) + "\\Plan" + (MPIndex + 1).ToString();
                    System.IO.Directory.CreateDirectory(PlanDirectory);

                    MachiningPlan thisMP = MachiningPlanList[MPIndex];
                    List<MachiningProcess> ProcessList = thisMP.MachiningProceses;
                    SelectionMgr SelMgr = Doc.SelectionManager;
                    List<RefPlane> RefplaneList = new List<RefPlane>();
                    List<Feature> FeatureList = new List<Feature>();

                    TraverseComponentFeatures(compName[0], ref FeatureList);

                    Body2 restbody = null;
                    foreach(MachiningProcess selectprocess in ProcessList)
                    {
                        Feature selectfeature = getSelectedPlane(selectprocess.MachiningReference, FeatureList);
                        bool bret = selectfeature.Select2(false, 0);
                        Array bodyarray = (Array)Doc.FeatureManager.PreSplitBody();

                        //スプリットに成功した場合
                        if (bodyarray.Length > 1)
                        {
                            Body2[] bodyCandidate = new Body2[bodyarray.Length];  
                            string[] Bodynames = new string[bodyarray.Length];  
                            Vertex[] BodyOrigins = new Vertex[bodyarray.Length];　
                            for (int j = 0; j < bodyarray.Length; j++)
                            {
                                bodyCandidate[j] = bodyarray.GetValue(j) as Body2;
                                Bodynames[j] = null;
                                BodyOrigins[j] = null;
                            }
                            Feature SplitFeature = Doc.FeatureManager.PostSplitBody(bodyCandidate.ToArray(), false, BodyOrigins, Bodynames);
                            Doc.ClearSelection2(true);

                            for (int i = 0; i < bodyCandidate.Length; i++)
                            {
                                Body2 tmpbody = bodyCandidate[i];
                                double[] massproperty = (double[])tmpbody.GetMassProperties(1);
                                double[] centroid = new double[3];
                                for (int j = 0; j < 3; j++)
                                {
                                    centroid[j] = massproperty[j];
                                }
                                if (checkBodyLocation(tmpbody, centroid, selectprocess.MachiningReference.CorrespondFeature, selectprocess.MachiningReference.PlaneNormal) == true)
                                {
                                    if (bodyCandidate.Length == 2)
                                    {
                                        if (i == 0)
                                        {
                                            restbody = bodyCandidate[1];
                                        }
                                        else if (i == 1)
                                        {
                                            restbody = bodyCandidate[0];
                                        }
                                        string bodyname = restbody.Name + "@" + Path.GetFileNameWithoutExtension(CompDocumentModel.GetPathName()) + "-1@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());
                                        bret = Doc.Extension.SelectByID2(bodyname, "SOLIDBODY", 0, 0, 0, false, 0, null, 0);
                                       
                                    }
                                    string selectbodyname = tmpbody.Name + "@" + Path.GetFileNameWithoutExtension(CompDocumentModel.GetPathName()) + "-1@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());
                                    //bret = tmpbody.Select2(true, tmpdata);
                                    bret = Doc.Extension.SelectByID2(selectbodyname, "SOLIDBODY", 0, 0, 0, false, 0, null, 0);
                                    int selectcount = SelMgr.GetSelectedObjectCount2(-1);



                                    //ModelDoc2 TmpFeat = (ModelDoc2)deletefeature;

                                    Feature deletefeature = (Feature)Doc.FeatureManager.InsertDeleteBody2(false);
                                    Doc.ClearSelection2(true);
                                    

                                    break;
                                }

                            }
                        }
                        else if (bodyarray.Length == 1)
                        {
                            Body2 tmpbody = (Body2)bodyarray.GetValue(0);
                            if (CheckTopologicalData(Doc, tmpbody) == true)
                            {
                                string selectbodyname = tmpbody.Name + "@" + Path.GetFileNameWithoutExtension(CompDocumentModel.GetPathName()) + "-1@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());
                                //bret = tmpbody.Select2(true, tmpdata);
                                bret = Doc.Extension.SelectByID2(selectbodyname, "SOLIDBODY", 0, 0, 0, false, 0, null, 0);

                                //ModelDoc2 TmpFeat = (ModelDoc2)deletefeature;

                                Feature deletefeature = (Feature)Doc.FeatureManager.InsertDeleteBody2(false);
                                Doc.ClearSelection2(true);
                            }
                            else
                            {
                                restbody = tmpbody;
                                string selectbodyname = restbody.Name + "@" + Path.GetFileNameWithoutExtension(CompDocumentModel.GetPathName()) + "-1@" + Path.GetFileNameWithoutExtension(Doc.GetPathName());
                                bret = Doc.Extension.SelectByID2(selectbodyname, "SOLIDBODY", 0, 0, 0, false, 0, null, 0);
                            }
                                
 
                        }

                    }
                    


                    

                }
            }

            return true;
        }

    }
}
