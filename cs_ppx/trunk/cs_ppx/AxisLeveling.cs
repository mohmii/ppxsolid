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
        //plane Calculator
        public void PlaneCalculatorZ()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            MathUtility swMathUtils = (MathUtility)SwApp.GetMathUtility();
            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            bool RegStatus = false;

            AssemblyDoc assyModel = null;
            Component2[] compName = null;
            List<AddedReferencePlane> refplaneX = new List<AddedReferencePlane>();
            List<AddedReferencePlane> refplaneY = new List<AddedReferencePlane>();
            List<AddedReferencePlane> refplaneZ = new List<AddedReferencePlane>();

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

                        //draw the real centroid
                        //Doc.SketchManager.Insert3DSketch(true);
                        centroid = (double[])bodyProperties.CenterOfMass;
                        SketchPoint skPoint1 = (SketchPoint)Doc.SketchManager.CreatePoint(centroid[0], centroid[1], centroid[2]);
                        Doc.SketchManager.Insert3DSketch(true);

                        //get the virtual centroid
                        //object boxVertices = compName[0].GetBox(false, false);
                        virtualCentroid = getMidPoint(MaxMinValue);

                        //draw the virtual centroid
                        //Doc.SketchManager.Insert3DSketch(true);
                        double[] tmpPoint = (double[])virtualCentroid;
                        SketchPoint skPoint2 = (SketchPoint)Doc.SketchManager.CreatePoint(tmpPoint[0], tmpPoint[1], tmpPoint[2]);
                        Doc.SketchManager.Insert3DSketch(true);

                        //get the part document of the raw material
                        ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                        PartDoc swPartDoc = (PartDoc)compModDoc;

                        //get the plane feature
                        boolStatus = getPlanes(swPartDoc, ref InitialRefPlanes);

                        if (boolStatus == true)
                        {
                            //get the plane normal
                            boolStatus = getPlaneNormal(swMathUtils, ref InitialRefPlanes);

                            if (boolStatus == true)
                            {
                                //set the distance
                                //boolStatus = setDistance(Doc, swMathUtils, ref InitialRefPlanes, tmpPoint);
                                boolStatus = setDistance(Doc, swMathUtils, ref InitialRefPlanes);

                                //find the centroid position refer to the normal direction
                                //boolStatus = findCPost(ref InitialRefPlanes, tmpPoint);
                                boolStatus = findCPost(ref InitialRefPlanes);

                                //find the outer reference plane
                                boolStatus = FindTheOuter(ref InitialRefPlanes);


                                List<int> removeId = new List<int>();

                                //store the plane feature, rank, and normal to planeList
                                //registerPlane(planeValue, planeNames, planeNormal, distance, ref removeId, swMathUtils);


                                RegStatus = registerPlane2(Doc,InitialRefPlanes, ref removeId, swMathUtils);
                                refplaneX.Clear();
                                refplaneY.Clear();
                                refplaneZ.Clear();

                                for (int i = 0; i < SelectedRefPlanes.Count ; i++)
                                {
                                    if (SelectedRefPlanes[i].NormalOrientation == "x-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "x-negative")
                                        refplaneX.Add(SelectedRefPlanes[i]);
                                    else if (SelectedRefPlanes[i].NormalOrientation == "y-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "y-negative")
                                        refplaneY.Add(SelectedRefPlanes[i]);
                                    else if (SelectedRefPlanes[i].NormalOrientation == "z-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "z-negative")
                                        refplaneZ.Add(SelectedRefPlanes[i]);

                                }

                                ProcessLog_TaskPaneHost.LogProcess("Calculate reference planes");
                                PPDetails_TaskPaneHost.LogProcess("Calculate reference planes");

                                if ((RegStatus == true) && (removeId.Count > 0))
                                {
                                    suppressFeature(Doc, assyModel, compName[0], InitialRefPlanes, removeId);

                                }

                            }
                        }
                        SelectedRefPlanes = new List<AddedReferencePlane>(refplaneZ);

                        for (int i = 0; i < SelectedRefPlanes.Count; i++)
                        {
                            //AddedReferencePlane tmpPlaneA = SelectedRefPlanes[i];
                            for (int j = 0; j < SelectedRefPlanes.Count; j++)
                            {
                                if (SelectedRefPlanes[i].name == SelectedRefPlanes[j].name && i != j)
                                {
                                    SelectedRefPlanes.RemoveAt(j);
                                    if (i > j) i--;
                                    j--;
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

        public void PlaneCalculatorXY()
        {
            ModelDoc2 Doc = (ModelDoc2)SwApp.ActiveDoc;
            MathUtility swMathUtils = (MathUtility)SwApp.GetMathUtility();
            int docType = (int)Doc.GetType();
            bool boolStatus = false;
            bool RegStatus = false;

            AssemblyDoc assyModel = null;
            Component2[] compName = null;
            List<AddedReferencePlane> refplaneX = new List<AddedReferencePlane>();
            List<AddedReferencePlane> refplaneY = new List<AddedReferencePlane>();
            List<AddedReferencePlane> refplaneZ = new List<AddedReferencePlane>();

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

                        //draw the real centroid
                        //Doc.SketchManager.Insert3DSketch(true);
                        centroid = (double[])bodyProperties.CenterOfMass;
                        SketchPoint skPoint1 = (SketchPoint)Doc.SketchManager.CreatePoint(centroid[0], centroid[1], centroid[2]);
                        Doc.SketchManager.Insert3DSketch(true);

                        //get the virtual centroid
                        //object boxVertices = compName[0].GetBox(false, false);
                        virtualCentroid = getMidPoint(MaxMinValue);

                        //draw the virtual centroid
                        //Doc.SketchManager.Insert3DSketch(true);
                        double[] tmpPoint = (double[])virtualCentroid;
                        SketchPoint skPoint2 = (SketchPoint)Doc.SketchManager.CreatePoint(tmpPoint[0], tmpPoint[1], tmpPoint[2]);
                        Doc.SketchManager.Insert3DSketch(true);

                        //get the part document of the raw material
                        ModelDoc2 compModDoc = (ModelDoc2)compName[0].GetModelDoc2();
                        PartDoc swPartDoc = (PartDoc)compModDoc;

                        //get the plane feature
                        boolStatus = getPlanes(swPartDoc, ref InitialRefPlanes);

                        if (boolStatus == true)
                        {
                            //get the plane normal
                            boolStatus = getPlaneNormal(swMathUtils, ref InitialRefPlanes);

                            if (boolStatus == true)
                            {
                                //set the distance
                                //boolStatus = setDistance(Doc, swMathUtils, ref InitialRefPlanes, tmpPoint);
                                boolStatus = setDistance(Doc, swMathUtils, ref InitialRefPlanes);

                                //find the centroid position refer to the normal direction
                                //boolStatus = findCPost(ref InitialRefPlanes, tmpPoint);
                                boolStatus = findCPost(ref InitialRefPlanes);

                                //find the outer reference plane
                                boolStatus = FindTheOuter(ref InitialRefPlanes);


                                List<int> removeId = new List<int>();

                                //store the plane feature, rank, and normal to planeList
                                //registerPlane(planeValue, planeNames, planeNormal, distance, ref removeId, swMathUtils);


                                RegStatus = registerPlane2(Doc, InitialRefPlanes, ref removeId, swMathUtils);
                                refplaneX.Clear();
                                refplaneY.Clear();
                                refplaneZ.Clear();

                                for (int i = 0; i < SelectedRefPlanes.Count; i++)
                                {
                                    if (SelectedRefPlanes[i].NormalOrientation == "x-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "x-negative")
                                        refplaneX.Add(SelectedRefPlanes[i]);
                                    else if (SelectedRefPlanes[i].NormalOrientation == "y-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "y-negative")
                                        refplaneY.Add(SelectedRefPlanes[i]);
                                    else if (SelectedRefPlanes[i].NormalOrientation == "z-plus" ||
                                        SelectedRefPlanes[i].NormalOrientation == "z-negative")
                                        refplaneZ.Add(SelectedRefPlanes[i]);

                                }

                                ProcessLog_TaskPaneHost.LogProcess("Calculate reference planes");
                                PPDetails_TaskPaneHost.LogProcess("Calculate reference planes");

                                if ((RegStatus == true) && (removeId.Count > 0))
                                {
                                    suppressFeature(Doc, assyModel, compName[0], InitialRefPlanes, removeId);

                                }

                            }
                        }
                        if (refplaneX.Count != 0 && CountSplitBody(Doc, refplaneX) < CountSplitBody(Doc, refplaneY))
                        {
                            SelectedRefPlanes = new List<AddedReferencePlane>(refplaneX);
                        }
                        else if (refplaneY.Count != 0 && CountSplitBody(Doc, refplaneX) > CountSplitBody(Doc, refplaneY))
                        {
                            SelectedRefPlanes = new List<AddedReferencePlane>(refplaneY);
                        }
                        else
                        {
                            SelectedRefPlanes = null;
                        }

                        for (int i = 0; i < SelectedRefPlanes.Count; i++)
                        {
                            //AddedReferencePlane tmpPlaneA = SelectedRefPlanes[i];
                            for (int j = 0; j < SelectedRefPlanes.Count; j++)
                            {
                                if (SelectedRefPlanes[i].name == SelectedRefPlanes[j].name && i != j)
                                {
                                    SelectedRefPlanes.RemoveAt(j);
                                    if (i > j) i--;
                                    j--;
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
        //OVERLOAD add all plane into one list
        public bool registerPlane2(ModelDoc2 Doc, List<AddedReferencePlane> ListOfPlanes, ref List<int> removeId, MathUtility mathUtils)
        {
            AssemblyDoc assyModel = (AssemblyDoc)Doc;
            AddedReferencePlane tmpSetPlane;
            bool retSim;

            //set the instance for the list that will keep the reference planes
            SelectedRefPlanes = new List<AddedReferencePlane>();　　　//表に出力される参照面
            planeList = new List<_planeProperties>();
            List<Feature> featureList = new List<Feature>();

            TraverseComponentFeatures(compName[0], ref featureList);
            bool boolStatus = compName[0].Select2(true, 0);
            assyModel.EditPart();

            for (int i = 0; i < ListOfPlanes.Count; i++)
            {
                tmpSetPlane = new AddedReferencePlane();
                tmpSetPlane = ListOfPlanes[i];
                retSim = false;

                if (SelectedRefPlanes.Count == 0)
                {
                    //オープンフェイス排除
                    Feature selectfeature = getSelectedPlane(tmpSetPlane, featureList);
                    bool bret = selectfeature.Select2(false, 0);
                    Array bodyarray = (Array)Doc.FeatureManager.PreSplitBody();
                    
                    if (bodyarray.Length > 1)
                    {
                        SelectedRefPlanes.Add(tmpSetPlane);
                    }
                    else
                    {
                        removeId.Add(i);
                    }
                    
                }
                else
                {
                    retSim = checkPlaneSim(SelectedRefPlanes, tmpSetPlane, mathUtils);

                    if (retSim == true)
                    {
                        removeId.Add(i);
                    }
                    else
                    {
                        Feature selectfeature = getSelectedPlane(tmpSetPlane,featureList);
                        selectfeature.Select2(false, 0);
                        Array bodyarray = (Array)Doc.FeatureManager.PreSplitBody();
                       
                        if (bodyarray.Length > 1)
                        {
                            SelectedRefPlanes.Add(tmpSetPlane);
                        }
                        else
                        {
                            removeId.Add(i);
                        }
                    }
                }
            }
           
            return true;
        }

        int CountSplitBody(ModelDoc2 Doc, List<AddedReferencePlane> RefPlaneList)
        {
            AssemblyDoc assyModel = (AssemblyDoc)Doc;
            FeatureManager featMgr = Doc.FeatureManager;
            List<Feature> featurelist = new List<Feature>();
            List<Feature> selectplanelist = new List<Feature>();
            int bodycount = 0;

            TraverseComponentFeatures(compName[0], ref featurelist);

            foreach (AddedReferencePlane tmpplane in RefPlaneList)
            {
                selectplanelist.Add(getSelectedPlane(tmpplane,featurelist));
                
            }

            bool bret = compName[0].Select2(false,0);
            assyModel.EditPart();

            foreach (Feature tmpfeature in selectplanelist)
            {
                bret = tmpfeature.Select2(true, 0);
            }

            Array bodyarray = (Array)featMgr.PreSplitBody();
            bodycount = bodyarray.Length;
            Doc.ClearSelection2(true);
            return bodycount;

        }



    }
}


