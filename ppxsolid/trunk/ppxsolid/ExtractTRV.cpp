#include "stdafx.h"
#include <iostream>

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// User Dll includes
#include "ppxsolid.h"
#include "SwDocument.h"

void Cppxsolid::ExtractTRV()
{
	CoInitialize(NULL);

	try
	{

	//declare smart pointers
	CComPtr<IModelDoc2> pSwModel;
	CComPtr<IPartDoc> pPart;
	CComPtr<IDrawingDoc> pDraw;
	CComPtr<IAssemblyDoc> pAssy;
	IFeature * pProd;
	CComPtr<IModelDocExtension> pSwDocExt;
	CComPtr<IFeatureManager> pFeatMgr;
	CComPtr<IBody2> mainBody, otherBody;
			
	long longstatus = NULL;
	long longwarnings =  NULL;
	long lNumSelections = NULL;
	long nBodyHighIndex, nBodyCounter = -1;
	long nFaces[2] = {0,0}; 

	HRESULT hres = NOERROR;
	VARIANT_BOOL retVal = VARIANT_FALSE;
	VARIANT pBodyArr, pBodies;

	//set the BSTR for raw material name
	CComBSTR RmName(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\Workpiece - Copy.SLDPRT");
	CComBSTR RmConfig(L"Default");
	
	//open the raw material
	hres = iSwApp->OpenDoc6(RmName, 1, 0, RmConfig, &longstatus, &longwarnings, &pSwModel);

	//Empty the Model Document
	pSwModel = NULL;

	//set raw material as the active document
	hres = iSwApp->IActivateDoc3(RmName, true, &longstatus, &pSwModel);

	pPart = pSwModel;
	
	//set the BSTR for product name
	CComBSTR prodName(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\Product - Copy.SLDPRT");

	//insert the product
	hres = pPart->InsertPart2(prodName, 273, &pProd);

	//set the movement of product
	pSwModel->get_Extension(&pSwDocExt);

	//add the code here

	
	//get all the bodies within document
	hres = pPart->GetBodies2(0, true, &pBodyArr);

	SAFEARRAY*              psaBody = V_ARRAY(&pBodyArr);
    LPDISPATCH*             pBodyDispArray = NULL;

	hres = SafeArrayAccessData(psaBody, (void **) &pBodyDispArray);
	hres = SafeArrayGetUBound(psaBody, 1, &nBodyHighIndex);
	nBodyCounter = nBodyHighIndex + 1;

	for(int i = 0; i<nBodyCounter; i++)
	{
		CComQIPtr<IBody2> tmpBody;
		tmpBody = pBodyDispArray[i];
		hres = tmpBody->GetFaceCount(&nFaces[i]);
		ASSERT(nFaces[i]);					
	}

	//check the mainbody by its faces, the smallest number of faces seems to be the main body
	long i, j = NULL;	i=0; j=1;
	
	if (nFaces[0] < nFaces[1])
	{
		hres = SafeArrayGetElement(psaBody, &i, &mainBody);
		hres = SafeArrayGetElement(psaBody, &j, &otherBody);
	}
	else
	{
		hres = SafeArrayGetElement(psaBody, &j, &mainBody);
		hres = SafeArrayGetElement(psaBody, &i, &otherBody);
	}
		
	//substract the raw material with product
	hres = pSwModel->get_FeatureManager(&pFeatMgr);

	IBody2 ** otherBodies = new IBody2*[1];
	otherBodies[0] = otherBody;

	hres = pFeatMgr->IInsertCombineFeature(SWBODYCUT, mainBody, 1, otherBodies, &pProd);
	
	//set the name for the trv and save it
	CComBSTR trvComponent(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\trvComponent.SLDPRT");
	
	pSwDocExt->SaveAs(trvComponent, 0, 1, NULL, &longstatus, &longwarnings, &retVal);
	
	hres = SafeArrayUnaccessData(psaBody);
	hres = SafeArrayDestroy(psaBody);

	CoUninitialize();

	}

	catch (HRESULT hres)
	{
		CoUninitialize();
	}
			
}
