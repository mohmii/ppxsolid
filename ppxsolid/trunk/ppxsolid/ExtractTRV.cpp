#include "stdafx.h"
#include <iostream>

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// User Dll includes
#include "ppxsolid.h"
#include "SwDocument.h"

void separateBodies(CComPtr<IPartDoc> pPartDoc, tagVARIANT pBodyArr, IBody2 ** mainBody,IBody2 ** otherBody);
void mateComponent(IBody2 * mainBody, IBody2 * otherBody);
void getTransformation(IBody2 * mainBody, IBody2 * otherBody, IMathTransform ** matrixTransform);
void getBoxCorners(VARIANT bodyBox, double * corners);

void Cppxsolid::ExtractTRV()
{
	CoInitialize(NULL);

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
	
	HRESULT hr = NOERROR;
	VARIANT_BOOL retVal = VARIANT_FALSE;
	VARIANT pBodyArr;

	try
	{
		//set the BSTR for raw material name
		CComBSTR RmName(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\Workpiece - Copy.SLDPRT");
		CComBSTR RmConfig(L"Default");
	
		//open the raw material
		hr = iSwApp->OpenDoc6(RmName, 1, 0, RmConfig, &longstatus, &longwarnings, &pSwModel);

		//Empty the Model Document
		pSwModel = NULL;

		//set raw material as the active document
		hr = iSwApp->IActivateDoc3(RmName, true, &longstatus, &pSwModel);

		pPart = pSwModel;
	
		//set the BSTR for product name
		CComBSTR prodName(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\Product - Copy.SLDPRT");

		//insert the product
		hr = pPart->InsertPart2(prodName, 273, &pProd);
				
		//get all the bodies within document
		hr = pPart->GetBodies2(0, true, &pBodyArr);

		//separete the bodies
		separateBodies(pPart, pBodyArr, &mainBody, &otherBody);
		
		//set the movement of product
		pSwModel->get_Extension(&pSwDocExt);

		//set the name for the trv and save it
		CComBSTR beforeCombine(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\beforeCombine.SLDPRT");
	
		pSwDocExt->SaveAs(beforeCombine, 0, 1, NULL, &longstatus, &longwarnings, &retVal);
		
		CComPtr<IMathTransform> matrixTransform;

		getTransformation(mainBody, otherBody, &matrixTransform);
		
		//add the code here
		mateComponent(mainBody, otherBody);
															
		//substract the raw material with product
		hr = pSwModel->get_FeatureManager(&pFeatMgr);

		IBody2 ** otherBodies = new IBody2*[1];
		otherBody.CopyTo(&otherBodies[0]);

		hr = pFeatMgr->IInsertCombineFeature(SWBODYCUT, mainBody, 1, otherBodies, &pProd);	
		
		//set the name for the trv and save it
		CComBSTR afterCombine(L"C:\\Users\\iis\\Documents\\Research\\solidworks\\illustrations\\afterCombine.SLDPRT");
	
		pSwDocExt->SaveAs(afterCombine, 0, 1, NULL, &longstatus, &longwarnings, &retVal);
				
		CoUninitialize();

	}

	catch (_com_error e)
	{
		CoUninitialize();
		//MessageBox(NULL, e.ErrorMessage(), NULL, MB_OK | MB_SETFOREGROUND);
		//exit(EXIT_FAILURE);
	}
			
}

//separate main and other body
void separateBodies(CComPtr<IPartDoc> pPartDoc, tagVARIANT pBodyArr, IBody2 ** mainBody, IBody2 ** otherBody)
{
	long nBodyHighIndex, nBodyCounter = -1;
	long nFaces[2] = {0,0};

	HRESULT hres = NOERROR;


	SAFEARRAY*              psaBody = V_ARRAY(&pBodyArr);
    LPDISPATCH*             pBodyDispArray = NULL;
	try
	{
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
			hres = SafeArrayGetElement(psaBody, &i, &*mainBody);
			hres = SafeArrayGetElement(psaBody, &j, &*otherBody);
		}
			else
		{
			hres = SafeArrayGetElement(psaBody, &j, &*mainBody);
			hres = SafeArrayGetElement(psaBody, &i, &*otherBody);
		}

		hres = SafeArrayUnaccessData(psaBody);
		hres = SafeArrayDestroy(psaBody);
	}

	catch (_com_error e)
	{
		CoUninitialize();
		hres = SafeArrayUnaccessData(psaBody);
		hres = SafeArrayDestroy(psaBody);

		MessageBox(NULL, e.ErrorMessage(), NULL, MB_OK | MB_SETFOREGROUND);
		
		exit(EXIT_FAILURE);
	}		
}

//mate main and other body into one solid
void mateComponent(IBody2 * mainBody, IBody2 * otherBody)
{
	IMathTransform * mTransform1 = NULL;
	//CComPtr<IMathTransform>  mTransform1, mTransform2;
	//double x, y, z, xform1[16], xform2[16];

	HRESULT hr = NOERROR;
	VARIANT_BOOL boolVal = VARIANT_FALSE;
	
	IDispatch *pBody;
		
	try
	{
		//Get Matrix
		mainBody->QueryInterface(__uuidof(IDispatch), (LPVOID*)&pBody);
		hr = mainBody->GetCoincidenceTransform(pBody, &mTransform1, &boolVal);
		
	}

	catch (_com_error e)
	{
		CoUninitialize();
		MessageBox(NULL, e.ErrorMessage(), NULL, MB_OK | MB_SETFOREGROUND);
		exit(EXIT_FAILURE);
	}
	
}

void getTransformation(IBody2 * mainBody, IBody2 * otherBody, IMathTransform ** matrixTransform)
{
	CoInitialize(NULL);

	{
		HRESULT hres = NOERROR;
		VARIANT bodyBox1, bodyBox2;
		
		double corners[6]={};
		
		try
		{
				
		hres = mainBody->IGetBodyBox(corners);
		
		//getBoxCorners(bodyBox1, &corners1);

		if (hres == S_OK) hres = otherBody->IGetBodyBox(corners);
		else throw;
				
		//getBoxCorners(bodyBox2, &corners2);
		}

		catch (_com_error e)
		{
			/*delete corners1;
			delete corners2;*/

			CoUninitialize();
			MessageBox(NULL, e.ErrorMessage(), NULL, MB_OK | MB_SETFOREGROUND);
			exit(EXIT_FAILURE);
		}
		/*delete corners1;
		delete corners2;
*/
	}

	CoUninitialize();

}

//get the bounding box corner
void getBoxCorners(_variant_t bodyBox, double ** corners = new double*[6])
{
	SAFEARRAY * psaCorner = NULL;
	LPDISPATCH * pCornerArray = NULL;
	HRESULT hres = NOERROR;
	long nHighIndex, nCounter = -1;

	psaCorner = V_ARRAY(&bodyBox);
	hres = SafeArrayAccessData(psaCorner, (void**) pCornerArray);
	
	hres = SafeArrayGetUBound(psaCorner, 1, &nHighIndex);
	nCounter = nHighIndex + 1;

	for(int i = 0; i<nCounter; i++)
	{
		//corners[i] = pCornerArray[i];		
	}

	hres = SafeArrayUnaccessData(psaCorner);
	hres = SafeArrayDestroy(psaCorner);
}