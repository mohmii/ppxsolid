#include "stdafx.h"

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// User Dll includes
#include "ppxsolid.h"
#include "SwDocument.h"

void Cppxsolid::FileOpen()//製品モデルと被削材モデルの位置を合わせる
{
	
	CoInitialize(NULL);

	{
			
	//IModelDoc2 *pModel = NULL;
	//IAssemblyDoc *pAssem = NULL;	//新規作成アセンブリへのポインタ（操作が失敗ならばNULL）
	//IComponent2 *pComp1, * pComp2 = NULL;		//付加された構成部品（Component）へのポインタ

	CComPtr<IModelDoc2> pModel = NULL;
	CComPtr<IAssemblyDoc> pAssem = NULL;	//新規作成アセンブリへのポインタ（操作が失敗ならばNULL）
	CComPtr<IComponent2> pComp1, pComp2 = NULL;		//付加された構成部品（Component）へのポインタ

	double x, y, z, xform1[16], xform2[16];   //Componentの位置確認用マトリクス

	// new api type IMathTransform to keep the Transform2 return value
	CComPtr<IMathTransform> mTransform1, mTransform2 = NULL;
	
	VARIANT_BOOL retval;            //変換の設定に成功ならばTRUE
	VARIANT_BOOL condition = FALSE;

	HRESULT res = NULL;				//返り値（成功ならばS_OK）

	try{

	//INewAssemblyの変更後
	BSTR templateName;				//新規ドキュメントのテンプレート名のようなもの															
	
	res = iSwApp->GetDocumentTemplate( swDocASSEMBLY,NULL,NULL, NULL, NULL, &templateName);		//
	if( res != S_OK ) throw(0);																		//
	res = iSwApp->INewDocument2 ( templateName, swDwgPaperAsizeVertical, NULL, NULL, &pModel );	//
	if( res != S_OK ) throw(0);
	SysFreeString( templateName );
	

	res = pModel->SetTitle2 (auT("tempAssembly.SLDASM"), &retval );
	res = pModel->QueryInterface( IID_IAssemblyDoc, (LPVOID *)&pAssem );

	x = y = 0.0; 
	z = 0.0;

	//new api AddComponent5 for adding the two components
	res = pAssem->AddComponent5 ( auT("Workpiece.SLDPRT"),0, auT("WP"),true, auT("WP"), x, y, z, &pComp1 );
	res = pAssem->AddComponent5 ( auT("Product.SLDPRT"), 0, auT("FS"),true, auT("FS"), x, y, z, &pComp2 );

	if(pComp1 == NULL || pComp2 == NULL){
		AfxMessageBox(_T("\"Workpiece.SLDPRT\" or \"Product.SLDPRT\" did not exist\n"));
		throw(0);
	}

	//new api Transform2 to get the component matrix
	res = pComp1->get_Transform2(&mTransform1);
	res = mTransform1->get_IArrayData(xform1);
	
	res = pComp2->get_Transform2(&mTransform2);
	res = mTransform2->get_IArrayData(xform2);

	//xform2[9] = xform1[9];
	//xform2[10] = xform1[10];
	xform2[11] = xform1[11];
	
	//set back to mathtransform
	res = mTransform1->put_IArrayData(xform1);
	res = mTransform2->put_IArrayData(xform2);
	
	//set back to Icomponent2
	res = pComp1->put_Transform2(mTransform1);
	res = pComp2->put_Transform2(mTransform2);

	if(pModel) pModel.Release();
	pModel = NULL;
	res = iSwApp->get_IActiveDoc2( &pModel );

	res = pModel->ShowNamedView2  ( auT("*等角投影"), 7 );

	}catch(...){
	if(pAssem) pAssem.Release();
	if(pModel) pModel.Release();
	if(pComp1) pComp1.Release();
	if(pComp2) pComp2.Release();
	
	if(mTransform1) mTransform1.Release();
	if(mTransform2) mTransform2.Release();
	
//	fout.close();             //チェック用
	AfxMessageBox(_T("Exception Happened."));
	}

	if(pAssem) pAssem.Release();
	if(pModel) pModel.Release();
	if(pComp1) pComp1.Release();
	if(pComp2) pComp2.Release();
	
	if(mTransform1 != NULL ) mTransform1.Release();
	if(mTransform2 != NULL ) mTransform2.Release();
	
	} //end of CoUninitialize()

	CoUninitialize();
}																// End function
