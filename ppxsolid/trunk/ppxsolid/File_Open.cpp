#include "stdafx.h"

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// User Dll includes
#include "ppxsolid.h"
#include "SwDocument.h"

void Cppxsolid::FileOpen()//���i���f���Ɣ��ރ��f���̈ʒu�����킹��
{
	
	CoInitialize(NULL);

	{

	CComPtr<ISldWorks> pSldWorks = NULL;

	//Cppxsolid *userAddin = NULL;
	IModelDoc2 *pModel = NULL;
	IAssemblyDoc *pAssem = NULL;	//�V�K�쐬�A�Z���u���ւ̃|�C���^�i���삪���s�Ȃ��NULL�j
	IComponent2 *pComp1= NULL, * pComp2 = NULL, * pComp3 =NULL, *pComp4 = NULL;		//�t�����ꂽ�\�����i�iComponent�j�ւ̃|�C���^
	double x, y, z;					//�A�Z���u���̍\�����i�̈ʒu

	double xform1[16],xform2[16],xform3[16],xform4[16];   //Component�̈ʒu�m�F�p�}�g���N�X

	// new api type
	IMathTransform * mTransform1, *mTransform2, *mTransform3, *mTransform4 = NULL;

     VARIANT_BOOL retval;            //�ϊ��̐ݒ�ɐ����Ȃ��TRUE
     VARIANT_BOOL condition = FALSE;

	HRESULT res = NULL;				//�Ԃ�l�i�����Ȃ��S_OK�j

	try{

	//INewAssembly�̕ύX��
	BSTR templateName;				//�V�K�h�L�������g�̃e���v���[�g���̂悤�Ȃ���															
	res = /*pSldWorks*/iSwApp-> GetDocumentTemplate( swDocASSEMBLY,NULL,NULL, NULL, NULL, &templateName);		//
	if( res != S_OK ) throw(0);																		//
	res = /*pSldWorks*/iSwApp->INewDocument2 ( templateName, swDwgPaperAsizeVertical, NULL, NULL, &pModel );	//
	if( res != S_OK ) throw(0);
	SysFreeString( templateName );
	

	res = pModel->SetTitle2 (auT("tempAssembly.SLDASM"), &retval );
	res = pModel->QueryInterface( IID_IAssemblyDoc, (LPVOID *)&pAssem );

	x = y = 0.0; 
	z = 0.0;

	//new api
	res = pAssem->AddComponent5 ( auT("Workpiece.SLDPRT"),0, auT("WP"),true, auT("WP"), x, y, z, &pComp1 );

	//new api
	res = pAssem->AddComponent5 ( auT("Product.SLDPRT"), 0, auT("FS"),true, auT("FS"), x, y, z, &pComp2 );

	if(pComp1 == NULL || pComp2 == NULL){
		AfxMessageBox(_T("\"Workpiece.SLDPRT\" or \"Product.SLDPRT\" did not exist\n"));
		throw(0);
	}

	//new api
	res = pComp1->get_Transform2(&mTransform1);
	res = mTransform1->get_IArrayData(xform1);
	
	res = pComp2->get_Transform2(&mTransform2);
	res = mTransform2->get_IArrayData(xform2);

	xform2[9] = xform1[9];
	xform2[10] = xform1[10];
	xform2[11] = xform1[11];
	
	//set back to mathtransform
	res = mTransform1->put_IArrayData(xform1);
	res = mTransform2->put_IArrayData(xform2);
	
	//set back to Icomponent2
	res = pComp1->put_Transform2(mTransform1);
	res = pComp2->put_Transform2(mTransform2);

	if(pModel) pModel->Release();
	pModel = NULL;
	res = /*UserApp->getSWApp()->*/iSwApp->get_IActiveDoc2( &pModel );

	res = pModel->ShowNamedView2  ( auT("*���p���e"), 7 );

	}catch(...){
	if(pAssem) pAssem->Release();
	if(pModel) pModel->Release();
	if(pComp1) pComp1->Release();
	if(pComp2) pComp2->Release();
	
	if(mTransform1) mTransform1->Release();
	if(mTransform2) mTransform2->Release();
	
//	fout.close();             //�`�F�b�N�p
	AfxMessageBox(_T("Exception Happened."));
	}

	if(pAssem) pAssem->Release();
	if(pModel) pModel->Release();
	if(pComp1) pComp1->Release();
	if(pComp2) pComp2->Release();
	
	if(mTransform1 != NULL ) mTransform1->Release();
	if(mTransform2 != NULL ) mTransform2->Release();
	
	} //end of CoUninitialize()

	CoUninitialize();
}																// End function
