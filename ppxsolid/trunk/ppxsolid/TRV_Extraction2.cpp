// Programmed and commented by Keiichi NAKAMOTO (2003/06/30)
#include "stdafx.h"

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// SolidWorks includes
#include <atlbase.h>

// User Dll includes
//#include "UserApp.h"
//#include "FPPS.h"

#include "math.h"
#include "ppxsolid.h"

/*
#include <fstream>
using namespace std;    ////std::�@�̏ȗ����\��
*/

void TRV_Extraction2()//���i���f���Ɣ��ރ��f������s�q�u�𒊏o
{
//	ofstream      fout("TRV_Extraction2.txt");
	CComQIPtr<ISldWorks> pSldWorks;	
	CComQIPtr<IModelDoc2> pModdc;
	CComQIPtr<IConfiguration> pConfiguration;	//�R���t�B�M�����[�V�����ւ̃|�C���^
	CComQIPtr<IComponent2> pRootComponent;		//���[�g�R���|�[�l���g�ւ̃|�C���^
	IComponent2 * pCompList[3] = {NULL};			//�R���|�[�l���g�ւ̃|�C���^�ւ̃|�C���^
	IBody2 * pBody1 = NULL;					//�{�f�B�ւ̃|�C���^
	IBody2 * pBody2 = NULL;					//�{�f�B�ւ̃|�C���^
	IBody2 * ptmpBody1 = NULL;				//�e���|�����{�f�B�ւ̃|�C���^
	IBody2 * ptmpBody2 = NULL;				//�e���|�����{�f�B�ւ̃|�C���^
//	LPFAULTENTITY pFaultEntity1 = NULL;
//	LPFAULTENTITY pFaultEntity2 = NULL;
	IEnumBodies2 * pBodyList = NULL;			//�{�f�B�̗񋓈ꗗ
	IBody2 * ptmpBody_attrib1 = NULL;		//�e���|�����{�f�B�ւ̃|�C���^
	IBody2 * ptmpBody_attrib2 = NULL;		//�e���|�����{�f�B�ւ̃|�C���^
	IModelDoc2 * pNewModdc = NULL;			//Boolean���ModelDoc�ւ̃|�C���^
	IPartDoc * pNewPart = NULL;				//PartDoc�ւ̃|�C���^
	IBody2 * pNewBody = NULL;				//Boolean��̃{�f�B�ւ̃|�C���^
	IFeature * pFeature = NULL;				//�t�B�[�`���ւ̃|�C���^
	IBody2 * tmppNewBody = NULL;				//Boolean��̃e���|�����E�{�f�B�ւ̃|�C���^
	VARIANT_BOOL accessGained = NULL;

	long type;								//Doc�̃I�v�V�����^�C�v
	int nChildren = 0;						//�q�R���|�[�l���g��
	char nameCh[16];
	int num_W;
	int num_P;
	int num_NDP;
	int num_NFP;
	int num_Comp1;
	int num_Comp2;
	double vntXform1[16];	//�}�ʃr���[�̈ʒu�ƃX�P�[���̔z��ւ̃|�C���^
	double vntXform2[16];	//�}�ʃr���[�̈ʒu�ƃX�P�[���̔z��ւ̃|�C���^


	int flag_W = 0;
	int flag_P = 0;
	int flag_NDP = 1;
	int flag_NFP = 1;
	int flag_Comp = 0;
	int flag = 0;

	HRESULT res = NULL;						//�Ԃ�l�i�����Ȃ��S_OK�j
	VARIANT_BOOL bres;						//�����Ȃ��TRUE
	VARIANT_BOOL retval;
	long lngError;							//�G���[�E�C���W�P�[�^

	try
	{
		pSldWorks = GetSldWorksPtr() /*UserApp->getSWApp()*/;
		//�h�L�������g�̃C���^�[�t�F�C�X�|�C���^���擾
		if( S_OK != pSldWorks->get_IActiveDoc2( &pModdc )){
			AfxMessageBox(_T("get_IActiveDoc ERROR\n"));
			throw(0);
		}
		if( pModdc == NULL ) throw(0);

		pModdc->GetType(&type);
		if( type =! swDocASSEMBLY){
			AfxMessageBox(_T("ActiveDoc is not AssemblyDoc ERROR\n"));
			throw(0);
		}

		//�h�L�������g�̃R���t�B�M�����[�V�����|�C���^���擾
		res = pModdc->IGetActiveConfiguration ( &pConfiguration );
		if( S_OK != res ){
			AfxMessageBox(_T("GetActiveConfiguration ERROR\n"));
			throw(0);
		}
		if( pConfiguration == NULL ) throw(0);

		//�h�L�������g�̃��[�g�R���|�[�l���g�|�C���^���擾
		res = pConfiguration->IGetRootComponent2 ( &pRootComponent );
		if( S_OK != res ){
			AfxMessageBox(_T("GetRootComponent ERROR\n"));
			throw(0);
		}
		if( pRootComponent == NULL ) throw(0);
//		pConfiguration->Release();

		//���[�g�R���|�[�l���g�̎q�R���|�[�l���g�����擾
		if( S_OK != pRootComponent->IGetChildrenCount ( &nChildren )){
			AfxMessageBox(_T("GetChildrenCount ERROR\n"));
			throw(0);
		}
		if(nChildren != 2 && nChildren != 3){
			AfxMessageBox(_T("This Doc is Fail\n"));
			throw(0);
		}

		//�q�R���|�[�l���g���擾
		if( S_OK != pRootComponent->IGetChildren(pCompList)){
			AfxMessageBox(_T("GetChildren ERROR1\n"));
			throw(0);
		}

		for(int i=0;i<nChildren;i++){
			BSTR NameBS = NULL;
			res = pCompList[i]->get_Name( &NameBS );
			if( S_OK != res ){
				AfxMessageBox(_T("get_Name ERROR\n"));
	//			fout<<"Name"<<"\t"<<nameCh<<"\n";
				throw(0);
			}
			int x = WideCharToMultiByte(	//BSTR�^����Char�^�ւ̕ϊ�
					CP_ACP,					// �R�[�h�y�[�W
					0,						// �������x�ƃ}�b�s���O���@�����肷��t���O�i����j
					NameBS,					// ���C�h������̃A�h���X
					16,						// ���C�h������̕�����
					nameCh,					// �V������������󂯎��o�b�t�@�̃A�h���X
					16,						// �V������������󂯎��o�b�t�@�̃T�C�Y
					NULL,					// �}�b�v�ł��Ȃ������̊���l�̃A�h���X
					NULL);					// ����̕������g�����Ƃ��ɃZ�b�g����t���O�̃A�h���X
	//			fout<<"Name"<<"\t"<<nameCh<<"\n";
			if(0 == strncmp(nameCh, "Workpiece", 9)){
				num_W = i;
				flag_W = 1;
			}
			else if(0 == strncmp(nameCh, "Product", 7)){
				num_P = i;
				flag_P = 1;
			}
			else if(0 == strncmp(nameCh, "nonDraftProduct", 15)){
				num_NDP = i;
				flag_NDP = 10;
			}

			else if(0 == strncmp(nameCh, "noFeatureProduct", 16)){
				num_NFP = i;
				flag_NFP = -1;
			}
			SysFreeString(NameBS);
		}

		flag_Comp = flag_W*flag_P*flag_NDP*flag_NFP;

		if(flag_Comp != 1 && flag_Comp != 10 && flag_Comp != -1 )
		{
			AfxMessageBox(_T("This Doc is Fail\n"));
			throw(0);
		}
	
//		fout <<"�R���|�[�l���g�擾����"<<endl;

		do
		{
			switch(flag_Comp)
			{
			case 10:
				num_Comp1 = num_NDP;
				num_Comp2 = num_P;
				break;
			case 1:
				num_Comp1 = num_W;
				num_Comp2 = num_P;
				break;
			case 100:
				num_Comp1 = num_W;
				num_Comp2 = num_NDP;
				break;
						
			case -1:
				num_Comp1 = num_NFP;
				num_Comp2 = num_P;
				break;
										
			case -10:
				num_Comp1 = num_W;
				num_Comp2 = num_NFP;
				break;
			}

		//�R���|�[�l���g�̃{�f�B�I�u�W�F�N�g���擾
			res = pCompList[num_Comp1]->IGetBody ( &pBody1 );
			res = pCompList[num_Comp2]->IGetBody ( &pBody2 );
			if( S_OK != res ){
				AfxMessageBox(_T("GetBody ERROR\n"));
				throw(0);
			}

			res = pBody1->ICopy ( &ptmpBody1 );
			res = pBody2->ICopy ( &ptmpBody2 );

		//�e���|�����{�f�B���A�Z���u���̍��W���ɍ��킹�z�u
			res = pCompList[num_Comp1]->IGetXform ( vntXform1 );
			res = pCompList[num_Comp2]->IGetXform( vntXform2 );
			vntXform1[9]=0.0; //�덷�J�b�g�i���j
			vntXform2[9]=0.0; //
			vntXform1[10]=0.0; //
			vntXform2[10]=0.0; //
			vntXform1[11]=0.0;
			vntXform2[11]=0.0;
			res = ptmpBody1->ISetXform ( vntXform1, &retval  );
			res = ptmpBody2->ISetXform ( vntXform2, &retval  );
/*
			ptmpBody1->IsTemporaryBody(&bres);
			ptmpBody2->IsTemporaryBody(&bres);
			ptmpBody1->get_Check3(&pFaultEntity1);
			ptmpBody2->get_Check3(&pFaultEntity2);
*/
			//Boolean CUT
			res =  ptmpBody1->IOperations2 ( SWBODYCUT, ptmpBody2, &lngError, &pBodyList );
	
			if( S_OK != res){
				AfxMessageBox(_T("Operations  ERROR\nCheck your choice\n"));
				throw(0);
			}
			if(ptmpBody1) ptmpBody1->Release();
			if(ptmpBody2) ptmpBody2->Release();

			flag = UserApp->Object_Check2( pBody2 );
			if(flag!=9) res = pBody1->ICopy ( &ptmpBody_attrib1);
			flag = UserApp->Object_Check2( pBody2 );
			if(flag!=9) res = pBody1->ICopy ( &ptmpBody_attrib2);

			if( flag == 0 )	UserApp->DefInitialize_FaceAccuracy( );		 //Attrib0�̏�����
			else if( flag == 3 ) UserApp->DefInitialize_EdgeAccuracy( ); //Attrib3�̏�����

//			fout <<"�V�����p�[�g�Ƃ��ďo��"<<endl;
			pBodyList->Reset();

		//�V�����p�[�g�Ƃ��ďo��
	//		res = UserApp->getSWApp()->INewPart ( &pNewPart );//sw2003
	//		res = pNewPart->QueryInterface( IID_IModelDoc2, (LPVOID *)&pModdc2 );
			BSTR templateName = NULL;						//�V�K�h�L�������g�̃e���v���[�g���̂悤�Ȃ���
			res = pSldWorks->GetDocumentTemplate( swDocPART,NULL,NULL, NULL, NULL, &templateName);			//
			res = pSldWorks->INewDocument2 ( templateName, swDwgPaperAsizeVertical, NULL, NULL, &pNewModdc );	//
			SysFreeString(templateName);
			if( res != S_OK ) throw(0);

			res = pNewModdc->QueryInterface(IID_IPartDoc, (LPVOID *)&pNewPart);

			res = pBodyList->Next ( 1, &pNewBody, NULL );
			do
			{
				res = pNewPart->ICreateFeatureFromBody4 ( pNewBody, false, 
														swCreateFeatureBodySimplify, &pFeature);
				if(pFeature == NULL){
					AfxMessageBox(_T("pFeature = NULL"));
					throw(0);
				}

				if(flag!=9)
				{
					//�t�B�[�`���[����{�f�B��
					res = pFeature->IGetBody2 ( &tmppNewBody );
					//�t�B�[�`���[�������{�f�B�ɑ������R�s�[
					if(ptmpBody_attrib1) UserApp->Copy_toTRV_fromProduct2( ptmpBody_attrib1, tmppNewBody );
					if(ptmpBody_attrib2) UserApp->Copy_toTRV_fromProduct2( ptmpBody_attrib2, tmppNewBody );
				}

				if(pNewBody) pNewBody->Release();
				pNewBody = NULL;
				if(pFeature) pFeature->Release();
				pFeature = NULL;
				res = pBodyList->Next ( 1, &pNewBody, NULL );
			}while(res == S_OK);

			res = pNewModdc->ShowNamedView2  ( auT("*���p���e"), 7 );//sw2003

			if( flag == 0 )
			{
				//Product��Face�Ƃ͐ڂ��Ă��Ȃ�TRV��Face�ɐV���������𐶐�
				UserApp->New_FaceAttrib( pNewPart );
			}
			else if( flag == 3 )
			{
				//Product��Edge�Ƃ͐ڂ��Ă��Ȃ�TRV��Edge�ɐV���������𐶐�
				UserApp->New_EdgeAttrib( pNewPart );
			}

			res = pNewModdc->EditRebuild3 ( &bres );
			res = pNewPart->ImportDiagnosis(TRUE,TRUE,TRUE,NULL,&lngError);

			if(flag_Comp == 10){
				res = pNewModdc->SetTitle2 ( auT("TRV of Draft.SLDPRT"), &retval );
			}
			
			else if(flag_Comp == -1){
				res = pNewModdc->SetTitle2 ( auT("TRV of Feature.SLDPRT"), &retval );
				res = pNewModdc->ShowNamedView2  ( auT("*���p���e"), 7 );

				res = pNewModdc->SaveAs4 ( auT("TRV of Feature.SLDPRT"), swSaveAsCurrentVersion, swSaveAsOptions_Silent , NULL, NULL, &accessGained );
			}
			
			else 
			{
				res = pNewModdc->SetTitle2 ( auT("TRV.SLDPRT"), &retval );
			}

			if(pBodyList) pBodyList->Release();
			if(pNewPart) pNewPart->Release();
			if(pNewModdc) pNewModdc->Release();
			if(ptmpBody_attrib1) ptmpBody_attrib1->Release();
			if(ptmpBody_attrib2) ptmpBody_attrib2->Release();
			if(pBody1) pBody1->Release();
			if(pBody2) pBody2->Release();

			pBodyList = NULL;
			pNewPart = NULL;
			pNewModdc = NULL;
			ptmpBody_attrib1 = NULL;
			ptmpBody_attrib2 = NULL;
			pBody1 = NULL;
			pBody2 = NULL;
/*
			if(pFaultEntity1) pFaultEntity1->Release();
			if(pFaultEntity2) pFaultEntity2->Release();
*/
			flag_Comp *= 10;

		}while(flag_Comp == 100 || flag_Comp == -10);

	}catch(...)
	{
		if(ptmpBody1) ptmpBody1->Release();
		if(ptmpBody2) ptmpBody2->Release();
		if(ptmpBody_attrib1) ptmpBody_attrib1->Release();
		if(ptmpBody_attrib2) ptmpBody_attrib2->Release();
		if(pBodyList) pBodyList->Release();
		if(pNewBody) pNewBody->Release();
		if(pNewPart) pNewPart->Release();
		if(pNewModdc) pNewModdc->Release();
		if(pBody1) pBody1->Release();
		if(pBody2) pBody2->Release();
		if(pCompList[0]) pCompList[0]->Release();
		if(pCompList[1]) pCompList[1]->Release();
		if(pCompList[2]) pCompList[2]->Release();
		/*
		if(pFaultEntity1) pFaultEntity1->Release();
		if(pFaultEntity2) pFaultEntity2->Release();
		*/
		AfxMessageBox(_T("Exception Happened.\nin TRV Extraction"));
		throw(0);
	}

	if(pCompList[0]) pCompList[0]->Release();
	if(pCompList[1]) pCompList[1]->Release();
	if(pCompList[2]) pCompList[2]->Release();
}																// End function
//--------------------------------------------------------------------------------------