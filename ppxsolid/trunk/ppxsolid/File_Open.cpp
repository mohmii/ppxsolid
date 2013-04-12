#include "stdafx.h"

// MFC includes
#include <tchar.h>		// MultiByte & Unicode support	

// User Dll includes
//#include "UserApp.h"
//#include "FPPS.h"
//#include "GlobalData.h"
//#include "amapp.h"

#include "ppxsolid.h"

/*	//チェック用
#include <fstream>
using namespace std;    ////std::　の省略が可能に
*/

void File_Open()//製品モデルと被削材モデルの位置を合わせる
{
//	ofstream      fout("File_Open.txt");		//チェック用

//	ofstream      fout("C:\\Documents and Settings\\SolidWorks\\My Documents\\Uko\\Yooshan\\Yooshan_sld\\File_Open.txt");

	CComPtr<ISldWorks> pSldWorks = NULL;
	IModelDoc2 *pModel = NULL;
	IAssemblyDoc *pAssem = NULL;	//新規作成アセンブリへのポインタ（操作が失敗ならばNULL）
	IComponent2 *pComp1= NULL, * pComp2 = NULL, * pComp3 =NULL, *pComp4 = NULL;		//付加された構成部品（Component）へのポインタ
	double x, y, z;					//アセンブリの構成部品の位置
//	double W_tlevel = 60.00*0.001;	//特定の条件なので要注意
//	double P_tlevel = 45.00*0.001;	//特定の条件なので要注意

	double xform1[16],xform2[16],xform3[16],xform4[16];   //Componentの位置確認用マトリクス
     VARIANT_BOOL retval;            //変換の設定に成功ならばTRUE
     VARIANT_BOOL condition = FALSE;

	HRESULT res = NULL;				//返り値（成功ならばS_OK）

	try{

	pSldWorks = GetSldWorksPtr()/* UserApp->getSWApp()*/;
	
//	res = pSldWorks->INewAssembly( &pAssem );  //INewAssemblyは廃止された
//	res = pAssem->QueryInterface( IID_IModelDoc2, (LPVOID *)&pModel );	
	//INewAssemblyの変更後
	BSTR templateName;				//新規ドキュメントのテンプレート名のようなもの															
	res = pSldWorks-> GetDocumentTemplate( swDocASSEMBLY,NULL,NULL, NULL, NULL, &templateName);		//
	if( res != S_OK ) throw(0);																		//
	res = pSldWorks->INewDocument2 ( templateName, swDwgPaperAsizeVertical, NULL, NULL, &pModel );	//
	if( res != S_OK ) throw(0);
	SysFreeString( templateName );

	res = pModel->SetTitle2 ( auT("tempAssembly.SLDASM"), &retval );
	res = pModel->QueryInterface( IID_IAssemblyDoc, (LPVOID *)&pAssem );

	x = y = 0.0; 
	z = 0.0;
//	x = y =  0.00000;
//	z = W_tlevel/2;
//	res = pAssem->IAddComponent2 ( auT("Workpiece.SLDPRT"), x, y, z, &pComp1 );
	res = pAssem->IAddComponent3 ( auT("Workpiece.SLDPRT"), x, y, z, &pComp1 );


//	if(res == S_OK) fout<<"Workpiece"<<"\n";		//チェック用

//	z = P_tlevel/2;
//	res = pAssem->IAddComponent2 ( auT("Product.SLDPRT"), x, y, z, &pComp2 );
	res = pAssem->IAddComponent3 ( auT("Product.SLDPRT"), x, y, z, &pComp2 );

	if(pComp1 == NULL || pComp2 == NULL){
		AfxMessageBox(_T("\"Workpiece.SLDPRT\" or \"Product.SLDPRT\" did not exist\n"));
		throw(0);
	}

/////////////////////////////////////// z軸位置補正（06/10/16） UKO

	res = pComp1->IGetXform(xform1);
/*//チェック用
	if(res == S_OK) {
		fout<<"xform1(Workpiece)"<<"\n";
		for(int i=0;i<13;i++){
			fout<<xform1[i]<<"\t";

			if(i==2||i==5||i==8||i==11||i==12){
				fout<<"\n";
			}
		}
	}
*/

	res = pComp2->IGetXform(xform2);
/*//チェック用
	if(res == S_OK) {
		fout<<"xform2(Produc)"<<"\n";
		for(int j=0;j<13;j++){
			fout<<xform2[j]<<"\t";

			if(j==2||j==5||j==8||j==11||j==12){
				fout<<"\n";
			}
		}
	}	
*/



	xform2[9] = xform1[9];
	xform2[10] = xform1[10];
	xform2[11] = xform1[11];

	res = pComp1->ISetXform ( xform1, &retval  );
	res = pComp2->ISetXform ( xform2, &retval  );

///////////////////////////////////////////

///////////////////////////////////////////抜き勾配考慮時（2006/11/27）UKO
//	res = pAssem->IAddComponent2 ( auT("nonDraftProduct.SLDPRT"), x, y, z, &pComp3 );
	res = pAssem->IAddComponent3 ( auT("nonDraftProduct.SLDPRT"), x, y, z, &pComp3 );
	if(pComp3 != NULL) {
		res = pComp3->IGetXform(xform3);
		xform3[9] = xform1[9];
		xform3[10] = xform1[10];
		xform3[11] = xform1[11];
		res = pComp3->ISetXform ( xform3, &retval  );
	}
////////////////////////////////////////////////////////////////

///////////////////////////////////////////Fillet考慮時///////
	res = pAssem->IAddComponent3 ( auT("noFeatureProduct.SLDPRT"), x, y, z, &pComp4 );
	if(pComp4 != NULL) {
		res = pComp4->IGetXform(xform4);
		xform4[9] = xform1[9];
		xform4[10] = xform1[10];
		xform4[11] = xform1[11];
		res = pComp4->ISetXform ( xform4, &retval  );
	}
//////////////////////////////////////////////////////////////////
	if(pModel) pModel->Release();
	pModel = NULL;
	res = UserApp->getSWApp()->get_IActiveDoc2 ( &pModel );

	res = pModel->ShowNamedView2  ( auT("*等角投影"), 7 );

	}catch(...){
	if(pAssem) pAssem->Release();
	if(pModel) pModel->Release();
	if(pComp1) pComp1->Release();
	if(pComp2) pComp2->Release();
	if(pComp3) pComp3->Release();

//	fout.close();             //チェック用
	AfxMessageBox(_T("Exception Happened."));
	}

	if(pAssem) pAssem->Release();
	if(pModel) pModel->Release();
	if(pComp1) pComp1->Release();
	if(pComp2) pComp2->Release();
	if(pComp3) pComp3->Release();
}																// End function
