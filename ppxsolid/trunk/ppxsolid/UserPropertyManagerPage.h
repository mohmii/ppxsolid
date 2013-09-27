// UserPropertyManagerPage.h : Declaration of the CUserPropertyManagerPage

#pragma once
#include "ppxsolid_i.h"
#include "resource.h"       // main symbols
#include <comsvcs.h>
class Cppxsolid;
class CPMPageHandler;

//Control IDs
#define GROUP1		0
#define GROUP2		1
#define	TEXT1		2
#define CHECK1		3
#define OPTION1		4
#define OPTION2		5
#define OPTION3		6	
#define LIST1		7
#define SELECTION1	8
#define NUM1		9
#define COMBO1		10
#define SELECTPRODUCT	11
#define SELECTRM	12

//PPX control IDS
#define PPXMODULE 77

// CUserPropertyManagerPage

class ATL_NO_VTABLE CUserPropertyManagerPage :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CUserPropertyManagerPage, &CLSID_UserPropertyManagerPage>,
	public IDispatchImpl<IUserPropertyManagerPage, &IID_IUserPropertyManagerPage, &LIBID_ppxsolidLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:

private:
	Cppxsolid *userAddin;
	CComPtr<ISldWorks> iSwApp;
	CComPtr<IPropertyManagerPage2> swPropertyPage;
	CComObject<CPMPageHandler> *handler;

protected: 
	//Generic Control
	IPropertyManagerPageControl *control;

	//Groups
	IPropertyManagerPageGroup *group1;
	IPropertyManagerPageGroup *group2;

	//ppx module
	IPropertyManagerPageGroup *ppxmodule;

	//In Group1
	IPropertyManagerPageTextbox		*text1;
	IPropertyManagerPageCheckbox	*check1;
	IPropertyManagerPageOption		*option1;
	IPropertyManagerPageOption		*option2;
	IPropertyManagerPageOption		*option3;
	IPropertyManagerPageListbox		*list1;

	//In Group2
	IPropertyManagerPageSelectionbox *selection1;
	IPropertyManagerPageNumberbox	 *num1;
	IPropertyManagerPageCombobox	 *combo1;
public:
	CUserPropertyManagerPage()
	{
	}

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_USERPROPERTYMANAGERPAGE)

DECLARE_NOT_AGGREGATABLE(CUserPropertyManagerPage)

BEGIN_COM_MAP(CUserPropertyManagerPage)
	COM_INTERFACE_ENTRY(IUserPropertyManagerPage)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()




// IUserPropertyManagerPage
public:
	void Init(Cppxsolid *app);
	void CreatePropertyManagerPage();
	void Destroy();
	void Show();
	void AddControls();
};

OBJECT_ENTRY_AUTO(__uuidof(UserPropertyManagerPage), CUserPropertyManagerPage)
