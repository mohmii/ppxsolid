// UserPropertyManagerPage.cpp : Implementation of CUserPropertyManagerPage

#include "stdafx.h"
#include "UserPropertyManagerPage.h"
#include "PMPageHandler.h"
#include "ppxsolid.h"

// CUserPropertyManagerPage

void CUserPropertyManagerPage::Init(Cppxsolid *app)
{
	userAddin = app;
	iSwApp = userAddin->GetSldWorksPtr();
	CreatePropertyManagerPage();
}

//Create the handler and property page, and add the controls
void CUserPropertyManagerPage::CreatePropertyManagerPage()
{
	CComObject<CPMPageHandler>::CreateInstance(&handler);
	CComBSTR title;
	title.LoadString(IDS_PMP_TITLE);
	long errors = 0;
	long options = swPropertyManagerOptions_OkayButton | swPropertyManagerOptions_CancelButton;
	iSwApp->ICreatePropertyManagerPage(title,options,handler,&errors, &swPropertyPage);
	if (swPropertyPage != NULL)
		AddControls();
}

void CUserPropertyManagerPage::Destroy()
{
	swPropertyPage->Close(VARIANT_FALSE); //same as pressing cancel
	swPropertyPage = NULL;
}

//Display the property page
void CUserPropertyManagerPage::Show()
{
	long rval;
	swPropertyPage->Show(&rval);
}

//Adds the two control groups, and adds the controls to each group
void CUserPropertyManagerPage::AddControls()
{
	//VARS
	CComBSTR caption;
	long options = -1;
	CComBSTR tip;

	//Add the groups
	caption.LoadString(IDS_PMP_GROUP1_TITLE);
	options = swGroupBoxOptions_Visible | swGroupBoxOptions_Expanded;

	swPropertyPage->IAddGroupBox(GROUP1, caption, options, &group1);

	options = swGroupBoxOptions_Checkbox | swGroupBoxOptions_Visible;
	swPropertyPage->IAddGroupBox(GROUP2, caption, options, &group2);

	//Add the controls to group1
	//Text1
	options = swControlOptions_Visible | swControlOptions_Enabled;
	caption.LoadString(IDS_PMP_TEXT1_CAPTION);
	group1->IAddControl(TEXT1, swControlType_Textbox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageTextbox), (void**)&text1);

	//Check1
	caption.LoadString(IDS_PMP_CHECK1_CAPTION);
	tip.LoadString(IDS_PMP_CHECK1_TIP);
	group1->IAddControl(CHECK1, swControlType_Checkbox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageCheckbox), (void**)&check1);

	//Option1
	caption.LoadString(IDS_PMP_OPTION1_CAPTION);
	tip.LoadString(IDS_PMP_OPTION1_TIP);
	group1->IAddControl(OPTION1, swControlType_Option, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageOption), (void**)&option1);

	//Option2
	caption.LoadString(IDS_PMP_OPTION2_CAPTION);
	tip.LoadString(IDS_PMP_OPTION2_TIP);
	group1->IAddControl(OPTION2, swControlType_Option, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageOption), (void**)&option2);
	option2->put_Checked(VARIANT_TRUE);

	//Option3
	caption.LoadString(IDS_PMP_OPTION3_CAPTION);
	tip.LoadString(IDS_PMP_OPTION3_TIP);
	group1->IAddControl(OPTION3, swControlType_Option, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageOption), (void**)&option3);

	//List1
	caption.LoadString(IDS_PMP_LIST1_CAPTION);
	tip.LoadString(IDS_PMP_LIST1_TIP);
	group1->IAddControl(LIST1, swControlType_Listbox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageListbox), (void**)&list1);
	if (list1 != NULL)
	{
		list1->put_Height(50);
		_TCHAR* items[] = { _T("One Fish"), _T("Two Fish"), _T("Red Fish"), _T("Blue Fish") };
		list1->IAddItems(4, (BSTR*)items);
	}


	//Add the controls to group2
	//Selection1
	caption.LoadString(IDS_PMP_SELECTION1_CAPTION);
	tip.LoadString(IDS_PMP_SELECTION1_TIP);
	group2->IAddControl(SELECTION1, swControlType_Selectionbox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageSelectionbox), (void**)&selection1);
	if (selection1 != NULL)
	{
		long filterArray[] = { swSelFACES, swSelVERTICES };
		selection1->ISetSelectionFilters(2, filterArray);
		selection1->put_Height(50);
		selection1->put_SingleEntityOnly(VARIANT_TRUE);
	}

	//Num1
	caption.LoadString(IDS_PMP_NUM1_CAPTION);
	tip.LoadString(IDS_PMP_NUM1_TIP);
	group2->IAddControl(NUM1, swControlType_Numberbox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageNumberbox), (void**)&num1);
	num1->SetRange(swNumberBox_UnitlessInteger, 0, 100, 1, VARIANT_TRUE);
	num1->put_Value(1);

	//Combo1
	caption.LoadString(IDS_PMP_COMBO1_CAPTION);
	tip.LoadString(IDS_PMP_COMBO1_TIP);
	group2->IAddControl(COMBO1, swControlType_Combobox, caption, swControlAlign_LeftEdge, options, tip, &control);
	control->QueryInterface(__uuidof(IPropertyManagerPageCombobox), (void**)&combo1);
	if (combo1 != NULL)
	{
		_TCHAR* items[] = { _T("One Fish"), _T("Two Fish"), _T("Red Fish"), _T("Blue Fish") };
		combo1->IAddItems(4, (BSTR*)items);
	}
}

