#pragma once


// CppxsolidDialog dialog

class CppxsolidDialog : public CDialog
{
	DECLARE_DYNAMIC(CppxsolidDialog)

public:
	CppxsolidDialog(CWnd* pParent = NULL);   // standard constructor
	virtual ~CppxsolidDialog();
		
// Dialog Data
	enum { IDD = IDD_CppxsolidDialog };

protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

	DECLARE_MESSAGE_MAP()
};
