#pragma once

using namespace ATL;

class CViewHelper
{
public:
	CViewHelper(LPDISPATCH Application);
	~CViewHelper();
	HRESULT CViewHelper::CreateTableView(BSTR bstrViewName, BSTR * bstrFieldName, ULONG ulFieldCount);
private:
	CComPtr<_Application> m_Application;
};

