// Connect.cpp : Implementation of CConnect

#include "stdafx.h"
#include "Connect.h"
#include <atlsafe.h>
// CConnect

#pragma region _IDTExtensibility2 Methods
STDMETHODIMP CConnect::OnConnection(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
{

	HRESULT hRes = S_OK;
	hRes = Application->QueryInterface(__uuidof(_Application), (void **)&m_Application);
	if (FAILED(hRes))
	{
		ATLTRACE(L"Unable to get Outlook application pointer");
	}
	hRes = ApplicationEventSink::DispEventAdvise((IDispatch*)Application, &__uuidof(ApplicationEvents_11));
	if (FAILED(hRes))
	{
		ATLTRACE(L"Unable to listen for ApplicationEvents");
		return hRes;
	}
	return S_OK;
}

STDMETHODIMP CConnect::OnDisconnection(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnAddInsUpdate(SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnStartupComplete(SAFEARRAY * * custom)
{
	return S_OK;
}

STDMETHODIMP CConnect::OnBeginShutdown(SAFEARRAY * * custom)
{
	return S_OK;
}
#pragma endregion

#pragma region ApplicationEvents_11 Methods
// ApplicationEvents_11 Methods
STDMETHODIMP CConnect::ItemSend(LPDISPATCH Item, VARIANT_BOOL * Cancel)
{
	return S_OK;
}

STDMETHODIMP CConnect::NewMail()
{
	return S_OK;
}

STDMETHODIMP CConnect::Reminder(LPDISPATCH Item)
{
	return S_OK;
}

STDMETHODIMP CConnect::OptionsPagesAdd(LPDISPATCH Pages)
{
	return S_OK;
}

STDMETHODIMP CConnect::Startup()
{
	return S_OK;
}

STDMETHODIMP CConnect::Quit()
{
	return S_OK;
}

STDMETHODIMP CConnect::AdvancedSearchComplete(LPDISPATCH SearchObject)
{
	return S_OK;
}

STDMETHODIMP CConnect::AdvancedSearchStopped(LPDISPATCH SearchObject)
{
	return S_OK;
}

STDMETHODIMP CConnect::MAPILogonComplete()
{
	return S_OK;
}

VOID CConnect::NewMailEx(BSTR EntryIDCollection)
{
	return;
}

STDMETHODIMP CConnect::AttachmentContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Attachments)
{
	return S_OK;
}

VOID CConnect::FolderContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Folder)
{
	return;
}

VOID CConnect::StoreContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Store)
{
	return;
}

VOID CConnect::ShortcutContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Shortcut)
{
	return;
}
VOID CConnect::ViewContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH View)
{
	return;
}
VOID CConnect::ItemContextMenuDisplay(LPDISPATCH CommandBar, LPDISPATCH Selection)
{
	return;
}
VOID CConnect::ContextMenuClose(long ContextMenu)
{
	return;
}

VOID CConnect::BeforeFolderSharingDialog(LPDISPATCH FolderToShare, VARIANT_BOOL * Cancel)
{
	return;
}

VOID CConnect::ItemLoad(LPDISPATCH Item)
{
	if (m_ItemWrapper)
		m_ItemWrapper->~CItemWrapper();
	CItemWrapper * m_ItemWrapper = new CItemWrapper(Item);
	return;
}
#pragma endregion

#pragma region IRibbonExtensibility Methods
STDMETHODIMP CConnect::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
	if (!RibbonXml)
		return E_POINTER;
	RibbonType ribbonType = NotSupported;
	ribbonType = CheckRibbonID(RibbonID);
	switch (ribbonType)
	{
	case RibbonType::Explorer:
		*RibbonXml = GetXMLResource(IDR_XML2);
		break;
	}
	return S_OK;
}

#pragma endregion

#pragma region Ribbon Extensibility helpers

RibbonType CConnect::CheckRibbonID(BSTR bstrRibbonID)
{
	int cchRibbonId = SysStringLen(bstrRibbonID);
	if (0 == wcscmp(L"Microsoft.Outlook.Explorer", bstrRibbonID))
	{
		return RibbonType::Explorer;
	}
	else if (0 == wcscmp(L"Microsoft.Outlook.Contact", bstrRibbonID))
	{
		return RibbonType::Contact;
	}
	else if (0 == wcscmp(L"Microsoft.Outlook.Task", bstrRibbonID))
	{
		return RibbonType::Task;
	}
	else if (0 == wcscmp(L"Microsoft.Outlook.Appointment", bstrRibbonID))
	{
		return RibbonType::Appointment;
	}
	else if ((0 == wcscmp(L"Microsoft.Outlook.Post.Read", bstrRibbonID)) ||
			(0 == wcscmp(L"Microsoft.Outlook.Post.Compose", bstrRibbonID)))
	{
		return RibbonType::Post;
	}
	else
	{
		return RibbonType::NotSupported;
	}
}

/*!-----------------------------------------------------------------------
	Retrieves a resource from the module.

	According to MSDN there is no need to clean up the resources
	because they will be released when the module is unloaded.
-----------------------------------------------------------------------!*/
HRESULT CConnect::HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
	if (!lpType || !ppvResourceData || !pdwSizeInBytes)
		return E_POINTER;

	HMODULE hModule = _AtlBaseModule.GetModuleInstance();

	if (!hModule)
		return E_UNEXPECTED;

	HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);

	if (!hRsrc)
		return HRESULT_FROM_WIN32(GetLastError());

	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);

	if (!hGlobal)
		return HRESULT_FROM_WIN32(GetLastError());

	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);

	return S_OK;
}

/*------------------------------------------------------------------------------
------------------------------------------------------------------------------*/
BSTR CConnect::GetXMLResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;

	if (FAILED(HrGetResource(nId, TEXT("XML"), &pResourceData, &dwSizeInBytes)))
		return NULL;

	// Assumes that the data is not stored in Unicode.
	return A2WBSTR(reinterpret_cast<LPCSTR>(pResourceData), dwSizeInBytes);
}

/*------------------------------------------------------------------------------
------------------------------------------------------------------------------*/
SAFEARRAY* CConnect::GetOFSResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;
	CComSafeArray<BYTE> psa;

	if (FAILED(HrGetResource(nId, TEXT("OFS"), &pResourceData, &dwSizeInBytes)))
		return NULL;

	if (FAILED(psa.Add(dwSizeInBytes, reinterpret_cast<BYTE*>(pResourceData))))
		return NULL;

	return psa.Detach();
}

#pragma endregion

#pragma region IRibbonCallback Methods
STDMETHODIMP CConnect::OnCreateCustomView(IRibbonControl *pRibbonControl)
{
	HRESULT hRes = S_OK;
	BSTR bViewName = ::SysAllocString(L"CustomView");
	BSTR * bFieldName = new BSTR[4];
	bFieldName[0] = ::SysAllocString(L"Subject");
	bFieldName[1] = ::SysAllocString(L"To");
	bFieldName[2] = ::SysAllocString(L"From");
	bFieldName[3] = ::SysAllocString(L"http://schemas.microsoft.com/mapi/proptag/0x001A001F");
	CViewHelper * pViewHelper = new CViewHelper(m_Application);
	EC_HRES_MSG(pViewHelper->CreateTableView(bViewName, bFieldName, 4), "calling CreateTableView");
Error:
	return hRes;
}

#pragma endregion

