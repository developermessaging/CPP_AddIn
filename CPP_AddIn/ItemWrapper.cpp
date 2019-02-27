#include "stdafx.h"
#include "ItemWrapper.h"




CItemWrapper::CItemWrapper(IDispatch * iDispItem)
{
	m_Item = iDispItem;
	HRESULT hRes = S_OK;
	hRes = m_Item->QueryInterface(__uuidof(_MailItem), (void **)&m_spMailItem);
	if SUCCEEDED(hRes)
	{
		hRes = m_spMailItem->get_MessageClass(&bstrMessageClass);
		if SUCCEEDED(hRes)
		{
			// We only wanna advise on items with the message class below 
			//because we don't wanna cancel changes to any other items
			if (0 == wcscmp(bstrMessageClass, L"IPM.Note.EnterpriseVaultShortcut"))
			{
				AdviseMailItem();
			}
		}
		else
		{
			m_Item->Release();
			hrError = hRes;
		}
	}
	else
	{
		m_Item->Release();
		hrError = hRes;
	}
}

CItemWrapper::~CItemWrapper()
{
	UnadviseMailItem();
	if (m_Item)
		m_Item->Release();
	::SysFreeString(bstrMessageClass);

}

void CItemWrapper::AdviseMailItem()
{
	if (m_Item)
		if (SUCCEEDED(ItemEventSink::DispEventAdvise(m_Item, &__uuidof(ItemEvents_10))))
		{
			bListeningForEvents = TRUE;
		}
}

void CItemWrapper::UnadviseMailItem()
{
	if (m_Item && bListeningForEvents)
		ItemEventSink::DispEventUnadvise(m_Item, &__uuidof(ItemEvents_10));
}

STDMETHODIMP CItemWrapper::Write(VARIANT_BOOL* isCancel)
{
	MessageBoxW(NULL, L"Cancelling write operation!", L"Native Addin", MB_OK);
	*isCancel = VARIANT_TRUE;
	return S_OK;
}
