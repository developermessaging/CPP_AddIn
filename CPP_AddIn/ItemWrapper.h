#pragma once

#include "stdafx.h"
#include <vector>
#include <algorithm>

using namespace ATL;
class ItemWrapper;

const DWORD dispidEventWrite = 0xf002;
__declspec(selectany) _ATL_FUNC_INFO WriteInfo = { CC_STDCALL, VT_EMPTY, 1,{ VT_BOOL | VT_BYREF } };
typedef public IDispEventSimpleImpl<1, ItemWrapper, &__uuidof(ItemEvents_10)> ItemEventSink;

class ItemWrapper : ItemEventSink
{
public:
	ItemWrapper(IDispatch * iDispItem);
	~ItemWrapper();

	STDMETHOD(Write)(VARIANT_BOOL * Cancel);
	BEGIN_SINK_MAP(ItemWrapper)
		SINK_ENTRY_INFO(1, __uuidof(ItemEvents_10),
			dispidEventWrite,
			Write,
			&WriteInfo)
	END_SINK_MAP()


private:
	IDispatch * m_Item = NULL;
	BOOL bListeningForEvents = FALSE;
	BSTR bstrMessageClass = L"";
	CComPtr<_MailItem> m_spMailItem;
	HRESULT hrError = S_OK;
	void AdviseMailItem();
	void UnadviseMailItem();
};

