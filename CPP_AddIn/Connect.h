// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "CPP_AddIn_i.h"
#include "ItemWrapper.h"
#define 	dispidEventItemLoad				((ULONG)(0xfba7))



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;
class CConnect;
static _ATL_FUNC_INFO ItemLoadInfo = { CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH} };

typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1> IDTImpl;
typedef public IDispEventSimpleImpl<1, CConnect, &__uuidof(ApplicationEvents_11)> ApplicationEventSink;

// CConnect

class ATL_NO_VTABLE CConnect : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<IConnect, &IID_IConnect, &LIBID_CPP_AddInLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<ApplicationEvents_11, &__uuidof(ApplicationEvents_11), &__uuidof(__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>,
	IDTImpl,
	ApplicationEventSink
{
public:
	CConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(106)


	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY(IConnect)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(ApplicationEvents_11)
	END_COM_MAP()



	BEGIN_SINK_MAP(CConnect)
		SINK_ENTRY_INFO(1, __uuidof(ApplicationEvents_11), dispidEventItemLoad, ItemLoad, &ItemLoadInfo)
	END_SINK_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

private:
	ItemWrapper * m_ItemWrapper = NULL;



	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		HRESULT hRes = S_OK;
		hRes = ApplicationEventSink::DispEventAdvise((IDispatch*)Application, &__uuidof(ApplicationEvents_11));
		if (FAILED(hRes))
		{
			ATLTRACE("Unable to listed for ApplicationEvents");
			return hRes;
		}
		return S_OK;
	}

	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return S_OK;
	}

// ApplicationEvents_11 Methods
public:
#pragma region 
	STDMETHOD(ItemSend)(LPDISPATCH Item, VARIANT_BOOL * Cancel)
	{
		return S_OK;
	}
	STDMETHOD(NewMail)()
	{
		return S_OK;
	}
	STDMETHOD(Reminder)(LPDISPATCH Item)
	{
		return S_OK;
	}
	STDMETHOD(OptionsPagesAdd)(LPDISPATCH Pages)
	{
		return S_OK;
	}
	STDMETHOD(Startup)()
	{
		return S_OK;
	}
	STDMETHOD(Quit)()
	{
		return S_OK;
	}
	STDMETHOD(AdvancedSearchComplete)(LPDISPATCH SearchObject)
	{
		return S_OK;
	}
	STDMETHOD(AdvancedSearchStopped)(LPDISPATCH SearchObject)
	{
		return S_OK;
	}
	STDMETHOD(MAPILogonComplete)()
	{
		return S_OK;
	}
	STDMETHOD_(void, NewMailEx)(BSTR EntryIDCollection)
	{
		return;
	}
	STDMETHOD(AttachmentContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Attachments)
	{
		return S_OK;
	}
	STDMETHOD_(void, FolderContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Folder)
	{
		return;
	}
	STDMETHOD_(void, StoreContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Store)
	{
		return;
	}
	STDMETHOD_(void, ShortcutContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Shortcut)
	{
		return;
	}
	STDMETHOD_(void, ViewContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH View)
	{
		return;
	}
	STDMETHOD_(void, ItemContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Selection)
	{
		return;
	}
	STDMETHOD_(void, ContextMenuClose)(long ContextMenu)
	{
		return;
	}

	STDMETHOD_(void, BeforeFolderSharingDialog)(LPDISPATCH FolderToShare, VARIANT_BOOL * Cancel)
	{
		return;
	}
#pragma endregion Events not relevant in this repro 

	STDMETHOD_(void, ItemLoad)(LPDISPATCH Item)
	{
		if (m_ItemWrapper)
			m_ItemWrapper->~ItemWrapper();
		ItemWrapper * m_ItemWrapper = new ItemWrapper(Item);
		return;
	}
};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
