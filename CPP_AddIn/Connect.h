// Connect.h : Declaration of the CConnect

#pragma once
#include "resource.h"       // main symbols
#include "CPP_AddIn_i.h"
#include "ItemWrapper.h"
#include "ViewHelper.h"
#define 	dispidEventItemLoad				((ULONG)(0xfba7))

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;
class CConnect;
static _ATL_FUNC_INFO ItemLoadInfo = { CC_STDCALL,VT_EMPTY,1,{VT_DISPATCH} };


typedef public IDispEventSimpleImpl<1, CConnect, &__uuidof(ApplicationEvents_11)> ApplicationEventSink;

typedef enum {
	NotSupported,
	Explorer,
	Contact,
	Task,
	Appointment,
	Post
} RibbonType;

// CConnect

class ATL_NO_VTABLE CConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CConnect, &CLSID_Connect>,
	public IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>,
	public IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ -1>,
	public IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback), &LIBID_CPP_AddInLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDispatchImpl<ApplicationEvents_11, &__uuidof(ApplicationEvents_11), &__uuidof(__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>,
	ApplicationEventSink
{
public:
	CConnect()
	{
	}

	DECLARE_REGISTRY_RESOURCEID(106)
	DECLARE_NOT_AGGREGATABLE(CConnect)

	BEGIN_COM_MAP(CConnect)
		COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
		COM_INTERFACE_ENTRY(ApplicationEvents_11)
		COM_INTERFACE_ENTRY(IRibbonExtensibility)
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
	CItemWrapper * m_ItemWrapper = NULL;

public:
	CComPtr<_Application> m_Application;

#pragma region _IDTExtensibility2 Methods
	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom);
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom);
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom);
#pragma endregion

#pragma region ApplicationEvents_11 Methods
	// ApplicationEvents_11 Methods
public:
	STDMETHOD(ItemSend)(LPDISPATCH Item, VARIANT_BOOL * Cancel);
	STDMETHOD(NewMail)();
	STDMETHOD(Reminder)(LPDISPATCH Item);
	STDMETHOD(OptionsPagesAdd)(LPDISPATCH Pages);
	STDMETHOD(Startup)();
	STDMETHOD(Quit)();
	STDMETHOD(AdvancedSearchComplete)(LPDISPATCH SearchObject);
	STDMETHOD(AdvancedSearchStopped)(LPDISPATCH SearchObject);
	STDMETHOD(MAPILogonComplete)();
	STDMETHOD_(void, NewMailEx)(BSTR EntryIDCollection);
	STDMETHOD(AttachmentContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Attachments);
	STDMETHOD_(void, FolderContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Folder);
	STDMETHOD_(void, StoreContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Store);
	STDMETHOD_(void, ShortcutContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Shortcut);
	STDMETHOD_(void, ViewContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH View);
	STDMETHOD_(void, ItemContextMenuDisplay)(LPDISPATCH CommandBar, LPDISPATCH Selection);
	STDMETHOD_(void, ContextMenuClose)(long ContextMenu);
	STDMETHOD_(void, BeforeFolderSharingDialog)(LPDISPATCH FolderToShare, VARIANT_BOOL * Cancel);
	STDMETHOD_(void, ItemLoad)(LPDISPATCH Item);
#pragma endregion

#pragma region IRibbonExtensibility Methods
	// IRibbonExtensibility Methods
public:
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);

#pragma endregion

#pragma region Ribbon Extensibility helpers
	HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes);
	BSTR GetXMLResource(int nId);
	SAFEARRAY* GetOFSResource(int nId);
	RibbonType CheckRibbonID(BSTR bstrRibbonID);

#pragma endregion

#pragma region IRibbonCallback Methods
	// IRibbonCallback Implementation:
	STDMETHOD(OnCreateCustomView)(IRibbonControl *pRibbonControl);

#pragma endregion

};

OBJECT_ENTRY_AUTO(__uuidof(Connect), CConnect)
