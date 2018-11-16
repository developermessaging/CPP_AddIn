// AddInConnect.h : Declaration of the CAddInConnect

#pragma once
#include "resource.h"       // main symbols



#include "CPPComAddIn_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;
typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), /* wMajor = */ 1>
IDTImpl;

// CAddInConnect

class ATL_NO_VTABLE CAddInConnect :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CAddInConnect, &CLSID_AddInConnect>,
	public IDispatchImpl<IAddInConnect, &IID_IAddInConnect, &LIBID_CPPComAddInLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IDTImpl

{
public:
	CAddInConnect()
	{
	}
	
	DECLARE_REGISTRY_RESOURCEID(IDR_ADDINCONNECT)


	BEGIN_COM_MAP(CAddInConnect)
		COM_INTERFACE_ENTRY(IAddInConnect)
		COM_INTERFACE_ENTRY2(IDispatch, _IDTExtensibility2)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:




	// _IDTExtensibility2 Methods
public:
	STDMETHOD(OnConnection)(LPDISPATCH Application, ext_ConnectMode ConnectMode, LPDISPATCH AddInInst, SAFEARRAY * * custom)
	{
		MessageBoxW(NULL, L"OnConnection", L"CPPComAddIn", MB_OK);
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
};

OBJECT_ENTRY_AUTO(__uuidof(AddInConnect), CAddInConnect)
