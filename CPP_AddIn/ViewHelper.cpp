#include "stdafx.h"
#include "ViewHelper.h"
#include "Connect.h"
#include <OleAuto.h>

CViewHelper::CViewHelper(LPDISPATCH Application)
{
	HRESULT hRes = S_OK;
	hRes = Application->QueryInterface(__uuidof(_Application), (void **)&m_Application);
	if (FAILED(hRes))
	{
		ATLTRACE(L"Unable to get Outlook application pointer");
	}
}

CViewHelper::~CViewHelper()
{
}


HRESULT CViewHelper::CreateTableView(BSTR bstrViewName, BSTR * bstrFieldName, ULONG ulFieldCount)
{
	HRESULT hRes = S_OK;
	CComPtr<Outlook::_Explorer> pExplorer;
	CComPtr<Outlook::MAPIFolder> pFolder;
	CComPtr<Outlook::_Views> pViews;

	if (!m_Application)
	{
		return E_FAIL;
	}
	// Getting a pointer to the current Explorer object
	EC_HRES_MSG(m_Application->ActiveExplorer(&pExplorer), "getting the active explorer");
	if (pExplorer)
	{
		// Getting a pointer to the current MAPIFolder object selected in the Explorer
		EC_HRES_MSG(pExplorer->get_CurrentFolder(&pFolder), "getting the current folder");
		if (pFolder)
		{
			// Getting a pointer to the views collection on the folder
			EC_HRES_MSG(pFolder->get_Views(&pViews), "getting the views collection");
			if (pViews)
			{
				long uCount = 0;
				// Getting the view count
				EC_HRES_MSG(pViews->get_Count(&uCount), "getting the view count");
				BOOL fExists = FALSE;
				if (0 < uCount)
				{
					
					VARIANT varIndex;
					varIndex.vt = VT_INT;

					// looping through the views to make sure a view doesn't already exist with that name
					for (unsigned int i = 1; i <= uCount; i++)
					{
						varIndex.lVal = i;
						CComPtr<Outlook::View> pView;
						EC_HRES_MSG(pViews->Item(varIndex, &pView), "walking views collection");
						if (pView)
						{
							BSTR bViewAtIndexName;
							EC_HRES_MSG(pView->get_Name(&bViewAtIndexName), "getting view name");
							if (0 == wcscmp(bViewAtIndexName, bstrViewName))
							{
								fExists = TRUE;
								break;
							}
						}

					}
				}
				if (fExists)
				{
					// bail if a view with that name exists
					ATLTRACE(L"View object already exists");
					MessageBoxW(NULL, L"A view with the same name already exists!", L"CPP_AddIn", MB_OK);
					return E_FAIL;
				}
				else
				{
					// otherwise create the view
					CComPtr<Outlook::View> pNewView;
					CComPtr<Outlook::_TableView> pTableView;
					CComPtr<Outlook::_ViewFields> pViewFields;
					EC_HRES_MSG(pViews->Add(bstrViewName, Outlook::OlViewType::olTableView, Outlook::OlViewSaveOption::olViewSaveOptionThisFolderOnlyMe, &pNewView), "creating new view");
					if (pNewView)
					{
						// cast the view as a TableView
						EC_HRES_MSG(pNewView->QueryInterface(__uuidof(Outlook::_TableView), (void **)&pTableView), "getting TableView object from newly created table");
						if (pTableView)
						{
							// Get the ViewFields collection
							EC_HRES_MSG(pTableView->get_ViewFields(&pViewFields), "getting fields collection");
							if (pViewFields)
							{
								for (unsigned int i = 0; i < ulFieldCount; i++)
								{
									// add the new fields
									CComPtr<Outlook::_ViewField> pViewField;
									if (S_OK != pViewFields->Add(bstrFieldName[i], &pViewField))
									{
										// if a fiels with the same name exists we'll just handle it
										std::wstring wszMessage = L"Unable to add field: ";
										wszMessage += bstrFieldName[i];
										ATLTRACE(wszMessage.c_str());
									}
								}
							}
							// Save and apply the view
							EC_HRES_MSG(pTableView->Save(), "saving view");
							EC_HRES_MSG(pTableView->Apply(), "applying view");
							MessageBoxW(NULL, L"New view applied!", L"CPP_AddIn", MB_OK);
						}
					}
				}
			}
		}
	}

Error:
	return hRes;
}