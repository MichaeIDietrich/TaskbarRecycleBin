#include "StdAfx.h"
#include "RecycleBinDeskBand.h"
#include "RecycleBinWindow.h"

#include "debug.h"

#define RECTWIDTH(x)   ((x).right - (x).left)
#define RECTHEIGHT(x)  ((x).bottom - (x).top)

#define IDM_SEPARATOR_OFFSET 0
#define IDM_EMPTY_RECYCLE_BIN_OFFSET 1

extern HINSTANCE   g_hInst;

static const WCHAR g_szDeskBandClass[] = L"RecycleBinDeskBand";


CRecycleBinDeskBand::CRecycleBinDeskBand() :
m_bRequiresSave(false), m_dwViewMode(0), m_bCompositionEnabled(FALSE), m_spSite(nullptr), m_nBandID(0), m_pWindow(nullptr)
{
    LOG("CRecycleBinDeskBan()");

    m_size = CalculateSize();
}

CRecycleBinDeskBand::~CRecycleBinDeskBand()
{
    LOG("~CRecycleBinDeskBand()");
}


////////////////////////////////////////////////////////////////////////////////
// 
HRESULT CRecycleBinDeskBand::FinalConstruct()
{
    LOG("FinalConstruct()");

    m_nBandID = 0;
    m_dwViewMode = DBIF_VIEWMODE_NORMAL;
    m_bRequiresSave = false;
    m_bCompositionEnabled = TRUE;

    return HandleShowRequest();
}


void CRecycleBinDeskBand::FinalRelease()
{
    LOG("FinalRelease()");
}


HRESULT CRecycleBinDeskBand::HandleShowRequest()
{
    LOG("HandleShowRequest()");

    return S_OK;

    // FIXME: seems not to work under Windows 8 any longer

    //CComPtr<ITrayDeskBand> spTrayDeskBand;
    //HRESULT hr = spTrayDeskBand.CoCreateInstance(CLSID_TrayDeskBand);

    //if (SUCCEEDED(hr))   // Vista and higher
    //{
    //    hr = spTrayDeskBand->DeskBandRegistrationChanged();
    //    ATLASSERT(SUCCEEDED(hr));

    //    if (SUCCEEDED(hr))
    //    {
    //        hr = spTrayDeskBand->IsDeskBandShown(CLSID_RecycleBinDeskBand);
    //        ATLASSERT(SUCCEEDED(hr));

    //        if (SUCCEEDED(hr) && hr == S_FALSE)
    //            hr = spTrayDeskBand->ShowDeskBand(CLSID_RecycleBinDeskBand);
    //    }
    //}
    //else
    //{
    //    LOG("INFO - Could not display deskband programmatically, maybe it is Windows XP")
    //}

    //return SUCCEEDED(hr);
}

HRESULT CRecycleBinDeskBand::UpdateDeskband()
{
    LOG("UpdateDeskband()");

    CComPtr<IInputObjectSite> spInputSite;
    HRESULT hr = GetSite(IID_IInputObjectSite, reinterpret_cast<void**>(&spInputSite));

    if (SUCCEEDED(hr))
    {
        CComQIPtr<IOleCommandTarget> spOleCmdTarget = spInputSite;

        if (spOleCmdTarget)
        {
            // m_nBandID must be `int' or bandID variant must be explicitly
            // set to VT_I4, otherwise IDeskBand::GetBandInfo won't
            // be called by the system.
            CComVariant bandID(m_nBandID);

            hr = spOleCmdTarget->Exec(&CGID_DeskBand, DBID_BANDINFOCHANGED, OLECMDEXECOPT_DODEFAULT, &bandID, NULL);
            ATLASSERT(SUCCEEDED(hr));

            //if (SUCCEEDED(hr) && m_pWindow)
            //    m_pWindow->UpdateCaption();
        }
    }

    return hr;
}


HRESULT WINAPI CRecycleBinDeskBand::UpdateRegistry(_In_ BOOL bRegister)
{
    LOG("UpdateRegistry()");
    LOG2("  bRegister", bRegister);

    HRESULT hr = S_OK;

    if (bRegister)
    {
        WCHAR szCLSID[MAX_PATH];
        StringFromGUID2(CLSID_RecycleBinDeskBand, szCLSID, ARRAYSIZE(szCLSID));

        WCHAR szSubkey[MAX_PATH];
        HKEY hKey;

        hr = StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
        if (SUCCEEDED(hr))
        {
            hr = E_FAIL;
            if (ERROR_SUCCESS == RegCreateKeyExW(HKEY_CLASSES_ROOT,
                szSubkey,
                0,
                NULL,
                REG_OPTION_NON_VOLATILE,
                KEY_WRITE,
                NULL,
                &hKey,
                NULL))
            {
                WCHAR const szName[] = L"Recyclebin"; // TODO: externalize
                //WCHAR szName[MAX_PATH];
                //ExtractRecycleBinName(szName, ARRAYSIZE(szName));

                if (ERROR_SUCCESS == RegSetValueExW(hKey,
                    NULL,
                    0,
                    REG_SZ,
                    (LPBYTE)szName,
                    sizeof(szName)))
                {
                    hr = S_OK;
                }

                RegCloseKey(hKey);
            }
        }

        if (SUCCEEDED(hr))
        {
            hr = StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s\\InprocServer32", szCLSID);
            if (SUCCEEDED(hr))
            {
                hr = HRESULT_FROM_WIN32(RegCreateKeyExW(HKEY_CLASSES_ROOT, szSubkey,
                    0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL));
                if (SUCCEEDED(hr))
                {
                    WCHAR szModule[MAX_PATH];
                    if (GetModuleFileNameW(g_hInst, szModule, ARRAYSIZE(szModule)))
                    {
                        DWORD cch = lstrlen(szModule);
                        hr = HRESULT_FROM_WIN32(RegSetValueExW(hKey, NULL, 0, REG_SZ, (LPBYTE)szModule, cch * sizeof(szModule[0])));
                    }

                    if (SUCCEEDED(hr))
                    {
                        WCHAR const szModel[] = L"Apartment";
                        hr = HRESULT_FROM_WIN32(RegSetValueExW(hKey, L"ThreadingModel", 0, REG_SZ, (LPBYTE)szModel, sizeof(szModel)));
                    }

                    RegCloseKey(hKey);
                }
            }
        }
    }
    else
    {
        WCHAR szCLSID[MAX_PATH];
        StringFromGUID2(CLSID_RecycleBinDeskBand, szCLSID, ARRAYSIZE(szCLSID));
        
        WCHAR szSubkey[MAX_PATH];
        HRESULT hr = StringCchPrintfW(szSubkey, ARRAYSIZE(szSubkey), L"CLSID\\%s", szCLSID);
        if (SUCCEEDED(hr))
        {
            if (ERROR_SUCCESS != RegDeleteTreeW(HKEY_CLASSES_ROOT, szSubkey))
            {
                hr = E_FAIL;
            }
        }
        
        return SUCCEEDED(hr) ? S_OK : SELFREG_E_CLASS;
    }

    LOG2("  hr", hr);

    return hr;
}


//
// IOleWindow
//
STDMETHODIMP CRecycleBinDeskBand::GetWindow(HWND* pHwnd)
{
    LOG("GetWindow()");

    if (m_pWindow)
    {
        *pHwnd = m_pWindow->Hwnd;
        return S_OK;
    }

    *pHwnd = nullptr;
    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::ContextSensitiveHelp(BOOL)
{
    LOG("ContextSensitiveHelp()");

    return E_NOTIMPL;
}

//
// IDockingWindow
//
STDMETHODIMP CRecycleBinDeskBand::ShowDW(BOOL fShow)
{
    LOG("ShowDW()")

    if (m_pWindow)
    {
        fShow ? m_pWindow->Show() : m_pWindow->Hide();
    }

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::CloseDW(DWORD)
{
    LOG("CloseDW()");

    if (m_pWindow)
    {
        m_pWindow->Kill();
        m_pWindow = nullptr;
    }

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::ResizeBorderDW(const RECT*, IUnknown*, BOOL)
{
    LOG("ResizeBorderDW()");

    return E_NOTIMPL;
}

//
// IDeskBand
//
HRESULT STDMETHODCALLTYPE CRecycleBinDeskBand::GetBandInfo(
    /* [in] */ DWORD dwBandID,
    /* [in] */ DWORD dwViewMode,
    /* [out][in] */ DESKBANDINFO *pdbi)
{
    LOG("GetBandInfo()");

    if (!pdbi) return E_INVALIDARG;

    m_nBandID = dwBandID;
    m_dwViewMode = dwViewMode;

    if (pdbi->dwMask & DBIM_MODEFLAGS)
    {
        pdbi->dwModeFlags = DBIMF_VARIABLEHEIGHT;
    }

    if (pdbi->dwMask & DBIM_MINSIZE)
    {
        pdbi->ptMinSize.x = m_size.cx;
        pdbi->ptMinSize.y = m_size.cy;
    }

    if (pdbi->dwMask & DBIM_MAXSIZE)
    {
        // the band object has no limit for its maximum height
        pdbi->ptMaxSize.x = -1;
        pdbi->ptMaxSize.y = -1;
    }

    if (pdbi->dwMask & DBIM_INTEGRAL)
    {
        pdbi->ptIntegral.x = 1;
        pdbi->ptIntegral.y = 1;
    }

    if (pdbi->dwMask & DBIM_ACTUAL)
    {
        pdbi->ptActual.x = m_size.cx;
        pdbi->ptActual.y = m_size.cy;

    }

    if (pdbi->dwMask & DBIM_TITLE)
    {
        if (dwViewMode == DBIF_VIEWMODE_FLOATING)
        {
            CString sTitle = L"RecycleBin"; // externalize
            lstrcpynW(pdbi->wszTitle, sTitle, ARRAYSIZE(pdbi->wszTitle));
        }
        else pdbi->dwMask &= ~DBIM_TITLE; // do not show title
    }

    if (pdbi->dwMask & DBIM_BKCOLOR)
    {
        //Use the default background color by removing this flag.
        pdbi->dwMask &= ~DBIM_BKCOLOR;
    }

    return S_OK;

}

//
// IDeskBand2
//
STDMETHODIMP CRecycleBinDeskBand::CanRenderComposited(BOOL* pfCanRenderComposited)
{
    LOG("CanRenderComposited()");

    if (pfCanRenderComposited)
    {
        *pfCanRenderComposited = TRUE;
    }

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::SetCompositionState(
    /* [in] */ BOOL fCompositionEnabled)
{
    LOG("SetCompositionState()");

    m_bCompositionEnabled = fCompositionEnabled;

    if (m_pWindow)
    {
        m_pWindow->Update();
    }

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::GetCompositionState(BOOL* pfCompositionEnabled)
{
    LOG("GetCompositionState()");

    if (pfCompositionEnabled)
    {
        *pfCompositionEnabled = m_bCompositionEnabled;
    }

    return S_OK;
}


//
// IObjectWithSite
//
STDMETHODIMP CRecycleBinDeskBand::SetSite(IUnknown* pUnkSite)
{
    LOG("SetSite()");

    HRESULT hr = __super::SetSite(pUnkSite);

    if (SUCCEEDED(hr) && pUnkSite) // pUnkSite is NULL when band is being destroyed
    {
        m_spSite = pUnkSite;
        CComQIPtr<IOleWindow> spOleWindow = pUnkSite;

        if (spOleWindow)
        {
            HWND hwndParent = NULL;
            hr = spOleWindow->GetWindow(&hwndParent);

            if (SUCCEEDED(hr))
            {
                if (m_pWindow)
                {
                    LOG("  kill Window");
                    m_pWindow->Kill();
                }

                LOG("  create Window");
                m_pWindow = new CRecycleBinWindow(hwndParent, this);

                if (m_pWindow == nullptr)
                {
                    hr = E_FAIL;
                }
            }
        }
    }

    return hr;

}


//
// IInputObject
//
STDMETHODIMP CRecycleBinDeskBand::UIActivateIO(BOOL fActivate, MSG*)
{
    LOG("UIActivateIO()");

    if (fActivate && m_pWindow)
    {
        m_pWindow->SetFocus(fActivate);
    }

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::HasFocusIO()
{
    LOG("HasFocusIO()");

    if (m_pWindow)
    {
        return (m_pWindow->HasFocus() ? S_OK : S_FALSE);
    }

    return S_FALSE;
}

STDMETHODIMP CRecycleBinDeskBand::TranslateAcceleratorIO(MSG*)
{
    LOG("TranslateAcceleratorIO()");

    return S_FALSE;
};

//
// IContextMenu
//
STDMETHODIMP CRecycleBinDeskBand::QueryContextMenu(HMENU hMenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
    LOG("QueryContextMenu()");

    if (CMF_DEFAULTONLY & uFlags)
    {
        return MAKE_HRESULT(SEVERITY_SUCCESS, 0, 0);
    }

    // Add a seperator
    ::InsertMenu(hMenu,
        indexMenu,
        MF_SEPARATOR | MF_BYPOSITION,
        idCmdFirst + IDM_SEPARATOR_OFFSET, nullptr);

    ::InsertMenu(hMenu,
        indexMenu,
        MF_STRING | MF_BYPOSITION,
        idCmdFirst + IDM_EMPTY_RECYCLE_BIN_OFFSET,
        L"Empty Recycle Bin"); // TODO: externalize

    return MAKE_HRESULT(SEVERITY_SUCCESS, FACILITY_NULL, IDM_EMPTY_RECYCLE_BIN_OFFSET + 1);
}

STDMETHODIMP CRecycleBinDeskBand::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
    LOG("GetCommandString()");

    return S_OK;
}

STDMETHODIMP CRecycleBinDeskBand::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
    LOG("InvokeCommand()");

    if (!pici) return E_INVALIDARG;

    if (LOWORD(pici->lpVerb) == IDM_EMPTY_RECYCLE_BIN_OFFSET)
    {
        HWND hwnd = nullptr;
        if (m_pWindow)
        {
            hwnd = m_pWindow->Hwnd;
        }
        SHEmptyRecycleBin(hwnd, nullptr, 0);
    }

    return S_OK;
}


////////////////////////////////////////////////////////////////////////////////
// IPersistStreamImpl

HRESULT CRecycleBinDeskBand::IPersistStreamInit_Load(
    LPSTREAM pStm,
    const ATL_PROPMAP_ENTRY* pMap)
{
    LOG("IPersistStreamInit_Load()");

    return S_OK;
}

HRESULT CRecycleBinDeskBand::IPersistStreamInit_Save(
    LPSTREAM pStm,
    BOOL fClearDirty,
    const ATL_PROPMAP_ENTRY* pMap)
{
    LOG("IpersistStreamInit_Save()");

    return __super::IPersistStreamInit_Save(pStm, fClearDirty, pMap);
}

////////////////////////////////////////////////////////////////////////////////


//
// CUSTOM METHODS
//
void CRecycleBinDeskBand::OnFocus(const BOOL fFocus)
{
    LOG("OnFocus()");

    if (m_spSite)
    {
        m_spSite->OnFocusChangeIS(static_cast<IOleWindow*>(this), fFocus);
    }
}

bool CRecycleBinDeskBand::HasCompositionEnabled()
{
    LOG("HasCompositionEnabled()");

    return m_bCompositionEnabled != 0;
}


SIZE CRecycleBinDeskBand::CalculateSize()
{
    LOG("CalculateSize()");

    SIZE size;

    APPBARDATA data = { 0 };

    data.cbSize = sizeof(data);

    SHAppBarMessage(ABM_GETTASKBARPOS, &data);

    LONG width = RECTWIDTH(data.rc);
    LONG height = RECTHEIGHT(data.rc);

    size.cx = size.cy = width < height ? width : height;

    return size;
}


void CRecycleBinDeskBand::ExtractRecycleBinName(LPWSTR name, size_t size)
{
    LPSHELLFOLDER pDesktop = nullptr;
    LPITEMIDLIST  pidlRecycleBin = nullptr;
    LPMALLOC      pMalloc = nullptr;

    SHGetMalloc(&pMalloc);

    if (SUCCEEDED(SHGetDesktopFolder(&pDesktop)) &&
        SUCCEEDED(SHGetSpecialFolderLocation(nullptr, CSIDL_BITBUCKET, &pidlRecycleBin)))
    {
        STRRET strRet;

        if (S_OK == pDesktop->GetDisplayNameOf(pidlRecycleBin, SHGDN_NORMAL, &strRet))
        {
            LPWSTR recycleBinName;
            StrRetToStr(&strRet, pidlRecycleBin, &recycleBinName);

            wcscpy_s(name, size, recycleBinName);

            CoTaskMemFree(recycleBinName);
        }
    }
}