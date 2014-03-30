#pragma once

#include "Guids.h"

class CRecycleBinWindow;

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

// CRecycleBinDeskBand

class CRecycleBinDeskBand :
    public CComObjectRootEx<CComSingleThreadModel>,
    public CComCoClass<CRecycleBinDeskBand, &CLSID_RecycleBinDeskBand>,
    public IObjectWithSiteImpl<CRecycleBinDeskBand>,
    public IPersistStreamInitImpl<CRecycleBinDeskBand>,
    public IInputObject,
    public IContextMenu,
    public IDeskBand2
{
    typedef IPersistStreamInitImpl<CRecycleBinDeskBand> IPersistStreamImpl;
public:
    CRecycleBinDeskBand();

    //DECLARE_REGISTRY_RESOURCEID(IDR_CALENDARDESKBAND)

    BEGIN_COM_MAP(CRecycleBinDeskBand)
        COM_INTERFACE_ENTRY(IOleWindow)
        COM_INTERFACE_ENTRY(IDockingWindow)
        COM_INTERFACE_ENTRY(IDeskBand)
        COM_INTERFACE_ENTRY(IDeskBand2)
        COM_INTERFACE_ENTRY(IInputObject)
        COM_INTERFACE_ENTRY(IContextMenu)
        COM_INTERFACE_ENTRY(IObjectWithSite)
        COM_INTERFACE_ENTRY_IID(IID_IPersist, IPersistStreamImpl)
        COM_INTERFACE_ENTRY_IID(IID_IPersistStream, IPersistStreamImpl)
        COM_INTERFACE_ENTRY_IID(IID_IPersistStreamInit, IPersistStreamImpl)
    END_COM_MAP()

    BEGIN_CATEGORY_MAP(CCalendarDeskBand)
        IMPLEMENTED_CATEGORY(CATID_DeskBand)
    END_CATEGORY_MAP()

    // IPersistStreamInitImpl requires property map.
    BEGIN_PROP_MAP(CRecycleBinDeskBand)
    END_PROP_MAP()

    DECLARE_PROTECT_FINAL_CONSTRUCT()

    HRESULT FinalConstruct();
    void FinalRelease();

    // Registry
    static HRESULT WINAPI UpdateRegistry(_In_ BOOL bRegister) throw();
    static void ExtractRecycleBinName(LPWSTR name, size_t size);


public:
    // IObjectWithSite
    //
    STDMETHOD(SetSite)(
        /* [in] */ IUnknown *pUnkSite);

    // IInputObject
    //
    STDMETHOD(UIActivateIO)(
        /* [in] */ BOOL fActivate,
        /* [unique][in] */ MSG *pMsg);

    STDMETHOD(HasFocusIO)();

    STDMETHOD(TranslateAcceleratorIO)(
        /* [in] */ MSG *pMsg);

    // IContextMenu
    //
    STDMETHOD(QueryContextMenu)(
        /* [in] */ HMENU hmenu,
        /* [in] */ UINT indexMenu,
        /* [in] */ UINT idCmdFirst,
        /* [in] */ UINT idCmdLast,
        /* [in] */ UINT uFlags);

    STDMETHOD(InvokeCommand)(
        /* [in] */ CMINVOKECOMMANDINFO *pici);

    STDMETHOD(GetCommandString)(
        /* [in] */ UINT_PTR idCmd,
        /* [in] */ UINT uType,
        /* [in] */ UINT *pReserved,
        /* [out] */ LPSTR pszName,
        /* [in] */ UINT cchMax);

    // IDeskBand
    //
    STDMETHOD(GetBandInfo)(
        /* [in] */ DWORD dwBandID,
        /* [in] */ DWORD dwViewMode,
        /* [out][in] */ DESKBANDINFO *pdbi);

    // IDeskBand2
    //
    STDMETHOD(CanRenderComposited)(
        /* [out] */ BOOL *pfCanRenderComposited);

    STDMETHOD(SetCompositionState)(
        /* [in] */ BOOL fCompositionEnabled);

    STDMETHOD(GetCompositionState)(
        /* [out] */ BOOL *pfCompositionEnabled);

    // IOleWindow
    //
    STDMETHOD(GetWindow)(
        /* [out] */ HWND *phwnd);

    STDMETHOD(ContextSensitiveHelp)(
        /* [in] */ BOOL fEnterMode);

    // IDockingWindow
    //
    STDMETHOD(ShowDW)(
        /* [in] */ BOOL fShow);

    STDMETHOD(CloseDW)(
        /* [in] */ DWORD dwReserved);

    STDMETHOD(ResizeBorderDW)(
        /* [in] */ LPCRECT prcBorder,
        /* [in] */ IUnknown *punkToolbarSite,
        /* [in] */ BOOL fReserved);

    // IPersistStreamImpl
    //
    HRESULT IPersistStreamInit_Load(
        /* [in] */ LPSTREAM pStm,
        /* [in] */ const ATL_PROPMAP_ENTRY* pMap);

    HRESULT IPersistStreamInit_Save(
        /* [in] */ LPSTREAM pStm,
        /* [in] */ BOOL fClearDirty,
        /* [in] */ const ATL_PROPMAP_ENTRY* pMap);

private:
    HRESULT HandleShowRequest();
    HRESULT UpdateDeskband();


protected:
    ~CRecycleBinDeskBand();


public:

    void OnFocus(const BOOL fFocus);
    
    bool HasCompositionEnabled();
    
    __declspec(property(get = HasCompositionEnabled)) bool CompositionEnabled;


    bool m_bRequiresSave; // used by IPersistStreamInitImpl


private:
    SIZE CalculateSize();

    int m_nBandID;
    DWORD m_dwViewMode;
    BOOL m_bCompositionEnabled;
    
    CComQIPtr<IInputObjectSite>  m_spSite;                // parent site that contains deskband
    
    CRecycleBinWindow* m_pWindow;
    SIZE               m_size;
};

OBJECT_ENTRY_AUTO(CLSID_RecycleBinDeskBand, CRecycleBinDeskBand)
