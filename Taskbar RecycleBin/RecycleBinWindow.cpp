#include "StdAfx.h"
#include "RecycleBinDeskBand.h"
#include "RecycleBinWindow.h"

#include "debug.h"

#define RECTWIDTH(x)   ((x).right - (x).left)
#define RECTHEIGHT(x)  ((x).bottom - (x).top)

#define LPARAM_TO_POINT(l) POINT{ GET_X_LPARAM(l), GET_Y_LPARAM(l) }

#define WM_SHELLNOTIFY (WM_USER + 5)

using namespace Gdiplus;


extern HINSTANCE   g_hInst;

static const WCHAR g_szDeskBandClass[] = L"RecycleBinDeskBand";


CRecycleBinWindow::CRecycleBinWindow(HWND parent, CRecycleBinDeskBand* pDeskBand) :
m_hWnd(nullptr), m_pDeskBand(pDeskBand), m_gdiplusToken(0), m_clicked(false), m_hovered(false), m_fHasFocus(FALSE), m_ulRecycleBinNotify(0)
{
    WNDCLASSW wc = { 0 };
    wc.style = CS_HREDRAW | CS_VREDRAW;
    wc.hCursor = LoadCursor(nullptr, IDC_ARROW);
    wc.hInstance = g_hInst;
    wc.lpfnWndProc = WndProc;
    wc.lpszClassName = g_szDeskBandClass;
    wc.hbrBackground = CreateSolidBrush(RGB(255, 255, 0));

    RegisterClassW(&wc);

    m_hWnd = CreateWindowExW(0,
        g_szDeskBandClass,
        nullptr,
        WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS,
        0,
        0,
        0,
        0,
        parent,
        nullptr,
        g_hInst,
        this);

    CoInitializeEx(nullptr, COINIT_APARTMENTTHREADED);

    GdiplusStartupInput gdiplusStartupInput;
    ::GdiplusStartup(&m_gdiplusToken, &gdiplusStartupInput, nullptr);

    RegisterRecycleBinNotify();
}

CRecycleBinWindow::~CRecycleBinWindow()
{
    GdiplusShutdown(m_gdiplusToken);

    CoUninitialize();

    UnregisterRecycleBinNotify();
}

void CRecycleBinWindow::OnPaint(const HDC hdcIn)
{
    HDC hdc = hdcIn;
    PAINTSTRUCT ps;

    if (!hdc)
    {
        hdc = BeginPaint(m_hWnd, &ps);
    }

    if (hdc)
    {
        RECT rc;
        GetClientRect(m_hWnd, &rc);

        LONG width  = RECTWIDTH(rc);
        LONG height = RECTHEIGHT(rc);

        if (CompositionEnabled)
        {
            HTHEME hTheme = OpenThemeData(nullptr, L"BUTTON");
            if (hTheme)
            {
                HDC hdcPaint = nullptr;
                HPAINTBUFFER hBufferedPaint = BeginBufferedPaint(hdc, &rc, BPBF_TOPDOWNDIB, nullptr, &hdcPaint);

                DrawThemeParentBackground(m_hWnd, hdcPaint, &rc);

                if (m_hovered || m_clicked)
                {
                    Graphics g(hdcPaint);

                    Pen        borderPen = { m_clicked ? Color{ 150, 255, 255, 255 } : Color{ 100, 255, 255, 255 } };
                    SolidBrush fillBrush = { m_clicked ? Color{ 100, 255, 255, 255 } : Color{ 50, 255, 255, 255 } };
                    Rect       selectionRect = { 5, 2, width - 10, height - 5 };
                    
                    g.DrawRectangle(&borderPen, selectionRect);
                    g.FillRectangle(&fillBrush, selectionRect);
                }

                LPITEMIDLIST pidl;
                if (SUCCEEDED(SHGetFolderLocation(0, CSIDL_BITBUCKET, 0, 0, &pidl)))
                {
                    SHFILEINFO sfi;
                    if (SHGetFileInfo((LPCWSTR)pidl, 0, &sfi, sizeof(sfi), SHGFI_ICON | SHGFI_PIDL))
                    {
                        SIZE  iconSize     = GetIconSize(sfi.hIcon);
                        POINT iconLocation = { (width - iconSize.cx) / 2, (height - iconSize.cy) / 2 };

                        if (m_clicked)
                        {
                            iconLocation.x++;
                            iconLocation.y++;
                        }

                        DrawIcon(hdcPaint, iconLocation.x, iconLocation.y, sfi.hIcon);

                        DestroyIcon(sfi.hIcon);
                    }
                    CoTaskMemFree(pidl);
                }

                EndBufferedPaint(hBufferedPaint, TRUE);

                CloseThemeData(hTheme);
            }
        }
        else
        {
            // TODO
        }
    }

    if (!hdcIn)
    {
        EndPaint(m_hWnd, &ps);
    }
}

void CRecycleBinWindow::OnFocus(const BOOL fFocus)
{
    SetFocus(fFocus);
}

void CRecycleBinWindow::OnMouseButtonDown(const int button, const POINT point)
{
    switch (button)
    {
    case LEFT_MOUSE_BUTTON:
        m_clicked = true;
        SetCapture(m_hWnd);
        Update();
        break;
    }
}

void CRecycleBinWindow::OnMouseButtonUp(const int button, const POINT point)
{
    switch (button)
    {
    case LEFT_MOUSE_BUTTON:
        m_clicked = false;
        ReleaseCapture();
        Update();
        break;
    }
}

void CRecycleBinWindow::OnMouseButtonClick(const int button, const POINT point)
{
    switch (button)
    {
    case LEFT_MOUSE_BUTTON:
        ShowRecycleBinFolder();
        break;
    //case RIGHT_MOUSE_BUTTON:
    //    HMENU hPopupMenu = CreatePopupMenu();
    //    InsertMenu(hPopupMenu, 0, MF_BYPOSITION | MF_STRING, 1, L"Exit");
    //    InsertMenu(hPopupMenu, 0, MF_BYPOSITION | MF_STRING, 2, L"Play");
    //    SetForegroundWindow(m_hwnd);
    //    TrackPopupMenu(hPopupMenu, TPM_BOTTOMALIGN | TPM_LEFTALIGN, 0, 0, 0, m_hwnd, NULL);
    //    break;
    }
}

void CRecycleBinWindow::OnMouseMove(const POINT point)
{
    if (!m_hovered)
    {
        TrackMouseEvents();

        m_hovered = true;
        Update();
    }
}

void CRecycleBinWindow::OnMouseHover(const POINT point)
{
}

void CRecycleBinWindow::OnMouseLeave(const POINT point)
{
    m_hovered = false;
    Update();
}


BOOL CRecycleBinWindow::HasFocus()
{
    return m_fHasFocus;
}

void CRecycleBinWindow::SetFocus(BOOL fFocus)
{
    m_fHasFocus = fFocus;

    if (fFocus)
    {
        ::SetFocus(m_hWnd);
    }

    m_pDeskBand->OnFocus(fFocus);
}

void CRecycleBinWindow::Show()
{
    ShowWindow(m_hWnd, SW_SHOW);
}

void CRecycleBinWindow::Hide()
{
    ShowWindow(m_hWnd, SW_HIDE);
}

void CRecycleBinWindow::Update()
{
    InvalidateRect(m_hWnd, nullptr, TRUE);
    ::UpdateWindow(m_hWnd);
}

void CRecycleBinWindow::Kill()
{
    ShowWindow(m_hWnd, SW_HIDE);
    DestroyWindow(m_hWnd);
    delete this;
    m_hWnd = nullptr;
}

HWND CRecycleBinWindow::GetHwnd()
{
    return m_hWnd;
}

bool CRecycleBinWindow::HasCompositionEnabled()
{
    return m_pDeskBand->CompositionEnabled;
}

LRESULT CALLBACK CRecycleBinWindow::WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
    LRESULT lResult = 0;
    POINT   point;

    CRecycleBinWindow* pWindow = reinterpret_cast<CRecycleBinWindow*>(GetWindowLongPtr(hwnd, GWLP_USERDATA));

    switch (uMsg)
    {
    case WM_CREATE:
        pWindow = reinterpret_cast<CRecycleBinWindow*>(reinterpret_cast<CREATESTRUCT*>(lParam)->lpCreateParams);
        SetWindowLongPtr(hwnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pWindow));
        break;
        
    case WM_LBUTTONDOWN:
        pWindow->OnMouseButtonDown(LEFT_MOUSE_BUTTON, LPARAM_TO_POINT(lParam));
        break;

    case WM_LBUTTONUP:
        point = LPARAM_TO_POINT(lParam);

        if (IsPointInWindow(hwnd, point))
        {
            pWindow->OnMouseButtonClick(LEFT_MOUSE_BUTTON, point);
        }

        pWindow->OnMouseButtonUp(LEFT_MOUSE_BUTTON, point);
        break;

    case WM_RBUTTONDOWN:
        pWindow->OnMouseButtonDown(RIGHT_MOUSE_BUTTON, LPARAM_TO_POINT(lParam));
        break;

    case WM_RBUTTONUP:
        point = LPARAM_TO_POINT(lParam);

        if (IsPointInWindow(hwnd, point))
        {
            pWindow->OnMouseButtonClick(RIGHT_MOUSE_BUTTON, point);
        }

        pWindow->OnMouseButtonUp(RIGHT_MOUSE_BUTTON, point);
        break;

    case WM_MOUSEMOVE:
        pWindow->OnMouseMove(LPARAM_TO_POINT(lParam));
        break;

    case WM_MOUSEHOVER:
        pWindow->OnMouseHover(LPARAM_TO_POINT(lParam));
        break;

    case WM_MOUSELEAVE:
        pWindow->OnMouseLeave(LPARAM_TO_POINT(lParam));
        break;

    case WM_SETFOCUS:
        pWindow->OnFocus(TRUE);
        break;

    case WM_KILLFOCUS:
        pWindow->OnFocus(FALSE);
        break;

    case WM_PAINT:
        pWindow->OnPaint(nullptr);
        break;

    case WM_PRINTCLIENT:
        pWindow->OnPaint(reinterpret_cast<HDC>(wParam));
        break;

    case WM_ERASEBKGND:
        if (pWindow->CompositionEnabled)
        {
            lResult = 1;
        }
        break;

    case WM_SHELLNOTIFY:
        pWindow->Update();
        break;
    }

    if (uMsg != WM_ERASEBKGND)
    {
        lResult = DefWindowProc(hwnd, uMsg, wParam, lParam);
    }

    return lResult;
}

SIZE CRecycleBinWindow::GetIconSize(const HICON icon)
{
    SIZE size = { 0 };

    ICONINFO info = { 0 };

    if (GetIconInfo(icon, &info))
    {
        BITMAP bmp = { 0 };

        if (info.hbmColor)
        {
            const int nWrittenBytes = GetObject(info.hbmColor, sizeof(bmp), &bmp);
            if (nWrittenBytes > 0)
            {
                size.cx = bmp.bmWidth;
                size.cy = bmp.bmHeight;
            }

            DeleteObject(info.hbmColor);
        }
        else if (info.hbmMask)
        {
            // Icon has no color plane, image data stored in mask
            const int nWrittenBytes = GetObject(info.hbmMask, sizeof(bmp), &bmp);
            if (nWrittenBytes > 0)
            {
                size.cx = bmp.bmWidth;
                size.cy = bmp.bmHeight;
            }

            DeleteObject(info.hbmMask);
        }
    }

    return size;
}

void CRecycleBinWindow::ShowRecycleBinFolder()
{
    ShellExecute(nullptr, nullptr, L"explorer", L"shell:recyclebinfolder", nullptr, SW_SHOWDEFAULT);
}

void CRecycleBinWindow::TrackMouseEvents()
{
    TRACKMOUSEEVENT tme;
    tme.cbSize = sizeof(TRACKMOUSEEVENT);
    tme.dwFlags = TME_HOVER | TME_LEAVE; //Type of events to track & trigger.
    tme.dwHoverTime = 1; //How long the mouse has to be in the window to trigger a hover event.
    tme.hwndTrack = m_hWnd;
    TrackMouseEvent(&tme);
}

bool CRecycleBinWindow::IsPointInWindow(const HWND hwnd, const POINT clientPoint)
{
    POINT screenPoint = clientPoint;
    RECT rect;
    ClientToScreen(hwnd, &screenPoint);
    GetWindowRect(hwnd, &rect);
    return PtInRect(&rect, screenPoint) != 0;
}

void CRecycleBinWindow::CreateToolTip(const HWND hwndParent, const LPWSTR text)
{
    // Create a tooltip.
    HWND hwndTT = CreateWindowEx(
        WS_EX_TOPMOST, 
        TOOLTIPS_CLASS, 
        nullptr,
        WS_POPUP | TTS_NOPREFIX | TTS_ALWAYSTIP,
        CW_USEDEFAULT, 
        CW_USEDEFAULT,
        CW_USEDEFAULT, 
        CW_USEDEFAULT,
        hwndParent, 
        nullptr,
        g_hInst, 
        nullptr);

    SetWindowPos(
        hwndTT, 
        HWND_TOPMOST, 
        0, 0, 0, 0,
        SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE);

    // Set up "tool" information. In this case, the "tool" is the entire parent window.

    TOOLINFO ti = { 0 };
    ti.cbSize = sizeof(TOOLINFO);
    ti.uFlags = TTF_SUBCLASS;
    ti.hwnd = hwndParent;
    ti.hinst = g_hInst;
    ti.lpszText = text;

    GetClientRect(hwndParent, &ti.rect);

    // Associate the tooltip with the "tool" window.
    SendMessage(hwndTT, TTM_ADDTOOL, 0, (LPARAM)(LPTOOLINFO)&ti);
}

void CRecycleBinWindow::RegisterRecycleBinNotify()
{
    LOG("RegisterRecycleBinNotify()");

    if (!m_ulRecycleBinNotify)
    {
        LPMALLOC pMalloc = NULL;
        SHGetMalloc(&pMalloc);

        LPITEMIDLIST pidlRecycleBin = { 0 };

        if (SHGetSpecialFolderLocation(m_hWnd, CSIDL_BITBUCKET, &pidlRecycleBin) == NOERROR)
        {
            SHChangeNotifyEntry pshcne;

            pshcne.pidl = pidlRecycleBin;
            pshcne.fRecursive = TRUE;
            m_ulRecycleBinNotify = SHChangeNotifyRegister(m_hWnd,
                SHCNRF_InterruptLevel | SHCNRF_ShellLevel,
                SHCNE_ALLEVENTS,
                WM_SHELLNOTIFY, /* Message that would be sent by the Shell */
                1,
                &pshcne);

            if (!m_ulRecycleBinNotify)
            {
                LOG("WARNING - Could not register to Recycle Bin shell notifications!");
            }
        }
        else
        {
            LOG("WARNING - Could not get Recycle Bin location!");
        }

        // free objects
        if (!pidlRecycleBin)
        {
            pMalloc->Free(pidlRecycleBin);
        }

        pMalloc->Release();
    }
}

void CRecycleBinWindow::UnregisterRecycleBinNotify()
{
    LOG("UnregisterRecycleBinNotify()");

    if (m_ulRecycleBinNotify)
    {
        SHChangeNotifyDeregister(m_ulRecycleBinNotify);
        m_ulRecycleBinNotify = 0;
    }
}
