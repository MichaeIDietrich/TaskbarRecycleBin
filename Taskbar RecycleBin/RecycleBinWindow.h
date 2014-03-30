#pragma once

#define LEFT_MOUSE_BUTTON 1
#define RIGHT_MOUSE_BUTTON 2
#define MIDDLE_MOUSE_BUTTON 3

class CRecycleBinDeskBand;

class CRecycleBinWindow
{

public:
    CRecycleBinWindow(HWND parent, CRecycleBinDeskBand* pDeskBand);


    void OnFocus(const BOOL fFocus);
    void OnPaint(const HDC hdcIn);
    void OnMouseButtonDown(const int button, const POINT point);
    void OnMouseButtonUp(const int button, const POINT point);
    void OnMouseButtonClick(const int button, const POINT point);
    void OnMouseMove(const POINT point);
    void OnMouseHover(const POINT point);
    void OnMouseLeave(const POINT point);

    BOOL HasFocus();
    void SetFocus(BOOL fFocus);
    void Show();
    void Hide();
    void Update();
    void Kill();

    HWND GetHwnd();
    bool HasCompositionEnabled();


    __declspec(property(get = GetHwnd)) HWND Hwnd;
    __declspec(property(get = HasCompositionEnabled)) bool CompositionEnabled;

    
protected:
    ~CRecycleBinWindow();


private:
    static LRESULT CALLBACK WndProc(HWND hwnd, UINT uMsg, WPARAM wParam, LPARAM lParam);
    static bool IsPointInWindow(const HWND hwnd, const POINT clientPoint);
    static void CreateToolTip(const HWND hwnd, const LPWSTR text);

    SIZE GetIconSize(const HICON icon);
    void ShowRecycleBinFolder();
    void TrackMouseEvents();
    void RegisterRecycleBinNotify();
    void UnregisterRecycleBinNotify();

    BOOL                 m_fHasFocus;
    ULONG                m_ulRecycleBinNotify;

    HWND                 m_hWnd;
    CRecycleBinDeskBand* m_pDeskBand;
    ULONG_PTR            m_gdiplusToken;
    bool                 m_hovered;
    bool                 m_clicked;
};