#include "stdafx.h"
#include "resource.h"
#include "helpers.h"

// Global Variables:
WNDCLASSEX g_wcex; //reusing g_wcex.hInstance and g_wcex.hIcon
#define DPS_WNDCLASS_NAME  _T("FocusXTrayWndClass")
#define WM_TRAY_NOTIFY (WM_APP+11)
HMENU g_htrayMenu;
NOTIFYICONDATA  g_trayicon;

HWINEVENTHOOK g_hWinEventHook;
HWND g_hwDUI;
DWORD g_lastProcessId;
DWORD g_lastThreadId;

/////////////////////////////////////////////////////////////////////////////////////////////////////////////
void CALLBACK WinEventProcCallback(HWINEVENTHOOK hook, DWORD dwEvent, HWND hwnd, LONG idObject, LONG idChild, DWORD dwEventThread, DWORD dwmsEventTime)
{
	//user holds Shift key or tabbed into folder tree pane
	if ((GetAsyncKeyState(VK_SHIFT)<0) || (GetAsyncKeyState(VK_TAB)<0)) return;
	switch (dwEvent)
	{
		case EVENT_OBJECT_FOCUS:
			//explorer window hierarchy (win7, win10):
			//SysTreeView32 -> NamespaceTreeControl -> CtrlNotifySink -> DirectUIHWND(*) -> DUIViewWndClassName -> ShellTabWindowClass -> CabinetWClass -> DESKTOP
			//DirectUIHWND  -> SHELLDLL_DefView     -> CtrlNotifySink -> DirectUIHWND(*)
			//from (*) and upward hwnds stay the same during navigation
			if (IsWindowClass(hwnd, _T("SysTreeView32")) && (g_hwDUI = GetParentWindow(hwnd, _T("DirectUIHWND"), 3)))
			{
				//check if this is really Windows Explorer
				HWND hw = GetParentWindow(g_hwDUI, _T("CabinetWClass"), 3);
				if (!(hw && (GetParent(hw) == HWND_DESKTOP))) return;
				//dwEventThread is window's thread but we need process id anyway...
				g_lastThreadId = GetWindowThreadProcessId(hwnd, &g_lastProcessId);
				//is there any guarantee that folder view renders on the same thread as explorer browser window?
			}
		break;
		case EVENT_OBJECT_SHOW:
			//folder view window is re-created when navigating
			if (idObject != OBJID_WINDOW) return;
			if (g_hwDUI == NULL) return;
			//checking process id should also handle Explorer crash case
			if (IsWindowThreadProcess(hwnd, g_lastThreadId, g_lastProcessId))
			{
				//luckily, shown window is DirectUIHWND inside SHELLDLL_DefView, just what we need
				//check if parent is the same: DirectUIHWND(*) <- CtrlNotifySink <- SHELLDLL_DefView <- DirectUIHWND
				if (GetParentWindow(hwnd, _T("DirectUIHWND"), 3) == g_hwDUI)
				{
					if (AttachThreadInput(GetCurrentThreadId(), g_lastThreadId, TRUE))
					{
						SetFocus(hwnd);  //GetFocus also works only if thread input is attached
						AttachThreadInput(GetCurrentThreadId(), g_lastThreadId, FALSE);
					}
				}
			}
			//events must be in correct order and from the same explorer window so last values are always reset in this event
			g_hwDUI = NULL;
			g_lastProcessId = 0;
			g_lastThreadId = 0;
		break;
	}
}
//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//tray menu window proc
LRESULT CALLBACK TrayWndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
		case WM_COMMAND:
			DestroyWindow(hWnd);
		break;
		case WM_TRAY_NOTIFY:
			if (LOWORD(lParam) == WM_RBUTTONUP)
			{
				POINT mousepos;
				if (!GetCursorPos(&mousepos)) return 0;
				::SetForegroundWindow(hWnd); //MSDN Q135788 fix 
				BOOL bT = ::TrackPopupMenuEx(GetSubMenu(g_htrayMenu,0), TPM_RIGHTALIGN, mousepos.x, mousepos.y, hWnd, NULL);
				::PostMessage(hWnd, WM_NULL, 0, 0);
			}
		break;
		case WM_ENDSESSION:
			if (wParam == TRUE) DestroyWindow(hWnd);
		break;
		case WM_DESTROY:
			PostQuitMessage(0);
		break;
		default: return DefWindowProc(hWnd, message, wParam, lParam);
	}
return 0;
}

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
int APIENTRY _tWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPTSTR lpCmdLine, int nCmdShow)
{
	CoInitialize(NULL);
	//init globals
	g_hwDUI = NULL;
	g_lastProcessId = 0;
	g_lastThreadId = 0;
	//init structs
	SecureZeroMemory(&g_wcex, sizeof(WNDCLASSEX));
	g_wcex.cbSize = sizeof(WNDCLASSEX);
	g_wcex.hInstance = GetModuleHandle(NULL);
	g_wcex.lpfnWndProc	= TrayWndProc;
	g_wcex.lpszMenuName	= MAKEINTRESOURCE(IDC_MAIN);
	g_wcex.lpszClassName = DPS_WNDCLASS_NAME;
	g_wcex.hIconSm = (HICON)::LoadImage(g_wcex.hInstance, MAKEINTRESOURCE(IDI_MAIN),
					IMAGE_ICON, ::GetSystemMetrics(SM_CXSMICON), ::GetSystemMetrics(SM_CYSMICON), LR_DEFAULTCOLOR);
	if (!RegisterClassEx(&g_wcex)) return 0;
	//create message-only window
	SecureZeroMemory(&g_trayicon, sizeof(NOTIFYICONDATA));
	g_trayicon.hWnd = CreateWindow(g_wcex.lpszClassName, NULL, WS_OVERLAPPEDWINDOW, CW_USEDEFAULT,
									0, CW_USEDEFAULT, 0, HWND_MESSAGE, NULL, g_wcex.hInstance, NULL);
	if (!g_trayicon.hWnd) return FALSE;
	g_htrayMenu=::GetMenu(g_trayicon.hWnd);
	//cmd switch to hide tray icon
	if (_tcsstr(GetCommandLine(), _T("/notray"))==NULL)
	{
		g_trayicon.cbSize = sizeof(NOTIFYICONDATA);
		g_trayicon.uFlags = (NIF_ICON | NIF_MESSAGE | NIF_TIP);
		//g_trayicon.hWnd is already set
		g_trayicon.hIcon = g_wcex.hIconSm;
		g_trayicon.dwInfoFlags = NIIF_USER;
		g_trayicon.uCallbackMessage = WM_TRAY_NOTIFY;
		g_trayicon.uTimeout = 3000;
		LoadString(GetModuleHandle(NULL), IDS_TRAYTOOLTIP, g_trayicon.szTip, sizeof(g_trayicon.szTip)/sizeof(TCHAR));
		Shell_NotifyIcon(NIM_ADD, &g_trayicon);
	}
	if (g_hWinEventHook = SetWinEventHook(EVENT_OBJECT_SHOW, EVENT_OBJECT_FOCUS, NULL, WinEventProcCallback, 0, 0, WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS))
	{
		MSG msg;
		while (GetMessage(&msg, NULL, 0, 0))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
		UnhookWinEvent(g_hWinEventHook);
	}
	if (g_trayicon.cbSize) Shell_NotifyIcon(NIM_DELETE, &g_trayicon);
	CoUninitialize();
return 0;
}
