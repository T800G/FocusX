#ifndef _HELPERS_9EF22F0D_46AB_423F_BA5A_D1BD482453D1_
#define _HELPERS_9EF22F0D_46AB_423F_BA5A_D1BD482453D1_
#include "debugtrace.h"

BOOL IsWindowCurrentThread(HWND hWnd)
{
return (GetWindowThreadProcessId(hWnd, NULL) == GetCurrentThreadId());
}

BOOL IsWindowCurrentProcess(HWND hWnd)
{
	DWORD dwpid;
return (GetWindowThreadProcessId(hWnd, &dwpid) && (dwpid == GetCurrentProcessId()));
}

BOOL IsWindowCurrentProcessThread(HWND hWnd)
{
	DWORD dwpid;
return (GetWindowThreadProcessId(hWnd, &dwpid)==GetCurrentThreadId()) && (dwpid==GetCurrentProcessId());
}

BOOL IsWindowThreadProcess(HWND hWnd, __in_opt const DWORD dwThreadId, __in_opt const DWORD dwProcessId)
{
	DWORD dwpid;
	DWORD dwtid = GetWindowThreadProcessId(hWnd, &dwpid);
	return (dwThreadId ? (dwThreadId == dwtid) : TRUE) && (dwProcessId ? (dwProcessId == dwpid) : TRUE);
}

BOOL IsWindowSameThreadProcess(HWND hWnd_1, HWND hWnd_2, __in_opt BOOL bThread = TRUE, __in_opt BOOL bProcess = TRUE)
{
	DWORD dwpid_1, dwpid_2;
	DWORD dwtid_1 = GetWindowThreadProcessId(hWnd_1, &dwpid_1);
	DWORD dwtid_2 = GetWindowThreadProcessId(hWnd_2, &dwpid_2);
	return (bThread ? (dwtid_1 == dwtid_2) : TRUE) && (bProcess ? (dwpid_1 == dwpid_2) : TRUE);
}

HWND GetOleWindow(IUnknown *punk, __in_opt const BOOL bQueryInterface = TRUE)
{
	IOleWindow* pow;
	if (bQueryInterface)
	{
		if FAILED(punk->QueryInterface((IID&)__uuidof(IOleWindow), (void**)&pow)) return NULL;
	}
	else pow = (IOleWindow*)punk;
	HWND how;
	if SUCCEEDED(pow->GetWindow(&how)) return how;
return NULL;
}

BOOL CopyWindowText(HWND hwFrom, HWND hwTo)
{
	//same thread and process only?
	//if (!IsWindowCurrentProcessThread(hwFrom) || !IsWindowCurrentProcessThread(hwTo)) return FALSE;

	//get text size and allocate buffer
	LRESULT lBufLen = SendMessage(hwFrom, WM_GETTEXTLENGTH, NULL, NULL) + 1;//terminating null is not counted
	LPTSTR pszBuf = (LPTSTR)LocalAlloc(LPTR, lBufLen*sizeof(TCHAR));
	if (pszBuf == NULL)
	{
		DBGTRACE2("LocalAlloc failed\n");
		return FALSE;
	}
	if (SendMessage(hwFrom, WM_GETTEXT, lBufLen, (LPARAM)pszBuf) == (lBufLen-1)) SendMessage(hwTo, WM_SETTEXT, NULL, (LPARAM)pszBuf);
	LocalFree(pszBuf);
return TRUE;
}

BOOL IsWindowClass(HWND hWnd, LPCTSTR pszClsName)
{
	TCHAR strBuf[256];
return ((GetClassName(hWnd, strBuf, 256) && lstrcmp(strBuf, pszClsName)==0));
}

BOOL IsWindowMaximized(HWND hWnd)
{
	if (NULL==hWnd) return FALSE;
	WINDOWPLACEMENT wpl;
	wpl.length=sizeof(WINDOWPLACEMENT);
	wpl.showCmd=SW_MAX;
	::GetWindowPlacement(hWnd, &wpl);
return (SW_SHOWMAXIMIZED==wpl.showCmd);
}

UINT GetChildWindowCount(HWND hWnd, __in_opt LPCTSTR lpszChildClass, __in_opt const BOOL bVisibleOnly = FALSE)
{
	UINT c=0;
	HWND hwch = FindWindowEx(hWnd, NULL, lpszChildClass, NULL);
	while (hwch)
	{
		if (bVisibleOnly) {if (IsWindowVisible(hwch)) c++;}
		else c++;
		hwch = FindWindowEx(hWnd, hwch, lpszChildClass, NULL);
	}
return c;
}

HWND FindGrandChildWindow(HWND hwParent,
							__in_opt LPCTSTR lpszChildClass,
							__in_opt LPCTSTR lpszGrandChildClass, 
							__in_opt const BOOL bVisibleChildOnly = FALSE,
							__in_opt const BOOL bVisibleGrandChildOnly = FALSE)
{
	HWND hwChild = FindWindowEx(hwParent, NULL, lpszChildClass, NULL);
	while (hwChild)
	{
		if (!(bVisibleChildOnly && !IsWindowVisible(hwChild)))
		{

			HWND hwGrandchild = FindWindowEx(hwChild, NULL, lpszGrandChildClass, NULL);
			while (hwGrandchild)
			{
				if (bVisibleGrandChildOnly && IsWindowVisible(hwChild)) return hwGrandchild;
				hwGrandchild = FindWindowEx(hwChild, hwGrandchild, lpszGrandChildClass, NULL);
			}
		}
		hwChild = FindWindowEx(hwParent, hwChild, lpszChildClass, NULL);
	}
return NULL;
}

HWND GetParentWindow(HWND hwChild, __in_opt LPCTSTR lpszParentClass = NULL, __in_opt const UINT zLevel=0)
{
	UINT c=0;
	HWND hwp = GetParent(hwChild);
	while (hwp)
	{
		c++;
		if (!(lpszParentClass && !IsWindowClass(hwp, lpszParentClass)))
			if (!(zLevel && (zLevel != c))) return hwp;
		hwp = GetParent(hwp);
	}
return hwp; //desktop (NULL) is topmost parent
}

HWND GetActiveMDIChild(WNDPROC mdiProc, HWND hwMDI, LPARAM lParam = NULL)
{
	return (HWND)mdiProc(hwMDI, WM_MDIGETACTIVE, NULL, lParam);
}

BOOL ScreenToClientRect(HWND hWindow, LPRECT lpRc)
{
	if (!::ScreenToClient(hWindow, reinterpret_cast<POINT*>(&lpRc->left)  )) return FALSE;
	if (!::ScreenToClient(hWindow, reinterpret_cast<POINT*>(&lpRc->right) )) return FALSE;
return TRUE;
}

HFONT CreateMenuFont()//LONG lfWeight)
{
	NONCLIENTMETRICS ncm;
	ncm.cbSize = sizeof(NONCLIENTMETRICS);
	if (!SystemParametersInfo(SPI_GETNONCLIENTMETRICS,0,&ncm,0)) return NULL;
	//ncm.lfMenuFont.lfWeight = lfWeight;//CLEARTYPE_QUALITY?
	return CreateFontIndirect(&ncm.lfMenuFont);
}

#ifndef CCoInitialize
//https://blogs.msdn.microsoft.com/oldnewthing/20040520-00/?p=39243
class CCoInitialize
{
	public:
	CCoInitialize() : m_hr(CoInitialize(NULL)) { }
	~CCoInitialize() { if (SUCCEEDED(m_hr)) CoUninitialize(); }
	operator HRESULT() const { return m_hr; }
	HRESULT m_hr;
};
#endif

#ifndef SafeRelease
template <class T> void SafeRelease(T **ppT)
{
    if (*ppT)
    {
        (*ppT)->Release();
        *ppT = NULL;
    }
}
#endif


#endif//_HELPERS_9EF22F0D_46AB_423F_BA5A_D1BD482453D1_
