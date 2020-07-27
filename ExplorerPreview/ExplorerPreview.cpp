#include "pch.h"
#include "framework.h"
#include "ExplorerPreview.h"

#define CHECKHR(__expr) { HRESULT hr = __expr; if (FAILED(hr)) { ATLTRACE(L"Failed hr:0x%08X", hr); return hr;} }

CComPtr<IPreviewHandler> handler;

class CFrame : public IPreviewHandlerFrame
{
	LONG m_ref;

public:
	CFrame() : m_ref(1) { }

	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		if (riid == IID_IUnknown || riid == IID_IPreviewHandlerFrame)
		{
			*ppv = static_cast<IPreviewHandlerFrame*>(static_cast<IUnknown*>(this));
			AddRef();
			return S_OK;
		}
		*ppv = NULL;

		CComHeapPtr<wchar_t> id;
		StringFromCLSID(riid, &id);
		ATLTRACE(L"QueryInterface iid: '%s'", id);
		return E_NOINTERFACE;
	}

	IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_ref); }
	IFACEMETHODIMP_(ULONG) Release() { LONG reg = InterlockedDecrement(&m_ref); if (!reg) delete this; return reg; }

	IFACEMETHODIMP GetWindowContext(PREVIEWHANDLERFRAMEINFO* pinfo)
	{
		ATLTRACE(L"GetWindowContext");
		return S_OK;
	}

	IFACEMETHODIMP TranslateAccelerator(MSG* pmsg)
	{
		ATLTRACE(L"TranslateAccelerator");
		return S_OK;
	}
};

CFrame test;

template<typename DispInterface>
class CDispInterfaceBase : public DispInterface
{
	LONG m_ref;
	CComPtr<IConnectionPoint> m_point;
	DWORD m_cookie;

public:
	CDispInterfaceBase() : m_ref(1), m_cookie(0) { }

	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		if (riid == IID_IUnknown || riid == IID_IDispatch || riid == __uuidof(DispInterface))
		{
			*ppv = static_cast<DispInterface*>(static_cast<IDispatch*>(this));
			AddRef();
			return S_OK;
		}
		*ppv = NULL;
		return E_NOINTERFACE;
	}

	IFACEMETHODIMP_(ULONG) AddRef() { return InterlockedIncrement(&m_ref); }
	IFACEMETHODIMP_(ULONG) Release() { LONG reg = InterlockedDecrement(&m_ref); if (!reg) delete this; return reg; }

	IFACEMETHODIMP GetTypeInfoCount(UINT* pctinfo) { *pctinfo = 0; return E_NOTIMPL; }
	IFACEMETHODIMP GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo) { *ppTInfo = nullptr; return E_NOTIMPL; }
	IFACEMETHODIMP GetIDsOfNames(REFIID, LPOLESTR* rgszNames, UINT cNames, LCID lcid, DISPID* rgDispId) { return E_NOTIMPL; }
	IFACEMETHODIMP Invoke(DISPID dispid, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		if (pvarResult)
		{
			VariantInit(pvarResult);
		}
		return Invoke(dispid, pdispparams, pvarResult);
	}

	virtual HRESULT Invoke(DISPID dispid, DISPPARAMS* pdispparams, VARIANT* pvarResult) = 0;

	HRESULT Connect(IUnknown* punk)
	{
		CComPtr<IConnectionPointContainer> container;
		CHECKHR(punk->QueryInterface(IID_PPV_ARGS(&container)));
		CHECKHR(container->FindConnectionPoint(__uuidof(DispInterface), &m_point));
		return m_point->Advise(this, &m_cookie);
	}

	void Disconnect()
	{
		if (m_cookie)
		{
			m_point->Unadvise(m_cookie);
			m_point.Release();
			m_cookie = 0;
		}
	}
};

class WebBrowserEvents : public CDispInterfaceBase<DWebBrowserEvents2>
{
	HRESULT Invoke(DISPID dispid, DISPPARAMS* pDispParams, VARIANT* pvarResult)
	{
		// there are other ways to find selection changed, but this one is simple since it only uses webbrowser (explorer) events
		if (dispid == DISPID_PROGRESSCHANGE)
		{
			// when progress = progress end, it's usable
			if (V_I4(&pDispParams->rgvarg[0]) == V_I4(&pDispParams->rgvarg[1]))
			{
				if (handler)
				{
					CComPtr<IObjectWithSite> site;
					if (SUCCEEDED(handler->QueryInterface(&site)))
					{
						site->SetSite(NULL);
					}

					handler->Unload();
					handler.Release();
				}

				CComPtr<IShellBrowser> browser;
				CHECKHR(CComQIPtr<IServiceProvider>(m_window)->QueryService(SID_STopLevelBrowser, IID_PPV_ARGS(&browser)));

				CComPtr<IShellView> view;
				CHECKHR(browser->QueryActiveShellView(&view));

				CComPtr<IFolderView2> folderView;
				CHECKHR(view->QueryInterface(&folderView));

				// get selected item
				int index = -1;
				folderView->GetSelectedItem(0, &index);
				if (index >= 0)
				{
					CComPtr<IShellItem2> item;
					CHECKHR(folderView->GetItem(index, IID_PPV_ARGS(&item)));

					CComHeapPtr<wchar_t> ext;
					CHECKHR(item->GetString(PKEY_ItemType, &ext));

					CComHeapPtr<wchar_t> parsingName;
					CHECKHR(item->GetDisplayName(SIGDN::SIGDN_DESKTOPABSOLUTEPARSING, &parsingName));

					ATLTRACE(L"WebBrowserEvent selected: '%s' ext: '%s'", parsingName, ext);

					CComPtr<IQueryAssociations> assoc;
					CHECKHR(item->BindToHandler(NULL, BHID_AssociationArray, IID_PPV_ARGS(&assoc)));

					WCHAR sclsid[48] = { 0 };
					DWORD size = 48;
					CHECKHR(assoc->GetString(ASSOCF_INIT_DEFAULTTOSTAR, ASSOCSTR_SHELLEXTENSION, L"{8895b1c6-b41f-4c1c-a562-0d564250836f}", sclsid, &size));
					ATLTRACE(L"WebBrowserEvent clsid: '%s'", sclsid);

					CLSID clsid;
					SHCLSIDFromString(sclsid, &clsid);
					CHECKHR(handler.CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER));

					CComPtr<IInitializeWithItem> iitem;
					if (SUCCEEDED(handler->QueryInterface(&iitem)))
					{
						CHECKHR(iitem->Initialize(item, STGM_READ));
					}
					else
					{
						CComPtr<IInitializeWithFile> ifile;
						if (SUCCEEDED(handler->QueryInterface(&ifile)))
						{
							CHECKHR(ifile->Initialize(parsingName, STGM_READ));
						}
						else
						{
							CComPtr<IInitializeWithStream> istream;
							CHECKHR(handler->QueryInterface(&istream));

							CComPtr<IStream> stream;
							CHECKHR(SHCreateStreamOnFile(parsingName, STGM_READ, &stream));
							CHECKHR(istream->Initialize(stream, STGM_READ));
						}
					}

					CComPtr<IObjectWithSite> site;
					if (SUCCEEDED(handler->QueryInterface(&site)))
					{
						site->SetSite(&test);
					}

					RECT rc;
					GetClientRect(m_hwnd, &rc);
					CHECKHR(handler->SetWindow(m_hwnd, &rc));
					CHECKHR(handler->DoPreview());
				}
			}
		}
		return S_OK;
	}

public:
	CComPtr<IDispatch> m_window;
	HWND m_hwnd;
};

class ShellWindowsEvents : public CDispInterfaceBase<DShellWindowsEvents>
{
	CAtlArray<WebBrowserEvents*> m_views;
	HWND m_hwnd;

	HRESULT HookViews()
	{
		Stop();
		CComPtr<IShellWindows> windows;
		CHECKHR(windows.CoCreateInstance(CLSID_ShellWindows));
		long count = 0;
		windows->get_Count(&count);
		for (long i = 0; i < count; i++)
		{
			CComPtr<IDispatch> window;
			CHECKHR(windows->Item(CComVariant(i), &window));

			CComPtr<IWebBrowserApp> ebrowser;
			CHECKHR(CComQIPtr<IServiceProvider>(window)->QueryService(SID_SWebBrowserApp, IID_PPV_ARGS(&ebrowser)));

			WebBrowserEvents* events = new WebBrowserEvents();
			events->m_hwnd = m_hwnd;
			window.QueryInterface(&events->m_window);
			if (FAILED(events->Connect(ebrowser)))
			{
				delete events;
			}
			else
			{
				m_views.Add(events);
			}
		}
		return S_OK;
	}

	HRESULT Invoke(DISPID dispid, DISPPARAMS* pDispParams, VARIANT* pvarResult)
	{
		// unfortunately, FindWindowSW doesn't seem to work with a cookie (!?)
		// so we reset the whole collection each time...
		HookViews();
		return S_OK;
	}

public:

	HRESULT Start(HWND hwnd)
	{
		m_hwnd = hwnd;
		CComPtr<IShellWindows> windows;
		CHECKHR(windows.CoCreateInstance(CLSID_ShellWindows));
		return Connect(windows);
	}

	void Stop()
	{
		for (size_t i = 0; i < m_views.GetCount(); i++)
		{
			WebBrowserEvents* p = m_views.GetAt(i);
			p->Disconnect();
			delete p;
		}
		m_views.RemoveAll();
	}
};

#define MAX_LOADSTRING 100

ShellWindowsEvents windowEvents;
HINSTANCE hInst;
WCHAR szTitle[MAX_LOADSTRING];
WCHAR szWindowClass[MAX_LOADSTRING];

ATOM MyRegisterClass(HINSTANCE hInstance);
HWND InitInstance(HINSTANCE, int);
LRESULT CALLBACK WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK About(HWND, UINT, WPARAM, LPARAM);

int APIENTRY wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nCmdShow)
{
	UNREFERENCED_PARAMETER(hPrevInstance);
	UNREFERENCED_PARAMETER(lpCmdLine);

	CoInitialize(NULL);
	LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
	LoadStringW(hInstance, IDC_EXPLORERPREVIEW, szWindowClass, MAX_LOADSTRING);
	MyRegisterClass(hInstance);

	HWND hwnd = InitInstance(hInstance, nCmdShow);
	if (!hwnd)
		return FALSE;

	windowEvents.Start(hwnd);
	HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_EXPLORERPREVIEW));
	MSG msg;
	while (GetMessage(&msg, nullptr, 0, 0))
	{
		if (handler)
		{
			//handler->TranslateAcceleratorW(&msg);
			//HWND hwnd;
			//handler->QueryFocus(&hwnd);
		}
		if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	windowEvents.Stop();
	CoUninitialize();
	return (int)msg.wParam;
}

ATOM MyRegisterClass(HINSTANCE hInstance)
{
	WNDCLASSEXW wcex;
	wcex.cbSize = sizeof(WNDCLASSEX);
	wcex.style = CS_HREDRAW | CS_VREDRAW;
	wcex.lpfnWndProc = WndProc;
	wcex.cbClsExtra = 0;
	wcex.cbWndExtra = 0;
	wcex.hInstance = hInstance;
	wcex.hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_EXPLORERPREVIEW));
	wcex.hCursor = LoadCursor(nullptr, IDC_ARROW);
	wcex.hbrBackground = (HBRUSH)(COLOR_WINDOW + 1);
	wcex.lpszMenuName = MAKEINTRESOURCEW(IDC_EXPLORERPREVIEW);
	wcex.lpszClassName = szWindowClass;
	wcex.hIconSm = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));
	return RegisterClassExW(&wcex);
}

HWND InitInstance(HINSTANCE hInstance, int nCmdShow)
{
	hInst = hInstance;
	HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW, 10, 10, 800, 800, nullptr, nullptr, hInstance, nullptr);
	if (!hWnd)
		return NULL;

	ShowWindow(hWnd, nCmdShow);
	UpdateWindow(hWnd);
	return hWnd;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_COMMAND:
	{
		int wmId = LOWORD(wParam);
		switch (wmId)
		{
		case IDM_ABOUT:
			DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
			break;

		case IDM_EXIT:
			DestroyWindow(hWnd);
			break;

		default:
			return DefWindowProc(hWnd, message, wParam, lParam);
		}
	}
	break;

	case WM_PAINT:
	{
		PAINTSTRUCT ps;
		HDC hdc = BeginPaint(hWnd, &ps);
		EndPaint(hWnd, &ps);
	}
	break;

	case WM_DESTROY:
		PostQuitMessage(0);
		break;

	default:
		return DefWindowProc(hWnd, message, wParam, lParam);
	}
	return 0;
}

INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
	UNREFERENCED_PARAMETER(lParam);
	switch (message)
	{
	case WM_INITDIALOG:
		return (INT_PTR)TRUE;

	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
		{
			EndDialog(hDlg, LOWORD(wParam));
			return (INT_PTR)TRUE;
		}
		break;
	}
	return (INT_PTR)FALSE;
}
