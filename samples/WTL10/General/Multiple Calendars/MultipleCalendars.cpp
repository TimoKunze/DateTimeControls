// MultipleCalendars.cpp : main source file for Events.exe
//

#include "stdafx.h"
#include "resource.h"
#include "MainDlg.h"

CAppModule _Module;

int Run(LPTSTR /*lpstrCmdLine*/ = NULL, int nCmdShow = SW_SHOWDEFAULT)
{
	#ifdef _DEBUG
		// enable CRT memory leak detection & report
		_CrtSetDbgFlag(_CRTDBG_ALLOC_MEM_DF     // enable debug heap allocs & block type IDs (ie _CLIENT_BLOCK)
		    //| _CRTDBG_CHECK_CRT_DF              // check CRT allocations too
		    //| _CRTDBG_DELAY_FREE_MEM_DF       // keep freed blocks in list as _FREE_BLOCK type
		    | _CRTDBG_LEAK_CHECK_DF             // do leak report at exit (_CrtDumpMemoryLeaks)

		    // pick only one of these heap check frequencies
		    //| _CRTDBG_CHECK_ALWAYS_DF         // check heap on every alloc/free
		    //| _CRTDBG_CHECK_EVERY_16_DF
		    //| _CRTDBG_CHECK_EVERY_128_DF
		    //| _CRTDBG_CHECK_EVERY_1024_DF
		    | _CRTDBG_CHECK_DEFAULT_DF          // by default, no heap checks
		    );
		//_CrtSetBreakAlloc(209);               // break debugger on numbered allocation
		// get ID number from leak detector report of previous run
	#endif

	CMessageLoop theLoop;
	_Module.AddMessageLoop(&theLoop);

	CMainDlg dlgMain;

	if(dlgMain.Create(NULL) == NULL)
	{
		ATLTRACE(TEXT("Main dialog creation failed!\n"));
		return 0;
	}

	dlgMain.ShowWindow(nCmdShow);

	int nRet = theLoop.Run();

	_Module.RemoveMessageLoop();
	return nRet;
}

int WINAPI _tWinMain(HINSTANCE hInstance, HINSTANCE /*hPrevInstance*/, LPTSTR lpstrCmdLine, int nCmdShow)
{
	HRESULT hRes = CoInitialize(NULL);
// If you are running on NT 4.0 or higher you can use the following call instead to 
// make the EXE free threaded. This means that calls come in on a random RPC thread.
//	HRESULT hRes = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	ATLASSERT(SUCCEEDED(hRes));

	DWORD majorVersion = 0;
	DWORD minorVersion = 0;
	AtlGetCommCtrlVersion(&majorVersion, &minorVersion);
	if(majorVersion < 6 || (majorVersion == 6 && minorVersion < 10)) {
		MessageBox(NULL, TEXT("This sample requires version 6.10 or newer of comctl32.dll."), TEXT("Error"), MB_ICONERROR | MB_OK);
		return IDABORT;
	}

	// this resolves ATL window thunking problem when Microsoft Layer for Unicode (MSLU) is used
	DefWindowProc(NULL, 0, 0, 0L);

	AtlInitCommonControls(ICC_DATE_CLASSES);	// add flags to support other controls

	hRes = _Module.Init(NULL, hInstance, &LIBID_DTCtlsLibU);
	ATLASSERT(SUCCEEDED(hRes));

	AtlAxWinInit();

	int nRet = Run(lpstrCmdLine, nCmdShow);

	_Module.Term();
	CoUninitialize();

	return nRet;
}
