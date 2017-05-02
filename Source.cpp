#pragma comment(linker,"\"/manifestdependency:type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

#pragma comment(lib, "shlwapi")

#include <windows.h>
#include <shlwapi.h>
#include <shlobj.h>
#include <string>
#include "resource.h"

TCHAR szClassName[] = TEXT("Window");

VOID MyCreateFileFromResource(TCHAR *szResourceName, TCHAR *szResourceType, TCHAR *szResFileName)
{
	HRSRC hRs;
	HGLOBAL hMem;
	HANDLE hFile;
	LPBYTE lpByte;
	DWORD dwWritten, dwResSize;
	hRs = FindResource(0, szResourceName, szResourceType);
	dwResSize = SizeofResource(0, hRs);
	hMem = LoadResource(0, hRs);
	lpByte = (BYTE *)LockResource(hMem);
	hFile = CreateFile(szResFileName, GENERIC_WRITE, 0, NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);
	WriteFile(hFile, lpByte, dwResSize, &dwWritten, NULL);
	CloseHandle(hFile);
}

std::wstring Replace(std::wstring String1, std::wstring String2, std::wstring String3)
{
	std::wstring::size_type  Pos(String1.find(String2));
	while (Pos != std::wstring::npos)
	{
		String1.replace(Pos, String2.length(), String3);
		Pos = String1.find(String2, Pos + String3.length());
	}
	return String1;
}

BOOL ReplaceUTF8File(LPCTSTR pszFileName, LPCTSTR lpszProjectName)
{
	HANDLE hFile = CreateFile(pszFileName, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);
	if (hFile == INVALID_HANDLE_VALUE)
	{
		return FALSE;
	}
	else
	{
		SetFilePointer(hFile, 3, 0, FILE_BEGIN);
		DWORD dwSize;
		const DWORD dwFileSize = GetFileSize(hFile, 0);
		LPSTR lpszText = (LPSTR)GlobalAlloc(0, dwFileSize - 2 + 1);
		ReadFile(hFile, lpszText, dwFileSize - 2, &dwSize, 0);
		CloseHandle(hFile);
		lpszText[dwSize] = 0;
		DWORD len = MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, 0, 0);
		LPWSTR pwsz = (LPWSTR)GlobalAlloc(0, len * sizeof(WCHAR));
		MultiByteToWideChar(CP_UTF8, 0, lpszText, -1, pwsz, len);
		GlobalFree(lpszText);
		std::wstring str(pwsz);
		GlobalFree(pwsz);
		str = Replace(str, TEXT("XXXXX"), lpszProjectName);
		hFile = CreateFile(pszFileName, GENERIC_WRITE, 0, 0, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0);
		if (hFile == INVALID_HANDLE_VALUE)
		{
			return FALSE;
		}
		const BYTE BOM[3] = { 0xEF, 0xBB, 0xBF };
		WriteFile(hFile, BOM, 3, &dwSize, NULL);
		len = WideCharToMultiByte(CP_UTF8, 0, str.c_str(), -1, 0, 0, 0, 0);
		char*psz = (char*)GlobalAlloc(GPTR, len * sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, str.c_str(), -1, psz, len, 0, 0);
		WriteFile(hFile, psz, lstrlenA(psz), &dwSize, 0);
		GlobalFree(psz);
		CloseHandle(hFile);
	}
	return TRUE;
}

LRESULT CALLBACK WndProc(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
	static HWND hEdit;
	static HWND hButton;
	static HWND hOpenCheck;
	static HWND hBuildCheck;
	switch (msg)
	{
	case WM_CREATE:
		hEdit = CreateWindowEx(WS_EX_CLIENTEDGE, TEXT("EDIT"), 0, WS_VISIBLE | WS_CHILD | WS_TABSTOP | ES_AUTOHSCROLL, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hButton = CreateWindow(TEXT("BUTTON"), TEXT("テンプレート作成"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_DEFPUSHBUTTON, 0, 0, 0, 0, hWnd, (HMENU)IDOK, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		hBuildCheck = CreateWindow(TEXT("BUTTON"), TEXT("ビルドする(&B)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hBuildCheck, BM_SETCHECK, BST_CHECKED, 0);
		hOpenCheck = CreateWindow(TEXT("BUTTON"), TEXT("Visual Studio で開く(&O)"), WS_VISIBLE | WS_CHILD | WS_TABSTOP | BS_AUTOCHECKBOX, 0, 0, 0, 0, hWnd, 0, ((LPCREATESTRUCT)lParam)->hInstance, 0);
		SendMessage(hOpenCheck, BM_SETCHECK, BST_CHECKED, 0);
		SendMessage(hEdit, EM_SETCUEBANNER, 1, (LPARAM)TEXT("プロジェクト名の入力"));
		break;
	case WM_SIZE:
		MoveWindow(hEdit, 10, 10, LOWORD(lParam) - 20, 32, TRUE);
		MoveWindow(hButton, 10, 50, 256, 32, TRUE);
		MoveWindow(hBuildCheck, 10, 90, 256, 32, TRUE);
		MoveWindow(hOpenCheck, 10, 130, 256, 32, TRUE);
		break;
	case WM_COMMAND:
		if (LOWORD(wParam) == IDOK)
		{
			TCHAR szDirectory[MAX_PATH];
			if (SHGetSpecialFolderPath(HWND_DESKTOP, szDirectory, CSIDL_DESKTOPDIRECTORY, FALSE))
			{
				TCHAR szProjectName[MAX_PATH];
				GetWindowText(hEdit, szProjectName, MAX_PATH);
				PathAppend(szDirectory, szProjectName);
				if (CreateDirectory(szDirectory, 0))
				{
					TCHAR szOutputFilePath[MAX_PATH];
					lstrcpy(szOutputFilePath, szDirectory);
					PathAppend(szOutputFilePath, TEXT("Source.cpp"));
					MyCreateFileFromResource(MAKEINTRESOURCE(IDR_CPP1), TEXT("CPP"), szOutputFilePath);

					lstrcpy(szOutputFilePath, szDirectory);
					PathAppend(szOutputFilePath, szProjectName);
					lstrcat(szOutputFilePath, TEXT(".sln"));
					MyCreateFileFromResource(MAKEINTRESOURCE(IDR_SLN1), TEXT("SLN"), szOutputFilePath);
					ReplaceUTF8File(szOutputFilePath, szProjectName);

					lstrcpy(szOutputFilePath, szDirectory);
					PathAppend(szOutputFilePath, szProjectName);
					lstrcat(szOutputFilePath, TEXT(".vcxproj"));
					MyCreateFileFromResource(MAKEINTRESOURCE(IDR_VCXPROJ1), TEXT("VCXPROJ"), szOutputFilePath);
					ReplaceUTF8File(szOutputFilePath, szProjectName);

					lstrcat(szOutputFilePath, TEXT(".filters"));
					MyCreateFileFromResource(MAKEINTRESOURCE(IDR_FILTERS1), TEXT("FILTERS"), szOutputFilePath);

					TCHAR szSolutionFilePath[MAX_PATH];
					lstrcpy(szSolutionFilePath, szDirectory);
					PathAppend(szSolutionFilePath, szProjectName);
					lstrcat(szSolutionFilePath, TEXT(".sln"));

					TCHAR szVisualStudioPath[MAX_PATH];
					HKEY hKey;
					DWORD dwPosition;
					DWORD dwType = REG_SZ;
					DWORD dwByte = MAX_PATH * sizeof(TCHAR);
					if (SendMessage(hBuildCheck, BM_GETCHECK, 0, 0))
					{
						if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,
#ifdef _WIN64
							TEXT("SOFTWARE\\WOW6432Node\\Microsoft\\MSBuild\\14.0"),
#else
							TEXT("SOFTWARE\\Microsoft\\VisualStudio\\14.0\\Setup\\vs"),
#endif
							0, 0, 0, KEY_READ, 0, &hKey, &dwPosition) == ERROR_SUCCESS)
						{
							if (RegQueryValueEx(hKey, TEXT("MSBuildOverrideTasksPath"), NULL, &dwType, (BYTE *)szVisualStudioPath, &dwByte) == ERROR_SUCCESS)
							{
								PathAppend(szVisualStudioPath, TEXT("MSBuild.exe"));
								TCHAR szParameters[1024];
								wsprintf(szParameters, TEXT("%s /t:Build /p:Configuration=Release;Platform=x86"), szSolutionFilePath);
								ShellExecute(NULL, TEXT("open"), szVisualStudioPath, szParameters, NULL, SW_SHOWNORMAL);
							}
							RegCloseKey(hKey);
						}
					}
					if (SendMessage(hOpenCheck, BM_GETCHECK, 0, 0))
					{
						if (RegCreateKeyEx(HKEY_LOCAL_MACHINE,
#ifdef _WIN64
							TEXT("SOFTWARE\\WOW6432Node\\Microsoft\\VisualStudio\\14.0\\Setup\\vs"),
#else
							TEXT("SOFTWARE\\Microsoft\\VisualStudio\\14.0\\Setup\\vs"),
#endif
							0, 0, 0, KEY_READ, 0, &hKey, &dwPosition) == ERROR_SUCCESS)
						{
							if (RegQueryValueEx(hKey, TEXT("EnvironmentPath"), NULL, &dwType, (BYTE *)szVisualStudioPath, &dwByte) == ERROR_SUCCESS)
							{
								ShellExecute(NULL, TEXT("open"), szVisualStudioPath, szSolutionFilePath, NULL, SW_SHOWNORMAL);
							}
							RegCloseKey(hKey);
						}
					}
				}
			}
		}
		break;
	case WM_CLOSE:
		DestroyWindow(hWnd);
		break;
	case WM_DESTROY:
		PostQuitMessage(0);
		break;
	default:
		return(DefDlgProc(hWnd, msg, wParam, lParam));
	}
	return 0;
}

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPreInst, LPSTR pCmdLine, int nCmdShow)
{
	MSG msg;
	WNDCLASS wndclass = {
		CS_HREDRAW | CS_VREDRAW,
		WndProc,
		0,
		DLGWINDOWEXTRA,
		hInstance,
		LoadIcon(hInstance, MAKEINTRESOURCE(IDI_ICON1)),
		LoadCursor(0,IDC_ARROW),
		0,
		0,
		szClassName
	};
	RegisterClass(&wndclass);
	HWND hWnd = CreateWindow(
		szClassName,
		TEXT("プロジェクト設定されたVisual Studio 2015のプロジェクトのテンプレートをデスクトップに出力"),
		WS_OVERLAPPEDWINDOW,
		CW_USEDEFAULT,
		0,
		CW_USEDEFAULT,
		0,
		0,
		0,
		hInstance,
		0
	);
	ShowWindow(hWnd, SW_SHOWDEFAULT);
	UpdateWindow(hWnd);
	while (GetMessage(&msg, 0, 0, 0))
	{
		if (!IsDialogMessage(hWnd, &msg))
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}
	return (int)msg.wParam;
}
