#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
#include <uxtheme.h>
#include <locale.h>
#include <commctrl.h>
#include <richedit.h>
#include <tchar.h>
#include <stdlib.h>
#include <stdio.h>
#include <ctype.h>
#include <sys/types.h>
#include <math.h>
#include "parson.h"

#define LVS_EX_AUTOSIZECOLUMNS 0x10000000

#define WMU_UPDATE_GRID        WM_USER + 1
#define WMU_UPDATE_RESULTSET   WM_USER + 2
#define WMU_UPDATE_FILTER_SIZE WM_USER + 3
#define WMU_SET_HEADER_FILTERS WM_USER + 4
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 5
#define WMU_RESET_CACHE        WM_USER + 6
#define WMU_SET_FONT           WM_USER + 7
#define WMU_SET_THEME          WM_USER + 8
#define WMU_UPDATE_TEXT        WM_USER + 9
#define WMU_UPDATE_HIGHLIGHT   WM_USER + 10

#define IDC_MAIN               100
#define IDC_TREE               101
#define IDC_TAB                102
#define IDC_GRID               103
#define IDC_TEXT               104
#define IDC_STATUSBAR          105
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROW           5001
#define IDM_FILTER_ROW         5002
#define IDM_DARK_THEME         5003
#define IDM_COPY_TEXT          5010
#define IDM_SELECTALL          5011

#define SB_VERSION             0
#define SB_CHILD_NO            1
#define SB_TYPE                2
#define SB_ROW_COUNT           3
#define SB_CURRENT_ROW         4
#define SB_AUXILIARY           5

#define SPLITTER_WIDTH         5
#define MAX_LENGTH             4096
#define MAX_COLUMN_LENGTH      2000
#define APP_NAME               TEXT("jsontab")
#define APP_VERSION            TEXT("0.9.3")

#define LCS_FINDFIRST          1
#define LCS_MATCHCASE          2
#define LCS_WHOLEWORDS         4
#define LCS_BACKWARDS          8

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
char* json_value_get_as_string(const JSON_Value* val);
BOOL addNode(HWND hTreeWnd, HTREEITEM hParentItem, JSON_Value* val);
int highlightBlock(HWND hWnd, TCHAR* text, int start);
HWND getMainWindow(HWND hWnd);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
TCHAR* getStoredString(TCHAR* name, TCHAR* defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords);
void setClipboardText(const TCHAR* text);
BOOL isNumber(TCHAR* val);
BOOL isUtf8(const char * string);
void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums);
HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam);
int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen);
LPARAM TreeView_GetItemParam(HWND hTreeWnd, HTREEITEM hItem);
int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);
void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
	if (ul_reason_for_call == DLL_PROCESS_ATTACH && iniPath[0] == 0) {
		TCHAR path[MAX_PATH + 1] = {0};
		GetModuleFileName(hModule, path, MAX_PATH);
		TCHAR* dot = _tcsrchr(path, TEXT('.'));
		_tcsncpy(dot, TEXT(".ini"), 5);
		if (_taccess(path, 0) == 0)
			_tcscpy(iniPath, path);	
	}
	return TRUE;
}

void __stdcall ListGetDetectString(char* DetectString, int maxlen) {
	snprintf(DetectString, maxlen, "MULTIMEDIA & ext=\"JSON\"");
}

void __stdcall ListSetDefaultParams(ListDefaultParamStruct* dps) {
	if (iniPath[0] == 0) {
		DWORD size = MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, NULL, 0);
		MultiByteToWideChar(CP_ACP, 0, dps->DefaultIniName, -1, iniPath, size);
	}
}

int __stdcall ListSearchTextW(HWND hWnd, TCHAR* searchString, int searchParameter) {
	HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);

	if (TabCtrl_GetCurSel(hTabWnd) == 1) { 
		HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
		DWORD len = _tcslen(searchString);
		int spos = SendMessage(hTextWnd, EM_GETSEL, 0, 0);
		int mode = 0;	
		FINDTEXTEXW ft = {{HIWORD(spos), -1}, searchString, {0, 0}};
		if (searchParameter & LCS_MATCHCASE)
			mode |= FR_MATCHCASE;
		if (searchParameter & LCS_WHOLEWORDS)
			mode |= FR_WHOLEWORD;
		if (!(searchParameter & LCS_BACKWARDS)) 
			mode |= FR_DOWN;
		else 
			ft.chrg.cpMin  = ft.chrg.cpMin > len ? ft.chrg.cpMin - len : ft.chrg.cpMin;
	
		int pos = SendMessage(hTextWnd, EM_FINDTEXTEXW, mode, (LPARAM)&ft);	
		if (pos != -1) 
			SendMessage(hTextWnd, EM_SETSEL, pos, pos + len);
		else	
			MessageBeep(0);
		SetFocus(hTextWnd);		
	} else {
		HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
		HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);	
		
		TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
		int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
		int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			
		BOOL isBackward = searchParameter & LCS_BACKWARDS;
		BOOL isMatchCase = searchParameter & LCS_MATCHCASE;
		BOOL isWholeWords = searchParameter & LCS_WHOLEWORDS;	
			
		int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
		if (rowNo == -1)
			rowNo = isBackward ? rowCount - 1 : 0;
			
		int* pColNo = (int*)GetProp(hWnd, TEXT("SEARCHCOLNO"));
	
		int pos = -1;
		do {
			int colNo = *pColNo;
			for (; (pos == -1) && colNo < colCount; colNo++) 
				pos = findString(cache[rowNo][colNo], searchString, isMatchCase, isWholeWords);
			*pColNo = pos != -1 ? colNo - 1 : 0;
			rowNo += (pos == -1) && (isBackward ? -1 : 1); 	
		} while ((pos == -1) && (isBackward ? rowNo > 0 : rowNo < rowCount - 1));
	
		TCHAR buf[256] = {0};
		if (pos != -1) {
			ListView_EnsureVisible(hGridWnd, rowNo, FALSE);
			ListView_SetItemState(hGridWnd, rowNo, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);
			
			TCHAR colName[MAX_COLUMN_LENGTH + 1];
			Header_GetItemText(ListView_GetHeader(hGridWnd), *pColNo, colName, MAX_COLUMN_LENGTH);
			TCHAR* val = cache[rowNo][*pColNo];
			int len = _tcslen(searchString);
			_sntprintf(buf, 255, TEXT("%ls: %ls%.*ls%ls"), colName, 
				pos >= 10 ? TEXT("...") : TEXT(""), 
				len + (pos < 10 ? pos : 10) + 10, pos < 10 ? val : val + pos - 10,
				_tcslen(val + pos + len) > 10 ? TEXT("...") : TEXT(""));
			*pColNo += 1;
		} else { 
			MessageBox(hWnd, searchString, TEXT("Not found:"), MB_OK);
		}
		SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)buf);	
		SetFocus(hGridWnd);	
	}

	return 0;
}

int __stdcall ListSearchText(HWND hWnd, char* searchString, int searchParameter) {
	DWORD len = MultiByteToWideChar(CP_ACP, 0, searchString, -1, NULL, 0);
	TCHAR* searchString16 = (TCHAR*)calloc (len, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, searchString, -1, searchString16, len);
	int rc = ListSearchTextW(hWnd, searchString16, searchParameter);
	free(searchString16);
	return rc;
}

HWND APIENTRY ListLoadW (HWND hListerWnd, TCHAR* fileToLoad, int showFlags) {
	int maxFileSize = getStoredValue(TEXT("max-file-size"), 1000000);
	struct _stat st = {0};
	if (_tstat(fileToLoad, &st) != 0 || maxFileSize > 0 && st.st_size > maxFileSize)
		return 0;

	char* data = calloc(st.st_size + 1, sizeof(char));
	FILE *f = _tfopen(fileToLoad, TEXT("rb"));
	fread(data, sizeof(char), st.st_size, f);
	fclose(f);

	json_set_escape_slashes(0);
	JSON_Value* json = isUtf8(data) ? json_parse_string(data) : 0;
	// ANSI
	if (!json) {
		DWORD len = MultiByteToWideChar(CP_ACP, 0, data, -1, NULL, 0);
		TCHAR* data16 = (TCHAR*)calloc (len, sizeof (TCHAR));
		MultiByteToWideChar(CP_ACP, 0, data, -1, data16, len);

		char* data8 = utf16to8(data16);
		json = json_parse_string(data8);
		free(data8);
		free(data16);
	}

	// UTF16LE
	if (!json) {
		char* data8 = utf16to8((TCHAR*)data);
		json = json_parse_string(data8);
		free(data8);
	}

	// UTF16BE
	if (!json) {
		for (int i = 0; i < st.st_size/2; i++) {
			int c = data[2 * i];
			data[2 * i] = data[2 * i + 1];
			data[2 * i + 1] = c;
		}

		char* data8 = utf16to8((TCHAR*)data);
		json = json_parse_string(data8);
		free(data8);
	}

	free(data);
	if (!json)
		return 0;

	INITCOMMONCONTROLSEX icex;
	icex.dwSize = sizeof(icex);
	icex.dwICC = ICC_LISTVIEW_CLASSES | ICC_TREEVIEW_CLASSES;
	InitCommonControlsEx(&icex);
	LoadLibrary(TEXT("msftedit.dll"));
	
	setlocale(LC_CTYPE, "");

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindowEx(WS_EX_CONTROLPARENT, WC_STATIC, APP_NAME, WS_CHILD | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("FILTERROW"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("JSON"), json);
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SEARCHCOLNO"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SPLITTERPOSITION"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTFAMILY"), getStoredString(TEXT("font"), TEXT("Arial")));	
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERALIGN"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH"), calloc(1, sizeof(int)));	
	
	SetProp(hMainWnd, TEXT("DARKTHEME"), calloc(1, sizeof(int)));			
	SetProp(hMainWnd, TEXT("TEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("BACKCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERTEXTCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FILTERBACKCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("SPLITTERCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("JSONTEXTCOLOR"), calloc(1, sizeof(int)));		
	SetProp(hMainWnd, TEXT("JSONKEYCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("JSONSTRINGCOLOR"), calloc(1, sizeof(int)));	
	SetProp(hMainWnd, TEXT("JSONBOOLEANCOLOR"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("JSONNULLCOLOR"), calloc(1, sizeof(int)));			

	*(int*)GetProp(hMainWnd, TEXT("SPLITTERPOSITION")) = getStoredValue(TEXT("splitter-position"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);
	*(int*)GetProp(hMainWnd, TEXT("MAXHIGHLIGHTLENGTH")) = getStoredValue(TEXT("max-highlight-length"), 64000);	
	*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) = getStoredValue(TEXT("filter-row"), 1);
	*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) = getStoredValue(TEXT("dark-theme"), 0);
	*(int*)GetProp(hMainWnd, TEXT("FILTERALIGN")) = getStoredValue(TEXT("filter-align"), 0);		

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	HDC hDC = GetDC(hMainWnd);
	float z = GetDeviceCaps(hDC, LOGPIXELSX) / 96.0; // 96 = 100%, 120 = 125%, 144 = 150%
	ReleaseDC(hMainWnd, hDC);	
	int sizes[6] = {35 * z, 140 * z, 200 * z, 400 * z, 500 * z, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 6, (LPARAM)&sizes);

	HWND hTreeWnd = CreateWindow(WC_TREEVIEW, NULL, WS_VISIBLE | WS_CHILD | TVS_HASBUTTONS | TVS_HASLINES | TVS_SHOWSELALWAYS | WS_TABSTOP,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TREE, GetModuleHandle(0), NULL);
	SetProp(hTreeWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTreeWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	HWND hTabWnd = CreateWindow(WC_TABCONTROL, NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | WS_TABSTOP, 100, 100, 100, 100,
		hMainWnd, (HMENU)IDC_TAB, GetModuleHandle(0), NULL);
	SetProp(hTabWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTabWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));	

	TCITEM tci;
	tci.mask = TCIF_TEXT | TCIF_IMAGE;
	tci.iImage = -1;
	tci.pszText = TEXT("Grid");
	tci.cchTextMax = 5;
	TabCtrl_InsertItem(hTabWnd, 0, &tci);

	tci.pszText = TEXT("Text");
	tci.cchTextMax = 5;
	TabCtrl_InsertItem(hTabWnd, 1, &tci);

	int tabNo = getStoredValue(TEXT("tab-no"), 0);
	HWND hGridWnd = CreateWindow(WC_LISTVIEW, NULL, (tabNo == 0 ? WS_VISIBLE : 0) | WS_CHILD  | LVS_REPORT | LVS_SHOWSELALWAYS | LVS_SINGLESEL | LVS_OWNERDATA | WS_TABSTOP,
		205, 0, 100, 100, hTabWnd, (HMENU)IDC_GRID, GetModuleHandle(0), NULL);
	ListView_SetExtendedListViewStyle(hGridWnd, LVS_EX_FULLROWSELECT | LVS_EX_GRIDLINES | LVS_EX_LABELTIP);
	SetProp(hGridWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hGridWnd, GWLP_WNDPROC, (LONG_PTR)cbHotKey));

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);
	SetWindowTheme(hHeader, TEXT(" "), TEXT(" "));
	SetProp(hHeader, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hHeader, GWLP_WNDPROC, (LONG_PTR)cbNewHeader));

	HWND hTextWnd = CreateWindowEx(0, TEXT("RICHEDIT50W"), NULL, (tabNo == 1 ? WS_VISIBLE : 0) | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP | ES_NOHIDESEL | ES_READONLY,
		0, 0, 100, 100, hTabWnd, (HMENU)IDC_TEXT, GetModuleHandle(0),  NULL);
	SetProp(hTextWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTextWnd, GWLP_WNDPROC, (LONG_PTR)cbNewText));
	TabCtrl_SetCurSel(hTabWnd, tabNo);

	HMENU hDataMenu = CreatePopupMenu();
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_ROW, TEXT("Copy row"));
	AppendMenu(hDataMenu, MF_STRING, 0, NULL);	
	AppendMenu(hDataMenu, (*(int*)GetProp(hMainWnd, TEXT("FILTERROW")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_FILTER_ROW, TEXT("Filters"));		
	AppendMenu(hDataMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));					
	SetProp(hMainWnd, TEXT("DATAMENU"), hDataMenu);

	HMENU hTextMenu = CreatePopupMenu();
	AppendMenu(hTextMenu, MF_STRING, IDM_COPY_TEXT, TEXT("Copy"));
	AppendMenu(hTextMenu, MF_STRING, IDM_SELECTALL, TEXT("Select all"));
	AppendMenu(hTextMenu, MF_STRING, 0, NULL);	
	AppendMenu(hTextMenu, (*(int*)GetProp(hMainWnd, TEXT("DARKTHEME")) != 0 ? MF_CHECKED : 0) | MF_STRING, IDM_DARK_THEME, TEXT("Dark theme"));					
	SetProp(hMainWnd, TEXT("TEXTMENU"), hTextMenu);

	TCHAR buf[255];
	_sntprintf(buf, 32, TEXT(" %ls"), APP_VERSION);	
	SendMessage(hStatusWnd, SB_SETTEXT, SB_VERSION, (LPARAM)buf);	

	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WMU_SET_THEME, 0, 0);
	HTREEITEM hRoot = TreeView_AddItem(hTreeWnd, TEXT("<<root>>"), TVI_ROOT, (LPARAM)json);
	addNode(hTreeWnd, hRoot, json);
	TreeView_Expand(hTreeWnd, hRoot, TVE_EXPAND);
	TreeView_Select(hTreeWnd, hRoot, TVGN_CARET);
	ShowWindow(hMainWnd, SW_SHOW);
	SetFocus(hTreeWnd);

	return hMainWnd;
}

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* fileToLoadW = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, fileToLoadW, size);
	HWND hWnd = ListLoadW(hListerWnd, fileToLoadW, showFlags);
	free(fileToLoadW);
	return hWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-position"), *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("tab-no"), TabCtrl_GetCurSel(GetDlgItem(hWnd, IDC_TAB)));
	setStoredValue(TEXT("filter-row"), *(int*)GetProp(hWnd, TEXT("FILTERROW")));	
	setStoredValue(TEXT("dark-theme"), *(int*)GetProp(hWnd, TEXT("DARKTHEME")));	

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);	
	json_value_free((JSON_Value*)GetProp(hWnd, TEXT("JSON")));
	free((int*)GetProp(hWnd, TEXT("FILTERROW")));
	free((int*)GetProp(hWnd, TEXT("DARKTHEME")));	
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLNO")));
	free((int*)GetProp(hWnd, TEXT("SEARCHCOLNO")));	
	free((int*)GetProp(hWnd, TEXT("SPLITTERPOSITION")));
	free((TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
	free((int*)GetProp(hWnd, TEXT("FILTERALIGN")));

	free((int*)GetProp(hWnd, TEXT("TEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("BACKCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")));
	free((int*)GetProp(hWnd, TEXT("JSONTEXTCOLOR")));
	free((int*)GetProp(hWnd, TEXT("JSONKEYCOLOR")));
	free((int*)GetProp(hWnd, TEXT("JSONSTRINGCOLOR")));	
	free((int*)GetProp(hWnd, TEXT("JSONBOOLEANCOLOR")));
	free((int*)GetProp(hWnd, TEXT("JSONNULLCOLOR")));			

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));	
	DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));
	DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("DATAMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("TEXTMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("FILTERROW"));	
	RemoveProp(hWnd, TEXT("DARKTHEME"));		
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLNO"));
	RemoveProp(hWnd, TEXT("SEARCHCOLNO"));	
	RemoveProp(hWnd, TEXT("JSON"));
	RemoveProp(hWnd, TEXT("SPLITTERPOSITION"));
	RemoveProp(hWnd, TEXT("FILTERALIGN"));

	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("FONTFAMILY"));
	RemoveProp(hWnd, TEXT("FONTSIZE"));
	RemoveProp(hWnd, TEXT("TEXTCOLOR"));
	RemoveProp(hWnd, TEXT("BACKCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERTEXTCOLOR"));
	RemoveProp(hWnd, TEXT("FILTERBACKCOLOR"));
	RemoveProp(hWnd, TEXT("SPLITTERCOLOR"));
	RemoveProp(hWnd, TEXT("JSONTEXTCOLOR"));	
	RemoveProp(hWnd, TEXT("JSONKEYCOLOR"));		
	RemoveProp(hWnd, TEXT("JSONSTRINGCOLOR"));			
	RemoveProp(hWnd, TEXT("JSONBOOLEANCOLOR"));		
	RemoveProp(hWnd, TEXT("JSONNULLCOLOR"));		
	RemoveProp(hWnd, TEXT("BACKBRUSH"));
	RemoveProp(hWnd, TEXT("FILTERBACKBRUSH"));		
	RemoveProp(hWnd, TEXT("SPLITTERBRUSH"));
	RemoveProp(hWnd, TEXT("DATAMENU"));
	RemoveProp(hWnd, TEXT("TEXTMENU"));

	DestroyWindow(hWnd);
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			GetClientRect(hWnd, &rc);
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			SetWindowPos(hTreeWnd, 0, 0, 0, splitterW, rc.bottom - statusH, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hTabWnd, 0, splitterW + SPLITTER_WIDTH, 0, rc.right - splitterW - SPLITTER_WIDTH, rc.bottom - statusH, SWP_NOZORDER);

			RECT rc2;
			GetClientRect(hTabWnd, &rc);
			TabCtrl_GetItemRect(hTabWnd, 0, &rc2);
			SetWindowPos(hTextWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - SPLITTER_WIDTH, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - SPLITTER_WIDTH, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
		}
		break;

		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			rc.right = rc.left + SPLITTER_WIDTH;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("SPLITTERBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;
		
		case WM_SETCURSOR: {
			SetCursor(LoadCursor(0, GetProp(hWnd, TEXT("ISMOUSEHOVER")) ? IDC_SIZEWE : IDC_ARROW));
			return TRUE;
		}
		break;		

		case WM_LBUTTONDOWN: {
			int x = GET_X_LPARAM(lParam);
			int pos = *(int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			if (x >= pos && x <= pos + SPLITTER_WIDTH) {
				SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
				SetCapture(hWnd);
			}
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;

		case WM_MOUSEMOVE: {
			DWORD x = GET_X_LPARAM(lParam);
			int* pPos = (int*)GetProp(hWnd, TEXT("SPLITTERPOSITION"));
			
			if (!GetProp(hWnd, TEXT("ISMOUSEHOVER")) && *pPos <= x && x <= *pPos + SPLITTER_WIDTH) {
				TRACKMOUSEEVENT tme = {sizeof(TRACKMOUSEEVENT), TME_LEAVE, hWnd, 0};
				TrackMouseEvent(&tme);	
				SetProp(hWnd, TEXT("ISMOUSEHOVER"), (HANDLE)1);
			}
			
			if (GetProp(hWnd, TEXT("ISMOUSEDOWN")) && x > 0 && x < 32000) {
				*pPos = x;
				SendMessage(hWnd, WM_SIZE, 0, 0);
			}
		}
		break;
		
		case WM_MOUSELEAVE: {
			SetProp(hWnd, TEXT("ISMOUSEHOVER"), 0);
		}
		break;

		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
			}
		}
		break;
		
		case WM_KEYDOWN: {
			if (wParam == VK_ESCAPE)
				SendMessage(GetParent(hWnd), WM_CLOSE, 0, 0);

			if (wParam == VK_TAB) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				BOOL isBackward = HIWORD(GetKeyState(VK_CONTROL));
				no += isBackward ? -1 : 1;
				SetFocus(wnds[no] && no >= 0 ? wnds[no] : (isBackward ? wnds[cnt - 1] : wnds[0]));
				return TRUE;
			}
			
			if (wParam == VK_F1) {
				ShellExecute(0, 0, TEXT("https://github.com/little-brother/jsontab-wlx/wiki"), 0, 0 , SW_SHOW);
				return TRUE;
			}
		}
		break;

		case WM_CONTEXTMENU: {
			POINT p = {LOWORD(lParam), HIWORD(lParam)};
			if (GetDlgCtrlID(WindowFromPoint(p)) == IDC_TEXT)
				TrackPopupMenu(GetProp(hWnd, TEXT("TEXTMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
		}
		break;

		case WM_COMMAND: {
			WORD cmd = LOWORD(wParam);
			if (cmd == IDM_COPY_CELL || cmd == IDM_COPY_ROW) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
				HWND hHeader = ListView_GetHeader(hGridWnd);
				int rowNo = ListView_GetNextItem(hGridWnd, -1, LVNI_SELECTED);
				if (rowNo != -1) {
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
					int rowCount = *(int*)GetProp(hWnd, TEXT("ROWCOUNT"));
					if (!resultset || rowNo >= rowCount)
						return 0;

					int colCount = Header_GetItemCount(hHeader);

					int startNo = cmd == IDM_COPY_CELL ? *(int*)GetProp(hWnd, TEXT("COLNO")) : 0;
					int endNo = cmd == IDM_COPY_CELL ? startNo + 1 : colCount;
					if (startNo > colCount || endNo > colCount)
						return 0;

					int len = 0;
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						len += _tcslen(cache[_rowNo][colNo]) + 1 /* column delimiter: TAB */;
					}

					TCHAR buf[len + 1];
					buf[0] = 0;
					for (int colNo = startNo; colNo < endNo; colNo++) {
						int _rowNo = resultset[rowNo];
						_tcscat(buf, cache[_rowNo][colNo]);
						if (colNo != endNo - 1)
							_tcscat(buf, TEXT("\t"));
					}

					setClipboardText(buf);
				}
			}

			if (cmd == IDM_COPY_TEXT || cmd == IDM_SELECTALL) {
				HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
				HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
				SendMessage(hTextWnd, cmd == IDM_COPY_TEXT ? WM_COPY : EM_SETSEL, 0, -1);
			}
			
			if (cmd == IDM_FILTER_ROW || cmd == IDM_DARK_THEME) {
				HMENU hDataMenu = (HMENU)GetProp(hWnd, TEXT("DATAMENU"));
				HMENU hTextMenu = (HMENU)GetProp(hWnd, TEXT("TEXTMENU"));				
				int* pOpt = (int*)GetProp(hWnd, cmd == IDM_FILTER_ROW ? TEXT("FILTERROW") : TEXT("DARKTHEME"));
				*pOpt = (*pOpt + 1) % 2;
				Menu_SetItemState(hDataMenu, cmd, *pOpt ? MFS_CHECKED : 0);
				Menu_SetItemState(hTextMenu, cmd, *pOpt ? MFS_CHECKED : 0);				
				
				UINT msg = cmd == IDM_FILTER_ROW ? WMU_SET_HEADER_FILTERS : WMU_SET_THEME;
				SendMessage(hWnd, msg, 0, 0);				
			}			
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_TAB && pHdr->code == TCN_SELCHANGE) {
				HWND hTabWnd = pHdr->hwndFrom;
				BOOL isText = TabCtrl_GetCurSel(hTabWnd) == 1;
				ShowWindow(GetDlgItem(hTabWnd, IDC_GRID), isText ? SW_HIDE : SW_SHOW);
				ShowWindow(GetDlgItem(hTabWnd, IDC_TEXT), isText ? SW_SHOW : SW_HIDE);
				if (!isText)
					SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			}

			if (pHdr->idFrom == IDC_TREE && pHdr->code == TVN_SELCHANGED) {
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);
				SendMessage(hWnd, WMU_UPDATE_TEXT, 0, 0);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_GETDISPINFO) {
				LV_DISPINFO* pDispInfo = (LV_DISPINFO*)lParam;
				LV_ITEM* pItem = &(pDispInfo)->item;

				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
				if(resultset && pItem->mask & LVIF_TEXT) {
					int rowNo = resultset[pItem->iItem];
					pItem->pszText = cache[rowNo][pItem->iSubItem];
				}
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == LVN_COLUMNCLICK) {
				NMLISTVIEW* pLV = (NMLISTVIEW*)lParam;
				int colNo = pLV->iSubItem + 1;
				int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));
				int orderBy = *pOrderBy;
				*pOrderBy = colNo == orderBy || colNo == -orderBy ? -orderBy : colNo;
				SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)NM_RCLICK) {
				NMITEMACTIVATE* ia = (LPNMITEMACTIVATE) lParam;
				POINT p;
				GetCursorPos(&p);
				*(int*)GetProp(hWnd, TEXT("COLNO")) = ia->iSubItem;
				TrackPopupMenu(GetProp(hWnd, TEXT("DATAMENU")), TPM_RIGHTBUTTON | TPM_TOPALIGN | TPM_LEFTALIGN, p.x, p.y, 0, hWnd, NULL);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_ITEMCHANGED) {
				HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);

				TCHAR buf[255] = {0};
				int pos = ListView_GetNextItem(pHdr->hwndFrom, -1, LVNI_SELECTED);
				if (pos != -1)
					_sntprintf(buf, 255, TEXT(" %i"), pos + 1);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_CURRENT_ROW, (LPARAM)buf);
				SendMessage(hStatusWnd, SB_SETTEXT, SB_AUXILIARY, (LPARAM)0);
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43 && HIWORD(GetKeyState(VK_CONTROL))) // Ctrl + C
					SendMessage(hWnd, WM_COMMAND, IDM_COPY_ROW, 0);
			}

			if (pHdr->code == HDN_ITEMCHANGED && pHdr->hwndFrom == ListView_GetHeader(GetDlgItem(GetDlgItem(hWnd, IDC_TAB), IDC_GRID)))
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_GRID: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			int filterAlign = *(int*)GetProp(hWnd, TEXT("FILTERALIGN"));

			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
			JSON_Value* val = (JSON_Value*)TreeView_GetItemParam(hTreeWnd, hItem);
			JSON_Value_Type type = json_value_get_type(val);

			HWND hHeader = ListView_GetHeader(hGridWnd);
			SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
			*(int*)GetProp(hWnd, TEXT("ORDERBY")) = 0;

			int colCount = Header_GetItemCount(hHeader);
			for (int colNo = 0; colNo < colCount; colNo++)
				DestroyWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo));

			for (int colNo = 0; colNo < colCount; colNo++)
				ListView_DeleteColumn(hGridWnd, colCount - colNo - 1);

			// Columns
			if (type == JSONArray) {
				JSON_Array* arr = json_value_get_array(val);
				if (json_array_get_count(arr) > 0) {
					JSON_Value* val1 = json_array_get_value(arr, 0);
					if (json_value_get_type(val1) == JSONObject) {
						JSON_Object* obj = json_value_get_object(val1);
						for (int itemNo = 0; itemNo < json_object_get_count(obj); itemNo++) {
							TCHAR* itemName = utf8to16(json_object_get_name(obj, itemNo));
							ListView_AddColumn(hGridWnd, itemName);
							free(itemName);
						}
					} else {
						ListView_AddColumn(hGridWnd, TEXT("Element"));
					}
				} else {
					ListView_AddColumn(hGridWnd, TEXT("Element"));
				}
			} else if (type == JSONObject) {
				ListView_AddColumn(hGridWnd, TEXT("Attribute"));
				ListView_AddColumn(hGridWnd, TEXT("Value"));
			} else {
				ListView_AddColumn(hGridWnd, TEXT("Value"));
			}

			colCount = Header_GetItemCount(hHeader);
			int align = filterAlign == -1 ? ES_LEFT : filterAlign == 1 ? ES_RIGHT : ES_CENTER;
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, align | ES_AUTOHSCROLL | WS_CHILD | WS_TABSTOP | WS_BORDER,
					0, 0, 0, 0, hHeader, (HMENU)(INT_PTR)(IDC_HEADER_EDIT + colNo), GetModuleHandle(0), NULL);
				SendMessage(hEdit, WM_SETFONT, (LPARAM)GetProp(hWnd, TEXT("FONT")), TRUE);
				SetProp(hEdit, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hEdit, GWLP_WNDPROC, (LONG_PTR)cbNewFilterEdit));
			}

			// Cache data
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int rowCount = 0;
			if (type == JSONArray) {
				JSON_Array* arr = json_value_get_array(val);
				rowCount = json_array_get_count(arr);

				if (rowCount > 0) {
					SetProp(hWnd, TEXT("CACHE"), calloc(rowCount, sizeof(TCHAR*)));
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));

					for (int rowNo = 0; rowNo < rowCount; rowNo++) {
						cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));

						JSON_Value* valRow = json_array_get_value(arr, rowNo);
						if (json_value_get_type(valRow) == JSONObject) {
							JSON_Object* obj = json_value_get_object(valRow);
							for (int colNo = 0; colNo < colCount; colNo++) {
								TCHAR colName16[MAX_LENGTH + 1];
								Header_GetItemText(hHeader, colNo, colName16, MAX_LENGTH);
								char* colName8 = utf16to8(colName16);
								JSON_Value* valValue = json_object_get_value(obj, colName8);
								free(colName8);

								if (valValue) {
									char* value8 = json_value_get_as_string(valValue);
									cache[rowNo][colNo] = utf8to16(value8);
									free(value8);
								} else {
									cache[rowNo][colNo] = calloc(1, sizeof(TCHAR));
								}
							}
						} else {
							char* value8 = json_value_get_as_string(valRow);
							cache[rowNo][0] = utf8to16(value8);
							free(value8);
						}
					}
				}
			} else if (type == JSONObject) {
				JSON_Object* obj = json_value_get_object(val);
				rowCount = json_object_get_count(obj);
				if (rowCount > 0) {
					SetProp(hWnd, TEXT("CACHE"), calloc(rowCount, sizeof(TCHAR*)));
					TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
					for (int rowNo = 0; rowNo < rowCount; rowNo++) {
						cache[rowNo] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));

						cache[rowNo][0] = utf8to16(json_object_get_name(obj, rowNo));
						JSON_Value* valValue = json_object_get_value_at(obj, rowNo);
						char* value8 = json_value_get_as_string(valValue);
						cache[rowNo][1] = utf8to16(value8);
						free(value8);
					}
				}
			} else {
				rowCount = 1;
				SetProp(hWnd, TEXT("CACHE"), calloc(rowCount, sizeof(TCHAR*)));
				TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
				cache[0] = (TCHAR**)calloc (colCount, sizeof (TCHAR*));
				char* value8 = json_value_get_as_string(val);
				cache[0][0] = utf8to16(value8);
				free(value8);
			}

			TCHAR buf[255];
			TCHAR *TYPES[10] = {TEXT(""), TEXT("NULL"), TEXT("STRING"), TEXT("NUMBER"), TEXT("OBJECT"), TEXT("ARRAY"), TEXT("BOOLEAN")};
			_sntprintf(buf, 255, TEXT(" %ls"), TYPES[type]);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_TYPE, (LPARAM)buf);

			int pos = 1;
			while((hItem = TreeView_GetPrevSibling(hTreeWnd, hItem)))
				pos++;
			_sntprintf(buf, 255, TEXT(" %i"), pos);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_CHILD_NO, (LPARAM)buf);

			*pTotalRowCount = rowCount;
			SendMessage(hWnd, WMU_UPDATE_RESULTSET, 0, 0);
			SendMessage(hWnd, WMU_SET_HEADER_FILTERS, 0, 0);
			SendMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

		case WMU_UPDATE_RESULTSET: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			HWND hHeader = ListView_GetHeader(hGridWnd);

			ListView_SetItemCount(hGridWnd, 0);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			int* pOrderBy = (int*)GetProp(hWnd, TEXT("ORDERBY"));

			if (!cache)
				return 1;
				
			if (*pTotalRowCount == 0)	
				return 1;
				
			int colCount = Header_GetItemCount(hHeader);
			BOOL* bResultset = (BOOL*)calloc(*pTotalRowCount, sizeof(BOOL));
			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++)
				bResultset[rowNo] = TRUE;

			for (int colNo = 0; colNo < colCount; colNo++) {
				HWND hEdit = GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo);
				TCHAR filter[MAX_LENGTH];
				GetWindowText(hEdit, filter, MAX_LENGTH);
				int len = _tcslen(filter);
				if (len == 0)
					continue;

				for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
					if (!bResultset[rowNo])
						continue;

					TCHAR* value = cache[rowNo][colNo];
					if (len > 1 && (filter[0] == TEXT('<') || filter[0] == TEXT('>')) && isNumber(filter + 1)) {
						TCHAR* end = 0;
						double df = _tcstod(filter + 1, &end);
						double dv = _tcstod(value, &end);
						bResultset[rowNo] = (filter[0] == TEXT('<') && dv < df) || (filter[0] == TEXT('>') && dv > df);
					} else {
						bResultset[rowNo] = len == 1 ? _tcsstr(value, filter) != 0 :
							filter[0] == TEXT('=') ? _tcscmp(value, filter + 1) == 0 :
							filter[0] == TEXT('!') ? _tcsstr(value, filter + 1) == 0 :
							filter[0] == TEXT('<') ? _tcscmp(value, filter + 1) < 0 :
							filter[0] == TEXT('>') ? _tcscmp(value, filter + 1) > 0 :
							_tcsstr(value, filter) != 0;
					}
				}
			}

			int rowCount = 0;
			int* resultset = (int*)calloc(*pTotalRowCount, sizeof(int));
			for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
				if (!bResultset[rowNo])
					continue;

				resultset[rowCount] = rowNo;
				rowCount++;
			}
			free(bResultset);

			if (rowCount > 0) {
				resultset = realloc(resultset, rowCount * sizeof(int));
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)resultset);
				int orderBy = *pOrderBy;

				if (orderBy) {
					int colNo = orderBy > 0 ? orderBy - 1 : - orderBy - 1;
					BOOL isBackward = orderBy < 0;
					
					BOOL isNum = TRUE;
					for (int i = 0; i < *pTotalRowCount && i <= 5; i++) 
						isNum = isNum && isNumber(cache[i][colNo]);
												
					if (isNum) {
						double* nums = calloc(*pTotalRowCount, sizeof(double));
						for (int i = 0; i < rowCount; i++)
							nums[resultset[i]] = _tcstod(cache[resultset[i]][colNo], NULL);

						mergeSort(resultset, (void*)nums, 0, rowCount - 1, isBackward, isNum);
						free(nums);
					} else {
						TCHAR** strings = calloc(*pTotalRowCount, sizeof(TCHAR*));
						for (int i = 0; i < rowCount; i++)
							strings[resultset[i]] = cache[resultset[i]][colNo];
						mergeSort(resultset, (void*)strings, 0, rowCount - 1, isBackward, isNum);
						free(strings);
					}
				}
			} else {
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)0);			
				free(resultset);
			}

			*pRowCount = rowCount;
			ListView_SetItemCount(hGridWnd, rowCount);
			InvalidateRect(hGridWnd, NULL, TRUE);

			TCHAR buf[255];
			_sntprintf(buf, 255, TEXT(" Rows: %i/%i"), rowCount, *pTotalRowCount);
			SendMessage(hStatusWnd, SB_SETTEXT, SB_ROW_COUNT, (LPARAM)buf);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;
		
		case WMU_UPDATE_TEXT: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);

			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);			
			JSON_Value* val = (JSON_Value*)TreeView_GetItemParam(hTreeWnd, hItem);
			JSON_Value_Type type = json_value_get_type(val);

			TCHAR* text16;
			if (type == JSONArray || type == JSONObject) {
				char* text8 = json_serialize_to_string_pretty(val);
				text16 = utf8to16(text8);
				json_free_serialized_string(text8);
			} else {
				char* text8 = json_value_get_as_string(val);
				text16 = utf8to16(text8);
				free(text8);
			}
			
			LockWindowUpdate(hTextWnd);
			SendMessage(hTextWnd, EM_EXLIMITTEXT, 0, _tcslen(text16) + 1);
			SetWindowText(hTextWnd, text16);
			LockWindowUpdate(0);
			free(text16);		
			
			SendMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);	
		}
		break;
		
		case WMU_UPDATE_HIGHLIGHT: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);

			GETTEXTLENGTHEX gtl = {GTL_NUMBYTES, 0};
			int len = SendMessage(hTextWnd, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 1200);
			if (len < *(int*)GetProp(hWnd, TEXT("MAXHIGHLIGHTLENGTH"))) {
				TCHAR* text = calloc(len + sizeof(TCHAR), sizeof(char));
				GETTEXTEX gt = {0};
				gt.cb = len + sizeof(TCHAR);
				gt.flags = 0;
				gt.codepage = 1200;
				SendMessage(hTextWnd, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)text);
				LockWindowUpdate(hTextWnd);

				CHARFORMAT2 cf2 = {0};
				cf2.cbSize = sizeof(CHARFORMAT2) ;
				cf2.dwMask = CFM_COLOR;
				cf2.crTextColor = *(int*)GetProp(hWnd, TEXT("JSONTEXTCOLOR"));			
				SendMessage(hTextWnd, EM_SETSEL, 0, -1);
				SendMessage(hTextWnd, EM_SETCHARFORMAT, SCF_ALL, (LPARAM) &cf2);

				highlightBlock(hTextWnd, text, 0);
				free(text);
				SendMessage(hTextWnd, EM_SETSEL, 0, 0);
				LockWindowUpdate(0);
				InvalidateRect(hTextWnd, NULL, TRUE);
			}
		}
		break;
		
		case WMU_UPDATE_FILTER_SIZE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);
			SendMessage(hHeader, WM_SIZE, 0, 0);
			for (int colNo = 0; colNo < colCount; colNo++) {
				RECT rc;
				Header_GetItemRect(hHeader, colNo, &rc);
				int h2 = round((rc.bottom - rc.top) / 2);
				SetWindowPos(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), 0, rc.left - (colNo > 0), h2, rc.right - rc.left + 1, h2 + 1, SWP_NOZORDER);
			}
		}
		break;
		
		case WMU_SET_HEADER_FILTERS: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int isFilterRow = *(int*)GetProp(hWnd, TEXT("FILTERROW"));
			int colCount = Header_GetItemCount(hHeader);
			
			SendMessage(hWnd, WM_SETREDRAW, FALSE, 0);
			LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
			styles = isFilterRow ? styles | HDS_FILTERBAR : styles & (~HDS_FILTERBAR);
			SetWindowLongPtr(hHeader, GWL_STYLE, styles);
					
			for (int colNo = 0; colNo < colCount; colNo++) 		
				ShowWindow(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), isFilterRow ? SW_SHOW : SW_HIDE);

			if (isFilterRow)				
				SendMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);											

			// Bug fix: force Windows to redraw header
			int w = ListView_GetColumnWidth(hGridWnd, 0);
			ListView_SetColumnWidth(hGridWnd, 0, w + 1);
			ListView_SetColumnWidth(hGridWnd, 0, w);			
			SendMessage(hWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hWnd, NULL, TRUE);
		}
		break;
		
		case WMU_AUTO_COLUMN_SIZE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			SendMessage(hGridWnd, WM_SETREDRAW, FALSE, 0);
			HWND hHeader = ListView_GetHeader(hGridWnd);
			int colCount = Header_GetItemCount(hHeader);

			for (int colNo = 0; colNo < colCount - 1; colNo++)
				ListView_SetColumnWidth(hGridWnd, colNo, colNo < colCount - 1 ? LVSCW_AUTOSIZE_USEHEADER : LVSCW_AUTOSIZE);

			if (colCount == 1 && ListView_GetColumnWidth(hGridWnd, 0) < 100)
				ListView_SetColumnWidth(hGridWnd, 0, 100);

			int maxWidth = getStoredValue(TEXT("max-column-width"), 300);
			if (colCount > 1) {
				for (int colNo = 0; colNo < colCount; colNo++) {
					if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
						ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
				}
			}

			// Fix last column
			if (colCount > 1) {
				int colNo = colCount - 1;
				ListView_SetColumnWidth(hGridWnd, colNo, LVSCW_AUTOSIZE);
				TCHAR name16[MAX_LENGTH + 1];
				Header_GetItemText(hHeader, colNo, name16, MAX_LENGTH);
				
				SIZE s = {0};
				HDC hDC = GetDC(hHeader);
				HFONT hOldFont = (HFONT)SelectObject(hDC, (HFONT)GetProp(hWnd, TEXT("FONT")));
				GetTextExtentPoint32(hDC, name16, _tcslen(name16), &s);
				SelectObject(hDC, hOldFont);
				ReleaseDC(hHeader, hDC);

				int w = s.cx + 12;
				if (ListView_GetColumnWidth(hGridWnd, colNo) < w)
					ListView_SetColumnWidth(hGridWnd, colNo, w);

				if (ListView_GetColumnWidth(hGridWnd, colNo) > maxWidth)
					ListView_SetColumnWidth(hGridWnd, colNo, maxWidth);
			}

			SendMessage(hGridWnd, WM_SETREDRAW, TRUE, 0);
			InvalidateRect(hGridWnd, NULL, TRUE);

			PostMessage(hWnd, WMU_UPDATE_FILTER_SIZE, 0, 0);
		}
		break;

		case WMU_RESET_CACHE: {
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			TCHAR*** cache = (TCHAR***)GetProp(hWnd, TEXT("CACHE"));
			int* pTotalRowCount = (int*)GetProp(hWnd, TEXT("TOTALROWCOUNT"));

			int colCount = Header_GetItemCount(ListView_GetHeader(hGridWnd));
			if (colCount > 0 && cache != 0) {
				for (int rowNo = 0; rowNo < *pTotalRowCount; rowNo++) {
					if (cache[rowNo]) {
						for (int colNo = 0; colNo < colCount; colNo++)
							if (cache[rowNo][colNo])
								free(cache[rowNo][colNo]);

						free(cache[rowNo]);
					}
					cache[rowNo] = 0;
				}
				free(cache);
			}

			int* resultset = (int*)GetProp(hWnd, TEXT("RESULTSET"));
			if (resultset)
				free(resultset);
			int* pRowCount = (int*)GetProp(hWnd, TEXT("ROWCOUNT"));
			*pRowCount = 0;

			SetProp(hWnd, TEXT("CACHE"), 0);
			*pTotalRowCount = 0;
		}
		break;

		// wParam - size delta
		case WMU_SET_FONT: {
			int* pFontSize = (int*)GetProp(hWnd, TEXT("FONTSIZE"));
			if (*pFontSize + wParam < 10 || *pFontSize + wParam > 48)
				return 0;
			*pFontSize += wParam;
			DeleteFont(GetProp(hWnd, TEXT("FONT")));

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, (TCHAR*)GetProp(hWnd, TEXT("FONTFAMILY")));
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			SendMessage(hTreeWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hGridWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hTextWnd, WM_SETFONT, (LPARAM)hFont, TRUE);
			SendMessage(hTabWnd, WM_SETFONT, (LPARAM)GetStockObject(DEFAULT_GUI_FONT), TRUE);			

			HWND hHeader = ListView_GetHeader(hGridWnd);
			for (int colNo = 0; colNo < Header_GetItemCount(hHeader); colNo++)
				SendMessage(GetDlgItem(hHeader, IDC_HEADER_EDIT + colNo), WM_SETFONT, (LPARAM)hFont, TRUE);

			SetProp(hWnd, TEXT("FONT"), hFont);

			SendMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);
			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;
		
		case WMU_SET_THEME: {
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			BOOL isDark = *(int*)GetProp(hWnd, TEXT("DARKTHEME"));
			
			int textColor = !isDark ? getStoredValue(TEXT("text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("text-color-dark"), RGB(220, 220, 220));
			int backColor = !isDark ? getStoredValue(TEXT("back-color"), RGB(255, 255, 255)) : getStoredValue(TEXT("back-color-dark"), RGB(32, 32, 32));
			int filterTextColor = !isDark ? getStoredValue(TEXT("filter-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("filter-text-color-dark"), RGB(255, 255, 255));
			int filterBackColor = !isDark ? getStoredValue(TEXT("filter-back-color"), RGB(240, 240, 240)) : getStoredValue(TEXT("filter-back-color-dark"), RGB(60, 60, 60));
			int splitterColor = !isDark ? getStoredValue(TEXT("splitter-color"), GetSysColor(COLOR_BTNFACE)) : getStoredValue(TEXT("splitter-color-dark"), GetSysColor(COLOR_BTNFACE));			
			
			*(int*)GetProp(hWnd, TEXT("TEXTCOLOR")) = textColor;
			*(int*)GetProp(hWnd, TEXT("BACKCOLOR")) = backColor;
			*(int*)GetProp(hWnd, TEXT("FILTERTEXTCOLOR")) = filterTextColor;
			*(int*)GetProp(hWnd, TEXT("FILTERBACKCOLOR")) = filterBackColor;
			*(int*)GetProp(hWnd, TEXT("SPLITTERCOLOR")) = splitterColor;
			
			*(int*)GetProp(hWnd, TEXT("JSONTEXTCOLOR")) = !isDark ? getStoredValue(TEXT("json-text-color"), RGB(0, 0, 0)) : getStoredValue(TEXT("json-text-color-dark"), RGB(220, 220, 220));
			*(int*)GetProp(hWnd, TEXT("JSONKEYCOLOR")) = !isDark ? getStoredValue(TEXT("json-key-color"), RGB(128, 0, 128)) : getStoredValue(TEXT("json-key-color-dark"), RGB(200, 0, 200));			
			*(int*)GetProp(hWnd, TEXT("JSONSTRINGCOLOR")) = !isDark ? getStoredValue(TEXT("json-string-color"), RGB(0, 128, 0)) : getStoredValue(TEXT("json-string-color-dark"), RGB(0, 128, 0));
			*(int*)GetProp(hWnd, TEXT("JSONBOOLEANCOLOR")) = !isDark ? getStoredValue(TEXT("json-boolean-color"), RGB(0, 0, 255)) : getStoredValue(TEXT("json-boolean-color-dark"), RGB(0, 0, 128));	
			*(int*)GetProp(hWnd, TEXT("JSONNULLCOLOR")) = !isDark ? getStoredValue(TEXT("json-null-color"), RGB(255, 0, 0)) : getStoredValue(TEXT("json-null-color-dark"), RGB(255, 0, 0));				

			DeleteObject(GetProp(hWnd, TEXT("BACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("FILTERBACKBRUSH")));			
			DeleteObject(GetProp(hWnd, TEXT("SPLITTERBRUSH")));			
			SetProp(hWnd, TEXT("BACKBRUSH"), CreateSolidBrush(backColor));
			SetProp(hWnd, TEXT("FILTERBACKBRUSH"), CreateSolidBrush(filterBackColor));
			SetProp(hWnd, TEXT("SPLITTERBRUSH"), CreateSolidBrush(splitterColor));			

			TreeView_SetTextColor(hTreeWnd, textColor);			
			TreeView_SetBkColor(hTreeWnd, backColor);

			ListView_SetTextColor(hGridWnd, textColor);			
			ListView_SetBkColor(hGridWnd, backColor);
			ListView_SetTextBkColor(hGridWnd, backColor);
						
			SendMessage(hTextWnd, EM_SETBKGNDCOLOR, 0, backColor);			
			
			SendMessage(hWnd, WMU_UPDATE_HIGHLIGHT, 0, 0);
			InvalidateRect(hWnd, NULL, TRUE);	
		}
		break;
		

	}
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewHeader(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_CTLCOLOREDIT) {
		HWND hMainWnd = getMainWindow(hWnd);
		SetBkColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERBACKCOLOR")));
		SetTextColor((HDC)wParam, *(int*)GetProp(hMainWnd, TEXT("FILTERTEXTCOLOR")));
		return (INT_PTR)(HBRUSH)GetProp(hMainWnd, TEXT("FILTERBACKBRUSH"));	
	}
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetClientRect(hWnd, &rc);
			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, *(int*)GetProp(getMainWindow(hWnd), TEXT("FILTERBACKCOLOR")));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

		case WM_CHAR: {
			return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN) {
				SendMessage(getMainWindow(hWnd), WMU_UPDATE_RESULTSET, 0, 0);
				return 0;
			}
			
			if (wParam == VK_TAB || wParam == VK_ESCAPE)
				return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);		
		}
		break;

		case WM_DESTROY: {
			RemoveProp(hWnd, TEXT("WNDPROC"));
		}
		break;
	}

	return CallWindowProc(cbDefault, hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == EM_SETZOOM)
		return 0;

	if (msg == WM_MOUSEWHEEL && LOWORD(wParam) == MK_CONTROL) {
		HWND hTabWnd = GetParent(hWnd);
		SendMessage(GetParent(hTabWnd), msg, wParam, lParam);
		return 0;
	}
	
	if (msg == WM_KEYDOWN && (wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || HIWORD(GetKeyState(VK_CONTROL)) && wParam == 0x46)) { // Ctrl + F
		SendMessage(GetAncestor(hWnd, GA_ROOT), WM_KEYDOWN, wParam, lParam);
		return 0;
	}
	
	if (msg == WM_KEYDOWN && wParam == VK_F1) {
		SendMessage(getMainWindow(hWnd), WM_KEYDOWN, wParam, lParam);
		return TRUE;
	}
		
	if (msg == WM_KEYDOWN && (wParam == VK_TAB || wParam == VK_ESCAPE)) { 
		return CallWindowProc(cbHotKey, hWnd, msg, wParam, lParam);
	}

	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbHotKey(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	if (msg == WM_KEYDOWN && (
		wParam == VK_TAB || wParam == VK_ESCAPE || 
		wParam == VK_F3 || wParam == VK_F5 || wParam == VK_F7 || (HIWORD(GetKeyState(VK_CONTROL)) && wParam == 0x46) || // Ctrl + F
		wParam == VK_F1 || wParam == VK_F11 ||
		(wParam >= 0x31 && wParam <= 0x42) && !getStoredValue(TEXT("disable-num-keys"), 0) || // 1 - 8
		(wParam == 0x4E || wParam == 0x50) && !getStoredValue(TEXT("disable-np-keys"), 0))) { // N, P
		HWND hMainWnd = getMainWindow(hWnd);
		if (wParam == VK_F7 || wParam == 0x46)
			*(int*)GetProp(hMainWnd, TEXT("SEARCHCOLNO")) = 0;
		SendMessage(wParam == VK_TAB || wParam == VK_ESCAPE || wParam == VK_F1 ? hMainWnd : GetParent(hMainWnd), WM_KEYDOWN, wParam, lParam);
		return 0;
	}
	
	// Prevent beep
	if (msg == WM_CHAR && (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB))
		return 0;
	
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

char* json_value_get_as_string(const JSON_Value* val) {
	char* res = 0;
	int type = json_value_get_type(val);
	if (type == JSONObject) {
		res = calloc(10, sizeof(char));
		snprintf(res, 10, "[Object]");
	} else if (type == JSONArray) {
		res = calloc(10, sizeof(char));
		snprintf(res, 10, "[Array]");
	} else if (type == JSONString) {
		int len = json_value_get_string_len(val);
		res = calloc(len + 1, sizeof(char));
		strncpy(res, json_value_get_string(val), len);
	} else {
		char* str = json_serialize_to_string_pretty(val);
		int len = strlen(str);
		res = calloc(len + 1, sizeof(char));
		strncpy(res, str, len);
		json_free_serialized_string(str);
	}

	return res;
}

BOOL addNode(HWND hTreeWnd, HTREEITEM hParentItem, JSON_Value* val) {
	JSON_Value_Type type = json_value_get_type(val);
	if (type == JSONArray) {
		JSON_Array* arr = json_value_get_array(val);
		int cnt = json_array_get_count(arr);
		for (int itemNo = 0; itemNo < cnt; itemNo++) {
			JSON_Value* val2 = json_array_get_value(arr, cnt - itemNo - 1);
			TCHAR itemName[64];
			_sntprintf(itemName, 64, TEXT("[%i]"), cnt - itemNo - 1);
			HTREEITEM hItem = TreeView_AddItem(hTreeWnd, itemName, hParentItem, (LPARAM)val2);
			addNode(hTreeWnd, hItem, val2);
		}

		return TRUE;
	}

	if (type == JSONObject) {
		JSON_Object* obj = json_value_get_object(val);
		int cnt = json_object_get_count(obj);
		for (int itemNo = 0; itemNo < cnt; itemNo++) {
			JSON_Value* val2 = json_object_get_value_at(obj, cnt - itemNo - 1);
			TCHAR* itemName = utf8to16(json_object_get_name(obj, cnt - itemNo - 1));
			HTREEITEM hItem = TreeView_AddItem(hTreeWnd, itemName, hParentItem, (LPARAM)val2);
			addNode(hTreeWnd, hItem, val2);
			free(itemName);
		}

		return TRUE;
	}

	char* value8 = json_value_get_as_string(val);
	TCHAR* value16 = utf8to16(value8);
	free(value8);

	TCHAR key16[MAX_LENGTH + 1];
	TreeView_GetItemText(hTreeWnd, hParentItem, key16, MAX_LENGTH);

	int len = _tcslen(key16) + _tcslen(value16) + 3;
	TCHAR pair16[len + 1];
	_sntprintf(pair16, len + 1, TEXT("%ls = %ls"), key16, value16);
	TreeView_SetItemText(hTreeWnd, hParentItem, pair16);
	free(value16);

	return TRUE;
}

int highlightBlock(HWND hWnd, TCHAR* text, int start) {
	int textLen = _tcslen(text);
	if (start < 0 || start >= textLen || (text[start] != TEXT('{') && text[start] != TEXT('[')))
		return 0;

	BOOL isObject = text[start] == TEXT('{');
	BOOL isKey = isObject;
	
	HWND hMainWnd = getMainWindow(hWnd);
	COLORREF keyColor = *(int*)GetProp(hMainWnd, TEXT("JSONKEYCOLOR"));
	COLORREF stringColor = *(int*)GetProp(hMainWnd, TEXT("JSONSTRINGCOLOR"));	
	COLORREF booleanColor = *(int*)GetProp(hMainWnd, TEXT("JSONBOOLEANCOLOR"));	
	COLORREF nullColor = *(int*)GetProp(hMainWnd, TEXT("JSONNULLCOLOR"));	

	CHARFORMAT2 cf2 = {0};
	cf2.cbSize = sizeof(CHARFORMAT2) ;
	cf2.dwMask = CFM_COLOR | CFM_BOLD;

	int pos = start + 1;
	while (pos < textLen) {
		TCHAR c = text[pos];		
		if (c == TEXT('[') || c == TEXT('{')) {
			pos += highlightBlock(hWnd, text, pos);
		} else if ((c == TEXT(']') && !isObject) || (c == TEXT('}') && isObject)) {
			break;
		} else if (c == TEXT('"')) {
			int end = pos + 1;
			int sCount = 0; // to process escaped quotes e.g. abc\\\"abc
			while (end < textLen) {
				if (sCount % 2 == 0 && text[end] == TEXT('"'))
					break;
				sCount = text[end] == TEXT('\\') ? sCount + 1 : 0;
				end++;
			}
			end++;

			cf2.crTextColor = isKey ? keyColor : stringColor;
			cf2.dwEffects = isKey ? CFM_BOLD : 0;
			SendMessage(hWnd, EM_SETSEL, pos, end);
			SendMessage(hWnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM) &cf2);

			pos = end;
			isKey = FALSE;
		} else if (_istalpha(c)) {
			TCHAR buf[10] = {0};
			_tcsncpy(buf, text + pos, 9);
			_tcslwr(buf);

			int end = pos;
			if (_tcsstr(buf, TEXT("true")) == buf) {
				cf2.crTextColor = booleanColor;
				end += 4;
			} else if (_tcsstr(buf, TEXT("false")) == buf) {
				cf2.crTextColor = booleanColor;
				end += 5;
			} else if (_tcsstr(buf, TEXT("null")) == buf) {
				cf2.crTextColor = nullColor;
				end += 4;
				cf2.dwEffects = CFM_BOLD;
			}

			SendMessage(hWnd, EM_SETSEL, pos, end);
			SendMessage(hWnd, EM_SETCHARFORMAT, SCF_SELECTION, (LPARAM) &cf2);
			pos = end;
		}

		isKey = isKey || (isObject && pos < textLen && text[pos] == TEXT(','));
		pos++;
	}

	return pos - start;
}

HWND getMainWindow(HWND hWnd) {
	HWND hMainWnd = hWnd;
	while (hMainWnd && GetDlgCtrlID(hMainWnd) != IDC_MAIN)
		hMainWnd = GetParent(hMainWnd);
	return hMainWnd;	
}

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
}

TCHAR* getStoredString(TCHAR* name, TCHAR* defValue) { 
	TCHAR* buf = calloc(256, sizeof(TCHAR));
	if (0 == GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) && defValue)
		_tcsncpy(buf, defValue, 255);
	return buf;	
}

int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam) {
	if (GetWindowLong(hWnd, GWL_STYLE) & WS_TABSTOP && IsWindowVisible(hWnd)) {
		int no = 0;
		HWND* wnds = (HWND*)lParam;
		while (wnds[no])
			no++;
		wnds[no] = hWnd;
	}

	return TRUE;
}

TCHAR* utf8to16(const char* in) {
	TCHAR *out;
	if (!in || strlen(in) == 0) {
		out = (TCHAR*)calloc (1, sizeof (TCHAR));
	} else  {
		DWORD size = MultiByteToWideChar(CP_UTF8, 0, in, -1, NULL, 0);
		out = (TCHAR*)calloc (size, sizeof (TCHAR));
		MultiByteToWideChar(CP_UTF8, 0, in, -1, out, size);
	}
	return out;
}

char* utf16to8(const TCHAR* in) {
	char* out;
	if (!in || _tcslen(in) == 0) {
		out = (char*)calloc (1, sizeof(char));
	} else  {
		int len = WideCharToMultiByte(CP_UTF8, 0, in, -1, NULL, 0, 0, 0);
		out = (char*)calloc (len, sizeof(char));
		WideCharToMultiByte(CP_UTF8, 0, in, -1, out, len, 0, 0);
	}
	return out;
}

int findString(TCHAR* text, TCHAR* word, BOOL isMatchCase, BOOL isWholeWords) {
	if (!text || !word)
		return -1;
		
	int res = -1;
	int tlen = _tcslen(text);
	int wlen = _tcslen(word);	
	if (!tlen || !wlen)
		return res;
	
	if (!isMatchCase) {
		TCHAR* ltext = calloc(tlen + 1, sizeof(TCHAR));
		_tcsncpy(ltext, text, tlen);
		text = _tcslwr(ltext);

		TCHAR* lword = calloc(wlen + 1, sizeof(TCHAR));
		_tcsncpy(lword, word, wlen);
		word = _tcslwr(lword);
	}

	if (isWholeWords) {
		for (int pos = 0; (res  == -1) && (pos <= tlen - wlen); pos++) 
			res = (pos == 0 || pos > 0 && !_istalnum(text[pos - 1])) && 
				!_istalnum(text[pos + wlen]) && 
				_tcsncmp(text + pos, word, wlen) == 0 ? pos : -1;
	} else {
		TCHAR* s = _tcsstr(text, word);
		res = s != NULL ? s - text : -1;
	}
	
	if (!isMatchCase) {
		free(text);
		free(word);
	}

	return res; 
}	

void setClipboardText(const TCHAR* text) {
	int len = (_tcslen(text) + 1) * sizeof(TCHAR);
	HGLOBAL hMem =  GlobalAlloc(GMEM_MOVEABLE, len);
	memcpy(GlobalLock(hMem), text, len);
	GlobalUnlock(hMem);
	OpenClipboard(0);
	EmptyClipboard();
	SetClipboardData(CF_UNICODETEXT, hMem);
	CloseClipboard();
}

BOOL isNumber(TCHAR* val) {
	int len = _tcslen(val);
	BOOL res = TRUE;
	int pCount = 0;
	for (int i = 0; res && i < len; i++) {
		pCount += val[i] == TEXT('.');
		res = _istdigit(val[i]) || val[i] == TEXT('.');
	}
	return res && pCount < 2;
}

// https://stackoverflow.com/a/1031773/6121703
BOOL isUtf8(const char * string) {
	if (!string)
		return FALSE;

	const unsigned char * bytes = (const unsigned char *)string;
	while (*bytes) {
		if((bytes[0] == 0x09 || bytes[0] == 0x0A || bytes[0] == 0x0D || (0x20 <= bytes[0] && bytes[0] <= 0x7E))) {
			bytes += 1;
			continue;
		}

		if (((0xC2 <= bytes[0] && bytes[0] <= 0xDF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF))) {
			bytes += 2;
			continue;
		}

		if ((bytes[0] == 0xE0 && (0xA0 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(((0xE1 <= bytes[0] && bytes[0] <= 0xEC) || bytes[0] == 0xEE || bytes[0] == 0xEF) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF)) ||
			(bytes[0] == 0xED && (0x80 <= bytes[1] && bytes[1] <= 0x9F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF))
		) {
			bytes += 3;
			continue;
		}

		if ((bytes[0] == 0xF0 && (0x90 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			((0xF1 <= bytes[0] && bytes[0] <= 0xF3) && (0x80 <= bytes[1] && bytes[1] <= 0xBF) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF)) ||
			(bytes[0] == 0xF4 && (0x80 <= bytes[1] && bytes[1] <= 0x8F) && (0x80 <= bytes[2] && bytes[2] <= 0xBF) && (0x80 <= bytes[3] && bytes[3] <= 0xBF))
		) {
			bytes += 4;
			continue;
		}

		return FALSE;
	}

	return TRUE;
}

void mergeSortJoiner(int indexes[], void* data, int l, int m, int r, BOOL isBackward, BOOL isNums) {
    int n1 = m - l + 1;
    int n2 = r - m;

    int L[n1], R[n2];

    for (int i = 0; i < n1; i++)
        L[i] = indexes[l + i];
    for (int j = 0; j < n2; j++)
        R[j] = indexes[m + 1 + j];

    int i = 0, j = 0, k = l;
    while (i < n1 && j < n2) {
    	int cmp = isNums ? ((double*)data)[L[i]] <= ((double*)data)[R[j]] : _tcscmp(((TCHAR**)data)[L[i]], ((TCHAR**)data)[R[j]]) <= 0;
    	if (isBackward)
    		cmp = !cmp;
    		
        if (cmp) {
            indexes[k] = L[i];
            i++;
        } else {
            indexes[k] = R[j];
            j++;
        }
        k++;
    }

    while (i < n1) {
        indexes[k] = L[i];
        i++;
        k++;
    }

    while (j < n2) {
        indexes[k] = R[j];
        j++;
        k++;
    }
}

void mergeSort(int indexes[], void* data, int l, int r, BOOL isBackward, BOOL isNums) {
    if (l < r) {
        int m = l + (r - l) / 2;
        mergeSort(indexes, data, l, m, isBackward, isNums);
        mergeSort(indexes, data, m + 1, r, isBackward, isNums);
        mergeSortJoiner(indexes, data, l, m, r, isBackward, isNums);
    }
}

HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam) {
	TVITEM tvi = {0};
	TVINSERTSTRUCT tvins = {0};
	tvi.mask = TVIF_TEXT | TVIF_PARAM;
	tvi.pszText = caption;
	tvi.cchTextMax = _tcslen(caption) + 1;
	tvi.lParam = lParam;

	tvins.item = tvi;
	tvins.hInsertAfter = TVI_FIRST;
	tvins.hParent = parent;
	return (HTREEITEM)SendMessage(hTreeWnd, TVM_INSERTITEM, 0, (LPARAM)(LPTVINSERTSTRUCT)&tvins);
};

int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_TEXT;
	tv.hItem = hItem;
	tv.cchTextMax = maxLen;
	tv.pszText = buf;
	return TreeView_GetItem(hTreeWnd, &tv);
}

LPARAM TreeView_GetItemParam(HWND hTreeWnd, HTREEITEM hItem) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_PARAM;
	tv.hItem = hItem;

	return TreeView_GetItem(hTreeWnd, &tv) ? tv.lParam : 0;
}

int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text) {
	TV_ITEM tv = {0};
	tv.mask = TVIF_TEXT;
	tv.hItem = hItem;
	tv.cchTextMax = _tcslen(text) + 1;
	tv.pszText = text;
	return TreeView_SetItem(hTreeWnd, &tv);
}

int ListView_AddColumn(HWND hListWnd, TCHAR* colName) {
	int colNo = Header_GetItemCount(ListView_GetHeader(hListWnd));
	LVCOLUMN lvc = {0};
	lvc.mask = LVCF_WIDTH | LVCF_TEXT | LVCF_SUBITEM;
	lvc.iSubItem = colNo;
	lvc.pszText = colName;
	lvc.cchTextMax = _tcslen(colName) + 1;
	lvc.cx = 100;
	return ListView_InsertColumn(hListWnd, colNo, &lvc);
}

int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax) {
	if (i < 0)
		return FALSE;

	TCHAR buf[cchTextMax];

	HDITEM hdi = {0};
	hdi.mask = HDI_TEXT;
	hdi.pszText = buf;
	hdi.cchTextMax = cchTextMax;
	int rc = Header_GetItem(hWnd, i, &hdi);

	_tcsncpy(pszText, buf, cchTextMax);
	return rc;
}

void Menu_SetItemState(HMENU hMenu, UINT wID, UINT fState) {
	MENUITEMINFO mii = {0};
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STATE;
	mii.fState = fState;
	SetMenuItemInfo(hMenu, wID, FALSE, &mii);
}
