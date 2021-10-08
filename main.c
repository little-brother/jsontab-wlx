#define UNICODE
#define _UNICODE

#include <windows.h>
#include <windowsx.h>
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
#define WMU_AUTO_COLUMN_SIZE   WM_USER + 4
#define WMU_RESET_CACHE        WM_USER + 5
#define WMU_SET_FONT           WM_USER + 6

#define IDC_MAIN               100
#define IDC_TREE               101
#define IDC_TAB                102
#define IDC_GRID               103
#define IDC_TEXT               104
#define IDC_STATUSBAR          105
#define IDC_HEADER_EDIT        1000

#define IDM_COPY_CELL          5000
#define IDM_COPY_ROW           5001
#define IDM_COPY_TEXT          5002
#define IDM_SELECTALL          5003

#define IDH_EXIT               6000
#define IDH_NEXT               6001
#define IDH_PREV               6002

#define SB_CHILD_NO            0
#define SB_TYPE                1
#define SB_ROW_COUNT           2
#define SB_CURRENT_ROW         3

#define MAX_LENGTH             4096
#define MAX_HIGHLIGHT_LENGTH   64000
#define APP_NAME               TEXT("json-wlx")

typedef struct {
	int size;
	DWORD PluginInterfaceVersionLow;
	DWORD PluginInterfaceVersionHi;
	char DefaultIniName[MAX_PATH];
} ListDefaultParamStruct;

static TCHAR iniPath[MAX_PATH] = {0};

LRESULT CALLBACK cbNewMain (HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
LRESULT CALLBACK cbNewText(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam);
char* json_value_get_as_string(const JSON_Value* val);
BOOL addNode(HWND hTreeWnd, HTREEITEM hParentItem, JSON_Value* val);
int highlightBlock(HWND hWnd, TCHAR* text, int start);
void setStoredValue(TCHAR* name, int value);
int getStoredValue(TCHAR* name, int defValue);
int CALLBACK cbEnumTabStopChildren (HWND hWnd, LPARAM lParam);
TCHAR* utf8to16(const char* in);
char* utf16to8(const TCHAR* in);
void setClipboardText(const TCHAR* text);
BOOL isNumber(TCHAR* val);
BOOL isUtf8(const char * string);
HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam);
int TreeView_GetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* buf, int maxLen);
int TreeView_SetItemText(HWND hTreeWnd, HTREEITEM hItem, TCHAR* text);
int ListView_AddColumn(HWND hListWnd, TCHAR* colName);
int Header_GetItemText(HWND hWnd, int i, TCHAR* pszText, int cchTextMax);

BOOL APIENTRY DllMain (HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved) {
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

HWND APIENTRY ListLoad (HWND hListerWnd, char* fileToLoad, int showFlags) {
	DWORD size = MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, NULL, 0);
	TCHAR* filepath = (TCHAR*)calloc (size, sizeof (TCHAR));
	MultiByteToWideChar(CP_ACP, 0, fileToLoad, -1, filepath, size);

	struct _stat st = {0};
	if (_tstat(filepath, &st) != 0 || st.st_size > getStoredValue(TEXT("max-file-size"), 1000000))
		return 0;

	char* data = calloc(st.st_size + 1, sizeof(char));
	FILE *f = fopen(fileToLoad, "rb");
	fread(data, sizeof(char), st.st_size, f);
	fclose(f);

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

	BOOL isStandalone = GetParent(hListerWnd) == HWND_DESKTOP;
	HWND hMainWnd = CreateWindowEx(WS_EX_CONTROLPARENT, WC_STATIC, TEXT("json-wlx"), WS_CHILD | WS_VISIBLE | (isStandalone ? SS_SUNKEN : 0),
		0, 0, 100, 100, hListerWnd, (HMENU)IDC_MAIN, GetModuleHandle(0), NULL);

	SetProp(hMainWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hMainWnd, GWLP_WNDPROC, (LONG_PTR)&cbNewMain));
	SetProp(hMainWnd, TEXT("JSON"), json);
	SetProp(hMainWnd, TEXT("CACHE"), 0);
	SetProp(hMainWnd, TEXT("RESULTSET"), 0);
	SetProp(hMainWnd, TEXT("ROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("TOTALROWCOUNT"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("ORDERBY"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("COLNO"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("SPLITTERWIDTH"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("FONT"), 0);
	SetProp(hMainWnd, TEXT("FONTSIZE"), calloc(1, sizeof(int)));
	SetProp(hMainWnd, TEXT("GRAYBRUSH"), CreateSolidBrush(GetSysColor(COLOR_BTNFACE)));

	*(int*)GetProp(hMainWnd, TEXT("SPLITTERWIDTH")) = getStoredValue(TEXT("splitter-width"), 200);
	*(int*)GetProp(hMainWnd, TEXT("FONTSIZE")) = getStoredValue(TEXT("font-size"), 16);

	HWND hStatusWnd = CreateStatusWindow(WS_CHILD | WS_VISIBLE |  (isStandalone ? SBARS_SIZEGRIP : 0), NULL, hMainWnd, IDC_STATUSBAR);
	int sizes[5] = {100, 200, 400, 500, -1};
	SendMessage(hStatusWnd, SB_SETPARTS, 5, (LPARAM)&sizes);

	HWND hTreeWnd = CreateWindow(WC_TREEVIEW, NULL, WS_VISIBLE | WS_CHILD | TVS_HASBUTTONS | TVS_HASLINES | TVS_SHOWSELALWAYS | WS_TABSTOP,
		0, 0, 100, 100, hMainWnd, (HMENU)IDC_TREE, GetModuleHandle(0), NULL);

	HWND hTabWnd = CreateWindow(WC_TABCONTROL, NULL, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | WS_TABSTOP, 100, 100, 100, 100,
		hMainWnd, (HMENU)IDC_TAB, GetModuleHandle(0), NULL);

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

	HWND hHeader = ListView_GetHeader(hGridWnd);
	LONG_PTR styles = GetWindowLongPtr(hHeader, GWL_STYLE);
	SetWindowLongPtr(hHeader, GWL_STYLE, styles | HDS_FILTERBAR);

	HWND hTextWnd = CreateWindowEx(0, TEXT("RICHEDIT50W"), NULL, (tabNo == 1 ? WS_VISIBLE : 0) | WS_CHILD | ES_MULTILINE | ES_AUTOVSCROLL | ES_AUTOHSCROLL | ES_WANTRETURN | WS_VSCROLL | WS_HSCROLL | WS_TABSTOP | ES_NOHIDESEL | ES_READONLY,
		0, 0, 100, 100, hTabWnd, (HMENU)IDC_TEXT, GetModuleHandle(0),  NULL);
	SetProp(hTextWnd, TEXT("WNDPROC"), (HANDLE)SetWindowLongPtr(hTextWnd, GWLP_WNDPROC, (LONG_PTR)cbNewText));
	TabCtrl_SetCurSel(hTabWnd, tabNo);

	HMENU hDataMenu = CreatePopupMenu();
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_CELL, TEXT("Copy cell"));
	AppendMenu(hDataMenu, MF_STRING, IDM_COPY_ROW, TEXT("Copy row"));
	SetProp(hMainWnd, TEXT("DATAMENU"), hDataMenu);

	HMENU hTextMenu = CreatePopupMenu();
	AppendMenu(hTextMenu, MF_STRING, IDM_COPY_TEXT, TEXT("Copy"));
	AppendMenu(hTextMenu, MF_STRING, IDM_SELECTALL, TEXT("Select all"));
	SetProp(hMainWnd, TEXT("TEXTMENU"), hTextMenu);

	SendMessage(hMainWnd, WMU_SET_FONT, 0, 0);
	SendMessage(hMainWnd, WM_SIZE, 0, 0);

	HTREEITEM hRoot = TreeView_AddItem(hTreeWnd, TEXT("<<root>>"), TVI_ROOT, (LPARAM)json);
	addNode(hTreeWnd, hRoot, json);
	TreeView_Expand(hTreeWnd, hRoot, TVE_EXPAND);
	TreeView_Select(hTreeWnd, hRoot, TVGN_CARET);

	RegisterHotKey(hMainWnd, IDH_EXIT, 0, VK_ESCAPE);
	RegisterHotKey(hMainWnd, IDH_NEXT, 0, VK_TAB);
	RegisterHotKey(hMainWnd, IDH_PREV, MOD_CONTROL, VK_TAB);
	SetFocus(hTreeWnd);

	return hMainWnd;
}

void __stdcall ListCloseWindow(HWND hWnd) {
	setStoredValue(TEXT("splitter-width"), *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")));
	setStoredValue(TEXT("font-size"), *(int*)GetProp(hWnd, TEXT("FONTSIZE")));
	setStoredValue(TEXT("tab-no"), TabCtrl_GetCurSel(GetDlgItem(hWnd, IDC_TAB)));

	SendMessage(hWnd, WMU_RESET_CACHE, 0, 0);
	json_value_free((JSON_Value*)GetProp(hWnd, TEXT("JSON")));
	free((int*)GetProp(hWnd, TEXT("ROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("TOTALROWCOUNT")));
	free((int*)GetProp(hWnd, TEXT("ORDERBY")));
	free((int*)GetProp(hWnd, TEXT("COLNO")));
	free((int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")));

	DeleteFont(GetProp(hWnd, TEXT("FONT")));
	DeleteObject(GetProp(hWnd, TEXT("GRAYBRUSH")));
	DestroyMenu(GetProp(hWnd, TEXT("DATAMENU")));
	DestroyMenu(GetProp(hWnd, TEXT("TEXTMENU")));

	RemoveProp(hWnd, TEXT("WNDPROC"));
	RemoveProp(hWnd, TEXT("CACHE"));
	RemoveProp(hWnd, TEXT("RESULTSET"));
	RemoveProp(hWnd, TEXT("ROWCOUNT"));
	RemoveProp(hWnd, TEXT("TOTALROWCOUNT"));
	RemoveProp(hWnd, TEXT("ORDERBY"));
	RemoveProp(hWnd, TEXT("COLNO"));
	RemoveProp(hWnd, TEXT("JSON"));
	RemoveProp(hWnd, TEXT("SPLITTERWIDTH"));

	RemoveProp(hWnd, TEXT("FONT"));
	RemoveProp(hWnd, TEXT("GRAYBRUSH"));
	RemoveProp(hWnd, TEXT("DATAMENU"));
	RemoveProp(hWnd, TEXT("TEXTMENU"));

	DestroyWindow(hWnd);
	return;
}

LRESULT CALLBACK cbNewMain(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	switch (msg) {
		case WM_HOTKEY: {
			WPARAM id = wParam;
			if (id == IDH_EXIT)
				SendMessage(GetParent(hWnd), WM_CLOSE, 0, 0);

			if (id == IDH_NEXT || id == IDH_PREV) {
				HWND hFocus = GetFocus();
				HWND wnds[1000] = {0};
				EnumChildWindows(hWnd, (WNDENUMPROC)cbEnumTabStopChildren, (LPARAM)wnds);

				int no = 0;
				while(wnds[no] && wnds[no] != hFocus)
					no++;

				int cnt = no;
				while(wnds[cnt])
					cnt++;

				BOOL isBackward = id == IDH_PREV;
				no += isBackward ? -1 : 1;
				SetFocus(wnds[no] ? wnds[no] : (isBackward ? wnds[cnt - 1] : wnds[0]));
			}
		}
		break;

		case WM_SIZE: {
			HWND hStatusWnd = GetDlgItem(hWnd, IDC_STATUSBAR);
			SendMessage(hStatusWnd, WM_SIZE, 0, 0);
			RECT rc;
			GetClientRect(hStatusWnd, &rc);
			int statusH = rc.bottom;

			int splitterW = *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH"));
			GetClientRect(hWnd, &rc);
			HWND hTreeWnd = GetDlgItem(hWnd, IDC_TREE);
			HWND hTabWnd = GetDlgItem(hWnd, IDC_TAB);
			HWND hGridWnd = GetDlgItem(hTabWnd, IDC_GRID);
			HWND hTextWnd = GetDlgItem(hTabWnd, IDC_TEXT);
			SetWindowPos(hTreeWnd, 0, 0, 0, splitterW, rc.bottom - statusH, SWP_NOMOVE | SWP_NOZORDER);
			SetWindowPos(hTabWnd, 0, splitterW + 5, 0, rc.right - splitterW - 5, rc.bottom - statusH, SWP_NOZORDER);

			RECT rc2;
			GetClientRect(hTabWnd, &rc);
			TabCtrl_GetItemRect(hTabWnd, 0, &rc2);
			SetWindowPos(hTextWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - 5, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
			SetWindowPos(hGridWnd, 0, 2, rc2.bottom + 3, rc.right - rc.left - 5, rc.bottom - rc2.bottom - 7, SWP_NOZORDER);
		}
		break;

		case WM_PAINT: {
			PAINTSTRUCT ps = {0};
			HDC hDC = BeginPaint(hWnd, &ps);

			RECT rc;
			GetClientRect(hWnd, &rc);
			rc.left = *(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH"));
			rc.right = rc.left + 5;
			FillRect(hDC, &rc, (HBRUSH)GetProp(hWnd, TEXT("GRAYBRUSH")));
			EndPaint(hWnd, &ps);

			return 0;
		}
		break;

		// https://groups.google.com/g/comp.os.ms-windows.programmer.win32/c/1XhCKATRXws
		case WM_NCHITTEST: {
			return 1;
		}
		break;

		case WM_LBUTTONDOWN: {
			SetProp(hWnd, TEXT("ISMOUSEDOWN"), (HANDLE)1);
			SetCapture(hWnd);
			return 0;
		}
		break;

		case WM_LBUTTONUP: {
			ReleaseCapture();
			RemoveProp(hWnd, TEXT("ISMOUSEDOWN"));
		}
		break;

		case WM_MOUSEMOVE: {
			if (wParam != MK_LBUTTON || !GetProp(hWnd, TEXT("ISMOUSEDOWN")))
				return 0;

			DWORD x = GET_X_LPARAM(lParam);
			if (x > 0 && x < 32000)
				*(int*)GetProp(hWnd, TEXT("SPLITTERWIDTH")) = x;
			SendMessage(hWnd, WM_SIZE, 0, 0);
		}
		break;

		case WM_MOUSEWHEEL: {
			if (LOWORD(wParam) == MK_CONTROL) {
				SendMessage(hWnd, WMU_SET_FONT, GET_WHEEL_DELTA_WPARAM(wParam) > 0 ? 1: -1, 0);
				return 1;
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
		}
		break;

		case WM_NOTIFY : {
			NMHDR* pHdr = (LPNMHDR)lParam;
			if (pHdr->idFrom == IDC_TAB && pHdr->code == TCN_SELCHANGE) {
				HWND hTabWnd = pHdr->hwndFrom;
				BOOL isText = TabCtrl_GetCurSel(hTabWnd) == 1;
				ShowWindow(GetDlgItem(hTabWnd, IDC_GRID), isText ? SW_HIDE : SW_SHOW);
				ShowWindow(GetDlgItem(hTabWnd, IDC_TEXT), isText ? SW_SHOW : SW_HIDE);
			}

			if (pHdr->idFrom == IDC_TREE && pHdr->code == TVN_SELCHANGED)
				SendMessage(hWnd, WMU_UPDATE_GRID, 0, 0);

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
			}

			if (pHdr->idFrom == IDC_GRID && pHdr->code == (DWORD)LVN_KEYDOWN) {
				NMLVKEYDOWN* kd = (LPNMLVKEYDOWN) lParam;
				if (kd->wVKey == 0x43 && GetKeyState(VK_CONTROL)) // Ctrl + C
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

			HTREEITEM hItem = TreeView_GetSelection(hTreeWnd);
			TV_ITEM tv = {0};
			tv.mask = TVIF_PARAM;
			tv.hItem = hItem;
			TreeView_GetItem(hTreeWnd, &tv);
			JSON_Value* val = (JSON_Value*)tv.lParam;
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
			if (_tcslen(text16) < MAX_HIGHLIGHT_LENGTH)
				highlightBlock(hTextWnd, text16, 0);
			SendMessage(hTextWnd, EM_SETSEL, 0, 0);
			LockWindowUpdate(0);
			InvalidateRect(hTextWnd, NULL, TRUE);
			free(text16);

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
			for (int colNo = 0; colNo < colCount; colNo++) {
				// Use WS_BORDER to vertical text aligment
				HWND hEdit = CreateWindowEx(WS_EX_TOPMOST, WC_EDIT, NULL, ES_CENTER | ES_AUTOHSCROLL | WS_VISIBLE | WS_CHILD | WS_TABSTOP | WS_BORDER,
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

			if (rowCount > 0) {
				resultset = realloc(resultset, rowCount * sizeof(int));
				SetProp(hWnd, TEXT("RESULTSET"), (HANDLE)resultset);
				int orderBy = *pOrderBy;

				if (orderBy) {
					// Bubble-sort
					for (int i = 0; i < rowCount; i++) {
						for (int j = i + 1; j < rowCount; j++) {
							int a = resultset[i];
							int b = resultset[j];
							if(orderBy > 0 && _tcscmp(cache[a][orderBy - 1], cache[b][orderBy - 1]) > 0) {
								resultset[i] = b;
								resultset[j] = a;
							}

							if(orderBy < 0 && _tcscmp(cache[a][-orderBy - 1], cache[b][-orderBy - 1]) < 0) {
								resultset[i] = b;
								resultset[j] = a;
							}
						}
					}
				}
			} else {
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
				RECT rc;
				HDC hDC = GetDC(hHeader);
				DrawText(hDC, name16, _tcslen(name16), &rc, DT_NOCLIP | DT_CALCRECT);
				ReleaseDC(hHeader, hDC);

				int w = rc.right - rc.left + 10;
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

			HFONT hFont = CreateFont (*pFontSize, 0, 0, 0, FW_DONTCARE, FALSE, FALSE, FALSE, ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, DEFAULT_PITCH | FF_SWISS, TEXT("Arial"));
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

			GETTEXTLENGTHEX gtl = {GTL_NUMBYTES, 0};
			int len = SendMessage(hTextWnd, EM_GETTEXTLENGTHEX, (WPARAM)&gtl, 1200);
			if (len < MAX_HIGHLIGHT_LENGTH) {
				TCHAR* text = calloc(len + sizeof(TCHAR), sizeof(char));
				GETTEXTEX gt = {0};
				gt.cb = len + sizeof(TCHAR);
				gt.flags = 0;
				gt.codepage = 1200;
				SendMessage(hTextWnd, EM_GETTEXTEX, (WPARAM)&gt, (LPARAM)text);
				LockWindowUpdate(hTextWnd);
				highlightBlock(hTextWnd, text, 0);
				free(text);
				SendMessage(hTextWnd, EM_SETSEL, 0, 0);
				LockWindowUpdate(0);
				InvalidateRect(hTextWnd, NULL, TRUE);
			}

			PostMessage(hWnd, WMU_AUTO_COLUMN_SIZE, 0, 0);
		}
		break;

	}
	return CallWindowProc((WNDPROC)GetProp(hWnd, TEXT("WNDPROC")), hWnd, msg, wParam, lParam);
}

LRESULT CALLBACK cbNewFilterEdit(HWND hWnd, UINT msg, WPARAM wParam, LPARAM lParam) {
	WNDPROC cbDefault = (WNDPROC)GetProp(hWnd, TEXT("WNDPROC"));

	switch(msg){
		// Win10+ fix: draw an upper border
		case WM_PAINT: {
			cbDefault(hWnd, msg, wParam, lParam);

			RECT rc;
			GetWindowRect(hWnd, &rc);
			HDC hDC = GetWindowDC(hWnd);
			HPEN hPen = CreatePen(PS_SOLID, 1, GetSysColor(COLOR_BTNFACE));
			HPEN oldPen = SelectObject(hDC, hPen);
			MoveToEx(hDC, 1, 0, 0);
			LineTo(hDC, rc.right - 1, 0);
			SelectObject(hDC, oldPen);
			DeleteObject(hPen);
			ReleaseDC(hWnd, hDC);

			return 0;
		}
		break;

		// Prevent beep
		case WM_CHAR: {
			if (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB)
				return 0;
		}
		break;

		case WM_KEYDOWN: {
			if (wParam == VK_RETURN || wParam == VK_ESCAPE || wParam == VK_TAB) {
				if (wParam == VK_RETURN) {
					HWND hHeader = GetParent(hWnd);
					HWND hGridWnd = GetParent(hHeader);
					HWND hTabWnd = GetParent(hGridWnd);
					HWND hMainWnd = GetParent(hTabWnd);
					SendMessage(hMainWnd, WMU_UPDATE_RESULTSET, 0, 0);
				}

				return 0;
			}
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
		for (int itemNo = 0; itemNo < json_array_get_count(arr); itemNo++) {
			JSON_Value* val2 = json_array_get_value(arr, itemNo);
			TCHAR itemName[64];
			_sntprintf(itemName, 64, TEXT("[%i]"), itemNo);
			HTREEITEM hItem = TreeView_AddItem(hTreeWnd, itemName, hParentItem, (LPARAM)val2);
			addNode(hTreeWnd, hItem, val2);
		}

		return TRUE;
	}

	if (type == JSONObject) {
		JSON_Object* obj = json_value_get_object(val);
		for (int itemNo = 0; itemNo < json_object_get_count(obj); itemNo++) {
			JSON_Value* val2 = json_object_get_value_at(obj, itemNo);
			TCHAR* itemName = utf8to16(json_object_get_name(obj, itemNo));
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

			cf2.crTextColor = isKey ? RGB(128, 0, 128) : RGB(0, 128, 0);
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
				cf2.crTextColor = RGB(0, 0, 255);
				end += 4;
			} else if (_tcsstr(buf, TEXT("false")) == buf) {
				cf2.crTextColor = RGB(0, 0, 255);
				end += 5;
			} else if (_tcsstr(buf, TEXT("null")) == buf) {
				cf2.crTextColor = RGB(255, 0, 0);
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

void setStoredValue(TCHAR* name, int value) {
	TCHAR buf[128];
	_sntprintf(buf, 128, TEXT("%i"), value);
	WritePrivateProfileString(APP_NAME, name, buf, iniPath);
}

int getStoredValue(TCHAR* name, int defValue) {
	TCHAR buf[128];
	return GetPrivateProfileString(APP_NAME, name, NULL, buf, 128, iniPath) ? _ttoi(buf) : defValue;
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

HTREEITEM TreeView_AddItem (HWND hTreeWnd, TCHAR* caption, HTREEITEM parent, LPARAM lParam) {
	TVITEM tvi = {0};
	TVINSERTSTRUCT tvins = {0};
	tvi.mask = TVIF_TEXT | TVIF_PARAM;
	tvi.pszText = caption;
	tvi.cchTextMax = _tcslen(caption) + 1;
	tvi.lParam = lParam;

	tvins.item = tvi;
	tvins.hInsertAfter = TVI_LAST;
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
