/*
 * cascextract - CASC Storage Browser and Extractor
 * A simple Win32 GUI application using CascLib
 */

#define _CRT_SECURE_NO_WARNINGS

/* CascLib.h includes windows.h with NOMINMAX and WIN32_LEAN_AND_MEAN */
#include <CascLib.h>
#include <commctrl.h>
#include <shlobj.h>
#include <ctype.h>
#include <stdarg.h>
#include <stdio.h>
#include <string.h>
#include <stdlib.h>

#pragma comment(lib, "comctl32.lib")
#pragma comment(lib, "shell32.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "advapi32.lib")

/* -----------------------------------------------------------------------
 * Control IDs
 * ----------------------------------------------------------------------- */
#define IDC_BTN_OPEN        100
#define IDC_BTN_CLOSE       101
#define IDC_BTN_EXTRACT     102
#define IDC_BTN_EXTRACT_ALL 103
#define IDC_BTN_ABOUT       109
#define IDC_LIST_FILES      104
#define IDC_STATIC_STATUS   105
#define IDC_EDIT_SEARCH     106
#define IDC_BTN_SEARCH      107
#define IDC_BTN_CLEAR       108

/* -----------------------------------------------------------------------
 * Custom window messages (posted from the worker thread)
 *
 *   WM_OPEN_DONE     lParam = OPEN_RESULT* (caller frees)
 *   WM_CASC_PROGRESS wParam = CASC_PROGRESS_MSG value
 * ----------------------------------------------------------------------- */
#define WM_OPEN_DONE     (WM_USER + 1)
#define WM_CASC_PROGRESS (WM_USER + 2)

/* -----------------------------------------------------------------------
 * Layout
 *
 *   Row 1 (y = BTN_MARGIN):              [Open] [Close] [Extract Sel] [Extract All]
 *   Row 2 (y = ROW2_Y):                  [<search edit>..............] [Search] [Clear]
 *   Listbox below TOOLBAR_H
 *   Status bar at the bottom
 * ----------------------------------------------------------------------- */
#define BTN_HEIGHT    26
#define BTN_WIDTH    150
#define BTN_MARGIN     6
#define SEARCH_BTN_W  72
#define CLEAR_BTN_W   55
#define ROW2_Y        (BTN_MARGIN + BTN_HEIGHT + BTN_MARGIN)
#define TOOLBAR_H     (ROW2_Y + BTN_HEIGHT + BTN_MARGIN)
#define STATUSBAR_H   22
#define STATUS_MARGIN  4

/* Registry key for persistent settings */
#define REG_KEY_PATH "SOFTWARE\\eezstreet\\cascextract"

/* 1 MB read buffer for extraction */
#define EXTRACT_CHUNK (1024 * 1024)

/* Initial pre-allocation sizes for the enumeration buffers */
#define INIT_FILE_COUNT  65536u
#define INIT_BUF_BYTES   (INIT_FILE_COUNT * 48u)

/* -----------------------------------------------------------------------
 * Data structures used by the background open thread
 * ----------------------------------------------------------------------- */

typedef struct _OPEN_PARAMS {
    char szPath[MAX_PATH];
    HWND hwndMain;
} OPEN_PARAMS;

typedef struct _OPEN_RESULT {
    HANDLE  hStorage;   /* valid handle, or NULL on failure */
    DWORD   dwError;    /* GetCascError() when hStorage == NULL */
    char   *pNameBuf;   /* flat buffer: all NUL-terminated file names */
    size_t  nBufUsed;   /* bytes used in pNameBuf */
    char  **ppNames;    /* pointers into pNameBuf; length = nFiles */
    DWORD   nFiles;
} OPEN_RESULT;

/* -----------------------------------------------------------------------
 * Global state
 *
 * The master file list (g_pNameBuf / g_ppNames / g_nFiles / g_nBufUsed)
 * is populated once when a storage is opened and freed when it is closed.
 * The search filter re-uses these arrays in-place; it never modifies them.
 * ----------------------------------------------------------------------- */
static HANDLE  g_hStorage   = NULL;
static char   *g_pNameBuf   = NULL;   /* master flat name buffer          */
static char  **g_ppNames    = NULL;   /* master pointer array             */
static DWORD   g_nFiles     = 0;      /* total file count                 */
static size_t  g_nBufUsed   = 0;      /* bytes used in g_pNameBuf         */

static HWND    g_hwndList   = NULL;
static HWND    g_hwndStatus = NULL;
static HWND    g_hwndMain   = NULL;
static HWND    g_hwndSearch = NULL;
static HFONT   g_hListFont  = NULL;

static char    g_szLastStoragePath[MAX_PATH];
static char    g_szLastExtractPath[MAX_PATH];

/* -----------------------------------------------------------------------
 * Registry helpers
 * ----------------------------------------------------------------------- */

static void RegLoadString(const char *name, char *out, DWORD size)
{
    HKEY  hKey;
    DWORD dwType = REG_SZ;
    DWORD dwSize = size;

    out[0] = '\0';
    if (RegOpenKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, KEY_READ, &hKey) == ERROR_SUCCESS)
    {
        RegQueryValueExA(hKey, name, NULL, &dwType, (LPBYTE)out, &dwSize);
        RegCloseKey(hKey);
    }
}

static void RegSaveString(const char *name, const char *value)
{
    HKEY hKey;
    if (RegCreateKeyExA(HKEY_CURRENT_USER, REG_KEY_PATH, 0, NULL,
                        REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS)
    {
        RegSetValueExA(hKey, name, 0, REG_SZ,
                       (const BYTE *)value, (DWORD)(strlen(value) + 1));
        RegCloseKey(hKey);
    }
}

/* -----------------------------------------------------------------------
 * Helpers
 * ----------------------------------------------------------------------- */

static void SetStatus(const char *fmt, ...)
{
    char    buf[512];
    va_list args;
    va_start(args, fmt);
    _vsnprintf_s(buf, sizeof(buf), _TRUNCATE, fmt, args);
    va_end(args);
    SetWindowTextA(g_hwndStatus, buf);
}

static int CALLBACK BrowseCallbackProc(HWND hwnd, UINT uMsg, LPARAM lParam, LPARAM lpData)
{
    (void)lParam;
    if (uMsg == BFFM_INITIALIZED && lpData)
        SendMessageA(hwnd, BFFM_SETSELECTION, TRUE, lpData);
    return 0;
}

static BOOL BrowseForFolder(HWND hwnd, const char *title, char *outPath, const char *initialPath)
{
    BROWSEINFOA  bi;
    LPITEMIDLIST pidl;

    memset(&bi, 0, sizeof(bi));
    bi.hwndOwner = hwnd;
    bi.lpszTitle = title;
    bi.ulFlags   = BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE;

    if (initialPath && initialPath[0])
    {
        bi.lpfn  = BrowseCallbackProc;
        bi.lParam = (LPARAM)initialPath;
    }

    pidl = SHBrowseForFolderA(&bi);
    if (pidl)
    {
        BOOL ok = SHGetPathFromIDListA(pidl, outPath);
        CoTaskMemFree(pidl);
        return ok;
    }
    return FALSE;
}

static void MakeDirs(const char *path)
{
    char  tmp[MAX_PATH * 2];
    char *p;

    strncpy_s(tmp, sizeof(tmp), path, _TRUNCATE);
    for (p = tmp + 1; *p; p++)
    {
        if (*p == '\\' || *p == '/')
        {
            char c = *p;
            *p = '\0';
            CreateDirectoryA(tmp, NULL);
            *p = c;
        }
    }
    CreateDirectoryA(tmp, NULL);
}

/* -----------------------------------------------------------------------
 * Fuzzy match
 *
 * Returns TRUE if every character of 'needle' appears in 'haystack' in
 * order (case-insensitive).  Consecutive matches are not required.
 * E.g. needle="inface" matches haystack="Interface\Glues\Main.blp".
 * ----------------------------------------------------------------------- */
static BOOL FuzzyMatch(const char *haystack, const char *needle)
{
    while (*needle)
    {
        char cn = (char)tolower((unsigned char)*needle);
        while (*haystack && (char)tolower((unsigned char)*haystack) != cn)
            ++haystack;
        if (!*haystack)
            return FALSE;
        ++haystack;
        ++needle;
    }
    return TRUE;
}

/* -----------------------------------------------------------------------
 * Filter / populate the listbox from the master name list.
 * Called whenever the search text changes (after the debounce timer fires)
 * or when the Search button is clicked.
 * ----------------------------------------------------------------------- */
static void ApplyFilter(void)
{
    char   szQuery[256];
    DWORD  i;
    DWORD  nShown    = 0;
    size_t nBufEst   = 0;

    if (!g_ppNames || g_nFiles == 0)
        return;

    GetWindowTextA(g_hwndSearch, szQuery, sizeof(szQuery));

    SendMessage(g_hwndList, WM_SETREDRAW,    FALSE, 0);
    SendMessage(g_hwndList, LB_RESETCONTENT, 0,     0);

    if (szQuery[0] == '\0')
    {
        /* Empty query: restore the full list with exact pre-allocation */
        SendMessage(g_hwndList, LB_INITSTORAGE, (WPARAM)g_nFiles, (LPARAM)g_nBufUsed);
        for (i = 0; i < g_nFiles; i++)
            SendMessageA(g_hwndList, LB_ADDSTRING, 0, (LPARAM)g_ppNames[i]);
        nShown = g_nFiles;
    }
    else
    {
        /*
         * Two-pass fuzzy filter:
         *   Pass 1 - count matches and total string bytes for LB_INITSTORAGE.
         *   Pass 2 - insert the matched strings.
         * Double-iterating the master list is cheap compared to the listbox
         * operations that follow, and LB_INITSTORAGE eliminates repeated
         * internal reallocs during insertion.
         */
        for (i = 0; i < g_nFiles; i++)
        {
            if (FuzzyMatch(g_ppNames[i], szQuery))
            {
                nShown++;
                nBufEst += strlen(g_ppNames[i]) + 1;
            }
        }

        SendMessage(g_hwndList, LB_INITSTORAGE, (WPARAM)nShown, (LPARAM)nBufEst);

        for (i = 0; i < g_nFiles; i++)
        {
            if (FuzzyMatch(g_ppNames[i], szQuery))
                SendMessageA(g_hwndList, LB_ADDSTRING, 0, (LPARAM)g_ppNames[i]);
        }
    }

    SendMessage(g_hwndList, LB_SETHORIZONTALEXTENT, 2000, 0);
    SendMessage(g_hwndList, WM_SETREDRAW, TRUE,  0);
    InvalidateRect(g_hwndList, NULL, TRUE);

    if (szQuery[0] == '\0')
        SetStatus("%u file(s).", g_nFiles);
    else
        SetStatus("Filter: %u / %u file(s) match \"%s\".", nShown, g_nFiles, szQuery);
}

/* -----------------------------------------------------------------------
 * Storage helpers
 * ----------------------------------------------------------------------- */

static void EnableStorageButtons(BOOL bHaveStorage)
{
    EnableWindow(GetDlgItem(g_hwndMain, IDC_BTN_CLOSE),       bHaveStorage);
    EnableWindow(GetDlgItem(g_hwndMain, IDC_BTN_EXTRACT),     bHaveStorage);
    EnableWindow(GetDlgItem(g_hwndMain, IDC_BTN_EXTRACT_ALL), bHaveStorage);
    EnableWindow(GetDlgItem(g_hwndMain, IDC_EDIT_SEARCH),     bHaveStorage);
    EnableWindow(GetDlgItem(g_hwndMain, IDC_BTN_SEARCH),      bHaveStorage);
    EnableWindow(GetDlgItem(g_hwndMain, IDC_BTN_CLEAR),       bHaveStorage);
}

static void CloseCurrentStorage(void)
{
    if (g_hStorage)
    {
        CascCloseStorage(g_hStorage);
        g_hStorage = NULL;
    }

    free(g_ppNames);  g_ppNames  = NULL;
    free(g_pNameBuf); g_pNameBuf = NULL;
    g_nFiles   = 0;
    g_nBufUsed = 0;

    SendMessage(g_hwndList, LB_RESETCONTENT, 0, 0);
    SetWindowTextA(g_hwndSearch, "");
    EnableStorageButtons(FALSE);
    SetStatus("Ready.");
}

/* -----------------------------------------------------------------------
 * Progress callback - runs on the worker thread, posts to the UI thread.
 * ----------------------------------------------------------------------- */
static BOOL WINAPI CascProgressCb(void *pParam, CASC_PROGRESS_MSG eMsg,
                                   LPCSTR szObject, DWORD dwCurrent, DWORD dwTotal)
{
    (void)szObject; (void)dwCurrent; (void)dwTotal;
    PostMessageA((HWND)pParam, WM_CASC_PROGRESS, (WPARAM)eMsg, 0);
    return FALSE;
}

/* -----------------------------------------------------------------------
 * Worker thread: opens the storage and enumerates all file names into a
 * flat heap buffer.  Sends WM_OPEN_DONE when finished.
 * ----------------------------------------------------------------------- */
static DWORD WINAPI OpenStorageThread(LPVOID lpParam)
{
    OPEN_PARAMS           *pParams = (OPEN_PARAMS *)lpParam;
    OPEN_RESULT           *pResult;
    CASC_OPEN_STORAGE_ARGS args;
    CASC_FIND_DATA         fd;
    HANDLE  hStorage  = NULL;
    HANDLE  hFind;
    char   *pBuf      = NULL;
    char  **ppNames   = NULL;
    size_t  nBufCap   = 0;
    size_t  nBufUsed  = 0;
    DWORD   nNameCap  = 0;
    DWORD   nFiles    = 0;
    size_t  nEstimate = 0;
    BOOL    bFailed   = FALSE;

    pResult = (OPEN_RESULT *)calloc(1, sizeof(OPEN_RESULT));
    if (!pResult) { free(pParams); return 1; }

    /* Open storage with a progress callback */
    memset(&args, 0, sizeof(args));
    args.Size                = sizeof(args);
    args.dwLocaleMask        = CASC_LOCALE_ALL;
    args.PfnProgressCallback = CascProgressCb;
    args.PtrProgressParam    = (void *)pParams->hwndMain;
	args.dwFlags = CASC_FEATURE_ALLOW_DOWNLOAD;

    if (!CascOpenStorageEx(pParams->szPath, &args, FALSE, &hStorage))
    {
        pResult->dwError = GetCascError();
        PostMessageA(pParams->hwndMain, WM_OPEN_DONE, 0, (LPARAM)pResult);
        free(pParams);
        return 1;
    }

    pResult->hStorage = hStorage;

    /* Pre-allocate using the known file count */
    CascGetStorageInfo(hStorage, CascStorageTotalFileCount,
                       &nEstimate, sizeof(nEstimate), NULL);
    if (nEstimate < INIT_FILE_COUNT) nEstimate = INIT_FILE_COUNT;

    nNameCap = (DWORD)nEstimate;
    nBufCap  = nEstimate * 48;
    if (nBufCap < INIT_BUF_BYTES) nBufCap = INIT_BUF_BYTES;

    ppNames = (char **)malloc(nNameCap * sizeof(char *));
    pBuf    = (char  *)malloc(nBufCap);
    if (!ppNames || !pBuf) { bFailed = TRUE; goto done; }

    /* Enumerate all files into the flat buffer */
    hFind = CascFindFirstFile(hStorage, "*", &fd, NULL);
    if (hFind != INVALID_HANDLE_VALUE)
    {
        do
        {
            size_t nameLen = strlen(fd.szFileName) + 1;

            /* Grow the name buffer, fixing up existing pointers if it moves */
            if (nBufUsed + nameLen > nBufCap)
            {
                char     *newBuf;
                size_t    newCap = (nBufUsed + nameLen) * 2;
                DWORD     i;

                newBuf = (char *)realloc(pBuf, newCap);
                if (!newBuf) { bFailed = TRUE; break; }
                if (newBuf != pBuf)
                {
                    ptrdiff_t delta = newBuf - pBuf;
                    for (i = 0; i < nFiles; i++) ppNames[i] += delta;
                }
                pBuf    = newBuf;
                nBufCap = newCap;
            }

            /* Grow the pointer array */
            if (nFiles >= nNameCap)
            {
                char **newPtrs;
                nNameCap *= 2;
                newPtrs = (char **)realloc(ppNames, nNameCap * sizeof(char *));
                if (!newPtrs) { bFailed = TRUE; break; }
                ppNames = newPtrs;
            }

            ppNames[nFiles] = pBuf + nBufUsed;
            memcpy(pBuf + nBufUsed, fd.szFileName, nameLen);
            nBufUsed += nameLen;
            nFiles++;
        }
        while (CascFindNextFile(hFind, &fd));

        CascFindClose(hFind);
    }

done:
    if (bFailed)
    {
        free(pBuf); free(ppNames);
        CascCloseStorage(hStorage);
        pResult->hStorage = NULL;
        pResult->dwError  = ERROR_NOT_ENOUGH_MEMORY;
    }
    else
    {
        pResult->pNameBuf = pBuf;
        pResult->nBufUsed = nBufUsed;
        pResult->ppNames  = ppNames;
        pResult->nFiles   = nFiles;
    }

    PostMessageA(pParams->hwndMain, WM_OPEN_DONE, 0, (LPARAM)pResult);
    free(pParams);
    return 0;
}

static void DoOpenStorage(HWND hwnd)
{
    OPEN_PARAMS *pParams;
    HANDLE       hThread;
    char         szPath[MAX_PATH];

    if (!BrowseForFolder(hwnd, "Select CASC Storage Folder (containing .build.info)", szPath, g_szLastStoragePath))
        return;

    strncpy_s(g_szLastStoragePath, sizeof(g_szLastStoragePath), szPath, _TRUNCATE);
    RegSaveString("LastStoragePath", szPath);

    CloseCurrentStorage();

    pParams = (OPEN_PARAMS *)malloc(sizeof(OPEN_PARAMS));
    if (!pParams) return;
    strncpy_s(pParams->szPath, sizeof(pParams->szPath), szPath, _TRUNCATE);
    pParams->hwndMain = hwnd;

    EnableWindow(GetDlgItem(hwnd, IDC_BTN_OPEN), FALSE);
    SetStatus("Opening storage: %s ...", szPath);

    hThread = CreateThread(NULL, 0, OpenStorageThread, pParams, 0, NULL);
    if (hThread)
    {
        CloseHandle(hThread);
    }
    else
    {
        free(pParams);
        EnableWindow(GetDlgItem(hwnd, IDC_BTN_OPEN), TRUE);
        SetStatus("Failed to create worker thread.");
    }
}

/* -----------------------------------------------------------------------
 * File extraction
 * ----------------------------------------------------------------------- */

static BOOL ExtractCascFile(const char *szFileName, const char *szOutDir)
{
    HANDLE     hCascFile = NULL;
    FILE      *hOut      = NULL;
    BYTE      *pBuf      = NULL;
    char       szOutPath[MAX_PATH * 2];
    char       szDirPart[MAX_PATH * 2];
    char       szFixed[MAX_PATH];
    char      *pSlash;
    char      *p;
    ULONGLONG  ullSize   = 0;
    ULONGLONG  ullDone   = 0;
    DWORD      dwRead;
    DWORD      dwWant;
    BOOL       bOK       = FALSE;

    /* Normalise separators and replace Windows-illegal characters */
    strncpy_s(szFixed, sizeof(szFixed), szFileName, _TRUNCATE);

	char* colon = strchr(szFixed, ':');
	if (colon && *(colon + 1))
		memmove(szFixed, colon + 1, strlen(colon + 1) + 1);

    for (p = szFixed; *p; p++)
    {
        if (*p == '/')
            *p = '\\';
        else if (*p == ':' || *p == '*' || *p == '?' ||
                 *p == '"' || *p == '<' || *p == '>' || *p == '|')
            *p = '_';
    }

    _snprintf_s(szOutPath, sizeof(szOutPath), _TRUNCATE, "%s\\%s", szOutDir, szFixed);

    strncpy_s(szDirPart, sizeof(szDirPart), szOutPath, _TRUNCATE);
    pSlash = strrchr(szDirPart, '\\');
    if (pSlash) { *pSlash = '\0'; MakeDirs(szDirPart); }

    if (!CascOpenFile(g_hStorage, szFileName, CASC_LOCALE_ALL, CASC_OVERCOME_ENCRYPTED, &hCascFile))
        return FALSE;

    if (!CascGetFileSize64(hCascFile, &ullSize))
    {
        CascCloseFile(hCascFile);
        return FALSE;
    }

    if (fopen_s(&hOut, szOutPath, "wb") != 0)
    {
        CascCloseFile(hCascFile);
        return FALSE;
    }

    pBuf = (BYTE *)malloc(EXTRACT_CHUNK);
    if (!pBuf) { fclose(hOut); CascCloseFile(hCascFile); return FALSE; }

    bOK = TRUE;
    while (ullDone < ullSize)
    {
        ULONGLONG ullRemain = ullSize - ullDone;
        dwWant = (ullRemain > EXTRACT_CHUNK) ? EXTRACT_CHUNK : (DWORD)ullRemain;

        if (!CascReadFile(hCascFile, pBuf, dwWant, &dwRead) || dwRead == 0)
        {
            bOK = FALSE;
            break;
        }
        fwrite(pBuf, 1, dwRead, hOut);
        ullDone += dwRead;
    }

    free(pBuf);
    fclose(hOut);
    CascCloseFile(hCascFile);

    if (!bOK) DeleteFileA(szOutPath);
    return bOK;
}

static void DoExtractSelected(HWND hwnd)
{
    char  szOutDir[MAX_PATH];
    char  szFileName[MAX_PATH];
    int  *pSel  = NULL;
    int   nSel  = 0;
    int   nOK   = 0;
    int   nFail = 0;
    int   i;
    char  msg[256];

    if (!g_hStorage) return;

    nSel = (int)SendMessage(g_hwndList, LB_GETSELCOUNT, 0, 0);
    if (nSel <= 0)
    {
        MessageBoxA(hwnd, "No files selected.", "Extract Selected", MB_OK | MB_ICONINFORMATION);
        return;
    }

    if (!BrowseForFolder(hwnd, "Select Output Folder", szOutDir, g_szLastExtractPath)) return;

    strncpy_s(g_szLastExtractPath, sizeof(g_szLastExtractPath), szOutDir, _TRUNCATE);
    RegSaveString("LastExtractPath", szOutDir);

    pSel = (int *)malloc(nSel * sizeof(int));
    if (!pSel) return;
    SendMessage(g_hwndList, LB_GETSELITEMS, (WPARAM)nSel, (LPARAM)pSel);

    for (i = 0; i < nSel; i++)
    {
        SendMessageA(g_hwndList, LB_GETTEXT, (WPARAM)pSel[i], (LPARAM)szFileName);
        SetStatus("Extracting %d / %d: %s", i + 1, nSel, szFileName);
        UpdateWindow(hwnd);
        if (ExtractCascFile(szFileName, szOutDir)) nOK++; else nFail++;
    }

    free(pSel);
    _snprintf_s(msg, sizeof(msg), _TRUNCATE,
        "Extraction complete.\n\n%d file(s) OK.\n%d file(s) failed.", nOK, nFail);
    SetStatus("Done. %d OK, %d failed.", nOK, nFail);
    MessageBoxA(hwnd, msg, "Extract Selected",
        MB_OK | (nFail ? MB_ICONWARNING : MB_ICONINFORMATION));
}

static void DoExtractAll(HWND hwnd)
{
    char  szOutDir[MAX_PATH];
    char  szFileName[MAX_PATH];
    int   nTotal;
    int   nOK   = 0;
    int   nFail = 0;
    int   i;
    char  msg[256];

    if (!g_hStorage) return;

    nTotal = (int)SendMessage(g_hwndList, LB_GETCOUNT, 0, 0);
    if (nTotal <= 0) return;

    if (!BrowseForFolder(hwnd, "Select Output Folder", szOutDir, g_szLastExtractPath)) return;

    strncpy_s(g_szLastExtractPath, sizeof(g_szLastExtractPath), szOutDir, _TRUNCATE);
    RegSaveString("LastExtractPath", szOutDir);

    for (i = 0; i < nTotal; i++)
    {
        SendMessageA(g_hwndList, LB_GETTEXT, (WPARAM)i, (LPARAM)szFileName);
        if (i % 25 == 0)
        {
            SetStatus("Extracting %d / %d: %s", i + 1, nTotal, szFileName);
            UpdateWindow(hwnd);
        }
        if (ExtractCascFile(szFileName, szOutDir)) nOK++; else nFail++;
    }

    _snprintf_s(msg, sizeof(msg), _TRUNCATE,
        "Extraction complete.\n\n%d file(s) OK.\n%d file(s) failed.", nOK, nFail);
    SetStatus("Done. %d OK, %d failed.", nOK, nFail);
    MessageBoxA(hwnd, msg, "Extract All",
        MB_OK | (nFail ? MB_ICONWARNING : MB_ICONINFORMATION));
}

/* -----------------------------------------------------------------------
 * Resize the search-row controls to fill the window width.
 * Called from WM_SIZE and once after WM_CREATE.
 * ----------------------------------------------------------------------- */
static void LayoutSearchRow(int cx)
{
    int editW  = cx - SEARCH_BTN_W - CLEAR_BTN_W - BTN_MARGIN * 4;
    int btnX   = BTN_MARGIN + editW + BTN_MARGIN;
    int clearX = btnX + SEARCH_BTN_W + BTN_MARGIN;

    if (editW < 1) editW = 1;

    SetWindowPos(GetDlgItem(g_hwndMain, IDC_EDIT_SEARCH), NULL,
        BTN_MARGIN, ROW2_Y, editW, BTN_HEIGHT, SWP_NOZORDER);
    SetWindowPos(GetDlgItem(g_hwndMain, IDC_BTN_SEARCH), NULL,
        btnX, ROW2_Y, SEARCH_BTN_W, BTN_HEIGHT, SWP_NOZORDER);
    SetWindowPos(GetDlgItem(g_hwndMain, IDC_BTN_CLEAR), NULL,
        clearX, ROW2_Y, CLEAR_BTN_W, BTN_HEIGHT, SWP_NOZORDER);
}

/* -----------------------------------------------------------------------
 * Window procedure
 * ----------------------------------------------------------------------- */

LRESULT CALLBACK WndProc(HWND hwnd, UINT msg, WPARAM wParam, LPARAM lParam)
{
    switch (msg)
    {
    case WM_CREATE:
    {
        int x = BTN_MARGIN;

        /* Row 1: storage buttons */
        CreateWindowA("BUTTON", "Open Storage...",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            x, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_OPEN, GetModuleHandleA(NULL), NULL);
        x += BTN_WIDTH + BTN_MARGIN;

        CreateWindowA("BUTTON", "Close Storage",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
            x, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_CLOSE, GetModuleHandleA(NULL), NULL);
        x += BTN_WIDTH + BTN_MARGIN;

        CreateWindowA("BUTTON", "Extract Selected",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
            x, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_EXTRACT, GetModuleHandleA(NULL), NULL);
        x += BTN_WIDTH + BTN_MARGIN;

        CreateWindowA("BUTTON", "Extract All",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
            x, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_EXTRACT_ALL, GetModuleHandleA(NULL), NULL);
        x += BTN_WIDTH + BTN_MARGIN;

        CreateWindowA("BUTTON", "About",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON,
            x, BTN_MARGIN, BTN_WIDTH, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_ABOUT, GetModuleHandleA(NULL), NULL);

        /* Row 2: search bar (positions corrected on first WM_SIZE) */
        g_hwndSearch = CreateWindowA("EDIT", "",
            WS_CHILD | WS_VISIBLE | WS_BORDER | ES_AUTOHSCROLL | WS_DISABLED,
            BTN_MARGIN, ROW2_Y, 100, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_EDIT_SEARCH, GetModuleHandleA(NULL), NULL);

        /* Use the same font as the list for the search box */
        g_hListFont = CreateFontA(14, 0, 0, 0, FW_NORMAL, FALSE, FALSE, FALSE,
            ANSI_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS,
            DEFAULT_QUALITY, FIXED_PITCH | FF_MODERN, "Consolas");

        if (g_hListFont)
            SendMessage(g_hwndSearch, WM_SETFONT, (WPARAM)g_hListFont, TRUE);

        CreateWindowA("BUTTON", "Search",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
            0, ROW2_Y, SEARCH_BTN_W, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_SEARCH, GetModuleHandleA(NULL), NULL);

        CreateWindowA("BUTTON", "Clear",
            WS_CHILD | WS_VISIBLE | BS_PUSHBUTTON | WS_DISABLED,
            0, ROW2_Y, CLEAR_BTN_W, BTN_HEIGHT,
            hwnd, (HMENU)(UINT_PTR)IDC_BTN_CLEAR, GetModuleHandleA(NULL), NULL);

        /* File list */
        g_hwndList = CreateWindowA("LISTBOX", NULL,
            WS_CHILD | WS_VISIBLE | WS_BORDER | WS_VSCROLL | WS_HSCROLL |
            LBS_MULTIPLESEL | LBS_EXTENDEDSEL | LBS_HASSTRINGS | LBS_NOINTEGRALHEIGHT,
            0, TOOLBAR_H, 100, 100,
            hwnd, (HMENU)(UINT_PTR)IDC_LIST_FILES, GetModuleHandleA(NULL), NULL);

        if (g_hListFont)
            SendMessage(g_hwndList, WM_SETFONT, (WPARAM)g_hListFont, TRUE);

        /* Status bar */
        g_hwndStatus = CreateWindowA("STATIC", "Ready.",
            WS_CHILD | WS_VISIBLE | SS_LEFT | SS_NOPREFIX,
            STATUS_MARGIN, 0, 100, STATUSBAR_H,
            hwnd, (HMENU)(UINT_PTR)IDC_STATIC_STATUS, GetModuleHandleA(NULL), NULL);

        g_hwndMain = hwnd;
        break;
    }

    case WM_SIZE:
    {
        int cx    = (int)LOWORD(lParam);
        int cy    = (int)HIWORD(lParam);
        int listH = cy - TOOLBAR_H - STATUSBAR_H;
        if (listH < 0) listH = 0;

        LayoutSearchRow(cx);

        if (g_hwndList)
            SetWindowPos(g_hwndList, NULL, 0, TOOLBAR_H, cx, listH, SWP_NOZORDER);
        if (g_hwndStatus)
            SetWindowPos(g_hwndStatus, NULL, STATUS_MARGIN, cy - STATUSBAR_H + 2,
                         cx - STATUS_MARGIN, STATUSBAR_H, SWP_NOZORDER);
        break;
    }

    /* Progress updates from the worker thread */
    case WM_CASC_PROGRESS:
        switch ((CASC_PROGRESS_MSG)wParam)
        {
        case CascProgressLoadingFile:               SetStatus("Loading file...");                break;
        case CascProgressLoadingManifest:           SetStatus("Loading manifest...");            break;
        case CascProgressDownloadingFile:           SetStatus("Downloading file...");            break;
        case CascProgressLoadingIndexes:            SetStatus("Loading index files...");         break;
        case CascProgressDownloadingArchiveIndexes: SetStatus("Downloading archive indexes..."); break;
        default:                                    SetStatus("Opening storage...");             break;
        }
        break;

    /* Worker thread finished */
    case WM_OPEN_DONE:
    {
        OPEN_RESULT *pRes = (OPEN_RESULT *)lParam;
        DWORD        i;

        EnableWindow(GetDlgItem(hwnd, IDC_BTN_OPEN), TRUE);

        if (!pRes || !pRes->hStorage)
        {
            char errmsg[256];
            _snprintf_s(errmsg, sizeof(errmsg), _TRUNCATE,
                "Failed to open CASC storage.\nError code: %u",
                pRes ? pRes->dwError : 0u);
            MessageBoxA(hwnd, errmsg, "Open Storage", MB_OK | MB_ICONERROR);
            free(pRes);
            SetStatus("Failed to open storage.");
            break;
        }

        /* Take ownership of the master name list */
        g_hStorage  = pRes->hStorage;
        g_pNameBuf  = pRes->pNameBuf;
        g_ppNames   = pRes->ppNames;
        g_nFiles    = pRes->nFiles;
        g_nBufUsed  = pRes->nBufUsed;
        free(pRes); /* struct only; pNameBuf/ppNames now owned by globals */

        /* Populate the listbox with LB_INITSTORAGE for one-shot allocation */
        SetStatus("Populating file list (%u files) ...", g_nFiles);
        UpdateWindow(hwnd);

        SendMessage(g_hwndList, WM_SETREDRAW,    FALSE,        0);
        SendMessage(g_hwndList, LB_RESETCONTENT, 0,            0);
        SendMessage(g_hwndList, LB_INITSTORAGE,  g_nFiles,     (LPARAM)g_nBufUsed);

        for (i = 0; i < g_nFiles; i++)
            SendMessageA(g_hwndList, LB_ADDSTRING, 0, (LPARAM)g_ppNames[i]);

        SendMessage(g_hwndList, LB_SETHORIZONTALEXTENT, 2000, 0);
        SendMessage(g_hwndList, WM_SETREDRAW, TRUE,  0);
        InvalidateRect(g_hwndList, NULL, TRUE);

        EnableStorageButtons(TRUE);
        SetStatus("%u file(s).", g_nFiles);
        break;
    }

    case WM_COMMAND:
        switch (LOWORD(wParam))
        {
        case IDC_BTN_OPEN:        DoOpenStorage(hwnd);     break;
        case IDC_BTN_CLOSE:       CloseCurrentStorage();   break;
        case IDC_BTN_EXTRACT:     DoExtractSelected(hwnd); break;
        case IDC_BTN_EXTRACT_ALL: DoExtractAll(hwnd);      break;

        case IDC_BTN_ABOUT:
            MessageBoxA(hwnd,
                "CASC Storage Extractor\n"
                "Version 1.01\n\n"
                "Author: eezstreet",
                "About", MB_OK | MB_ICONINFORMATION);
            break;

        case IDC_BTN_SEARCH:
            /* Immediate search on button click */
            ApplyFilter();
            break;

        case IDC_BTN_CLEAR:
            /* Clearing the text triggers EN_CHANGE -> debounce -> ApplyFilter */
            SetWindowTextA(g_hwndSearch, "");
            SetFocus(g_hwndSearch);
            break;

        }
        break;

    case WM_DESTROY:
        CloseCurrentStorage();
        if (g_hListFont) { DeleteObject(g_hListFont); g_hListFont = NULL; }
        PostQuitMessage(0);
        break;

    default:
        return DefWindowProcA(hwnd, msg, wParam, lParam);
    }

    return 0;
}

/* -----------------------------------------------------------------------
 * WinMain
 * ----------------------------------------------------------------------- */

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance,
                   LPSTR lpCmdLine, int nCmdShow)
{
    WNDCLASSA wc;
    HWND      hwnd;
    MSG       msg;

    (void)hPrevInstance;
    (void)lpCmdLine;

    CoInitialize(NULL);
    InitCommonControls();

    RegLoadString("LastStoragePath", g_szLastStoragePath, sizeof(g_szLastStoragePath));
    RegLoadString("LastExtractPath", g_szLastExtractPath, sizeof(g_szLastExtractPath));

    memset(&wc, 0, sizeof(wc));
    wc.style         = CS_HREDRAW | CS_VREDRAW;
    wc.lpfnWndProc   = WndProc;
    wc.hInstance     = hInstance;
    wc.hIcon         = LoadIconA(NULL, (LPCSTR)IDI_APPLICATION);
    wc.hCursor       = LoadCursorA(NULL, (LPCSTR)IDC_ARROW);
    wc.hbrBackground = (HBRUSH)(COLOR_BTNFACE + 1);
    wc.lpszClassName = "CascExtractWnd";

    if (!RegisterClassA(&wc))
    {
        MessageBoxA(NULL, "RegisterClass failed.", "Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    hwnd = CreateWindowA("CascExtractWnd", "CASC Storage Extractor",
        WS_OVERLAPPEDWINDOW, CW_USEDEFAULT, CW_USEDEFAULT, 960, 680,
        NULL, NULL, hInstance, NULL);

    if (!hwnd)
    {
        MessageBoxA(NULL, "CreateWindow failed.", "Error", MB_OK | MB_ICONERROR);
        return 1;
    }

    ShowWindow(hwnd, nCmdShow);
    UpdateWindow(hwnd);

    while (GetMessageA(&msg, NULL, 0, 0))
    {
        TranslateMessage(&msg);
        DispatchMessageA(&msg);
    }

    CoUninitialize();
    return (int)msg.wParam;
}
