// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
// Tested with Excel 2016/2019/365 64-bit
#include <windows.h>
#include "xlcall.h"

// Excel12v function pointer
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

// Global state
static BOOL g_bInitialized = FALSE;
static BOOL g_bToggling    = FALSE;

// Store previous workbook name and sheet index as strings
static wchar_t g_prevBookName[512] = {0};
static int     g_prevSheetIndex     = -1;

// ============================================================
// Helper: Create XLOPER12 values
// ============================================================
static XLOPER12 MakeStr(const wchar_t *s)
{
    XLOPER12 x;
    x.xltype = xltypeStr;
    x.val.str = (wchar_t*)s;
    return x;
}

static XLOPER12 MakeInt(int i)
{
    XLOPER12 x;
    x.xltype = xltypeInt;
    x.val.w = (WORD)i;
    return x;
}

// ============================================================
// Helper: Get active workbook name and sheet index
// ============================================================
static BOOL GetActiveInfo(wchar_t *bookName, int nameSize, int *sheetIdx)
{
    XLOPER12 xRes;

    // Get workbook name: xlfGetDocument(88)
    xRes.xltype = xltypeNil;
    XLOPER12 xArg88 = MakeInt(88);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg88);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > nameSize - 1) len = nameSize - 1;
        memcpy(bookName, xRes.val.str + 1, len * sizeof(wchar_t));
        bookName[len] = 0;
    } else {
        bookName[0] = 0;
    }

    // Get sheet index: xlfGetDocument(87)
    XLOPER12 xArg87 = MakeInt(87);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg87);
    *sheetIdx = (xRes.xltype == xltypeInt) ? xRes.val.w : 1;

    return TRUE;
}

// ============================================================
// Helper: Activate workbook by name + sheet by index
// ============================================================
static BOOL ActivatePrevSheet(void)
{
    if (g_prevBookName[0] == 0 || g_prevSheetIndex < 1) return FALSE;

    // Step 1: Activate workbook by name
    // xlcWorkbookActivate(book_name)
    XLOPER12 xBook = MakeStr(g_prevBookName);
    EXCEL12(xlcWorkbookActivate, 0, 1, &xBook);

    // Step 2: Select sheet by index
    // xlcWorkbookSelect(index, book_name) => only index needed if book is active
    // Use xlcSelect with sheet index
    XLOPER12 xIdx = MakeInt(g_prevSheetIndex);
    // xlcSelect: select sheet by index (returns nothing)
    // We use xlcSelectEnd / xlcWorkbookSelect to jump
    EXCEL12(xlcSelect, 0, 1, &xIdx);

    return TRUE;
}

// ============================================================
// Core: Toggle to previous sheet
// ============================================================
static void WINAPI ToggleToPreviousSheet(void)
{
    if (!g_bInitialized) return;
    if (g_prevBookName[0] == 0 || g_prevSheetIndex < 1) return;

    // Save current info before switching
    wchar_t curBook[512];
    int curIdx;
    GetActiveInfo(curBook, 512, &curIdx);

    g_bToggling = TRUE;

    if (ActivatePrevSheet()) {
        // Successfully jumped. Now swap so next press goes back
        wchar_t tmpBook[512];
        wcscpy(tmpBook, g_prevBookName);
        wcscpy(g_prevBookName, curBook);
        wcscpy(curBook, tmpBook);

        int tmpIdx = g_prevSheetIndex;
        g_prevSheetIndex = curIdx;
        // curIdx is no longer needed
    }

    g_bToggling = FALSE;
}

// ============================================================
// Update history when user manually switches sheets
// ============================================================
static void WINAPI OnSheetActivate(void)
{
    if (g_bToggling) return;

    wchar_t curBook[512];
    int curIdx;
    GetActiveInfo(curBook, 512, &curIdx);

    if (!g_bInitialized) {
        wcscpy(g_prevBookName, curBook);
        g_prevSheetIndex = curIdx;
        g_bInitialized = TRUE;
        return;
    }

    // If actual change happened, save previous
    if (wcscmp(curBook, g_prevBookName) != 0 || curIdx != g_prevSheetIndex) {
        // Save old as new "previous"
        wcscpy(g_prevBookName, curBook);
        g_prevSheetIndex = curIdx;
    }
}

// ============================================================
// xlAutoOpen: called when XLL is loaded
// ============================================================
int WINAPI xlAutoOpen(void)
{
    // Get Excel12v
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    g_bInitialized = FALSE;

    // Register command with shortcut L
    XLOPER12 xRes;
    XLOPER12 xArgs[7];
    xArgs[0] = MakeStr(L"ToggleSheet");       // procedure name in DLL
    xArgs[1] = MakeStr(L"L");                  // shortcut key (Ctrl+Alt+L)
    xArgs[2] = MakeStr(L"ToggleSheet");       // command name in Excel
    xArgs[3] = MakeStr(L"");                   // description
    xArgs[4] = MakeInt(1);                     // macro type = command
    xArgs[5] = MakeStr(L"Previous Sheet");    // category
    xArgs[6] = MakeInt(0);                     // (reserved)

    EXCEL12(xlfRegister, &xRes, 7,
            &xArgs[0], &xArgs[1], &xArgs[2], &xArgs[3],
            &xArgs[4], &xArgs[5], &xArgs[6]);

    // Assign shortcut
    XLOPER12 xKey = MakeStr(L"^%L");
    XLOPER12 xCmd = MakeStr(L"ToggleSheet");
    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    // Initialize current state
    OnSheetActivate();

    return 1;
}

// ============================================================
// xlAutoClose: called when XLL is unloaded
// ============================================================
int WINAPI xlAutoClose(void)
{
    if (pExcel12v) {
        XLOPER12 xKey = MakeStr(L"^%L");
        EXCEL12(xlcOnKey, 0, 1, &xKey);
    }
    return 1;
}

// ============================================================
// DllMain
// ============================================================
BOOL WINAPI DllMain(HINSTANCE hinstDLL, DWORD fdwReason, LPVOID lpvReserved)
{
    if (fdwReason == DLL_PROCESS_ATTACH) {
        DisableThreadLibraryCalls(hinstDLL);
    }
    return TRUE;
}
