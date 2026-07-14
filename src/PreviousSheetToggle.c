// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#include <windows.h>

// ============================================================
// Excel12v function pointer and helper
// ============================================================
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

// ============================================================
// XLOPER12 type constants and struct
// ============================================================
#define xltypeNum     0x0001
#define xltypeStr     0x0002
#define xltypeMissing 0x0080
#define xltypeNil     0x0100
#define xltypeInt     0x0800

typedef struct {
    union {
        double num;
        wchar_t *str;
        unsigned short w;
    } val;
    unsigned short xltype;
} XLOPER12;

// Excel function indices
#define xlfRegister         20001
#define xlfGetDocument      20032
#define xlcOnKey            30061
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

// ============================================================
// Global state
// ============================================================
static BOOL    g_bToggling = FALSE;
static BOOL    g_bHasPrev  = FALSE;
static wchar_t g_prevBook[512] = {0};
static wchar_t g_prevSheet[512] = {0};

// ============================================================
// Helpers to create XLOPER12 values
// ============================================================
static XLOPER12 OpStr(const wchar_t *s)
{
    XLOPER12 x = {{0}, xltypeStr};
    x.val.str = (wchar_t*)s;
    return x;
}

static XLOPER12 OpInt(int i)
{
    XLOPER12 x = {{0}, xltypeInt};
    x.val.w = (unsigned short)i;
    return x;
}

// ============================================================
// Get current workbook name and sheet name
// ============================================================
static void GetCurrent(wchar_t *book, int bsize, wchar_t *sheet, int ssize)
{
    XLOPER12 xArg, xRes;

    // Workbook name
    xArg = OpInt(88);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > bsize - 1) len = bsize - 1;
        memcpy(book, xRes.val.str + 1, len * sizeof(wchar_t));
        book[len] = 0;
    } else {
        book[0] = 0;
    }

    // Sheet name
    xArg = OpInt(76);
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > ssize - 1) len = ssize - 1;
        memcpy(sheet, xRes.val.str + 1, len * sizeof(wchar_t));
        sheet[len] = 0;
    } else {
        sheet[0] = 0;
    }
}

// ============================================================
// Activate a workbook and sheet
// ============================================================
static BOOL Activate(const wchar_t *book, const wchar_t *sheet)
{
    if (book[0] == 0 || sheet[0] == 0) return FALSE;

    // Activate workbook
    XLOPER12 xBook = OpStr(book);
    EXCEL12(xlcWorkbookActivate, 0, 1, &xBook);

    // Select sheet
    XLOPER12 xSheet = OpStr(sheet);
    EXCEL12(xlcSelect, 0, 1, &xSheet);

    return TRUE;
}

// ============================================================
// Toggle command - called by shortcut
// ============================================================
__declspec(dllexport) void WINAPI ToggleCmd(void)
{
    if (!g_bHasPrev) return;

    // Save current
    wchar_t curBook[512], curSheet[512];
    GetCurrent(curBook, 512, curSheet, 512);

    g_bToggling = TRUE;

    // Jump to previous
    if (Activate(g_prevBook, g_prevSheet)) {
        // Update "previous" to where we just came from
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
    }

    g_bToggling = FALSE;
}

// ============================================================
// Update history (call before toggle to record current position)
// ============================================================
static void UpdateHistory(void)
{
    if (g_bToggling) return;

    wchar_t curBook[512], curSheet[512];
    GetCurrent(curBook, 512, curSheet, 512);

    if (!g_bHasPrev) {
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
        g_bHasPrev = TRUE;
        return;
    }

    // If actually changed, record old position
    if (wcscmp(curBook, g_prevBook) != 0 || wcscmp(curSheet, g_prevSheet) != 0) {
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
    }
}

// ============================================================
// Another exported command that does history + toggle
// ============================================================
__declspec(dllexport) void WINAPI SwitchCmd(void)
{
    UpdateHistory();
    ToggleCmd();
}

// ============================================================
// xlAutoOpen
// ============================================================
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    // Register SwitchCmd command with shortcut L
    XLOPER12 xRes, xArgs[7];
    xArgs[0] = OpStr(L"SwitchCmd");              // DLL export name
    xArgs[1] = OpStr(L"L");                       // shortcut key
    xArgs[2] = OpStr(L"SwitchCmd");              // Excel command name
    xArgs[3] = OpStr(L"Switch to previous sheet"); // description
    xArgs[4] = OpInt(1);                          // macro type: command
    xArgs[5] = OpStr(L"Previous Sheet");         // category
    xArgs[6] = OpInt(0);

    EXCEL12(xlfRegister, &xRes, 7,
            &xArgs[0], &xArgs[1], &xArgs[2], &xArgs[3],
            &xArgs[4], &xArgs[5], &xArgs[6]);

    // Bind Ctrl+Alt+L
    XLOPER12 xKey = OpStr(L"^%L");
    XLOPER12 xCmd = OpStr(L"SwitchCmd");
    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    // Init history
    UpdateHistory();

    return 1;
}

// ============================================================
// xlAutoClose
// ============================================================
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    if (pExcel12v) {
        XLOPER12 xKey = OpStr(L"^%L");
        EXCEL12(xlcOnKey, 0, 1, &xKey);
    }
    return 1;
}

// ============================================================
// DllMain
// ============================================================
BOOL WINAPI DllMain(HINSTANCE h, DWORD r, LPVOID p)
{
    if (r == DLL_PROCESS_ATTACH) DisableThreadLibraryCalls(h);
    return TRUE;
}
