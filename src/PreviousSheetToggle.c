// PreviousSheetToggle.c
// Excel XLL - Quick switch to previous sheet (Ctrl+Alt+L)
#include <windows.h>

// Excel12v function pointer
typedef int (__stdcall *LPFNEXCEL12V)(int xlfn, void *pxResult, int count, const void *rgpx[]);
static LPFNEXCEL12V pExcel12v = NULL;

#define EXCEL12(fn, pxRes, count, ...) \
    do { \
        const void *args[] = { __VA_ARGS__ }; \
        pExcel12v((fn), (pxRes), (count), args); \
    } while(0)

// XLOPER12
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

#define xlfRegister         20001
#define xlfGetDocument      20032
#define xlcOnKey            30061
#define xlcWorkbookActivate 30096
#define xlcSelect           30243

// Global state
static BOOL    g_bHasPrev  = FALSE;
static BOOL    g_bToggling = FALSE;
static wchar_t g_prevBook[512] = {0};
static wchar_t g_prevSheet[512] = {0};

// ============================================================
// Excel string: first wchar_t = length, then characters, NO null terminator
// We pre-build these strings in static buffers
// ============================================================

// Helper: build an Excel string in a static buffer, returns XLOPER12
static XLOPER12 ExcelStr(const wchar_t *s, wchar_t *buf, int bufSize)
{
    int len = (int)wcslen(s);
    if (len > bufSize - 2) len = bufSize - 2;
    buf[0] = (wchar_t)len;
    memcpy(buf + 1, s, len * sizeof(wchar_t));
    XLOPER12 x = {{0}, xltypeStr};
    x.val.str = buf;
    return x;
}

static XLOPER12 OpInt(int i)
{
    XLOPER12 x = {{0}, xltypeInt};
    x.val.w = (unsigned short)i;
    return x;
}

// Static buffers for Excel strings (one per use)
static wchar_t g_buf1[256], g_buf2[256], g_buf3[256], g_buf4[256], g_buf5[256];

// ============================================================
// Get current workbook name and sheet name
// ============================================================
static void GetCurrent(wchar_t *book, int bsize, wchar_t *sheet, int ssize)
{
    XLOPER12 xArg, xRes;

    xArg = OpInt(88);
    xRes.xltype = xltypeNil;
    EXCEL12(xlfGetDocument, &xRes, 1, &xArg);
    if (xRes.xltype == xltypeStr && xRes.val.str) {
        int len = xRes.val.str[0];
        if (len > bsize - 1) len = bsize - 1;
        memcpy(book, xRes.val.str + 1, len * sizeof(wchar_t));
        book[len] = 0;
    } else {
        book[0] = 0;
    }

    xArg = OpInt(76);
    xRes.xltype = xltypeNil;
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
// Activate workbook and sheet
// ============================================================
static BOOL Activate(const wchar_t *book, const wchar_t *sheet)
{
    if (book[0] == 0 || sheet[0] == 0) return FALSE;

    XLOPER12 xBook = ExcelStr(book, g_buf4, 256);
    EXCEL12(xlcWorkbookActivate, 0, 1, &xBook);

    XLOPER12 xSheet = ExcelStr(sheet, g_buf5, 256);
    EXCEL12(xlcSelect, 0, 1, &xSheet);

    return TRUE;
}

// ============================================================
// Toggle command
// ============================================================
__declspec(dllexport) void WINAPI ToggleCmd(void)
{
    if (!g_bHasPrev) return;

    wchar_t curBook[512], curSheet[512];
    GetCurrent(curBook, 512, curSheet, 512);

    g_bToggling = TRUE;
    Activate(g_prevBook, g_prevSheet);

    wcscpy(g_prevBook, curBook);
    wcscpy(g_prevSheet, curSheet);

    g_bToggling = FALSE;
}

// ============================================================
// Update history
// ============================================================
__declspec(dllexport) void WINAPI SwitchCmd(void)
{
    wchar_t curBook[512], curSheet[512];
    GetCurrent(curBook, 512, curSheet, 512);

    if (!g_bHasPrev) {
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);
        g_bHasPrev = TRUE;
        return;
    }

    if (wcscmp(curBook, g_prevBook) != 0 || wcscmp(curSheet, g_prevSheet) != 0) {
        wchar_t oldBook[512], oldSheet[512];
        wcscpy(oldBook, g_prevBook);
        wcscpy(oldSheet, g_prevSheet);
        wcscpy(g_prevBook, curBook);
        wcscpy(g_prevSheet, curSheet);

        g_bToggling = TRUE;
        Activate(oldBook, oldSheet);
        g_bToggling = FALSE;
    }
}

// ============================================================
// xlAutoOpen
// ============================================================
__declspec(dllexport) int WINAPI xlAutoOpen(void)
{
    HMODULE hMod = GetModuleHandle(NULL);
    pExcel12v = (LPFNEXCEL12V)GetProcAddress(hMod, "Excel12v");
    if (!pExcel12v) return 0;

    XLOPER12 xRes, xArgs[7];

    xArgs[0] = ExcelStr(L"SwitchCmd", g_buf1, 256);
    xArgs[1] = ExcelStr(L"L", g_buf2, 256);
    xArgs[2] = ExcelStr(L"SwitchCmd", g_buf3, 256);
    xArgs[3] = OpStrStatic(L"", 0);
    xArgs[4] = OpInt(1);
    xArgs[5] = OpStrStatic(L"Previous Sheet", 14);
    xArgs[6] = OpInt(0);

    // Fix xArgs[3] and xArgs[5]
    {
        static wchar_t b3[2] = {0, 0};
        b3[0] = 0;
        XLOPER12 x3 = {{0}, xltypeStr};
        x3.val.str = b3;
        xArgs[3] = x3;
    }
    {
        static wchar_t b5[32] = {14, L'P',L'r',L'e',L'v',L'i',L'o',L'u',L's',L' ',L'S',L'h',L'e',L'e',L't'};
        XLOPER12 x5 = {{0}, xltypeStr};
        x5.val.str = b5;
        xArgs[5] = x5;
    }

    EXCEL12(xlfRegister, &xRes, 7,
            &xArgs[0], &xArgs[1], &xArgs[2], &xArgs[3],
            &xArgs[4], &xArgs[5], &xArgs[6]);

    // Bind Ctrl+Alt+L
    static wchar_t keyBuf[8] = {4, L'^', L'%', L'L', 0};
    static wchar_t cmdBuf[16] = {9, L'S',L'w',L'i',L't',L'c',L'h',L'C',L'm',L'd'};

    XLOPER12 xKey = {{0}, xltypeStr};  xKey.val.str = keyBuf;
    XLOPER12 xCmd = {{0}, xltypeStr};  xCmd.val.str = cmdBuf;
    EXCEL12(xlcOnKey, 0, 2, &xKey, &xCmd);

    // Init
    wchar_t b[512], s[512];
    GetCurrent(b, 512, s, 512);
    wcscpy(g_prevBook, b);
    wcscpy(g_prevSheet, s);
    g_bHasPrev = TRUE;

    return 1;
}

// Need this helper
static XLOPER12 OpStrStatic(const wchar_t *s, int len)
{
    static wchar_t buf[256];
    buf[0] = (wchar_t)len;
    if (len > 0) memcpy(buf + 1, s, len * sizeof(wchar_t));
    XLOPER12 x = {{0}, xltypeStr};
    x.val.str = buf;
    return x;
}

// ============================================================
// xlAutoClose
// ============================================================
__declspec(dllexport) int WINAPI xlAutoClose(void)
{
    if (pExcel12v) {
        static wchar_t k[8] = {4, L'^', L'%', L'L', 0};
        XLOPER12 xKey = {{0}, xltypeStr};
        xKey.val.str = k;
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
