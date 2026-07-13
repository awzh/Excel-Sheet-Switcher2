#ifndef XLCALL_H
#define XLCALL_H
#ifdef __cplusplus
extern "C" {
#endif

typedef char XCHAR;
typedef short RW;
typedef short COL;
typedef unsigned char BYTE;
typedef double XLFP12;

#define xltypeNum        0x0001
#define xltypeStr        0x0002
#define xltypeBool       0x0004
#define xltypeRef        0x0008
#define xltypeErr        0x0010
#define xltypeFlow       0x0020
#define xltypeMulti      0x0040
#define xltypeMissing    0x0080
#define xltypeNil        0x0100
#define xltypeSRef       0x0400
#define xltypeInt        0x0800

typedef struct xlref12 {
    WORD rwFirst;
    WORD rwLast;
    BYTE colFirst;
    BYTE colLast;
} XLREF12;

typedef struct xlmref12 {
    WORD count;
    XLREF12 reftbl[1];
} XLMREF12;

typedef struct xloper12 {
    union {
        double num;
        XCHAR *str;
        WORD boolval;
        WORD err;
        short int w;
        struct {
            WORD count;
            XLREF12 reftbl[1];
        } sref;
        XLMREF12 *mref;
        struct {
            struct xloper12 *lparray;
            WORD rows;
            WORD columns;
        } array;
        struct {
            union {
                XCHAR *lpbData;
                HANDLE hdata;
            } h;
            long cbData;
        } bigdata;
    } val;
    WORD xltype;
} XLOPER12;

typedef XLOPER12 *LPXLOPER12;

extern int __stdcall Excel12v(int xlfn, LPXLOPER12 operRes, int count, LPXLOPER12 opers[]);

#define xlFree           10000
#define xlfRegister      20001
#define xlfGetWorkbook   20008
#define xlfGetDocument   20032
#define xlfActiveCell    20051
#define xlSheetId        20059
#define xlGetName        20088
#define xlcOnKey         30061
#define xlcWorkbookActivate 30096

static XLOPER12 g_tempOper;
static XCHAR* g_tempStr = NULL;

static LPXLOPER12 TempStr12Impl(const XCHAR *s) {
    if (g_tempStr) free(g_tempStr);
    int len = (int)s[0];
    g_tempStr = (XCHAR*)malloc((len + 2) * sizeof(XCHAR));
    memcpy(g_tempStr, s, (len + 1) * sizeof(XCHAR));
    g_tempStr[len + 1] = 0;
    g_tempOper.xltype = xltypeStr;
    g_tempOper.val.str = g_tempStr;
    return &g_tempOper;
}
static LPXLOPER12 TempInt12Impl(int i) {
    g_tempOper.xltype = xltypeInt;
    g_tempOper.val.w = (WORD)i;
    return &g_tempOper;
}
static LPXLOPER12 TempMissing12Impl() {
    g_tempOper.xltype = xltypeMissing;
    return &g_tempOper;
}
static LPXLOPER12 TempActiveRef12Impl() {
    g_tempOper.xltype = xltypeRef;
    return &g_tempOper;
}
#define TempStr12(st)      TempStr12Impl(st)
#define TempInt12(i)       TempInt12Impl(i)
#define TempMissing12()    TempMissing12Impl()
#define TempActiveRef12()  TempActiveRef12Impl()

#ifdef __cplusplus
}
#endif
#endif
