// xlcall.h - Excel XLL SDK definitions
#ifndef XLCALL_H
#define XLCALL_H

#ifdef __cplusplus
extern "C" {
#endif

#include <stddef.h>

// Basic types
typedef wchar_t XCHAR;  // Wide character for strings

// XLOPER12 type constants
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

// XLOPER12 structure (simplified for our use)
typedef struct xloper12 {
    union {
        double num;
        wchar_t *str;
        WORD boolval;
        WORD err;
        short int w;
    } val;
    WORD xltype;
} XLOPER12;

// Excel12v function signature
extern int __stdcall Excel12v(int xlfn, void *pxResult, int count, const void *rgpx[]);

// Function indices
#define xlFree            10000
#define xlfCaller         10010
#define xlfRegister       20001
#define xlfGetWorkbook    20008
#define xlfGetDocument    20032
#define xlSheetId         20059
#define xlGetName         20088
#define xlcOnKey          30061
#define xlcWorkbookActivate 30096

// Global missing XLOPER12 for convenience
static XLOPER12 g_xlMissing = { {0}, xltypeMissing };

#ifdef __cplusplus
}
#endif

#endif
