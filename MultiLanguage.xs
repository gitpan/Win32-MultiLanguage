#define VC_EXTRALEAN
#define CINTERFACE
#define COBJMACROS

#include <windows.h>
#include <mlang.h>

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"

MODULE = Win32::MultiLanguage		PACKAGE = Win32::MultiLanguage		

PROTOTYPES: DISABLE

void
DetectInputCodepage(octets, ...)
    SV* octets

  PREINIT:
    DWORD dwFlag = MLDETECTCP_NONE;
    DWORD dwPrefWinCodePage = 0;
    CHAR* pSrcStr;
    INT cSrcSize;
    DetectEncodingInfo* lpEncoding;
    INT nScores = 64;
    STRLEN nOctets;
    HRESULT hr;
    int i;
    AV* av;
    IMultiLanguage2* p;

  PPCODE:

    if (items > 1)
        dwFlag = (DWORD)SvIV(ST(1));

    if (items > 2)
        dwPrefWinCodePage = (DWORD)SvIV(ST(2));

    /**/
    if (CoInitialize(NULL) != S_OK)
    {
        warn("CoInitialize failed\n");
        XSRETURN_EMPTY;
    }

    /**/
    if (CoCreateInstance(&CLSID_CMultiLanguage,
                         NULL,
                         CLSCTX_ALL,
                         &IID_IMultiLanguage2,
                         (VOID**)&p) != S_OK)
    {
        warn("CoCreateInstance failed\n");
        CoUninitialize();
        XSRETURN_EMPTY;
    }

    pSrcStr = SvPV(octets, nOctets);
    cSrcSize = (INT)nOctets;

    New(42, lpEncoding, nScores, DetectEncodingInfo);

    /**/
    hr = IMultiLanguage2_DetectInputCodepage(p,
                                             dwFlag,
                                             dwPrefWinCodePage,
                                             pSrcStr,
                                             &cSrcSize,
                                             lpEncoding,
                                             &nScores);

    if (hr == S_FALSE)
    {
        /* warn("The method cannot determine the code page of the input stream.\n"); */
        Safefree(lpEncoding);
        CoUninitialize();
        XSRETURN_EMPTY;
    }
    else if (hr == E_FAIL || hr != S_OK)
    {
        warn("An error occured while calling DetectInputCodepage\n");
        Safefree(lpEncoding);
        CoUninitialize();
        XSRETURN_EMPTY;
    }
    
    warn("Read %d bytes\n", cSrcSize);

    av = newAV();

    for (i = 0; i < nScores; ++i)
    {
        HV* hv = newHV();

        hv_store(hv, "LangID", 6, newSViv(lpEncoding[i].nLangID), 0);
        hv_store(hv, "CodePage", 8, newSViv(lpEncoding[i].nCodePage), 0);
        hv_store(hv, "DocPercent", 10, newSViv(lpEncoding[i].nDocPercent), 0);
        hv_store(hv, "Confidence", 10, newSViv(lpEncoding[i].nConfidence), 0);

        av_push(av, newRV_noinc((SV*)hv));
    }

    /**/
    Safefree(lpEncoding);
    CoUninitialize();

    XPUSHs(sv_2mortal(newRV_noinc((SV*)av)));

