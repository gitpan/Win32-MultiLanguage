/* $Id$ */

/* max number of code pages */
#define MULTILANGUAGE_XS_NSCORES 32

/* :-( */
#define MY_CO_CREATE_AND_INITIALIZE(p) \
    if (CoInitialize(NULL) != S_OK) \
    { \
        warn("CoInitialize failed\n"); \
        XSRETURN_EMPTY; \
    } \
    if (CoCreateInstance(&CLSID_CMultiLanguage, \
                         NULL, \
                         CLSCTX_ALL, \
                         &IID_IMultiLanguage2, \
                         (VOID**)&p) != S_OK) \
    { \
        warn("CoCreateInstance failed\n"); \
        CoUninitialize(); \
        XSRETURN_EMPTY; \
    }

#define MY_CLEANUP_AND_RETURN(p) \
        IMultiLanguage2_Release(p); \
        CoUninitialize(); \
        XSRETURN_EMPTY


#define VC_EXTRALEAN
#define CINTERFACE
#define COBJMACROS

#include <windows.h>
#include <mlang.h>

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"

MODULE = Win32::MultiLanguage       PACKAGE = Win32::MultiLanguage      

PROTOTYPES: DISABLE

void
DetectInputCodepage(octets, ...)
    SV* octets

  PREINIT:
    DWORD dwFlag = MLDETECTCP_NONE;
    DWORD dwPrefWinCodePage = 0;
    CHAR* pSrcStr;
    INT cSrcSize;
    DetectEncodingInfo lpEncoding[MULTILANGUAGE_XS_NSCORES];
    INT nScores = MULTILANGUAGE_XS_NSCORES;
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

    MY_CO_CREATE_AND_INITIALIZE(p)

    pSrcStr = SvPV(octets, nOctets);
    cSrcSize = (INT)nOctets;

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
        MY_CLEANUP_AND_RETURN(p);
    }
    else if (hr == E_FAIL || hr != S_OK)
    {
        warn("An error occured while calling DetectInputCodepage\n");
        MY_CLEANUP_AND_RETURN(p);
    }
    
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
    IMultiLanguage2_Release(p);
    CoUninitialize();
    XPUSHs(sv_2mortal(newRV_noinc((SV*)av)));

void
GetRfc1766FromLcid(svLocale)
    SV* svLocale

  PREINIT:
    BSTR bstrRfc1766;
    LCID lcidLocale;
    HRESULT hr;
    LPSTR strBuf;
    int nRequired;
    int nBstrLen;
    SV* result;
    IMultiLanguage2* p;

  PPCODE:
    lcidLocale = SvUV(svLocale);

    MY_CO_CREATE_AND_INITIALIZE(p)
    
    hr = IMultiLanguage2_GetRfc1766FromLcid(p, lcidLocale, &bstrRfc1766);
    
    if (hr == E_INVALIDARG)
    {
        warn("One or more of the arguments are invalid.\n");
        MY_CLEANUP_AND_RETURN(p);
    }
    else if (hr == E_FAIL || hr != S_OK || !bstrRfc1766)
    {
        MY_CLEANUP_AND_RETURN(p);
    }
    
    nBstrLen = wcslen(bstrRfc1766);
    
    /* determine size for character buffer first */
    nRequired = WideCharToMultiByte(65001, 0, bstrRfc1766, nBstrLen, NULL, 0, NULL, NULL);
    
    if (!nRequired)
    {
        warn("Unexpected result from WideCharToMultiByte\n");
        SysFreeString(bstrRfc1766);
        MY_CLEANUP_AND_RETURN(p);
    }

    New(42, strBuf, nRequired, char);
    
    if (!strBuf)
    {
        warn("Insufficient memory\n");
        SysFreeString(bstrRfc1766);
        MY_CLEANUP_AND_RETURN(p);
    }
    
    if (!WideCharToMultiByte(65001, 0, bstrRfc1766, nBstrLen, strBuf, nRequired, NULL, NULL))
    {
        warn("WideCharToMultiByte failed\n");
        SysFreeString(bstrRfc1766);
        Safefree(strBuf);
        MY_CLEANUP_AND_RETURN(p);
    }
    
    /* todo: creates a copy, should reuse string */
    result = newSVpvn(strBuf, nRequired);
    SvUTF8_on(result);
    
    /*
      it is not documented that the caller is responsible for freeing the bstr
      but if this is not done here the application leaks memory, so we free it
    */
    SysFreeString(bstrRfc1766);
    
    Safefree(strBuf);
    IMultiLanguage2_Release(p);
    CoUninitialize();
    
    XPUSHs(sv_2mortal(result));

