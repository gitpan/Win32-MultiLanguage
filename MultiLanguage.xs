/* $Id$ */

/* max number of code pages */
#define MULTILANGUAGE_XS_NSCORES 32

/* max number of code pages */
#define MULTILANGUAGE_XS_DCPS    128

/* :-( */
#define MyCoCreateMlang(p, iid) \
    if (CoCreateInstance(&CLSID_CMultiLanguage, \
                         NULL, \
                         CLSCTX_ALL, \
                         iid, \
                         (VOID**)&p) != S_OK) \
    { \
        warn("CoCreateInstance failed\n"); \
        XSRETURN_EMPTY; \
    }

#define VC_EXTRALEAN
#define CINTERFACE
#define COBJMACROS

#include <windows.h>
#include <mlang.h>

#include "EXTERN.h"
#include "perl.h"
#include "XSUB.h"

SV* wchar2sv(LPCWSTR lpS, UINT nLen)
{
    LPSTR strBuf;
    int nRequired;
    SV* result;

    /* determine size for character buffer first */
    nRequired = WideCharToMultiByte(65001, 0, lpS, nLen, NULL, 0, NULL, NULL);
    
    if (!nRequired)
    {
        warn("Unexpected result from WideCharToMultiByte\n");
        return &PL_sv_undef;
    }

    New(42, strBuf, nRequired, char);
    
    if (!strBuf)
    {
        warn("Insufficient memory\n");
        return &PL_sv_undef;
    }
    
    if (!WideCharToMultiByte(65001, 0, lpS, nLen, strBuf, nRequired, NULL, NULL))
    {
        warn("WideCharToMultiByte failed\n");
        Safefree(strBuf);
        return &PL_sv_undef;
    }
    
    result = newSV(0);
    sv_usepvn(result, strBuf, nRequired);
    SvUTF8_on(result);

    return result;
}

LPWSTR sv2wchar(SV* sv, UINT* len)
{
    LPWSTR lpNew;
	STRLEN svlen;
	LPSTR lpOld;
	int nRequired;
	
	if (!sv)
	  return NULL;
	  
	if (!SvUTF8(sv))
	{
	    /* warn("non-utf8 data in sv2wchar\n"); */
	}
	
	if(!len)
	  return NULL;
	
	*len = 0;
	
	/* upgrade to utf-8 if necessary */
	lpOld = SvPVutf8(sv, svlen);
	
	if (!svlen)
	{
	    New(42, lpNew, 1, WCHAR);
	    lpNew[0] = 0;
	    return lpNew;
	}

    nRequired = MultiByteToWideChar(65001, 0, lpOld, svlen, NULL, 0);
    
    if (!nRequired)
    {
        warn("Unexpected result from MultiByteToWideChar\n");
        return NULL;
    }

    New(42, lpNew, nRequired + 1, WCHAR);

    if (!lpNew)
    {
        warn("Insufficient memory\n");
        return NULL;
    }
    
    /* null-terminate string */
    lpNew[nRequired] = 0;

    if (!MultiByteToWideChar(65001, 0, lpOld, svlen, lpNew, nRequired))
    {
        warn("MultiByteToWideChar failed\n");
        Safefree(lpNew);
        return NULL;
    }
    
    /* set length */
    *len = nRequired;
    
    return lpNew;
}

MODULE = Win32::MultiLanguage       PACKAGE = Win32::MultiLanguage      

PROTOTYPES: DISABLE

BOOT:
    if (CoInitialize(NULL) != S_OK)
    {
        /* todo: check whether this is the best thing to do */
        croak("CoInitialize failed\n");
        XSRETURN_NO;
    }

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

    MyCoCreateMlang(p, &IID_IMultiLanguage2)

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

	/* no longer needed */
    IMultiLanguage2_Release(p);

    if (hr == S_FALSE)
    {
        XSRETURN_EMPTY;
    }
    else if (hr == E_FAIL || hr != S_OK)
    {
        warn("An error occured while calling DetectInputCodepage\n");
        XSRETURN_EMPTY;
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

    XPUSHs(sv_2mortal(newRV_noinc((SV*)av)));

void
GetRfc1766FromLcid(svLocale)
    SV* svLocale

  PREINIT:
    BSTR bstrRfc1766;
    LCID lcidLocale;
    HRESULT hr;
    IMultiLanguage2* p;
    SV* result;

  PPCODE:
    lcidLocale = SvUV(svLocale);

    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    
    hr = IMultiLanguage2_GetRfc1766FromLcid(p, lcidLocale, &bstrRfc1766);
    IMultiLanguage2_Release(p);
    
    if (hr == E_INVALIDARG)
    {
        warn("One or more of the arguments are invalid.\n");
        XSRETURN_EMPTY;
    }
    else if (hr == E_FAIL || hr != S_OK || !bstrRfc1766)
    {
        XSRETURN_EMPTY;
    }
    
    result = wchar2sv(bstrRfc1766, wcslen(bstrRfc1766));
    
    /*
      it is not documented that the caller is responsible for freeing the bstr
      but if this is not done here the application leaks memory, so we free it
    */
    SysFreeString(bstrRfc1766);
    XPUSHs(sv_2mortal(result));

void
GetCodePageInfo(svCodePage, svLangId)
    SV* svCodePage
    SV* svLangId

  PREINIT:
    UINT uiCodePage;
    LANGID LangId;
    MIMECPINFO cpi;
    IMultiLanguage2* p;
    HV* hv;
    HRESULT hr;

  PPCODE:
    uiCodePage = SvUV(svCodePage);
    LangId = (LANGID)SvUV(svLangId);

    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    
    hr = IMultiLanguage2_GetCodePageInfo(p, uiCodePage, LangId, &cpi);

	/* no longer needed */
    IMultiLanguage2_Release(p);

    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    hv = newHV();
    
    hv_store(hv, "Flags",             5, newSVuv(cpi.dwFlags),                                               0);
    hv_store(hv, "CodePage",          8, newSVuv(cpi.uiCodePage),                                            0);
    hv_store(hv, "FamilyCodePage",   14, newSVuv(cpi.uiFamilyCodePage),                                      0);
    hv_store(hv, "Description",      11, wchar2sv(cpi.wszDescription, wcslen(cpi.wszDescription)),           0);
    hv_store(hv, "WebCharset",       10, wchar2sv(cpi.wszWebCharset, wcslen(cpi.wszWebCharset)),             0);
    hv_store(hv, "HeaderCharset",    13, wchar2sv(cpi.wszHeaderCharset, wcslen(cpi.wszHeaderCharset)),       0);
    hv_store(hv, "BodyCharset",      11, wchar2sv(cpi.wszBodyCharset, wcslen(cpi.wszBodyCharset)),           0);
    hv_store(hv, "FixedWidthFont",   14, wchar2sv(cpi.wszFixedWidthFont, wcslen(cpi.wszFixedWidthFont)),     0);
    hv_store(hv, "ProportionalFont", 16, wchar2sv(cpi.wszProportionalFont, wcslen(cpi.wszProportionalFont)), 0);
    hv_store(hv, "GDICharset",       10, newSViv(cpi.bGDICharset),                                           0);

    XPUSHs(sv_2mortal(newRV_noinc((SV*)hv)));

  /* todo: GetRfc1766Info */
  /* todo: GetFamilyCodePage */

void
GetCodePageDescription(svCodePage, svLcid)
    SV* svCodePage
    SV* svLcid

  PREINIT:
    UINT uiCodePage;
    LCID lcid;
    WCHAR lpWideCharStr[MAX_MIMECP_NAME];
    int cchWideChar = MAX_MIMECP_NAME;
    IMultiLanguage2* p;
    HRESULT hr;

  PPCODE:
  
    uiCodePage = SvUV(svCodePage);
    lcid = SvUV(svLcid);

    MyCoCreateMlang(p, &IID_IMultiLanguage2)

    hr = IMultiLanguage2_GetCodePageDescription(p, uiCodePage, lcid, lpWideCharStr, cchWideChar);
    
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }

    XPUSHs(sv_2mortal(wchar2sv(lpWideCharStr, wcslen(lpWideCharStr))));
  
void
GetCharsetInfo(svCharset)
    SV* svCharset
    
  PREINIT:
    HV* hv;
    MIMECSETINFO csi;
    IMultiLanguage2* p;
    BSTR bstrCharset;
    UINT len;
    HRESULT hr;
    
  PPCODE:

    bstrCharset = (BSTR)sv2wchar(svCharset, &len);
    
    if (!bstrCharset)
    {
        warn("conversion to wide string failed\n");
        XSRETURN_EMPTY;
    }
    
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetCharsetInfo(p, bstrCharset, &csi);

    /* no longer needed */
    Safefree(bstrCharset);
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    hv = newHV();
    hv_store(hv, "CodePage",          8, newSVuv(csi.uiCodePage),                          0);
    hv_store(hv, "InternetEncoding", 16, newSVuv(csi.uiInternetEncoding),                  0);
    hv_store(hv, "Charset",           7, wchar2sv(csi.wszCharset, wcslen(csi.wszCharset)), 0);

    XPUSHs(sv_2mortal(newRV_noinc((SV*)hv)));

void
DetectOutboundCodePage(sv, ...)
    SV* sv

  PREINIT:
    DWORD dwFlags = 0;
    LPWSTR lpWideCharStr = NULL;
    UINT cchWideChar = 0;
    UINT* puiPreferredCodePages = NULL;
    UINT nPreferredCodePages = 0;
    UINT puiDetectedCodePages[MULTILANGUAGE_XS_DCPS];
    UINT nDetectedCodePages = MULTILANGUAGE_XS_DCPS;
    LPWSTR wcSpecialChar = NULL;
    HRESULT hr;
    IMultiLanguage3* p;
    UINT i;
    
  PPCODE:
  
    if (items > 1)
        dwFlags = (DWORD)SvIV(ST(1));

    if (items > 2)
    {
        warn("Third parameter not yet implemented\n");
    }
    
    lpWideCharStr = sv2wchar(sv, &cchWideChar);
    
    if (!lpWideCharStr)
    {
        warn("Conversion to wide string failed\n");
        XSRETURN_EMPTY;
    }
    
    MyCoCreateMlang(p, &IID_IMultiLanguage3)

    hr = IMultiLanguage3_DetectOutboundCodePage(p,
                                                dwFlags,
                                                lpWideCharStr,
                                                cchWideChar, 
                                                puiPreferredCodePages, 
                                                nPreferredCodePages,
                                                puiDetectedCodePages,
                                                &nDetectedCodePages,
                                                wcSpecialChar);

    /* no longer needed */
	Safefree(lpWideCharStr);
    IMultiLanguage3_Release(p);

    if (hr != S_OK || !nDetectedCodePages)
    {
        XSRETURN_EMPTY;
    }
    
    for (i = 0; i < nDetectedCodePages; ++i)
    {
        XPUSHs(sv_2mortal(newSViv(puiDetectedCodePages[i])));
    }

SV*
IsConvertible(svSrcEncoding, svDstEncoding)
    SV* svSrcEncoding
    SV* svDstEncoding

  PREINIT:
    DWORD dwSrcEncoding;
    DWORD dwDstEncoding;
    HRESULT hr;
    IMultiLanguage2* p;

  CODE:
    dwSrcEncoding = (DWORD)SvUV(svSrcEncoding);
    dwDstEncoding = (DWORD)SvUV(svDstEncoding);
    
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_IsConvertible(p, dwSrcEncoding, dwDstEncoding);
    IMultiLanguage2_Release(p);

    if (hr == S_FALSE)
        XSRETURN_NO;
        
    if (hr == S_OK)
        XSRETURN_YES;
        
    XSRETURN_UNDEF;

void
GetRfc1766Info(svLocale, svLangId)
    SV* svLocale
    SV* svLangId

  PREINIT:
    RFC1766INFO Rfc1766Info;
    LCID Locale;
    LANGID LangId;
    HRESULT hr;
    HV* hv;
    IMultiLanguage2* p;
    
  PPCODE:
    Locale = (LCID)SvUV(svLocale);
    LangId = (LANGID)SvUV(svLangId);
    
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetRfc1766Info(p, Locale, LangId, &Rfc1766Info);
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    hv = newHV();
    
    hv_store(hv, "Lcid",        4, newSVuv(Rfc1766Info.lcid),                                              0);
    hv_store(hv, "Rfc1766",     7, wchar2sv(Rfc1766Info.wszRfc1766, wcslen(Rfc1766Info.wszRfc1766)),       0);
    hv_store(hv, "LocaleName", 10, wchar2sv(Rfc1766Info.wszLocaleName, wcslen(Rfc1766Info.wszLocaleName)), 0);

    XPUSHs(sv_2mortal(newRV_noinc((SV*)hv)));

void
GetLcidFromRfc1766(svRfc1766)
    SV* svRfc1766

  PREINIT:
    LCID Locale;
    BSTR bstrRfc1766;
    HRESULT hr;
    IMultiLanguage2* p;
    UINT len;
  
  PPCODE:
  
    bstrRfc1766 = sv2wchar(svRfc1766, &len);
    
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetLcidFromRfc1766(p, &Locale, bstrRfc1766);
    IMultiLanguage2_Release(p);
    Safefree(bstrRfc1766);
    
    if (hr != S_FALSE && hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    XPUSHs(sv_2mortal(newSViv( Locale )));
    XPUSHs(sv_2mortal(newSViv( hr == S_FALSE )));

void
GetFamilyCodePage(svCodePage)
    SV* svCodePage
  PREINIT:
    HRESULT hr;
    IMultiLanguage2* p;
	UINT uiCodePage;
	UINT uiFamilyCodePage;

  PPCODE:
    uiCodePage = SvUV(svCodePage);
    
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetFamilyCodePage(p, uiCodePage, &uiFamilyCodePage);
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    XPUSHs(sv_2mortal(newSViv( uiFamilyCodePage )));

void
GetNumberOfCodePageInfo()

  PREINIT:
    HRESULT hr;
    IMultiLanguage2* p;
	UINT uiCodePage = 0;

  PPCODE:
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetNumberOfCodePageInfo(p, &uiCodePage);
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    XPUSHs(sv_2mortal(newSViv( uiCodePage )));

void
GetNumberOfScripts()

  PREINIT:
    HRESULT hr;
    IMultiLanguage2* p;
	UINT nScripts = 0;

  PPCODE:
    MyCoCreateMlang(p, &IID_IMultiLanguage2)
    hr = IMultiLanguage2_GetNumberOfScripts(p, &nScripts);
    IMultiLanguage2_Release(p);
    
    if (hr != S_OK)
    {
        XSRETURN_EMPTY;
    }
    
    XPUSHs(sv_2mortal(newSViv( nScripts )));



void
END()
  CODE:
    CoUninitialize();
    XSRETURN_YES; /* todo: ... */

