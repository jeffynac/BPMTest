//************************************************************************************************
// THIS SOURCE CODE IS PROVIDED AS-IS WITH NO WARRANTY.
//************************************************************************************************
// External Serialization Server (ESS) Host
// ----------------------------------------
// This is the source code for the EssHost.exe executable distributed with BPM Microsystem's BPWin software.
// This source code implements a Microsoft Component Object Model (COM) Out-of-processes object with a CLSID of
// 7A61987B-66D9-41c8-973B-846CC37FC1A5, which is the CLSID required by the BPWin software's COM Serialization feature.
// The COM object implemented by this source code does not adhere to the specification put forth by BPM Microsystems
// for performing COM Serialization; instead, the COM object fully implements the OLE Automation IDispatch interface,
// and forwards all of its method calls to a 3rd-party COM object, which does adhere to the specification:
//    BPWin -> EssHost -> 3rd-party-object
// Since the BPWin software restricts the types of COM Serialization ESS objects to only Out-of-process COM objects,
// components that are instead impemented in such a way as to be usable only as in-process can be activated by the
// EssHost object.  BPWin can then be configured to use the EssHost object and the EssHost will, as stated above,
// in turn use a "hosted" object that actually implements BPWin's COM Serialization-required interface.
// As well as enabling the use of in-process COM objects with BPWin COM Serialization, the EssHost object also
// enables other forms of object activation that are not directly supported by BPWin.
// There are three forms of COM object activation supported by the EssHost, referred to in this description as 
// "GetObject", "ProgId", and "CLSID".  Each form of activation requires a string to be specified by the user as
// extra command-line arguments in the BPWin Software's Serialization dialog's "Path:" field.
// An explanation of what each does with that string, and how to specify that string, follows.
//
// GetObject
//    The string is passed to the ::CoGetObject() Windows API function.  This form of activation results in supporting
//    various forms of COM objects that are able to be activated with "monikers", the most useful of which are
//    "Windows Script Component" (.wsc) files, which are a very simple form of coding COM objects.
//    To specify this string to the EssHost: in the Serialization dialog, use the "-getobj" switch with this syntax:
//
//       -getobj"[Moniker Display Name]"
//
//    Here is an example that shows the entire contents of the "Path:" field where the user wants to specify a COM
//    object that adhere's to the COM Serialization specification, and it is implemented in a .wsc script file:
//
//       EssHost.exe -getobj"script:c:\Ess\Script.wsc"
//
// ProgId
//    The string is interpretted by the EssHost as being a Programmatic Identifier (ProgId).  The ProgId is looked up
//    in the registry to find its corresponding CLSID, and that CLSID is used in a call to the ::CoCreateInstance()
//    Windows API function to create the hosted object.
//    To specify this string to the EssHost: in the Serialization dialog, use the "-progid" switch with this syntax:
//
//       -progid"[ProgID]"
//
//    Here is an example that shows the entire contents of the "Path:" field where the user wants to specify a COM
//    object that adhere's to the COM Serialization specification, and it is registered on the system with a
//    ProgID of "MyCompany.Ess":
//
//       EssHost.exe -progid"MyCompany.Ess"
//
// CLSID
//    The string is interpretted by the EssHost as being a CLSID.  The CLSID is used in a call to the 
//    ::CoCreateInstance() Windows API function to create the hosted object.
//    To specify this string to the EssHost: in the Serialization dialog, use the "-clsid" switch with this syntax:
//
//       -clsid"[CLSID]"
//
//    Here is an example that shows the entire contents of the "Path:" field where the user wants to specify a COM
//    object that adhere's to the COM Serialization specification, and it is registered on the system with a
//    CLSID of "EB32A2A8-1F13-4b31-AFB6-33211AF989E0":
//
//       EssHost.exe -clsid"EB32A2A8-1F13-4b31-AFB6-33211AF989E0"
//    
// To build EssHost.exe, run this command at a Visual Studio 2008 Command Prompt:
// For debug:
//  cl /EHsc /Zi EssHost.cpp
// For release:
//  cl /EHsc EssHost.cpp

#define _ATL_ATTRIBUTES
#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#include <atlbase.h>
#include <atlcom.h>
#include <atlsafe.h>
#include <comutil.h>
#include <comdef.h>
#include <winerror.h>
#include <vector>

using namespace ATL;

[module(exe, uuid = "{CAACA545-BB06-4490-8611-2EA464E75D4F}", name = "ExternalSerialServer")]
class CEssHostModule
{
public:

   int WinMain(int nShowCmd ) throw();

   struct HostSpecInfo
   {
      enum Type
      {
         GETOBJ,
         CLSID,
         PROGID,
         NONE
      };
      Type m_eType;
      CString m_sSpec;

      HostSpecInfo():
         m_eType(NONE)
      {}
   };

   HostSpecInfo GetHostSpecInfo(const CString & sCmdLine);
   HostSpecInfo GetAndDeleteSavedHostSpecInfo();
   
private:
   static const CString k_sCmdLineKeyName;
   static const CString k_sEssHostKeyName;

   bool ConfigureSpecInfo(CString sLabel, HostSpecInfo::Type eType, const CString & sToken, HostSpecInfo & oSpecInfo);
   bool TokenMatches(const CString & sCriterion, const CString & sToken);
};

[
	coclass,
	threading(apartment),
	vi_progid("ExternalSerialServer.Host"),
	progid("ExternalSerialServer.Host.1"),
	version(1.0),
	uuid("7A61987B-66D9-41c8-973B-846CC37FC1A5"),
   registration_script(script="EssHost.rgs")
]
class CEssHost : public IDispatch
{
public:
	DECLARE_PROTECT_FINAL_CONSTRUCT()

   virtual HRESULT FinalConstruct()
   {
      HRESULT hReturn = S_OK;

      const CEssHostModule::HostSpecInfo oSpecInfo = _AtlModule.GetAndDeleteSavedHostSpecInfo();
      switch(oSpecInfo.m_eType)
      {
         case CEssHostModule::HostSpecInfo::CLSID:  ActivateFromClsId    (oSpecInfo.m_sSpec, hReturn); break;
         case CEssHostModule::HostSpecInfo::PROGID: ActivateFromProgId   (oSpecInfo.m_sSpec, hReturn); break;
         case CEssHostModule::HostSpecInfo::GETOBJ: ActivateWithGetObject(oSpecInfo.m_sSpec, hReturn); break;

         default:
         {
            hReturn = E_UNEXPECTED;
            ReportOnError(hReturn,
                          "There was a problem processing this specification:\n"
                          "%s\n"
                          "Usage in the BPWin Serialization Dialog \"Path:\" field:\n"
                          "   EssHost.exe -[getobj | clsid | progid]\"[string]\"\n"
                          "For example:\n"
                          "   To specify a .wsc script file:\n"
                          "      EssHost.exe -getobj\"script:c:\\MyScripts\\Serialization5.wsc\"\n"
                          "   To specify a ProgId:\n"
                          "      EssHost.exe -progid\"Project.ClassName\"\n"
                          "   To specify a CLSID:\n"
                          "      EssHost.exe -clsid\"xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx\"",
                          oSpecInfo.m_sSpec.GetString());
         }
      }

      return hReturn;
   }

   virtual HRESULT STDMETHODCALLTYPE GetTypeInfoCount(UINT *pctinfo)
   {
      return m_spoHosted->GetTypeInfoCount(pctinfo);
   }
        
   virtual HRESULT STDMETHODCALLTYPE GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo **ppTInfo)
   {
     return m_spoHosted->GetTypeInfo(iTInfo, lcid, ppTInfo);
   }
        
   virtual HRESULT STDMETHODCALLTYPE GetIDsOfNames(REFIID riid,
                                                   LPOLESTR *rgszNames,
                                                   UINT cNames,
                                                   LCID lcid,
                                                   DISPID *rgDispId)
   {
      return m_spoHosted->GetIDsOfNames(riid, rgszNames, cNames, lcid, rgDispId);
   }

   virtual HRESULT STDMETHODCALLTYPE Invoke(DISPID dispIdMember,
                                            REFIID riid,
                                            LCID lcid,
                                            WORD wFlags,
                                            DISPPARAMS *pDispParams,
                                            VARIANT *pVarResult,
                                            EXCEPINFO *pExcepInfo,
                                            UINT *puArgErr)
   {
      return m_spoHosted->Invoke(dispIdMember,
                                 riid,
                                 lcid,
                                 wFlags,
                                 pDispParams,
                                 pVarResult,
                                 pExcepInfo,
                                 puArgErr);
   }

private:

   void ReportOnError(HRESULT hResult, LPCTSTR pszFormat, ...)
   {
      if(SUCCEEDED(hResult)) return;

       va_list oArgs;
       va_start(oArgs, pszFormat);

       CString sResult;
       sResult.FormatV(pszFormat, oArgs);
       va_end(oArgs);

       AtlReportError(__uuidof(CEssHost), sResult, GUID_NULL, hResult);
   }

   void ActivateWithGetObject(const CString & sName, HRESULT & hResult)
   {
      hResult = ::CoGetObject(_bstr_t(sName), NULL, IID_IDispatch, reinterpret_cast<void**>(&m_spoHosted));
      ReportOnError(hResult,
                    "CoGetObject(\"%s\") failed with HRESULT 0x%X:\n"
                    "%s",
                    sName.GetString(),
                    hResult,
                    SystemErrorDescription(hResult).GetString());
   }

   void ActivateFromProgId(const CString & sProgId, HRESULT & hResult)
   {
      CLSID oClsid = {0};
      hResult = ::CLSIDFromProgID(CStringW(sProgId), &oClsid);
      ReportOnError(hResult,
                    "CLSIDFromProgID(\"%s\") returned HRESULT 0x%X.",
                    sProgId.GetString(),
                    hResult);
      if(FAILED(hResult)) return;

      LPOLESTR pszClsid = NULL;
      ::StringFromCLSID(oClsid, &pszClsid);

      hResult = m_spoHosted.CoCreateInstance(oClsid);

      ReportOnError(hResult,
                    "The CLSID \"%s\", which corresponds to ProgID \"%s\", could not be used to instantiate"
                    " a COM object. CoCreateInstance(%s) failed with 0x%X\n"
                    "%s",
                    pszClsid,
                    sProgId.GetString(),
                    pszClsid,
                    hResult,
                    SystemErrorDescription(hResult).GetString());

      ::CoTaskMemFree(pszClsid);
   }

   void ActivateFromClsId(const CString & sClsId, HRESULT & hResult)
   {
      CLSID oClsid = {0};
      hResult = ::CLSIDFromString(_bstr_t(sClsId), &oClsid);

      ReportOnError(hResult,
                    "CLSIDFromString(\"%s\") returned HRESULT 0x%X.\n"
                    "%s",
                    sClsId.GetString(),
                    hResult,
                    SystemErrorDescription(hResult).GetString());
      if(FAILED(hResult)) return;

      hResult = m_spoHosted.CoCreateInstance(oClsid);

      ReportOnError(hResult,
                    "CoCreateInstance(CLSID = \"%s\") returned HRESULT 0x%X.\n"
                    "%s",
                    sClsId,
                    hResult,
                    SystemErrorDescription(hResult).GetString());
   }

   CString SystemErrorDescription(HRESULT hResult)
   {
       LPTSTR pszBuffer = NULL;
       const DWORD dwResult = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM   |
                                              FORMAT_MESSAGE_IGNORE_INSERTS |
                                              FORMAT_MESSAGE_ALLOCATE_BUFFER,
                                              NULL,
                                              hResult,
                                              LANG_NEUTRAL,
                                              reinterpret_cast<LPTSTR>(&pszBuffer),
                                              0,
                                              NULL);

       if(dwResult != 0)
       {
          const CString sLoadedMsg = pszBuffer;
          ::LocalFree(pszBuffer);
          return sLoadedMsg.Left(sLoadedMsg.GetLength() - 2);
       }

       return CString();
   }

   CComPtr<IDispatch> m_spoHosted;
};

const CString CEssHostModule::k_sCmdLineKeyName    = "CommandLine";
const CString CEssHostModule::k_sEssHostKeyName = "Software\\ExternalSerialServerHost";

int CEssHostModule::WinMain(int nShowCmd ) throw( )
{
   const CString sCmdLine = ::GetCommandLine();
   if(HostSpecInfo::NONE != GetHostSpecInfo(sCmdLine).m_eType)
   {
      CRegKey oCmdLineKey;
      oCmdLineKey.Create(HKEY_CURRENT_USER, k_sEssHostKeyName);
      oCmdLineKey.SetKeyValue(k_sCmdLineKeyName, sCmdLine);
   }

   AtlSetPerUserRegistration(true);
   
   return ATL::CAtlExeModuleT<CEssHostModule>::WinMain(nShowCmd);
}

bool CEssHostModule::TokenMatches(const CString & sCriterion, const CString & sToken)
{
   return sToken.Left(sCriterion.GetLength()).CompareNoCase(sCriterion) == 0;
}

bool CEssHostModule::ConfigureSpecInfo(CString sLabel, 
                                       HostSpecInfo::Type eType,
                                       const CString & sToken,
                                       HostSpecInfo & oSpecInfo)
{
   if(!TokenMatches(sLabel, sToken)) return false;
   
   oSpecInfo.m_sSpec = sToken.Left(sToken.ReverseFind('\"'));
   oSpecInfo.m_sSpec = oSpecInfo.m_sSpec.Right(oSpecInfo.m_sSpec.GetLength() - sLabel.GetLength());
   oSpecInfo.m_sSpec.Remove('\"');
   oSpecInfo.m_eType = eType;
   return true;
}

CEssHostModule::HostSpecInfo CEssHostModule::GetHostSpecInfo(const CString & sCmdLine)
{
   TCHAR szTokens[] = _T("-/");
   CString sToken = FindOneOf(sCmdLine, szTokens);
   HostSpecInfo oSpecInfo;
	while (!sToken.IsEmpty())
	{
      ConfigureSpecInfo("clsid",  HostSpecInfo::CLSID,  sToken, oSpecInfo);
      ConfigureSpecInfo("getobj", HostSpecInfo::GETOBJ, sToken, oSpecInfo);
      ConfigureSpecInfo("progid", HostSpecInfo::PROGID, sToken, oSpecInfo);
      if(TokenMatches("unregserver", sToken)) return HostSpecInfo();

      sToken = FindOneOf(sToken, szTokens);
   }
   return oSpecInfo;
}

CEssHostModule::HostSpecInfo CEssHostModule::GetAndDeleteSavedHostSpecInfo()
{
   CRegKey oCmdLineKey;
   oCmdLineKey.Create(HKEY_CURRENT_USER, k_sEssHostKeyName + "\\" + k_sCmdLineKeyName);

   CString sCmdLine;
   ULONG ulNchars = _MAX_PATH * 2;
   LONG lResult = oCmdLineKey.QueryValue(NULL,
                                         NULL,
                                         reinterpret_cast<void*>(sCmdLine.GetBuffer(ulNchars + 1)),
                                         &ulNchars);
   sCmdLine.ReleaseBuffer();

   oCmdLineKey.DeleteValue(NULL);

   return GetHostSpecInfo(sCmdLine);
}