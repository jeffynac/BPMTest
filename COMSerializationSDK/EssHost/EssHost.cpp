////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// Copyright 2009 BPM Microsystems
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT
// SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE
// FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE,
// ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
//                                                                                                                    //
// $Header: //sesca/src/Winprog/COMSerializationSDK/EssHost/rcs/EssHost.cpp 1.3 2009/02/05 13:54:02 arturoc Exp $
//                                                                                                                    //
///////// 120 columns, Tab = Insert 3 Spaces ///////////////////////////////////////////////////////////////////////////

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
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
#include <sstream>

using namespace ATL;

[module(exe, uuid = "{CAACA545-BB06-4490-8611-2EA464E75D4F}", name = "ExternalSerialServer")]
class CEssHostModule
{
public:

   CEssHostModule();

   /**
      overrides CAtlModuleT's virtual function to customize registration
   */
   HRESULT RegisterServer(BOOL bRegTypeLib = FALSE, const CLSID* pCLSID = NULL) throw( );

   /**
      overrides CAtlModuleT's virtual function to customize unregistration
   */
   HRESULT UnregisterServer(BOOL bUnRegTypeLib, const CLSID* pCLSID = NULL) throw( );

private:

   /**
      Implementation of the RegisterServer() and UnregisterServer() overrides; in other words, all those overrides
      are defined to do here is forward to this member function.  They specify pointers to member functions
      that we bind here to appropriate objects - it's done this way since both registration and unregistration are
      very similar, with the only difference between them being the names of the member functions that need to be
      called.
      Therefore we have this member function, which defines that code only once, and is parameterized on the only
      difference:  the member functions to execute.
      Note that upon error detection, a dialog box is displayed to notify the user as to the nature of the error.
      This is done since registration occurs without a way of sending BPWin detailed error information since, at this
      point, BPWin is merely executing the .exe with /regserver or /unregserver on the command line.

      @param pfnRegFunc Pointer to either IRegistrar::StringRegister or IRegistrar::StringUnregister, depending
                        on the type of registry operation to perform.
      @param pfnModFunc Pointer to either CAtlModuleT::RegisterServer or CAtlModuleT::UnregisterServer, depending
                        on the type of registry operation to perform.
      @param bRegTypeLib Whether to also register the type library - this value is merely forwarded from the argument
                         provided by the framework.
      @param pCLSID Just like the bRegTypeLib, this is merely forwarded from the value provided by the framework.
      @return The value returned when pfnModFunc was invoked or, if an exception occurred, the result of handling
              the exception.  In either case, this return value is what is returned to the framework.
   */
   template<class TRegistrarFunc, class TModuleFunc>
   HRESULT RegUnReg(TRegistrarFunc pfnRegFunc, TModuleFunc pfnModFunc, BOOL bRegTypeLib, const CLSID* pCLSID) throw( );
         
   /**
      Creates and returns an instance of a Registrar object (see "Creating Registrar Scripts" in MSDN).
      Also, initializes the registrar object with the registry script substitutions we want to perform based 
      on runtime data (path to the current module and whether to add reg info to HKLM or HKCU).
   */
   CComPtr<IRegistrar> PrepareRegistrar();

   /**
      Determines the current path to this executable; this is for adding it to the registry information
      for the LocalServer32 value.
   */
   std::wstring GetModulePath();

   std::string Narrow(std::wstring const& sSource);

   /**
      Adds a replacement key/value pair to the registrar and detects errors with that operation
      and throws exceptions based on those errors.
   */
   void AddReplacement(CComPtr<IRegistrar> & spoRegistrar,
                       const std::wstring & sKey,
                       const std::wstring & sValue);

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
   CString SystemErrorDescription(DWORD dwError);
   
private:
   static const CString k_sCmdLineKeyName;
   static const CString k_sEssHostKeyName;
   const _bstr_t m_sRegScript;

   bool ConfigureSpecInfo(CString sLabel, HostSpecInfo::Type eType, const CString & sToken, HostSpecInfo & oSpecInfo);
   bool TokenMatches(const CString & sCriterion, const CString & sToken);

   /**
      Class for using RAII to free a security identifier allocated in the IsUserAdmin() function, below.
   */
   struct CSidFreer
   {
      /**
         The security identifier to free on destruction.
      */
      PSID m_pSidToFree;

      /**
         On construction, specify the security identifier to free on destruction.
      */
      explicit CSidFreer(PSID pSidToFree):
         m_pSidToFree(pSidToFree)
      {}

      /**
         Frees the security identifier.
      */
      ~CSidFreer()
      {
         ::FreeSid(m_pSidToFree);
      }
   };

   bool IsUserAdmin();
};

[
	coclass,
	threading(apartment),
	vi_progid("ExternalSerialServer.Host"),
	progid("ExternalSerialServer.Host.1"),
	version(1.0),
	uuid("7A61987B-66D9-41c8-973B-846CC37FC1A5"),
   registration_script(script="none")
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
                    _AtlModule.SystemErrorDescription(hResult).GetString());
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
                    _AtlModule.SystemErrorDescription(hResult).GetString());

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
                    _AtlModule.SystemErrorDescription(hResult).GetString());
      if(FAILED(hResult)) return;

      hResult = m_spoHosted.CoCreateInstance(oClsid);

      ReportOnError(hResult,
                    "CoCreateInstance(CLSID = \"%s\") returned HRESULT 0x%X.\n"
                    "%s",
                    sClsId,
                    hResult,
                    _AtlModule.SystemErrorDescription(hResult).GetString());
   }

   CComPtr<IDispatch> m_spoHosted;
};

const CString CEssHostModule::k_sCmdLineKeyName    = "CommandLine";
const CString CEssHostModule::k_sEssHostKeyName = "Software\\ExternalSerialServerHost";

int CEssHostModule::WinMain(int nShowCmd ) throw( )
{
   try
   {
      const CString sCmdLine = ::GetCommandLine();
      if(HostSpecInfo::NONE != GetHostSpecInfo(sCmdLine).m_eType)
      {
         CRegKey oCmdLineKey;
         oCmdLineKey.Create(HKEY_CURRENT_USER, k_sEssHostKeyName);
         oCmdLineKey.SetKeyValue(k_sCmdLineKeyName, sCmdLine);
      }

      return ATL::CAtlExeModuleT<CEssHostModule>::WinMain(nShowCmd);
   }
   catch(std::exception & oException)
   {
      ::MessageBox(NULL, oException.what(), "EssHost Error", MB_OK);
   }

   return -1;
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

bool CEssHostModule::IsUserAdmin()
{
   SID_IDENTIFIER_AUTHORITY oNtAuthority = SECURITY_NT_AUTHORITY;
   PSID pAdministratorsGroup = NULL;
   if(0 == ::AllocateAndInitializeSid(&oNtAuthority,
                                      2,
                                      SECURITY_BUILTIN_DOMAIN_RID,
                                      DOMAIN_ALIAS_RID_ADMINS,
                                      0, 0, 0, 0, 0, 0,
                                      &pAdministratorsGroup))
   {
      const DWORD dwLastError = ::GetLastError();

      std::ostringstream oOutStrm;
      oOutStrm << "::AllocateAndInitializeSid() failed, could not determine if this EssHost process is a member of the"
                  " Administrators local group.  " << SystemErrorDescription(dwLastError).GetString();
                  
      throw std::runtime_error(oOutStrm.str());
   }

   CSidFreer oSidFreer(pAdministratorsGroup);

   BOOL bIsMember = FALSE;
   if(0 == ::CheckTokenMembership(NULL, pAdministratorsGroup, &bIsMember))
   {
      const DWORD dwLastError = ::GetLastError();

      std::ostringstream oOutStrm;
      oOutStrm << "::CheckTokenMembership() failed, could not determine if this EssHost process is a member of the"
                  " Administrators local group.  " << SystemErrorDescription(dwLastError).GetString();
                  
      throw std::runtime_error(oOutStrm.str());
   }

   return bIsMember != FALSE;
}

CString CEssHostModule::SystemErrorDescription(DWORD dwError)
{
    LPTSTR pszBuffer = NULL;
    const DWORD dwResult = ::FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM   |
                                           FORMAT_MESSAGE_IGNORE_INSERTS |
                                           FORMAT_MESSAGE_ALLOCATE_BUFFER,
                                           NULL,
                                           dwError,
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

CEssHostModule::CEssHostModule():
   m_sRegScript(
      "%ROOT%\n"
      "{\n"
      "   NoRemove Software\n"
      "   {\n"
      "      NoRemove Classes\n"
      "      {\n"         
      "         ExternalSerialServer.Host.1 = s 'CEssHost Object'\n"
      "         {\n"
      "            CLSID = s '{7A61987B-66D9-41c8-973B-846CC37FC1A5}'\n"
      "         }\n"
      "         ExternalSerialServer.Host = s 'CEssHost Object'\n"
      "         {\n"
      "            CLSID = s '{7A61987B-66D9-41c8-973B-846CC37FC1A5}'\n"
      "            CurVer = s 'ExternalSerialServer.Host.1'\n"
      "         }\n"
      "         NoRemove CLSID\n"
      "         {\n"
      "           ForceRemove {7A61987B-66D9-41c8-973B-846CC37FC1A5} = s 'CEssHost Object'\n"
      "            {\n"
      "               ProgID = s 'ExternalSerialServer.Host.1'\n"
      "               VersionIndependentProgID = s 'ExternalSerialServer.Host'\n"
      "               ForceRemove 'Programmable'\n"
      "               LocalServer32 = s '%MODULE%'\n"
      "            }\n"
      "         }\n"
      "      }\n"
      "   }\n"
      "}")
{}

HRESULT CEssHostModule::RegisterServer(BOOL bRegTypeLib /*= FALSE*/, const CLSID* pCLSID/* = NULL*/) throw( )
{
   return RegUnReg(&IRegistrar::StringRegister, &CAtlModuleT::RegisterServer, bRegTypeLib, pCLSID);
}

HRESULT CEssHostModule::UnregisterServer(BOOL bUnRegTypeLib, const CLSID* pCLSID /*= NULL*/) throw( )
{
   return RegUnReg(&IRegistrar::StringUnregister, &CAtlModuleT::UnregisterServer, bUnRegTypeLib, pCLSID);
}

template<class TRegistrarFunc, class TModuleFunc>
HRESULT CEssHostModule::RegUnReg(TRegistrarFunc pfnRegFunc,
                                 TModuleFunc pfnModFunc,
                                 BOOL bRegTypeLib,
                                 const CLSID* pCLSID) throw( )
{
   try
   {
      CComPtr<IRegistrar> spoRegistrar = PrepareRegistrar();
      const HRESULT hResult = (spoRegistrar->*pfnRegFunc)(m_sRegScript);
      
      if(FAILED(hResult))
      {
         std::ostringstream oOutStream;
         oOutStream << "Registrar script method call failed with HRESULT " <<
            std::hex << hResult << "h: " << SystemErrorDescription(hResult);
         ::MessageBox(NULL, oOutStream.str().c_str(), "Registration error", MB_OK);
      }

      // Must cast to bind the member function pointer to the base class, not CEssHostModule.
      //
      return (static_cast<CAtlModuleT*>(this)->*pfnModFunc)(bRegTypeLib, pCLSID);
   }
   catch(std::exception & oException)
   {
      ::MessageBox(NULL, oException.what(), "Error", MB_OK);
   }
   return E_FAIL;
}

CComPtr<IRegistrar> CEssHostModule::PrepareRegistrar()
{
	CComPtr<IRegistrar> spoRegistrar;

	HRESULT hResult = CoCreateInstance(CLSID_Registrar,   
									           NULL,
									           CLSCTX_INPROC_SERVER,
									           IID_IRegistrar,
									           reinterpret_cast<void**>(&spoRegistrar));

	if(FAILED(hResult))
	{
      std::ostringstream oOutStream;
      oOutStream << "Could not create instance of IRegistrar object to perform registration.\n"
                    "CoCreateInstance() returned " << std::hex << hResult << "h "
                 << SystemErrorDescription(hResult);
      throw std::runtime_error(oOutStream.str().c_str());
	}

   AddReplacement(spoRegistrar, L"MODULE", GetModulePath());

   // From "Application Compatibility: UAC: COM Per-User Configuration" in MSDN
   //    http://msdn.microsoft.com/en-us/library/bb756926.aspx
   //
   //       "Beginning with Windows Vista® and Windows Server® 2008, if the integrity level of a process is higher
   //        than Medium, the COM runtime ignores per-user COM configuration and accesses only per-machine COM
   //        configuration. This action reduces the surface area for elevation of privilege attacks, preventing a
   //        process with standard user privileges from configuring a COM object with arbitrary code and having this
   //        code called from an elevated process."
   // 
   // Instead of bothering to additionally detect whether we're running Vista or later, we will only detect
   // administrator, since we know that this kind of registration will also work correctly under all supported
   // operating systems.
   //
   AddReplacement(spoRegistrar, L"ROOT", IsUserAdmin() ? L"HKLM" : L"HKCU");

   return spoRegistrar;
}

std::wstring CEssHostModule::GetModulePath()
{
   std::vector<wchar_t> oModule(_MAX_PATH + 1);
   if(0 == ::GetModuleFileNameW(_AtlBaseModule.m_hInst, &oModule.at(0), _MAX_PATH))
	{
      const DWORD dwLastError = ::GetLastError();
      std::ostringstream oOutStream;
      oOutStream << "Could not get path to executable to use in registration data.  ::GetModuleFileNameW() set last "
                    "error to " << dwLastError << " " << SystemErrorDescription(dwLastError);
      throw std::runtime_error(oOutStream.str().c_str());
	}

   return &oModule.at(0);
}

std::string CEssHostModule::Narrow(std::wstring const& sSource)
{
   typedef std::ctype<std::wstring::value_type> CType;
   std::vector<std::string::value_type> oResult(sSource.size() + 1);
   CType const& oCtype = std::use_facet<CType>(std::locale());
   oCtype.narrow(sSource.data(), sSource.data() + sSource.size(), ' ', &oResult.at(0));
   return &oResult.at(0);
}

/**
   Adds a replacement key/value pair to the registrar and detects errors with that operation
   and throws exceptions based on those errors.
*/
void CEssHostModule::AddReplacement(CComPtr<IRegistrar> & spoRegistrar,
                    const std::wstring & sKey,
                    const std::wstring & sValue)
{
   const HRESULT hResult = spoRegistrar->AddReplacement(_bstr_t(sKey.c_str()), _bstr_t(sValue.c_str()));

   if(FAILED(hResult))
   {
      std::wostringstream oOutStrm;
      oOutStrm << "Could not add replacement with key \"" << sKey << "\" and value \""
               << sValue << "\" to registrar to perform registration. IRegistrar.AddReplacement() returned " <<
               std::hex << hResult << "h " << SystemErrorDescription(hResult);

      throw std::runtime_error(Narrow(oOutStrm.str()).c_str());
   }
}