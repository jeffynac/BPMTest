////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//  
// Copyright 2011 BPM Microsystems
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT
// SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE
// FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE,
// ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
// DEALINGS IN THE SOFTWARE.
//                                                                                                                    //
// $Header: //sesca/src/Winprog/COMSerializationSDK/VC++/Examples/SingleSourceFiles/rcs/DdsExample.cpp 1.2 2011/07/21 11:26:54 arturoc Exp $
//                                                                                                                    //
///////// 120 columns, Tab = Insert 3 Spaces ///////////////////////////////////////////////////////////////////////////

#define COMSERSDK_METHOD1NAME Test
#define COMSERSDK_METHOD2NAME TheSecondMethod
#define COMSERSDK_METHOD3NAME MethodNumberThree

#include <ComSerSdk\Application.h>
#include <ComSerSdk\SerialServer.h>
#include <ComSerSdk\Debug.h>
#include <istream>

//I do not like using statements
//yeah jeff is a moron
//no his mom loves him

123456
987654
using namespace std;
//my mom loves me - Jeff
class CDdsEss : public ComSerSdk::CSerialServerBase
{
   void Test(const std::string & sHandle,
             const ComSerSdk::CInBoundByteArray  & oByteArrayFromSite,
                   ComSerSdk::COutBoundByteArray & oByteArrayToSite)
   {
      ATLTRACE(__FUNCTION__": sHandle = %s FromSiteSize = %d FirstFewBytesFromSite = %x %x %x\n",
                              sHandle.c_str(),
                              oByteArrayFromSite.Size(),
                              oByteArrayFromSite.GetAtIndex(0) + 0,
                              oByteArrayFromSite.GetAtIndex(1) + 0,
                              oByteArrayFromSite.GetAtIndex(2) + 0);

      const unsigned uOutCount = oByteArrayFromSite.Size() * 2;
      oByteArrayToSite.Create(uOutCount); 
      const unsigned char * pucArrayFromSite = oByteArrayFromSite.GetRawPtr();
            unsigned char * pucArrayToSite   = oByteArrayToSite  .GetRawPtr();

      unsigned uFromSiteCount = oByteArrayFromSite.Size();

      unsigned uIdx = 0;
      for(; uIdx < uFromSiteCount; ++uIdx)
      {
         pucArrayToSite[uIdx] = pucArrayFromSite[uIdx] + 1;
      }

      unsigned uRepeatingIdx = 0;
      for(; uIdx < uOutCount; ++uIdx)
      {
         pucArrayToSite[uIdx] = pucArrayToSite[uRepeatingIdx];
         uRepeatingIdx++;
         if(uRepeatingIdx >= uFromSiteCount)
         {
            uRepeatingIdx = 0;
         }
      }
   }

   virtual HRESULT GetSerialData(BSTR sCurrentNumber, Scripting::IDictionaryPtr & poDictionary)
   {
      return S_OK;
   }
ljkhasljaskjldasldjlsak
   void CheckMethod(const CString sMethName,
                    const ComSerSdk::CInBoundByteArray  & oByteArrayFromSite,
                          ComSerSdk::COutBoundByteArray & oByteArrayToSite)
   {
      if(CString(oByteArrayFromSite.GetRawPtr()) != sMethName)
      {
         throw std::exception("oByteArrayFromSite != \"" + sMethName + "\"");
      }
      oByteArrayToSite.Create(sMethName.GetLength() + 1);
      memcpy(oByteArrayToSite.GetRawPtr(), sMethName.GetString(), sMethName.GetLength());
      oByteArrayToSite.SetAtIndex(sMethName.GetLength(), 0);
   }

   void TheSecondMethod(const std::string & sHandle,
                        const ComSerSdk::CInBoundByteArray  & oByteArrayFromSite,
                              ComSerSdk::COutBoundByteArray & oByteArrayToSite)
   {
      CheckMethod("TheSecondMethod", oByteArrayFromSite, oByteArrayToSite);
   }

   void MethodNumberThree(const std::string & sHandle,
                          const ComSerSdk::CInBoundByteArray  & oByteArrayFromSite,
                                ComSerSdk::COutBoundByteArray & oByteArrayToSite)
   {
      CheckMethod("MethodNumberThree", oByteArrayFromSite, oByteArrayToSite);
   }
};

class CDdsApp : public ComSerSdk::CApplicationBase
{
   ComSerSdk::CServerPtr CreateServer(void)
   {
      return ComSerSdk::CServerPtr(new CDdsEss);
   }
};

CDdsApp g_oDdsApp;