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
// $Header: //sesca/src/Winprog/COMSerializationSDK/VC++/Examples/SingleSourceFiles/FromDocumentation/rcs/FunctionalVectorServer.cpp 1.2 2009/02/03 16:22:59 arturoc Exp $
//                                                                                                                    //
///////// 120 columns, Tab = Insert 3 Spaces ///////////////////////////////////////////////////////////////////////////

#include <ComSerSdk/Application.h>
#include <ComSerSdk/VectorServer.h>

class CVectorServer : public ComSerSdk::CVectorServerBase
{
public:
   CVectorServer():
      m_uNum00(0),
      m_uNumAA(0xFFFFFFFF)
   {}

private:

   unsigned m_uNum00;
   unsigned m_uNumAA;

   void SetNum(ComSerSdk::CLockedVector & oVec,
               void * pvData)
   {
      memcpy(oVec.GetRawPtr(), pvData, oVec.Size());
   }

   void SetVectors(const std::string &,
                         std::string &,
                   ComSerSdk::CVectorSetter & oVectorSetter)
   {
      SetNum(oVectorSetter.Lock(0x00, sizeof m_uNum00), &m_uNum00);
      SetNum(oVectorSetter.Lock(0xAA, sizeof m_uNumAA), &m_uNumAA);
      m_uNum00++;
      m_uNumAA--;
   }
};

class CEssApp : public ComSerSdk::CApplicationBase
{
   ComSerSdk::CServerPtr CreateServer()
   {
      return ComSerSdk::CServerPtr(new CVectorServer());
   }
};

CEssApp g_oEssApp;
