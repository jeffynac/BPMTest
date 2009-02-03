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
// $Header: //sesca/src/Winprog/COMSerializationSDK/VC++/Examples/SingleSourceFiles/FromDocumentation/rcs/MinimalStaticVectorServer.cpp 1.2 2009/02/03 16:23:00 arturoc Exp $
//                                                                                                                    //
///////// 120 columns, Tab = Insert 3 Spaces ///////////////////////////////////////////////////////////////////////////

#include <ComSerSdk/Application.h>
#include <ComSerSdk/StaticVectorServer.h>

typedef ComSerSdk::CStaticVector<0, 1> CVector;

class CStaticVectorServer : public   
   ComSerSdk::CStaticVectorServerBase<CVector>
{
   void SetVector(const std::string &, std::string &, CVector &)
   {}
};

class CEssApp : public ComSerSdk::CApplicationBase
{
public:
   virtual ComSerSdk::CServerPtr CreateServer()
   {
      return ComSerSdk::CServerPtr(new CStaticVectorServer());
   }
};

CEssApp g_oEssApp;
