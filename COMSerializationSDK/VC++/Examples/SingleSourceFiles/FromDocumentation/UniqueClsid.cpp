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
// $Header: //sesca/src/Winprog/COMSerializationSDK/VC++/Examples/SingleSourceFiles/FromDocumentation/rcs/UniqueClsid.cpp 1.1 2009/08/14 08:48:06 arturoc Exp $
//                                                                                                                    //
///////// 120 columns, Tab = Insert 3 Spaces ///////////////////////////////////////////////////////////////////////////

#define COMSERSDK_ESS_NAME M29W160ETSerializer

#include <ComSerSdk/Application.h>
#include <ComSerSdk/VectorServer.h>

class CMinimalVectorServer :
   public ComSerSdk::CVectorServerBase
{ghnfghfghfghfggfhf
                         std::string &,
                   ComSerSdk::CVectorSetter & oVectorSetter)
   {
   }
};

class CMyEssApp : public ComSerSdk::CApplicationBase
{
public:
   ComSerSdk::CServerPtr CreateServer()
   {
      return ComSerSdk::CServerPtr(new CMinimalVectorServer());
   }
};

CMyEssApp g_oMyEssApp;
