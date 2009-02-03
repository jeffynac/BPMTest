/**
   Copyright 2009 BPM Microsystems
   THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
   IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
   FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT
   SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE
   FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE,
   ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
   DEALINGS IN THE SOFTWARE.
*/
#include <ComSerSdk/Application.h>
#include <ComSerSdk/StaticVectorServer.h>

typedef ComSerSdk::CStaticVector<0x40,   8>      CSerialNumber;
typedef ComSerSdk::CStaticVector<0xFF0,  128/8>  CHeader;
typedef ComSerSdk::CStaticVector<0x1000, 0x1000> CPayload;

class CStaticVectorServer :
   public ComSerSdk::CStaticVectorServerBase<CSerialNumber,
                                             CHeader,
                                             CPayload>
{
   void SetVector(const std::string &,
                  std::string &,
                  CSerialNumber & oSerNumVec)
   {
      RandomizeBuffer(oSerNumVec.GetRawPtr(), oSerNumVec.Size());
   }

   void SetVector(const std::string &,
                  std::string &,
                  CHeader & oHeaderVec)
   {
      RandomizeBuffer(oHeaderVec.GetRawPtr(), oHeaderVec.Size());
   }

   void SetVector(const std::string &,
                  std::string &,
                  CPayload & oPayloadVec)
   {
      RandomizeBuffer(oPayloadVec.GetRawPtr(), oPayloadVec.Size());
   }

   void RandomizeBuffer(unsigned char * pucBuffer, unsigned uSize)
   {
      for(unsigned i = 0; i < uSize; ++i)
      {
         pucBuffer[i] = (i % 2 == 0) ? 0xFF : (rand() % 0xFF);
      }
   }
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
