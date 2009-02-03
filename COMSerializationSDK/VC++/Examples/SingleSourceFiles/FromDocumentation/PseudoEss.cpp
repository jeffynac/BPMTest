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

#include /*TODO: Specify correct server base class header.*/
#include <ComSerSdk/Application.h>

class /*TODO: Name your server concrete class.*/ :
   public ComSerSdk::/*TODO: Specify server base class.*/
{
   /*TODO: Override server base class virtuals.*/
};

class CEssApp : public ComSerSdk::CApplicationBase
{
   ComSerSdk::CServerPtr CreateServer()
   {
      return ComSerSdk::CServerPtr(/* TODO: 'new' server class */
                                   /* here with constructor    */
                                   /* arguments, if any.     */);
   }                             
};

CEssApp g_oEssApp;
