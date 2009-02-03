'
'  Copyright 2009 BPM Microsystems
'  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE, TITLE AND NON-INFRINGEMENT. IN NO EVENT
'  SHALL THE COPYRIGHT HOLDERS OR ANYONE DISTRIBUTING THE SOFTWARE BE LIABLE
'  FOR ANY DAMAGES OR OTHER LIABILITY, WHETHER IN CONTRACT, TORT OR OTHERWISE,
'  ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
'  DEALINGS IN THE SOFTWARE.
'  
'  $Header: //sesca/src/Winprog/COMSerializationSDK/Script/Library/rcs/EssLib.vbs 1.1 2009/02/03 15:00:40 arturoc Exp $

Dim m_oDictionary
Set m_oDictionary = CreateObject("Scripting.Dictionary")

Dim m_oVectors()
ReDim m_oVectors(0)

Class Vector

   Private m_oAddress
   Private m_oBytes

   Public Sub Lock(oAddress, lNumElems)
      m_oAddress = oAddress
      If m_oDictionary.Exists(oAddress) Then
         m_oBytes = m_oDictionary.Item(oAddress)
         m_oDictionary.Remove(oAddress)
      Else
         ReDim m_oBytes(lNumElems - 1)
         Dim lIndex
         lIndex = 0
         For Each oElem in m_oBytes
            SetAtIndex lIndex, &HFF
            lIndex = lIndex + 1
         Next
      End If
   End Sub

   Public Sub Unlock()
      m_oDictionary.Item(m_oAddress) = m_oBytes
      ReDim m_oBytes(0)
   End Sub

   Public Sub SetAtIndex(lIndex, lValue)
      m_oBytes(lIndex) = CByte(lValue)
   End Sub

   Public Sub SetAtAddress(oAddress, lValue)
      SetAtIndex oAddress - m_oAddress, lValue
   End Sub

   Public Function GetAddress()
      GetAddress = m_oAddress
   End Function

   Public Sub CopyFromArray(oJscriptArray)
      Dim lIndex
      lIndex = 0
      For Each oElem in oJscriptArray
         SetAtIndex lIndex, oElem
         lIndex = lIndex + 1
      Next
   End Sub

End Class

Public Function LockVector(oAddress, lSize)
   Dim oNewVector
   Set oNewVector = new Vector
   oNewVector.Lock oAddress, lSize
   Set LockVector = oNewVector
End Function

Public Function GetDictionary()
   Set GetDictionary = m_oDictionary
End Function

Public Sub SetNextHandle(sNextHandle)
   m_oDictionary.Item("NextHandle") = sNextHandle
End Sub

Public Sub SetVector(oAddress, oArray, lSize)
  Dim oVec
  Set oVec = LockVector(oAddress, lSize)
  oVec.CopyFromArray oArray
  oVec.Unlock
End Sub

Public Sub ThrowError(sDescription)
   Err.Clear
   Err.Raise &H80040300, "ESS .wsc file", sDescription
End Sub