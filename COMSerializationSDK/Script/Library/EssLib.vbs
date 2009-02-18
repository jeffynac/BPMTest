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
'  $Header: //sesca/src/Winprog/COMSerializationSDK/Script/Library/rcs/EssLib.vbs 1.2 2009/02/18 16:08:00 arturoc Exp $

' The collection required by BPWin that specifies the data pattern modifications.
'
Dim m_oDictionary
Set m_oDictionary = CreateObject("Scripting.Dictionary")

' The key associated with the item in the dictionary whose value will be added as an entry into the serialization
' log file by BPWin.
'
Private Const k_sLogTextKey = "LogText"

' Type of object returned by the LockVector() routine for specifying modifications for a 
' range of contiguous bytes in the data pattern.
'
Class Vector

   ' Address in the data pattern that is the beginning of the range to modify.
   '
   Private m_oAddress
   
   ' The array that will hold the new values of a range in the data pattern.
   '
   Private m_oBytes

   ' Prepares the address and array data members for modification by the user of the library.
   ' Important: All calls to Lock() require a subsequent call to Unlock().
   '
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

   ' Commits the pattern modifications specified by the library user to the dictionary object.
   '
   Public Sub Unlock()
      m_oDictionary.Item(m_oAddress) = m_oBytes
      ReDim m_oBytes(0)
   End Sub

   ' Precondition: This vector must be locked.
   ' Sets the element in the array at lIndex to lValue.
   '
   Public Sub SetAtIndex(lIndex, lValue)
      m_oBytes(lIndex) = CByte(lValue)
   End Sub

   ' Precondition: This vector must be locked.
   ' Sets the element in the array at the data pattern address specified in oAddress to lValue.
   '
   Public Sub SetAtAddress(oAddress, lValue)
      SetAtIndex oAddress - m_oAddress, lValue
   End Sub

   ' Precondition: The vector must have been locked at some point.
   ' Returns the address this vector was last configured with.
   '
   Public Function GetAddress()
      GetAddress = m_oAddress
   End Function

   ' Precondition: This vector must be locked.
   ' Helper for setting the array in this vector with the array passed in. 
   '
   Public Sub CopyFromArray(oJscriptOrVbScriptArray)
      Dim lIndex
      lIndex = 0
      For Each oElem in oJscriptArray
         SetAtIndex lIndex, oElem
         lIndex = lIndex + 1
      Next
   End Sub

End Class

' Returns a vector object configured with the address and size vector attributes passed in.
' IMPORTANT: The returned vector object must also be subsequently unlocked (by calling the vector object's Unlock()
' method) after the client code has completed modifying the vector's array.
'
Public Function LockVector(oAddress, lSize)
   Dim oNewVector
   Set oNewVector = new Vector
   oNewVector.Lock oAddress, lSize
   Set LockVector = oNewVector
End Function

' Sets the "NextHandle" dictionary item.  BPWin will write the value of sNextHandle to the serialization data file (if 
' one was specified in its serialization dialog) and will also pass this in the next call of GetSerialData() for the
' sCurrentHandle parameter.
'
Public Sub SetNextHandle(sNextHandle)
   m_oDictionary.Item("NextHandle") = sNextHandle
End Sub

' Modifies lSize bytes in the data pattern beginning at oAddress, with the new values of those bytes specified with
' oArray. 
'
Public Sub SetVector(oAddress, oArray, lSize)
  Dim oVec
  Set oVec = LockVector(oAddress, lSize)
  oVec.CopyFromArray oArray
  oVec.Unlock
End Sub

' Helper for raising an exception.
'
Private Sub RaiseException(sDescription, lNumber)
   Err.Clear
   Err.Raise lNumber, "ESS .wsc file", sDescription
End Sub

' Aborts the job session. BPWin will display an error message that will contain the text in the sDescription parameter.
'
Public Sub ThrowError(sDescription)
   RaiseException &H80040300, sDescription
End Sub

' IMPORTANT: For the job session to continue, this can only be called from the GetSerialData() method.
'
' Throws an exception configured to tell BPWin to fail the device operation instead of aborting the job session.
' The text in sDescription will be added to the serialization log file.
'
Public Sub FailDeviceOperation(sDescription)
   RaiseException &H80040220, sDescription
End Sub

' Set the return value of this function as the return value of the GetSerialData() method.  Returns the dictionary
' object BPWin requires the GetSerialData() method to return.
' Also, this function removes the "LogText" entry from the dictionary object.
'
Public Function GetDictionary()
   If m_oDictionary.Exists(k_sLogTextKey) Then
      m_oDictionary.Remove(k_sLogTextKey)
   End If
   Set GetDictionary = m_oDictionary
End Function

' Equivalent to GetDictionary(), but additionally adds an entry to the dictionary object that BPWin will write to the
' serialization log file.
'
Public Sub GetDictionaryAndAddLogEntry(sLogEntryContent)
   m_oDictionary.Item(k_sLogTextKey) = sLogEntryContent
   Set GetDictionaryAndLog = m_oDictionary
end sub
