Public Module OWInterface

  Public Interface BusMasterInterface

    Function ResetInterface() As Boolean
    Function OW_SearchROM(Optional ByVal FamilyCode As Byte = 0, Optional ByVal Conditional As Boolean = False) As OW_Response
    Function OW_AddressDevice(ByVal Address As String) As OW_Response
    Function OW_MatchROM() As OW_Response
    Function OW_ResetBus() As OW_Response
    Function OW_PowerDownBus() As OW_Response
    Function OW_WriteBlock(ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OW_Response
    Function OW_ReadBit() As OW_Response
    Function OW_WriteBit(ByVal BitToWrite As Byte) As OW_Response
    Function OW_ReadPages(ByVal StartPage As Byte, Optional ByVal PagesToRead As Byte = 1, Optional ByVal Address As String = "") As OW_Response
    Function OW_ReadFileRecords(ByVal StartRecord As Byte, Optional ByVal RecordsToRead As Byte = 1, Optional ByVal Address As String = "") As OW_Response
    Function OW_WriteFileRecord(ByVal RecordNumber As Byte, ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OW_Response
    Function OW_GetLock() As OW_Response
    Function OW_ReleaseLock() As OW_Response
    Sub log(ByVal WhatToLog As String)
    Function getLog() As String
    Sub clearLog()
    Sub incrementLogLevel()
    Sub decrementLogLevel()

  End Interface

  Public Class OW_Response
    Public ResponseTime As Date
    Public Exception_Code As String
    Public Exception_Description As String
    Public Data As ArrayList
  End Class

  Public Function Hex2(ByVal dec As Integer) As String
    Dim retVal As String
    retVal = Hex(dec)
    If Len(retVal) Mod 2 <> 0 Then
      retVal = "0" & retVal
    End If
    Return retVal
  End Function

  Public MustInherit Class EDS_OWInterface
    Implements BusMasterInterface

    Public EnableLogging As Boolean = False
    Public InterfaceID As String        ' Tracks Interface ID
    Public InterfaceType As String      ' Tracks Interface Type
    Public DeviceId As Integer = 0      ' Tracks the Device ID

    Private logData As String
    Private logLevel As Integer = 0

    Public MustOverride Function OW_AddressDevice(ByVal Address As String) As OW_Response Implements BusMasterInterface.OW_AddressDevice
    Public MustOverride Function OW_GetLock() As OW_Response Implements BusMasterInterface.OW_GetLock
    Public MustOverride Function OW_MatchROM() As OW_Response Implements BusMasterInterface.OW_MatchROM
    Public MustOverride Function OW_PowerDownBus() As OW_Response Implements BusMasterInterface.OW_PowerDownBus
    Public MustOverride Function OW_ReadBit() As OW_Response Implements BusMasterInterface.OW_ReadBit
    Public MustOverride Function OW_ReadFileRecords(ByVal StartRecord As Byte, Optional ByVal RecordsToRead As Byte = 1, Optional ByVal Address As String = "") As OW_Response Implements BusMasterInterface.OW_ReadFileRecords
    Public MustOverride Function OW_ReadPages(ByVal StartPage As Byte, Optional ByVal PagesToRead As Byte = 1, Optional ByVal Address As String = "") As OW_Response Implements BusMasterInterface.OW_ReadPages
    Public MustOverride Function OW_ReleaseLock() As OW_Response Implements BusMasterInterface.OW_ReleaseLock
    Public MustOverride Function OW_ResetBus() As OW_Response Implements BusMasterInterface.OW_ResetBus
    Public MustOverride Function OW_SearchROM(Optional ByVal FamilyCode As Byte = 0, Optional ByVal Conditional As Boolean = False) As OW_Response Implements BusMasterInterface.OW_SearchROM
    Public MustOverride Function OW_WriteBit(ByVal BitToWrite As Byte) As OW_Response Implements BusMasterInterface.OW_WriteBit
    Public MustOverride Function OW_WriteBlock(ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OW_Response Implements BusMasterInterface.OW_WriteBlock
    Public MustOverride Function OW_WriteFileRecord(ByVal RecordNumber As Byte, ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OW_Response Implements BusMasterInterface.OW_WriteFileRecord
    Public MustOverride Function ResetInterface() As Boolean Implements BusMasterInterface.ResetInterface

    'Converts a string of characters in hex notation to a byte array
    Public Shared Function Hex2CharArray(ByVal hexString As String) As Char()
      Dim retVal As Char()
      'Allocate storage space for the resulting byte()
      ReDim retVal((Len(hexString) / 2) - 1)
      For t As Integer = 1 To Len(hexString) Step 2
        retVal((t - 1) / 2) = Microsoft.VisualBasic.ChrW(CInt(("&h" & Mid$(hexString, t, 2))))
      Next
      Return retVal
    End Function
    'Converts a byte array into a hex string
    Public Shared Function ByteArray2Hex(ByVal byteArray As Byte()) As String
      Dim retVal As String = ""
      For t As Integer = 1 To byteArray.Length()
        retVal = retVal & Hex2(byteArray(t - 1))
      Next
      Return retVal
    End Function

    'Simplistic logging facility used by EDS to generate
    'product specific usage documentation
    Public Sub log(ByVal whatToLog As String) Implements BusMasterInterface.log
      If EnableLogging Then
        Me.logData = Me.logData & StrDup(Me.logLevel, vbTab) & whatToLog
      End If
    End Sub
    Public Sub clearLog() Implements BusMasterInterface.clearLog
      Me.logData = ""
      Me.logLevel = 0
    End Sub
    Public Sub incrementLogLevel() Implements BusMasterInterface.incrementLogLevel
      Me.logLevel = Me.logLevel + 1
    End Sub
    Public Sub decrementLogLevel() Implements BusMasterInterface.decrementLogLevel
      If Me.logLevel > 0 Then
        Me.logLevel = Me.logLevel - 1
      End If
    End Sub
    Public Function getLog() As String Implements BusMasterInterface.getLog
      Return Me.logData
    End Function
  End Class
End Module
