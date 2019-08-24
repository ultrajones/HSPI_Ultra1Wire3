Imports HSPI_ULTRA1WIRE3.OWInterface.OW_Response

Public Class HA7E
  Inherits EDS_OWInterface

  ' Provides access to the HA7E RS232 to 1-Wire Bus Master interface
  ' manufactured by Embedded Data Systems, LLC.

  Private SerialPort As HSPI_ULTRA1WIRE3.SerialPort
  Private last1WireAddress As String

  Const CMD_ADDRESS_DEVICE As String = "A"
  Const CMD_MATCH_DEVICE As String = "M"
  Const CMD_READ_FILE_RECORDS As String = "L"
  Const CMD_READ_PAGES As String = "G"
  Const CMD_READ_NEXT_PAGE As String = "g"
  Const CMD_RESET_BUS As String = "R"
  Const CMD_SEARCH As String = "S"
  Const CMD_SEARCH_NEXT As String = "s"
  Const CMD_CONDITIONAL_SEARCH As String = "C"
  Const CMD_CONDITIONAL_SEARCH_NEXT As String = "c"
  Const CMD_FAMILY_SEARCH As String = "F"
  Const CMD_FAMILY_SEARCH_NEXT As String = "f"
  Const CMD_WRITE_BIT As String = "B"
  Const CMD_WRITE_BLOCK As String = "W"
  Const CMD_WRITE_FILE_RECORD As String = "U"

  Public Sub New(ByVal DeviceId As Integer, ByVal portName As String)

    SerialPort = New SerialPort(portName.Replace(":", ""), 9600, SerialPort.NOPARITY, 8, SerialPort.ONESTOPBIT)
    SerialPort.flushComIn()

    Me.DeviceId = DeviceId
    Me.InterfaceType = "HA7E"
    Me.InterfaceID = portName

  End Sub

  Public Overrides Function OW_AddressDevice(ByVal Address As String) As OW_Response

    Dim retVal As New OWInterface.OW_Response

    Try
      Dim retData As New ArrayList
      Dim ha7Response As String
      Dim ha7Request As String

      '
      ' Build the request
      '
      ha7Request = CMD_ADDRESS_DEVICE & Address & vbCr

      '
      ' Get data from the HA7E
      '
      WriteMessage("HA7E Addressing Device having ROMId " & Address & ".", MessageType.Debug)

      ha7Response = txHA7(ha7Request)
      Me.last1WireAddress = Address

      '
      ' Address data is returned for this request
      '
      retData.Add(Left(ha7Response, 16))
      retVal.Data = retData
      retVal.Exception_Code = "0"
      retVal.Exception_Description = "None"
      retVal.ResponseTime = Now

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  'Pretends to obtain an exclusive 1-Wire lock.
  Public Overrides Function OW_GetLock() As OWInterface.OW_Response

    Dim retVal As New OWInterface.OW_Response

    Try
      Dim retData As New ArrayList

      retData.Add("1234567890")
      retVal.Data = retData
      retVal.Exception_Code = "0"
      retVal.Exception_Description = "None"
      retVal.ResponseTime = Now

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  Public Overrides Function OW_MatchROM() As OWInterface.OW_Response

    Dim retVal As New OW_Response

    Try

      Dim retData As New ArrayList
      Dim ha7Response As String
      Dim ha7Request As String

      'No required parameters

      '
      ' Build the request
      '
      ha7Request = CMD_ADDRESS_DEVICE & last1WireAddress & vbCr

      '
      ' Get data from the HA7
      '
      WriteMessage("HA7E Matching Device having ROMId " & last1WireAddress & ".", MessageType.Debug)
      ha7Response = txHA7(ha7Request)

      '
      ' Address data is returned for this request
      '
      retVal.Data = retData

      '
      ' Create stats / exception info
      '
      retVal.ResponseTime = Now()

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  Public Overrides Function OW_PowerDownBus() As OWInterface.OW_Response

  End Function

  Public Overrides Function OW_ReadBit() As OWInterface.OW_Response

  End Function

  Public Overrides Function OW_ReadFileRecords(ByVal StartRecord As Byte, Optional ByVal RecordsToRead As Byte = 1, Optional ByVal Address As String = "") As OWInterface.OW_Response

  End Function

  'Reads one or more memory pages from a 1-Wire device
  Public Overrides Function OW_ReadPages(ByVal StartPage As Byte, Optional ByVal PagesToRead As Byte = 1, Optional ByVal Address As String = "") As OWInterface.OW_Response

    Dim retVal As New OWInterface.OW_Response

    Try

      Dim retData As New ArrayList
      Dim ha7Response As String
      Dim ha7InitialRequest As String
      Dim ha7SubsequentRequest As String

      'Process the Optional parameters
      If Address.Length > 0 Then
        'Need to address the device first
        Me.OW_AddressDevice(Address)
      End If

      'Build the request, including any standard parameters
      ha7InitialRequest = CMD_READ_PAGES & "," & Hex2(1) & Hex2(StartPage)
      ha7SubsequentRequest = CMD_READ_NEXT_PAGE

      'Get first page from the HA7
      Me.log("Reading " & CStr(PagesToRead) & " memory pages from Device having ROMId " & Address & ", starting at page " & CStr(StartPage) & "." & vbCrLf)
      Me.log("Reading first page." & vbCrLf)
      ha7Response = txHA7(ha7InitialRequest)
      retData.Add(Left(ha7Response, Len(ha7Response) - 1))

      'Get subsequent pages from the HA7
      For t As Integer = 1 To PagesToRead - 1
        ha7Response = txHA7(ha7SubsequentRequest)
        'Me.log("Reading subsequent page." & vbCrLf)
        retData.Add(Left(ha7Response, Len(ha7Response) - 1))
      Next
      retVal.Data = retData
      retVal.Exception_Code = "0"
      retVal.Exception_Description = "None"
      retVal.ResponseTime = Now

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  '
  ' Pretends to release an exclusive lock on the 1-Wire bus
  '
  Public Overrides Function OW_ReleaseLock() As OWInterface.OW_Response

    Dim retVal As New OWInterface.OW_Response

    Try

      Dim retData As New ArrayList

      retData.Add("1234567890")
      retVal.Data = retData
      retVal.Exception_Code = "0"
      retVal.Exception_Description = "None"
      retVal.ResponseTime = Now

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  'Resets the 1-Wire bus
  Public Overrides Function OW_ResetBus() As OWInterface.OW_Response

    txHA7(CMD_RESET_BUS)

  End Function

  Public Overrides Function OW_SearchROM(Optional ByVal FamilyCode As Byte = 0, _
                                         Optional ByVal Conditional As Boolean = False) As OWInterface.OW_Response

    Dim retVal As New OW_Response

    Try

      Dim ha7Response As String
      Dim ha7Request As String
      Dim family As String = ""
      Dim retData As New ArrayList

      '
      ' Process the optional parameters
      '
      If FamilyCode <> 0 Then
        family = Hex2(FamilyCode)
      End If

      '
      ' Perform the search
      '
      If Not Conditional Then
        If FamilyCode = 0 Then
          '
          ' Performing a regular search
          '
          WriteMessage("HA7E Performing standard device search.", MessageType.Debug)
          ha7Request = CMD_SEARCH ' Performs initial search, returning only 1 address at a time
          Do
            ha7Response = txHA7(ha7Request)
            WriteMessage("HA7E Response [" & ha7Response & "]", MessageType.Debug)
            If ha7Response.Length() > 1 Then
              retData.Add(Left(ha7Response, 16))
            Else
              Exit Do
            End If
            ha7Request = CMD_SEARCH_NEXT ' Performs a subsequent search, returning only 1 address at a time
          Loop
        Else
          '
          ' Performing a regular family search
          '
          WriteMessage("HA7E Performing family search.", MessageType.Debug)
          ha7Request = CMD_FAMILY_SEARCH & family ' Performs initial family search, returns 1 address
          Do
            ha7Response = txHA7(ha7Request)
            If ha7Response.Length() > 1 Then
              retData.Add(Left(ha7Response, 16))
            Else
              Exit Do
            End If
            ha7Request = CMD_FAMILY_SEARCH_NEXT ' Performs a subsequent family search, returning 1 address at a time
          Loop
        End If
      End If
      If Conditional Then
        '
        ' Performing a conditional search
        '
        WriteMessage("HA7E Performing conditional search.", MessageType.Debug)
        ha7Request = CMD_CONDITIONAL_SEARCH ' Performs initial conditional search, returning only 1 address
        Do
          ha7Response = txHA7(ha7Request)
          If ha7Response.Length() > 1 Then
            If FamilyCode > 0 Then
              '
              ' User requested a Conditional Family Search, so now throw out the ROM-ID's that don't match the requested family code
              '
              If Right(ha7Response, 2) = family Then
                '
                ' Family code matches
                '
                retData.Add(Left(ha7Response, 16))
              End If
            Else
              '
              ' No family code restriction requested
              '
              retData.Add(Left(ha7Response, 16))
            End If
          Else
            Exit Do
          End If
          ha7Request = CMD_CONDITIONAL_SEARCH_NEXT 'Performs a subsequent conditional search, returning 1 address at a time
        Loop
      End If

      retVal.Data = retData

      '
      ' Set stats / exception info
      '
      retVal.ResponseTime = Now()

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  Public Overrides Function OW_WriteBit(ByVal BitToWrite As Byte) As OWInterface.OW_Response

  End Function

  Public Overrides Function OW_WriteBlock(ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OWInterface.OW_Response

    Dim retVal As New OWInterface.OW_Response

    Try

      Dim retData As New ArrayList

      Dim ha7Request As String
      Dim ha7Response As String

      'Process the Optional parameters
      If Address.Length > 0 Then
        'Need to address the device first
        Me.OW_AddressDevice(Address)
      End If

      'Build the request, including any standard parameters
      ha7Request = CMD_WRITE_BLOCK & Hex2(Len(WhatToWrite) / 2) & WhatToWrite & vbCr

      'Get data from the HA7
      WriteMessage("HA7E Writing block of data: " & WhatToWrite & ".", MessageType.Debug)

      ha7Response = txHA7(ha7Request)

      retData.Add(Left(ha7Response, Len(ha7Response) - 1))

      'Set stats / exception info
      retVal.Data = retData
      retVal.Exception_Code = "0"
      retVal.Exception_Description = "None"
      retVal.ResponseTime = Now

    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  Public Overrides Function OW_WriteFileRecord(ByVal RecordNumber As Byte, ByVal WhatToWrite As String, Optional ByVal Address As String = "") As OWInterface.OW_Response

  End Function

  Public Overrides Function ResetInterface() As Boolean

  End Function

  'Transmits data to the HA7, and returns the response
  Private Function txHA7(ByVal whatToSend As String) As String

    Dim retVal As String = ""

    Try
      Dim rx As String = ""

      WriteMessage("HA7E TX [" & whatToSend & "].", MessageType.Debug)
      SerialPort.tx(whatToSend)
      Do While InStr(retVal, Chr(13)) = 0
        rx = SerialPort.rx()
        If rx.Length() = 0 Then
          'Timed out while reading from the port
          Exit Do
        Else
          retVal = retVal & rx
        End If
      Loop
      WriteMessage("HA7E RX [" & retVal & "].", MessageType.Debug)
    Catch pEx As Exception

    End Try

    '
    ' Return results
    '
    Return retVal

  End Function

  Protected Overrides Sub Finalize()
    MyBase.Finalize()
  End Sub
End Class

