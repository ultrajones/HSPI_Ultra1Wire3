Imports System.Text.RegularExpressions
Imports System.Globalization

Public Class HA7NetServer

  Public HostName As String = ""        ' Hostname of the HA7Net represented by this instance
  Public HTTPPort As Integer = 80       ' Port number of the http server on the HA7Net represented by this instance
  Public HTTPSPort As Integer = 443     ' Port number of the https server on the HA7Net represented by this instance
  Public UseSSL As Boolean = False      ' Flag to indicate if all http communications with this HA7Net should be made via https
  Public InterfaceID As String = ""     ' Tracks Interface ID
  Public InterfaceType As String = ""   ' Tracks Interface Type
  Public DeviceId As Integer = 0

  Const URL_SEARCH As String = "/1Wire/Search.html"
  Const URL_READ_TEMPERATURE As String = "/1Wire/ReadTemperature.html"
  Const URL_READ_DS18B20 As String = "/1Wire/ReadDS18B20.html"
  Const URL_READ_HUMIDITY As String = "/1Wire/ReadHumidity.html"
  Const URL_READ_COUNTER As String = "/1Wire/ReadCounter.html"
  Const URL_READ_SWITCH As String = "/1Wire/SwitchControl.html"

  Private LockID As String

  ''' <summary>
  ''' Contains data that describes a single HA7Net found by using the Multicast discovery mechanism.
  ''' </summary>
  Public Class DiscoveryResponse
    Public Signature As String = ""           ' HA for standard HA7Net models
    Public Command As Int32                   ' 8001h for standard HA7Net models
    Public Port As Int32                      ' Port number the http server can be reached on
    Public SSLPort As Int32                   ' Port number the https server can be reached on
    Public SerialNumber As String             ' MAC address of the HA7Net
    Public DeviceName As String               ' User configurable device name returned by the HA7Net
    Public IPAddress As System.Net.IPAddress  ' IP Address of the HA7Net
  End Class

  Public Class OW_Response
    Public ResponseTime As Date
    Public Exception_Code As String
    Public Exception_Description As String
    Public Data As Hashtable
  End Class

  ''' <summary>
  ''' Parameters used to build the URL of requests made to the OWServer
  ''' </summary>
  Private Class HttpParameter

    Public Name As String = ""
    Public Value As String = ""

    ''' <summary>
    ''' Default constructor for HttpParameter class
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="Value"></param>
    Public Sub New(ByVal Name As String, ByVal Value As String)
      Me.Name = Name
      Me.Value = Value
    End Sub

  End Class

  ''' <summary>
  ''' Default HA7Net constructor
  ''' </summary>
  ''' <param name="DeviceId"></param>
  ''' <param name="HostName">String representation of the hostname or IP address of the HA7Net represented by this instance</param>
  ''' <param name="InterfaceId"></param>
  Public Sub New(ByVal DeviceId As Integer, ByVal HostName As String, InterfaceId As String)

    Me.HostName = HostName
    Me.InterfaceType = "HA7Net"
    Me.InterfaceID = InterfaceId
    Me.DeviceId = DeviceId

  End Sub

  ''' <summary>
  ''' Hex to String
  ''' </summary>
  ''' <param name="dec"></param>
  ''' <returns></returns>
  Public Function Hex2(ByVal dec As Integer) As String
    Dim retVal As String
    retVal = Hex(dec)
    If Len(retVal) Mod 2 <> 0 Then
      retVal = "0" & retVal
    End If
    Return retVal
  End Function

  ' Uses a multicast mechanism to discover any HA7Nets within
  ' reach of a Multicast packet. This mechanism works by
  ' transmitting a particularly formatted packet to the multicast
  ' group 224.1.2.3 on port 4567. Any HA7Nets that hear the
  ' packet will respond by transmitting a directed UDP packet
  ' back to the client that sent the multicast packet. The packet
  ' received from the HA7Nets contain enough information that the
  ' client should be able to figure out how to communicate
  ' directly to the HA7Net.
  ' 
  ' Returns
  ' An ArrayList of HA7Net.DiscoveryResponse, which details the
  ' information about each HA7Net that was discovered.           
  Shared Function DiscoverHA7Nets() As ArrayList

    Dim HA7Nets As New ArrayList
    Dim HA7NetResponse As DiscoveryResponse
    Dim HA7NetResponseIP As New Specialized.StringDictionary

    Dim multicastDestination As New System.Net.IPEndPoint(System.Net.IPAddress.Parse("224.1.2.3"), 4567)
    Dim receiveBytes(256) As Byte 'Buffer to hold incoming network data
    Dim RemoteIpEndPoint As New System.Net.IPEndPoint(System.Net.IPAddress.Any, 0)
    Dim txPacket As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes("HA" & Chr(0) & Chr(1))

    Dim IPAddresssList As ArrayList = NetworkAdapter.GetConfiguration

    '
    ' Enumerate through the sensors and update the assocated HomeSeer device
    '
    Dim MyEnum As IEnumerator = IPAddresssList.GetEnumerator

    While MyEnum.MoveNext

      Dim strIPAddress As String = CStr(MyEnum.Current.ToString)

      Dim udpSocket As New System.Net.Sockets.Socket(Net.Sockets.AddressFamily.InterNetwork, _
                                                     Net.Sockets.SocketType.Dgram, _
                                                     Net.Sockets.ProtocolType.Udp)

      WriteMessage("Sending HA7Net multicast query to " & strIPAddress & " ...", MessageType.Debug)

      Try
        '
        ' Transmit the Multicast Packet
        '
        Dim LocalIpEndPoint As New System.Net.IPEndPoint(System.Net.IPAddress.Parse(strIPAddress), 0)
        udpSocket.Bind(LocalIpEndPoint)
        udpSocket.SendTo(txPacket, multicastDestination)

        '
        ' Give the responses 2 seconds to come back
        '
        System.Threading.Thread.Sleep(1000 * 2)

        If udpSocket.Available() Then
          udpSocket.ReceiveFrom(receiveBytes, receiveBytes.Length, Net.Sockets.SocketFlags.None, RemoteIpEndPoint)
          HA7NetResponse = ParseMulticastResponse(receiveBytes)

          If Not (HA7NetResponse Is Nothing) Then
            'Found a valid response...
            'Store the IP address that this response came from
            HA7NetResponse.IPAddress = RemoteIpEndPoint.Address

            '
            ' Attempt to prevent the same interface from being added more than once
            '
            Dim strKey As String = HA7NetResponse.SerialNumber
            If HA7NetResponseIP.ContainsKey(strKey) = False Then
              'It is safe to add this HA7Net to the list
              HA7Nets.Add(HA7NetResponse)
              HA7NetResponseIP.Add(strKey, strKey)
            End If
          End If

        End If

      Catch pEx As Exception

      Finally
        '
        ' Close the UDP socket
        '
        udpSocket.Close()
      End Try

    End While

    Return HA7Nets

  End Function

  ' Parses rawBytes looking for a valid response to a multicast discovery packet.
  ' If a valid response is found, then returns a DiscoveryResponse object containing 
  ' the data parsed from the response, which should contain information regarding how
  ' to communicate directly with the HA7Net.
  ' It rawBytes does not contain a valid response, then returns Nothing.
  Shared Function ParseMulticastResponse(ByVal rawBytes As Byte()) As DiscoveryResponse

    Dim deviceNameLength As Int16
    Dim retVal As DiscoveryResponse
    Dim serialNumberLength As Int16

    retVal = Nothing

    If (System.Text.Encoding.ASCII.GetString(rawBytes, 0, 2)) = "HA" Then
      If ((rawBytes(2) * 256) + (rawBytes(3))) = &H8001 Then
        'This seems like an HA7 response to the multicast discovery request
        retVal = New DiscoveryResponse
        retVal.Signature = (System.Text.Encoding.ASCII.GetString(rawBytes, 0, 2))
        retVal.Command = (rawBytes(2) * 256) + rawBytes(3)
        retVal.Port = (rawBytes(4) * 256) + rawBytes(5)
        retVal.SSLPort = (rawBytes(6) * 256) + rawBytes(7)

        retVal.SerialNumber = System.Text.Encoding.ASCII.GetString(rawBytes, 8, 12)
        'Check for null termination of serial number, and trim the return value to match
        serialNumberLength = InStr(retVal.SerialNumber, Chr(0)) - 1
        If (serialNumberLength > 0) Then
          retVal.SerialNumber = Left$(retVal.SerialNumber, serialNumberLength)
        End If

        retVal.DeviceName = System.Text.Encoding.ASCII.GetString(rawBytes, 20, 64)
        'Check for null termination of device name, and trim the return value to match
        deviceNameLength = InStr(retVal.DeviceName, Chr(0)) - 1
        If (deviceNameLength > 0) Then
          retVal.DeviceName = Left$(retVal.DeviceName, deviceNameLength)
        End If
      End If
    End If

    Return retVal

  End Function

  ''' <summary>
  ''' Builds the HTTPWebRequest
  ''' </summary>
  ''' <param name="url"></param>
  ''' <param name="parameters"></param>
  ''' <returns></returns>
  Private Function buildHttpWebRequest(ByVal url As String, ByVal parameters As ArrayList) As System.Net.HttpWebRequest

    Dim httpWebReqeust As System.Net.HttpWebRequest

    Dim uri As String = ""
    Dim port As Integer = 0

    '
    ' Determine if we are using HTTPS
    '
    If UseSSL Then
      uri = "https://"
      port = HTTPSPort
    Else
      uri = "http://"
      port = HTTPPort
    End If

    '
    ' Build URI
    '
    uri = String.Format("{0}{1}:{2}", uri, HostName, port.ToString)

    If parameters Is Nothing Then
      parameters = New ArrayList
    End If

    '
    ' Append the lockid to the parameter list, if one exists
    '
    If Me.LockID <> "" Then
      Dim lockParameter As New HttpParameter("LockID", Me.LockID)
      parameters.Add(lockParameter)
    End If

    '
    ' Append any optional parameters to the URL
    '
    Dim firstParameter As Boolean = True
    For Each parameter As HttpParameter In parameters
      If firstParameter Then
        url &= "?"
        firstParameter = False
      Else
        url &= "&"
      End If
      url &= String.Format("{0}={1}", parameter.Name, parameter.Value)
    Next

    '
    ' Create the HttpWebRequest
    '
    httpWebReqeust = System.Net.WebRequest.Create(uri & url)

    Return httpWebReqeust

  End Function

  ''' <summary>
  ''' Returns the HTML Response
  ''' </summary>
  ''' <param name="Response"></param>
  ''' <returns></returns>
  Private Function getHTML(ByVal Response As System.Net.HttpWebResponse) As String

    Try
      '
      ' Extracts the HTML from the httpWebResponse object
      '
      Dim responseStream As System.IO.Stream = Response.GetResponseStream()
      Dim responseStreamReader As New System.IO.StreamReader(responseStream)

      Return responseStreamReader.ReadToEnd()

    Catch ex As Exception
      Return ""
    End Try

  End Function

  '
  ' Parses the result time, and any exceptions from the response, and places these
  ' values into the OW_Response structure.
  '
  Private Sub ParseResultMetaData(ByRef retVal As OW_Response, ByRef htmlData As String)

    Dim re As Regex
    Dim m As Match
    Dim epochTime As New System.DateTime(1970, 1, 1)

    Try

      '
      ' Calculate response time, and store into the OW_Response structure
      '  NAME="Completed_0" VALUE="1147829057"
      '
      re = New Regex("INPUT.*NAME=" & Chr(34) & "Completed_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*VALUE=" & Chr(34) & "(?<Value>([^" & Chr(34) & "]*))" & Chr(34))
      m = re.Match(htmlData)
      If Not m Is Nothing Then
        retVal.ResponseTime = epochTime.Add(New TimeSpan(0, 0, CLng(m.Groups("Value").ToString())))
      End If

      '
      ' Get the HA7Net Exception Data Code
      '
      re = New Regex("INPUT.*NAME=" & Chr(34) & "Exception_Code_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*VALUE=" & Chr(34) & "(?<Value>([^" & Chr(34) & "]*))" & Chr(34))
      m = re.Match(htmlData)
      If Not m Is Nothing Then
        retVal.Exception_Code = m.Groups("Value").ToString()
      End If

      '
      ' Get the HA7net Description
      '
      re = New Regex("INPUT.*NAME=" & Chr(34) & "Exception_String_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*VALUE=" & Chr(34) & "(?<Value>([^" & Chr(34) & "]*))" & Chr(34))
      m = re.Match(htmlData)
      If Not m Is Nothing Then
        retVal.Exception_Description = m.Groups("Value").ToString()
      End If

    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "ParseResultMetaData()")
    End Try

  End Sub

  ''' <summary>
  ''' Retrieves response from HA7Net Device
  ''' </summary>
  ''' <param name="whatToSend"></param>
  ''' <returns></returns>
  Private Function txHA7(ByVal whatToSend As System.Net.HttpWebRequest) As System.Net.HttpWebResponse

    Dim retVal As System.Net.HttpWebResponse

    Try

      WriteMessage(String.Format("Getting URL: {0}", whatToSend.Address.PathAndQuery), MessageType.Debug)
      retVal = CType(whatToSend.GetResponse(), System.Net.HttpWebResponse)

    Catch pEx As System.Net.WebException
      '
      ' Proces program exception
      '
      Throw New Exception(pEx.Message)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      Throw New Exception(pEx.Message)
    End Try

    Return retVal

  End Function

  ' Implements the 1-Wire search algorithm including each of the
  ' regular search, family search, and conditional search
  ' functionalities. This function is used to discover the 64-bit
  ' ROM codes (addresses) of devices connected to the 1-Wire bus.
  ' The search function can optionally restrict the returned list
  ' \of addresses to include only those devices belonging to a
  ' particular family, and/or those that are in a device defined
  ' conditional state.
  ' 
  ' Parameters
  ' FamilyCode :   One byte decimal number used to restrict the results
  '                to 1\-Wire devices belonging to the given family.
  ' Conditional :  True if only devices in a conditional state are to be
  '                returned.
  ' 
  ' Returns
  ' Addresses –Table that contains a list of 8 byte 1-Wire ROM
  ' address codes, with each ROM code residing a text field named
  ' ‘Address_x’ where x is a 0 based sequential integer. This is
  ' stored as an ArrayList of Strings in the Data field of the
  ' OWInterFace.OW_Response object.                                     
  Public Function OW_SearchROM(Optional ByVal FamilyCode As Byte = 0, Optional ByVal Conditional As Boolean = False) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      '
      ' Process the optional parameters
      '
      If FamilyCode <> 0 Then
        parameterList.Add(New HttpParameter("FamilyCode", Hex2(FamilyCode)))
      End If
      If Conditional = True Then
        parameterList.Add(New HttpParameter("Conditional", "TRUE"))
      End If

      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_SEARCH, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      '
      ' Copy the 1-Wire Addresses into a string array for returning to caller
      '
      Dim re As New Regex("INPUT.*NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      Dim m As Match
      For Each m In re.Matches(htmlData)
        retData.Add(m.Groups("ROMId").ToString(), "")
      Next
      retVal.Data = retData

      '
      ' Get stats / exception info
      '
      ParseResultMetaData(retVal, htmlData)

    Catch pEx As Exception
      '
      ' Proces program exception
      '
      WriteMessage(String.Format("Unable to connect to HA7Net [{0}]: {1}", Me.HostName, pEx.Message), MessageType.Error)
    End Try

    Return retVal

  End Function

  ''' <summary>
  ''' Read HA7Net Counter Data
  ''' </summary>
  ''' <param name="Address_Array"></param>
  ''' <returns></returns>
  Public Function OW_ReadCounter(ByVal Address_Array As String) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      ' Process the required address parameter
      parameterList.Add(New HttpParameter("Address_Channel_Array", Address_Array))
      '/1Wire/ReadCounter.html?Address_Channel_Array={ B80000000451C61D,A},{ B80000000451C61D,B}

      '
      ' Build the http request, including any standard parameters
      '
      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_READ_COUNTER, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      ' <INPUT CLASS="HA7Value" NAME="Exception_Code_0" ID="Exception_Code_0" TYPE="hidden" VALUE="0" Size="5" disabled>
      ' <INPUT CLASS="HA7Value" NAME="Exception_String_0" ID="Exception_String_0" TYPE="hidden" VALUE="None" Size="5" disabled>
      ' <INPUT CLASS="HA7Value" NAME="Address_0" ID="Address_0" TYPE="text" VALUE="1F0000000770011D">
      ' <INPUT CLASS="HA7Value" NAME="Count_0" ID="Count_0" TYPE="text" VALUE="16">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_0" ID="Device_Exception_0" TYPE="text" VALUE="OK">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_Code_0" ID="Device_Exception_Code_0" TYPE="hidden" VALUE="0">
      ' <INPUT CLASS="HA7Value" NAME="Address_1" ID="Address_1" TYPE="text" VALUE="1F0000000770011D">
      ' <INPUT CLASS="HA7Value" NAME="Count_1" ID="Count_1" TYPE="text" VALUE="0">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_1" ID="Device_Exception_1" TYPE="text" VALUE="OK">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_Code_1" ID="Device_Exception_Code_1" TYPE="hidden" VALUE="0">

      '
      ' Address data is returned for this request
      '
      Dim AddressRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      For Each AddressMatch As Match In AddressRegex.Matches(htmlData)
        Dim RecordNumber As String = AddressMatch.Groups("RecordNumber").Value
        Dim ROMId As String = AddressMatch.Groups("ROMId").Value

        '
        ' Set up the data for the return
        '
        If retData.ContainsKey(ROMId) = False Then
          retData.Add(ROMId, New Hashtable)
          retData(ROMId)("ValueA") = "0"
          retData(ROMId)("ValueB") = "0"
          retData(ROMId)("Channel") = 0
        End If

        '
        ' Get the Counter value
        '
        Dim CounterRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Count_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In CounterRegex.Matches(htmlData)

          If retData(ROMId).ContainsKey("Channel") Then

            retData(ROMId)("Channel") += 1
            Dim iChannel As Integer = retData(ROMId)("Channel")

            Select Case iChannel
              Case 1
                retData(ROMId)("ValueA") = m.Groups("Value").Value
              Case 2
                retData(ROMId)("ValueB") = m.Groups("Value").Value
            End Select

          End If

        Next

      Next

      retVal.Data = retData

      'Get stats / exception info
      ParseResultMetaData(retVal, htmlData)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "OW_ReadCounter()")
    End Try

    'Return results
    Return retVal

  End Function

  ''' <summary>
  ''' Read HA7Net Switch Data
  ''' </summary>
  ''' <param name="Address_Array"></param>
  ''' <returns></returns>
  Public Function OW_ReadSwitch(ByVal Address_Array As String) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      ' Process the required address parameter
      parameterList.Add(New HttpParameter("SwitchRequest", Address_Array))
      'http://192.168.2.5/1Wire/SwitchControl.html?SwitchRequest={1F0000000770011D, {{1, Read, False}}}
      'http://192.168.2.5/1Wire/SwitchControl.html?SwitchRequest={1F0000000770011D, {{2, Read, False}}}

      '
      ' Build the http request, including any standard parameters
      '
      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_READ_SWITCH, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      ' <INPUT CLASS="HA7Value" NAME="Address_0" ID="Address_0" TYPE="text" VALUE="2B000000346E3612">
      ' <INPUT CLASS="HA7Value" NAME="Channel_0" ID="Channel_0" TYPE="text" VALUE="1">
      ' <INPUT CLASS="HA7Value" NAME="OriginalState_0" ID="OriginalState_0" TYPE="text" VALUE="InputLow">
      ' <INPUT CLASS="HA7Value" NAME="CurrentState_0" ID="CurrentState_0" TYPE="text" VALUE="InputLow">
      ' <INPUT CLASS="HA7Value" NAME="Activity_0" ID="Activity_0" TYPE="text" VALUE="FALSE">
      ' <INPUT CLASS="HA7Value" NAME="Status_0" ID="Status_0" TYPE="text" VALUE="OK">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_Code_0" ID="Device_Exception_Code_0" TYPE="hidden" VALUE="0">
      ' <INPUT CLASS="HA7Value" NAME="Address_1" ID="Address_1" TYPE="text" VALUE="2B000000346E3612">
      ' <INPUT CLASS="HA7Value" NAME="Channel_1" ID="Channel_1" TYPE="text" VALUE="2">
      ' <INPUT CLASS="HA7Value" NAME="OriginalState_1" ID="OriginalState_1" TYPE="text" VALUE="InputLow">
      ' <INPUT CLASS="HA7Value" NAME="CurrentState_1" ID="CurrentState_1" TYPE="text" VALUE="InputLow">
      ' <INPUT CLASS="HA7Value" NAME="Activity_1" ID="Activity_1" TYPE="text" VALUE="FALSE">
      ' <INPUT CLASS="HA7Value" NAME="Status_1" ID="Status_1" TYPE="text" VALUE="OK">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_Code_1" ID="Device_Exception_Code_1" TYPE="hidden" VALUE="0">

      '
      ' Address data is returned for this request
      '
      Dim AddressRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      For Each AddressMatch As Match In AddressRegex.Matches(htmlData)
        Dim RecordNumber As String = AddressMatch.Groups("RecordNumber").Value
        Dim ROMId As String = AddressMatch.Groups("ROMId").Value

        '
        ' Set up the data for the return
        '
        If retData.ContainsKey(ROMId) = False Then
          retData.Add(ROMId, New Hashtable)
          retData(ROMId)("ValueA") = "0"
          retData(ROMId)("ValueB") = "0"
          retData(ROMId)("Channel") = 0
        End If

        Dim CounterRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "CurrentState_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In CounterRegex.Matches(htmlData)
          Dim Value As String = m.Groups("Value").Value

          If retData(ROMId).ContainsKey("Channel") Then

            retData(ROMId)("Channel") += 1
            Dim iChannel As Integer = retData(ROMId)("Channel")

            Select Case iChannel
              Case 1
                retData(ROMId)("ValueA") = IIf(Value = "InputLow", "0", "1")
              Case 2
                retData(ROMId)("ValueB") = IIf(Value = "InputLow", "0", "1")
            End Select

          End If

        Next

      Next

      retVal.Data = retData

      'Get stats / exception info
      ParseResultMetaData(retVal, htmlData)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "OW_ReadSwitch()")
    End Try

    'Return results
    Return retVal

  End Function

  ''' <summary>
  ''' Read HA7Net Temperature Data
  ''' </summary>
  ''' <param name="Address_Array"></param>
  ''' <returns></returns>
  Public Function OW_ReadTemperature(ByVal Address_Array As String) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      ' Process the required address parameter
      parameterList.Add(New HttpParameter("Address_Array", Address_Array))

      '
      ' Build the http request, including any standard parameters
      '
      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_READ_TEMPERATURE, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      ' <INPUT CLASS="HA7Value" NAME="Address_0" ID="Address_0" TYPE="text" VALUE="5F00080009478E10">
      ' <INPUT CLASS="HA7Value" NAME="Temperature_0" ID="Temperature_0" TYPE="text" VALUE="18.625">
      ' <INPUT CLASS="HA7Value" NAME="Resolution_0" ID="Resolution_0" TYPE="text" VALUE="9+">

      '
      ' Address data is returned for this request
      '
      Dim AddressRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      For Each AddressMatch As Match In AddressRegex.Matches(htmlData)
        Dim RecordNumber As String = AddressMatch.Groups("RecordNumber").Value
        Dim ROMId As String = AddressMatch.Groups("ROMId").Value

        '
        ' Set up the data for the return
        '
        retData.Add(ROMId, New Hashtable)
        retData(ROMId)("Temperature") = "-185"
        retData(ROMId)("Resolution") = "0"

        Dim TemperatureRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Temperature_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In TemperatureRegex.Matches(htmlData)
          retData(ROMId)("Temperature") = m.Groups("Value").Value
        Next

        Dim ResolutionRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Resolution_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In ResolutionRegex.Matches(htmlData)
          retData(ROMId)("Resolution") = m.Groups("Value").Value
        Next
      Next

      retVal.Data = retData

      'Get stats / exception info
      ParseResultMetaData(retVal, htmlData)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "OW_ReadTemperature()")
    End Try

    'Return results
    Return retVal

  End Function

  ''' <summary>
  ''' Read HA7Net DS18B20 Data
  ''' </summary>
  ''' <param name="Address_Array"></param>
  ''' <returns></returns>
  Public Function OW_ReadDS18B20(ByVal Address_Array As String) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      ' Process the required address parameter
      parameterList.Add(New HttpParameter("DS18B20Request", Address_Array))

      '
      ' Build the http request, including any standard parameters
      '
      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_READ_DS18B20, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      ' <INPUT CLASS="HA7Value" NAME="Address_0" ID="Address_0" TYPE="text" VALUE="2C0000025E996128">
      ' <INPUT CLASS="HA7Value" NAME="Temperature_0" ID="Temperature_0" TYPE="text" VALUE="20.6875">
      ' <INPUT CLASS="HA7Value" NAME="Resolution_0" ID="Resolution_0" TYPE="text" VALUE="12">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_0" ID="Device_Exception_0" TYPE="text" VALUE="OK">

      '
      ' Address data is returned for this request
      '
      Dim AddressRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      For Each AddressMatch As Match In AddressRegex.Matches(htmlData)
        Dim RecordNumber As String = AddressMatch.Groups("RecordNumber").Value
        Dim ROMId As String = AddressMatch.Groups("ROMId").Value

        '
        ' Set up the data for the return
        '
        retData.Add(ROMId, New Hashtable)
        retData(ROMId)("Temperature") = "-185"
        retData(ROMId)("Resolution") = "0"

        Dim TemperatureRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Temperature_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In TemperatureRegex.Matches(htmlData)
          retData(ROMId)("Temperature") = m.Groups("Value").Value
        Next

        Dim ResolutionRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Resolution_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In ResolutionRegex.Matches(htmlData)
          retData(ROMId)("Resolution") = m.Groups("Value").Value
        Next
      Next

      retVal.Data = retData

      'Get stats / exception info
      ParseResultMetaData(retVal, htmlData)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "OW_ReadDS18B20()")
    End Try

    'Return results
    Return retVal

  End Function

  ''' <summary>
  ''' Read HA7Net Humidity Data
  ''' </summary>
  ''' <param name="Address_Array"></param>
  ''' <returns></returns>
  Public Function OW_ReadHumidity(ByVal Address_Array As String) As HA7NetServer.OW_Response

    Dim HA7NetResponse As System.Net.HttpWebResponse
    Dim parameterList As New ArrayList
    Dim retData As New Hashtable
    Dim retVal As New OW_Response

    Dim htmlData As String

    Try
      '
      ' Make sure we return something on error
      '
      retVal.Data = retData

      ' Process the required address parameter
      parameterList.Add(New HttpParameter("Address_Array", Address_Array))

      '
      ' Build the http request, including any standard parameters
      '
      Dim HA7NetRequest As System.Net.HttpWebRequest = buildHttpWebRequest(URL_READ_HUMIDITY, parameterList)

      '
      ' Get data from the HA7NET
      '
      HA7NetResponse = txHA7(HA7NetRequest)
      htmlData = getHTML(HA7NetResponse)

      ' <INPUT CLASS="HA7Value" NAME="Address_0" ID="Address_0" TYPE="text" VALUE="3600000085459726">
      ' <INPUT CLASS="HA7Value" NAME="Humidity_0" ID="Humidity_0" TYPE="text" VALUE="37.3519">
      ' <INPUT CLASS="HA7Value" NAME="Temperature_0" ID="Temperature_0" TYPE="text" VALUE="19.5312">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_0" ID="Device_Exception_0" TYPE="text" VALUE="OK">
      ' <INPUT CLASS="HA7Value" NAME="Device_Exception_Code_0" ID="Device_Exception_Code_0" TYPE="hidden" VALUE="0">

      '
      ' Address data is returned for this request
      '
      Dim AddressRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Address_(?<RecordNumber>([0-9A-F]+))" & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<ROMId>([0-9A-F]{16}))" & Chr(34))
      For Each AddressMatch As Match In AddressRegex.Matches(htmlData)
        Dim RecordNumber As String = AddressMatch.Groups("RecordNumber").Value
        Dim ROMId As String = AddressMatch.Groups("ROMId").Value

        '
        ' Set up the data for the return
        '
        retData.Add(ROMId, New Hashtable)
        retData(ROMId)("Temperature") = "-185"
        retData(ROMId)("Humidity") = "0"

        Dim TemperatureRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Temperature_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In TemperatureRegex.Matches(htmlData)
          retData(ROMId)("Temperature") = m.Groups("Value").Value
        Next

        Dim HumidityRegex As New Regex("INPUT.*?NAME=" & Chr(34) & "Humidity_" & RecordNumber & Chr(34) & "\B.*?VALUE=" & Chr(34) & "(?<Value>(.*?))" & Chr(34))
        For Each m As Match In HumidityRegex.Matches(htmlData)
          retData(ROMId)("Humidity") = m.Groups("Value").Value
        Next
      Next

      retVal.Data = retData

      'Get stats / exception info
      ParseResultMetaData(retVal, htmlData)
    Catch pEx As Exception
      '
      ' Proces program exception
      '
      ProcessError(pEx, "OW_ReadHumidity()")
    End Try

    'Return results
    Return retVal

  End Function

  ''' <summary>
  ''' Convert Temperature Reading
  ''' </summary>
  ''' <param name="value"></param>
  ''' <param name="Farenheight"></param>
  ''' <returns></returns>
  Public Function ConvertTemperature(ByVal value As String, ByVal Farenheight As Boolean) As Single

    Dim temperature As Single = -185

    Try

      temperature = Single.Parse(value, nfi)

      If Farenheight = True Then
        Return (temperature * 1.8@) + 32@
      Else
        Return temperature
      End If

    Catch pEx As Exception
      Return temperature
    End Try

  End Function

End Class
