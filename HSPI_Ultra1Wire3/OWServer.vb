Imports System.Globalization
Imports System.Text.RegularExpressions

Public Class OWServer

  Public HostName As String = ""        ' Hostname of the OWServer represented by this instance
  Public HTTPPort As Integer = 80       ' Port number of the http server on the OWServer represented by this instance
  Public HTTPSPort As Integer = 443     ' Port number of the https server on the OWServer represented by this instance
  Public UseSSL As Boolean = False      ' Flag to indicate if all http communications with this OWServer should be made via https
  Public InterfaceID As String = ""     ' Tracks Interface ID
  Public InterfaceType As String = ""   ' Tracks Interface Type
  Public DeviceId As Integer = 0

  '
  ' Contains data that describes a single OWServer found
  ' by using the UDP broadcast discovery mechanism.
  '
  Public Class DiscoveryResponse
    Public Signature As String = ""           ' Unused for OWServer models
    Public Command As Int32                   ' Unused for OWServer models
    Public Port As Int32                      ' Port number the http server can be reached on
    Public SSLPort As Int32                   ' Port number the https server can be reached on
    Public SerialNumber As String             ' MAC address of the OWServer
    Public DeviceName As String               ' User configurable device name returned by the OWServer
    Public DeviceType As String               ' OWServer Type
    Public IPAddress As System.Net.IPAddress  ' IP Address of the OWServer
  End Class

  '
  ' Parameters used to build the URL of requests made to the HA7Net
  '
  Private Class HttpParameter

    Public Name As String = ""
    Public Value As String = ""

    ' Default constructor for HttpParameter class
    Public Sub New(ByVal Name As String, ByVal Value As String)
      Me.Name = Name
      Me.Value = Value
    End Sub

  End Class

  ' Default OWServer constructor
  ' 
  ' Parameters
  ' HostName :  String representation of the hostname or IP address of
  '             the OWServer represented by this instance.              
  Public Sub New(ByVal DeviceId As Integer, ByVal HostName As String, InterfaceID As String, InterfaceType As String)

    Me.HostName = HostName
    Me.InterfaceType = InterfaceType
    Me.InterfaceID = InterfaceID
    Me.DeviceId = DeviceId

  End Sub

  ' Uses a UDP broadcast to discover any HA7Nets within
  ' reach of a UDP packet. This mechanism works by
  ' transmitting a particularly formatted packet to the broadcast
  ' address on UDP port 30303. Any OWServers that hear the
  ' packet will respond by transmitting a directed UDP packet
  ' back to the client that sent the UDP packet. The packet
  ' received from the OWServer contain enough information that the
  ' client should be able to figure out how to communicate
  ' directly to the OWServer.
  ' 
  ' Returns
  ' An ArrayList of OWServer.DiscoveryResponse, which details the
  ' information about each OWServer that was discovered.           
  Shared Function DiscoverOWServers() As ArrayList

    Dim OWServers As New ArrayList
    Dim OWServerResponse As DiscoveryResponse
    Dim OWServerResponseIP As New Specialized.StringDictionary

    Dim broadcast As New System.Net.IPEndPoint(System.Net.IPAddress.Broadcast, 30303)
    Dim RemoteIpEndPoint As New System.Net.IPEndPoint(System.Net.IPAddress.Any, 0)

    Dim receiveBytes(256) As Byte 'Buffer to hold incoming network data
    Dim sendbuf As Byte() = System.Text.Encoding.ASCII.GetBytes("Discovery: Who is out there?")

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

      WriteMessage("Sending OWServer broadcast query to " & strIPAddress & " ...", MessageType.Debug)

      Try

        '
        ' Transmit the UDP Broadcast Packet
        '
        udpSocket.SetSocketOption(System.Net.Sockets.SocketOptionLevel.Socket, System.Net.Sockets.SocketOptionName.Broadcast, 1)
        udpSocket.SendTo(sendbuf, broadcast)

        '
        ' Give the responses 2 seconds to come back
        '
        System.Threading.Thread.Sleep(1000 * 2)

        While udpSocket.Available() > 0
          udpSocket.ReceiveFrom(receiveBytes, receiveBytes.Length, Net.Sockets.SocketFlags.None, RemoteIpEndPoint)
          OWServerResponse = ParseResponse(receiveBytes)

          If Not (OWServerResponse Is Nothing) Then
            '
            ' Store the IP address that this response came from
            '
            OWServerResponse.IPAddress = RemoteIpEndPoint.Address

            '
            ' Attempt to prevent the same interface from being added more than once
            '
            Dim strKey As String = OWServerResponse.SerialNumber

            If OWServerResponseIP.ContainsKey(strKey) = False Then
              '
              ' It is safe to add this OWServer to the list
              '
              OWServers.Add(OWServerResponse)
              OWServerResponseIP.Add(strKey, strKey)
            End If
          End If

        End While

      Catch pEx As Exception

      Finally
        '
        ' Close the UDP socket
        '
        udpSocket.Close()
      End Try

    End While

    Return OWServers

  End Function

  ' Parses rawBytes looking for a valid response to a UDP discovery packet.
  ' If a valid response is found, then returns a DiscoveryResponse object containing 
  ' the data parsed from the response, which should contain information regarding how
  ' to communicate directly with the OWServer.
  ' It rawBytes does not contain a valid response, then returns Nothing.
  Shared Function ParseResponse(ByVal rawBytes As Byte()) As DiscoveryResponse

    Dim retVal As DiscoveryResponse

    retVal = Nothing

    Try

      retVal = New DiscoveryResponse
      retVal.Signature = ""
      retVal.Command = 0
      retVal.Port = 80
      retVal.SSLPort = 0

      ' {NETBios: OW-SERVER-ENET , MAC: 00-50-C2-91-B0-F5, IP: 192.168.2.9,   Product: OW_SERVER-Enet,          FWVer: 2.22, Name: OW-SERVER-ENET,          HTTPPort: 80, Bootloader: TFTP, TCPIntfPort: 0, }
      ' {NETBios: EDSWIRELESSCTRL, MAC: 00-04-A3-BE-90-DB, IP: 192.168.2.118, Product: WirelessController-Enet, FWVer: 1.41, Name: WirelessController-Enet, HTTPPort: 80, Bootloader: POST, TCPIntfPort: 0 }

      Dim strResponse As String = System.Text.Encoding.UTF8.GetString(rawBytes).Replace("""", "")

      Dim SerialNumber As String = Regex.Match(strResponse, "MAC:\s(?<MAC>([^,]+)),").Groups("MAC").ToString()
      Dim DeviceName As String = Regex.Match(strResponse, "NETBios:\s(?<NETBios>([^,]+)),").Groups("NETBios").ToString()
      Dim Product As String = Regex.Match(strResponse, "Product:\s(?<Product>([^,]+)),").Groups("Product").ToString()
      Dim HTTPPort As String = Regex.Match(strResponse, "HTTPPort:\s(?<HTTPPort>(\d+)),").Groups("HTTPPort").ToString()

      Dim DeviceType As String = "OWServer"
      If Regex.IsMatch(Product, "Wireless", RegexOptions.IgnoreCase) = True Then
        DeviceType = "WirelessController"
      End If

      If IsNumeric(HTTPPort) Then
        retVal.Port = Integer.Parse(HTTPPort)
      End If

      retVal.DeviceType = DeviceType
      retVal.DeviceName = DeviceName.Trim
      retVal.SerialNumber = SerialNumber.Trim

      If retVal.DeviceName.Length = 0 Then Return Nothing
      If retVal.SerialNumber.Length = 0 Then Return Nothing

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

    Return retVal

  End Function

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

  Public Function ConvertPressure(ByVal value As String, ByVal Metric As Boolean) As Single

    Dim pressure As Single = 0.0

    Try
      pressure = Math.Abs(Single.Parse(value, nfi))
      If Metric = False Then pressure *= 0.02953
    Catch ex As Exception

    End Try

    Return pressure

  End Function

End Class

