Imports System.Net.NetworkInformation

Public Class NetworkAdapter

  Public Shared Function GetConfiguration() As ArrayList

    Dim IPAddressList As New ArrayList

    Try
      Dim networkinterfaces() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces()
      Dim found As Boolean = False

      For Each ni As NetworkInterface In networkinterfaces

        If ni.NetworkInterfaceType <> NetworkInterfaceType.Loopback And ni.OperationalStatus = OperationalStatus.Up Then
          For Each UnicastIPAddress As UnicastIPAddressInformation In ni.GetIPProperties().UnicastAddresses
            If UnicastIPAddress.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then

              Dim ipaddress As String = UnicastIPAddress.Address.ToString()
              If IPAddressList.Contains(ipaddress) = False Then
                IPAddressList.Add(ipaddress)
              End If

            End If

          Next
        End If

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "GetConfiguration()")
    End Try

    Return IPAddressList

  End Function

End Class
