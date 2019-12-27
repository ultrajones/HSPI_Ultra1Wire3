Imports System.Threading
Imports System.Text.RegularExpressions
Imports System.Web.Script.Serialization
Imports System.Text
Imports System.Net
Imports System.Xml
Imports System.Data.Common
Imports System.Data.SQLite
Imports System.IO

Module hspi_plugin

  Public Const IFACE_NAME As String = "Ultra1Wire3"

  Public Const LINK_TARGET As String = "hspi_ultra1wire3/hspi_ultra1wire3.aspx"
  Public Const LINK_URL As String = "hspi_ultra1wire3.html"
  Public Const LINK_TEXT As String = "Ultra1Wire3"
  Public Const LINK_PAGE_TITLE As String = "Ultra1Wire3 HSPI"
  Public Const LINK_HELP As String = "/hspi_ultra1wire3/Ultra1Wire3_HSPI_Users_Guide.pdf"

  Public gBaseCode As String = ""
  Public gIOEnabled As Boolean = True
  Public gEthernetMode As Boolean = False
  Public gImageDir As String = "/images/hspi_ultra1wire3/"
  Public gHSInitialized As Boolean = False
  Public gINIFile As String = "hspi_" & IFACE_NAME.ToLower & ".ini"

  Public gUnitType As String = "US"
  Public gTempDegreeUnit As Boolean = True
  Public gTempDegreeIcon As Boolean = False

  Public TEMP08Interfaces As New List(Of TEMP08)
  Public HA7EInterfaces As New List(Of HA7E)
  Public HA7NetInterfaces As New List(Of HA7NetServer)
  Public OWServerInterfaces As New List(Of OWServer)
  Public OneWireSensors As New List(Of OneWireSensor)

  Public DatabaseInsertQueue As New Queue

  Public gMonitoring As Boolean = True

  Public gRTCReceived As Boolean = False
  Public gDeviceValueType As String = "1"

  Public strCmdWait As String = ""
  Public iCmdAttempt As Byte = 0

  Public MAX_DATABASE_QUEUE As Integer = 5000

  Public DEV_DATABASE_INTERFACE As Byte = 1
  Public DEV_PLUGIN_INTERFACE As Byte = 2

#Region "HSPI - Misc"

  ''' <summary>
  ''' Gets plug-in setting from INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValueDefault"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSetting(ByVal strSection As String, _
                             ByVal strKey As String, _
                             ByVal strValueDefault As String) As String

    Dim strMessage As String = ""

    Try
      strMessage = "Entered GetSetting() function."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Get the ini settings
      '
      Dim strValue As String = hs.GetINISetting(strSection, strKey, strValueDefault, gINIFile)

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      Return strValue

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSetting()")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Saves plug-in setting to INI file
  ''' </summary>
  ''' <param name="strSection"></param>
  ''' <param name="strKey"></param>
  ''' <param name="strValue"></param>
  ''' <remarks></remarks>
  Public Sub SaveSetting(ByVal strSection As String, _
                         ByVal strKey As String, _
                         ByVal strValue As String)

    Dim strMessage As String = ""

    Try
      strMessage = "Entered SaveSetting() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Check to see if we need to encrypt the data
      '
      If strKey = "UserPass" Then
        If strValue.Length = 0 Then Exit Sub
        strValue = hs.EncryptString(strValue, "&Cul8r#1")
      End If

      strMessage = String.Format("Section: {0}, Key: {1}, Value: {2}", strSection, strKey, strValue)
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Save the settings
      '
      hs.SaveINISetting(strSection, strKey, strValue, gINIFile)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "SaveSetting()")
    End Try

  End Sub

#End Region

#Region "HS Sensors"

  ''' <summary>
  ''' Returns plug-in statistics
  ''' </summary>
  ''' <param name="StatisticsType"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetStatistics(ByVal StatisticsType As String) As ULong

    Select Case StatisticsType
      Case "TEMP08Interfaces" : Return TEMP08Interfaces.Count
      Case "HA7EInterfaces" : Return HA7EInterfaces.Count
      Case "HA7NetInterfaces" : Return HA7NetInterfaces.Count
      Case "OWServerInterfaces" : Return OWServerInterfaces.Count
      Case "DBInsQueue" : Return DatabaseInsertQueue.Count
      Case "DBInsSuccess" : Return gDBInsertSuccess
      Case "DBInsFailure" : Return gDBInsertFailure
      Case Else
        Return 0
    End Select

  End Function

  ''' <summary>
  ''' Gets the chart data from the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="ChartType"></param>
  ''' <param name="AMChartType"></param>
  ''' <param name="strEndDate"></param>
  ''' <param name="Interval"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSensorChartJSON(ByVal device_id As Integer, _
                                     ByVal ChartType As String, _
                                     ByVal AMChartType As String, _
                                     ByVal strEndDate As String, _
                                     ByVal Interval As String) As String

    Dim JSON As New StringBuilder(1024)

    Try
      '
      ' Determine Chart Options
      '
      Dim strDBTable As String = ""
      Dim strDBField As String = ""

      Select Case ChartType
        Case "Counter"
          strDBTable = "tblCounterHistory"
          strDBField = "SUM(value)"
        Case "Humidity"
          strDBTable = "tblHumidityHistory"
          strDBField = "AVG(value)"
        Case "Temperature"
          strDBTable = "tblTemperatureHistory"
          strDBField = "AVG(value)"
        Case "Pressure"
          strDBTable = "tblPressureHistory"
          strDBField = "AVG(value)"
        Case Else : Return ""
      End Select

      Select Case AMChartType
        Case "column", "line", "step", "smoothedLine"
        Case Else
          AMChartType = "smoothedLine"
      End Select

      Dim dEndDateTime As Date
      Date.TryParse(strEndDate, dEndDateTime)

      '
      ' Calcualate the dates
      '
      If dEndDateTime > DateTime.Now Then dEndDateTime = DateTime.Now
      dEndDateTime = New DateTime(dEndDateTime.Year, dEndDateTime.Month, dEndDateTime.Day, 23, 59, 59)

      Dim DateInterval As DateInterval = DateInterval.Day
      Dim DateNumber As Double = 0

      Dim iInterval As Integer = 60 * 60 * 24
      Select Case Interval
        Case "1 day"
          iInterval = 60 * 60
          DateInterval = DateInterval.Day
          DateNumber = -1
        Case "1 week"
          iInterval = 60 * 60
          DateInterval = DateInterval.Day
          DateNumber = -7
        Case "1 month"
          iInterval = 60 * 60
          DateInterval = DateInterval.Month
          DateNumber = -1
        Case "3 months"
          iInterval = 60 * 60 * 24
          DateInterval = DateInterval.Month
          DateNumber = -3
        Case "6 months"
          iInterval = 60 * 60 * 24
          DateInterval = DateInterval.Month
          DateNumber = -6
        Case "1 year"
          iInterval = 60 * 60 * 24
          DateInterval = DateInterval.Month
          DateNumber = -12
      End Select

      '
      ' Calculate the begin date
      '
      Dim dBeginDateTime As DateTime = DateAdd(DateInterval, DateNumber, dEndDateTime)

      '
      ' Convert date to long (EPOCH)
      '
      Dim lngBeginTS As Long = ConvertDateTimeToEpoch(dBeginDateTime)
      Dim lngEndTS As Long = ConvertDateTimeToEpoch(dEndDateTime)

      Dim strDateQuery As String = String.Format("ts >= {0} AND ts <= {1}", lngBeginTS.ToString, lngEndTS.ToString)

      '
      ' Sorted lists for devices and results
      '
      Dim MyResults As New SortedList()

      '
      ' Build SQL Query
      '      
      Dim strSQL As String = String.Format("SELECT {0}*(ts/{0}) as ts_epoch, {1} as value, sensor_id " _
                                         & "FROM {2} as A " _
                                         & "WHERE device_id >= {3} " _
                                         & "AND {4} " _
                                         & "GROUP BY ts_epoch, sensor_id " _
                                         & "ORDER BY ts_epoch ASC", iInterval.ToString, strDBField, strDBTable, device_id.ToString, strDateQuery)

      WriteMessage(strSQL, MessageType.Debug)

      '
      ' Execute Query
      '
      Dim MyDataSet As Data.DataSet = QueryDatabase(strSQL)
      Using MyDataTable As Data.DataTable = MyDataSet.Tables(0)
        '
        ' Begin Process Query Results
        '     
        For Each MyRow As Data.DataRow In MyDataTable.Rows
          Dim ts As String = MyRow("ts_epoch")
          Dim sensor_id As Integer = Int32.Parse(MyRow("sensor_id"))
          Dim value As String = Convert.ToDecimal(MyRow("value")).ToString("F2")

          If MyResults.ContainsKey(ts) = False Then
            MyResults.Add(ts, New SortedList)
          End If
          MyResults(ts)(sensor_id) = value
        Next
      End Using

      '
      ' Start Chart Info
      '
      JSON.AppendLine("""chartinfo"":[")
      JSON.AppendLine("{")
      JSON.AppendFormat("""{0}"":""{1}"",", "Title", dEndDateTime.ToString)
      JSON.AppendFormat("""{0}"":""{1}"",", "Type", AMChartType)
      JSON.AppendFormat("""{0}"":""{1}""", "LabelSuffix", ChartType)
      JSON.AppendLine("}")
      JSON.AppendLine("],")

      ' Start Data
      JSON.AppendLine("""data"":[")

      Dim dayName As String = ""
      Dim i As Integer = 0
      For Each ts As String In MyResults.Keys
        Dim MyTS As DateTime = ConvertEpochToDateTime(ts)

        Dim strDate As String = ""
        Dim MyDayName As String = MyTS.ToString("yyyy-MM-dd HH:mm") ' MyTS.ToString("ddd, h tt")

        JSON.AppendLine("{")
        JSON.AppendFormat("""{0}"":""{1}"",", "date", MyDayName)

        Dim j As Integer = 0
        For Each sensor_id As Integer In MyResults(ts).Keys
          '
          ' If we have a value for the timestamp and channel, then use it
          '
          Dim value As String = ""
          If MyResults(ts).ContainsKey(sensor_id) = True Then
            value = MyResults(ts)(sensor_id)
          End If

          '
          ' Write the value (or an empty string if no value)
          '
          JSON.AppendFormat("""{0}"":{1}", sensor_id, value)
          j += 1
          If j < MyResults(ts).Count Then
            JSON.AppendLine(",")
          End If

        Next

        JSON.AppendLine("}")
        i += 1
        If i < MyResults.Count Then
          JSON.AppendLine(",")
        End If

      Next

      JSON.AppendLine("],")

      Dim hs_devices As Hashtable = GetHomeSeerDevices(ChartType)

      ' Start Sensors
      JSON.AppendLine("""sensors"":[")

      Using MyDataTable As DataTable = GetOneWireSensors(ChartType)
        Dim iRowCount As Integer = 0
        If MyDataTable.Columns.Contains("sensor_id") Then
          For Each row As DataRow In MyDataTable.Rows

            iRowCount += 1

            Dim sensor_id As String = row("sensor_id")
            Dim sensor_name As String = row("sensor_name")
            Dim sensor_type As String = row("sensor_type")
            Dim sensor_subtype As String = row("sensor_subtype")
            Dim sensor_units As String = row("sensor_units")
            Dim sensor_color As String = row("sensor_color")

            Dim dv_addr = row("sensor_addr")
            If hs_devices.ContainsKey(dv_addr) = True Then
              sensor_name = String.Format("{0}-{1}", hs_devices(dv_addr)("name"), hs_devices(dv_addr)("location"))
            End If

            JSON.Append("{")

            JSON.AppendFormat("""{0}"":""{1}"",", "Id", sensor_id)
            JSON.AppendFormat("""{0}"":""{1}"",", "Color", sensor_color)
            JSON.AppendFormat("""{0}"":""{1}"",", "Type", sensor_type)
            JSON.AppendFormat("""{0}"":""{1}"",", "Subtype", sensor_subtype)
            JSON.AppendFormat("""{0}"":""{1}"",", "Units", sensor_units)
            JSON.AppendFormat("""{0}"":""{1}""", "Name", sensor_name)

            If iRowCount < MyDataTable.Rows.Count Then
              JSON.AppendLine("},")
            Else
              JSON.AppendLine("}")
            End If

          Next row
        End If

      End Using

      JSON.AppendLine("]")

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Error)
    End Try

    Return "{" & JSON.ToString() & "}"

  End Function

  ''' <summary>
  ''' Returns the HomeSeer devices
  ''' </summary>
  ''' <param name="device_type"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function GetHomeSeerDevices(device_type As String) As Hashtable

    Dim hs_devices As New Hashtable

    Dim dv As Scheduler.Classes.DeviceClass

    Try

      Dim EN As Scheduler.Classes.clsDeviceEnumeration = hs.GetDeviceEnumerator
      If EN Is Nothing Then Throw New Exception(IFACE_NAME & " failed to get a device enumerator from HomeSeer.")

      Do
        dv = EN.GetNext
        If dv Is Nothing Then Continue Do
        If dv.Interface(Nothing) IsNot Nothing Then
          If dv.Interface(Nothing) = IFACE_NAME Then
            If dv.Device_Type_String(hs) = device_type Then
              Dim dv_addr As String = dv.Address(hs)
              If hs_devices.ContainsKey(dv_addr) = False Then
                Dim hs_device As New Hashtable
                hs_device.Add("name", dv.Name(hs))
                hs_device.Add("location", dv.Location(hs))
                hs_device.Add("location2", dv.Location2(hs))
                hs_devices.Add(dv_addr, hs_device)
              End If
            End If
          End If
        End If
      Loop Until EN.Finished

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
    End Try

    Return hs_devices

  End Function

  ''' <summary>
  ''' Get the 1-Wire Device List from the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireDeviceList() As ArrayList

    Dim OneWireDevices As New ArrayList

    Try
      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, device_name, device_type, device_conn, device_addr FROM tblDevices WHERE device_conn <> '{0}'", "Disabled")

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Using dtrResults As IDataReader = MyDbCommand.ExecuteReader()
            '
            ' Process the resutls
            '
            While dtrResults.Read()
              Dim OneWireDevice As New OneWireDevice

              OneWireDevice.device_id = dtrResults("device_id")
              OneWireDevice.device_name = dtrResults("device_name")
              OneWireDevice.device_type = dtrResults("device_type")
              OneWireDevice.device_conn = dtrResults("device_conn")
              OneWireDevice.device_addr = dtrResults("device_addr")

              OneWireDevices.Add(OneWireDevice)
            End While

            dtrResults.Close()
          End Using

        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "Get1WireDeviceList()")
    End Try

    Return OneWireDevices

  End Function

  ''' <summary>
  ''' Returns the list of configured 1-wire devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireDevices() As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered Get1WireDevices() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblDevices")

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

      MyDT.Dispose()
      MyDS.Dispose()
      MyDbCommand.Dispose()

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "Get1WireDevices()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Returns the 1-Wire Device
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireDevice(ByVal device_id As String) As OneWireDevice

    Dim OneWireDevice As New OneWireDevice

    Try
      WriteMessage("Entered Get1WireDevice() function.", MessageType.Debug)

      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, device_name, device_type, device_conn, device_addr FROM tblDevices WHERE device_id={0}", device_id)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          If dtrResults.Read() Then
            OneWireDevice.device_id = dtrResults("device_id")
            OneWireDevice.device_name = dtrResults("device_name")
            OneWireDevice.device_type = dtrResults("device_type")
            OneWireDevice.device_conn = dtrResults("device_conn")
            OneWireDevice.device_addr = dtrResults("device_addr")
          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "Get1WireDevice()")
    End Try

    Return OneWireDevice

  End Function

  ''' <summary>
  ''' Inserts a new 1Wire Device into the database
  ''' </summary>
  ''' <param name="device_serial"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_type"></param>
  ''' <param name="device_image"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Insert1WireDevice(ByVal device_serial As String,
                                    ByVal device_name As String,
                                    ByVal device_type As String,
                                    ByVal device_image As String,
                                    ByVal device_conn As String,
                                    ByVal device_addr As String) As Boolean

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_name.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to insert new TEMP08 Device into the database.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("The IP address and port fields are required.  Unable to insert new TEMP08 Device into the database.")
      End If

      '
      ' Try inserting the 1-Wire Device into one of the 20 available slots
      '
      For device_id As Integer = 1 To 9

        Dim strSQL As String = String.Format("INSERT INTO tblDevices (" _
                                     & " device_id, device_serial, device_name, device_type, device_image, device_conn, device_addr" _
                                     & ") VALUES (" _
                                     & " {0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}'" _
                                     & ")", device_id, device_serial, device_name, device_type, device_image, device_conn, device_addr)

        Dim dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        Try

          SyncLock SyncLockMain
            iRecordsAffected = dbcmd.ExecuteNonQuery()
          End SyncLock

        Catch pEx As Exception
          '
          ' Ignore this error
          '
        Finally
          dbcmd.Dispose()
        End Try

        If iRecordsAffected > 0 Then
          Return True
        End If

      Next

      Throw New Exception("Unable to insert 1-Wire Device into the database.  Please ensure you are not attempting to connect more than 9 1-Wire Devices to the plug-in.")

    Catch pEx As Exception
      Call ProcessError(pEx, "Insert1WireDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  '''   Updates 1Wire Device in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="device_serial"></param>
  ''' <param name="device_name"></param>
  ''' <param name="device_type"></param>
  ''' <param name="device_conn"></param>
  ''' <param name="device_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Update1WireDevice(ByVal device_id As Integer,
                                    ByVal device_serial As String,
                                    ByVal device_name As String,
                                    ByVal device_type As String,
                                    ByVal device_conn As String,
                                    ByVal device_addr As String) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      If device_name.Length = 0 Then
        Throw New Exception("One or more required fields are empty.  Unable to update the 1Wire Device in the database.")
      ElseIf device_conn = "Ethernet" And (device_addr.Length = 0) Then
        Throw New Exception("One or more required fields are empty.  Unable to update the 1Wire Device in the database.")
      End If

      Dim strSql As String = String.Format("UPDATE tblDevices SET " _
                                          & " device_serial='{0}', " _
                                          & " device_name='{1}', " _
                                          & " device_type='{2}'," _
                                          & " device_conn='{3}'," _
                                          & " device_addr='{4}' " _
                                          & "WHERE device_id={5}", device_serial, device_name, device_type, device_conn, device_addr, device_id.ToString)

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSql

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "Update1WireDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "Update1WireDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Removes existing 1-Wire Device stored in the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Delete1WireDevice(ByVal device_id As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = String.Format("DELETE FROM tblDevices WHERE device_id={0}", device_id.ToString)

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "Delete1WireDevice() removed " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Warning)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        'Delete1WirePulseCounters(device_id)
        'Delete1WireTempSensors(device_id)
        Return True
      Else
        Return False
      End If

      Return True

    Catch pEx As Exception
      Call ProcessError(pEx, "Delete1WireDevice()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Returns the OneWire Sensor Config from the database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <param name="sensorAddr"></param>
  ''' <param name="sensorRomId"></param>
  ''' <param name="sensorType"></param>
  ''' <param name="sensorChannel"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireSensor(device_id As Integer, sensorAddr As String, sensorRomId As String, sensorType As String, sensorChannel As String) As OneWireSensorConfig

    Dim OneWireSensor As New OneWireSensor(0, device_id, sensorAddr, sensorType, sensorChannel)
    Dim strMessage As String = ""

    strMessage = "Entered Get1WireSensor() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim sensorUnits As String = ""
      Select Case sensorType
        Case "Switch" : sensorUnits = ""
        Case "Temperature" : sensorUnits = "º"
        Case "Counter" : sensorUnits = ""
        Case "Pressure" : sensorUnits = IIf(gUnitType = "Metric", "mb", "inHg")
        Case "Light" : sensorUnits = "Lux"
        Case "Humidity" : sensorUnits = "%RH"
        Case "Voltage" : sensorUnits = "V"
      End Select

      Dim sensor_addr As String = String.Format("{0}:{1}:{2}", sensorRomId, sensorType.Substring(0, 1), sensorChannel).ToUpper

      Dim strSQL As String = String.Format("INSERT OR IGNORE INTO tblSensorConfig (" &
                                           "device_id, sensor_addr, sensor_name, sensor_type, sensor_subtype, sensor_channel, sensor_units" &
                                           ") VALUES (" &
                                           "{0}, '{1}', '{2}', '{3}', '{4}', '{5}', '{6}'); SELECT * FROM tblSensorConfig WHERE sensor_addr = '{1}' ", device_id, sensor_addr, OneWireSensor.sensorName, sensorType, sensorType, sensorChannel, sensorUnits)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          If dtrResults.Read() Then

            OneWireSensor.sensorId = dtrResults("sensor_id")
            OneWireSensor.sensorName = dtrResults("sensor_name")
            OneWireSensor.sensorSubtype = dtrResults("sensor_subtype")
            OneWireSensor.sensorColor = dtrResults("sensor_color")
            OneWireSensor.sensorImage = dtrResults("sensor_image")
            OneWireSensor.sensorUnits = dtrResults("sensor_units")
            OneWireSensor.sensorResolution = dtrResults("sensor_resolution")
            OneWireSensor.postEnabled = dtrResults("post_enabled")

            OneWireSensor.dev_00d = dtrResults("dev_00d")
            OneWireSensor.dev_01d = dtrResults("dev_01d")
            OneWireSensor.dev_07d = dtrResults("dev_07d")
            OneWireSensor.dev_30d = dtrResults("dev_30d")

          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

      Try
        If OneWireSensor.sensorColor.Length = 0 Then
          OneWireSensor.sensorColor = GetSafeColors.Dequeue
          Update1WireSensor(OneWireSensor)
        End If
      Catch ex As Exception

      End Try

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "Get1WireSensor()")
    End Try

    Return OneWireSensor

  End Function

  ''' <summary>
  ''' Updates the 1-Wire Sensor Object
  ''' </summary>
  ''' <param name="OneWireSensor"></param>
  ''' <remarks></remarks>
  Friend Function Update1WireSensor(ByRef OneWireSensor As OneWireSensor) As Boolean

    Dim strMessage As String = ""

    Try

      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database transaction because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim sensorUnits As String = OneWireSensor.sensorUnits
      Select Case OneWireSensor.sensorType
        Case "Switch" : sensorUnits = ""
        Case "Temperature" : sensorUnits = "º"
        Case "Counter"
          Select Case OneWireSensor.sensorSubtype
            Case "Lightning Counter"
              sensorUnits = "Strikes"
            Case "Rain Gauge"
              sensorUnits = IIf(gUnitType = "Metric", "Millimeters ", "Inches")
            Case "Water Meter"
              sensorUnits = IIf(gUnitType = "Metric", "Liters", "Gallons")
            Case "Gas Meter"
              sensorUnits = IIf(gUnitType = "Metric", "m3 ", "ft3")
          End Select
        Case "Pressure" : sensorUnits = IIf(gUnitType = "Metric", "mb", "inHg")
        Case "Light" : sensorUnits = "Lux"
        Case "Humidity"
          Select Case OneWireSensor.sensorSubtype
            Case "Light Level"
              sensorUnits = "Lux"
            Case Else
              sensorUnits = "%RH"
          End Select
        Case "Voltage" : sensorUnits = "V"
      End Select
      OneWireSensor.sensorUnits = sensorUnits

      Dim strSql As String = String.Format("UPDATE tblSensorConfig SET " _
                                          & " sensor_subtype='{0}', " _
                                          & " sensor_color='{1}'," _
                                          & " sensor_image='{2}'," _
                                          & " sensor_units='{3}', " _
                                          & " sensor_resolution={4}," _
                                          & " post_enabled={5}," _
                                          & " dev_enabled=1," _
                                          & " dev_00d={6}," _
                                          & " dev_01d={7}," _
                                          & " dev_07d={8}," _
                                          & " dev_30d={9} " _
                                          & "WHERE sensor_addr='{10}'",
                                          OneWireSensor.sensorSubtype,
                                          OneWireSensor.sensorColor,
                                          OneWireSensor.sensorImage,
                                          OneWireSensor.sensorUnits,
                                          OneWireSensor.sensorResolution,
                                          OneWireSensor.postEnabled,
                                          OneWireSensor.dev_00d,
                                          OneWireSensor.dev_01d,
                                          OneWireSensor.dev_30d,
                                          OneWireSensor.postEnabled,
                                          OneWireSensor.sensorAddr)

      '
      ' Build the insert/update/delete query
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSql

      Dim iRecordsAffected As Integer = 0
      SyncLock SyncLockMain
        iRecordsAffected = MyDbCommand.ExecuteNonQuery()
      End SyncLock

      strMessage = "Update1WireDevice() updated " & iRecordsAffected & " row(s)."
      Call WriteMessage(strMessage, MessageType.Debug)

      MyDbCommand.Dispose()

      If iRecordsAffected > 0 Then
        Return True
      Else
        Return False
      End If

    Catch pEx As Exception
      Call ProcessError(pEx, "Update1WireSensor()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Updates the 1-Wire Sensor Object
  ''' </summary>
  ''' <param name="DeviceId"></param>
  ''' <param name="SensorClass"></param>
  ''' <param name="SensorType"></param>
  ''' <param name="SensorChannel"></param>
  ''' <param name="ROMId"></param>
  ''' <param name="Value"></param>
  ''' <remarks></remarks>
  Public Sub UpdateOneWireSensor(ByVal DeviceId As Integer,
                                 ByVal SensorClass As String,
                                 ByVal SensorType As String,
                                 ByVal SensorChannel As Char,
                                 ByVal ROMId As String,
                                 ByVal Value As Double)

    Dim sensorAddr As String = String.Format("{0}:{1}:{2}", ROMId, SensorType.Substring(0, 1), SensorChannel).ToUpper

    If OneWireSensors.Any(Function(s) s.sensorAddr = sensorAddr) = False Then
      Dim OneWireSensor As OneWireSensor = Get1WireSensor(DeviceId, sensorAddr, ROMId, SensorType, SensorChannel)
      OneWireSensor.sensorClass = SensorClass
      OneWireSensor.Value = Value

      OneWireSensors.Add(OneWireSensor)

      Call WriteMessage(String.Format("Updating {0} device [{1}], channel {2} to {3}.", OneWireSensor.sensorType, OneWireSensor.romId, SensorChannel, Value.ToString), MessageType.Debug)
    Else

      Dim OneWireSensor As OneWireSensor = OneWireSensors.Find(Function(s) s.sensorAddr = sensorAddr)
      OneWireSensor.Value = Value

      Call WriteMessage(String.Format("Updating {0} device [{1}], channel {2} to {3}.", OneWireSensor.sensorType, OneWireSensor.romId, SensorChannel, Value.ToString), MessageType.Debug)
    End If

  End Sub

  ''' <summary>
  ''' Gets the 1-Wire Sensors from the underlying database
  ''' </summary>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetOneWireSensors(ByVal sensor_type As String, Optional device_id As Integer = 0) As DataTable

    Dim ResultsDT As New DataTable
    Dim strMessage As String = ""

    strMessage = "Entered GetOneWireSensors() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblSensorConfig where sensor_type='{0}' AND dev_enabled=1", sensor_type)

      '
      ' Initialize the command object
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      '
      ' Get our DataTable
      '
      Dim MyDT As DataTable = MyDS.Tables(0)

      '
      ' Get record count
      '
      Dim iRecordCount As Integer = MyDT.Rows.Count

      If iRecordCount > 0 Then
        '
        ' Build field names
        '
        Dim iFieldCount As Integer = MyDS.Tables(0).Columns.Count() - 1
        For iFieldNum As Integer = 0 To iFieldCount
          '
          ' Create the columns
          '
          Dim ColumnName As String = MyDT.Columns.Item(iFieldNum).ColumnName
          Dim MyDataColumn As New DataColumn(ColumnName, GetType(String))

          '
          ' Add the columns to the DataTable's Columns collection
          '
          ResultsDT.Columns.Add(MyDataColumn)
        Next

        '
        ' Let's output our records	
        '
        Dim i As Integer = 0
        For i = 0 To iRecordCount - 1
          '
          ' Create the rows
          '
          Dim dr As DataRow
          dr = ResultsDT.NewRow()
          For iFieldNum As Integer = 0 To iFieldCount
            dr(iFieldNum) = MyDT.Rows(i)(iFieldNum)
          Next
          ResultsDT.Rows.Add(dr)
        Next

      End If

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetOneWireSensors()")

    End Try

    Return ResultsDT

  End Function

  ''' <summary>
  ''' Returns the OneWire Sensor Config from the database
  ''' </summary>
  ''' <param name="sensor_addr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireSensor(ByVal sensor_addr As String) As OneWireSensor

    Dim OneWireSensor As New OneWireSensor
    Dim strMessage As String = ""

    strMessage = "Entered Get1WireSensor() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Make sure the datbase is open before attempting to use it
      '
      Select Case DBConnectionMain.State
        Case ConnectionState.Broken, ConnectionState.Closed
          strMessage = "Unable to complete database query because the database " _
                     & "connection has not been initialized."
          Throw New System.Exception(strMessage)
      End Select

      Dim strSQL As String = String.Format("SELECT * FROM tblSensorConfig where sensor_addr='{0}' AND dev_enabled=1", sensor_addr)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

        MyDbCommand.Connection = DBConnectionMain
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        SyncLock SyncLockMain
          Dim dtrResults As IDataReader = MyDbCommand.ExecuteReader()

          If dtrResults.Read() Then
            OneWireSensor = New OneWireSensor(dtrResults("sensor_id"), dtrResults("device_id"), dtrResults("sensor_addr"), dtrResults("sensor_type"), dtrResults("sensor_channel"))

            OneWireSensor.sensorName = dtrResults("sensor_name")
            OneWireSensor.sensorSubtype = dtrResults("sensor_subtype")
            OneWireSensor.sensorColor = dtrResults("sensor_color")
            OneWireSensor.sensorImage = dtrResults("sensor_image")
            OneWireSensor.sensorUnits = dtrResults("sensor_units")
            OneWireSensor.sensorResolution = dtrResults("sensor_resolution")
            OneWireSensor.postEnabled = dtrResults("post_enabled")

            OneWireSensor.dev_00d = dtrResults("dev_00d")
            OneWireSensor.dev_01d = dtrResults("dev_01d")
            OneWireSensor.dev_07d = dtrResults("dev_07d")
            OneWireSensor.dev_30d = dtrResults("dev_30d")

          End If

          dtrResults.Close()
        End SyncLock

        MyDbCommand.Dispose()

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "Get1WireSensor()")
    End Try

    Return OneWireSensor

  End Function

  ''' <summary>
  ''' Sends UDP packet to discover EDS network adapters
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Discover1WireDevices()

    Try
      '
      ' Initialize the HA7Nets
      '
      Dim bHA7NetEnabled As Boolean = CBool(hs.GetINISetting("Interface", "HA7Net", "True", gINIFile))
      Call DiscoverHA7Nets(bHA7NetEnabled)

      '
      ' Initialize the OWServers
      '
      Dim bOWServerEnabled As Integer = CBool(hs.GetINISetting("Interface", "OWServer", "True", gINIFile))
      Call DiscoverOWServers(bOWServerEnabled)

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      ProcessError(pEx, "Discover1WireDevices()")
    End Try

  End Sub

  ''' <summary>
  ''' Refresh 1Wire devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Refresh1WireDevices()

    Dim OneWireDevices As New ArrayList

    Try

      OneWireDevices = Get1WireDeviceList()
      For Each OneWireDevice As OneWireDevice In OneWireDevices

        Select Case OneWireDevice.device_type
          Case "TEMP08"
            '
            ' Check to see if we need to add the device
            '
            If TEMP08Interfaces.Any(Function(TEMP08) TEMP08.InterfaceID = OneWireDevice.device_addr) = False Then
              Dim TEMP08 As New TEMP08(OneWireDevice.device_id, OneWireDevice.device_type, OneWireDevice.device_conn, OneWireDevice.device_addr)
              Dim strResult As String = TEMP08.ConnectToDevice()

              TEMP08Interfaces.Add(TEMP08)

              WriteMessage(String.Format("Addding New Device: {0}->{1}->{2}->{3}->{4}",
                                         OneWireDevice.device_id,
                                         OneWireDevice.device_name,
                                         OneWireDevice.device_type,
                                         OneWireDevice.device_conn,
                                         OneWireDevice.device_addr), MessageType.Debug)

            End If

          Case "HA7E"
            '
            ' Check to see if we need to add the device
            '
            If HA7EInterfaces.Any(Function(HA7E) HA7E.InterfaceID = OneWireDevice.device_addr) = False Then
              Dim ha7e As New HA7E(OneWireDevice.device_id, OneWireDevice.device_addr)
              Dim strResult As String = HS_HA7E.InitHA7E(ha7e)

              HA7EInterfaces.Add(ha7e)
            End If

          Case "HA7Net"
            '
            ' Check to see if we need to add the device
            '
            If HA7NetInterfaces.Any(Function(HA7NetServer) HA7NetServer.InterfaceID = OneWireDevice.device_addr) = False Then

              Dim regexPattern As String = "(?<ipaddr>.+):(?<port>\d+)"
              If Regex.IsMatch(OneWireDevice.device_addr, regexPattern) = True Then

                Dim ip_addr As String = Regex.Match(OneWireDevice.device_addr, regexPattern).Groups("ipaddr").ToString()
                Dim ip_port As String = Regex.Match(OneWireDevice.device_addr, regexPattern).Groups("port").ToString()

                Dim HA7NetServer As New HA7NetServer(OneWireDevice.device_id, ip_addr, OneWireDevice.device_addr)
                HA7NetServer.HTTPPort = ip_port
                HA7NetServer.HTTPSPort = ip_port
                HA7NetServer.UseSSL = False

                HA7NetInterfaces.Add(HA7NetServer)
              End If

            End If

          Case "OWServer", "WirelessController"
            '
            ' Check to see if we need to add the device
            '
            If OWServerInterfaces.Any(Function(OWServer) OWServer.InterfaceID = OneWireDevice.device_addr) = False Then

              Dim regexPattern As String = "(?<ipaddr>.+):(?<port>\d+)"
              If Regex.IsMatch(OneWireDevice.device_addr, regexPattern) = True Then

                Dim ip_addr As String = Regex.Match(OneWireDevice.device_addr, regexPattern).Groups("ipaddr").ToString()
                Dim ip_port As String = Regex.Match(OneWireDevice.device_addr, regexPattern).Groups("port").ToString()

                Dim OWServer As New OWServer(OneWireDevice.device_id, ip_addr, OneWireDevice.device_addr, OneWireDevice.device_type)
                OWServer.HTTPPort = ip_port
                OWServer.HTTPSPort = ip_port
                OWServer.UseSSL = False

                OWServerInterfaces.Add(OWServer)
              End If

            End If

        End Select

      Next

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "Refresh1WireDevices()")
    End Try

  End Sub

  ''' <summary>
  ''' heck One Wire Sensors Thread
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub CheckOneWireSensors()

    Dim bAbortThread As Boolean = False
    Dim iCheckInterval As Integer = 0

    Try
      '
      ' Begin Main While Loop
      '
      While bAbortThread = False

        If gMonitoring = True Then

          Refresh1WireDevices()
          Discover1WireDevices()

          For Each HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer In HA7NetInterfaces
            CheckHA7NetDevices(HA7Net)
          Next

          For Each OWServer As HSPI_ULTRA1WIRE3.OWServer In OWServerInterfaces
            CheckOWServerDevices(OWServer)
          Next

          For Each HA7E As HSPI_ULTRA1WIRE3.HA7E In HA7EInterfaces
            CheckHA7EDevices(HA7E)
          Next

          For Each TEMP08 As HSPI_ULTRA1WIRE3.TEMP08 In TEMP08Interfaces
            TEMP08.CheckOneWireSensors()
          Next

          CheckDatabase()
          UpdateOneWireSensors()
          UpdateConnectedSensors()

        End If

        '
        ' Sleep the requested number of minutes between runs
        '
        iCheckInterval = CInt(hs.GetINISetting("Sensor", "CheckInterval", "1", gINIFile))
        SleepMinutes(iCheckInterval)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("CheckHA7NetThread received abort request, terminating normally."), MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckHA7NetThread()")
    End Try

  End Sub

  ''' <summary>
  ''' Updates the environmental sensor devices
  ''' </summary>
  ''' <remarks></remarks>
  Friend Sub UpdateOneWireSensors()

    Dim strMessage As String = ""
    Dim UpdateInterval As Date = Date.Now
    Dim iCheckInterval As Integer = 0
    Dim bAlwaysInsert As Boolean = True

    strMessage = "Entered UpdateEnviromentalSensors() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Get the polling interval
      '
      iCheckInterval = CInt(hs.GetINISetting("Sensor", "CheckInterval", "1", gINIFile))
      bAlwaysInsert = CBool(hs.GetINISetting("Sensor", "AlwaysInsert", "1", gINIFile))

      Dim ts As Long = ConvertDateTimeToEpoch(DateTime.Now)

      For Each OneWireSensor In OneWireSensors
        Dim dv_addr As String = OneWireSensor.dvAddress

        Dim sensorId As Integer = OneWireSensor.sensorId
        Dim sensorClass As String = OneWireSensor.sensorClass

        '
        ' Determine if we need to create the HomeSeer Device
        '
        If dv_addr.Length = 0 Then
          dv_addr = GetHomeSeerDevice(OneWireSensor)
          OneWireSensor.dvAddress = dv_addr
        End If

        If dv_addr.Length > 0 Then
          Select Case OneWireSensor.sensorType
            Case "Counter"
              '
              ' Process Counter Sensor Type
              '
              Dim sensorSubtype As String = OneWireSensor.sensorSubtype

              '
              ' Process Counter Resolution
              '
              Dim sensorValue As Single = OneWireSensor.Value
              Dim sensorReset As Integer = 0
              Dim strSensorUnits As String = OneWireSensor.sensorUnits
              Dim sensorResolution As Single = OneWireSensor.sensorResolution

              ''
              '' Determine if counter has been reset
              ''
              'Dim strSensorReset As String = hs.GetINISetting(ioMisc, "SensorReset", 0, gINIFile)
              'Try
              '  iSensorReset = Integer.Parse(strSensorReset)
              '  If iSensorValue < iSensorReset Then
              '    hs.SaveINISetting(ioMisc, "SensorReset", "0", gINIFile)
              '    iSensorReset = 0
              '  End If
              'Catch pEx As Exception
              '  Call ProcessError(pEx, "Update1WireCounters()")
              'End Try

              sensorValue -= sensorReset
              sensorValue *= sensorResolution

              '
              ' Write Debug Message
              '
              strMessage = String.Format("Updating HomeSeer device {0} with raw={1}, resolution={2}, reset={3}, value={4}.", OneWireSensor.dvAddress, OneWireSensor.Value, sensorResolution.ToString, sensorReset.ToString, sensorValue.ToString)
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Update Counter Device
              '
              SetDeviceValue(OneWireSensor, sensorValue)

              '
              ' Insert into database
              '
              InsertSensorData("tblCounterData", ts, OneWireSensor)

            Case "Humidity"
              '
              ' Process Humidity Sensor Type
              '
              Dim sensorSubtype As String = OneWireSensor.sensorSubtype
              Dim sensorValue As Integer = OneWireSensor.Value

              Select Case sensorSubtype
                Case "Light Level"
                  '
                  ' Update Light Level
                  '
                  sensorValue += 25
                  sensorValue /= 2
                  If sensorValue < 0 Then sensorValue = 0
                  If sensorValue > 100 Then sensorValue = 0

                  '
                  ' Update Light Level Device
                  '
                  SetDeviceValue(OneWireSensor, sensorValue)

                Case Else
                  '
                  ' Update Humidity device
                  '
                  SetDeviceValue(OneWireSensor, OneWireSensor.Value)

                  '
                  ' Insert into database
                  '
                  InsertSensorData("tblHumidityData", ts, OneWireSensor)
              End Select

              '
              ' Write Debug Message
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

            Case "Light"
              '
              ' Process Light Sensor Type
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Update Light Device
              '
              SetDeviceValue(OneWireSensor, OneWireSensor.Value)

            Case "Pressure"
              '
              ' Process Pressure Sensor Type
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Update Pressure Device
              '
              SetDeviceValue(OneWireSensor, OneWireSensor.Value)

              '
              ' Insert into database
              '
              InsertSensorData("tblPressureData", ts, OneWireSensor)

            Case "Switch"
              '
              ' Process Switch Sensor Type
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Format Switch
              '
              SetDeviceValue(OneWireSensor, OneWireSensor.Value)

            Case "Vibration"
              '
              ' Process Vibration Sensor Type
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Format Vibration
              '
              SetDeviceValue(OneWireSensor, OneWireSensor.Value)

            Case "Temperature"
              '
              ' Process Temperature Sensor Type
              '
              If OneWireSensor.Value = -185 Then
                strMessage = String.Format("HomeSeer device {0} did not report a valid value; database insert skipped.", dv_addr, CLng(OneWireSensor.Value))
                WriteMessage(strMessage, MessageType.Warning)
              ElseIf DateDiff(DateInterval.Minute, OneWireSensor.ValueTs, Date.Now) > CLng(iCheckInterval) Then
                '
                ' This device hasn't reported a temperature
                '
                If bAlwaysInsert = True Then
                  strMessage = String.Format("HomeSeer device {0} did not report a value; using last reported temperature.", dv_addr, CLng(OneWireSensor.Value))
                  WriteMessage(strMessage, MessageType.Notice)
                  '
                  ' Insert into database
                  '
                  InsertSensorData("tblTemperatureData", ts, OneWireSensor)
                Else
                  strMessage = String.Format("HomeSeer device {0} did not report a value.", dv_addr)
                  WriteMessage(strMessage, MessageType.Notice)
                End If
              Else
                '
                ' Update the sensor value
                '
                strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
                WriteMessage(strMessage, MessageType.Debug)

                '
                ' Update Temperature Device
                '
                SetDeviceValue(OneWireSensor, OneWireSensor.Value)

                '
                ' Insert into database
                '
                InsertSensorData("tblTemperatureData", ts, OneWireSensor)
              End If

            Case "Voltage"
              '
              ' Process Voltage Sensor Type
              '
              strMessage = String.Format("Updating HomeSeer device {0} with {1}.", dv_addr, CLng(OneWireSensor.Value))
              WriteMessage(strMessage, MessageType.Debug)

              '
              ' Update Voltage
              '
              SetDeviceValue(OneWireSensor, OneWireSensor.Value)

            Case Else
              '
              ' Process Unsupported Sensor Type
              '
              strMessage = String.Format("Unable to update {0}; unsupported device ", OneWireSensor.sensorType)
              WriteMessage(strMessage, MessageType.Notice)

          End Select

        Else
          '
          ' Unable to find the HomeSeer device address assoicated to the ROMId
          '
          strMessage = String.Format("Unable to update {0}; unable to determine HomeSeer address.", OneWireSensor.romId)
          WriteMessage(strMessage, MessageType.Notice)
        End If

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateEnvironmentalSensors()")
    End Try

  End Sub

  ''' <summary>
  ''' Inserts sensor reading into the database
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="OneWireSensor"></param>
  ''' <remarks></remarks>
  Private Sub InsertSensorData(dbTableName As String, ts As Long, ByRef OneWireSensor As OneWireSensor)

    Try
      '
      ' Insert value into database
      '
      Dim strSQL As String = String.Format("INSERT INTO {0} (" & _
                                           " device_id, sensor_id, ts, value " & _
                                           ") VALUES (" & _
                                           " {1}, {2}, {3}, {4}" & _
                                           ")", dbTableName, OneWireSensor.deviceId, OneWireSensor.sensorId, ts, OneWireSensor.DBValue)

      SyncLock (DatabaseInsertQueue.SyncRoot)
        DatabaseInsertQueue.Enqueue(strSQL)
      End SyncLock

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "InsertSensorData()")
    End Try

  End Sub

  ''' <summary>
  ''' Gets the 1-Wire Sensor History
  ''' </summary>
  ''' <param name="ts_start"></param>
  ''' <param name="ts_end"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSensorHistory(dbTable As String, ByVal ts_start As Long, ByVal ts_end As Long) As DataSet

    Try
      '
      ' Define the SQL Query
      '
      Dim strSQL As String = String.Format("SELECT device_id, channel_id, AVG(watt) as wattAvg, SUM(kwh) as kwhSum FROM tblSensorHistory WHERE ts >= {0} and ts <= {1} GROUP BY channel_id", ts_start.ToString, ts_end.ToString)

      '
      ' Execute the data reader
      '
      Dim MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

      MyDbCommand.Connection = DBConnectionMain
      MyDbCommand.CommandType = CommandType.Text
      MyDbCommand.CommandText = strSQL

      '
      ' Initialize the dataset, then populate it
      '
      Dim MyDS As DataSet = New DataSet

      Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
      MyDA.SelectCommand = MyDbCommand

      SyncLock SyncLockMain
        MyDA.Fill(MyDS)
      End SyncLock

      MyDbCommand.Dispose()

      Return MyDS

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetChannelHistory()")
      Return New DataSet
    End Try

  End Function

  ''' <summary>
  ''' Flushes the sensor list
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub FlushSensorList()

    Try

      OneWireSensors.Clear()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "FlushSensorList()")
    End Try

  End Sub

#End Region

#Region "Database Threads"

  ''' <summary>
  ''' Process Sensor History Data
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessSensorHistory()

    Dim bAbortThread As Boolean = False

    Try
      '
      ' Give up some time to allow plug-in to start
      '
      Thread.Sleep(1000 * 15)

      While bAbortThread = False

        'Try
        '  '
        '  ' Update watts and kwh history
        '  '
        '  Dim days() As Integer = {0, 1, 7, 30}
        '  For Each iDay As Integer In days

        '    Dim ts_start As Long = 0
        '    Dim ts_end As Long = 0
        '    Select Case iDay
        '      Case 0
        '        ts_start = ConvertDateTimeToEpoch(DateTime.Today)
        '        ts_end = ConvertDateTimeToEpoch(DateTime.Now)
        '      Case 1
        '        ts_start = ConvertDateTimeToEpoch(DateAdd(DateInterval.Day, -iDay, DateTime.Today))
        '        ts_end = ConvertDateTimeToEpoch(DateTime.Today)
        '      Case Else
        '        ts_start = ConvertDateTimeToEpoch(DateAdd(DateInterval.Day, -iDay, DateTime.Now))
        '        ts_end = ConvertDateTimeToEpoch(DateTime.Now)
        '    End Select

        '    Using MyDS As DataSet = GetSensorHistory(ts_start, ts_end)

        '      For Each DataRow As DataRow In MyDS.Tables(0).Rows

        '        Dim channel_id As Integer = Convert.ToInt32(DataRow("channel_id"))
        '        Dim strSQL As String = String.Format("UPDATE tblChannelConfig SET watt_{0}d={1}, kwh_{0}d={2} WHERE channel_id={3}", _
        '                                              iDay.ToString.PadLeft(2, "0"), _
        '                                              DataRow("wattAvg"), _
        '                                              DataRow("kwhSum"), _
        '                                              channel_id.ToString)

        '        Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        '          dbcmd.Connection = DBConnectionMain
        '          dbcmd.CommandType = CommandType.Text
        '          dbcmd.CommandText = strSQL

        '          SyncLock SyncLockMain
        '            dbcmd.ExecuteNonQuery()
        '          End SyncLock

        '          dbcmd.Dispose()

        '        End Using

        '        '
        '        ' Check HomeSeer Watt Devices
        '        '
        '        Dim strDbFieldWatt As String = String.Format("dev_watt_{0}d", iDay.ToString.PadLeft(2, "0"))
        '        UpdateChannelDevice(channel_id, strDbFieldWatt, DataRow("wattAvg"))

        '        '
        '        ' Check HomeSeer KWh Devices
        '        '
        '        Dim strDbFieldKWh As String = String.Format("dev_kwh_{0}d", iDay.ToString.PadLeft(2, "0"))
        '        UpdateChannelDevice(channel_id, strDbFieldKWh, DataRow("kwhSum"))

        '      Next

        '    End Using

        '  Next

        'Catch pEx As Exception
        '  '
        '  ' Return message
        '  '
        '  ProcessError(pEx, "ProcessSensorHistory()")
        'End Try

        '
        ' Give up some time to allow the main thread to populate the queue with more data
        '
        SleepMinutes(60)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessSensorHistory thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessSensorHistory()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessSensorHistory terminated."), MessageType.Debug)

    End Try

  End Sub

  ''' <summary>
  ''' Processes Database queue
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessDatabaseQueue()

    Dim bAbortThread As Boolean = False

    Try

      While bAbortThread = False

        Try
          '
          ' Check the Database Insert Queue
          '
          If DatabaseInsertQueue.Count > 0 Then
            '
            ' Process commands in command queue
            '
            Call CheckDatabase()

            While DatabaseInsertQueue.Count > 0 And gIOEnabled = True

              Dim strSQL As String = ""
              SyncLock (DatabaseInsertQueue.SyncRoot)
                strSQL = DatabaseInsertQueue.Dequeue
              End SyncLock

              If strSQL.Length > 0 Then
                '
                ' Insert into database
                '
                Call hspi_database.InsertData(strSQL)

                Thread.Sleep(0)

              End If

            End While

          End If

          '
          ' Check queue counts
          '
          If DatabaseInsertQueue.Count > MAX_DATABASE_QUEUE Then
            Dim strMessage As String = "DatabaseInsertQueue count exceeded threshold!  Clearing queue to prevent memory issues."
            Call WriteMessage(strMessage, MessageType.Critical)

            SyncLock DatabaseInsertQueue.SyncRoot
              DatabaseInsertQueue.Clear()
            End SyncLock
          End If

        Catch pEx As Exception

        End Try

        '
        ' Give up some time to allow the main thread to populate the queue with more data
        '
        Thread.Sleep(1000)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessDatabaseQueue thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessDatabaseQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessDatabaseQueue terminated."), MessageType.Debug)

    End Try

  End Sub

#End Region

#Region "Helper Functions"

  '------------------------------------------------------------------------------------
  'Purpose:   Determines if IP Address is valid
  'Input:     ByVal strIPAddr As String
  'Output:    Boolean
  '------------------------------------------------------------------------------------
  Private Function IsIPAddr(ByVal strIPAddr As String) As Boolean

    Try
      ' Create an instance of IPAddress for the specified address string 
      ' (in dotted-quad, or colon-hexadecimal notation).
      Dim address As System.Net.IPAddress = System.Net.IPAddress.Parse(strIPAddr)
      Return True
    Catch pEx As ArgumentNullException
      Return False
    Catch pEx As FormatException
      Return False
    Catch pEx As Exception
      Return False
    End Try

  End Function

#End Region

End Module

Module HS_HA7E

  Dim ds18S20s As New List(Of DS18S20)
  Dim ds18B20s As New List(Of DS18B20)

  ''' <summary>
  ''' Initialize the HA7E
  ''' </summary>
  ''' <param name="HA7E"></param>
  ''' <remarks></remarks>
  Public Function InitHA7E(ByRef HA7E As HSPI_ULTRA1WIRE3.HA7E) As String

    Dim strMessage As String = ""

    strMessage = "Entered InitHA7E() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      HA7E.OW_ResetBus()
      HA7E.OW_SearchROM()
      HA7E.OW_ResetBus()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      strMessage = String.Format("HA7E failed to initialize on {0} due to error {1}.", HA7E.InterfaceID, pEx.Message)
      Return strMessage
    End Try

  End Function

  '------------------------------------------------------------------------------------
  'Purpose:   Process HA7Net connected devices
  'Input:     HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer
  'Output:    None
  '------------------------------------------------------------------------------------
  Sub CheckHA7EDevices(ByRef HA7E As HSPI_ULTRA1WIRE3.HA7E)

    Try
      Dim Sensors As ArrayList

      Dim ds18S20 As DS18S20
      Dim ds18B20 As DS18B20

      Dim SensorClass As String = "Environmental"
      Dim SensorType As String = "Temperature"

      Dim strMessage As String = ""

      '
      ' Indicate we are looking for sensors
      '
      WriteMessage("Getting HA7E devices from:  " & HA7E.InterfaceID, MessageType.Debug)

      Dim OWResponse As HSPI_ULTRA1WIRE3.OW_Response = HA7E.OW_SearchROM()
      Sensors = OWResponse.Data

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)
      WriteMessage(String.Format("HA7E [{0}] has {1} 1-wire devices connected.", HA7E.InterfaceID, Sensors.Count), MessageType.Debug)

      '
      ' Process each sensor
      '
      For Each ROMId As String In Sensors

        ' Check for DS18S20's [high precision digital thermometer]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "10" Then
          Dim SensorKey As String = String.Format("{0}:{1}", ROMId, SensorType)

          If ds18S20s.Any(Function(s) s.ROMId = ROMId) = False Then
            ds18S20 = New DS18S20(ROMId, HA7E)
            ds18S20s.Add(ds18S20)

            strMessage = String.Format("Found new ds18S20 sensor [{0}] ...", ROMId)
            Call WriteMessage(strMessage, MessageType.Debug)
          End If

        End If

        ' Check for DS18B20's [programmable resolution digital thermometer]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "28" Then
          Dim SensorKey As String = String.Format("{0}:{1}", ROMId, SensorType)

          If ds18B20s.Any(Function(s) s.ROMId = ROMId) = False Then
            ds18B20 = New DS18B20(ROMId, HA7E)
            ds18B20s.Add(ds18B20)

            strMessage = String.Format("Found new ds18B20 sensor [{0}] ...", ROMId)
            Call WriteMessage(strMessage, MessageType.Debug)
          End If

        End If

      Next

      Call CheckDS18S20(HA7E.DeviceId)
      Call CheckDS18B20(HA7E.DeviceId)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckHA7EDevices()")

    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Iterates through each of the previously discovered DS18S20 sensors, and 
  '           reads the state of the power supply and gets the current temperature.
  'Input:     None
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub CheckDS18S20(deviceId As Integer)

    Dim strMessage As String = ""
    Dim iAttempts As Integer = CInt(hs.GetINISetting("Sensor", "Attempts", 1, gINIFile))

    strMessage = "Entered CheckDS18S20() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Determine the requested temperature conversion and acceptable range
      '
      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)
      Dim iMinTemp As Integer = -67
      Dim iMaxTemp As Integer = 257
      If bFarenheight = False Then
        iMinTemp = -55
        iMaxTemp = 125
      End If

      '
      ' Test each DS18S20 sensor
      '
      For Each ds18S20 As HSPI_ULTRA1WIRE3.DS18S20 In ds18S20s

        Dim iTries As Integer = 0
        Dim Value As Single

        Do
          '
          ' Increment tries counter
          '
          iTries += 1

          ds18S20.BusMaster.clearLog()
          ds18S20.BusMaster.OW_GetLock()

          '
          ' Checking power supply type...
          '
          If ds18S20.ReadPowerSupply() = 0 Then
            ' Sensor is being parasitically powered.
            strMessage = String.Format("{0} is {1} powered.", ds18S20.ROMId, "parasitically")
            WriteMessage(strMessage, MessageType.Debug)
          Else
            ' Sensor is being externally powered.
            strMessage = String.Format("{0} is {1} powered.", ds18S20.ROMId, "externally")
            WriteMessage(strMessage, MessageType.Debug)
          End If

          '
          ' Attempt to read temperature
          '
          ds18S20.BusMaster.clearLog()

          Value = ds18S20.GetTemperature(bFarenheight)
          ds18S20.BusMaster.OW_ReleaseLock()

          '
          ' Check for bad temperature reading
          '
          If Value <> 185 And (Value > iMinTemp And Value <= iMaxTemp) Then
            Exit Do
          Else
            Thread.Sleep(1000)
          End If
        Loop While iTries < iAttempts

        Select Case Value
          Case 185, 30.875
          Case Else
            UpdateOneWireSensor(deviceId, "Environmental", "Temperature", "A", ds18S20.ROMId, Value)
        End Select

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDS18S20()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Iterates through each of the previously discovered DS18B20 sensors, and 
  '           reads the state of the power supply and gets the current temperature.
  'Input:     None
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub CheckDS18B20(ByVal deviceId As Integer)

    Dim strMessage As String = ""
    Dim iAttempts As Integer = CInt(hs.GetINISetting("Sensor", "Attempts", 1, gINIFile))

    strMessage = "Entered CheckDS18B20() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Determine the requested temperature conversion and acceptable range
      '
      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)
      Dim iMinTemp As Integer = -67
      Dim iMaxTemp As Integer = 257
      If bFarenheight = False Then
        iMinTemp = -55
        iMaxTemp = 125
      End If

      '
      ' Test each DS18B20 sensor
      '
      For Each ds18B20 As HSPI_ULTRA1WIRE3.DS18B20 In ds18B20s

        Dim iTries As Integer = 0
        Dim Value As Single

        Do
          '
          ' Increment tries counter
          '
          iTries += 1

          ds18B20.BusMaster.clearLog()
          ds18B20.BusMaster.OW_GetLock()

          '
          ' Checking power supply type...
          '
          If ds18B20.ReadPowerSupply() = 0 Then
            ' Sensor is being parasitically powered.
            strMessage = String.Format("{0} is {1} powered.", ds18B20.ROMId, "parasitically")
            WriteMessage(strMessage, MessageType.Debug)
          Else
            ' Sensor is being externally powered.
            strMessage = String.Format("{0} is {1} powered.", ds18B20.ROMId, "externally")
            WriteMessage(strMessage, MessageType.Debug)
          End If

          '
          ' Attempt to read temperature
          '
          ds18B20.BusMaster.clearLog()

          Value = ds18B20.GetTemperature(bFarenheight)
          ds18B20.BusMaster.OW_ReleaseLock()

          '
          ' Check for bad temperature reading
          '
          If Value <> 185 And (Value > iMinTemp And Value <= iMaxTemp) Then
            Exit Do
          Else
            Thread.Sleep(1000)
          End If
        Loop While iTries < iAttempts

        Select Case Value
          Case 185, 30.875
          Case Else
            UpdateOneWireSensor(deviceId, "Environmental", "Temperature", "A", ds18B20.ROMId, Value)
        End Select

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDS18B20()")
    End Try

  End Sub

End Module

Module HS_HA7Net

  '------------------------------------------------------------------------------------
  'Purpose:   Subroutine to discover HA7Net interfaces
  'Input:     bEnabled As Boolean
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub DiscoverHA7Nets(ByVal bEnabled As Boolean)

    Dim strMessage As String = ""

    strMessage = "Entered InitHA7Net() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Adds an HA7Net owInterface to the Arraylist of owInterfaces 
      '
      Dim discoveredHA7Nets As ArrayList = HSPI_ULTRA1WIRE3.HA7NetServer.DiscoverHA7Nets()

      If discoveredHA7Nets.Count = 0 Then
        WriteMessage("No HA7Net units were found via discovery process ...", MessageType.Debug)
      End If

      For Each tempHA7NetInfo As HSPI_ULTRA1WIRE3.HA7NetServer.DiscoveryResponse In discoveredHA7Nets
        With tempHA7NetInfo

          If HA7NetInterfaces.Any(Function(HA7NetServer) HA7NetServer.HostName = tempHA7NetInfo.IPAddress.ToString) = False Then
            '
            ' Add this item to the database
            '
            Insert1WireDevice(.SerialNumber, .DeviceName, "HA7Net", "", "Ethernet", String.Concat(.IPAddress.ToString, ":", .Port))

            strMessage = String.Format("Found HA7Net [{0}], hostname {1} at IP address {2} on port {3}", .SerialNumber, .DeviceName, .IPAddress.ToString(), .Port)
            WriteMessage(strMessage, MessageType.Informational)
          End If

        End With
      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "DiscoverHA7Nets()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process HA7Net connected devices
  'Input:     HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer
  'Output:    None
  '------------------------------------------------------------------------------------
  Sub CheckHA7NetDevices(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer)

    Try

      Dim Sensors As Hashtable

      Dim DS18S20s As String = "" ' Temperature
      Dim DS18B20s As String = "" ' Temperature
      Dim DS2438s As String = ""  ' Humidity
      Dim DS2406s As String = ""  ' Switch
      Dim DS2423s As String = ""  ' Counter

      Dim strMessage As String = ""

      ' Indicate we are looking for sensors
      WriteMessage("Getting HA7Net devices from:  " & HA7Net.HostName, MessageType.Debug)

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_SearchROM()
      Sensors = OWResponse.Data

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)
      WriteMessage(String.Format("HA7Net [{0}] has {1} 1-wire devices connected.", HA7Net.HostName, Sensors.Count), MessageType.Debug)

      '
      ' Process each sensor
      '
      For Each ROMId As String In Sensors.Keys

        ' Check for DS18S20's [high precision digital thermometer]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "10" Then
          DS18S20s &= String.Format("{0}{1}", IIf(DS18S20s.Length = 0, "", ","), ROMId)
        End If

        ' Check for DS18B20's [programmable resolution digital thermometer]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "28" Then
          DS18B20s &= String.Format("{0}{{{1},{2}}}", IIf(DS18B20s.Length = 0, "", ","), ROMId, "12")
        End If

        ' Check for DS2438's [smart battery monitor]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "26" Then
          DS2438s &= String.Format("{0}{1}", IIf(DS2438s.Length = 0, "", ","), ROMId)
        End If

        ' Check for DS2406's [Dual addressable switch plus memory]
        ' http://192.168.2.5/1Wire/SwitchControl.html?SwitchRequest={2B000000346E3612, {{1, Read, False}}
        'URL: /1Wire/SwitchControl.html?SwitchRequest={Address1, {{Channel1, Action1, Reset}, {Channel2, Action2, Reset}}}
        If Microsoft.VisualBasic.Right(ROMId, 2) = "12" Then
          DS2406s &= String.Format("{0}{{{1},{{{{{2},{3},{4}}},{{{5},{6},{7}}}}}}}", IIf(DS2423s.Length = 0, "", ","), ROMId, "1", "READ", "FALSE", "2", "READ", "FALSE")
        End If

        ' Check for DS2423's [counter]
        If Microsoft.VisualBasic.Right(ROMId, 2) = "1D" Then
          DS2423s &= String.Format("{0}{{{1},{2}}}", IIf(DS2423s.Length = 0, "", ","), ROMId, "A")
          DS2423s &= String.Format("{0}{{{1},{2}}}", IIf(DS2423s.Length = 0, "", ","), ROMId, "B")
        End If

      Next

      If DS18S20s.Length > 0 Then
        ReadDS18S20Sensors(HA7Net, DS18S20s)
      End If

      If DS18B20s.Length > 0 Then
        ReadDS18B20Sensors(HA7Net, DS18B20s)
      End If

      If DS2438s.Length > 0 Then
        ReadDS2438Sensors(HA7Net, DS2438s)
      End If

      If DS2406s.Length > 0 Then
        ReadDS2406Switches(HA7Net, DS2406s)
      End If

      If DS2423s.Length > 0 Then
        ReadDS2423Counters(HA7Net, DS2423s)
      End If

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckHA7NetDevices()")

    End Try

  End Sub

  Private Sub ReadDS18S20Sensors(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer, ByVal Sensors As String)

    Try

      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_ReadTemperature(Sensors)
      Dim OWData As Hashtable = OWResponse.Data
      Dim SensorType As String = "DS18S20"

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)

      For Each ROMId As String In OWData.Keys

        Dim Value As Single = HA7Net.ConvertTemperature(OWData(ROMId)("Temperature"), bFarenheight)
        UpdateOneWireSensor(HA7Net.DeviceId, "Environmental", "Temperature", "A", ROMId, Value)

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReadDS18S20Sensors()")
    End Try

  End Sub

  Private Sub ReadDS18B20Sensors(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer, ByVal Sensors As String)

    Try
      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_ReadDS18B20(Sensors)
      Dim OWData As Hashtable = OWResponse.Data
      Dim SensorType As String = "DS18B20"

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)

      For Each ROMId As String In OWData.Keys

        Dim Value As Single = HA7Net.ConvertTemperature(OWData(ROMId)("Temperature"), bFarenheight)
        UpdateOneWireSensor(HA7Net.DeviceId, "Environmental", "Temperature", "A", ROMId, Value)

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReadDS18B20Sensors()")
    End Try

  End Sub

  Private Sub ReadDS2438Sensors(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer, ByVal Sensors As String)

    Try

      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_ReadHumidity(Sensors)
      Dim OWData As Hashtable = OWResponse.Data

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)

      For Each ROMId As String In OWData.Keys

        Dim Value1 As Single = HA7Net.ConvertTemperature(OWData(ROMId)("Temperature"), bFarenheight)
        UpdateOneWireSensor(HA7Net.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

        Dim Value2 As Single = Math.Abs(Single.Parse(OWData(ROMId)("Humidity"), nfi))
        UpdateOneWireSensor(HA7Net.DeviceId, "Environmental", "Humidity", "A", ROMId, Value2)
      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReadDS2438Sensors()")
    End Try

  End Sub

  Private Sub ReadDS2406Switches(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer, ByVal Switches As String)

    Try

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_ReadSwitch(Switches)
      Dim OWData As Hashtable = OWResponse.Data
      Dim SensorType As String = "Switch"

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)

      For Each ROMId As String In OWData.Keys

        Dim ValueA As Byte = 0
        Dim ValueB As Byte = 0

        If OWData.ContainsKey(ROMId) = True Then
          If OWData(ROMId).ContainsKey("ValueA") = True Then
            ValueA = CByte(OWData(ROMId)("ValueA"))
            UpdateOneWireSensor(HA7Net.DeviceId, "Switch", SensorType, "A", ROMId, ValueA)
          End If
          If OWData(ROMId).ContainsKey("ValueB") = True Then
            ValueB = CByte(OWData(ROMId)("ValueB"))
            UpdateOneWireSensor(HA7Net.DeviceId, "Switch", SensorType, "B", ROMId, ValueB)
          End If
        End If

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReadDS2406Switches()")
    End Try

  End Sub

  Private Sub ReadDS2423Counters(ByRef HA7Net As HSPI_ULTRA1WIRE3.HA7NetServer, ByVal Counters As String)

    Try

      Dim OWResponse As HA7NetServer.OW_Response = HA7Net.OW_ReadCounter(Counters)
      Dim OWData As Hashtable = OWResponse.Data
      Dim SensorType As String = "Counter"

      WriteMessage(String.Format("Exception Code: {0}", OWResponse.Exception_Code), MessageType.Debug)
      WriteMessage(String.Format("Exception Desc: {0}", OWResponse.Exception_Description), MessageType.Debug)
      WriteMessage(String.Format("Response Date:  {0}", OWResponse.ResponseTime), MessageType.Debug)

      For Each ROMId As String In OWData.Keys

        Dim ValueA As Integer = 0
        Dim ValueB As Integer = 0

        If OWData.ContainsKey(ROMId) = True Then
          If OWData(ROMId).ContainsKey("ValueA") = True Then
            ValueA = Integer.Parse(OWData(ROMId)("ValueA"), nfi)
            UpdateOneWireSensor(HA7Net.DeviceId, "Counter", SensorType, "A", ROMId, ValueA)
          End If
          If OWData(ROMId).ContainsKey("ValueB") = True Then
            ValueB = Integer.Parse(OWData(ROMId)("ValueB"), nfi)
            UpdateOneWireSensor(HA7Net.DeviceId, "Counter", SensorType, "B", ROMId, ValueB)
          End If
        End If

      Next

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ReadDS2423Counters()")
    End Try

  End Sub

End Module

Module HS_OWSERVER

  '------------------------------------------------------------------------------------
  'Purpose:   Subroutine to discover OWServer interfaces
  'Input:     bEnabled As Boolean
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub DiscoverOWServers(ByVal bEnabled As Boolean)

    Dim strMessage As String = ""

    strMessage = "Entered InitOWServer() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Adds an OWServer owInterface to the Arraylist of owInterfaces 
      '
      Dim discoveredOWServers As ArrayList = HSPI_ULTRA1WIRE3.OWServer.DiscoverOWServers

      If discoveredOWServers.Count = 0 Then
        strMessage = "No OWServer units were found via discovery process ..."
        WriteMessage(strMessage, MessageType.Debug)
      End If

      For Each tempOWServerInfo As HSPI_ULTRA1WIRE3.OWServer.DiscoveryResponse In discoveredOWServers
        With tempOWServerInfo

          If OWServerInterfaces.Any(Function(OWServer) OWServer.HostName = tempOWServerInfo.IPAddress.ToString) = False Then

            '
            ' Add this item to the database
            '
            Insert1WireDevice(.SerialNumber, .DeviceName, .DeviceType, "", "Ethernet", String.Concat(.IPAddress.ToString, ":", .Port))

            strMessage = String.Format("Found OWServer [{0}], hostname {1} at IP address {2} on port {3}", .SerialNumber, .DeviceName, .IPAddress.ToString(), .Port)
            WriteMessage(strMessage, MessageType.Informational)
          End If

        End With
      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "DiscoverOWServers()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process OWServer connected devices
  'Input:     OWServer As HSPI_ULTRA1WIRE3.OWServer
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub CheckOWServerDevices(ByVal OWServer As HSPI_ULTRA1WIRE3.OWServer)

    Try

      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)
      Dim bMetric As Boolean = IIf(gUnitType = "Metric", True, False)

      Dim BarometricPressureXPath As String = IIf(gUnitType = "Metric", "xsi:BarometricPressureMb", "xsi:BarometricPressureHg")

      Dim xmlDoc As New Xml.XmlDocument()

      Dim strURL As String = String.Format("http://{0}:{1}/{2}", OWServer.HostName, OWServer.HTTPPort, "details.xml")
      WriteMessage("Getting OWServer devices from:  " & strURL, MessageType.Debug)

      '
      ' Get the XML data
      '
      Using ResponseStream As IO.Stream = WebRequest.Create(strURL).GetResponse().GetResponseStream()
        xmlDoc.Load(ResponseStream)
      End Using

      Dim nsMgr As New XmlNamespaceManager(xmlDoc.NameTable)
      If OWServer.InterfaceType = "WirelessController" Then
        nsMgr.AddNamespace("xsi", "http://www.embeddeddatasystems.com/schema/wirelesscontroller")
      Else
        nsMgr.AddNamespace("xsi", "http://www.embeddeddatasystems.com/schema/owserver")
      End If

      '
      ' Get WebURL
      ' 
      Dim XmlDevicesConnected As XmlNode = xmlDoc.SelectSingleNode("//xsi:DevicesConnected", nsMgr)
      If Not XmlDevicesConnected Is Nothing Then
        Try
          WriteMessage(String.Format("OWServer [{0}] has {1} 1-wire devices connected.", OWServer.HostName, XmlDevicesConnected.InnerText), MessageType.Debug)
        Catch ex As Exception

        End Try
      End If

      '<owd_DS18S20 Description="Parasite power thermometer">
      ' <Name>DS18S20</Name>
      ' <Family>10</Family> 
      ' <ROMId>1057479200080078</ROMId> 
      ' <Health>7</Health> 
      ' <RawData>22004B46000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000</RawData> 
      ' <PrimaryValue>17.0 Deg C</PrimaryValue> 
      ' <Temperature Units="Centigrade">17.0</Temperature> 
      ' <UserByte1 Writable="True">75</UserByte1> 
      ' <UserByte2 Writable="True">70</UserByte2> 
      '</owd_DS18S20>

      '
      ' Get the DS18S20
      ' 
      Dim xmlDS18S20s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_DS18S20", nsMgr)
      If xmlDS18S20s.Count > 0 Then

        For Each xmlDS18S20 As XmlNode In xmlDS18S20s

          Dim Family As String = xmlDS18S20.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlDS18S20.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlDS18S20.SelectSingleNode("xsi:Temperature", nsMgr).InnerText

          Dim Value As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value)

        Next
      End If

      '<owd_DS18B20 Description="Programmable resolution thermometer">
      ' <Name>DS18B20</Name> 
      ' <Family>28</Family> 
      ' <ROMId>2815C50701000073</ROMId> 
      ' <Health>7</Health> 
      ' <RawData>A9014B467F0000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000</RawData> 
      ' <PrimaryValue>26.5625 Deg C</PrimaryValue> 
      ' <Temperature Units="Centigrade">26.5625</Temperature> 
      ' <UserByte1 Writable="True">75</UserByte1> 
      ' <UserByte2 Writable="True">70</UserByte2> 
      ' <Resolution>12</Resolution> 
      ' <PowerSource>0</PowerSource> 
      '</owd_DS18B20>

      '
      ' Get the DS18B20
      ' 
      Dim xmlDS18B20s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_DS18B20", nsMgr)
      If xmlDS18B20s.Count > 0 Then
        Dim SensorType As String = "DS18B20"
        For Each xmlDS18B20 As XmlNode In xmlDS18B20s

          Dim Family As String = xmlDS18B20.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlDS18B20.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlDS18B20.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Resolution As String = xmlDS18B20.SelectSingleNode("xsi:Resolution", nsMgr).InnerText

          Dim Value As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value)

        Next
      End If

      '<owd_DS2438 Description="Smart battery monitor">
      ' <Name>DS2438</Name> 
      ' <Family>26</Family> 
      ' <ROMId>2642C54F000000F0</ROMId> 
      ' <Health>7</Health> 
      ' <RawData>B01024010200B0100000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000</RawData> 
      ' <PrimaryValue>16.6875 Deg C</PrimaryValue> 
      ' <Temperature Units="Centigrade">16.6875</Temperature> 
      ' <Vdd Units="Volts">2.920000</Vdd> 
      ' <Vsense Units="Millivolts">42.719997</Vsense> 
      ' <Vad Units="Millivolts">0.488200</Vad> 
      ' <Humidity Units="PercentRelativeHumidity">-24.85</Humidity> 
      '</owd_DS2438>

      '
      ' Get the DS2438
      ' 
      Dim xmlDS2438s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_DS2438", nsMgr)
      If xmlDS2438s.Count > 0 Then
        For Each xmlDS2438 As XmlNode In xmlDS2438s
          Dim Family As String = xmlDS2438.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlDS2438.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlDS2438.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlDS2438.SelectSingleNode("xsi:Humidity", nsMgr).InnerText

          If Humidity.Contains("-") = False Then
            Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)
          End If

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

        Next
      End If

      '
      ' Get the EDS0064 (Temp)
      ' 
      Dim xmlEDS0064s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0064", nsMgr)
      If xmlEDS0064s.Count > 0 Then
        For Each xmlEDS0064 As XmlNode In xmlEDS0064s
          Dim Family As String = xmlEDS0064.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0064.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0064.SelectSingleNode("xsi:Temperature", nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

        Next
      End If

      '
      ' Get the EDS0065 (Temp, humidity)
      ' 
      Dim xmlEDS0065s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0065", nsMgr)
      If xmlEDS0065s.Count > 0 Then
        For Each xmlEDS0065 As XmlNode In xmlEDS0065s
          Dim Family As String = xmlEDS0065.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0065.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0065.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS0065.SelectSingleNode("xsi:Humidity", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

        Next
      End If

      '
      ' Get the EDS0066 (Temp, barometric pressure)
      ' 
      Dim xmlEDS0066s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0066", nsMgr)
      If xmlEDS0066s.Count > 0 Then
        For Each xmlEDS0066 As XmlNode In xmlEDS0066s
          Dim Family As String = xmlEDS0066.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0066.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0066.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim BarometricPressure As String = xmlEDS0066.SelectSingleNode(BarometricPressureXPath, nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

          Dim Value3 As Single = Math.Abs(Single.Parse(BarometricPressure, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Pressure", "A", ROMId, Value3)

        Next
      End If

      '
      ' Get the EDS0066 (Temp, light)
      ' 
      Dim xmlEDS0067s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0067", nsMgr)
      If xmlEDS0067s.Count > 0 Then
        For Each xmlEDS0067 As XmlNode In xmlEDS0067s
          Dim Family As String = xmlEDS0067.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0067.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0067.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Light As String = xmlEDS0067.SelectSingleNode("xsi:Light", nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

          Dim Value4 As Single = Math.Abs(Single.Parse(Light, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Light", "A", ROMId, Value4)

        Next
      End If

      '
      ' Get the EDS0068 (Temp, humidity, barometric pressure and light)
      ' 
      Dim xmlEDS0068s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0068", nsMgr)
      If xmlEDS0068s.Count > 0 Then
        For Each xmlEDS0068 As XmlNode In xmlEDS0068s
          Dim Family As String = xmlEDS0068.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0068.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0068.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS0068.SelectSingleNode("xsi:Humidity", nsMgr).InnerText
          Dim BarometricPressure As String = xmlEDS0068.SelectSingleNode(BarometricPressureXPath, nsMgr).InnerText
          Dim Light As String = xmlEDS0068.SelectSingleNode("xsi:Light", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

          Dim Value3 As Single = Math.Abs(Single.Parse(BarometricPressure, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Pressure", "A", ROMId, Value3)

          Dim Value4 As Single = Math.Abs(Single.Parse(Light, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Light", "A", ROMId, Value4)

        Next
      End If

      '
      ' Get the EDS0071 (Temp)
      ' 
      Dim xmlEDS0071s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0071", nsMgr)
      If xmlEDS0071s.Count > 0 Then
        For Each xmlEDS0071 As XmlNode In xmlEDS0071s
          Dim Family As String = xmlEDS0071.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlEDS0071.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Temperature As String = xmlEDS0071.SelectSingleNode("xsi:Temperature", nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

        Next
      End If

      '<owd_DS2423 Description="RAM with counters">
      ' <Name>DS2423</Name> 
      ' <Family>1D</Family> 
      ' <ROMId>1D1A8C0C000000AF</ROMId> 
      ' <Health>7</Health> 
      ' <RawData>01000000010000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000</RawData> 
      ' <PrimaryValue>1, 1</PrimaryValue> 
      ' <Counter_A>1</Counter_A> 
      ' <Counter_B>1</Counter_B> 
      '</owd_DS2423>

      '
      ' Get the DS2423
      '
      Dim xmlDS2423s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_DS2423", nsMgr)
      If xmlDS2423s.Count > 0 Then
        Dim SensorType As String = "Counter"
        For Each xmlDS2423 As XmlNode In xmlDS2423s
          Dim Family As String = xmlDS2423.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlDS2423.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim Counter_A As String = xmlDS2423.SelectSingleNode("xsi:Counter_A", nsMgr).InnerText
          Dim Counter_B As String = xmlDS2423.SelectSingleNode("xsi:Counter_B", nsMgr).InnerText

          Dim ValueA As Integer = Integer.Parse(Counter_A, nfi)
          Dim ValueB As Integer = Integer.Parse(Counter_B, nfi)
          UpdateOneWireSensor(OWServer.DeviceId, "Counter", SensorType, "A", ROMId, ValueA)
          UpdateOneWireSensor(OWServer.DeviceId, "Counter", SensorType, "B", ROMId, ValueB)

        Next
      End If

      '<owd_DS2406 Description="Dual addressable switch plus memory">
      ' <Name>DS2406</Name> 
      ' <Family>12</Family> 
      ' <ROMId>122CBE1C0000008D</ROMId> 
      ' <Health>7</Health> 
      ' <RawData>7F000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000</RawData> 
      ' <PrimaryValue>A=1, B=1</PrimaryValue> 
      ' <InputLevel_A>1</InputLevel_A> 
      ' <InputLevel_B>1</InputLevel_B> 
      ' <FlipFlop_A Writable="True">1</FlipFlop_A> 
      ' <FlipFlop_B Writable="True">1</FlipFlop_B> 
      ' <ActivityLatch_A>1</ActivityLatch_A> 
      ' <ActivityLatch_B>1</ActivityLatch_B> 
      ' <NumberOfChannels>2</NumberOfChannels> 
      ' <PowerSource>0</PowerSource> 
      ' <ActivityLatchReset Writable="True">-</ActivityLatchReset> 
      '</owd_DS2406>

      '
      ' Get the DS2406
      '
      Dim xmlDS2406s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_DS2406", nsMgr)
      If xmlDS2406s.Count > 0 Then
        Dim SensorType As String = "Switch"
        For Each xmlDS2406 As XmlNode In xmlDS2406s
          Dim Family As String = xmlDS2406.SelectSingleNode("xsi:Family", nsMgr).InnerText
          Dim ROMId As String = xmlDS2406.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim InputLevel_A As String = xmlDS2406.SelectSingleNode("xsi:InputLevel_A", nsMgr).InnerText
          Dim InputLevel_B As String = xmlDS2406.SelectSingleNode("xsi:InputLevel_B", nsMgr).InnerText

          Dim ValueA As Single = CByte(InputLevel_A)
          Dim ValueB As Single = CByte(InputLevel_B)
          UpdateOneWireSensor(OWServer.DeviceId, "Switch", SensorType, "A", ROMId, ValueA)
          UpdateOneWireSensor(OWServer.DeviceId, "Switch", SensorType, "B", ROMId, ValueB)

        Next
      End If

      ' <owd_EDS0070 Description="Vibration Sensor">
      ' <Name>EDS0070</Name>
      ' <Family>7E</Family>
      ' <ROMId>4A00100000237F7E</ROMId>
      ' <Health>7</Health>
      ' <Channel>1</Channel>
      ' <RawData>
      ' 70000000000000005C020000000000000000000000000000000000000EC60F0000000000FF0300000000000000000000000000000000000000000000000000000000000000000000
      ' </RawData>
      ' <PrimaryValue>0</PrimaryValue>
      ' <VibrationInstant>0</VibrationInstant>
      ' <VibrationPeak>0</VibrationPeak>
      ' <VibrationMaximum>604</VibrationMaximum>
      ' <VibrationMinimum>0</VibrationMinimum>
      ' <LED>0</LED>
      ' <Relay>0</Relay>
      ' <VibrationHighAlarmState>0</VibrationHighAlarmState>
      ' <VibrationLowAlarmState>0</VibrationLowAlarmState>
      ' <Counter>1033742</Counter>
      ' <ClearAlarms Writable="True">-</ClearAlarms>
      ' <VibrationHighConditionalSearchState Writable="True">0</VibrationHighConditionalSearchState>
      ' <VibrationLowConditionalSearchState Writable="True">0</VibrationLowConditionalSearchState>
      ' <VibrationHighAlarmValue Writable="True">1023</VibrationHighAlarmValue>
      ' <VibrationLowAlarmValue Writable="True">0</VibrationLowAlarmValue>
      ' <LEDFunction Writable="True">0</LEDFunction>
      ' <RelayFunction Writable="True">0</RelayFunction>
      ' <LEDState Writable="True">0</LEDState>
      ' <RelayState Writable="True">0</RelayState>
      ' <Version>1.00</Version>
      ' </owd_EDS0070>

      '
      ' Get the EDS0070 (Vibration)
      ' 
      Dim xmlEDS0070s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS0070", nsMgr)
      If xmlEDS0070s.Count > 0 Then
        For Each xmlEDS0070 As XmlNode In xmlEDS0070s
          Dim ROMId As String = xmlEDS0070.SelectSingleNode("xsi:ROMId", nsMgr).InnerText
          Dim VibrationPeak As String = xmlEDS0070.SelectSingleNode("xsi:VibrationPeak", nsMgr).InnerText

          Dim Value1 As Double = Math.Abs(Double.Parse(VibrationPeak, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Vibration", "A", ROMId, Value1)

        Next
      End If

      '
      ' New Sensors
      '

      '
      ' Get the EDS1064 (Temp)
      ' 
      Dim xmlEDS1064s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS1064", nsMgr)
      If xmlEDS1064s.Count > 0 Then
        For Each EDS1064 As XmlNode In xmlEDS1064s
          Dim ROMId As String = EDS1064.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = EDS1064.SelectSingleNode("xsi:Temperature", nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

        Next
      End If

      '
      ' Get the EDS0065 (Temp, humidity)
      ' 
      Dim xmlEDS1065s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS1065", nsMgr)
      If xmlEDS1065s.Count > 0 Then
        For Each xmlEDS1065 As XmlNode In xmlEDS1065s
          Dim ROMId As String = xmlEDS1065.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS1065.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS1065.SelectSingleNode("xsi:Humidity", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

        Next
      End If

      '
      ' Get the EDS0066 (Temp, barometric pressure)
      ' 
      Dim xmlEDS1066s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS1066", nsMgr)
      If xmlEDS1066s.Count > 0 Then
        For Each xmlEDS1066 As XmlNode In xmlEDS1066s
          Dim ROMId As String = xmlEDS1066.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS1066.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim BarometricPressure As String = xmlEDS1066.SelectSingleNode(BarometricPressureXPath, nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

          Dim Value3 As Single = OWServer.ConvertPressure(BarometricPressure, bMetric)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Pressure", "A", ROMId, Value3)

        Next
      End If

      '
      ' Get the EDS0066 (Temp, light)
      ' 
      Dim xmlEDS1067s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS1067", nsMgr)
      If xmlEDS1067s.Count > 0 Then
        For Each xmlEDS1067 As XmlNode In xmlEDS1067s
          Dim ROMId As String = xmlEDS1067.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS1067.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Light As String = xmlEDS1067.SelectSingleNode("xsi:Light", nsMgr).InnerText

          Dim Value1 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value1)

          Dim Value4 As Single = Math.Abs(Single.Parse(Light, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Light", "A", ROMId, Value4)

        Next
      End If

      '
      ' Get the EDS0068 (Temp, humidity, barometric pressure and light)
      ' 
      Dim xmlEDS1068s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS1068", nsMgr)
      If xmlEDS1068s.Count > 0 Then
        For Each xmlEDS1068 As XmlNode In xmlEDS1068s
          Dim ROMId As String = xmlEDS1068.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS1068.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS1068.SelectSingleNode("xsi:Humidity", nsMgr).InnerText
          Dim BarometricPressure As String = xmlEDS1068.SelectSingleNode("xsi:BarometricPressure", nsMgr).InnerText
          Dim Light As String = xmlEDS1068.SelectSingleNode("xsi:Light", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

          Dim Value3 As Single = OWServer.ConvertPressure(BarometricPressure, bMetric)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Pressure", "A", ROMId, Value3)

          Dim Value4 As Single = Math.Abs(Single.Parse(Light, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Light", "A", ROMId, Value4)

        Next
      End If

      '
      ' Get the EDS2040 (Temperature, Humidity, Discrete Input, and External Probe Support with High Power Radio)
      '
      Dim xmlEDS2040s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS2040", nsMgr)
      If xmlEDS2040s.Count > 0 Then
        For Each xmlEDS2040 As XmlNode In xmlEDS2040s
          Dim ROMId As String = xmlEDS2040.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS2040.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS2040.SelectSingleNode("xsi:Humidity", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

          '
          ' Process External Probes
          '
          Dim OW1Value As String = xmlEDS2040.SelectSingleNode("xsi:OW1Value", nsMgr).InnerText
          Dim OW2Value As String = xmlEDS2040.SelectSingleNode("xsi:OW2Value", nsMgr).InnerText
          Dim OW3Value As String = xmlEDS2040.SelectSingleNode("xsi:OW3Value", nsMgr).InnerText

          Dim OW1Health As String = xmlEDS2040.SelectSingleNode("xsi:OW1Health", nsMgr).InnerText
          Dim OW2Health As String = xmlEDS2040.SelectSingleNode("xsi:OW2Health", nsMgr).InnerText
          Dim OW3Health As String = xmlEDS2040.SelectSingleNode("xsi:OW3Health", nsMgr).InnerText

          Dim OW1ROMID As String = xmlEDS2040.SelectSingleNode("xsi:OW1ROMID", nsMgr).InnerText
          Dim OW2ROMID As String = xmlEDS2040.SelectSingleNode("xsi:OW2ROMID", nsMgr).InnerText
          Dim OW3ROMID As String = xmlEDS2040.SelectSingleNode("xsi:OW3ROMID", nsMgr).InnerText

          Dim OW1DataType As String = xmlEDS2040.SelectSingleNode("xsi:OW1DataType", nsMgr).InnerText
          Dim OW2DataType As String = xmlEDS2040.SelectSingleNode("xsi:OW2DataType", nsMgr).InnerText
          Dim OW3DataType As String = xmlEDS2040.SelectSingleNode("xsi:OW3DataType", nsMgr).InnerText

          If OW1DataType = "Temperature" AndAlso OW1Health > 0 Then
            Dim OW1Value1 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW1ROMID, OW1Value1)
          End If

          If OW2DataType = "Temperature" AndAlso OW2Health > 0 Then
            Dim OW2Value2 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW2ROMID, OW2Value2)
          End If

          If OW3DataType = "Temperature" AndAlso OW3Health > 0 Then
            Dim OW3Value3 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW3ROMID, OW3Value3)
          End If

        Next

      End If

      '
      ' Get the EDS2040 (Temperature, Humidity, Discrete Input, and External Probe Support with High Power Radio)
      '
      Dim xmlEDS2041s As XmlNodeList = xmlDoc.SelectNodes("//xsi:owd_EDS2040", nsMgr)
      If xmlEDS2041s.Count > 0 Then
        For Each xmlEDS2041 As XmlNode In xmlEDS2041s
          Dim ROMId As String = xmlEDS2041.SelectSingleNode("xsi:EUI", nsMgr).InnerText
          Dim Temperature As String = xmlEDS2041.SelectSingleNode("xsi:Temperature", nsMgr).InnerText
          Dim Humidity As String = xmlEDS2041.SelectSingleNode("xsi:Humidity", nsMgr).InnerText

          Dim Value1 As Single = Math.Abs(Single.Parse(Humidity, nfi))
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Humidity", "A", ROMId, Value1)

          Dim Value2 As Single = OWServer.ConvertTemperature(Temperature, bFarenheight)
          UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", ROMId, Value2)

          '
          ' Process External Probes
          '
          Dim OW1Value As String = xmlEDS2041.SelectSingleNode("xsi:OW1Value", nsMgr).InnerText
          Dim OW2Value As String = xmlEDS2041.SelectSingleNode("xsi:OW2Value", nsMgr).InnerText
          Dim OW3Value As String = xmlEDS2041.SelectSingleNode("xsi:OW3Value", nsMgr).InnerText

          Dim OW1Health As String = xmlEDS2041.SelectSingleNode("xsi:OW1Health", nsMgr).InnerText
          Dim OW2Health As String = xmlEDS2041.SelectSingleNode("xsi:OW2Health", nsMgr).InnerText
          Dim OW3Health As String = xmlEDS2041.SelectSingleNode("xsi:OW3Health", nsMgr).InnerText

          Dim OW1ROMID As String = xmlEDS2041.SelectSingleNode("xsi:OW1ROMID", nsMgr).InnerText
          Dim OW2ROMID As String = xmlEDS2041.SelectSingleNode("xsi:OW2ROMID", nsMgr).InnerText
          Dim OW3ROMID As String = xmlEDS2041.SelectSingleNode("xsi:OW3ROMID", nsMgr).InnerText

          Dim OW1DataType As String = xmlEDS2041.SelectSingleNode("xsi:OW1DataType", nsMgr).InnerText
          Dim OW2DataType As String = xmlEDS2041.SelectSingleNode("xsi:OW2DataType", nsMgr).InnerText
          Dim OW3DataType As String = xmlEDS2041.SelectSingleNode("xsi:OW3DataType", nsMgr).InnerText

          If OW1DataType = "Temperature" AndAlso OW1Health > 0 Then
            Dim OW1Value1 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW1ROMID, OW1Value1)
          End If

          If OW2DataType = "Temperature" AndAlso OW2Health > 0 Then
            Dim OW2Value2 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW2ROMID, OW2Value2)
          End If

          If OW3DataType = "Temperature" AndAlso OW3Health > 0 Then
            Dim OW3Value3 As Single = OWServer.ConvertTemperature(OW1Value, bFarenheight)
            UpdateOneWireSensor(OWServer.DeviceId, "Environmental", "Temperature", "A", OW3ROMID, OW3Value3)
          End If

        Next

      End If

    Catch pEx As WebException
      '
      ' Process the error
      '
      WriteMessage("The OWServer connection failed due to an error: " & pEx.Message, MessageType.Error)
    Catch pEx As XmlException
      '
      ' Process the error
      '
      WriteMessage("The XML document returned by the OWServer was not valid XML.", MessageType.Error)
    Catch pEx As Exception
      '
      ' Process the error
      '
      WriteMessage("The OWServer connection failed.", MessageType.Error)
    End Try

  End Sub

End Module