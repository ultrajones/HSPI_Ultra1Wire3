Imports System.Data
Imports System.Data.Common
Imports System.Threading
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports System.Configuration
Imports System.Text
Imports System.IO
Imports System.Data.SQLite

Module hspi_database

  Public DBConnectionMain As SQLite.SQLiteConnection  ' Our main database connection
  Public DBConnectionTemp As SQLite.SQLiteConnection  ' Our temp database connection

  Public gDBInsertSuccess As ULong = 0            ' Tracks DB insert success
  Public gDBInsertFailure As ULong = 0            ' Tracks DB insert success

  Public bDBInitialized As Boolean = False        ' Indicates if database successfully initialized

  Public SyncLockMain As New Object
  Public SyncLockTemp As New Object

#Region "Database Initilization"

  ''' <summary>
  ''' Initializes the database
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitializeMainDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered InitializeMainDatabase() function.", MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionMain Is Nothing Then
        If CloseDBConn(DBConnectionMain) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Create the database directory if it does not exist
      '
      Dim databaseDir As String = FixPath(String.Format("{0}\Data\{1}\", hs.GetAppPath, IFACE_NAME.ToLower))
      If Directory.Exists(databaseDir) = False Then
        Directory.CreateDirectory(databaseDir)
      End If

      '
      ' Determine the database filename
      '
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionMain, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeDatabase()")
    End Try

    Return bSuccess

  End Function

  '------------------------------------------------------------------------------------
  'Purpose: Initializes the temporary database
  'Inputs:  None
  'Outputs: True or False indicating if database was initialized
  '------------------------------------------------------------------------------------
  Public Function InitializeTempDatabase() As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    strMessage = "Entered InitializeChannelDatabase() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Close database if it's open
      '
      If Not DBConnectionTemp Is Nothing Then
        If CloseDBConn(DBConnectionTemp) = False Then
          Throw New Exception("An existing database connection could not be closed.")
        End If
      End If

      '
      ' Determine the database filename
      '
      Dim dtNow As DateTime = DateTime.Now
      Dim iHour As Integer = dtNow.Hour
      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}_{2}.db3", hs.GetAppPath(), IFACE_NAME.ToLower, strDBDate))

      '
      ' Determine the database provider factory and connection string
      '
      Dim strConnectionString As String = String.Format("Data Source={0}; Version=3; Journal Mode=Off;", strDataSource)

      '
      ' Attempt to open the database connection
      '
      bSuccess = OpenDBConn(DBConnectionTemp, strConnectionString)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      bSuccess = False
      Call ProcessError(pEx, "InitializeTempDatabase()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Opens a connection to the database
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <param name="strConnectionString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function OpenDBConn(ByRef objConn As SQLite.SQLiteConnection,
                              ByVal strConnectionString As String) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered OpenDBConn() function.", MessageType.Debug)

    Try
      '
      ' Open database connection
      '
      objConn = New SQLite.SQLiteConnection()
      objConn.ConnectionString = strConnectionString
      objConn.Open()

      '
      ' Run database vacuum
      '
      WriteMessage("Running SQLite database vacuum.", MessageType.Debug)
      Using MyDbCommand As DbCommand = objConn.CreateCommand()

        MyDbCommand.Connection = objConn
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = "VACUUM"
        MyDbCommand.ExecuteNonQuery()

        MyDbCommand.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "OpenDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = objConn.State = ConnectionState.Open

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database initialization complete."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Database initialization failed using [" & strConnectionString & "]."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Closes database connection
  ''' </summary>
  ''' <param name="objConn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CloseDBConn(ByRef objConn As SQLite.SQLiteConnection) As Boolean

    Dim strMessage As String = ""               ' Holds informational messages
    Dim bSuccess As Boolean = False             ' Indicate default success

    WriteMessage("Entered CloseDBConn() function.", MessageType.Debug)

    Try
      '
      ' Attempt to the database
      '
      If objConn.State <> ConnectionState.Closed Then
        objConn.Close()
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CloseDBConn()")
    End Try

    '
    ' Determine database connection status
    '
    bSuccess = objConn.State = ConnectionState.Closed

    '
    ' Record connection state to HomeSeer log
    '
    If bSuccess = True Then
      strMessage = "Database connection closed successfuly."
      Call WriteMessage(strMessage, MessageType.Debug)
    Else
      strMessage = "Unable to close database; Try restarting HomeSeer."
      Call WriteMessage(strMessage, MessageType.Debug)
    End If

    Return bSuccess

  End Function

  ''' <summary>
  ''' Checks to ensure a table exists
  ''' </summary>
  ''' <param name="strTableName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckDatabaseTable(ByVal strTableName As String) As Boolean

    Dim strMessage As String = ""
    Dim bSuccess As Boolean = False

    WriteMessage("Entered CheckDatabaseTable() function.", MessageType.Debug)

    Try
      '
      ' Build SQL delete statement
      '
      If String.Compare(strTableName, "tblDevices", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue

            sqlQueue.Enqueue("CREATE TABLE " & strTableName & "(" _
                            & "device_id INTEGER PRIMARY KEY," _
                            & "device_serial varchar(25) NOT NULL," _
                            & "device_name varchar(25) NOT NULL," _
                            & "device_type varchar(255) NOT NULL," _
                            & "device_image varchar(255) NOT NULL," _
                            & "device_conn varchar(255) NOT NULL," _
                            & "device_addr varchar(25) NOT NULL" _
                          & ")")

            sqlQueue.Enqueue("CREATE UNIQUE INDEX idxDevices1 ON tblDevices(device_addr)")

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf String.Compare(strTableName, "tblSensorConfig", True) = 0 Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue
            sqlQueue.Enqueue("CREATE TABLE " & strTableName & "(" _
                           & "sensor_id INTEGER PRIMARY KEY," _
                           & "device_id INTEGER NOT NULL," _
                           & "sensor_addr TEXT," _
                           & "sensor_name TEXT NOT NULL," _
                           & "sensor_type TEXT NOT NULL," _
                           & "sensor_subtype TEXT NOT NULL," _
                           & "sensor_channel TEXT NOT NULL," _
                           & "sensor_color TEXT NOT NULL DEFAULT ''," _
                           & "sensor_image TEXT NOT NULL DEFAULT ''," _
                           & "sensor_units TEXT NOT NULL DEFAULT ''," _
                           & "sensor_resolution REAL NOT NULL DEFAULT 1," _
                           & "sensor_reset INTEGER NOT NULL DEFAULT 0," _
                           & "post_enabled INTEGER NOT NULL DEFAULT 0," _
                           & "dev_enabled INTEGER NOT NULL DEFAULT 1," _
                           & "dev_00d INTEGER NOT NULL DEFAULT 0," _
                           & "dev_01d INTEGER NOT NULL DEFAULT 0," _
                           & "dev_07d INTEGER NOT NULL DEFAULT 0," _
                           & "dev_30d INTEGER NOT NULL DEFAULT 0" _
                           & ")")

            sqlQueue.Enqueue(String.Format("CREATE        INDEX idxSensorConfig1 ON {0} (device_id)", strTableName))
            sqlQueue.Enqueue(String.Format("CREATE UNIQUE INDEX idxSensorConfig2 ON {0} (sensor_addr)", strTableName))

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf Regex.IsMatch(strTableName, "Data") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionTemp.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionTemp.CreateCommand()

            Dim sqlQueue As New Queue
            sqlQueue.Enqueue("CREATE TABLE " & strTableName & "(" _
                           & "device_id INTEGER NOT NULL," _
                           & "sensor_id INTEGER NOT NULL," _
                           & "ts INTEGER NOT NULL," _
                           & "value INTEGER NOT NULL" _
                           & ")")

            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}1 ON {0} (device_id)", strTableName))
            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}2 ON {0} (sensor_id)", strTableName))
            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}3 ON {0} (ts)", strTableName))

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionTemp
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      ElseIf Regex.IsMatch(strTableName, "History") = True Then
        '
        ' Retrieve schema information about tblTableName
        '
        Dim SchemaTable As DataTable = DBConnectionMain.GetSchema("Columns", New String() {Nothing, Nothing, strTableName})

        If SchemaTable.Rows.Count <> 0 Then
          WriteMessage("Table " & SchemaTable.Rows(0)!TABLE_NAME.ToString & " exists.", MessageType.Debug)
        Else
          '
          ' Create the table
          '
          WriteMessage("Creating " & strTableName & " table ...", MessageType.Debug)

          Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

            Dim sqlQueue As New Queue
            sqlQueue.Enqueue("CREATE TABLE " & strTableName & "(" _
                           & "device_id INTEGER NOT NULL," _
                           & "sensor_id INTEGER NOT NULL," _
                           & "ts INTEGER NOT NULL," _
                           & "value INTEGER NOT NULL" _
                           & ")")

            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}1 ON {0} (device_id)", strTableName))
            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}2 ON {0} (sensor_id)", strTableName))
            sqlQueue.Enqueue(String.Format("CREATE INDEX idx{0}3 ON {0} (ts)", strTableName))

            While sqlQueue.Count > 0
              Dim strSQL As String = sqlQueue.Dequeue

              dbcmd.Connection = DBConnectionMain
              dbcmd.CommandType = CommandType.Text
              dbcmd.CommandText = strSQL

              Dim iRecordsAffected As Integer = dbcmd.ExecuteNonQuery()
              If iRecordsAffected <> 1 Then
                'Throw New Exception("Database schemea update failed due to error.")
              End If

            End While

            dbcmd.Dispose()
          End Using

        End If

      Else
        Throw New Exception(strTableName & " not currently supported.")
      End If

      bSuccess = True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckDatabaseTable()")
    End Try

    Return bSuccess

  End Function

  ''' <summary>
  ''' Returns the size of the selected database
  ''' </summary>
  ''' <param name="databaseName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDatabaseSize(ByVal databaseName As String)

    Try

      Select Case databaseName
        Case "DBConnectionMain"
          '
          ' Determine the database filename
          '
          Dim strDataSource As String = FixPath(String.Format("{0}\Data\{1}\{1}.db3", hs.GetAppPath(), IFACE_NAME.ToLower))
          Dim file As New FileInfo(strDataSource)
          Return FormatFileSize(file.Length)

      End Select

    Catch pEx As Exception

    End Try

    Return "Unknown"

  End Function

  ''' <summary>
  ''' Converts filesize to String
  ''' </summary>
  ''' <param name="FileSizeBytes"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function FormatFileSize(ByVal FileSizeBytes As Long) As String

    Try

      Dim sizeTypes() As String = {"B", "KB", "MB", "GB"}
      Dim Len As Decimal = FileSizeBytes
      Dim sizeType As Integer = 0

      Do While Len > 1024
        Len = Decimal.Round(Len / 1024, 2)
        sizeType += 1
        If sizeType >= sizeTypes.Length - 1 Then Exit Do
      Loop

      Dim fileSize As String = String.Format("{0} {1}", Len.ToString("N0"), sizeTypes(sizeType))
      Return fileSize

    Catch pEx As Exception

    End Try

    Return FileSizeBytes.ToString

  End Function

#End Region

#Region "Database Queries"

  ''' <summary>
  ''' Return values from database
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function QueryDatabase(ByVal strSQL As String) As DataSet

    Dim strMessage As String = ""

    WriteMessage("Entered QueryDatabase() function.", MessageType.Debug)

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

      '
      ' Initialize the command object
      '
      Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

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

      End Using

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "QueryDatabase()")
      Return New DataSet
    End Try

  End Function

  ''' <summary>
  ''' Insert data into database
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <remarks></remarks>
  Public Sub InsertData(ByVal strSQL As String)

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    Select Case DBConnectionTemp.State
      Case ConnectionState.Broken, ConnectionState.Closed
        Exit Sub
    End Select

    Try

      Using dbcmd As DbCommand = DBConnectionTemp.CreateCommand()

        dbcmd.Connection = DBConnectionTemp
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockTemp
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strSQL, MessageType.Debug)
    Finally
      '
      ' Update counter
      '
      If iRecordsAffected = 1 Then
        gDBInsertSuccess += 1
      Else
        gDBInsertFailure += 1
      End If
    End Try

  End Sub

  ''' <summary>
  ''' Insert History data into database
  ''' </summary>
  ''' <param name="strSQL"></param>
  ''' <remarks></remarks>
  Public Sub InsertHistoryData(ByVal strSQL As String)

    Dim strMessage As String = ""
    Dim iRecordsAffected As Integer = 0

    '
    ' Ensure database is loaded before attempting to use it
    '
    Select Case DBConnectionMain.State
      Case ConnectionState.Broken, ConnectionState.Closed
        Exit Sub
    End Select

    Try

      Using dbcmd As DbCommand = DBConnectionMain.CreateCommand()

        dbcmd.Connection = DBConnectionMain
        dbcmd.CommandType = CommandType.Text
        dbcmd.CommandText = strSQL

        SyncLock SyncLockMain
          iRecordsAffected = dbcmd.ExecuteNonQuery()
        End SyncLock

        dbcmd.Dispose()
      End Using

    Catch pEx As Exception
      '
      ' Process error
      '
      strMessage = "InsertHistoryData() Reports Error: [" & pEx.ToString & "], " _
                  & "Failed on SQL: " & strSQL & "."
      Call WriteMessage(strSQL, MessageType.Notice)
    Finally
      '
      ' Update counter
      '
      If iRecordsAffected = 1 Then
        gDBInsertSuccess += 1
      Else
        gDBInsertFailure += 1
      End If
    End Try

  End Sub

#End Region

#Region "Database Maintenance"

  ''' <summary>
  ''' Update the database to disable orphand sensors
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateOrphanedSensors()

    Try

      Dim sensors As New ArrayList

      Dim ts As Long = ConvertDateTimeToEpoch(DateAdd(DateInterval.Day, -7, Date.Now))

      Dim strSQL As String = String.Format("SELECT sensor_id from tblTemperatureHistory WHERE ts > {0} " & _
                                     "UNION SELECT sensor_id from tblHumidityHistory WHERE ts > {0} " & _
                                     "UNION SELECT sensor_id from tblCounterHistory WHERE ts > {0} " & _
                                     "UNION SELECT sensor_id from tblPressureHistory WHERE ts > {0}", ts.ToString)

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
              sensors.Add(dtrResults("sensor_id"))
            End While

            dtrResults.Close()
          End Using

        End SyncLock

        MyDbCommand.Dispose()

      End Using

      '
      ' Updated the active sensors
      '
      If sensors.Count > 0 Then
        Dim sensorList As String = String.Join(",", sensors.ToArray)

        '
        ' Build the insert/update/delete query
        '
        Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

          MyDbCommand.Connection = DBConnectionMain
          MyDbCommand.CommandType = CommandType.Text
          MyDbCommand.CommandText = String.Format("UPDATE tblSensorConfig SET dev_enabled=0 WHERE sensor_id NOT IN ({0})", sensorList)

          Dim iRecordsAffected As Integer = 0
          SyncLock SyncLockMain
            iRecordsAffected = MyDbCommand.ExecuteNonQuery()
          End SyncLock

          MyDbCommand.Dispose()

          If iRecordsAffected > 0 Then
            Dim strMessage As String = "UpdateOrphanedSensors() disabled " & iRecordsAffected & " orphaned sensors."
            Call WriteMessage(strMessage, MessageType.Warning)
          End If

        End Using

      End If

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      ProcessError(pEx, "UpdateOrphanedSensors()")
    End Try

  End Sub

  ''' <summary>
  ''' Updated the database to reflect the enabled state of the connected sensors
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateConnectedSensors()

    Try

      Dim sensors As New ArrayList

      '
      ' Update any sensors that are reporting
      '
      For Each OneWireSensor In OneWireSensors
        Dim dv_addr As String = OneWireSensor.dvAddress

        Dim sensorId As Integer = OneWireSensor.sensorId
        If sensors.Contains(sensorId) = False Then
          sensors.Add(sensorId)
        End If
      Next

      '
      ' Updated the active sensors
      '
      If sensors.Count > 0 Then
        Dim sensorList As String = String.Join(",", sensors.ToArray)

        '
        ' Build the insert/update/delete query
        '
        Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

          MyDbCommand.Connection = DBConnectionMain
          MyDbCommand.CommandType = CommandType.Text
          MyDbCommand.CommandText = String.Format("UPDATE tblSensorConfig SET dev_enabled=0 WHERE sensor_id NOT IN ({0}) AND dev_enabled=1", sensorList)

          Dim iRecordsAffected As Integer = 0
          SyncLock SyncLockMain
            iRecordsAffected = MyDbCommand.ExecuteNonQuery()
          End SyncLock

          MyDbCommand.Dispose()

          If iRecordsAffected > 0 Then
            Dim strMessage As String = "UpdateConnectedSensors() disabled " & iRecordsAffected & " offline sensors."
            Call WriteMessage(strMessage, MessageType.Warning)
          End If

        End Using

        '
        ' Build the insert/update/delete query
        '
        Using MyDbCommand As DbCommand = DBConnectionMain.CreateCommand()

          MyDbCommand.Connection = DBConnectionMain
          MyDbCommand.CommandType = CommandType.Text
          MyDbCommand.CommandText = String.Format("UPDATE tblSensorConfig SET dev_enabled=1 WHERE sensor_id IN ({0}) AND dev_enabled=0", sensorList)

          Dim iRecordsAffected As Integer = 0
          SyncLock SyncLockMain
            iRecordsAffected = MyDbCommand.ExecuteNonQuery()
          End SyncLock

          MyDbCommand.Dispose()

          If iRecordsAffected > 0 Then
            Dim strMessage As String = "UpdateConnectedSensors() enabled " & iRecordsAffected & " online sensors."
            Call WriteMessage(strMessage, MessageType.Informational)
          End If

        End Using

      End If

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      ProcessError(pEx, "UpdateConnectedSensors()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose: Checks to see if new database should be created
  'Inputs:  None
  'Outputs: None
  '------------------------------------------------------------------------------------
  Public Sub CheckDatabase()

    Try
      '
      ' Check to see if the database needs to be initialized
      '
      If DBConnectionTemp Is Nothing Then
        Call InitializeTempDatabase()

        Call CheckDatabaseTable("tblTemperatureData")
        Call CheckDatabaseTable("tblHumidityData")
        Call CheckDatabaseTable("tblCounterData")
        Call CheckDatabaseTable("tblPressureData")
      End If

      '
      ' Define the database connection
      '
      Dim dtNow As DateTime = DateTime.Now
      Dim iHour As Integer = dtNow.Hour

      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strConnectionString As String = DBConnectionTemp.ConnectionString

      If Regex.IsMatch(strConnectionString.ToLower, strDBDate) = True Then

        Dim strMessage As String = String.Format("CheckDatabase() reports hour database {0} exists.", strDBDate)
        Call WriteMessage(strMessage, MessageType.Debug)

      Else
        '
        ' Process the existing sensor history database
        '
        Call ProcessTemperatureHistoryData()
        Call ProcessHumidityHistoryData()
        Call ProcessCounterHistoryData()
        Call ProcessPressureHistoryData()

        Dim strMessage As String = String.Format("CheckDatabase() reports hour database {0} created.", strDBDate)
        Call WriteMessage(strMessage, MessageType.Debug)

        SyncLock SyncLockTemp
          Call InitializeTempDatabase()
          Call CheckDatabaseTable("tblTemperatureData")
          Call CheckDatabaseTable("tblHumidityData")
          Call CheckDatabaseTable("tblCounterData")
          Call CheckDatabaseTable("tblPressureData")
        End SyncLock

        Call PurgeDatabase(iHour)
      End If

    Catch pEx As Exception
      WriteMessage(pEx.Message, MessageType.Debug)
    End Try

  End Sub

  ''' <summary>
  ''' Process temperature history
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessTemperatureHistoryData()

    Try

      Dim ts As Long = ConvertDateTimeToEpoch(DateAdd(DateInterval.Hour, -1, DateTime.Now))
      Dim MyDS As DataSet = GetSensorHistoryData("tblTemperatureData", ts)

      For Each DataRow As DataRow In MyDS.Tables(0).Rows

        Dim strSQLInsert As String = String.Format("INSERT INTO tblTemperatureHistory (device_id, sensor_id, ts, value) VALUES ({0}, {1}, {2}, {3})", _
                                                   DataRow("device_id"), _
                                                   DataRow("sensor_id"), _
                                                   ts.ToString(), _
                                                   DataRow("sensor_value"))

        InsertHistoryData(strSQLInsert)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Notice)
    End Try

  End Sub

  ''' <summary>
  ''' Process Humidity history
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessHumidityHistoryData()

    Try

      Dim ts As Long = ConvertDateTimeToEpoch(DateAdd(DateInterval.Hour, -1, DateTime.Now))
      Dim MyDS As DataSet = GetSensorHistoryData("tblHumidityData", ts)

      For Each DataRow As DataRow In MyDS.Tables(0).Rows

        Dim strSQLInsert As String = String.Format("INSERT INTO tblHumidityHistory (device_id, sensor_id, ts, value) VALUES ({0}, {1}, {2}, {3})", _
                                                   DataRow("device_id"), _
                                                   DataRow("sensor_id"), _
                                                   ts.ToString(), _
                                                   DataRow("sensor_value"))

        InsertHistoryData(strSQLInsert)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Notice)
    End Try

  End Sub

  ''' <summary>
  ''' Process counter history
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessCounterHistoryData()

    Try

      Dim ts As Long = ConvertDateTimeToEpoch(DateAdd(DateInterval.Hour, -1, DateTime.Now))
      Dim MyDS As DataSet = GetSensorHistoryData("tblCounterData", ts)

      For Each DataRow As DataRow In MyDS.Tables(0).Rows

        Dim strSQLInsert As String = String.Format("INSERT INTO tblCounterHistory (device_id, sensor_id, ts, value) VALUES ({0}, {1}, {2}, {3})", _
                                                   DataRow("device_id"), _
                                                   DataRow("sensor_id"), _
                                                   ts.ToString(), _
                                                   DataRow("sensor_value"))

        InsertHistoryData(strSQLInsert)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Notice)
    End Try

  End Sub

  ''' <summary>
  ''' Process Pressure history
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ProcessPressureHistoryData()

    Try

      Dim ts As Long = ConvertDateTimeToEpoch(DateAdd(DateInterval.Hour, -1, DateTime.Now))
      Dim MyDS As DataSet = GetSensorHistoryData("tblPressureData", ts)

      For Each DataRow As DataRow In MyDS.Tables(0).Rows

        Dim strSQLInsert As String = String.Format("INSERT INTO tblPressureHistory (device_id, sensor_id, ts, value) VALUES ({0}, {1}, {2}, {3})", _
                                                   DataRow("device_id"), _
                                                   DataRow("sensor_id"), _
                                                   ts.ToString(), _
                                                   DataRow("sensor_value"))

        InsertHistoryData(strSQLInsert)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Notice)
    End Try

  End Sub

  ''' <summary>
  ''' Get the 1-Wire Sensor History Data
  ''' </summary>
  ''' <param name="ts"></param>
  ''' <param name="device_id"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSensorHistoryData(dbTable As String, ByVal ts As Long, Optional ByVal device_id As Integer = 0) As DataSet

    Try
      '
      ' Define the SQL Query
      '
      Dim dbWhere As String = String.Format("ts >= {0}", ts.ToString)
      If device_id > 0 Then
        dbWhere &= String.Format(" AND device_id={0}", device_id.ToString)
      End If

      Dim dbField As String
      Select Case dbTable
        Case "tblCounterData"
          dbField = "SUM(value)"
        Case Else
          dbField = "AVG(value)"
      End Select

      Dim strSQL As String = String.Format("SELECT device_id, sensor_id, {0} as sensor_value FROM {1} WHERE {2} GROUP BY sensor_id", dbField, dbTable, dbWhere)

      '
      ' Execute the data reader
      '
      Using MyDbCommand As DbCommand = DBConnectionTemp.CreateCommand()

        MyDbCommand.Connection = DBConnectionTemp
        MyDbCommand.CommandType = CommandType.Text
        MyDbCommand.CommandText = strSQL

        '
        ' Initialize the dataset, then populate it
        '
        Dim MyDS As DataSet = New DataSet

        Dim MyDA As System.Data.IDbDataAdapter = New SQLiteDataAdapter(MyDbCommand)
        MyDA.SelectCommand = MyDbCommand

        SyncLock SyncLockTemp
          MyDA.Fill(MyDS)
        End SyncLock

        MyDbCommand.Dispose()

        Return MyDS

      End Using

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "GetSensorHistoryData()")
      Return New DataSet
    End Try

  End Function

  ''' <summary>
  ''' Purge the previous database
  ''' </summary>
  ''' <param name="iHour"></param>
  ''' <remarks></remarks>
  Public Sub PurgeDatabase(ByVal iHour As Integer)

    Try
      '
      ' Determine the database filename
      '
      iHour -= 1
      If iHour < 0 Then iHour = 23

      Dim strDBDate As String = iHour.ToString.PadLeft(2, "0")
      Dim strDBFile As String = String.Format("{0}\data\{1}\{1}_{2}.db3", hs.GetAppPath(), IFACE_NAME, strDBDate).ToLower

      If System.IO.File.Exists(strDBFile) = True Then
        '
        ' Delete the file
        '
        System.IO.File.Delete(strDBFile)

        Dim strMessage As String = String.Format("PurgeDatabase() reports database {0} deleted.", strDBFile)
        Call WriteMessage(strMessage, MessageType.Debug)

      Else

        Dim strMessage As String = String.Format("PurgeDatabase() reports database {0} not found.", strDBFile)
        Call WriteMessage(strMessage, MessageType.Debug)

      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Debug)
    End Try

  End Sub

#End Region

#Region "Database Date Formatting"

  ''' <summary>
  ''' dateTime as DateTime
  ''' </summary>
  ''' <param name="dateTime"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertDateTimeToEpoch(ByVal dateTime As DateTime) As Long

    Dim baseTicks As Long = 621355968000000000
    Dim tickResolution As Long = 10000000

    Return (dateTime.ToUniversalTime.Ticks - baseTicks) / tickResolution

  End Function

  ''' <summary>
  ''' Converts Epoch to datetime
  ''' </summary>
  ''' <param name="epochTicks"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function ConvertEpochToDateTime(ByVal epochTicks As Long) As DateTime

    '
    ' Create a new DateTime value based on the Unix Epoch
    '
    Dim converted As New DateTime(1970, 1, 1, 0, 0, 0, 0)

    '
    ' Return the value in string format
    '
    Return converted.AddSeconds(epochTicks).ToLocalTime

  End Function

#End Region

End Module

