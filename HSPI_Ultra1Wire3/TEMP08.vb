Imports System.Threading
Imports System.Text
Imports System.Net.Sockets
Imports System.Net
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Collections.Generic

Public Class TEMP08

  ' Provides access to the TEMP08 1-Wire Bus Master interface

  Public InterfaceID As String        ' Tracks Interface ID
  Public InterfaceType As String      ' Tracks Interface Type

  Public MAX_ATTEMPTS As Byte = 2
  Public MAX_WAIT_TIME As Single = 1.5

  Protected CheckTEMP08Thread As Thread
  Protected WatchdogThread As Thread

  Protected CommandQueue As New Queue

  Protected m_DeviceId As Integer = 0
  Protected m_DeviceSerial As String = ""
  Protected m_DeviceConnectionType As String = ""
  Protected m_DeviceConnectionAddr As String = ""

  Protected m_DeviceType As String = ""

  Protected gDeviceInitialized As Boolean = False
  Protected gDeviceConnected As Boolean = False
  Protected gDeviceResponse As Boolean = False

  Protected m_Initialized As Boolean = False
  Protected m_WatchdogActive As Boolean = False
  Protected m_WatchdogDisabled As Boolean = False

  Protected m_LastReport As DateTime = DateTime.Now

  Protected strCmdWait As String = ""
  Protected iCmdAttempt As Byte = 0

  Dim WithEvents serialPort As New IO.Ports.SerialPort

#Region "TEMP08 Object"

  Public ReadOnly Property GetDeviceId() As Integer
    Get
      Return Me.m_DeviceId
    End Get
  End Property

  Public ReadOnly Property GetDeviceSerial() As String
    Get
      Return Me.m_DeviceSerial
    End Get
  End Property

  Public ReadOnly Property GetConnectionType() As String
    Get
      Return Me.m_DeviceConnectionType
    End Get
  End Property

  Public ReadOnly Property GetConnectionAddr() As String
    Get
      Return Me.m_DeviceConnectionAddr
    End Get
  End Property

  Public ReadOnly Property GetConnectionStatus() As String
    Get
      Select Case gDeviceConnected
        Case True
          Return "Connected"
        Case Else
          Return "Disconnected"
      End Select
    End Get
  End Property

  Public ReadOnly Property Type() As String
    Get
      Return Me.m_DeviceType
    End Get
  End Property

  Public Property LastReport() As DateTime
    Set(ByVal value As DateTime)
      m_LastReport = value
    End Set
    Get
      Return m_LastReport
    End Get
  End Property

  ''' <summary>
  ''' Creates new TEMP08 Device Object
  ''' </summary>
  ''' <param name="deviceId"></param>
  ''' <param name="deviceType"></param>
  ''' <param name="connectionType"></param>
  ''' <param name="connectionAddr"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal deviceId As Integer, ByVal deviceType As String, ByVal connectionType As String, ByVal connectionAddr As String)

    MyBase.New()

    Dim strMessage As String = ""

    Me.InterfaceType = "TEMP08"
    Me.InterfaceID = connectionAddr

    m_DeviceId = deviceId
    m_DeviceType = deviceType

    m_DeviceConnectionType = connectionType
    m_DeviceConnectionAddr = connectionAddr

    '
    ' Start the TEMP08 command queue thread
    '
    CheckTEMP08Thread = New Thread(New ThreadStart(AddressOf ProcessTEMP08CommandQueue))
    CheckTEMP08Thread.Name = "CheckTEMP08"
    CheckTEMP08Thread.Start()
    WriteMessage(String.Format("{0} Thread Started", CheckTEMP08Thread.Name), MessageType.Debug)

    '
    ' Start the watchdog thread
    '
    WatchdogThread = New Thread(New ThreadStart(AddressOf TEMP08WatchdogThread))
    WatchdogThread.Name = "Watchdog"
    WatchdogThread.Start()

    WriteMessage(String.Format("{0} Thread Started", WatchdogThread.Name), MessageType.Debug)

  End Sub

  ''' <summary>
  ''' Distroy Object
  ''' </summary>
  ''' <remarks></remarks>
  Protected Overrides Sub Finalize()

    Try

      '
      ' Abort WatchdogThread
      '
      If WatchdogThread.IsAlive = True Then
        WatchdogThread.Abort()
      End If

      '
      ' Disconnect
      '
      Disconnect()

    Catch pEx As Exception

    End Try

    MyBase.Finalize()

  End Sub

#End Region

#Region "TEMP08 Device Connection"

  ''' <summary>
  ''' Initialize the connection to TEMP08 Device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConnectToDevice() As String

    Dim strMessage As String = ""
    Dim strPortName As String = ""
    Dim strPortAddr As Integer = 0

    strMessage = "Entered ConnectToDevice() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Get the Connection Information
      '
      Dim TEMP08Device As OneWireDevice = Get1WireDevice(Me.m_DeviceId)

      m_DeviceConnectionType = TEMP08Device.device_conn
      m_DeviceConnectionAddr = TEMP08Device.device_addr

      Select Case m_DeviceConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Try connection via the TEMP08
          '
          strMessage = String.Format("Initiating TEMP08 device '{0}' connection ...", m_DeviceConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Debug)

          '
          ' Inititalize Ethernet connection
          '
          Dim regexPattern As String = "(?<ipaddr>.+):(?<port>\d+)"
          If Regex.IsMatch(m_DeviceConnectionAddr, regexPattern) = True Then

            Dim ip_addr As String = Regex.Match(m_DeviceConnectionAddr, regexPattern).Groups("ipaddr").ToString()
            Dim ip_port As String = Regex.Match(m_DeviceConnectionAddr, regexPattern).Groups("port").ToString()

            gDeviceConnected = ConnectToEthernet(ip_addr, ip_port)
            If gDeviceInitialized = False Then
              gDeviceInitialized = gDeviceConnected
            End If

          Else
            '
            ' Unable to connect
            '
            gDeviceConnected = False

          End If

          If gDeviceConnected = False Then
            Throw New Exception(String.Format("Unable to connect to TEMP08 device '{0}'.", m_DeviceConnectionAddr))
          End If

        Case "SERIAL"
          '
          ' Try connecting to the serial port
          '
          Dim strComPort As String = m_DeviceConnectionAddr.Replace(":", "").ToUpper

          strMessage = String.Format("Initiating TEMP08 device '{0}' connection ...", m_DeviceConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Debug)

          '
          ' Close port if already open
          '
          If serialPort.IsOpen Then
            serialPort.Close()
          End If

          Try
            With serialPort
              .PortName = strComPort
              .BaudRate = 9600
              .Parity = IO.Ports.Parity.None
              .DataBits = 8
              .StopBits = IO.Ports.StopBits.One
              .ReadTimeout = 100
            End With
            serialPort.Open()
            serialPort.DiscardInBuffer()

            gDeviceConnected = serialPort.IsOpen
            If gDeviceInitialized = False Then
              gDeviceInitialized = gDeviceConnected
            End If

          Catch pEx As Exception
            Throw New Exception(String.Format("Unable to connect to TEMP08 device '{0}' because '{1}'.", m_DeviceConnectionAddr, pEx.ToString))
          End Try

        Case Else
          '
          ' Bail out when no port is defined (user has not set a port yet)
          '
          strMessage = String.Format("TEMP08 '{0}' interface is disabled.", Me.m_DeviceConnectionAddr)
          Call WriteMessage(strMessage, MessageType.Warning)

          '
          ' Unable to connect because the Interface is disabled
          '
          Throw New Exception(strMessage)

      End Select

      '
      ' Initialize the TEMP08
      '
      Call InitTEMP08()

      '
      ' We are connected here
      '
      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      WriteMessage(pEx.Message, MessageType.Error)
      gDeviceConnected = False
      Return pEx.ToString
    Finally
      '
      ' Update the TEMP08 connection status
      '
      'UpdateECMConnectionDevice(Me.m_DeviceConnectionType, Me.m_DeviceId, gDeviceConnected)
    End Try

  End Function

  ''' <summary>
  ''' Disconnect the connection to the TEMP08
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function DisconnectFromDevice() As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered DisconnectFromDevice() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      Select Case m_DeviceConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Close the ethernet connection
          '
          DisconnectEthernet()
        Case "SERIAL"
          '
          ' Close the serial port
          '
          If serialPort.IsOpen Then
            serialPort.Close()
          End If
      End Select

      '
      ' Reset Global Variables
      '
      gDeviceResponse = False
      gDeviceConnected = False

      Return True

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "DisconnectFromDevice()")
      Return False
    Finally
      '
      ' Update the TEMP08 connection status
      '
      'UpdateECMConnectionDevice(Me.m_DeviceConnectionType, Me.m_DeviceId, gDeviceConnected)
    End Try

  End Function

  ''' <summary>
  ''' Reconnect to the TEMP08 Device
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Reconnect()

    '
    ' Ensure plug-in is disconnected
    '
    DisconnectFromDevice()

    '
    ' Ensure watchdog is not disabled
    '
    m_WatchdogDisabled = False

    '
    ' Ensure TEMP08 is marked as initialized
    '
    gDeviceInitialized = True

    '
    ' Interrupt the watchdog thread
    '
    If WatchdogThread.ThreadState = ThreadState.WaitSleepJoin Then
      If m_WatchdogActive = False Then
        WatchdogThread.Interrupt()
      End If
    End If

  End Sub

  ''' <summary>
  ''' Disconnect from the TEMP08 Device
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub Disconnect()

    '
    ' Ensure the watchdog is disabled
    '
    m_WatchdogDisabled = True

    '
    ' Disconnect from the TEMP08 Device
    '
    DisconnectFromDevice()

  End Sub

#End Region

#Region "HSPI - Watchdog"

  ''' <summary>
  ''' TEMP08 Watchdog Thread
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub TEMP08WatchdogThread()

    Dim strMessage As String = ""
    Dim bAbortThread As Boolean = False
    Dim dblTimerInterval As Single = 1000 * 30
    Dim iSeconds As Long = 0

    Try
      '
      ' Stay in TEMP08WatchdogThread for duration of program
      '
      While bAbortThread = False

        Try

          If m_WatchdogDisabled = True Then

            dblTimerInterval = 1000 * 60

            strMessage = String.Format("Watchdog Timer indicates the TEMP08 device '{0}' auto reconnect is disabled.", m_DeviceConnectionAddr)
            WriteMessage(strMessage, MessageType.Debug)

          ElseIf gDeviceInitialized = True Then

            If IsDate(m_LastReport) Then
              iSeconds = DateDiff(DateInterval.Second, m_LastReport, DateTime.Now)
            End If

            strMessage = String.Format("Watchdog Timer indicates a response from the TEMP08 device '{0}' was received at {1}.", m_DeviceConnectionAddr, m_LastReport.ToString)
            WriteMessage(strMessage, MessageType.Debug)

            '
            ' Test to see if we are connected and that we have received a response within the past 300 seconds
            '
            Call CheckPhysicalConnection()

            If iSeconds > 300 Or m_WatchdogActive = True Or gDeviceConnected = False Then
              '
              ' Action for initial watchdog trigger
              '
              If m_WatchdogActive = False Then
                m_WatchdogActive = True
                dblTimerInterval = 1000 * 30

                Dim strWatchdogReason As String = String.Format("No response response from the TEMP08 device '{0}' for {1} seconds.", m_DeviceConnectionAddr, iSeconds)
                If gDeviceConnected = False Then
                  strWatchdogReason = String.Format("Connection to TEMP08 device '{0}' was lost", m_DeviceConnectionAddr)
                End If

                strMessage = String.Format("Watchdog Timer indicates {0}.  Attempting to reconnect ...", strWatchdogReason)
                WriteMessage(strMessage, MessageType.Warning)

                '
                ' Check watchdog trigger
                '
                Dim strTrigger As String = IFACE_NAME & Chr(2) & "TEMP08 Watchdog Trigger" & Chr(2) & "Connection Failure" & Chr(2) & "*"
                'callback.CheckTrigger(strTrigger)
              End If

              '
              ' Ensure everything is closed properly and attempt a reconnect
              '
              Call DisconnectFromDevice()
              Call ConnectToDevice()

              If gDeviceConnected = False Then

                WriteMessage("Watchdog Timer reconnect attempt failed.", MessageType.Warning)

                dblTimerInterval *= 2
                If dblTimerInterval > 3600000 Then
                  dblTimerInterval = 3600000
                End If

              Else

                WriteMessage("Watchdog Timer reconnect attempt succeeded.", MessageType.Informational)
                m_WatchdogActive = False
                dblTimerInterval = 1000 * 30

                '
                ' Check watchdog trigger
                '
                Dim strTrigger As String = IFACE_NAME & Chr(2) & "TEMP08 Watchdog Trigger" & Chr(2) & "Connection Restore" & Chr(2) & "*"
                'callback.CheckTrigger(strTrigger)

                ' Was a check to power status
                Call CheckPhysicalConnection()

              End If

            Else
              '
              ' Plug-in is connected to the Global Cache device
              '
              m_WatchdogActive = False
              dblTimerInterval = 1000 * 30

              strMessage = String.Format("Watchdog Timer indicates a response from the TEMP08 device '{0}' was received {1} seconds ago.", m_DeviceConnectionAddr, iSeconds.ToString)
              WriteMessage(strMessage, MessageType.Debug)

              ' Was a check to power status
              Call CheckPhysicalConnection()

            End If

          End If

          '
          ' Sleep Watchdog Thread
          '
          strMessage = String.Format("Watchdog Timer thread for the TEMP08 device '{0}' sleeping for {1}.", m_DeviceConnectionAddr, dblTimerInterval.ToString)
          WriteMessage(strMessage, MessageType.Debug)

          Thread.Sleep(dblTimerInterval)

        Catch pEx As ThreadInterruptedException
          '
          ' Thread sleep was interrupted
          '
          gDeviceInitialized = True
          strMessage = String.Format("Watchdog Timer thread for the TEMP08 device '{0}' was interrupted.", m_DeviceConnectionAddr, iSeconds.ToString)
          WriteMessage(strMessage, MessageType.Debug)

        Catch pEx As Exception
          '
          ' Process Exception
          '
          Call ProcessError(pEx, "TEMP08WatchdogThread()")
        End Try

      End While ' Stay in thread until we get an abort/exit request

    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      bAbortThread = True      ' Not actually needed
      Call WriteMessage("Abort requested on TEMP08WatchdogThread", MessageType.Debug)
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "TEMP08WatchdogThread()")
    Finally

    End Try

  End Sub

#End Region

  '------------------------------------------------------------------------------------
  'Purpose:   Subroutine to initialize TEMP08 interface
  'Input:     port as Integer
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Function InitTEMP08() As String

    Dim strMessage As String = ""

    strMessage = "Entered InitTEMP08() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      '
      ' Begin TEMP08 Initialization
      '
      If gDeviceConnected = True Then

        strMessage = "Setting TEMP08 Clock ..."
        WriteMessage(strMessage, MessageType.Informational)

        Dim MyDate As DateTime = DateTime.Now
        Dim MyDOW As String = (MyDate.DayOfWeek + 1).ToString.PadLeft(2, "0")
        Dim MyHour As String = MyDate.Hour.ToString.PadLeft(2, "0")
        Dim MyMin As String = MyDate.Minute.ToString.PadLeft(2, "0")
        Dim MySec As String = MyDate.Second.ToString.PadLeft(2, "0")

        AddCommand(String.Format("{0} {1}, {2}, {3}, {4}", "SCK", MyDOW, MyHour, MyMin, MySec) & vbCrLf)

        strMessage = "Enabling TEMP08 Serial Number ID display ..."
        WriteMessage(strMessage, MessageType.Informational)
        AddCommand("SIDon")

        strMessage = "Setting TEMP08 temperature format ..."
        WriteMessage(strMessage, MessageType.Informational)
        AddCommand("STDC")

        strMessage = "Disabling TEMP08 update interval ..."
        WriteMessage(strMessage, MessageType.Informational)
        AddCommand("SPT00")

        strMessage = "Getting TEMP08 version ..."
        WriteMessage(strMessage, MessageType.Informational)
        AddCommand("VER")
      End If

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "InitTEMP08()")
      Return pEx.ToString
    End Try

  End Function

#Region "TEMP08 Command Processing"

  '------------------------------------------------------------------------------------
  'Purpose:   Adds command to command buffer for processing
  'Input:     strCommand as String, bForce as Boolean
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub AddCommand(ByVal strCommand As String, Optional ByVal bForce As Boolean = False)

    '
    ' bForce may be used to add a command to speak the same word/phrase more than once
    '
    If CommandQueue.Contains(strCommand) = False Or bForce = True Then
      CommandQueue.Enqueue(strCommand)
    End If

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Processes commands and waits for the response
  'Input:     None
  'Output:    None
  '------------------------------------------------------------------------------------
  Public Sub ProcessTEMP08CommandQueue()

    Dim bAbortThread As Boolean = False
    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim iCmdAttempt As Integer = 0
    Dim strCommand As String = ""
    Dim iMaxWaitTime As Single = 0

    Try

      While bAbortThread = False
        '
        ' Don't send a command if the TEMP08 won't reply
        '
        If CommandQueue.Count > 0 And gDeviceConnected = False Then
          '
          ' Clear out the existing commands
          '
          CommandQueue.Clear()
        Else

          '
          ' Process commands in command queue
          '
          While CommandQueue.Count > 0 And gDeviceConnected = True And gIOEnabled = True
            '
            ' Set the command response we are waiting for
            '
            strCommand = CommandQueue.Peek
            iMaxWaitTime = 0

            '
            ' Determine if we need to wait for the command response
            '
            If Regex.IsMatch(strCommand, "^a[0-9]") = True Then
              strCmdWait = ""
              iMaxWaitTime = MAX_WAIT_TIME
            Else
              strCmdWait = ""
              iMaxWaitTime = MAX_WAIT_TIME
            End If

            '
            ' Increment the counter
            '
            iCmdAttempt += 1
            WriteMessage(String.Format("Sending command: '{0}' to TEMP08, attempt # {1}", strCommand, iCmdAttempt), MessageType.Debug)
            SendDataToTEMP08(strCommand)

            '
            ' Determine if we need to wait for a response
            '
            If iMaxWaitTime > 0 And strCmdWait.Length > 0 Then
              '
              ' A response to our command is expected, so lets wait for it
              '
              WriteMessage(String.Format("Waiting for the TEMP08 to respond with '{0}' for up to {1} seconds...", strCmdWait, iMaxWaitTime), MessageType.Debug)

              '
              ' Keep track of when we started waiting for the response
              '
              dtStartTime = Now

              '
              '  Wait for the proper response to come back, or the maximum wait time
              '
              Do
                '
                ' Sleep this thread for 50ms giving the receive function time to get the response
                '
                Thread.Sleep(50)
                '
                ' Find out how long we have been waiting in total
                '
                etElapsedTime = Now.Subtract(dtStartTime)
                iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)

                '
                ' Loop until the expected command was received (strCmdWait is cleared) or we ran past the maximum wait time
                '
              Loop Until strCmdWait.Length = 0 Or iMillisecondsWaited > iMaxWaitTime * 1000 ' Now abort if the command was recieved or we ran out of time

              WriteMessage(String.Format("Waited {0} milliseconds for the command response.", iMillisecondsWaited), MessageType.Debug)

              If strCmdWait.Length > 0 Or iCmdAttempt > MAX_ATTEMPTS Then
                '
                ' Command failed, so lets stop trying to send this commmand
                '
                If strCmdWait <> "CC" Then
                  WriteMessage(String.Format("No response/improper response from TEMP08 to command '{0}'", strCommand), MessageType.Warning)
                End If
                strCmdWait = String.Empty

                '
                ' Only Dequeue the command if we have tried more than MAX_ATTEMPTS times
                '
                If iCmdAttempt > MAX_ATTEMPTS Then
                  CommandQueue.Dequeue()
                  strCmdWait = String.Empty
                  iCmdAttempt = 0
                End If
              Else
                CommandQueue.Dequeue()
                iCmdAttempt = 0
              End If
            Else
              '
              ' No response expected, so remove command from queue
              '
              WriteMessage(String.Format("Command {0} does not produce a result.", strCommand), MessageType.Debug)
              CommandQueue.Dequeue()
              strCmdWait = String.Empty
              iCmdAttempt = 0
            End If

            Thread.Sleep(2000)

          End While ' Done with all commands in queue

        End If

        '
        ' Give up some time to allow the main thread to populate the command queue with more commands
        '
        Thread.Sleep(50)

      End While ' Stay in thread until we get an abort/exit request

    Catch pEx As ThreadAbortException
      ' 
      ' There was a normal request to terminate the thread.  
      '
      bAbortThread = True      ' Not actually needed
      WriteMessage(String.Format("ProcessTEMP08CommandQueue thread received abort request, terminating normally."), MessageType.Debug)

    Catch pEx As Exception
      '
      ' Return message
      '
      ProcessError(pEx, "ProcessTEMP08CommandQueue()")

    Finally
      '
      ' Notify that we are exiting the thread 
      '
      WriteMessage(String.Format("ProcessTEMP08CommandQueue terminated."), MessageType.Debug)

    End Try

  End Sub

#End Region

#Region "TEMP08 Protocol Processing"

  ''' <summary>
  ''' Sends command to TEMP08
  ''' </summary>
  ''' <param name="strDataToSend"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SendDataToTEMP08(ByVal strDataToSend As String) As Boolean

    Dim strMessage As String = ""

    strMessage = "Entered SendToTEMP08() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Format packet
      '
      Dim strPacket As String = FormatDataPacket(strDataToSend)

      Select Case m_DeviceConnectionType.ToUpper
        Case "ETHERNET"
          '
          ' Set data to Ethernet connection
          '
          If gDeviceConnected = True And gIOEnabled = True Then
            strMessage = String.Format("Sending '{0}' to TEMP08 device '{1}' via Ethernet.", strPacket, m_DeviceConnectionAddr)
            Call WriteMessage(strMessage, MessageType.Debug)
            Return SendMessageToEthernet(strPacket)
          Else
            Return False
          End If

        Case "SERIAL"
          '
          ' Send data using the serial port
          '
          If serialPort.IsOpen = True Then
            strMessage = String.Format("Sending '{0}' to TEMP08 device '{1}' via Serial.", strPacket, m_DeviceConnectionAddr)
            Call WriteMessage(strMessage, MessageType.Debug)

            serialPort.Write(strPacket)
          End If
          Return serialPort.IsOpen

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SendDataToTEMP08()")
      Return False
    End Try

    Return False

  End Function

  '------------------------------------------------------------------------------------
  'Purpose:   Formats data packet before sending to the TEMP08
  'Input:     strData as String
  'Output:    String
  '------------------------------------------------------------------------------------
  Function FormatDataPacket(ByVal strData As String) As String
    Return strData
  End Function

  '------------------------------------------------------------------------------------
  'Purpose:   Event triggered when data is available from com port
  'Input:     Multiple, see below
  'Output:    None
  '------------------------------------------------------------------------------------
  Private Sub ProcessReceived(ByVal strDataRec As String)

    Dim strMessage As String = ""

    strMessage = "Entered TEMP08 ProcessReceived() thread subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try

      strMessage = "TEMP08 Data Received:  " & strDataRec
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Update the last date/time response
      '
      m_LastReport = DateTime.Now

      ' Voltage #01[060000003770D026]=03.13V 05.08V 00mV
      ' Quad Voltage #01[3000000002202920]=05.05V,05.04V,03.18V,03.10V
      ' Wind Speed[9F0000000335111D]=10 MPH, Gust = 10
      ' Rain #01[8E0000000115111D]=02.56 Inch
      ' Counter #01[3D0000000C8C351D]=08957 06228
      ' 1WIO #01[A400000001042829]=Off,Off,Off,Off
      ' Temp #01[91000800135B9B10]=76.55F

      '
      ' Process Temp08 Sensors
      ' 
      If Regex.IsMatch(strDataRec, "(Temp|Sensor|Humidity|Barometer) #") Then
        Call ProcessOneWireSensor(strDataRec)
      ElseIf Regex.IsMatch(strDataRec, "(Counter|Rain) #") Then
        ProcessOneWireCounter(strDataRec)
      ElseIf Regex.IsMatch(strDataRec, "Switch #") Then
        ProcessOneWireSwitch(strDataRec)
      ElseIf Regex.IsMatch(strDataRec, "Voltage #") Then
        ProcessVoltageSensor(strDataRec)
      ElseIf Regex.IsMatch(strDataRec, "Wind (Dirn|Speed") Then
        'ProcessWindSensor(strDataRec)
      End If

      Thread.Sleep(0)

    Catch pEx As Exception
      '
      ' Nothing to do here
      '
    End Try

  End Sub

#End Region

#Region "TEMP08 Sensor Processing"

  '-------------------------------------------------------------------------------
  ' Purpose:   Checks TEMP08 sensors
  ' Input:     None
  ' Output:    None
  '-------------------------------------------------------------------------------
  Sub CheckOneWireSensors()

    Dim strMessage As String = ""

    strMessage = "Entered CheckOneWireSensors() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Begin Sensor Reading
      '
      AddCommand("TMP")
      Thread.Sleep(1000 * 15)

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "CheckOneWireSensors()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process Temp08 temperature sensor
  'Input:     strRecData as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Sub ProcessOneWireSensor(ByVal strRecData As String)

    Dim colMatches As MatchCollection
    Dim strMessage As String = ""

    strMessage = "Entered ProcessOneWireSensor() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Process 1-wire sensors
      '
      Dim bFarenheight As Boolean = IIf(gUnitType = "Metric", False, True)

      ' Note:  Temp05 uses "Sensor" for a temperature probe)
      colMatches = Regex.Matches(strRecData, "(Temp|Sensor) #(?<sensorid>\d\d)\[(?<romid>.*)\]=(?<value>.*)(F|C)")
      For Each objMatch As Match In colMatches

        Dim SensorClass As String = "Environmental"
        Dim SensorType As String = "Temperature"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim Temperature As String = objMatch.Groups("value").Value

        Dim Value As Single = ConvertTemperature(Temperature, bFarenheight)
        UpdateOneWireSensor(m_DeviceId, SensorClass, SensorType, "A", ROMId, Value)

      Next

      colMatches = Regex.Matches(strRecData, "Humidity #(?<sensorid>\d\d)\[(?<romid>.*)\]=>?(?<value>.*)%")
      For Each objMatch As Match In colMatches

        Dim SensorClass As String = "Environmental"
        Dim SensorType As String = "Humidity"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim strValue As String = objMatch.Groups("value").Value

        Dim Value As Single = Math.Abs(Single.Parse(strValue, nfi))
        UpdateOneWireSensor(m_DeviceId, SensorClass, SensorType, "A", ROMId, Value)

      Next

      colMatches = Regex.Matches(strRecData, "Barometer #(?<sensorid>\d\d)\[(?<romid>.*)\]=(?<value>.*)\sinHg")
      For Each objMatch As Match In colMatches

        Dim SensorClass As String = "Environmental"
        Dim SensorType As String = "Pressure"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim strValue As String = objMatch.Groups("value").Value

        Dim Value As Single = Math.Abs(Single.Parse(strValue, nfi))
        UpdateOneWireSensor(m_DeviceId, SensorClass, SensorType, "A", ROMId, Value)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ProcessOneWireSensor()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process Temp08 Counter sensor
  'Input:     strRecData as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Sub ProcessOneWireCounter(ByVal strRecData As String)

    Dim colMatches As MatchCollection
    Dim strMessage As String = ""

    strMessage = "Entered ProcessOneWireCounter() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Process 1-wire Generic Counter
      '
      ' TEMP08 Data Received: Counter #01[1F0000000770011D]=00000 00000 
      colMatches = Regex.Matches(strRecData, "Counter #(?<sensorid>\d\d)\[(?<romid>.*)\]=(?<value1>\d+)\s(?<value2>\d+)")
      For Each objMatch As Match In colMatches

        Dim SensorType As String = "Counter"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim strValueA As String = objMatch.Groups("value1").Value
        Dim strValueB As String = objMatch.Groups("value2").Value

        Dim ValueA As Integer = Math.Abs(Integer.Parse(strValueA, nfi))
        UpdateOneWireSensor(m_DeviceId, "Counter", SensorType, "A", ROMId, ValueA)

        Dim ValueB As Integer = Math.Abs(Integer.Parse(strValueB, nfi))
        UpdateOneWireSensor(m_DeviceId, "Counter", SensorType, "B", ROMId, ValueB)

      Next

      '
      ' Process 1-wire Rain Sensor
      '
      ' Rain #01[8E0000000115111D]=02.56 Inch
      colMatches = Regex.Matches(strRecData, "Rain #(?<sensorid>\d\d)\[(?<romid>.*)\]=(?<value1>\d+\.\d+)")
      For Each objMatch As Match In colMatches

        Dim SensorType As String = "Counter"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim strValueA As String = objMatch.Groups("value1").Value

        Dim Value As Single = Math.Abs(Single.Parse(strValueA, nfi))
        Dim ValueA As Integer = Value / 0.01
        UpdateOneWireSensor(m_DeviceId, "Counter", SensorType, "A", ROMId, ValueA)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ProcessOneWireSensor()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process Temp08 Switch sensor
  'Input:     strRecData as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Sub ProcessOneWireSwitch(ByVal strRecData As String)

    Dim colMatches As MatchCollection
    Dim strMessage As String = ""

    strMessage = "Entered ProcessOneWireSwitch() subroutine."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Process 1-wire Switch
      ' TEMP08 Data Received: Switch #04[D6000000221D1512]=Off
      colMatches = Regex.Matches(strRecData, "Switch #(?<sensorid>\d\d)\[(?<romid>.*)\]=?<value>(On|Off)")
      For Each objMatch As Match In colMatches

        Dim SensorType As String = "Switch"
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value

        Dim Value As Byte = IIf(objMatch.Groups("value").Value = "On", 1, 0)
        UpdateOneWireSensor(m_DeviceId, "Switch", SensorType, "A", ROMId, Value)

      Next

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "ProcessOneWireSensor()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Subroutine to update a voltage Temp08 Sensor
  'Input:     strRecData as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Private Sub ProcessVoltageSensor(ByVal strRecData As String)

    Dim colMatches As MatchCollection
    Dim strMessage As String = ""

    Try
      '
      ' Write the debug message
      '
      strMessage = "Entered ProcessMiscSensor() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Process misc sensors
      '
      'Voltage #01[060000003770D026]=03.13V 05.08V 00mV
      colMatches = Regex.Matches(strRecData, "(?<type>Voltage) #(?<sensorid>\d\d)\[(?<romid>.*)\]=(?<value>\d+\.\d+)V")
      For Each objMatch As Match In colMatches

        Dim SensorType As String = objMatch.Groups("type").Value
        Dim strSensorID As String = objMatch.Groups("sensorid").Value
        Dim ROMId As String = objMatch.Groups("romid").Value
        Dim Value As String = objMatch.Groups("value").Value

        UpdateOneWireSensor(m_DeviceId, "Voltage", SensorType, "A", ROMId, Value)

      Next

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "ProcessMiscSensor()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Process Temp08 Wind Sensor
  'Input:     strRecData as String
  'Output:    None
  '------------------------------------------------------------------------------------
  Private Sub ProcessWindSensor(ByVal strRecData As String)

    Dim colMatches As MatchCollection
    Dim strMessage As String = ""

    Try
      '
      ' Write debug message
      '
      strMessage = "Entered ProcessWindSensor() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Process Wind sensors
      '
      colMatches = Regex.Matches(strRecData, "(?<type>Wind Dirn|Wind Speed)\[(?<romid>.*)\]=(?<value>.*)")
      For Each objMatch As Match In colMatches

        Dim strSensorType As String = objMatch.Groups("type").Value
        Dim strSensorID As String = "01"
        Dim strSensorRomID As String = objMatch.Groups("romid").Value
        Dim strSensorValue As String = objMatch.Groups("value").Value

      Next

    Catch pEx As Exception
      '
      ' Process Program Exception
      '
      Call ProcessError(pEx, "ProcessWindSensor()")
    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Converts temperature value
  'Input:     Value as String, Farenheight As Boolean
  'Output:    None
  '------------------------------------------------------------------------------------
  Private Function ConvertTemperature(ByVal value As String, ByVal Farenheight As Boolean) As Single

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

#End Region

#Region "TEMP08 Connection Status"

  ''' <summary>
  ''' Checks to see if we are connected to the TEMP08
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CheckTEMP08Connection() As Boolean

    Dim dtStartTime As Date
    Dim etElapsedTime As TimeSpan
    Dim iMillisecondsWaited As Integer
    Dim Attempts As Integer = 0

    Try
      '
      ' We are going to use this function as a verification test to see if the TEMP08 is actually connected
      '
      Call WriteMessage(String.Format("Checking if the TEMP08 device '{0}' is connected ...", m_DeviceConnectionAddr), MessageType.Debug)

      If gDeviceConnected = False Then
        Return False
      End If

      '
      ' Reset global variable
      '
      gDeviceResponse = False

      Do
        '
        ' Block for until we get our TEMP08 response
        '
        dtStartTime = Now
        Do
          Thread.Sleep(50)

          If IsDate(m_LastReport) Then
            Dim iSeconds As Long = DateDiff(DateInterval.Second, m_LastReport, DateTime.Now)
            If iSeconds <= 30 Then gDeviceResponse = True
          End If

          etElapsedTime = Now.Subtract(dtStartTime)
          iMillisecondsWaited = etElapsedTime.Milliseconds + (etElapsedTime.Seconds * 1000)
        Loop While gDeviceResponse = False And iMillisecondsWaited < 3000

        If gDeviceResponse = False Then
          Attempts = Attempts + 1
        End If

      Loop While Attempts < 3 And gDeviceResponse = False

      Return gDeviceResponse

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CheckTEMP08Connection()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Determine if TEMP08 is active
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub CheckPhysicalConnection()

    Try

      Select Case m_DeviceConnectionType.ToUpper
        Case "ETHERNET"
          If TcpClient.Connected = False Or TcpClient.Client.Connected = False Then
            gDeviceConnected = False
          Else

            If NetworkStream.CanRead = False Or NetworkStream.CanWrite = False Then
              gDeviceConnected = False
            End If

          End If
        Case "SERIAL"
          gDeviceConnected = serialPort.IsOpen
      End Select

    Catch pEx As Exception

    End Try

  End Sub

#End Region

#Region "Serial Support"

  ''' <summary>
  ''' Event handler for Com Port data received
  ''' </summary>
  ''' <param name="sender"></param>
  ''' <param name="e"></param>
  ''' <remarks></remarks>
  Private Sub DataReceived(ByVal sender As Object, _
                           ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) _
                           Handles serialPort.DataReceived

    Dim strMessage As String = "Entered Serial DataReceived() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    '
    ' Read from com port and buffer till we get a vbLf
    '
    Try
      Dim Str As New StringBuilder

      Do
        Do

          Dim By As Integer = serialPort.ReadByte()
          Dim Dat As Char = Chr(By)
          If Dat = vbLf Then
            If Str.Length > 2 Then
              ProcessReceived(Str.ToString)
            End If
            Str.Length = 0
          Else
            Str.Append(Dat)
            Thread.Sleep(20)
          End If

        Loop While serialPort.BytesToRead > 0

        Thread.Sleep(1000)
      Loop While serialPort.BytesToRead > 0

    Catch pEx As Exception
      WriteMessage(pEx.ToString, MessageType.Error)
    End Try

  End Sub

#End Region

#Region "Ethernet Support"

  Dim NetworkStream As NetworkStream
  Dim TCPListener As TcpListener
  Dim TcpClient As TcpClient

  Dim ReadThread As Threading.Thread

  ''' <summary>
  ''' Establish connection to Ethernet
  ''' </summary>
  ''' <param name="Ip"></param>
  ''' <param name="Port"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConnectToEthernet(ByVal Ip As String, ByVal Port As Integer) As Boolean

    Dim strMessage As String = ""

    Try

      Dim IPAddress As String = ResolveAddress(Ip)

      If IPAddress = "0.0.0.0" Then
        '
        ' Start TCP Listener
        '
        TCPListener = New TcpListener(Net.IPAddress.Any, Port)
        TCPListener.Start(1)

        strMessage = String.Format("TEMP08 Device listening on TCP Port {0} ...", Port.ToString)
        Call WriteMessage(strMessage, MessageType.Informational)

        '
        ' Create TCPClient
        '
        TcpClient = TCPListener.AcceptTcpClient()

        strMessage = String.Format("TEMP08 Device connection accepted on TCP Port {0}", Port.ToString)
        Call WriteMessage(strMessage, MessageType.Informational)

        NetworkStream = TcpClient.GetStream()
        ReadThread = New Thread(New ThreadStart(AddressOf EthernetReadThreadProc))
        ReadThread.Name = "EthernetReadThreadProc"
        ReadThread.IsBackground = True
        ReadThread.Start()

      Else

        Try
          '
          ' Create TCPClient
          '
          TcpClient = New TcpClient(IPAddress, Port)

          strMessage = String.Format("TEMP08 Device connection to {0} ({1}:{2})", IPAddress, Ip, Port.ToString)
          Call WriteMessage(strMessage, MessageType.Informational)

        Catch pEx As SocketException
          '
          ' Process Exception
          '
          strMessage = String.Format("Ethernet connection could not be made to {0} ({1}:{2}) - {3}", _
                                    IPAddress, Ip.ToString, Port.ToString, pEx.Message)
          Call WriteMessage(strMessage, MessageType.Debug)
          Return False
        End Try

        NetworkStream = TcpClient.GetStream()
        ReadThread = New Thread(New ThreadStart(AddressOf EthernetReadThreadProc))
        ReadThread.Name = "EthernetReadThreadProc"
        ReadThread.Start()

      End If

      Return True
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "ConnectToEthernet()")
      Return False
    End Try

    Return True

  End Function

  ''' <summary>
  ''' Check ip string to be an ip address or if not try to resolve using DNS
  ''' </summary>
  ''' <param name="hostNameOrAddress"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function ResolveAddress(ByVal hostNameOrAddress As String) As String

    Try
      '
      ' Attempt to identify fqdn as an IP address
      '
      IPAddress.Parse(hostNameOrAddress)

      '
      ' If this did not throw then it is a valid IP address
      '
      Return hostNameOrAddress
    Catch ex As Exception
      Try
        ' Try to resolve it through DNS if it is not in IP address form
        ' and use the first IP address if defined as round robbin in DNS
        Dim ipAddress As IPAddress = Dns.GetHostEntry(hostNameOrAddress).AddressList(0)

        Return ipAddress.ToString
      Catch pEx As Exception
        Return ""
      End Try

    End Try

  End Function

  ''' <summary>
  ''' Disconnection From Ethernet
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub DisconnectEthernet()

    Try

      If ReadThread.IsAlive = True Then
        ReadThread.Abort()
      End If

      NetworkStream.Close()
      TcpClient.Close()
      TCPListener.Stop()

    Catch pEx As Exception
      '
      ' Ignore Exception
      '
    End Try

  End Sub

  ''' <summary>
  ''' Send Message to connected IP address (first send buffer length and then the buffer holding message)
  ''' </summary>
  ''' <param name="Message"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Protected Function SendMessageToEthernet(ByVal Message As String) As Boolean

    Try

      Dim Buffer() As Byte = Encoding.ASCII.GetBytes(Message.ToCharArray)

      If TcpClient.Connected = True Then
        NetworkStream.Write(Buffer, 0, Buffer.Length)
        Return True
      Else
        Call WriteMessage("Attempted to write to a closed ethernet stream in SendMessageToEthernet()", MessageType.Warning)
        Return False
      End If

    Catch ex As Exception
      Call ProcessError(ex, "SendMessageToEthernet()")
      Return False
    End Try

  End Function

  ''' <summary>
  ''' Process to Read Data From TCP Client
  ''' </summary>
  ''' <remarks></remarks>
  Protected Sub EthernetReadThreadProc()

    Try

      Dim r As New BinaryReader(NetworkStream)
      Dim Str As New StringBuilder

      Dim strMessage As String = "Entered EthernetReadThreadProc() subroutine."
      WriteMessage(strMessage, MessageType.Debug)

      '
      ' Stay in EthernetReadThreadProc while client is connected
      '
      Do While TcpClient.Connected = True

        Do While NetworkStream.DataAvailable = True

          Dim By As Integer = r.ReadByte()
          Dim Dat As Char = Chr(By)
          If Dat = vbLf Then
            If Str.Length > 2 Then
              ProcessReceived(Str.ToString)
            End If
            Str.Length = 0
          Else
            Str.Append(Dat)
          End If

          Thread.Sleep(10)
        Loop

        Thread.Sleep(50)

      Loop

    Catch ab As ThreadAbortException
      '
      ' Process Thread Abort Exception
      '
      Call WriteMessage("Abort requested on EthernetReadThreadProc", MessageType.Debug)
      Return
    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "EthernetReadThreadProc()")
    Finally
      '
      ' Indicate we are no longer connected to the TEMP08 Device
      '
      gDeviceConnected = False
    End Try

  End Sub

#End Region

End Class


