Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service

Imports System.Reflection
Imports System.Reflection.Assembly
Imports System.Diagnostics.FileVersionInfo

Imports System.Text
Imports System.Threading

Public Class HSPI
  Inherits ScsService
  Implements IPlugInAPI               ' This API is required for ALL plugins

  '
  ' Web Config Global Variables
  '
  Dim WebConfigPage As hspi_webconfig
  Dim WebPage As Object

  Public OurInstanceFriendlyName As String = ""
  Public instance As String = ""

  '
  ' Define Class Global Variables
  '
  Public SensorHistoryThread As Thread
  Public DatabaseQueueThread As Thread
  Public CheckOneWireSensorsThread As Thread

#Region "HSPI - Base Plugin API"

  ''' <summary>
  ''' Probably one of the most important properties, the Name function in your plug-in is what the plug-in is identified with at all times.  
  ''' The filename of the plug-in is irrelevant other than when HomeSeer is searching for plug-in files, but the Name property is key to many things, 
  ''' including how plug-in created triggers and actions are stored by HomeSeer.  
  ''' If this property is changed from one version of the plug-in to the next, all triggers, actions, and devices created by the plug-in will have to be re-created by the user.  
  ''' Please try to keep the Name property value short, e.g. 14 to 16 characters or less.  Web pages, trigger and action forms created by your plug-in can use a longer, 
  ''' more elaborate name if you so desire.  In the sample plug-ins, the constant IFACE_NAME is commonly used in the program to return the name of the plug-in. 
  ''' No spaces or special characters are allowed other than a dash or underscore.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
    Get
      Return IFACE_NAME
    End Get
  End Property

  ''' <summary>
  ''' Return the API's that this plug-in supports. This is a bit field. All plug-ins must have bit 3 set for I/O.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
    Return HomeSeerAPI.Enums.eCapabilities.CA_IO
  End Function

  ''' <summary>
  ''' This determines whether the plug-in is free, or is a licensed plug-in using HomeSeer's licensing service.  
  ''' Return a value of 1 for a free plug-in, a value of 2 indicates that the plug-in is licensed using HomeSeer's licensing.
  ''' </summary>
  ''' <returns>Integer</returns>
  ''' <remarks></remarks>
  Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
    Return 2
  End Function

  ''' <summary>
  ''' Returns the instance name of this instance of the plug-in. Only valid if SupportsMultipleInstances returns TRUE. 
  ''' The instance is set when the plug-in is started, it is passed as a command line parameter. 
  ''' The initial instance name is set when a new instance is created on the HomeSeer interfaces page. 
  ''' A plug-in needs to associate this instance name with any local status that it is keeping for this instance. 
  ''' See the multiple instances section for more information.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName

    '
    ' Write the debug message
    '
    WriteMessage("Entered InstanceFriendlyName() function.", MessageType.Debug)

    Return Me.instance
  End Function

  ''' <summary>
  ''' Return TRUE if the plug-in supports multiple instances. 
  ''' The plug-in may be launched multiple times and will be passed a unique instance name as a command line parameter to the Main function. 
  ''' The plug-in then needs to associate all local status with this particular instance.
  ''' This feature is ideal for cases where multiple hardware modules need to be supported. 
  ''' For example, an single irrigation controller supports 8 zones but the user needs 16. 
  ''' They can add a second controller as a new instance to control 8 more zones. 
  ''' This assumes that the second controller would use a different COM port or IP address.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
    Return False
  End Function

  ''' <summary>
  ''' No Summary in API Docs
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstancesSingleEXE
    Return False
  End Function

  ''' <summary>
  ''' Your plug-in should return True if you wish to use HomeSeer's interfaces page of the configuration for the user to enter a serial port number for your plug-in to use.  
  ''' If enabled, HomeSeer will return this COM port number to your plug-in in the InitIO call.  
  ''' If you wish to have your own configuration UI for the serial port, or if your plug-in does not require a serial port, return False
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
    Get
      Return False
    End Get
  End Property

  ''' <summary>
  ''' SetIOMulti is called by HomeSeer when a device that your plug-in owns is controlled.  Your plug-in owns a device when it's INTERFACE property is set to the name of your plug-in.
  ''' The parameters passed to SetIOMulti are as follows - depending upon what generated the SetIO call, not all parameters will contain data.  
  ''' Be sure to test for "Is Nothing" before testing for values or your plug-in may generate an exception error when a variable passed is uninitialized. 
  ''' </summary>
  ''' <param name="colSend">This is a collection of CAPIControl objects, one object for each device that needs to be controlled. Look at the ControlValue property to get the value that device needs to be set to.</param>
  ''' <remarks></remarks>
  Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti

    '
    ' Write the debug message
    '
    WriteMessage(String.Format("Entered {0} {1}", "SetIOMulti", "Subroutine"), MessageType.Debug)

    Dim CC As CAPIControl
    For Each CC In colSend
      If CC Is Nothing Then Continue For

      WriteMessage(String.Format("SetIOMulti set value: {0}, type {1}, ref:{2}", CC.ControlValue.ToString, CC.ControlType.ToString, CC.Ref.ToString), MessageType.Debug)

      Try
        '
        ' Device exists, so lets get a reference to it
        '
        Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(CC.Ref)
        If dv Is Nothing Then Continue For

        '
        ' Get the device type
        '
        Dim dev_type As String = dv.Device_Type_String(hs)

        '
        ' Process the SetIO action
        '
        Select Case dev_type
          Case "Ultra1Wire3 Monitoring"
            Dim dev_value As Integer = Int32.Parse(CC.ControlValue)
            Select Case dev_value
              Case -1
                '
                ' Enable 1-Wire Checking
                '
                gMonitoring = True
                hs.SaveINISetting("Options", "Monitoring", gMonitoring.ToString, gINIFile)

                Dim dev_status As Integer = IIf(gMonitoring = True, 1, 0)
                hs.SetDeviceValueByRef(CC.Ref, dev_status, True)

                WriteMessage("1-Wire sensor checking is enabled.", MessageType.Informational)

              Case -2
                '
                ' Disable 1-Wire Checking
                '
                gMonitoring = False
                hs.SaveINISetting("Options", "Monitoring", gMonitoring.ToString, gINIFile)

                Dim dev_status As Integer = IIf(gMonitoring = True, 1, 0)
                hs.SetDeviceValueByRef(CC.Ref, dev_status, True)

                WriteMessage("1-Wire sensor checking is disabled.", MessageType.Informational)

            End Select

          Case Else
            '
            ' Write warning message
            '
            WriteMessage(String.Format("Received unsupported SetIO action for {0}", dev_type), MessageType.Warning)

        End Select

      Catch pEx As Exception
        '
        ' Process program exception
        '
        ProcessError(pEx, "SetIOMulti")
      End Try

    Next

  End Sub

  ''' <summary>
  ''' HomeSeer may call this function at any time to get the status of the plug-in. Normally it is displayed on the Interfaces page.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus

    Dim es As New IPlugInAPI.strInterfaceStatus
    es.intStatus = IPlugInAPI.enumInterfaceStatus.OK

    Return es

  End Function

  ''' <summary>
  ''' If your plugin is set to start when HomeSeer starts, or is enabled from the interfaces page, then this function will be called to initialize your plugin. 
  ''' If you returned TRUE from HSComPort then the port number as configured in HomeSeer will be passed to this function. 
  ''' Here you should initialize your plugin fully. The hs object is available to you to call the HomeSeer scripting API as well as the callback object so you can call into the HomeSeer plugin API.  
  ''' HomeSeer's startup routine waits for this function to return a result, so it is important that you try to exit this procedure quickly.  
  ''' If your hardware has a long initialization process, you can check the configuration in InitIO and if everything is set up correctly, start a separate thread to initialize the hardware and exit InitIO.  
  ''' If you encounter an error, you can always use InterfaceStatus to indicate this.
  ''' </summary>
  ''' <param name="port"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitIO(ByVal port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO

    Dim InterfaceStatus As String = ""
    Dim logMessage As String = ""

    Try
      '
      ' Write startup message to disk
      '
      WriteToDisk("HomeSeerV", hs.Version(), False)
      WriteToDisk("HomeSeerL", hs.IsLicensed())
      WriteToDisk("HomeSeerU", hs.SystemUpTime())

      '
      ' Write the debug message
      '
      WriteMessage("Entered InitIO() function.", MessageType.Debug)

      '
      ' Write the startup message
      '
      logMessage = String.Format("{0} version {1} initializing...", IFACE_NAME, HSPI.Version())
      Call WriteMessage(logMessage, MessageType.Informational)

      '
      ' Let's set the initial HomeSeer options for our plug-in
      '
      gLogLevel = Int32.Parse(hs.GetINISetting("Options", "LogLevel", LogLevel.Informational, gINIFile))

      '
      ' Let's set the initial HomeSeer options for our plug-in
      '
      gUnitType = hs.GetINISetting("Options", "UnitType", "US", gINIFile)
      gTempDegreeUnit = CBool(hs.GetINISetting("Options", "TempDegreeUnit", "True", gINIFile))
      gTempDegreeIcon = CBool(hs.GetINISetting("Options", "TempDegreeIcon", "False", gINIFile))
      gMonitoring = CBool(hs.GetINISetting("Options", "Monitoring", True, gINIFile))
      gDeviceValueType = hs.GetINISetting("Options", "DeviceValueType", "1", gINIFile)

      '
      ' Initialize our devices
      '
      hspi_devices.InitPluginDevices()

      bDBInitialized = hspi_database.InitializeMainDatabase()
      If bDBInitialized = True Then
        '
        ' Check to see if the devices need to be refreshed (the version # has changed)
        '
        Dim bVersionChanged As Boolean = False
        If GetSetting("Settings", "Version", "") <> HSPI.Version() Then
          bVersionChanged = True
        End If

        If bVersionChanged Then
          '
          ' Run maintenance tasks
          '
          WriteMessage("Running version has changed, performing maintenance ...", MessageType.Informational)
          SaveSetting("Settings", "Version", HSPI.Version())
        End If

        '
        ' Check database tables
        '
        CheckDatabaseTable("tblDevices")
        CheckDatabaseTable("tblSensorConfig")
        CheckDatabaseTable("tblTemperatureHistory")
        CheckDatabaseTable("tblHumidityHistory")
        CheckDatabaseTable("tblCounterHistory")
        CheckDatabaseTable("tblPressureHistory")

        '
        ' Start Plug-in Threads
        '
        UpdateOrphanedSensors()

        '
        ' Start Database Queue Thread
        '
        DatabaseQueueThread = New Thread(New ThreadStart(AddressOf ProcessDatabaseQueue))
        DatabaseQueueThread.Name = "DatabaseQueue"
        DatabaseQueueThread.Start()
        WriteMessage(String.Format("{0} Thread Started", DatabaseQueueThread.Name), MessageType.Debug)

        '
        ' Start History Queue Thread
        '
        SensorHistoryThread = New Thread(New ThreadStart(AddressOf ProcessSensorHistory))
        SensorHistoryThread.Name = "SensorHistory"
        SensorHistoryThread.Start()
        WriteMessage(String.Format("{0} Thread Started", SensorHistoryThread.Name), MessageType.Debug)


        '
        ' Start the CheckOneWireSensors thread
        '
        CheckOneWireSensorsThread = New Thread(New ThreadStart(AddressOf CheckOneWireSensors))
        CheckOneWireSensorsThread.Name = "CheckOneWireSensors"
        CheckOneWireSensorsThread.Start()
        WriteMessage(String.Format("{0} Thread Started", CheckOneWireSensorsThread.Name), MessageType.Debug)

      End If

      '
      ' Register the Configuration Web Page
      '
      WebConfigPage = New hspi_webconfig(IFACE_NAME)
      WebConfigPage.hspiref = Me
      RegisterWebPage(WebConfigPage.PageName, LINK_TEXT, LINK_PAGE_TITLE, instance)

      '
      ' Register the ASXP Web Page
      '
      'RegisterASXPWebPage(LINK_URL, LINK_TEXT, LINK_PAGE_TITLE, instance)

      '
      ' Register the Help File
      '
      RegisterHelpPage(LINK_HELP, IFACE_NAME & " Help File", IFACE_NAME)

      '
      ' Register for events from homeseer
      '
      'callback.RegisterEventCB(Enums.HSEvent.VALUE_CHANGE, IFACE_NAME, "")

      '
      ' Write the startup message
      '
      logMessage = String.Format("{0} version {1} initialization complete.", IFACE_NAME, HSPI.Version())
      Call WriteMessage(logMessage, MessageType.Informational)

      '
      ' Indicate the plug-in has been initialized
      '
      gHSInitialized = True

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "InitIO")

      gHSInitialized = False
    End Try

    Return InterfaceStatus

  End Function

  ''' <summary>
  ''' When HomeSeer shuts down or a plug-in is disabled from the interfaces page this function is then called. 
  ''' You should terminate any threads that you started, close any COM ports or TCP connections and release memory. 
  ''' After you return from this function the plugin EXE will terminate and must be allowed to terminate cleanly.
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered ShutdownIO() subroutine.", MessageType.Debug)

      Try
        hs.SaveEventsDevices()
      Catch pEx As Exception
        WriteMessage("Could not save devices", MessageType.Error)
      End Try

      Try

        If SensorHistoryThread.IsAlive Then
          SensorHistoryThread.Abort()
        End If
        If DatabaseQueueThread.IsAlive = True Then
          DatabaseQueueThread.Abort()
        End If
        If CheckOneWireSensorsThread.IsAlive = True Then
          CheckOneWireSensorsThread.Abort()
        End If

      Catch ex As Exception
        WriteMessage("Error shutting down threads.", MessageType.Error)
      End Try

      If instance = "" Then
        bShutDown = True
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "ShutdownIO")
    Finally
      gHSInitialized = False
    End Try

  End Sub

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="parms"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction

    Try

      Dim ty As Type = Me.GetType
      Dim mi As MethodInfo = ty.GetMethod(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Method {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
        Return Nothing
      End If
      Return (mi.Invoke(Me, parms))

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginFunction")
    End Try

    Return Nothing

  End Function

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="parms"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet

    Try
      Dim ty As Type = Me.GetType
      Dim mi As PropertyInfo = ty.GetProperty(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Property {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
        Return Nothing
      End If
      Return mi.GetValue(Me, parms)
    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginPropertyGet")
    End Try

    Return Nothing

  End Function

  ''' <summary>
  ''' There may be times when you need to offer a custom function that is not part of the plugin API. 
  ''' The following API functions allow users to call your plugin from scripts and web pages by calling the functions by name.
  ''' </summary>
  ''' <param name="proc"></param>
  ''' <param name="value"></param>
  ''' <remarks></remarks>
  Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet

    Try

      Dim ty As Type = Me.GetType
      Dim mi As PropertyInfo = ty.GetProperty(proc)
      If mi Is Nothing Then
        WriteMessage(String.Format("Property {0} does not exist in {1}.", proc, IFACE_NAME), MessageType.Error)
      End If
      mi.SetValue(Me, value, Nothing)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "PluginPropertySet")
    End Try

  End Sub

  ''' <summary>
  ''' This procedure will be called in your plug-in by HomeSeer whenever the user uses the search function of HomeSeer, and your plug-in is loaded and initialized.  
  ''' Unlike ActionReferencesDevice and TriggerReferencesDevice, this search is not being specific to a device, it is meant to find a match anywhere in the resources managed by your plug-in.  
  ''' This could include any textual field or object name that is utilized by the plug-in.
  ''' </summary>
  ''' <param name="SearchString"></param>
  ''' <param name="RegEx"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Search(ByVal SearchString As String,
                         ByVal RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements IPlugInAPI.Search

  End Function

#End Region

#Region "HSPI - Devices"

  ''' <summary>
  ''' Return TRUE if your plug-in allows for configuration of your devices via the device utility page. 
  ''' This will allow you to generate some HTML controls that will be displayed to the user for modifying the device
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
    Return True
  End Function

  ''' <summary>
  ''' If your plug-in manages all devices in the system, you can return TRUE from this function. Your configuration page will be available for all devices.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
    Return False
  End Function

  ''' <summary>
  ''' If SupportsConfigDevice returns TRUE, this function will be called when the device properties are displayed for your device. 
  ''' The device properties is displayed from the Device Utility page. This page displays a tab for each plug-in that controls the device. 
  ''' Normally, only one plug-in will be associated with a single device. 
  ''' If there is any configuration that needs to be set on the device, you can return any HTML that you would like displayed. 
  ''' Normally this would be any jquery controls that allow customization of the device. The returned HTML is just an HTML fragment and not a complete page.
  ''' </summary>
  ''' <param name="ref"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="newDevice"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String Implements HomeSeerAPI.IPlugInAPI.ConfigDevice

    Dim stb As New StringBuilder
    Dim jqButton As New clsJQuery.jqButton("btnSave", "Save", "DeviceUtility", True)

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered ConfigDevice() function.", MessageType.Debug)

      Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(ref)
      If dv Is Nothing Then Return "Error"

      Dim sensorAddr = dv.Address(hs)
      Dim OneWireSensor As OneWireSensor = OneWireSensors.Find(Function(s) s.sensorAddr = sensorAddr)
      If OneWireSensor Is Nothing Then Return "1-Wire Sensor not found."

      Dim OneWireDevice As OneWireDevice = Get1WireDevice(OneWireSensor.deviceId)
      If OneWireDevice Is Nothing Then Return "1-Wire Device not found."

      stb.AppendLine("<form id='frmConfigDevice' name='SampleTab' method='Post'>")
      stb.AppendLine("<table cellspacing='0' width='100%'>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Device Name:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireDevice.device_name)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Device Id:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireDevice.device_id)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Device Type:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireDevice.device_type)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Connection Type:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireDevice.device_conn)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Connection Address:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireDevice.device_addr)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Sensor ROM Id:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireSensor.romId)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label' style='width:150px'>{0}</td>", "Sensor Channel:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireSensor.sensorChannel)
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label'>{0}</td>", "Sensor Type:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", OneWireSensor.sensorType)
      stb.AppendLine(" </tr>")

      Dim cp As New clsJQuery.jqColorPicker("txtColorPicker", "", 40, OneWireSensor.sensorColor)

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell_label'>{0}</td>", "Sensor Color:")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", cp.Build)
      stb.AppendLine(" </tr>")

      Select Case OneWireSensor.sensorType
        Case "Counter"
          '
          ' Counter Subtype
          '
          Dim selSubtype As New clsJQuery.jqDropList("selSubtype", "", False)
          selSubtype.toolTip = "Select the sensor subtype."
          selSubtype.autoPostBack = False

          Dim sensorTypes() As String = {"Counter", "Lightning Counter", "Rain Gauge", "Water Meter", "Gas Meter"}
          For Each sensorType As String In sensorTypes
            selSubtype.AddItem(sensorType, sensorType, OneWireSensor.sensorSubtype = sensorType)
          Next

          stb.AppendLine(" <tr>")
          stb.AppendFormat("  <td class='tablecell_label'>{0}</td>", "Sensor Subtype:")
          stb.AppendFormat("  <td class='tablecell'>{0}</td>", selSubtype.Build)
          stb.AppendLine(" </tr>")

          '
          ' Counter Resolution
          '
          Dim tb As New clsJQuery.jqTextBox("txtSensorResolution", "text", OneWireSensor.sensorResolution, "", 10, True)
          tb.promptText = "Enter the counter resolution (e.g.  1 pulse equals 0.01)"
          tb.toolTip = tb.promptText

          stb.AppendLine(" <tr>")
          stb.AppendLine("  <td class='tablecell_label'>Sensor Resolution:</td>")
          stb.AppendFormat("  <td class='tablecell'>1 pulse equals {0}</td>{1}", tb.Build, vbCrLf)
          stb.AppendLine(" </tr>")

        Case "Humidity"
          '
          ' Humidity Subtype
          '
          Dim selSubtype As New clsJQuery.jqDropList("selSubtype", "", False)
          selSubtype.toolTip = "Select the sensor subtype."
          selSubtype.autoPostBack = False

          Dim sensorTypes() As String = {"Humidity", "Light Level"}
          For Each sensorType As String In sensorTypes
            selSubtype.AddItem(sensorType, sensorType, OneWireSensor.sensorSubtype = sensorType)
          Next

          stb.AppendLine(" <tr>")
          stb.AppendFormat("  <td class='tablecell_label'>{0}</td>", "Sensor Subtype:")
          stb.AppendFormat("  <td class='tablecell'>{0}</td>", selSubtype.Build)
          stb.AppendLine(" </tr>")

      End Select

      stb.AppendLine("</table>")
      stb.AppendFormat("<p>{0}</p>", jqButton.Build)
      stb.AppendLine("</form>")

    Catch pEx As Exception

    End Try

    'Dim jqButton As New clsJQuery.jqButton("button", "Press", "deviceutility", True)
    'stb.Append(jqButton.Build)

    stb.Append(clsPageBuilder.DivStart("div_config_device", ""))
    stb.Append(clsPageBuilder.DivEnd)

    Return stb.ToString

  End Function

  ''' <summary>
  ''' This function is called when a user posts information from your plugin tab on the device utility page.
  ''' </summary>
  ''' <param name="ref"></param>
  ''' <param name="data"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <returns>
  '''  DoneAndSave = 1            Any changes to the config are saved and the page is closed and the user it returned to the device utility page
  '''  DoneAndCancel = 2          Changes are not saved and the user is returned to the device utility page
  '''  DoneAndCancelAndStay = 3   No action is taken, the user remains on the plugin tab
  '''  CallbackOnce = 4           Your plugin ConfigDevice is called so your tab can be refereshed, the user stays on the plugin tab
  '''  CallbackTimer = 5          Your plugin ConfigDevice is called and a page timer is called so ConfigDevicePost is called back every 2 seconds
  ''' </returns>
  ''' <remarks></remarks>
  Function ConfigDevicePost(ByVal ref As Integer, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost

    Try

      '
      ' Write the debug message
      '
      WriteMessage("Entered ConfigDevicePost() function.", MessageType.Debug)

      '
      ' Check if device exists and get a reference to it
      '
      Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(ref)
      If dv Is Nothing Then Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay

      '
      ' Check if OneWireSensor exits and get a reference to it
      '
      Dim sensorAddr = dv.Address(hs)
      Dim OneWireSensor As OneWireSensor = OneWireSensors.Find(Function(s) s.sensorAddr = sensorAddr)
      If OneWireSensor Is Nothing Then Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay

      WriteMessage(String.Format("ConfigDevicePost: {0}", data), MessageType.Debug)

      '
      ' Handle form items
      '
      Dim parts As Collections.Specialized.NameValueCollection
      parts = System.Web.HttpUtility.ParseQueryString(data)

      Dim bDeviceUpdated As Boolean = False
      Dim bForceUpdate As Boolean = False

      If parts("id") = "btnSave" Then
        If Not parts.Get("selSubtype") Is Nothing Then
          If OneWireSensor.sensorSubtype <> parts("selSubtype") Then bForceUpdate = True
          OneWireSensor.sensorSubtype = parts("selSubtype")

          If Not parts.Get("txtSensorResolution") Is Nothing Then
            OneWireSensor.sensorResolution = Double.Parse(parts("txtSensorResolution"))
          End If
        End If

          OneWireSensor.sensorColor = parts("txtColorPicker")
          bDeviceUpdated = Update1WireSensor(OneWireSensor)

        Dim dv_addr As String = GetHomeSeerDevice(OneWireSensor, bForceUpdate)
        OneWireSensor.dvAddress = dv_addr

        hs.SetDeviceValueByRef(ref, OneWireSensor.Value, True)
        End If

        If bDeviceUpdated = True Then
          callback.ConfigDivToUpdateAdd("div_config_device", "Device options have been saved.")
        End If

    Catch pEx As Exception

    End Try

    Return Enums.ConfigDevicePostReturn.DoneAndCancelAndStay
  End Function

  ''' <summary>
  ''' Return TRUE if the plugin supports the ability to add devices through the Add Device link on the device utility page. 
  ''' If TRUE a tab appears on the add device page that allows the user to configure specific options for the new device.
  ''' When ConfigDevice is called the newDevice parameter will be True if this is the first time the device config screen is being displayed and a new device is being created. 
  ''' See ConfigDevicePost  for more information.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks>
  ''' </remarks>
  Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
    Return False
  End Function

  ''' <summary>
  ''' If a device is owned by your plug-in (interface property set to the name of the plug-in) and the device's status_support property is set to True, 
  ''' then this procedure will be called in your plug-in when the device's status is being polled, such as when the user clicks "Poll Devices" on the device status page.
  ''' Normally your plugin will automatically keep the status of its devices updated. 
  ''' There may be situations where automatically updating devices is not possible or CPU intensive. 
  ''' In these cases the plug-in may not keep the devices updated. HomeSeer may then call this function to force an update of a specific device. 
  ''' This request is normally done when a user displays the status page, or a script runs and needs to be sure it has the latest status.
  ''' </summary>
  ''' <param name="dvref"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PollDevice(ByVal dvref As Integer) As IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice

    '
    ' Write the debug message
    '
    WriteMessage("Entered PollDevice() function.", MessageType.Debug)

  End Function

#End Region

#Region "HSPI - Triggers"

  ''' <summary>
  ''' Return True if the given trigger can also be used as a condition, for the given trigger number.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
    Get
      Return False
    End Get
  End Property

  ''' <summary>
  ''' Return True if your plugin contains any triggers, else return false  ''' 
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property HasTriggers() As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
    Get
      Return False
    End Get
  End Property

  ''' <summary>
  ''' Not documented in API
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerTrue
    Return True
  End Function

  ''' <summary>
  ''' Return the number of triggers that your plugin supports.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
    Get
      Return 0
    End Get
  End Property

  ''' <summary>
  ''' Return the number of sub triggers your plugin supports.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
    Get
      Return 0
    End Get
  End Property

  ''' <summary>
  ''' Return the name of the given trigger based on the trigger number passed.
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
    Get
      Return 0
    End Get
  End Property

  ''' <summary>
  ''' Return the text name of the sub trigger given its trigger number and sub trigger number
  ''' </summary>
  ''' <param name="TriggerNumber"></param>
  ''' <param name="SubTriggerNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
    Get
      Return ""
    End Get
  End Property

  ''' <summary>
  ''' Given a strTrigActInfo object detect if this this trigger is configured properly, if so, return True, else False.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
    Get
      Return True
    End Get
  End Property

  ''' <summary>
  ''' Return HTML controls for a given trigger.
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
    Return ""
  End Function

  ''' <summary>
  ''' Process a post from the events web page when a user modifies any of the controls related to a plugin trigger. 
  ''' After processing the user selctions, create and return a strMultiReturn object. 
  ''' </summary>
  ''' <param name="PostData"></param>
  ''' <param name="TrigInfoIn"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, _
                                          ByVal TrigInfoIn As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI

  End Function

  ''' <summary>
  ''' After the trigger has been configured, this function is called in your plugin to display the configured trigger. 
  ''' Return text that describes the given trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
    Return ""
  End Function

  ''' <summary>
  ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  
  ''' Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
    Set(ByVal value As Boolean)

    End Set
    Get
      Return False
    End Get
  End Property

  ''' <summary>
  ''' Return True if the given device is referenced by the given trigger.
  ''' </summary>
  ''' <param name="TrigInfo"></param>
  ''' <param name="dvRef"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function TriggerReferencesDevice(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, _
                                                                           ByVal dvRef As Integer) As Boolean _
                                                                           Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
    Return False
  End Function

#End Region

#Region "HSPI - Actions"

  ''' <summary>
  ''' Return the number of actions the plugin supports.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
    Return 0
  End Function

  ''' <summary>
  ''' When an event is triggered, this function is called to carry out the selected action. Use the ActInfo parameter to determine what action needs to be executed then execute this action.
  ''' Return TRUE if the action was executed successfully, else FALSE if there was an error.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
    Return True
  End Function

  ''' <summary>
  ''' Missing in the API Docs
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI

  End Function

  Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, _
                                      ByVal ActInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI


  End Function

  ''' <summary>
  ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <param name="dvRef"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionReferencesDevice(ByVal ActInfo As IPlugInAPI.strTrigActInfo, _
                                         ByVal dvRef As Integer) As Boolean _
                                         Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice

  End Function

  ''' <summary>
  ''' This function is called from the HomeSeer event page when an event is in edit mode. Your plug-in needs to return HTML controls so the user can make action selections. 
  ''' Normally this is one of the HomeSeer jquery controls such as a clsJquery.jqueryCheckbox.
  ''' </summary>
  ''' <param name="sUnique"></param>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
    Return ""
  End Function

  ''' <summary>
  ''' Return TRUE if the given action is configured properly. 
  ''' There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.
  ''' </summary>
  ''' <param name="ActInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
    Return True
  End Function

  ''' <summary>
  ''' Return the name of the action given an action number. The name of the action will be displayed in the HomeSeer events actions list.
  ''' </summary>
  ''' <param name="ActionNumber"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
    Get
      Return ""
    End Get
  End Property

  ''' <summary>
  ''' The HomeSeer events page has an option to set the editing mode to "Advanced Mode". 
  ''' This is typically used to enable options that may only be of interest to advanced users or programmers. The Set in this function is called when advanced mode is enabled. 
  ''' Your plug-in can also enable this mode if an advanced selection was saved and needs to be displayed.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
    Set(ByVal value As Boolean)
      mvarActionAdvanced = value
    End Set
    Get
      Return mvarActionAdvanced
    End Get
  End Property
  Private mvarActionAdvanced As Boolean

  ''' <summary>
  ''' Not in API Docs
  ''' </summary>
  ''' <param name="PlugName"></param>
  ''' <param name="evRef"></param>
  ''' <param name="ActionInfo"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function UpdatePlugAction(ByVal PlugName As String, ByVal evRef As Integer, ByVal ActionInfo As IPlugInAPI.strTrigActInfo) As String
    Return ""
  End Function

#End Region

#Region "HSPI - WebPage"

  ''' <summary>
  ''' When your plug-in web page has form elements on it, and the form is submitted, this procedure is called to handle the HTTP "Put" request.  
  ''' There must be one PagePut procedure in each plug-in object or class that is registered as a web page in HomeSeer.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin

    Try
      '
      ' Build and return the actual page
      '
      WebPage = SelectPage(pageName)
      Return WebPage.GetPagePlugin(pageName, user, userRights, queryString, instance)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' Determine what page we need to display
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function SelectPage(ByVal pageName As String) As Object

    WriteMessage("Entered SelectPage() Function", MessageType.Debug)
    WriteMessage("pageName=" & pageName, MessageType.Debug)

    SelectPage = Nothing
    Try

      Select Case pageName
        Case WebConfigPage.PageName
          SelectPage = WebConfigPage
        Case Else
          SelectPage = WebConfigPage
      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "SelectPage")
    End Try

  End Function

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. 
  ''' You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="data"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
    Try
      '
      ' Build and return the actual page
      '
      WebPage = SelectPage(pageName)
      Return WebPage.postBackProc(pageName, data, user, userRights)

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "PostBackProc")
    End Try
  End Function

  ''' <summary>
  ''' This function is called by HomeSeer from the form or class object that a web page was registered with using RegisterConfigLink.  
  ''' You must have a GenPage procedure per web page that you register with HomeSeer.  
  ''' This page is called when the user requests the web page with an HTTP Get command, which is the default operation when the browser requests a page.
  ''' </summary>
  ''' <param name="link"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GenPage(ByVal link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage

    Dim sb As New StringBuilder()

    Try
      '
      ' Generate the HTML re-direct
      '
      sb.Append("HTTP/1.1 301 Moved Permanently" & vbCrLf)
      sb.AppendFormat("Location: {0}" & vbCrLf, LINK_TARGET)
      sb.Append(vbCrLf)

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process the error
      '
      Return ""
    End Try

  End Function

  ''' <summary>
  ''' When your plug-in web page has form elements on it, and the form is submitted, this procedure is called to handle the HTTP "Put" request.  
  ''' There must be one PagePut procedure in each plug-in object or class that is registered as a web page in HomeSeer
  ''' </summary>
  ''' <param name="data"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function PagePut(ByVal data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
    Return ""
  End Function

#End Region

#Region "HSPI - Speak Proxy"

  ''' <summary>
  ''' If your plug-in is registered as a Speak proxy plug-in, then when HomeSeer is asked to speak something, it will pass the speak information to your plug-in using this procedure.  
  ''' When your plug-in is ready to do the actual speaking, it should call SpeakProxy, and pass the information that it got from this procedure to SpeakProxy.  
  ''' It may be necessary or a feature of your plug-in to modify the text being spoken or the host/instance list provided in the host parameter - this is acceptable.
  ''' </summary>
  ''' <param name="device"></param>
  ''' <param name="txt"></param>
  ''' <param name="w"></param>
  ''' <param name="host"></param>
  ''' <remarks></remarks>
  Public Sub SpeakIn(device As Integer, txt As String, w As Boolean, host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn

  End Sub

#End Region

#Region "HSPI - Callbacks"

  Public Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent

    Console.WriteLine("HSEvent: " & EventType.ToString)
    Select Case EventType
      Case Enums.HSEvent.VALUE_CHANGE
        '
        ' Do stuff
        '
    End Select

  End Sub

  ''' <summary>
  ''' Return True if the given device is referenced in the given action.
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
    Return False
  End Function

#End Region

#Region "HSPI - Plug-in Version Information"

  ''' <summary>
  ''' Returns the full version number of the HomeSeer plug-in
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Shared Function Version() As String

    Dim strVersion As String = ""

    Try
      '
      ' Write the debug message
      '
      WriteMessage("Entered Version() function.", MessageType.Debug)

      '
      ' Build the plug-in's version
      '
      strVersion = GetVersionInfo(GetExecutingAssembly.Location).ProductMajorPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductMinorPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductBuildPart() & "." _
                  & GetVersionInfo(GetExecutingAssembly.Location).ProductPrivatePart()

    Catch pEx As Exception
      '
      ' Process error condtion
      '
      strVersion = "??.??.??"
    End Try

    Return strVersion

  End Function

#End Region

#Region "HSPI - Public Properties"

  ''' <summary>
  ''' Returns the JSON sensor data
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

    Return hspi_plugin.GetSensorChartJSON(device_id, ChartType, AMChartType, strEndDate, Interval)

  End Function

  ''' <summary>
  ''' Returns various statistics
  ''' </summary>
  ''' <param name="StatisticsType"></param>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property GetStatistics(ByVal StatisticsType As String) As Integer

    Get
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
    End Get

  End Property

  ''' <summary>
  ''' Returns the size of the selected database
  ''' </summary>
  ''' <param name="databaseName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetDatabaseSize(ByVal databaseName As String)
    Return hspi_database.GetDatabaseSize(databaseName)
  End Function

  ''' <summary>
  ''' Property to control log level
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Property LoggingLevel() As Integer
    Get
      Return gLogLevel
    End Get
    Set(ByVal value As Integer)
      gLogLevel = value
    End Set
  End Property

#End Region

#Region "HSPI - Public Sub/Functions"

  ''' <summary>
  ''' Returns the list of configured 1-wire devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function Get1WireDevices() As DataTable

    Return HSPI_ULTRA1WIRE3.Get1WireDevices()

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
  Public Function Insert1WireDevice(ByVal device_serial As String, _
                                    ByVal device_name As String, _
                                    ByVal device_type As String, _
                                    ByVal device_image As String, _
                                    ByVal device_conn As String, _
                                    ByVal device_addr As String) As Boolean


    Return HSPI_ULTRA1WIRE3.Insert1WireDevice(device_serial, device_name, device_type, device_image, device_conn, device_addr)

  End Function

  '------------------------------------------------------------------------------------
  'Purpose:   Updates existing TEMP08 Device stored in the database
  'Input:     Multiple, see below
  'Output:    Boolean
  '------------------------------------------------------------------------------------
  Public Function Update1WireDevice(ByVal device_id As Integer, _
                                    ByVal device_serial As String, _
                                    ByVal device_name As String, _
                                    ByVal device_type As String, _
                                    ByVal device_conn As String, _
                                    ByVal device_addr As String) As Boolean

    Return HSPI_ULTRA1WIRE3.Update1WireDevice(device_id, device_serial, device_name, device_type, device_conn, device_addr)

  End Function

  '------------------------------------------------------------------------------------
  'Purpose:   Removes existing TEMP08 Device stored in the database
  'Input:     device_id As Integer
  'Output:    Boolean
  '------------------------------------------------------------------------------------
  Public Function Delete1WireDevice(ByVal device_id As Integer) As Boolean

    'Try

    '  SyncLock ECMObjects.SyncRoot

    '    Dim strDeviceId As String = String.Format("ECM{0}", device_id.ToString)
    '    If ECMObjects.ContainsKey(strDeviceId) = True Then

    '      Dim BrultechDevice As BrultechDevice = ECMObjects(strDeviceId)

    '      '
    '      ' Disconnect from the Brultech ECM Device
    '      '
    '      BrultechDevice.Disconnect()
    '      BrultechDevice = Nothing

    '      ECMObjects.Remove(strDeviceId)

    '    End If

    '  End SyncLock

    'Catch pEx As Exception
    '  '
    '  ' Process Exception
    '  '
    '  Call ProcessError(pEx, "DeleteBrultechECM()")
    'End Try

    Return HSPI_ULTRA1WIRE3.Delete1WireDevice(device_id)

  End Function

  ''' <summary>
  ''' Returns a list of valid serial ports
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetSerialPortNames() As Specialized.StringCollection

    Dim PortNames As New Specialized.StringCollection

    Try
      '
      ' Find all available serial port names on this computer
      '
      For Each strPortName As String In My.Computer.Ports.SerialPortNames
        PortNames.Add(strPortName)
      Next
    Catch pEx As Exception

    End Try

    Return PortNames

  End Function

  ''' <summary>
  ''' Resets database statistics
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub ResetStatistics()

    Dim strMessage As String = ""

    Try
      strMessage = "Entered ResetStatistics() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Reset global variables
      '
      gDBInsertSuccess = 0
      gDBInsertFailure = 0

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "ResetStatistics()")
    End Try

  End Sub

  ''' <summary>
  ''' Flushes sensor list
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub FlushSensorList()

    Dim strMessage As String = ""

    Try
      strMessage = "Entered FlushSensorList() subroutine."
      Call WriteMessage(strMessage, MessageType.Debug)

      '
      ' Reset the discovered sensors
      '
      FlushSensorList()

    Catch pEx As Exception
      '
      ' Process the error
      '
      Call ProcessError(pEx, "FlushSensorList()")
    End Try

  End Sub

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
      '
      ' Write the debug message
      '
      WriteMessage("Entered GetSetting() function.", MessageType.Debug)

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
  '''  Saves plug-in settings to INI file
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
      '
      ' Write the debug message
      '
      WriteMessage("Entered SaveSetting() subroutine.", MessageType.Debug)

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

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the list of users authorized to access the web page
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function WEBUserRolesAuthorized() As Integer

    Return hspi_main.WEBUserRolesAuthorized

  End Function

  '----------------------------------------------------------------------
  'Purpose: Determine if logged in user is authorized to view the web page
  'Inputs:  LoggedInUser as String
  'Outputs: Boolean (True indicates user is authorized)
  '----------------------------------------------------------------------
  ''' <summary>
  ''' Determine if logged in user is authorized to view the web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <param name="USER_ROLES_AUTHORIZED"></param>
  ''' <returns>Boolean (True indicates user is authorized)</returns>
  ''' <remarks></remarks>
  Public Function WEBUserIsAuthorized(ByVal LoggedInUser As String, _
                                      ByVal USER_ROLES_AUTHORIZED As Integer) As Boolean

    Return hspi_main.WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED)

  End Function

#End Region

End Class
