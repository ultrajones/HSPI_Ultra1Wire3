Imports HomeSeerAPI
Imports Scheduler
Imports HSCF
Imports HSCF.Communication.ScsServices.Service
Imports System.Text.RegularExpressions

Module hspi_devices

  Public DEV_STATUS As Byte = 1

  Dim bCreateRootDevice = True

  ''' <summary>
  ''' Update the list of monitored devices
  ''' </summary>
  ''' <remarks></remarks>
  Public Sub UpdateMiscDeviceSettings(ByVal bDeviceNoLog As Boolean)

    Dim dv As Scheduler.Classes.DeviceClass
    Dim DevEnum As Scheduler.Classes.clsDeviceEnumeration

    Dim strMessage As String = ""

    strMessage = "Entered UpdateMiscDeviceSettings() function."
    Call WriteMessage(strMessage, MessageType.Debug)

    Try
      '
      ' Go through devices to see if we have one assigned to our plug-in
      '
      DevEnum = hs.GetDeviceEnumerator

      If Not DevEnum Is Nothing Then

        Do While Not DevEnum.Finished
          dv = DevEnum.GetNext
          If dv Is Nothing Then Continue Do
          If dv.Interface(Nothing) IsNot Nothing Then
            If dv.Interface(Nothing) = IFACE_NAME Then
              '
              ' We found our device, so process based on device type
              '
              Dim dv_type As String = dv.Device_Type_String(hs)

              '
              ' Set options based on root or child device
              '
              If dv.Relationship(hs) = Enums.eRelationship.Parent_Root Then
                dv.MISC_Set(hs, Enums.dvMISC.NO_STATUS_TRIGGER)               ' When set, the device status values will Not appear in the device change trigger
              ElseIf dv.Relationship(hs) = Enums.eRelationship.Child Then
                dv.MISC_Clear(hs, Enums.dvMISC.NO_STATUS_TRIGGER)             ' When set, the device status values will Not appear in the device change trigger

                Select Case dv_type
                  Case "Ultra1Wire3 Monitoring"
                    dv.MISC_Clear(hs, Enums.dvMISC.STATUS_ONLY)               ' When set, the device cannot be controlled
                  Case Else
                    dv.MISC_Set(hs, Enums.dvMISC.STATUS_ONLY)                 ' When set, the device cannot be controlled
                End Select
              End If

              '
              ' None of the devices can be controlled by Voice
              '
              dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)              ' When set, this device is included in the voice recognition context for device commands

              '
              ' Apply Logging Options
              '
              If bDeviceNoLog = False Then
                dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)                          ' When set, no logging to the log for this device
              Else
                dv.MISC_Clear(hs, Enums.dvMISC.NO_LOG)                        ' When set, no logging to the log for this device
              End If

              '
              ' This property indicates (when True) that the device supports the retrieval of its status on-demand through the "Poll" feature on the device utility page.
              '
              dv.Status_Support(hs) = False

              '
              ' If an event or device was modified by a script, this function should be called to update HomeSeer with the changes.
              '
              hs.SaveEventsDevices()

            End If
          End If
        Loop
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "UpdateMiscDeviceSettings()")
    End Try

    DevEnum = Nothing
    dv = Nothing

  End Sub

  ''' <summary>
  ''' Create the HomeSeer Root Device
  ''' </summary>
  ''' <param name="strRootId"></param>
  ''' <param name="strRootName"></param>
  ''' <param name="dv_ref_child"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreateRootDevice(ByVal strRootId As String,
                                   ByVal strRootName As String,
                                   ByVal dv_ref_child As Integer) As Integer

    Dim dv As Scheduler.Classes.DeviceClass

    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Set the local variables
      '
      If strRootId = "Plugin" Then
        dv_name = "Ultra1Wire3 Plugin"
        dv_addr = String.Format("{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = dv_name
      Else
        dv_name = String.Format("{0} Root Device", strRootName)
        dv_addr = String.Format("1Wire-{0}-Root", strRootName.Replace(" ", "-"))
        dv_type = strRootName
      End If

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} root device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} root device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = IIf(strRootId = "Plugin", "Plug-ins", "Environmental")
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this a parent root device
      '
      dv.Relationship(hs) = Enums.eRelationship.Parent_Root
      dv.AssociatedDevice_Add(hs, dv_ref_child)

      Dim image As String = "device_root.png"

      Dim VSPair As VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
      VSPair.PairType = VSVGPairType.SingleValue
      VSPair.Value = 0
      VSPair.Status = "Root"
      VSPair.Render = Enums.CAPIControlType.Values
      hs.DeviceVSP_AddPair(dv_ref, VSPair)

      Dim VGPair As VGPair = New VGPair()
      VGPair.PairType = VSVGPairType.SingleValue
      VGPair.Set_Value = 0
      VGPair.Graphic = String.Format("{0}{1}", gImageDir, image)
      hs.DeviceVGP_AddPair(dv_ref, VGPair)

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.STATUS_ONLY)           ' When set, the device cannot be controlled
      dv.MISC_Set(hs, Enums.dvMISC.NO_STATUS_TRIGGER)     ' When set, the device status values will Not appear in the device change trigger
      dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)  ' When set, this device is included in the voice recognition context for device commands

      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)         ' When set, device control options will be displayed
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)              ' When set, no logging to the log for this device
      End If

      dv.Status_Support(hs) = False

      hs.SaveEventsDevices()

    Catch pEx As Exception

    End Try

    Return dv_ref

  End Function

  ''' <summary>
  ''' Function to initilize our plug-ins devices
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function InitPluginDevices() As String

    Dim strMessage As String = ""

    WriteMessage("Entered InitPluginDevices() function.", MessageType.Debug)

    Try
      Dim Devices As Byte() = {DEV_PLUGIN_INTERFACE}
      For Each dev_cod As Byte In Devices
        Dim strResult As String = CreatePluginDevice(IFACE_NAME, dev_cod)
        If strResult.Length > 0 Then Return strResult
      Next

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "InitPluginDevices()")
      Return pEx.ToString
    End Try

  End Function

  ''' <summary>
  ''' Subroutine to create a HomeSeer device
  ''' </summary>
  ''' <param name="base_code"></param>
  ''' <param name="dev_code"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function CreatePluginDevice(ByVal base_code As String, ByVal dev_code As String) As String

    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0

    Dim dv_name As String = ""
    Dim dv_type As String = ""
    Dim dv_addr As String = ""

    Dim DeviceShowValues As Boolean = False
    Dim DevicePairs As New ArrayList

    Try

      Select Case dev_code
        Case DEV_PLUGIN_INTERFACE.ToString
          '
          ' Create the Monitoring State device
          '
          dv_name = "1-Wire Sensor Checking"
          dv_type = "Ultra1Wire3 Monitoring"
          dv_addr = String.Concat(base_code, ":", dev_code)

        Case DEV_DATABASE_INTERFACE.ToString
          '
          ' Create the Monitoring State device
          '
          dv_name = "Database Connection"
          dv_type = "Ultra1Wire3 Database"
          dv_addr = String.Concat(base_code, ":", dev_code)

        Case Else
          Throw New Exception(String.Format("Unable to create plug-in device for unsupported device name: {0}", dv_name))
      End Select

      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Plug-ins"
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Make this device a child of the root
      '
      If bCreateRootDevice = True Then
        dv.AssociatedDevice_ClearAll(hs)
        Dim dvp_ref As Integer = CreateRootDevice("Plugin", IFACE_NAME, dv_ref)
        If dvp_ref > 0 Then
          dv.AssociatedDevice_Add(hs, dvp_ref)
        End If
        dv.Relationship(hs) = Enums.eRelationship.Child
      End If

      '
      ' Update the last change date
      ' 
      dv.Last_Change(hs) = DateTime.Now

      Dim VSPair As VSPair
      Dim VGPair As VGPair

      Select Case dv_type
        Case "Ultra1Wire3 Monitoring"

          DevicePairs.Clear()
          DevicePairs.Add(New hspi_device_pairs(-3, "", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-2, "Disable", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(-1, "Enable", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Control))
          DevicePairs.Add(New hspi_device_pairs(0, "Disaled", "state_disabled.png", HomeSeerAPI.ePairStatusControl.Status))
          DevicePairs.Add(New hspi_device_pairs(1, "Enabled", "state_enabled.png", HomeSeerAPI.ePairStatusControl.Status))

          '
          ' Add the Status Graphic Pairs
          '
          For Each Pair As hspi_device_pairs In DevicePairs

            VSPair = New VSPair(Pair.Type)
            VSPair.PairType = VSVGPairType.SingleValue
            VSPair.Value = Pair.Value
            VSPair.Status = Pair.Status
            VSPair.Render = Enums.CAPIControlType.Values
            hs.DeviceVSP_AddPair(dv_ref, VSPair)

            VGPair = New VGPair()
            VGPair.PairType = VSVGPairType.SingleValue
            VGPair.Set_Value = Pair.Value
            VGPair.Graphic = String.Format("{0}{1}", gImageDir, Pair.Image)
            hs.DeviceVGP_AddPair(dv_ref, VGPair)

          Next

          Dim dev_status As Integer = IIf(gMonitoring = True, 1, 0)
          hs.SetDeviceValueByRef(dv_ref, dev_status, False)

          DeviceShowValues = True

        Case "Ultra1Wire3 Database"

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -3
          VSPair.Status = ""
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -2
          VSPair.Status = "Close"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Control)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = -1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Values
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Closed"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "Open"
          VSPair.Render = Enums.CAPIControlType.Button
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          DeviceShowValues = True
      End Select

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Clear(hs, Enums.dvMISC.STATUS_ONLY)         ' When set, the device cannot be controlled
      dv.MISC_Clear(hs, Enums.dvMISC.NO_STATUS_TRIGGER)   ' When set, the device status values will Not appear in the device change trigger
      dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)  ' When set, this device is included in the voice recognition context for device commands

      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)         ' When set, device control options will be displayed
      End If

      If bDeviceExists = False Then
        dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)              ' When set, no logging to the log for this device
      End If

      '
      ' This property indicates (when True) that the device supports the retrieval of its status on-demand through the "Poll" feature on the device utility page.
      '
      dv.Status_Support(hs) = False

      '
      ' If an event or device was modified by a script, this function should be called to update HomeSeer with the changes.
      '
      hs.SaveEventsDevices()

      Return ""

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "CreatePluinDevice()")
      Return "Failed to create HomeSeer device due to error."
    End Try

  End Function

  ''' <summary>
  ''' Subroutine to create the HomeSeer device
  ''' </summary>
  ''' <param name="OneWireSensor"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetHomeSeerDevice(ByVal OneWireSensor As OneWireSensor,
                                    Optional bForceUpdate As Boolean = False) As String


    Dim dv As Scheduler.Classes.DeviceClass
    Dim dv_ref As Integer = 0
    Dim dv_misc As Integer = 0
    Dim dv_offset As Double = 0.0

    Dim dv_name As String = OneWireSensor.sensorName
    Dim dv_type As String = OneWireSensor.sensorType
    Dim dv_addr As String = OneWireSensor.sensorAddr
    Dim dv_subtype As String = OneWireSensor.sensorSubtype

    Dim DeviceShowValues As Boolean = False

    Try
      '
      ' Define local variables
      '
      dv = LocateDeviceByAddr(dv_addr)
      Dim bDeviceExists As Boolean = Not dv Is Nothing

      If bDeviceExists = True Then
        '
        ' Lets use the existing device
        '
        dv_addr = dv.Address(hs)
        dv_ref = dv.Ref(hs)

        Call WriteMessage(String.Format("Updating existing HomeSeer {0} device.", dv_name), MessageType.Debug)

        Try
          Dim VSP As VSPair = hs.DeviceVSP_Get(dv_ref, 0, ePairStatusControl.Status)
          dv_offset = VSP.ValueOffset
        Catch pEx As Exception

        End Try

      Else
        '
        ' Create A HomeSeer Device
        '
        dv_ref = hs.NewDeviceRef(dv_name)
        If dv_ref > 0 Then
          dv = hs.GetDeviceByRef(dv_ref)
        End If

        Call WriteMessage(String.Format("Creating new HomeSeer {0} device.", dv_name), MessageType.Debug)

      End If

      '
      ' Make this device a child of the root
      '
      If dv.Relationship(hs) <> Enums.eRelationship.Child Or bForceUpdate = True Then

        If bCreateRootDevice = True Then
          dv.AssociatedDevice_ClearAll(hs)
          Dim dv_type_root As String = dv_type
          Select Case dv_subtype
            Case "Light Level"
              dv_type_root = "Light"
          End Select
          Dim dvp_ref As Integer = CreateRootDevice("", dv_type_root, dv_ref)
          If dvp_ref > 0 Then
            dv.AssociatedDevice_Add(hs, dvp_ref)
          End If
          dv.Relationship(hs) = Enums.eRelationship.Child
        End If

        hs.SaveEventsDevices()
      End If

      '
      ' Exit if our device exists
      '
      If bDeviceExists = True And bForceUpdate = False Then Return dv_addr

      '
      ' Define the HomeSeer device
      '
      dv.Address(hs) = dv_addr
      dv.Interface(hs) = IFACE_NAME
      dv.InterfaceInstance(hs) = Instance

      '
      ' Update location properties on new devices only
      '
      If bDeviceExists = False Then
        dv.Location(hs) = IFACE_NAME & " Plugin"
        dv.Location2(hs) = "Environmental"

        '
        ' Update the last change date
        ' 
        dv.Last_Change(hs) = DateTime.Now
      End If

      '
      ' The following simply shows up in the device properties but has no other use
      '
      dv.Device_Type_String(hs) = dv_type

      '
      ' Set the DeviceTypeInfo
      '
      Dim DT As New DeviceTypeInfo
      DT.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
      DT.Device_Type = DeviceTypeInfo.eDeviceType_Plugin.Root
      dv.DeviceType_Set(hs) = DT

      '
      ' Clear the value status pairs
      '
      hs.DeviceVSP_ClearAll(dv_ref, True)
      hs.DeviceVGP_ClearAll(dv_ref, True)
      hs.SaveEventsDevices()

      Dim VSPair As VSPair
      Dim VGPair As VGPair
      Select Case dv_type
        Case "Switch"
          '
          ' Format Switch Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 0
          VSPair.Status = "Off"
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.SingleValue
          VSPair.Value = 1
          VSPair.Status = "On"
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "off.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.SingleValue
          VGPair.Set_Value = 1
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "on.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Temperature"
          '
          ' Format Temperature Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)
          If gTempDegreeUnit = True Then
            strStatusSuffix &= IIf(gUnitType = "Metric", "C", "F")
          End If

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = -350
          VSPair.RangeEnd = 1300
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.RangeStatusDecimals = 1
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -350
          VGPair.RangeEnd = 1300
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "temperature.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Counter"
          '
          ' Format Counter Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = Double.MaxValue
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.RangeStatusDecimals = 2
          VSPair.ValueOffset = dv_offset
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = Double.MaxValue
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "counter.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Vibration"
          '
          ' Format Vibration Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = -1
          VSPair.RangeEnd = 0
          VSPair.RangeStatusSuffix = "None"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (No Vibration)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = -1
          VGPair.RangeEnd = 0
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "vibration_none.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Low Vibration)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 1
          VSPair.RangeEnd = 256
          VSPair.RangeStatusSuffix = "Low"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Low Vibration)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 1
          VGPair.RangeEnd = 256
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "vibration_low.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Medium Vibration)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 257
          VSPair.RangeEnd = 512
          VSPair.RangeStatusSuffix = "Medium"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Medium Vibration)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 257
          VGPair.RangeEnd = 512
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "vibration_medium.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (High Vibration)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 513
          VSPair.RangeEnd = 768
          VSPair.RangeStatusSuffix = "High"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (High Vibration)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 513
          VGPair.RangeEnd = 768
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "vibration_high.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

          '
          ' Add VSPair (Very High Vibration)
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 769
          VSPair.RangeEnd = 1023
          VSPair.RangeStatusSuffix = "Very High"
          VSPair.IncludeValues = False
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair (Very High Vibration)
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 769
          VGPair.RangeEnd = 1023
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "vibration_very_high.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Pressure"
          '
          ' Format Pressure Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 1200
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.RangeStatusDecimals = 2
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 1200
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "pressure.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Light"
          '
          ' Format Light Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 130000
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 130000
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "light.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Humidity"
          '
          ' Format Humidity Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)
          Dim strImage As String = IIf(dv_subtype = "Light Level", "light.png", "humidity.png")

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 100
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.RangeStatusDecimals = 1
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 100
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, strImage)
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

        Case "Voltage"
          '
          ' Format Voltage Status Suffix
          '
          Dim strStatusSuffix As String = String.Concat(" ", OneWireSensor.sensorUnits)

          '
          ' Add VSPair
          '
          VSPair = New VSPair(ePairStatusControl.Status)
          VSPair.PairType = VSVGPairType.Range
          VSPair.RangeStart = 0
          VSPair.RangeEnd = 240
          VSPair.RangeStatusDecimals = 2
          VSPair.RangeStatusPrefix = ""
          VSPair.RangeStatusSuffix = strStatusSuffix
          VSPair.Render = Enums.CAPIControlType.ValuesRangeSlider
          hs.DeviceVSP_AddPair(dv_ref, VSPair)

          '
          ' Add VSPair
          '
          VGPair = New VGPair()
          VGPair.PairType = VSVGPairType.Range
          VGPair.RangeStart = 0
          VGPair.RangeEnd = 240
          VGPair.Graphic = String.Format("{0}{1}", gImageDir, "voltage.png")
          hs.DeviceVGP_AddPair(dv_ref, VGPair)

      End Select

      '
      ' Update the Device Misc Bits
      '
      dv.MISC_Set(hs, Enums.dvMISC.STATUS_ONLY)           ' When set, the device cannot be controlled
      dv.MISC_Clear(hs, Enums.dvMISC.NO_STATUS_TRIGGER)   ' When set, the device status values will Not appear in the device change trigger
      dv.MISC_Clear(hs, Enums.dvMISC.AUTO_VOICE_COMMAND)  ' When set, this device is included in the voice recognition context for device commands

      If DeviceShowValues = True Then
        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)         ' When set, device control options will be displayed
      End If

      If bDeviceExists = False Then
        Dim dv_log As Boolean = CBool(GetSetting("Sensor", "DeviceLogging", False))
        Select Case dv_log
          Case True
            dv.MISC_Clear(hs, Enums.dvMISC.NO_LOG)        ' When set, no logging to the log for this device
          Case False
            dv.MISC_Set(hs, Enums.dvMISC.NO_LOG)          ' When set, no logging to the log for this device
        End Select
      End If

      '
      ' This property indicates (when True) that the device supports the retrieval of its status on-demand through the "Poll" feature on the device utility page.
      '
      dv.Status_Support(hs) = False

      '
      ' If an event or device was modified by a script, this function should be called to update HomeSeer with the changes.
      '
      hs.SaveEventsDevices()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "GetHomeSeerDevice()")
    End Try

    Return dv_addr

  End Function

  ''' <summary>
  ''' Locates device by device code
  ''' </summary>
  ''' <param name="strDeviceAddr"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function LocateDeviceByAddr(ByVal strDeviceAddr As String) As Object

    Dim objDevice As Object
    Dim dev_ref As Long = 0

    Try

      dev_ref = hs.DeviceExistsAddress(strDeviceAddr, False)
      objDevice = hs.GetDeviceByRef(dev_ref)
      If Not objDevice Is Nothing Then
        Return objDevice
      End If

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      Call ProcessError(pEx, "LocateDeviceByAddr")
    End Try
    Return Nothing ' No device found

  End Function

  ''' <summary>
  ''' Sets the OneWireSensor Value
  ''' </summary>
  ''' <param name="OneWireSensor"></param>
  Public Sub SetDeviceValue(ByRef OneWireSensor As OneWireSensor, SensorValue As Double)

    Try

      WriteMessage(String.Format("{0}->{1}->{2}->{3}",
                                 OneWireSensor.dvAddress,
                                 OneWireSensor.romId,
                                 OneWireSensor.sensorType,
                                 OneWireSensor.Value),
                                 MessageType.Debug)

      Dim dv_ref As Integer = hs.DeviceExistsAddress(OneWireSensor.dvAddress, False)
      Dim bDeviceExists As Boolean = dv_ref <> -1

      WriteMessage(String.Format("Device address {0} was found.", OneWireSensor.dvAddress), MessageType.Debug)

      If bDeviceExists = True Then

        If IsNumeric(SensorValue) Then

          Dim dblDeviceValue As Double = Double.Parse(hs.DeviceValueEx(dv_ref))
          Dim dblSensorValue As Double = Double.Parse(SensorValue)

          If dblDeviceValue <> dblSensorValue Then
            hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)

            '
            ' Write the value change to the log
            '
            Try
              Dim dev_ref As Long = hs.DeviceExistsAddress(OneWireSensor.dvAddress, False)
              Dim dv As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(dev_ref)
              If Not dv Is Nothing Then
                OneWireSensor.dvNoLog = dv.MISC_Check(hs, dvMISC.NO_LOG)
                OneWireSensor.dvName = String.Format("{0} {1} {2} {3}", dv.Location(hs), dv.Location2(hs), dv.Name(hs), dv.Device_Type_String(hs))

                If OneWireSensor.dvNoLog = False Then
                  If OneWireSensor.Diff >= 1 Then
                    Dim dv_format As String = "F0"
                    Select Case OneWireSensor.sensorType
                      Case "Humidity", "Pressure", "Temperature", "Voltage", "Counter"
                        dv_format = "F2"
                    End Select

                    Dim dv_value As String = dblSensorValue.ToString(dv_format)
                    Dim dv_string As String = hs.DeviceVSP_GetStatus(dev_ref, dblSensorValue, ePairStatusControl.Status)
                    hs.SetDeviceValueByRef(dv_ref, dblSensorValue, True)
                    Dim dv_logmsg As String = String.Format("Device: <font color=""{0}"">{1}</font> Set to <font color=""{2}"">{3} ({4})</font>", "#000080", OneWireSensor.dvName, "#008000", dv_value, dv_string)
                    WriteMessage(dv_logmsg, MessageType.Informational)
                    OneWireSensor.ResetDiff()
                  End If
                End If
              End If
            Catch pEx As Exception

            End Try

          End If

        End If

      Else
        WriteMessage(String.Format("Device address {0} cannot be found.", OneWireSensor.dvAddress), MessageType.Warning)
      End If

    Catch pEx As Exception
      '
      ' Process program exception
      '
      Call ProcessError(pEx, "SetDeviceValue()")

    End Try

  End Sub

End Module
