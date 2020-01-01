Imports System.Text.RegularExpressions

Public Class OneWireSensorConfig

  Private m_sensor_id As Integer = 0
  Private m_device_id As Integer = 0
  Private m_sensor_addr As String = ""
  Private m_sensor_name As String = ""
  Private m_sensor_type As String = ""
  Private m_sensor_subtype As String = ""
  Private m_sensor_class As String = ""
  Private m_sensor_device As String = ""
  Private m_sensor_channel As String = ""
  Private m_sensor_color As String = ""
  Private m_sensor_image As String = ""
  Private m_sensor_units As String = ""
  Private m_sensor_resolution As Double = 0.0
  Private m_sensor_reset As Double = 0.0

  Private m_post_enabled As Integer = 0

  Private m_dev_00d As Integer = 0
  Private m_dev_01d As Integer = 0
  Private m_dev_07d As Integer = 0
  Private m_dev_30d As Integer = 0

  Private m_value As Double = 0

  Sub New()

  End Sub

  Sub New(sensorId As Integer, deviceId As Integer, sensorAddr As String, sensorType As String, sensorChannel As String)

    Me.m_sensor_id = sensorId
    Me.m_device_id = deviceId
    Me.m_sensor_addr = sensorAddr
    Me.m_sensor_type = sensorType
    Me.m_sensor_channel = sensorChannel
    Me.m_sensor_name = GetSensorName(romId)
    Me.m_sensor_device = GetSensorDevice(romId)

  End Sub

  Public ReadOnly Property deviceId
    Get
      Return m_device_id
    End Get
  End Property

  Public ReadOnly Property sensorAddr
    Get
      Return m_sensor_addr
    End Get
  End Property

  Public ReadOnly Property romId
    Get
      Return Regex.Replace(m_sensor_addr, ":.:.", "")
    End Get
  End Property

  Public Property sensorId As Integer
    Get
      Return m_sensor_id
    End Get
    Set(value As Integer)
      Me.m_sensor_id = value
    End Set
  End Property

  Public ReadOnly Property sensorChannel As String
    Get
      Return m_sensor_channel
    End Get
  End Property

  Public ReadOnly Property sensorType As String
    Get
      Return m_sensor_type
    End Get
  End Property

  Public Property sensorName As String
    Get
      Return m_sensor_name
    End Get
    Set(value As String)
      m_sensor_name = value
    End Set
  End Property

  Public Property sensorDevice As String
    Get
      Return m_sensor_device
    End Get
    Set(value As String)
      m_sensor_device = value
    End Set
  End Property

  Public Property sensorSubtype As String
    Get
      Return m_sensor_subtype
    End Get
    Set(value As String)
      m_sensor_subtype = value
    End Set
  End Property

  Public Property sensorClass As String
    Get
      Return m_sensor_class
    End Get
    Set(value As String)
      m_sensor_class = value
    End Set
  End Property

  Public Property sensorColor As String
    Get
      Return m_sensor_color
    End Get
    Set(value As String)
      m_sensor_color = value
    End Set
  End Property

  Public Property sensorImage As String
    Get
      Return m_sensor_image
    End Get
    Set(value As String)
      m_sensor_image = value
    End Set
  End Property

  Public Property sensorUnits As String
    Get
      Return m_sensor_units
    End Get
    Set(value As String)
      m_sensor_units = value
    End Set
  End Property

  Public Property sensorResolution As Double
    Get
      Return m_sensor_resolution
    End Get
    Set(value As Double)
      m_sensor_resolution = value
    End Set
  End Property

  Public Property sensorReset As Double
    Get
      Return m_sensor_reset
    End Get
    Set(value As Double)
      m_sensor_reset = value
    End Set
  End Property

  Public Property postEnabled As Integer
    Get
      Return m_post_enabled
    End Get
    Set(value As Integer)
      m_post_enabled = value
    End Set
  End Property

  Public Property dev_00d As Integer
    Get
      Return m_dev_00d
    End Get
    Set(value As Integer)
      m_dev_00d = value
    End Set
  End Property

  Public Property dev_01d As Integer
    Get
      Return m_dev_01d
    End Get
    Set(value As Integer)
      m_dev_01d = value
    End Set
  End Property

  Public Property dev_07d As Integer
    Get
      Return m_dev_07d
    End Get
    Set(value As Integer)
      m_dev_07d = value
    End Set
  End Property

  Public Property dev_30d As Integer
    Get
      Return m_dev_30d
    End Get
    Set(value As Integer)
      m_dev_30d = value
    End Set
  End Property

  Private Function GetSensorName(ByVal ROMId As String) As String

    Select Case Microsoft.VisualBasic.Right(ROMId, 2)
      Case "10" : Return "Digital Thermometer"
      Case "12" : Return "Switch"
      Case "1D" : Return "Counter"
      Case "26" : Return "Smart Battery Monitor"
      Case "28" : Return "Digital Thermometer"
      Case "7E" : Return "EDS Environmental Sensor"
      Case Else
        If Regex.IsMatch(ROMId, "^(00|D8|80)") = True Then
          Return "EDS Environmental Sensor"
        Else
          Return "Unsupported"
        End If
    End Select

  End Function

  Private Function GetSensorDevice(ByVal ROMId As String) As String

    Select Case Microsoft.VisualBasic.Right(ROMId, 2)
      Case "10" : Return "DS18S20"
      Case "12" : Return "Switch"
      Case "1D" : Return "DS2406"
      Case "26" : Return "DS2438"
      Case "28" : Return "DS18B20"
      Case "7E" : Return "EDS"
      Case Else
        If Regex.IsMatch(ROMId, "^(00|D8|80)") = True Then
          Return "EDS"
        Else
          Return "Unsupported"
        End If
    End Select

  End Function

End Class
