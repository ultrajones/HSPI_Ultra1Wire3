Public Class OneWireSensor
  Inherits OneWireSensorConfig

  Private m_dv_name As String = String.Empty
  Private m_dv_addr As String = String.Empty
  Private m_dv_nolog As Boolean = False

  Private m_value As Double = 0.0
  Private m_value_ts As DateTime = DateTime.Now

  Private m_initialized As Boolean = False

  Private _valueMin As Double = 0.0
  Private _valueMax As Double = 0.0

  Private _diffMin As Double = 0.0
  Private _diffMax As Double = 0.0

  Sub New()

    MyBase.New()

  End Sub

  Sub New(sensorId As Integer, deviceId As Integer, sensorAddr As String, sensorType As String, sensorChannel As String)

    MyBase.New(sensorId, deviceId, sensorAddr, sensorType, sensorChannel)

    _valueMin = 0
    _valueMax = 0

    _diffMin = 0
    _diffMax = 0

  End Sub

  ''' <summary>
  ''' Get/Set the Sensor Value
  ''' </summary>
  ''' <returns></returns>
  Public Property Value As Double
    Get
      Return m_value
    End Get
    Set(value As Double)
      m_value = value
      m_value_ts = DateTime.Now

      If m_initialized = False Then
        Call ResetCounters()
        Call ResetDiff()
        m_initialized = True
      End If

      If value < _valueMin Then _valueMin = value
      If value > _valueMax Then _valueMax = value

      If value < _diffMin Then _diffMin = value
      If value > _diffMax Then _diffMax = value
    End Set
  End Property

  ''' <summary>
  ''' Get/Set the Sensor Value Timestamp
  ''' </summary>
  ''' <returns></returns>
  Public Property ValueTs As DateTime
    Get
      Return m_value_ts
    End Get
    Set(value As DateTime)
      m_value_ts = value
    End Set
  End Property

  ''' <summary>
  ''' Gets the Database Sensor Value
  ''' </summary>
  ''' <returns></returns>
  Public ReadOnly Property DBValue() As Double
    Get
      If MyBase.sensorType = "Counter" Then
        Dim Counter As Double = _valueMax - _valueMin
        ResetCounters()
        Return Counter
      Else
        Return m_value
      End If
    End Get
  End Property

  ''' <summary>
  ''' Gets the Database Sensor Value
  ''' </summary>
  ''' <returns></returns>
  Public ReadOnly Property Diff() As Double
    Get
      Return _diffMax - _diffMin
    End Get
  End Property

  ''' <summary>
  ''' Gets the Sensor Device Address
  ''' </summary>
  ''' <returns></returns>
  Public Property dvAddress As String
    Get
      Return m_dv_addr
    End Get
    Set(value As String)
      m_dv_addr = value
    End Set
  End Property

  ''' <summary>
  ''' Gets the Sensor Device Name
  ''' </summary>
  ''' <returns></returns>
  Public Property dvName As String
    Get
      Return m_dv_name
    End Get
    Set(value As String)
      m_dv_name = value
    End Set
  End Property

  ''' <summary>
  ''' Gets the Sensor Logging Value
  ''' </summary>
  ''' <returns></returns>
  Public Property dvNoLog As Boolean
    Get
      Return m_dv_nolog
    End Get
    Set(value As Boolean)
      m_dv_nolog = value
    End Set
  End Property

  ''' <summary>
  ''' Resets the Sensor Counter Values
  ''' </summary>
  Public Sub ResetDiff()
    _diffMax = Value
    _diffMin = Value
  End Sub

  ''' <summary>
  ''' Resets the Sensor Counter Values
  ''' </summary>
  Private Sub ResetCounters()
    _valueMax = Value
    _valueMin = Value
  End Sub

End Class
