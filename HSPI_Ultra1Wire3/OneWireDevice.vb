<Serializable()> _
Public Class OneWireDevice

  Private m_device_id As Integer = 0
  Private m_device_name As String = ""
  Private m_device_type As String = ""
  Private m_device_conn As String = ""
  Private m_device_addr As String = ""

  Public Sub New()

  End Sub

  Public Sub New(ByVal device_id As Integer, device_name As String, device_type As String, device_conn As String, device_addr As String)

    Me.m_device_id = m_device_id
    Me.m_device_name = m_device_name
    Me.m_device_type = m_device_type
    Me.m_device_conn = m_device_conn
    Me.m_device_addr = m_device_addr

  End Sub

  Public Property device_id As Integer
    Get
      Return m_device_id
    End Get

    Set(value As Integer)
      Me.m_device_id = value
    End Set
  End Property

  Public Property device_name As String
    Get
      Return m_device_name
    End Get

    Set(value As String)
      Me.m_device_name = value
    End Set
  End Property

  Public Property device_type As String
    Get
      Return m_device_type
    End Get

    Set(value As String)
      Me.m_device_type = value
    End Set
  End Property

  Public Property device_conn As String
    Get
      Return m_device_conn
    End Get

    Set(value As String)
      Me.m_device_conn = value
    End Set
  End Property

  Public Property device_addr As String
    Get
      Return m_device_addr
    End Get

    Set(value As String)
      Me.m_device_addr = value
    End Set
  End Property

End Class
