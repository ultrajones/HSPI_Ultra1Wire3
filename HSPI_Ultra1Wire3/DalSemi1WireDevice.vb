' 
' Provides base functionality common to all 1-Wire devices
'
Public Class Dallas1WireDevice

  Private m_ROMID As String
  Private m_BusMaster As OWInterface.BusMasterInterface

  ' An instance of a Dallas1WireDevice class is instanciated after the
  ' ROM Id for the particular device has been discovered via an
  ' external method. Creating a new Dallas1WireDevice object requires
  ' passing in both the ROM Id of the Dallas1WireDevice, and an object that
  ' represents the 1-Wire bus master through which the physical
  ' 1-Wire device can be communicated with.
  ' 
  ' Parameters
  ' ROMId :      A string containing the 16 byte hex representation of
  '              the 1-Wire Device's 8 byte ROM ID code.
  ' busMaster :  An object representing the physical 1\-Wire bus master
  '              through which this 1-Wire device can be communicated with.  
  Public Sub New(ByVal ROMId As String, ByVal busMaster As OWInterface.BusMasterInterface)
    Me.m_BusMaster = busMaster
    Me.m_ROMID = ROMId
  End Sub

  ''' <summary>
  ''' ROMId - String representation of the 16 byte hex representation of the 1-Wire device's 8 byte ROM Id code.
  ''' </summary>
  ''' <value></value>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public ReadOnly Property ROMId() As String
    Get
      Return Me.m_ROMID
    End Get
  End Property

  Public ReadOnly Property BusMaster() As OWInterface.BusMasterInterface
    Get
      Return Me.m_BusMaster
    End Get
  End Property

End Class