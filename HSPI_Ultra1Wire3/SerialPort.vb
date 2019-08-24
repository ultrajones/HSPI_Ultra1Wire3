'Simple RS232 Serial Port VB .Net class provided by Embedded Data Systems, LLC.
'
'This code should be usable on both desktop and compact framework platforms
'by setting the constant 'COMPACT_FRAMEWORK' at compile time.
'This code may be freely copied providing you are using it to integrated
'products manufactured by Embedded Data Systems.
'This code is provided for sample purposes only, and is without warranty
'or guarantee of any kind.

Option Strict On
Public Class SerialPort
  Implements IDisposable

#If COMPACT_FRAMEWORK Then
  private Declare Function CreateFile Lib "coredll.dll" _
     (ByVal lpFileName As String, ByVal dwDesiredAccess As Int32, _
      ByVal dwShareMode As Int32, ByVal lpSecurityAttributes As IntPtr, _
      ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32, _
      ByVal hTemplateFile As IntPtr) As IntPtr
  Private Declare Function GetCommState Lib "coredll.dll" (ByVal nCid As IntPtr, ByRef lpDCB As DCB) As Boolean
  Private Declare Function SetCommState Lib "coredll.dll" (ByVal nCid As IntPtr, ByRef lpDCB As DCB) As Boolean
  Private Declare Function GetCommTimeouts Lib "coredll.dll" (ByVal hFile As IntPtr, ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean
  Private Declare Function SetCommTimeouts Lib "coredll.dll" (ByVal hFile As IntPtr, ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean
  Private Declare Function WriteFile Lib "coredll.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToWrite As Int32, ByRef lpNumberOfBytesWritten As Int32, ByVal lpOverlapped As IntPtr) As Boolean
  Private Declare Function ReadFile Lib "coredll.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToRead As Int32, ByRef lpNumberOfBytesRead As Int32, ByVal lpOverlapped As IntPtr) As Boolean
  Private Declare Function CloseHandle Lib "coredll.dll" (ByVal hObject As IntPtr) As Boolean
#Else
  Private Declare Auto Function CreateFile Lib "kernel32.dll" _
       (ByVal lpFileName As String, ByVal dwDesiredAccess As Int32, _
        ByVal dwShareMode As Int32, ByVal lpSecurityAttributes As IntPtr, _
        ByVal dwCreationDisposition As Int32, ByVal dwFlagsAndAttributes As Int32, _
        ByVal hTemplateFile As IntPtr) As IntPtr
  Private Declare Auto Function GetCommState Lib "kernel32.dll" (ByVal nCid As IntPtr, ByRef lpDCB As DCB) As Boolean
  Private Declare Auto Function SetCommState Lib "kernel32.dll" (ByVal nCid As IntPtr, ByRef lpDCB As DCB) As Boolean
  Private Declare Auto Function GetCommTimeouts Lib "kernel32.dll" (ByVal hFile As IntPtr, ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean
  Private Declare Auto Function SetCommTimeouts Lib "kernel32.dll" (ByVal hFile As IntPtr, ByRef lpCommTimeouts As COMMTIMEOUTS) As Boolean
  Private Declare Auto Function WriteFile Lib "kernel32.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToWrite As Int32, ByRef lpNumberOfBytesWritten As Int32, ByVal lpOverlapped As IntPtr) As Boolean
  Private Declare Auto Function ReadFile Lib "kernel32.dll" (ByVal hFile As IntPtr, ByVal lpBuffer As Byte(), ByVal nNumberOfBytesToRead As Int32, ByRef lpNumberOfBytesRead As Int32, ByVal lpOverlapped As IntPtr) As Boolean
  Private Declare Auto Function CloseHandle Lib "kernel32.dll" (ByVal hObject As IntPtr) As Boolean
#End If

  Public Const NOPARITY As Int32 = 0    'No Parity
  Public Const ONESTOPBIT As Int32 = 0  '1 Stop Bit

  Private Structure DCB
    Public DCBlength As Int32
    Public BaudRate As Int32
    Public fBitFields As Int32 'See Comments in Win32API.Txt
    Public wReserved As Int16
    Public XonLim As Int16
    Public XoffLim As Int16
    Public ByteSize As Byte
    Public Parity As Byte
    Public StopBits As Byte
    Public XonChar As Byte
    Public XoffChar As Byte
    Public ErrorChar As Byte
    Public EofChar As Byte
    Public EvtChar As Byte
    Public wReserved1 As Int16 'Reserved; Do Not Use
  End Structure
  Private fBinary As Integer = 1  'binary mode, no EOF check
  Private fParity As Integer = 2 '  enable parity checking
  Private fOutxCtsFlow As Integer = 4 'CTS output flow control
  Private fOutxDsrFlow As Integer = 8 'DSR output flow control
  Private fDTRControlDisable As Integer = 0 ' DTR flow control type (2 bits)
  Private fDtrControlEnable As Integer = 16 ' Enables DTR Line, and leaves it on.
  Private fDTRControlHandshake As Integer = 32 'Standard DTR handshaking
  Private fDsrSensitivity As Integer = 64 'DSR sensitivity
  Private fTXContinueOnXoff As Integer = 128 'XOFF continues Tx
  Private fOutX As Integer = 256 'XON/XOFF out flow control
  Private fInX As Integer = 512 'XON/XOFF in flow control
  Private fErrorChar As Integer = 1024 'enable error replacement
  Private fNull As Integer = 2048 'enable null stripping
  Private fRTSControlDisable As Integer = 0 'Flow control Line (2 bits)
  Private fRTSControlEnable As Integer = 4096 'Enable RTS Flow control line, and leave it on
  Private fRTSControlHandshake As Integer = 8192 'Enables standard RTS handshaking
  Private fAbortOnError As Integer = 16384 'abort reads/writes on error
  Private fDummy2 As Integer = 32768 'reserved 

  Private hSerialPort As IntPtr
  Private MyDCB As DCB
  Private MyCommTimeouts As COMMTIMEOUTS
  Private oEncoder As New System.Text.ASCIIEncoding
  Private oEnc As System.Text.Encoding = Text.Encoding.GetEncoding(1252)
  Private portName As String

  Private Const GENERIC_READ As Int32 = &H80000000
  Private Const GENERIC_WRITE As Int32 = &H40000000
  Private Const OPEN_EXISTING As Int32 = 3
  Private Const FILE_ATTRIBUTE_NORMAL As Int32 = &H80

  Private Structure COMMTIMEOUTS
    Public ReadIntervalTimeout As Int32
    Public ReadTotalTimeoutMultiplier As Int32
    Public ReadTotalTimeoutConstant As Int32
    Public WriteTotalTimeoutMultiplier As Int32
    Public WriteTotalTimeoutConstant As Int32
  End Structure

  'Opens the serial port using the parameters supplied.
  Public Sub New(ByVal portName As String, ByVal baudRate As Integer, ByVal parity As Byte, ByVal byteSize As Byte, ByVal stopBits As Byte)
    Me.portName = portName
    openPort(portName, baudRate, parity, byteSize, stopBits)
  End Sub
  'Transmits up to 48 bytes out the port
  Public Sub tx(ByVal whatToSend As String)
    Dim bytesWritten As Int32
    Dim outbuffer(48) As Byte
    Dim success As Boolean

    outbuffer = oEnc.GetBytes(whatToSend)
    success = WriteFile(hSerialPort, outbuffer, outbuffer.Length, bytesWritten, IntPtr.Zero)

    If success = False Then
      Throw New CommException("Unable to write to " & Me.portName)
    End If

  End Sub
  'Reads from port until some data is recieved, or timeOut is exceeded
  'Timeout is currently hardcoded at 3000 mS, in openPort()
  Public Function rx() As String
    Dim bytesRead As Int32
    Dim inBuffer(48) As Byte
    Dim retVal As String
    Dim success As Boolean

    success = ReadFile(hSerialPort, inBuffer, inBuffer.Length, bytesRead, IntPtr.Zero)
    If success = False Then
      Throw New CommException("Unable to read from " & portName)
    End If
    retVal = oEnc.GetString(inBuffer, 0, bytesRead)
    Return retVal
  End Function

  Private Sub openPort(ByVal portName As String, ByVal baudRate As Integer, ByVal parity As Byte, ByVal byteSize As Byte, ByVal stopBits As Byte)

    Dim Success As Boolean
    ' Obtain a handle to the serial port.
    hSerialPort = CreateFile(portName, GENERIC_READ Or GENERIC_WRITE, 0, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero)

    ' Verify that the obtained handle is valid.
    If hSerialPort.ToInt32 = -1 Then
      Throw New CommException("Unable to obtain a handle to the COM port")
    End If

    ' Retrieve the current control settings.
    Success = GetCommState(hSerialPort, MyDCB)
    If Success = False Then
      Throw New CommException("Unable to retrieve the current control settings")
    End If

    ' Modify the properties of MyDCB 
    MyDCB.BaudRate = baudRate
    MyDCB.ByteSize = byteSize
    MyDCB.Parity = parity
    MyDCB.StopBits = stopBits
    MyDCB.fBitFields = MyDCB.fBitFields Or fRTSControlDisable Or fDTRControlDisable

    ' Reconfigure port based on the properties of MyDCB.
    Success = SetCommState(hSerialPort, MyDCB)
    If Success = False Then
      Throw New CommException("Unable to reconfigure COM")
    End If

    ' Retrieve the current time-out settings.
    Success = GetCommTimeouts(hSerialPort, MyCommTimeouts)
    If Success = False Then
      Throw New CommException("Unable to retrieve current time-out settings")
    End If

    ' Modify the properties of MyCommTimeouts
    MyCommTimeouts.ReadIntervalTimeout = 5
    MyCommTimeouts.ReadTotalTimeoutConstant = 3000
    MyCommTimeouts.ReadTotalTimeoutMultiplier = 0
    MyCommTimeouts.WriteTotalTimeoutConstant = 0
    MyCommTimeouts.WriteTotalTimeoutMultiplier = 0

    ' Reconfigure the time-out settings, based on the properties of MyCommTimeouts.
    Success = SetCommTimeouts(hSerialPort, MyCommTimeouts)
    If Success = False Then
      Throw New CommException("Unable to reconfigure the time-out settings")
    End If

    'Flush any garbage on port
    flushComIn()
  End Sub

  'Closes the physical port, and prepares this object for destruction
  Public Sub Dispose() Implements System.IDisposable.Dispose
    Static disposed As Boolean = False
    ' Release the handle to port.
    CloseHandle(hSerialPort)
    disposed = True
  End Sub

  'Flushes any data already received on the port
  Public Sub flushComIn()
    Dim ExistingCommTimeouts As COMMTIMEOUTS
    Dim tempCommTimeouts As COMMTIMEOUTS
    Dim success As Boolean
    Dim BytesRead As Int32
    Dim Buffer(32) As Byte

    ' Retrieve the current time-out settings.
    success = GetCommTimeouts(hSerialPort, ExistingCommTimeouts)

    With tempCommTimeouts
      .ReadIntervalTimeout = 1
      .WriteTotalTimeoutConstant = 0
      .WriteTotalTimeoutMultiplier = 0
      .ReadTotalTimeoutConstant = 1
    End With

    ' Temporarily reconfigure the time-out settings
    success = SetCommTimeouts(hSerialPort, tempCommTimeouts)

    'Read data from port.
    success = ReadFile(hSerialPort, Buffer, 32, BytesRead, IntPtr.Zero)

    'Reset the time-out settings to previous values
    success = SetCommTimeouts(hSerialPort, ExistingCommTimeouts)

  End Sub

  Class CommException
    Inherits ApplicationException

    Sub New(ByVal Reason As String)
      MyBase.New(Reason)
    End Sub

  End Class

End Class
