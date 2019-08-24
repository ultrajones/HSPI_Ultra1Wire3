' DS18B20 is a VB.Net class that demonstrates how to work with
' a Dallas Semiconductor DS18B20 Temperature Sensor using an
' EDS 1-Wire Bus Master.
' 
' HARDWARE DESCRIPTION
' 
' The DS18B20 Digital Thermometer provides 9 to 12–bit
' centigrade temperature measurements and has an alarm function
' with nonvolatile user-programmable upper and lower trigger
' points. The DS18B20 communicates over a 1-Wire bus that by
' definition requires only one data line (and ground) for
' communication with a central microprocessor. It has an
' \operating temperature range of –55°C to +125°C and is
' accurate to 0.5°C over the range of –10°C to +85°C. In
' addition, the DS18B20 can derive power directly from the data
' line (“parasite power”), eliminating the need for an external
' power supply. Each DS18B20 has a unique 64-bit serial code,
' which allows multiple DS18B20s to function on the same 1–wire
' bus; thus, it is simple to use one microprocessor to control
' many DS18B20s distributed over a large area. Applications
' that can benefit from this feature include HVAC environmental
' controls, temperature monitoring systems inside buildings,
' equipment or machinery, and process monitoring and control
' systems.
' 
' TYPICAL USAGE:
' 
' A typical usage scenario would look like:
' 
'     * Discover and address the device.
'     * Configure desired conversion resolution by updating the
'       configuration register in the DS18B20's scratchpad.
'     * Perform a temperature conversion.
'     * Read the scratchpad from the DS18S20, which contains
'       the temperature information in the first two bytes.
'     * Parse the temperature information out of the
'       scratchpad. (This varies depending on the bits of resolution)
' 
' 
' 
' The following code demonstrates a typical usage of this
' class:
' <CODE>
'     Dim temperature As Single
'     For Each ds18B20 As EmbeddedDataSystems.DS18B20 In ds18B20s
'       txtResults.Text = "Reading Temperature..." & vbCrLf
'       txtResults.Refresh()
'       temperature = ds18B20.GetTemperature(True, 12) 'Farenheight, w/12 bit resolution
'       txtResults.Text = txtResults.Text & "Current Temp: " & CStr(temperature) & " F" & vbCrLf
'     Next
' </CODE>                                                                                       

Public Class DS18B20

  Inherits Dallas1WireDevice

  Const PROTOCOL_CONVERT_T As String = "44"         'DS18B20 Function Command "CONVERT T"
  Const PROTOCOL_READ_SCRATCHPAD As String = "BE"   'DS18B20 Function Command "READ SCRATCHPAD"
  Const PROTOCOL_WRITE_SCRATCHPAD As String = "4E"  'DS18B20 Function Command "WRITE SCRATCHPAD"
  Const PROTOCOL_COPY_SCRATCHPAD As String = "48"   'DS18B20 Function Command "COPY SCRATCHPAD"
  Const PROTOCOL_RECALL_E2 As String = "B8"         'DS18B20 Function Commnad "RECALL E2"
  Const PROTOCOL_READ_POWER_SUPPLY As String = "B4" 'DS18B20 Function Command "READ POWER SUPPLY"

  Const T_CONV_12Bit As Integer = 750 'Maximum time (mS) required for 12 Bit DS18S20 temperature conversion.
  Const T_CONV_11Bit As Integer = 375 'Maximum time (mS) required for 11 Bit DS18S20 temperature conversion.
  Const T_CONV_10Bit As Integer = 188 'Maximum time (mS) required for 10 Bit DS18S20 temperature conversion.
  Const T_CONV_9Bit As Integer = 94   'Maximum time (mS) required for 9 Bit DS18S20 temperature conversion.

  Public Sub New(ByVal ROMId As String, ByVal busMaster As OWInterface.BusMasterInterface)
    MyBase.New(ROMId, busMaster)
  End Sub

  ' Used to initiate a temperature conversion on the DS18B20.
  ' 
  ' Since the DS18B20's are very power conservative, they do not
  ' continuously sample the temperature. Normally, the DS18B20
  ' will sit idle consuming almost no power until you
  ' specifically tell it to perform a temperature conversion,
  ' which is what this function accomplishes.
  '  
  ' 
  ' \Note that the act of performing a temperature conversion
  ' does not in itself give you the temperature data. Instead,
  ' the sensor will perform a conversion, and store the resulting
  ' data in its internal scratchpad memory. After the conversion
  ' is complete, you have to explicitly read the scratchpad data
  ' from the DS18B20 in order to determine the temperature at the
  ' time it was sampled.
  '                                                              
  Public Sub ConvertTemperature(Optional ByVal resolution As Byte = 0)

    Dim response As OW_Response

    Me.BusMaster.log("DS18B20- Initiating a temperature conversion. " & vbCrLf)
    If resolution > 0 Then SetConversionResolution(resolution)
    With BusMaster
      response = .OW_WriteBlock(PROTOCOL_CONVERT_T, Me.ROMId)
      'EDS Bus Masters automatically provide strong	pullup for the conversion.
      'Just	need to	sleep	for	the	duration of	the	conversion so	the	sensor is	provided
      'adequate	power.

      'Note that the DS18B20 has different max conversion times depending
      'on the coversion resolution.  This sample makes no attempt to alter
      'the wait time based on the configured resolution, but this would be
      'an obvious optimization to be applied to production code.
      System.Threading.Thread.Sleep(T_CONV_12Bit)
      .OW_ResetBus()
    End With

  End Sub

  'This will write the new conversion resolution into the DS18B20's configuration register
  Public Sub SetConversionResolution(ByVal newResolution As Byte)

    Dim scratchPad As DS18B20.ScratchPad
    'Read existing scratchpad data in order to preserve any alarm settings
    'when writing new scratchpad contents
    Me.BusMaster.log("DS18B20- Setting conversion resolution to " & newResolution & " bits." & vbCrLf)
    scratchPad = ReadScratchpad()
    'Update byte 4 of the scratchpad with the new conversion resolution
    Select Case newResolution
      Case Is = 9
        scratchPad.ScratchPadBytes(4) = 31
      Case Is = 10
        scratchPad.ScratchPadBytes(4) = 63
      Case Is = 11
        scratchPad.ScratchPadBytes(4) = 95
      Case Is = 12
        scratchPad.ScratchPadBytes(4) = 127
    End Select
    WriteScratchpad(scratchPad)

  End Sub

  ' This command allows the master to read the contents of the
  ' scratchpad. The data transfer starts with the least
  ' significant bit of byte 0 and continues through the
  ' scratchpad until the 9th byte (byte 8 – CRC) is read. The
  ' master may issue a reset to terminate reading at any time if
  ' \only part of the scratchpad data is needed.
  ' 
  ' Returns
  ' DS18B20.ScratchPad - ScratchPad object containing both the
  ' raw scratchpad data, and fields representing the individual
  ' pieces of data parsed from the scratchpad.                  
  Public Function ReadScratchpad() As DS18B20.ScratchPad

    Dim response As OW_Response
    Dim retVal As DS18B20.ScratchPad

    Me.BusMaster.log("DS18B20- Reading scratchpad from physical device. " & vbCrLf)
    With BusMaster
      'Write the 'BE' command to the DS18B20, followed by 9 bytes of time slots to read the reply from the DS18B20.
      response = .OW_WriteBlock(PROTOCOL_READ_SCRATCHPAD & "FFFFFFFFFFFFFFFFFF", Me.ROMId)
      .OW_ResetBus()
    End With
    retVal = New DS18B20.ScratchPad(Mid(CType(response.Data(0), String), 3, 18))
    Return retVal

  End Function

  ' This command allows the master to write 3 bytes of data to
  ' the DS18B20’s scratchpad. The first byte is written into the
  ' TH register (byte 2 of the scratchpad), and the second byte
  ' is written into the TL register (byte 3 of the scratchpad).
  ' The 3rd byte is written to the configuration register.
  ' Data must be transmitted least significant bit first. Both
  ' bytes MUST be written before the master issues a reset, or
  ' the data may be corrupted.                                  
  Public Sub WriteScratchpad(ByVal scratchPad As DS18B20.ScratchPad)

    Dim response As OW_Response
    Me.BusMaster.log("DS18B20- Writing scratchpad to physical device. " & vbCrLf)
    With BusMaster
      response = .OW_WriteBlock(PROTOCOL_WRITE_SCRATCHPAD & Hex2(scratchPad.ScratchPadBytes(2)) _
                  & Hex2(scratchPad.ScratchPadBytes(3)) & Hex2(scratchPad.ScratchPadBytes(4)), Me.ROMId)
      .OW_ResetBus()
    End With

  End Sub

  ' This command copies the contents of the hardware's scratchpad
  ' TH, TL, and configuration registers (bytes 2, 3, and 4) to EEPROM. If the device
  ' is being used in parasite power mode, within 10 µs (max)
  ' after this command is issued the master must enable a strong
  ' pullup on the 1-wire bus for at least 10 ms.                 
  Public Sub CopyScratchpad()
    'TODO - Implement the CopyScratchpad function
  End Sub

  ' This command recalls the alarm trigger values (TH and TL)
  ' from EEPROM and places the data in bytes 2 and 3,
  ' respectively, in the scratchpad memory. The master device can
  ' issue read time slots following the Recall E2 command and the
  ' DS18S20 will indicate the status of the recall by
  ' transmitting 0 while the recall is in progress and 1 when the
  ' recall is done. The recall operation happens automatically at
  ' powerup, so valid data is available in the scratchpad as soon
  ' as power is applied to the device.                           
  'Public Sub RecallE2()
  '  'TODO - Implement the RecallE2 function
  'End Sub
  ' Used to determine whether the DS18S20 is operating from
  ' parasite power, or whether it is being powered externally.
  ' 
  ' In some situations the bus master may not know whether the
  ' DS18S20s on the bus are parasite powered or powered by
  ' external supplies. The master needs this information to
  ' determine if the strong bus pullup should be used during
  ' temperature conversions. To get this information, the master
  ' can address the device and issue a Read Power Supply [B4h]
  ' command followed by a “read-time slot”. During the read-time
  ' slot, parasite powered DS18S20s will pull the bus low, and
  ' externally powered DS18S20s will let the bus remain high.
  ' 
  ' Note
  ' The use of parasite power is not reccommended for
  ' temperatures above 100 C since the DS18S20 may not be able to
  ' sustain communications due to higher leakage currents that
  ' can exist at those temperatures.
  ' 
  ' Returns
  ' 0 - if sensor is being parasitically powered, 1 - if sensor is
  ' being externally powered                                     
  Public Function ReadPowerSupply() As Byte

    Dim response As OW_Response

    Me.BusMaster.log("DS18B20- Reading power supply status. " & vbCrLf)
    With BusMaster
      'Sends the 'B4' function command to the DS18S20, followed by 8 read time slots
      response = .OW_WriteBlock(PROTOCOL_READ_POWER_SUPPLY & "FF", Me.ROMId)
      .OW_ResetBus()
    End With
    Return CByte("&h" & Mid(response.Data(0).ToString(), 3, 2))

  End Function

  ' Reads the temperature data stored in the DS18S20's scratchpad
  ' during the last temperature conversion.
  ' 
  ' Note
  ' This function does not initiate a new temperature conversion.
  ' Also, this function assumes that the values currently in the
  ' DS18B20's configuration register are the same values that were
  ' there when the temperature convertion took place.  Since it is
  ' possible that another process might alter the configuration register
  ' without detection by this process, this assumption
  ' can only be guaranteed valid if the ConvertTemperature->ReadScratchpad
  ' process is performed within a single lock of the 1-Wire bus.
  ' 
  ' Parameters
  ' Farenheight :  If True, temperature will be returned in Farehenheight
  '                units, otherwise Centegrade units will be used.       
  Public Function GetLastTemperature(Optional ByVal Farenheight As Boolean = False) As Single
    'Read	scratchpad
    Dim scratchpad As DS18B20.ScratchPad
    scratchpad = Me.ReadScratchpad()

    'Return	the	Temperature	Data contained in	the	scratchpad
    Return scratchpad.calcTemperature(Farenheight)

  End Function

  ' Initiates a temperature conversion and then reads the temperature data from
  ' the DS18S20's scratchpad.
  ' This function effectively combines ConvertTemperature() and GetLastTemperature() into
  ' a single call
  '
  ' Note- If Resolution is specified (9-12), then the sensor will be reconfigured for that resolution prior
  ' to performing the temperature conversion.  If not specified, then
  ' the device's current configuration will be used.
  Public Function GetTemperature(Optional ByVal Farenheight As Boolean = False, _
                                 Optional ByVal Resolution As Byte = 0) As Single
    ' TODO - Restructure this to utilize a single outer lock instead of having
    '        each of the following function calls acquire and release individual locks.
    Me.ConvertTemperature(Resolution)
    Return Me.GetLastTemperature(Farenheight)
  End Function

  ' Class for working with the DS18B20's 9 byte scratchpad.
  ' 
  ' 
  ' 
  ' HARDWARE DESCRIPTION:
  ' 
  ' The DS18B20's memory consists of an SRAM scratchpad with
  ' nonvolatile EEPROM storage for the high and low alarm trigger
  ' registers (TH and TL) and configuration register.
  ' 
  ' Byte 0 and byte 1 of the scratchpad contain the LSB and the
  ' MSB of the temperature register, respectively. These bytes
  ' are read-only. Bytes 2 and 3 provide access to TH and TL
  ' registers. Byte 4 contains the configuration register data,
  ' which is explained in detail in the CONFIGURATION REGISTER
  ' section of the DS18B20 datasheet. Bytes 5, 6, and 7 are
  ' reserved for internal use by the device and cannot be
  ' overwritten; these bytes will return all 1s when read. Byte
  ' 8 of the scratchpad is read-only and contains the cyclic
  ' redundancy check (CRC) code for bytes 0 through 7 of the
  ' scratchpad. The DS18B20 generates this CRC using the method
  ' described in the CRC GENERATION section of the DS18B20
  ' datasheet.
  ' 
  ' Data in the EEPROM registers is retained when the device is
  ' powered down; at power-up the EEPROM data is reloaded into
  ' the corresponding scratchpad locations. Data can also be
  ' reloaded from EEPROM to the scratchpad at any time using the
  ' Recall E2 [B8h] command.                                     
  Public Class ScratchPad

    Public ScratchPadBytes(9) As Byte      'Scratchpad bytes
    Public RawBytes As String 'Raw, unparsed bytes read from scratchpad

    'Constructor
    'Requires the rawbytes read from the physical device
    'Parameters:
    ' rawBytes - Bytes read from the scratchpad using the DS18B20's 'BE' command.
    Public Sub New(ByVal rawBytes As String)

      Me.RawBytes = rawBytes
      parseRawBytes()

    End Sub
    Public Function ValidateCRC() As Boolean
      'TODO	-	Implement	DS18B20.ScratchPad.ValidateCRC()
      Return True
    End Function

    'Parses	the	raw	bytes	and	populates	the
    'public	byte array
    Protected Sub parseRawBytes()

      Me.ScratchPadBytes(0) = CByte("&H" & Mid(RawBytes, 1, 2))
      Me.ScratchPadBytes(1) = CByte("&H" & Mid(RawBytes, 3, 2))
      Me.ScratchPadBytes(2) = CByte("&H" & Mid(RawBytes, 5, 2))
      Me.ScratchPadBytes(3) = CByte("&H" & Mid(RawBytes, 7, 2))
      Me.ScratchPadBytes(4) = CByte("&H" & Mid(RawBytes, 9, 2))
      Me.ScratchPadBytes(5) = CByte("&H" & Mid(RawBytes, 11, 2))
      Me.ScratchPadBytes(6) = CByte("&H" & Mid(RawBytes, 13, 2))
      Me.ScratchPadBytes(7) = CByte("&H" & Mid(RawBytes, 15, 2))
      Me.ScratchPadBytes(8) = CByte("&H" & Mid(RawBytes, 17, 2))

    End Sub

    ' Description
    ' Calculates the temperature from the existing scratchpad data.
    ' 
    ' The resolution of the temperature sensor is user-configurable
    ' to 9, 10, 11, or 12 bits, corresponding to increments of
    ' 0.5C, 0.25C, 0.125C, and 0.0625C, respectively. The
    ' default resolution at power-up is 12-bit. The current
    ' configuration can be determined by reading the configuration
    ' register, which is contained in the DS18B20's scratchpad.
    ' 
    ' The DS18B20 output temperature data is calibrated in degrees
    ' centigrade; for Fahrenheit applications, a lookup table or
    ' conversion routine must be used. The temperature data is
    ' stored as a 16-bit sign-extended two’s complement number in
    ' the temperature register (see Figure 2). The sign bits (S)
    ' indicate if the temperature is positive or negative: for
    ' positive numbers S = 0 and for negative numbers S = 1. If the
    ' DS18B20 is configured for 12-bit resolution, all bits in the
    ' temperature register will contain valid data. For 11-bit
    ' resolution, bit 0 is undefined. For 10-bit resolution, bits 1
    ' and 0 are undefined, and for 9-bit resolution bits 2, 1 and 0
    ' are undefined.
    ' 
    ' Returns
    ' Temperature in either Centegrade or Farenheight.
    ' 
    ' Parameters
    ' Farenheight :  Flag to indicate if the temperature is to be returned
    '                in Farenheight. If not True, then the returned value
    '                will be in Centegrade.
    ' 
    ' See Also
    ' DS18B20.ConvertTemperature(), DS18B20.GetLastTemperature(),
    ' DS18B20.GetTemperature()                                            
    Public Function calcTemperature(Optional ByVal Farenheight As Boolean = False) As Decimal

      Dim celcius As Decimal
      Dim tempRead As Int16, tempSign As Integer
      Dim temperature_LSB As Byte, temperature_MSB As Byte, bitsOfResolution As Byte
      Dim bitMask As Int16

      temperature_LSB = ScratchPadBytes(0)
      temperature_MSB = ScratchPadBytes(1)

      'Have a look at the 'S' bits of the Temperature Register to determine the
      'whether the measured temperature is positive or negative
      'Combine the LS and MS Temperature register bytes
      tempRead = (CType(temperature_MSB, Int16) << 8) Or temperature_LSB

      If (temperature_MSB And 248) > 0 Then
        'Measured temperature is negative
        tempSign = -1
        tempRead = ((Not tempRead) + 1) And &HFFFF
      Else
        'Measured temperature is positive
        tempSign = 1
      End If

      'Look at the configuration register to determine how many bits of resolution
      'were used during the conversion.  Based on this, construct a bitmask 
      'that is used to mask off any undefined bits in the LS byte
      'of the temperature register.
      'See page 3 of the DS18B20 spec sheet.
      bitsOfResolution = Me.ParseResolution()
      bitMask = 2040      'Will mask off the sign bits
      Select Case bitsOfResolution
        Case Is = 12
          bitMask = bitMask Or 7          'Will pickup all bits of the LS byte
        Case Is = 11
          bitMask = bitMask Or 6          'Will mask out bit 0 of the LS byte
        Case Is = 10
          bitMask = bitMask Or 4          'Will mask out bits 0 and 1 of the LS byte
        Case Is = 9
          bitMask = bitMask Or 0          'Will mask out bits 0, 1, and 2 of the LS byte
      End Select

      'Mask off the sign bits and any bits undefined by the current resolution
      tempRead = tempRead And bitMask

      'Now convert the resulting count to Celcius based on the resolution
      'See page 3 of the DS18B20 manual
      celcius = tempRead * 0.0625@

      'Apply the correct sign, determined previously
      celcius = celcius * tempSign

      If Farenheight Then
        Return (celcius * 1.8@) + 32@
      Else
        Return celcius
      End If
    End Function

    'Looks at the configuration register of the scratchpad
    'and returns the number of bits of resolution that the
    'DS18B20 is configured for.
    Public Function ParseResolution() As Byte

      Dim retVal As Byte

      Select Case Me.ScratchPadBytes(4)
        Case Is = 31
          retVal = 9
        Case Is = 63
          retVal = 10
        Case Is = 95
          retVal = 11
        Case Is = 127
          retVal = 12
        Case Else
          retVal = 0
      End Select
      Return retVal

    End Function

  End Class

End Class
