' DS18S20 is a VB.Net class that demonstrates how to work with
' a Dallas Semiconductor DS18S20 Temperature Sensor using an
' EDS 1-Wire Bus Master.
' 
' 
' 
' HARDWARE DESCRIPTION
' 
' The DS18S20 Digital Thermometer provides 9–bit centigrade
' temperature measurements and has an alarm function with
' nonvolatile user-programmable upper and lower trigger points.
' The DS18S20 communicates over a 1-wire bus that by definition
' requires only one data line (and ground) for communication
' with a central microprocessor. It has an operating
' temperature range of –55°C to +125°C and is accurate to
' ±0.5°C over the range of –10°C to +85°C. In addition, the
' DS18S20 can derive power directly from the data line
' (“parasite power”), eliminating the need for an external
' power supply. Each DS18S20 has a unique 64-bit serial code,
' which allows multiple DS18S20s to function on the same 1–wire
' bus; thus, it is simple to use one microprocessor to control
' many DS18S20s distributed over a large area. Applications
' that can benefit from this feature include HVAC environmental
' controls, temperature monitoring systems inside buildings,
' equipment or machinery, and process monitoring and control
' systems.                                                     

Public Class DS18S20

  Inherits Dallas1WireDevice

  Const PROTOCOL_CONVERT_T As String = "44"           'DS18S20 Function Command "CONVERT T"
  Const PROTOCOL_READ_SCRATCHPAD As String = "BE"     'DS18S20 Function Command "READ SCRATCHPAD"
  Const PROTOCOL_WRITE_SCRATCHPAD As String = "4E"    'DS18S20 Function Command "WRITE SCRATCHPAD"
  Const PROTOCOL_COPY_SCRATCHPAD As String = "48"     'DS18S20 Function Command "COPY SCRATCHPAD"
  Const PROTOCOL_RECALL_E2 As String = "B8"           'DS18S20 Function Commnad "RECALL E2"
  Const PROTOCOL_READ_POWER_SUPPLY As String = "B4"   'DS18s20 Function Command "READ POWER SUPPLY"

  Const T_CONV As Integer = 750 'Maximum time required for DS18S20 temperature conversion.

  Public Sub New(ByVal ROMId As String, ByVal busMaster As OWInterface.BusMasterInterface)
    MyBase.New(ROMId, busMaster)
  End Sub

  ' Class for working with the DS18S20's 9 byte scratchpad.
  ' 
  ' 
  ' 
  ' HARDWARE DESCRIPTION:
  ' 
  ' The scratchpad memory contains the 2-byte temperature
  ' register that stores the digital output from the temperature
  ' sensor. In addition, the scratchpad provides access to the
  ' 1-byte upper and lower alarm trigger registers (TH and TL).
  ' The TH and TL registers are nonvolatile (EEPROM), so they
  ' will retain data when the device is powered down.           
  Public Class ScratchPad

    Public ScratchPadBytes(9) As Byte      'Scratchpad bytes
    Public RawBytes As String 'Raw, unparsed bytes read from scratchpad

    'Constructor
    'Requires the rawbytes read from the physical device
    'Parameters:
    ' rawBytes - Bytes read from the scratchpad using the DS18S20's 'BE' command.
    Public Sub New(ByVal rawBytes As String)
      Me.RawBytes = rawBytes
      parseRawBytes()
    End Sub

    Public Function ValidateCRC() As Boolean
      'TODO	-	Implement	DS18x20.ScratchPad.ValidateCRC()
      Return True
    End Function

    'Parses	the	raw	bytes and populates the
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
    ' The DS18S20 output data is calibrated in degrees centigrade;
    ' for Fahrenheit applications, a lookup table or conversion
    ' routine must be used. The temperature sensor output has 9-bit
    ' resolution, which corresponds to 0.5°C steps. The temperature
    ' data is stored as a 16-bit sign-extended two’s complement
    ' number in the temperature register (see Figure 2). The sign
    ' bits (S) indicate if the temperature is positive or negative:
    ' for positive numbers S = 0 and for negative numbers S = 1.
    ' Table 2 gives examples of digital output data and the
    ' corresponding temperature reading. Resolutions greater than 9
    ' bits can be calculated using the data from the temperature,
    ' COUNT REMAIN and COUNT PER °C registers in the scratchpad.
    ' \Note that the COUNT PER °C register is hard-wired to 16
    ' (10h). After reading the scratchpad, the TEMP_READ value is
    ' \obtained by truncating the 0.5°C bit (bit 0) from the
    ' temperature data. The extended resolution temperature can
    ' then be calculated using the following equation:
    ' 
    ' TEMPERATURE = TEMP_READ - 0.25 + ((COUNT_PER_C -
    ' COUNT_REMAIN) / COUNT_PER_C)
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
    ' DS18S20.ConvertTemperature(), DS18S20.GetLastTemperature(),
    ' DS18S20.GetTemperature()                                            
    Public Function calcTemperature(Optional ByVal Farenheight As Boolean = False) As Decimal

      Dim celcius As Decimal
      Dim tempRead As Integer
      Dim countPerC As Byte, countRemain As Byte
      Dim temperature_LSB As Byte, temperature_MSB As Byte

      temperature_LSB = ScratchPadBytes(0)
      temperature_MSB = ScratchPadBytes(1)

      countPerC = ScratchPadBytes(7)
      countRemain = ScratchPadBytes(6)

      tempRead = (CType(temperature_MSB, Int16) << 8) Or temperature_LSB

      celcius = CDec((tempRead - 0.25@ + CDec((CType(countPerC, Integer) - CType(countRemain, Integer))) / countPerC) / 2@)

      If Farenheight Then
        Return (celcius * 1.8@) + 32@
      Else
        Return celcius
      End If

    End Function

  End Class

  ' Used to initiate a temperature conversion on the DS18S20.
  ' 
  ' Since the DS18S20's are very power conservative, they do not
  ' continuously sample the temperature. Normally, the DS18S20
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
  ' from the DS18S20 in order to determine the temperature at the
  ' time it was sampled.
  '                                                              
  Public Sub ConvertTemperature()
    Dim response As OW_Response
    With BusMaster
      response = .OW_WriteBlock(PROTOCOL_CONVERT_T, Me.ROMId)
      ' EDS Bus Masters automatically provide strong pullup for the conversion.
      ' Just need to sleep for the duration of the conversion so the sensor is provided
      ' adequate	power.
      System.Threading.Thread.Sleep(T_CONV)
    End With
  End Sub

  ' This command allows the master to read the contents of the
  ' scratchpad. The data transfer starts with the least
  ' significant bit of byte 0 and continues through the
  ' scratchpad until the 9th byte (byte 8 – CRC) is read. The
  ' master may issue a reset to terminate reading at any time if
  ' \only part of the scratchpad data is needed.
  ' 
  ' Returns
  ' DS18S20.ScratchPad - ScratchPad object containing both the
  ' raw scratchpad data, and fields representing the individual
  ' pieces of data parsed from the scratchpad.                  
  Public Function ReadScratchpad() As DS18S20.ScratchPad
    Dim response As OW_Response
    Dim retVal As DS18S20.ScratchPad
    With BusMaster
      'Write the 'BE' command to the DS18S20, followed by 9 bytes of time slots to read the reply from the DS18S20.
      response = .OW_WriteBlock(PROTOCOL_READ_SCRATCHPAD & "FFFFFFFFFFFFFFFFFF", Me.ROMId)
    End With
    retVal = New DS18S20.ScratchPad(Mid(CType(response.Data(0), String), 3, 18))
    Return retVal
  End Function

  ' This command allows the master to write 2 bytes of data to
  ' the DS18S20’s scratchpad. The first byte is written into the
  ' TH register (byte 2 of the scratchpad), and the second byte
  ' is written into the TL register (byte 3 of the scratchpad).
  ' Data must be transmitted least significant bit first. Both
  ' bytes MUST be written before the master issues a reset, or
  ' the data may be corrupted.                                  
  Public Sub WriteScratchpad()
    'TODO - Implement the WriteScratchpad function
  End Sub

  ' This command copies the contents of the hardware's scratchpad
  ' TH and TL registers (bytes 2 and 3) to EEPROM. If the device
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
    With BusMaster
      'Send the 'B4' function command to the DS18S20, followed by 8 read time slots
      response = .OW_WriteBlock(PROTOCOL_READ_POWER_SUPPLY & "FF", Me.ROMId)
    End With
    Return CByte("&h" & Mid(response.Data(0).ToString(), 3, 2))
  End Function

  ' Reads the temperature data stored in the DS18S20's scratchpad
  ' during the last temperature conversion.
  ' 
  ' Note
  ' This function does not initiate a new temperature conversion.
  ' 
  ' Parameters
  ' Farenheight :  If True, temperature will be returned in Farehenheight
  '                units, otherwise Centegrade units will be used.       
  Public Function GetLastTemperature(Optional ByVal Farenheight As Boolean = False) As Single
    'Read	scratchpad
    Dim scratchpad As DS18S20.ScratchPad
    scratchpad = Me.ReadScratchpad()

    'Return	the	Temperature	Data contained in	the	scratchpad
    Return scratchpad.calcTemperature(Farenheight)
  End Function

  ' Initiates a temperature conversion and then reads the temperature data from
  ' the DS18S20's scratchpad.
  ' This function effectively combines ConvertTemperature() and GetLastTemperature() into
  ' a single call
  Public Function GetTemperature(Optional ByVal Farenheight As Boolean = False) As Single
    ' TODO - Restructure this to utilize a single outer lock instead of having
    '        each of the following function calls acquire and release individual locks.
    Me.ConvertTemperature()
    Return Me.GetLastTemperature(Farenheight)
  End Function

End Class