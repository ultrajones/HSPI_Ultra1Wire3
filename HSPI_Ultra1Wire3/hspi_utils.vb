'Imports System.IO
'Imports System.Runtime.Serialization.Formatters
'Imports HomeSeerAPI
Imports System.Data.Common
Imports System.Drawing

Module hspi_utils

  ''' <summary>
  ''' Fixes the filename path
  ''' </summary>
  ''' <param name="strFileName"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function FixPath(ByVal strFileName As String) As String

    Try

      Dim OSType As HomeSeerAPI.eOSType = hs.GetOSType()

      If OSType = HomeSeerAPI.eOSType.linux Then

        strFileName = strFileName.Replace("\", "/")
        strFileName = strFileName.Replace("//", "/")

      Else

        strFileName = strFileName.Replace("/", "\")
        strFileName = strFileName.Replace("\\", "\")

      End If

    Catch pEx As Exception

    End Try

    Return strFileName

  End Function


  ''' <summary>
  ''' Registers a web page with HomeSeer
  ''' </summary>
  ''' <param name="link"></param>
  ''' <param name="linktext"></param>
  ''' <param name="page_title"></param>
  ''' <param name="Instance"></param>
  ''' <remarks></remarks>
  Public Sub RegisterWebPage(ByVal link As String, _
                             Optional linktext As String = "", _
                             Optional page_title As String = "", _
                             Optional Instance As String = "")

    Try

      hs.RegisterPage(link, IFACE_NAME, Instance)

      If linktext = "" Then linktext = link
      linktext = linktext.Replace("_", " ")

      If page_title = "" Then page_title = linktext

      Dim wpd As New HomeSeerAPI.WebPageDesc
      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      callback.RegisterLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterWebPage")

    End Try

  End Sub

  ''' <summary>
  ''' Registers links to AXPX web page
  ''' </summary>
  ''' <param name="link"></param>
  ''' <param name="linktext"></param>
  ''' <param name="page_title"></param>
  ''' <param name="Instance"></param>
  ''' <remarks></remarks>
  Public Sub RegisterASXPWebPage(ByVal link As String, _
                                 linktext As String, _
                                 page_title As String, _
                                 Optional Instance As String = "")

    Try

      Dim wpd As New HomeSeerAPI.WebPageDesc

      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      callback.RegisterLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterASXPWebPage")

    End Try

  End Sub

  Public Sub RegisterHelpPage(ByVal link As String, _
                              linktext As String, _
                              page_title As String, _
                              Optional Instance As String = "")

    Try

      Dim wpd As New HomeSeerAPI.WebPageDesc

      wpd.plugInName = IFACE_NAME
      wpd.plugInInstance = Instance
      wpd.link = link
      wpd.linktext = linktext & Instance
      wpd.page_title = page_title & Instance

      hs.RegisterHelpLink(wpd)

    Catch pEx As Exception
      '
      ' Process the program exception
      '
      ProcessError(pEx, "RegisterHelpPage")

    End Try

  End Sub

  '------------------------------------------------------------------------------------
  'Purpose:   Gets the safe colors used for charting
  'Input:     None
  'Output:    Queue
  '------------------------------------------------------------------------------------
  Public Function GetSafeColors() As Queue

    Static safeColors As New Queue()
    If safeColors.Count > 1 Then Return safeColors

    Try

      '
      ' Get the color names from the Known color enum
      '
      Dim colorNames As String() = [Enum].GetNames(GetType(KnownColor))
      '
      ' Iterate thru each string in the colorNames array
      '
      For Each colorName As String In colorNames
        '
        ' Cast the colorName into a KnownColor
        '
        Dim knownColor As KnownColor = DirectCast([Enum].Parse(GetType(KnownColor), colorName), KnownColor)

        '
        ' Check if the knownColor variable is a System color
        '
        If (knownColor > KnownColor.Transparent) Then
          '
          ' Add it to our list
          '
          Dim myColor As Color = Color.FromKnownColor(knownColor)
          If (384 - myColor.R - myColor.G - myColor.B) > 0 Then
            Dim hexValue As String = myColor.R.ToString("x2") & myColor.G.ToString("x2") & myColor.B.ToString("x2")
            If hexValue <> "000000" Then
              safeColors.Enqueue(hexValue)
            End If
          End If
        End If
      Next

      For Each colorName As String In colorNames
        '
        ' Cast the colorName into a KnownColor
        '
        Dim knownColor As KnownColor = DirectCast([Enum].Parse(GetType(KnownColor), colorName), KnownColor)

        '
        ' Check if the knownColor variable is a System color
        '
        If (knownColor > KnownColor.Transparent And colorName.Contains("Dark")) Then
          '
          ' Add it to our list
          '
          Dim myColor As Color = Color.FromKnownColor(knownColor)
          Dim hexValue As String = myColor.R.ToString("x2") & myColor.G.ToString("x2") & myColor.B.ToString("x2")

          safeColors.Enqueue(hexValue)
        End If
      Next

    Catch pEx As Exception
      '
      ' Process Exception
      '
      Call ProcessError(pEx, "GetSafeColors()")
    End Try

    Return safeColors

  End Function

End Module
