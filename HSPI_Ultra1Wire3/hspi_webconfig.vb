Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI
Imports System.Collections.Specialized
Imports System.Web.UI.WebControls

Public Class hspi_webconfig
  Inherits clsPageBuilder

  Public hspiref As HSPI

  Dim TimerEnabled As Boolean

  ''' <summary>
  ''' Initializes new webconfig
  ''' </summary>
  ''' <param name="pagename"></param>
  ''' <remarks></remarks>
  Public Sub New(ByVal pagename As String)
    MyBase.New(pagename)
  End Sub

#Region "Page Building"

  ''' <summary>
  ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function. 
  ''' A complete page needs to be created and returned.
  ''' </summary>
  ''' <param name="pageName"></param>
  ''' <param name="user"></param>
  ''' <param name="userRights"></param>
  ''' <param name="queryString"></param>
  ''' <param name="instance"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String, instance As String) As String

    Try

      Dim stb As New StringBuilder

      '
      ' Called from the start of your page to reset all internal data structures in the clsPageBuilder class, such as menus.
      '
      Me.reset()

      '
      ' Determine if user is authorized to access the web page
      '
      Dim LoggedInUser As String = hs.WEBLoggedInUser()
      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      '
      ' Handle any queries like mode=something
      '
      Dim parts As Collections.Specialized.NameValueCollection = Nothing
      If (queryString <> "") Then
        parts = HttpUtility.ParseQueryString(queryString)
      End If

      Dim Header As New StringBuilder
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultra1wire3/css/jquery.dataTables.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultra1wire3/css/dataTables.tableTools.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultra1wire3/css/dataTables.editor.min.css"" rel=""stylesheet"" />")
      Header.AppendLine("<link type=""text/css"" href=""/hspi_ultra1wire3/css/jquery.dataTables_themeroller.css"" rel=""stylesheet"" />")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/jquery.dataTables.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/dataTables.tableTools.min.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/dataTables.editor.min.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/amcharts.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/serial.js""></script>")

      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/hspi_ultra1wire3_devices.js""></script>")
      Header.AppendLine("<script type=""text/javascript"" src=""/hspi_ultra1wire3/js/hspi_ultra1wire3_charts.js""></script>")
      Me.AddHeader(Header.ToString)

      Dim pageTile As String = String.Format("{0} {1}", pageName, instance).TrimEnd
      stb.Append(hs.GetPageHeader(pageName, pageTile, "", "", False, False))

      '
      ' Start the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

      '
      ' A message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
      '
      stb.Append(clsPageBuilder.DivStart("divErrorMessage", "class='errormessage'"))
      stb.Append(clsPageBuilder.DivEnd)

      Me.RefreshIntervalMilliSeconds = 3000
      stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName))

      If WEBUserIsAuthorized(LoggedInUser, USER_ROLES_AUTHORIZED) = False Then
        '
        ' Current user not authorized
        '
        stb.Append(WebUserNotUnauthorized(LoggedInUser))
      Else
        '
        ' Specific page starts here
        '
        stb.Append(BuildContent)
      End If

      '
      ' End the page plug-in document division
      '
      stb.Append(clsPageBuilder.DivEnd)

      '
      ' Add the body html to the page
      '
      Me.AddBody(stb.ToString)

      '
      ' Return the full page
      '
      Return Me.BuildPage()

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "GetPagePlugin")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the HTML content
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildContent() As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table border='0' cellpadding='0' cellspacing='0' width='1000'>")
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td width='1000' align='center' style='color:#FF0000; font-size:14pt; height:30px;'><strong><div id='divMessage'>&nbsp;</div></strong></td>")
      stb.AppendLine(" </tr>")
      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>", BuildTabs())
      stb.AppendLine(" </tr>")
      stb.AppendLine("</table>")

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildContent")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Builds the jQuery Tabss
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Public Function BuildTabs() As String

    Try

      Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
      Dim tab As New clsJQuery.Tab

      tabs.postOnTabClick = True

      tab.tabTitle = "Status"
      tab.tabDIVID = "tabStatus"
      tab.tabContent = "<div id='divStatus'>" & BuildTabStatus(False) & "</div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Options"
      tab.tabDIVID = "tabOptions"
      tab.tabContent = "<div id='divOptions'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Devices"
      tab.tabDIVID = "tabDevices"
      tab.tabContent = "<div id='divDevices'></div>"
      tabs.tabs.Add(tab)

      tab = New clsJQuery.Tab
      tab.tabTitle = "Charts"
      tab.tabDIVID = "tabCharts"
      tab.tabContent = "<div id='divCharts'>" & BuildTabCharts(False) & "</div>"
      tabs.tabs.Add(tab)

      Return tabs.Build

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabs")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Status Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabStatus(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine(clsPageBuilder.FormStart("frmStatus", "frmStatus", "Post"))

      stb.AppendLine("<div>")
      stb.AppendLine("<table>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Plug-In Status </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("    <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Name:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", IFACE_NAME)
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Status:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", "OK")
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Version:</strong></td>")
      stb.AppendFormat("    <td style=""text-align: right"">{0}</td>", HSPI.Version)
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("    <legend> Database Statistics </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Inserts:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("DBInsSuccess")).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Failures:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("DBInsFailure")).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>In&nbsp;Queue:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("DBInsQueue")).ToString("N0"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>Size:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", GetDatabaseSize("DBConnectionMain"))
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td>")
      stb.AppendLine("   <fieldset>")
      stb.AppendLine("   <legend> 1-Wire Interface Adapters </legend>")
      stb.AppendLine("    <table style=""width: 100%"">")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>HA7E:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("HA7EInterfaces")).ToString())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>HA7Net:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("HA7NetInterfaces")).ToString())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>OWServer:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("OWServerInterfaces")).ToString())
      stb.AppendLine("     </tr>")
      stb.AppendLine("     <tr>")
      stb.AppendLine("      <td style=""width: 20%""><strong>TEMP08:</strong></td>")
      stb.AppendFormat("    <td align=""right"">{0}</td>", Convert.ToInt32(GetStatistics("TEMP08Interfaces")).ToString())
      stb.AppendLine("     </tr>")
      stb.AppendLine("    </table>")
      stb.AppendLine("   </fieldset>")
      stb.AppendLine("  </td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table>")
      stb.AppendLine("</div>")

      stb.AppendLine(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divStatus", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabStatus")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Options Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabOptions(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.AppendLine("<table cellspacing='0' width='100%'>")

      stb.Append(clsPageBuilder.FormStart("frmOptions", "frmOptions", "Post"))

      '
      ' HA7Net Configuration
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>HA7Net Configuration</td>")
      stb.AppendLine(" </tr>")

      '
      ' HA7Net Configuration (HA7Net Discovery)
      '
      Dim selHA7Net As New clsJQuery.jqDropList("selHA7Net", Me.PageName, False)
      selHA7Net.id = "selHA7Net"
      selHA7Net.toolTip = "If enabled, the plug-in will send an auto-discovery packet out each enabled network interface in an attempt to find connected HA7Net devices."

      Dim txtHA7Net As String = GetSetting("Interface", "HA7Net", "1")
      selHA7Net.AddItem("Enabled+Manual", "2", txtHA7Net = "2")
      selHA7Net.AddItem("Enabled", "1", txtHA7Net = "1")
      selHA7Net.AddItem("Disabled", "0", txtHA7Net = "0")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>HA7Net&nbsp;Discovery</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selHA7Net.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' OWServer Configuration
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>OWServer Configuration</td>")
      stb.AppendLine(" </tr>")

      '
      ' OWServer Configuration (OWServer Discovery)
      '
      Dim selOWServer As New clsJQuery.jqDropList("selOWServer", Me.PageName, False)
      selOWServer.id = "selOWServer"
      selOWServer.toolTip = "If enabled, the plug-in will send an auto-discovery packet out each enabled network interface in an attempt to find connected OW-Server devices."

      Dim txtOWServert As String = GetSetting("Interface", "OWServer", "1")
      selOWServer.AddItem("Enabled+Manual", "2", txtOWServert = "2")
      selOWServer.AddItem("Enabled", "1", txtOWServert = "1")
      selOWServer.AddItem("Disabled", "0", txtOWServert = "0")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>OWServer&nbsp;Discovery</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selOWServer.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Sensor Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Sensor Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' 1-Wire Options (Check Interval)
      '
      Dim selCheckInterval As New clsJQuery.jqDropList("selCheckInterval", Me.PageName, False)
      selCheckInterval.id = "selCheckInterval"
      selCheckInterval.toolTip = "Specifies how often to read the connected sensors."

      Dim txtCheckInterval As String = GetSetting("Sensor", "CheckInterval", "1")
      For index As Integer = 1 To 10 Step 1
        Dim value As String = index.ToString
        Dim desc As String = index.ToString.PadLeft(2, "0")
        selCheckInterval.AddItem(desc, value, index.ToString = txtCheckInterval)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Reading&nbsp;Interval</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Minute(s)</td>{1}", selCheckInterval.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' 1-Wire Options (Attempts)
      '
      Dim selCheckAttemps As New clsJQuery.jqDropList("selCheckAttemps", Me.PageName, False)
      selCheckAttemps.id = "selCheckAttemps"
      selCheckAttemps.toolTip = "Specifies the number of times to attempt a sensor reading before giving up."

      Dim txtAttempts As String = GetSetting("Sensor", "Attempts", "1")
      For index As Integer = 1 To 5 Step 1
        Dim value As String = index.ToString
        Dim desc As String = index.ToString.PadLeft(2, "0")
        selCheckAttemps.AddItem(desc, value, index.ToString = txtAttempts)
      Next

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Reading&nbsp;Attempts</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}&nbsp;Time(s)</td>{1}", selCheckAttemps.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Sensor Options (Sensor Logging)
      '
      Dim selDeviceLogging As New clsJQuery.jqDropList("selDeviceLogging", Me.PageName, False)
      selDeviceLogging.id = "selDeviceLogging"
      selDeviceLogging.toolTip = "If set to Yes, sensor value changes will be written to the HomeSeer log."

      Dim txtDeviceLogging As String = CBool(GetSetting("Sensor", "DeviceLogging", False)).ToString
      selDeviceLogging.AddItem("No", "False", txtDeviceLogging = "False")
      selDeviceLogging.AddItem("Yes", "True", txtDeviceLogging = "True")

      Dim ttDeviceLogging As New clsJQuery.jqToolTip("This option will enable or disable device logging for *all* sensors.  You can also set individual sensor logging by editing the HomeSeer device and setting ""Do Not log commands from this device.""")

      stb.AppendLine(" <tr>")
      stb.AppendFormat(" <td class='tablecell'>Enable Sensor Logging&nbsp;{0}</td>", ttDeviceLogging.build)
      stb.AppendFormat(" <td class='tablecell'>{0}</td>{1}", selDeviceLogging.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Sensor Options (On No Report)
      '
      Dim selAlwaysInsert As New clsJQuery.jqDropList("selAlwaysInsert", Me.PageName, False)
      selAlwaysInsert.id = "selAlwaysInsert"
      selAlwaysInsert.toolTip = "Specifies what action to take when a sensor does not respond."

      Dim txtAlwaysInsert As String = GetSetting("Sensor", "AlwaysInsert", "1")
      selAlwaysInsert.AddItem("Do not insert value into database", "0", txtAlwaysInsert = "0")
      selAlwaysInsert.AddItem("Insert last value into database", "1", txtAlwaysInsert = "1")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>On No Report</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selAlwaysInsert.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>HomeSeer Device Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options (Unit Type)
      '
      Dim selUnitType As New clsJQuery.jqDropList("selUnitType", Me.PageName, False)
      selUnitType.id = "selUnitType"
      selUnitType.toolTip = "The format used to display temperatures, rainfall and barometric pressure.  The default format is U.S customary units."

      Dim strUnitType As String = GetSetting("Options", "UnitType", "US")
      selUnitType.AddItem("U.S. customary units (miles, °F, etc...)", "US", strUnitType = "US")
      selUnitType.AddItem("Metric system units (kms, °C, etc...)", "Metric", strUnitType = "Metric")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Unit&nbsp;Type</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selUnitType.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options (Display Degree Units)
      '
      Dim selTempDegreeUnit As New clsJQuery.jqDropList("selTempDegreeUnit", Me.PageName, False)
      selTempDegreeUnit.id = "selTempDegreeUnit"
      selTempDegreeUnit.toolTip = "If set to Yes, the degree unit (e.g.  F or C) will be displayed on the HomeSeer status web page."

      Dim txtTempDegreeUnit As String = CBool(GetSetting("Options", "TempDegreeUnit", "True")).ToString
      selTempDegreeUnit.AddItem("No", "False", txtTempDegreeUnit = "False")
      selTempDegreeUnit.AddItem("Yes", "True", txtTempDegreeUnit = "True")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Display&nbsp;Degree&nbsp;Units</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selTempDegreeUnit.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' HomeSeer Device Options (Display Degree Icon)
      '
      Dim selTempDegreeIcon As New clsJQuery.jqDropList("selTempDegreeIcon", Me.PageName, False)
      selTempDegreeIcon.id = "selTempDegreeIcon"
      selTempDegreeIcon.toolTip = "If set to Yes, a degree icon will be displayed on the HomeSeer status web page."

      Dim txtTempDegreeIcon As String = CBool(GetSetting("Options", "TempDegreeIcon", "True")).ToString
      selTempDegreeIcon.AddItem("No", "False", txtTempDegreeIcon = "False")
      selTempDegreeIcon.AddItem("Yes", "True", txtTempDegreeIcon = "True")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell'>Display&nbsp;Degree&nbsp;Icon</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selTempDegreeIcon.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Web Page Access (Authorized User Roles)
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Web Page Access</td>")
      stb.AppendLine(" </tr>")

      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tablecell' style=""width: 20%"">Authorized User Roles</td>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", BuildWebPageAccessCheckBoxes, vbCrLf)
      stb.AppendLine(" </tr>")

      '
      ' Application Options
      '
      stb.AppendLine(" <tr>")
      stb.AppendLine("  <td class='tableheader' colspan='2'>Application Options</td>")
      stb.AppendLine(" </tr>")

      '
      ' Application Options (Logging Level)
      '
      Dim selLogLevel As New clsJQuery.jqDropList("selLogLevel", Me.PageName, False)
      selLogLevel.id = "selLogLevel"
      selLogLevel.toolTip = "Specifies the plug-in logging level."

      Dim itemValues As Array = System.Enum.GetValues(GetType(LogLevel))
      Dim itemNames As Array = System.Enum.GetNames(GetType(LogLevel))

      For i As Integer = 0 To itemNames.Length - 1
        Dim itemSelected As Boolean = IIf(gLogLevel = itemValues(i), True, False)
        selLogLevel.AddItem(itemNames(i), itemValues(i), itemSelected)
      Next
      selLogLevel.autoPostBack = True

      stb.AppendLine(" <tr>")
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", "Logging&nbsp;Level", vbCrLf)
      stb.AppendFormat("  <td class='tablecell'>{0}</td>{1}", selLogLevel.Build, vbCrLf)
      stb.AppendLine(" </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divOptions", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Devices Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabDevices(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmDevices", "frmDevices", "Post"))

      stb.AppendLine("<table width='100%' class='display cell-border' id='table_devices' cellspacing='0'>")

      '
      ' HA7Net Configuration
      '
      stb.AppendLine(" <thead>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <th>Device Serial</th>")
      stb.AppendLine("   <th>Device Name</th>")
      stb.AppendLine("   <th>Device Type</th>")
      stb.AppendLine("   <th>Connection Type</th>")
      stb.AppendLine("   <th>Connection Address</th>")
      stb.AppendLine("   <th>Action</th>")
      stb.AppendLine("  </tr>")
      stb.AppendLine(" </thead>")

      stb.AppendLine(" <tbody>")
      Dim MyDataTable As DataTable = hspi_plugin.Get1WireDevices()
      For Each row As DataRow In MyDataTable.Rows
        stb.AppendFormat("  <tr id='{0}'>{1}", row("device_id"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_serial"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_name"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_type"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_conn"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", row("device_addr"), vbCrLf)
        stb.AppendFormat("   <td>{0}</td>{1}", "", vbCrLf)
        stb.AppendLine("  </tr>")
      Next
      stb.AppendLine(" </tbody>")

      stb.AppendLine("</table")

      Dim strHint As String = "You will need to manually add 1-Wire devices that do not support discovery (e.g.  TEMP08, HA7E)."
      Dim strInfo As String = "You can temporarily disable a 1-Wire device by setting the Connection Type to Disabled."

      stb.AppendLine(" <div>&nbsp;</div>")
      stb.AppendLine(" <p>")
      stb.AppendFormat("<img alt='Hint' src='/images/hspi_ultra1wire3/ico_hint.gif' width='16' height='16' border='0' />&nbsp;{0}<br/>", strHint)
      stb.AppendFormat("<img alt='Info' src='/images/hspi_ultra1wire3/ico_info.gif' width='16' height='16' border='0' />&nbsp;{0}", strInfo)
      stb.AppendLine(" </p>")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divDevices", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabOptions")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Charts Tab
  ''' </summary>
  ''' <param name="Rebuilding"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Function BuildTabCharts(Optional ByVal Rebuilding As Boolean = False) As String

    Try

      Dim stb As New StringBuilder

      stb.Append(clsPageBuilder.FormStart("frmCharts", "frmCharts", "Post"))

      stb.AppendLine("<table width='100%'>")
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <td>")

      '
      ' Chart Type
      '
      Dim selChartType As New clsJQuery.jqDropList("selChartType", Me.PageName, True)
      selChartType.id = "selChartType"
      selChartType.toolTip = "Select the Sensor Type"

      Dim strChartType As String = GetSetting("WebPage", "FilterCompare", "is")
      selChartType.AddItem("Temperature", "Temperature", "Temperature" = strChartType)
      selChartType.AddItem("Humidity", "Humidity", "Humidity" = strChartType)
      selChartType.AddItem("Counter", "Counter", "Counter" = strChartType)
      selChartType.AddItem("Pressure", "Pressure", "Pressure" = strChartType)
      stb.AppendLine(selChartType.Build())

      '
      ' Chart Type
      '
      Dim selAMChartType As New clsJQuery.jqDropList("selAMChartType", Me.PageName, True)
      selAMChartType.id = "selAMChartType"
      selAMChartType.toolTip = "Select the Chart Type"

      Dim strAMChartType As String = GetSetting("WebPage", "FilterCompare", "is")
      selAMChartType.AddItem("Smoothed Line Chart", "smoothedLine", "smoothedLine" = strAMChartType)
      selAMChartType.AddItem("Line Chart", "line", "line" = strAMChartType)
      selAMChartType.AddItem("Step Chart", "step", "step" = strAMChartType)
      selAMChartType.AddItem("Column Chart", "column", "column" = strAMChartType)
      stb.AppendLine(selAMChartType.Build())

      '
      ' Chart Interval
      '
      Dim selChartInterval As New clsJQuery.jqDropList("selChartInterval", Me.PageName, True)
      selChartInterval.id = "selChartInterval"
      selChartInterval.toolTip = "Select the Chart Type"

      Dim strChartInterval As String = GetSetting("WebPage", "FilterCompare", "is")
      selChartInterval.AddItem("For the last 1 day", "1 day", "1 day" = strChartInterval)
      selChartInterval.AddItem("For the last 1 week", "1 week", "1 week" = strChartInterval)
      selChartInterval.AddItem("For the last 1 month", "1 month", "1 month" = strChartInterval)
      selChartInterval.AddItem("For the last 3 months", "3 months", "3 months" = strChartInterval)
      selChartInterval.AddItem("For the last 6 months", "6 months", "6 months" = strChartInterval)
      selChartInterval.AddItem("For the last 1 year", "1 year", "1 year" = strChartInterval)
      stb.AppendLine(selChartInterval.Build())
      stb.AppendLine("&nbsp;Ending&nbsp;")

      Dim tbEndDate As New clsJQuery.jqTextBox("txtEndDate", "text", Date.Now.ToLongDateString, PageName, 30, True)
      tbEndDate.id = "txtEndDate"
      tbEndDate.toolTip = ""
      tbEndDate.editable = True
      stb.AppendLine(tbEndDate.Build)

      Dim btnChart As New clsJQuery.jqButton("btnChart", "Filter", Me.PageName, True)
      stb.AppendLine(btnChart.Build())

      stb.AppendLine("   </td>")
      stb.AppendLine("  </tr>")

      '
      ' HA7Net Configuration
      '
      stb.AppendLine("  <tr>")
      stb.AppendLine("   <td><div id='chartdiv' style='width: 960px; height: 720px;'></div></td>")
      stb.AppendLine("  </tr>")

      stb.AppendLine("</table")

      stb.Append(clsPageBuilder.FormEnd())

      If Rebuilding Then Me.divToUpdate.Add("divCharts", stb.ToString)

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildTabCharts")
      Return "error - " & Err.Description
    End Try

  End Function

  ''' <summary>
  ''' Build the Web Page Access Checkbox List
  ''' </summary>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function BuildWebPageAccessCheckBoxes()

    Try

      Dim stb As New StringBuilder

      Dim USER_ROLES_AUTHORIZED As Integer = WEBUserRolesAuthorized()

      Dim cb1 As New clsJQuery.jqCheckBox("chkWebPageAccess_Guest", "Guest", Me.PageName, True, True)
      Dim cb2 As New clsJQuery.jqCheckBox("chkWebPageAccess_Admin", "Admin", Me.PageName, True, True)
      Dim cb3 As New clsJQuery.jqCheckBox("chkWebPageAccess_Normal", "Normal", Me.PageName, True, True)
      Dim cb4 As New clsJQuery.jqCheckBox("chkWebPageAccess_Local", "Local", Me.PageName, True, True)

      cb1.id = "WebPageAccess_Guest"
      cb1.checked = CBool(USER_ROLES_AUTHORIZED And USER_GUEST)

      cb2.id = "WebPageAccess_Admin"
      cb2.checked = CBool(USER_ROLES_AUTHORIZED And USER_ADMIN)
      cb2.enabled = False

      cb3.id = "WebPageAccess_Normal"
      cb3.checked = CBool(USER_ROLES_AUTHORIZED And USER_NORMAL)

      cb4.id = "WebPageAccess_Local"
      cb4.checked = CBool(USER_ROLES_AUTHORIZED And USER_LOCAL)

      stb.Append(clsPageBuilder.FormStart("frmWebPageAccess", "frmWebPageAccess", "Post"))

      stb.Append(cb1.Build())
      stb.Append(cb2.Build())
      stb.Append(cb3.Build())
      stb.Append(cb4.Build())

      stb.Append(clsPageBuilder.FormEnd())

      Return stb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "BuildWebPageAccessCheckBoxes")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

#Region "Page Processing"

  ''' <summary>
  ''' Post a message to this web page
  ''' </summary>
  ''' <param name="sMessage"></param>
  ''' <remarks></remarks>
  Sub PostMessage(ByVal sMessage As String)

    Try

      Me.divToUpdate.Add("divMessage", sMessage)

      Me.pageCommands.Add("starttimer", "")

      TimerEnabled = True

    Catch pEx As Exception

    End Try

  End Sub

  ''' <summary>
  ''' When a user clicks on any controls on one of your web pages, this function is then called with the post data. You can then parse the data and process as needed.
  ''' </summary>
  ''' <param name="page">The name of the page as registered with hs.RegisterLink or hs.RegisterConfigLink</param>
  ''' <param name="data">The post data</param>
  ''' <param name="user">The name of logged in user</param>
  ''' <param name="userRights">The rights of logged in user</param>
  ''' <returns>Any serialized data that needs to be passed back to the web page, generated by the clsPageBuilder class</returns>
  ''' <remarks></remarks>
  Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

    Try

      WriteMessage("Entered postBackProc() function.", MessageType.Debug)

      Dim postData As NameValueCollection = HttpUtility.ParseQueryString(data)

      '
      ' Write debug to console
      '
      If gLogLevel >= MessageType.Debug Then
        For Each keyName As String In postData.AllKeys
          Console.WriteLine(String.Format("{0}={1}", keyName, postData(keyName)))
        Next
      End If

      Select Case postData("action")
        Case "edit"
          Dim device_id As String = postData("id")
          Dim device_serial As String = postData("data[device_serial]").Trim
          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_type As String = postData("data[device_type]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          ''
          '' Update the 1-Wire Device
          ''
          Dim bSuccess As Boolean = Update1WireDevice(device_id, device_serial, device_name, device_type, device_conn, device_addr)
          If bSuccess = False Then
            Dim sb As New StringBuilder

            sb.AppendLine("{")
            sb.AppendFormat(" ""error"": {0}{1}", "Unable to update 1-Wire device due to an error.", vbCrLf)
            sb.AppendLine("}")

            Return sb.ToString
          Else
            BuildTabDevices(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return "{ }"
          End If

        Case "create"
          Dim device_serial As String = postData("data[device_serial]").Trim
          Dim device_name As String = postData("data[device_name]").Trim
          Dim device_type As String = postData("data[device_type]").Trim
          Dim device_conn As String = postData("data[device_conn]").Trim
          Dim device_addr As String = postData("data[device_addr]").Trim

          '
          ' Update the 1-Wire Device
          '
          Dim bSuccess As Boolean = Insert1WireDevice(device_serial, device_name, device_type, "", device_conn, device_addr)
          If bSuccess = False Then
            Dim sb As New StringBuilder

            sb.AppendLine("{")
            sb.AppendFormat(" ""error"": {0}{1}", "Unable to insert new 1-Wire device due to an error.", vbCrLf)
            sb.AppendLine("}")

            Return sb.ToString
          Else
            BuildTabDevices(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return "{ }"
          End If

        Case "remove"
          Dim device_id As String = Val(postData("id[]"))

          Dim bSuccess As Boolean = Delete1WireDevice(device_id)
          If bSuccess = False Then
            Dim sb As New StringBuilder

            sb.AppendLine("{")
            sb.AppendFormat(" ""error"": {0}{1}", "Unable to delete 1-Wire device due to an error.", vbCrLf)
            sb.AppendLine("}")

            Return sb.ToString
          Else
            BuildTabDevices(True)
            Me.pageCommands.Add("executefunction", "reDraw()")

            Return "{ }"
          End If

      End Select

      '
      ' Process the post data
      '
      Select Case postData("id")
        Case "tabStatus"
          BuildTabStatus(True)

        Case "tabOptions"
          BuildTabOptions(True)

        Case "tabDevices"
          BuildTabDevices(True)
          Me.pageCommands.Add("executefunction", "reDraw()")

        Case "tabCharts"
          Me.pageCommands.Add("executefunction", "initChart()")

        Case "selHA7Net"
          Dim value As String = postData(postData("id"))
          SaveSetting("Interface", "HA7Net", value)

          PostMessage("The HA7Net Discovery option has been updated.")
        Case "selOWServer"
          Dim value As String = postData(postData("id"))
          SaveSetting("Interface", "OWServer", value)

          PostMessage("The OWServer Discovery option has been updated.")
        Case "selCheckInterval"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sensor", "CheckInterval", value)

          PostMessage("The Sensor Check Interval option has been updated.")
        Case "selCheckAttemps"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sensor", "Attempts", value)

          PostMessage("The Sensor Check Attempts option has been updated.")
        Case "selAlwaysInsert"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sensor", "AlwaysInsert", value)

          PostMessage("The Sensor No Report option has been updated.")
        Case "selDeviceLogging"
          Dim value As String = postData(postData("id"))
          SaveSetting("Sensor", "DeviceLogging", value)

          UpdateMiscDeviceSettings(CBool(value))

          PostMessage("The Sensor Device Logging option has been updated.")
        Case "selUnitType"
          Dim strValue As String = postData(postData("id"))
          SaveSetting("Options", "UnitType", strValue)

          PostMessage("The Unit Type has been updated.")
        Case "selTempDegreeUnit"
          Dim strValue As String = CBool(postData(postData("id"))).ToString
          SaveSetting("Options", "TempDegreeUnit", strValue)

          PostMessage("The Temperature Degree Unit option has been updated.")
        Case "selTempDegreeIcon"
          Dim strValue As String = CBool(postData(postData("id"))).ToString
          SaveSetting("Options", "TempDegreeIcon", strValue)

          PostMessage("The Temperature Degree Icon option has been updated.")
        Case "selLogLevel"
          gLogLevel = Int32.Parse(postData("selLogLevel"))
          hs.SaveINISetting("Options", "LogLevel", gLogLevel.ToString, gINIFile)

          PostMessage("The application logging level has been updated.")
        Case "btnChart"
          Me.pageCommands.Add("executefunction", "initChart()")

        Case "sensorvalues"
          Dim device_id As Integer = Int32.Parse(postData("device_id"))
          Dim chart_type As String = postData("chart_type")
          Dim amchart_type As String = postData("amchart_type")
          Dim end_date As String = postData("end_date")
          Dim internval As String = postData("interval")

          Dim strJSON As String = GetSensorChartJSON(device_id, chart_type, amchart_type, end_date, internval)
          Return strJSON

        Case "WebPageAccess_Guest"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Guest") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_GUEST
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_GUEST
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Normal"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Normal") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_NORMAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_NORMAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "WebPageAccess_Local"

          Dim AUTH_ROLES As Integer = WEBUserRolesAuthorized()
          If postData("chkWebPageAccess_Local") = "checked" Then
            AUTH_ROLES = AUTH_ROLES Or USER_LOCAL
          Else
            AUTH_ROLES = AUTH_ROLES Xor USER_LOCAL
          End If
          hs.SaveINISetting("WEBUsers", "AuthorizedRoles", AUTH_ROLES.ToString, gINIFile)

        Case "timer" ' This stops the timer and clears the message
          If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
            TimerEnabled = False
          Else
            Me.pageCommands.Add("stoptimer", "")
            Me.divToUpdate.Add("divMessage", "&nbsp;")
          End If

      End Select

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "postBackProc")
    End Try

    Return MyBase.postBackProc(page, data, user, userRights)

  End Function

#End Region

#Region "HSPI - Web Authorization"

  ''' <summary>
  ''' Returns the HTML Not Authorized web page
  ''' </summary>
  ''' <param name="LoggedInUser"></param>
  ''' <returns></returns>
  ''' <remarks></remarks>
  Private Function WebUserNotUnauthorized(LoggedInUser As String) As String

    Try

      Dim sb As New StringBuilder

      sb.AppendLine("<table border='0' cellpadding='2' cellspacing='2' width='575px'>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td nowrap>")
      sb.AppendLine("     <h4>The Web Page You Were Trying To Access Is Restricted To Authorized Users ONLY</h4>")
      sb.AppendLine("   </td>")
      sb.AppendLine("  </tr>")
      sb.AppendLine("  <tr>")
      sb.AppendLine("   <td>")
      sb.AppendLine("     <p>This page is displayed if the credentials passed to the web server do not match the ")
      sb.AppendLine("      credentials required to access this web page.</p>")
      sb.AppendFormat("     <p>If you know the <b>{0}</b> user should have access,", LoggedInUser)
      sb.AppendFormat("      then ask your <b>HomeSeer Administrator</b> to check the <b>{0}</b> plug-in options", IFACE_NAME)
      sb.AppendFormat("      page to make sure the roles assigned to the <b>{0}</b> user allow access to this", LoggedInUser)
      sb.AppendLine("        web page.</p>")
      sb.AppendLine("  </td>")
      sb.AppendLine(" </tr>")
      sb.AppendLine(" </table>")

      Return sb.ToString

    Catch pEx As Exception
      '
      ' Process program exception
      '
      ProcessError(pEx, "WebUserNotUnauthorized")
      Return "error - " & Err.Description
    End Try

  End Function

#End Region

End Class
