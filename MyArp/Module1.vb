' Copyright Rob Latour, 2024
' https://rlatour.com/myarp

Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Configuration
Module Module1

    Private gCommandLine As String
    Private Structure Switches

        Dim Help As Boolean
        Dim ADD As Boolean
        Dim Delete As Boolean
        Dim Clear As Boolean
        Dim DatabaseBackup As Boolean
        Dim DatabaseDelete As Boolean
        Dim DatabaseEdit As Boolean
        Dim DatabaseRestore As Boolean
        Dim Hyper_V As Boolean
        Dim Refresh As Boolean
        Dim Pause As Boolean
        Dim NoActive As Boolean
        Dim NoInactive As Boolean
        Dim NoPing As Boolean
        Dim NoDynamic As Boolean
        Dim NoStatic As Boolean
        Dim Quite As Boolean

        Dim NotQuite As Boolean

    End Structure

    Private gCommandLineSwitch As Switches

    Private gProgramName As String

    Friend Structure DeviceTableRowStructure

        Friend Active As String
        Friend IPAddress As String
        Friend SortableIPAddress As Long
        Friend MAC As String
        Friend DeviceName As String
        Friend Description As String
        Friend LastSeen As String
        Friend StaticOrDynamic As String

    End Structure

    Private ExistingDeviceTable(5000) As DeviceTableRowStructure
    Private ExistingDeviceTableIndex As Integer = 0

    Private NewDeviceTable(1000) As DeviceTableRowStructure
    Private NewDeviceTableIndex As Integer = 0

    Private gDatabasePathAndFilename As String
    Private gBackupDatabasePathAndFilename As String

    Private gSettingsPathAndFilename As String

    Private gUpdateInventory As Boolean = True
    Private gWriteReport As Boolean = True

    Private gOriginalForegroundColour As ConsoleColor

    Private Dashes As String = StrDup(80, "-")
    Const cDynamic As String = "Dynamic"
    Const cStatic As String = "Static"
    Const cUnknown As String = "Unknown"

    Const cAPIAPrefix As String = "169.254"
    Const cUsingAPIA As String = " (using Automatic Private IP Addressing)"
    'ref: https://askleo.com/why_cant_i_connect_with_a_169254xx_ip_address/

    Const cHyperVPrefix As String = "234."

    Private CommandLineOverrideForTesting As String = String.Empty

    Sub Main()

#If DEBUG Then

        ' used for testing:

        'CommandLineOverrideForTesting = "/DBR /P /NP"
        'CommandLineOverrideForTesting = "/P /DBE /NP "
        'CommandLineOverrideForTesting = "/DEL 01:02:03:04:05:06 "
        'CommandLineOverrideForTesting = "/ADD 01:02:03:04:05:06 This is another test "
        'CommandLineOverrideForTesting = "/ADD 7C:DD:90:A0:AE:E7 Raspberry PI Wireless"
        'CommandLineOverrideForTesting = "/DEL 02:0F:B5:43:0D:F6"
        'CommandLineOverrideForTesting = "/? /P"
        'CommandLineOverrideForTesting = "/NP /NRD /P"
        'CommandLineOverrideForTesting = "/NP /NRS /P"
        'CommandLineOverrideForTesting = "/NRA /NRI /P"   ' nothing to report
        'CommandLineOverrideForTesting = "/NRD /NRS /P"   ' nothing to report
        'CommandLineOverrideForTesting = "/NRD /P"
        'CommandLineOverrideForTesting = " /HYPERV /P"
        'CommandLineOverrideForTesting = "/Q"
        'CommandLineOverrideForTesting = "/R /HYPERV /P"
        CommandLineOverrideForTesting = "/P"

#End If

        Initialize()

        If gUpdateInventory Then
            PingDevicesLoadedInMemory()
            RefreshFromThisComputer()
            RefreshFromNetwork()
            UpdateInventoryForHyperV()
            UpdateInventoryForExtender()
            UpdateInventoryForAutomaticPrivateIPAddressing()
        Else
            NewDeviceTable = Nothing
        End If

        WriteDatabaseFromMemory()

        If gWriteReport Then

            WriteHeaderLine(True)
            WriteActiveRecords()
            WriteHeaderLine(False)
            WriteInactiveRecords()

            Console.WriteLine("")

        End If

        Wrapup()

    End Sub

#Region "Setup"
    Private Sub Initialize()

        Try

            SetSettingsVersion()

            InitalizeWindow()

            SetDatabaseName()

            SetCommandLineSwitches()

            If gCommandLineSwitch.Clear Then Console.Clear()

            If gCommandLineSwitch.Quite Then

                gWriteReport = False

            Else

                If (gCommandLineSwitch.NoActive AndAlso gCommandLineSwitch.NoInactive) Then
                    WriteAColouredConsoleLine("Warning: you have selected to not report either active or inactive devices - so nothing will be reported.", ConsoleColor.Yellow)
                    gUpdateInventory = False
                    gWriteReport = False
                End If

                If gCommandLineSwitch.NoDynamic AndAlso gCommandLineSwitch.NoStatic Then
                    WriteAColouredConsoleLine("Warning: you have selected not to report either dynamic or static addresses - so nothing will be reported.", ConsoleColor.Yellow)
                    gUpdateInventory = False
                    gWriteReport = False
                End If

                If gCommandLineSwitch.Help Then
                    ShowHelp()
                    gUpdateInventory = False
                    gWriteReport = False
                End If

            End If

            If gCommandLineSwitch.DatabaseBackup Then

                If File.Exists(gDatabasePathAndFilename) Then

                    File.Copy(gDatabasePathAndFilename, gBackupDatabasePathAndFilename, True)

                    WriteAColouredConsoleLine(gDatabasePathAndFilename & " was backed up to " & gBackupDatabasePathAndFilename, ConsoleColor.Green)

                Else

                    WriteAColouredConsoleLine(gDatabasePathAndFilename & " was not found; no backup was performed.", ConsoleColor.Red)

                End If

            End If

            If gCommandLineSwitch.DatabaseDelete Then

                If File.Exists(gDatabasePathAndFilename) Then

                    If gCommandLineSwitch.Quite OrElse Console_PromptForYesorNo("Do you want to delete the database?") Then

                        File.Delete(gDatabasePathAndFilename)
                        WriteAColouredConsoleLine(gDatabasePathAndFilename & " was deleted.", ConsoleColor.Green)

                    Else

                        WriteAColouredConsoleLine("ok, nothing was deleted.", ConsoleColor.Gray)

                    End If

                Else

                    WriteAColouredConsoleLine(gDatabasePathAndFilename & " was not found; so it wasn't deleted.", ConsoleColor.Yellow)

                End If

            End If

            If gCommandLineSwitch.DatabaseRestore Then

                If File.Exists(gBackupDatabasePathAndFilename) Then

                    If File.Exists(gDatabasePathAndFilename) Then

                        If gCommandLineSwitch.Quite OrElse Console_PromptForYesorNo("Do you want to restore the database?") Then
                            File.Copy(gBackupDatabasePathAndFilename, gDatabasePathAndFilename, True)
                            WriteAColouredConsoleLine("ok, the datebase was restored.", ConsoleColor.Green)

                        Else

                            WriteAColouredConsoleLine("ok, nothing was deleted.", ConsoleColor.Gray)

                        End If

                    Else

                        File.Copy(gBackupDatabasePathAndFilename, gDatabasePathAndFilename)
                        WriteAColouredConsoleLine("ok, the datebase was restored.", ConsoleColor.Green)

                    End If

                Else

                    WriteAColouredConsoleLine(gBackupDatabasePathAndFilename & " was not found; so the database could not be restored.", ConsoleColor.Yellow)

                End If

            End If

            If gCommandLineSwitch.DatabaseEdit Then

                If File.Exists(gDatabasePathAndFilename) Then

                    WriteAColouredConsoleLine(gDatabasePathAndFilename & " has been opened for editing in Notepad;", ConsoleColor.White,, False)
                    WriteAColouredConsoleLine("processing will continue once Notepad has been exited.", ConsoleColor.White, False)

                    Dim ps As New System.Diagnostics.ProcessStartInfo("notepad", gDatabasePathAndFilename)

                    ps.RedirectStandardOutput = True
                    ps.UseShellExecute = False
                    ps.WindowStyle = ProcessWindowStyle.Normal
                    ps.CreateNoWindow = False

                    Using proc As New System.Diagnostics.Process()

                        proc.StartInfo = ps
                        proc.Start()

                        While Not proc.HasExited
                            System.Threading.Thread.Sleep(100)
                        End While

                    End Using

                Else

                    WriteAColouredConsoleLine(gDatabasePathAndFilename & " does not exist; it can not be edited.", ConsoleColor.Yellow)

                End If

            End If

            LoadDatabaseInMemory()

            If gCommandLineSwitch.Delete Then

                gUpdateInventory = False
                gWriteReport = False

                Dim FormatError As Boolean = False

                Try

                    Dim WorkingLine As String = gCommandLine.Trim

                    If WorkingLine.ToUpper.StartsWith("/DEL") Then
                        WorkingLine = WorkingLine.Remove(0, 4).TrimStart
                    Else
                        FormatError = True
                    End If

                    If (WorkingLine.Length <> 17) OrElse (WorkingLine.Count(Function(c) c = ":"c) <> 5) Then
                        FormatError = True
                    End If

                    If FormatError Then

                        WriteAColouredConsoleLine("Delete request incorrectly formatted, should be: /DEL [Physical Address]", ConsoleColor.Red,, False)
                        WriteAColouredConsoleLine("For example: /DEL 01:02:03:04:05:06", ConsoleColor.Red, False)

                        Exit Try

                    Else

                        If IsPhsyicalAddressInDatabase(ExistingDeviceTable, WorkingLine.ToUpper) Then

                            For x = 0 To ExistingDeviceTable.Count - 1

                                If WorkingLine.ToUpper = ExistingDeviceTable(x).MAC Then

                                    Dim aList As List(Of DeviceTableRowStructure) = ExistingDeviceTable.ToList
                                    aList.RemoveAt(x)
                                    ExistingDeviceTable = aList.ToArray
                                    ExistingDeviceTableIndex = ExistingDeviceTable.Count - 1

                                    WriteAColouredConsoleLine("Delete succeeded.", ConsoleColor.Green)

                                    Exit For

                                End If

                            Next

                        Else

                            WriteAColouredConsoleLine("Physical Address " & WorkingLine & " was not found.", ConsoleColor.Red)

                            FormatError = True

                            Exit Try

                        End If

                    End If

                Catch ex As Exception

                End Try

                If FormatError Then
                    gUpdateInventory = False
                    gWriteReport = False
                End If

                Exit Try

            End If

            If gCommandLineSwitch.ADD Then

                gUpdateInventory = False
                gWriteReport = False

                Dim FormatError As Boolean = False

                Try

                    Dim WorkingLine As String = gCommandLine.Trim

                    If WorkingLine.ToUpper.StartsWith("/ADD") Then
                        WorkingLine = WorkingLine.Remove(0, 4).TrimStart
                    Else
                        FormatError = True
                    End If

                    Dim MAC As String = String.Empty

                    If (WorkingLine.Length < 17) Then
                        FormatError = True
                    Else
                        MAC = WorkingLine.Substring(0, 17).ToUpper
                        If (MAC.Count(Function(c) c = ":"c) <> 5) Then
                            FormatError = True
                        End If
                    End If

                    If FormatError Then

                        WriteAColouredConsoleLine("Add request incorrectly formatted, should be: /ADD [Physical Address] (Description)", ConsoleColor.Red,, False)
                        WriteAColouredConsoleLine("For example: /ADD 01:02:03:04:05:06 or /ADD 01:02:03:04:05:06 Rob's Laptop", ConsoleColor.Red, False)
                        Exit Try

                    End If

                    Dim Description As String = String.Empty

                    Description = WorkingLine.Remove(0, 17).Trim

                    ' update description if entry already exists

                    If IsPhsyicalAddressInDatabase(ExistingDeviceTable, MAC) Then

                        If Description.Length > 0 Then

                            For x = 0 To ExistingDeviceTable.Count - 1

                                If MAC = ExistingDeviceTable(x).MAC Then
                                    ExistingDeviceTable(x).Description = Description
                                    If Description.Length > 0 Then
                                        WriteAColouredConsoleLine("Description update succeeded.", ConsoleColor.Green)
                                    Else
                                        WriteAColouredConsoleLine("Add succeeded.", ConsoleColor.Green)
                                    End If
                                    Exit For

                                End If

                            Next

                        Else

                            WriteAColouredConsoleLine("Add request not required, the database already contains an entry for " & MAC & ".", ConsoleColor.Yellow)
                            FormatError = True
                            Exit Try

                        End If

                    Else

                        ExistingDeviceTableIndex += 1
                        ReDim Preserve ExistingDeviceTable(ExistingDeviceTableIndex)
                        With ExistingDeviceTable(ExistingDeviceTableIndex)

                            .Active = "Inactive"
                            .IPAddress = "0.0.0.0"
                            .SortableIPAddress = 0
                            .MAC = MAC
                            .DeviceName = cUnknown
                            .StaticOrDynamic = cUnknown
                            .Description = Description
                            .LastSeen = "never"

                        End With

                        WriteAColouredConsoleLine("Add succeeded.", ConsoleColor.Green)

                    End If

                Catch ex As Exception

                End Try

            End If

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

    End Sub
    Private Sub SetSettingsVersion()

        Dim a As System.Reflection.Assembly = System.Reflection.Assembly.GetExecutingAssembly()
        Dim appVersion As Version = a.GetName().Version
        Dim appVersionString As String = appVersion.ToString

        If My.Settings.ApplicationVersion <> appVersion.ToString Then
            My.Settings.Upgrade()
            My.Settings.ApplicationVersion = appVersionString
        End If

        Dim UserConfig As Configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal)
        gSettingsPathAndFilename = UserConfig.FilePath

    End Sub
    Private Sub InitalizeWindow()

        gOriginalForegroundColour = Console.ForegroundColor

        Dim myHeight As Integer = Console.LargestWindowHeight / 2

        If My.Settings.Height > 0 Then
            myHeight = My.Settings.Height
        End If

        Dim myWidth As Integer = Console.LargestWindowWidth / 2

        If My.Settings.Width > 0 Then
            myWidth = My.Settings.Width
        End If

        Console.SetWindowSize(myWidth, myHeight)

        SetConsoleCtrlHandler(New HandlerRoutine(AddressOf WindowClosingHandler), True) ' used to catch window closing

    End Sub

    Private Sub SetCommandLineSwitches()

        gCommandLine = Microsoft.VisualBasic.Command.ToString.Trim

        If CommandLineOverrideForTesting <> String.Empty Then gCommandLine = CommandLineOverrideForTesting

        Dim UCommands As String = gCommandLine.Replace("\", "/").ToUpper

        With gCommandLineSwitch

            .Help = UCommands.Contains("/?")

            .ADD = UCommands.Contains("/ADD")
            .Delete = UCommands.Contains("/DEL")

            .Clear = UCommands.Contains("/C")

            .DatabaseBackup = UCommands.Contains("/DBB")
            .DatabaseDelete = UCommands.Contains("/DBD")
            .DatabaseEdit = UCommands.Contains("/DBE")
            .DatabaseRestore = UCommands.Contains("/DBR")

            .Hyper_V = UCommands.Contains("/HYPERV") OrElse UCommands.Contains("/HYPER-V")

            .NoActive = UCommands.Contains("/NRA")
            .NoInactive = UCommands.Contains("/NRI")
            .NoDynamic = UCommands.Contains("/NRD")
            .NoStatic = UCommands.Contains("/NRS")

            .NoPing = UCommands.Contains("/NP")

            .Pause = UCommands.Contains("/P")
            .Quite = UCommands.Contains("/Q")
            .Refresh = UCommands.Contains("/R")

            .NotQuite = Not .Quite

        End With

    End Sub

#End Region

#Region "Processing"

    Private Sub ShowHelp()

        Dim ProgramName As String = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name

        Dim SomeSpaces = StrDup(20, " ")

        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("")
        Console.WriteLine("Application:  MyArp version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor)
        Console.WriteLine("Copyright:    2024 Rob Latour")
        Console.WriteLine("License:      MIT - https://github.com/roblatour/myarp/blob/main/LICENSE")
        Console.WriteLine("Website:      https://github.com/roblatour/myarp")

        Console.WriteLine("")
        Console.WriteLine("Loading from: " & Environment.CurrentDirectory)
        Console.WriteLine("Database:     " & gDatabasePathAndFilename)
        Console.WriteLine("Settings:     " & gSettingsPathAndFilename)
        Console.WriteLine("")
        Console.Write("Usage:        ")
        Console.ForegroundColor = ConsoleColor.White
        Console.WriteLine(ProgramName & " /? /ADD [Physical Address] (Description) /DEL [Physical Address] /C /DBB /DBD /DBE /DBR /NP /NRA /NRI /NRD /NRS /P /Q /R ")
        Console.ForegroundColor = ConsoleColor.Gray
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/?                                    = show (this) help and exit")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/ADD [Physical address] (Description) = Add or update a database entry and description")
        Console.WriteLine(SomeSpaces & "                                        only one /ADD statement is allowed at at time")
        Console.WriteLine(SomeSpaces & "                                        an /ADD statement must be the only statement on a line")
        Console.WriteLine(SomeSpaces & "                                        example /ADD statements look like this:")
        Console.WriteLine(SomeSpaces & "                                                /ADD 7C:DD:90:00:00:01")
        Console.WriteLine(SomeSpaces & "                                                /ADD 7C:DD:90:00:00:02 Raspberry Pi Wireless")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/DEL [Physical address]               = Delete a database entry")
        Console.WriteLine(SomeSpaces & "                                        only one /DEL statement is allowed at at time")
        Console.WriteLine(SomeSpaces & "                                        a /DEL statement must be the only statement on a line")
        Console.WriteLine(SomeSpaces & "                                        an example /DEL statement looks like this:")
        Console.WriteLine(SomeSpaces & "                                                   /DEL 7C:DD:90:00:00:01")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/C                                    = clear console window")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/DBB                                  = database backup")
        Console.WriteLine(SomeSpaces & "/DBD                                  = database delete")
        Console.WriteLine(SomeSpaces & "/DBE                                  = database edit")
        Console.WriteLine(SomeSpaces & "/DBR                                  = database restore from backup")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/HYPERV                               = consolodate Hyper-V related entries as 234.-.-.-")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/NP                                   = do not ping (saves time but results may be less accurate)")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/NRA                                  = do not report active devices")
        Console.WriteLine(SomeSpaces & "/NRI                                  = do not report inactive devices")
        Console.WriteLine(SomeSpaces & "/NRD                                  = do not report dynamic IP addresses")
        Console.WriteLine(SomeSpaces & "/NRS                                  = do not report static IP addresses")
        Console.WriteLine("")
        Console.WriteLine(SomeSpaces & "/P                                    = pause prompt before exit")
        Console.WriteLine(SomeSpaces & "/Q                                    = no prompts or reports (just update the database)")
        Console.WriteLine(SomeSpaces & "/R                                    = refresh device names (may take a long time)")
        Console.WriteLine("")
        Console.ForegroundColor = gOriginalForegroundColour

    End Sub

    Private Sub RefreshFromThisComputer()

        Dim gThisPCsMACs As String = String.Empty
        Dim ThisPCsIPAddress As String = String.Empty

        Dim FirstMacFound As Boolean = True

        Dim Count As Integer = 0
        Dim IP_Address, MAC_Address As String

        Dim mc As System.Management.ManagementClass
        Dim mo As System.Management.ManagementObject
        mc = New System.Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances()

        For Each mo In moc

            If mo.Item("IPEnabled") = True Then

                Dim IPAddresses() As String = mo.Item("IPAddress")
                IP_Address = IPAddresses(0)

                MAC_Address = mo.Item("MacAddress").ToString.ToUpper

                NewDeviceTable(NewDeviceTableIndex) = UpdateFromExistingDeviceInfo(MAC_Address, IP_Address, True)
                NewDeviceTable(NewDeviceTableIndex).StaticOrDynamic = cDynamic

                NewDeviceTableIndex += 1

            End If

        Next

    End Sub

    Private Sub RefreshFromNetwork()

        ' Polls the network and loads the database with information related to the currently network connected machines

        Try

            If gCommandLineSwitch.NotQuite Then
                Console.ForegroundColor = ConsoleColor.White
                Console.Write("Checking network for new devices")
            End If

            ' Run the ARP command

            Dim ps As New System.Diagnostics.ProcessStartInfo("arp", "-a ")

            ps.RedirectStandardOutput = True
            ps.UseShellExecute = False
            ps.WindowStyle = ProcessWindowStyle.Hidden
            ps.CreateNoWindow = True

            Dim sbResults As New StringBuilder

            Using proc As New System.Diagnostics.Process()

                Dim Timeout As DateTime = Now.AddSeconds(5) ' v1.1 to do testing here

                proc.StartInfo = ps
                proc.Start()
                Dim sr As System.IO.StreamReader = proc.StandardOutput
                System.Threading.Thread.Sleep(250)

                While Not proc.HasExited

                    If Timeout < Now Then Exit While
                    System.Threading.Thread.Sleep(100)

                End While

                sbResults.Append(sr.ReadToEnd)

            End Using

            Dim AllOutputLines() As String = sbResults.ToString.Split(vbCrLf)

            Dim DynamicOutputLines() As String = Nothing

            If gCommandLineSwitch.NoDynamic Then
            Else
                DynamicOutputLines = Filter(AllOutputLines, "dynamic", True, CompareMethod.Text)
            End If

            Dim StaticOutputLines() As String = Nothing

            If gCommandLineSwitch.NoStatic Then
            Else
                StaticOutputLines = Filter(AllOutputLines, "static", True, CompareMethod.Text)
            End If

            Dim FilteredOutputLines() As String = Nothing

            If StaticOutputLines IsNot Nothing Then

                If DynamicOutputLines IsNot Nothing Then
                    FilteredOutputLines = DynamicOutputLines.Concat(StaticOutputLines).Distinct.ToArray
                Else
                    FilteredOutputLines = StaticOutputLines.Distinct.ToArray
                End If

            Else

                If DynamicOutputLines IsNot Nothing Then
                    FilteredOutputLines = DynamicOutputLines.Distinct.ToArray
                Else
                    NewDeviceTableIndex = -1
                    Exit Try
                End If

            End If

            Dim IP_Address, MAC_Address As String

            If gCommandLineSwitch.NotQuite Then

                If gCommandLineSwitch.Refresh Then

                    Console.ForegroundColor = ConsoleColor.White
                    Console.WriteLine("Refreshing device names")
                    Console.WriteLine("(this may take a while)")
                    Console.WriteLine("")

                End If

            End If

            Array.Sort(FilteredOutputLines)

            For Each IndividualOutputLine As String In FilteredOutputLines

                Dim Entries() As String = IndividualOutputLine.Split(New String() {}, StringSplitOptions.RemoveEmptyEntries)

                IP_Address = Entries(0)
                MAC_Address = Entries(1).ToUpper.Replace("-", ":")

                If IsPhsyicalAddressInDatabase(NewDeviceTable, MAC_Address) Then

                    ' this should not happen but don't let the same mac be added twice

                Else

                    NewDeviceTable(NewDeviceTableIndex) = UpdateFromExistingDeviceInfo(MAC_Address, IP_Address, False)

                    If Entries(2).ToUpper = cStatic.ToUpper Then

                        NewDeviceTable(NewDeviceTableIndex).StaticOrDynamic = cStatic

                        If NewDeviceTable(NewDeviceTableIndex).Description = cUnknown Then

                            NewDeviceTable(NewDeviceTableIndex).Description = LookupIPSpecialStaticAddresses(IP_Address)

                        End If

                    Else

                        NewDeviceTable(NewDeviceTableIndex).StaticOrDynamic = cDynamic

                    End If

                    NewDeviceTableIndex += 1

                End If

            Next

            NewDeviceTableIndex -= 1

        Catch ex As Exception

            MsgBox(ex.ToString)

        End Try

        If NewDeviceTableIndex < 0 Then
            NewDeviceTable = Nothing
        Else
            ReDim Preserve NewDeviceTable(NewDeviceTableIndex)
        End If

        If gCommandLineSwitch.NotQuite Then

            If gCommandLineSwitch.Refresh Then
                Console.WriteLine("")
                Console.WriteLine("Refreshing device names complete")
            End If

        End If

        If gCommandLineSwitch.NotQuite Then
            'Erase the working line
            Console.SetCursorPosition(0, Console.CursorTop)
            Console.Write(StrDup(99, " "))
            Console.SetCursorPosition(0, Console.CursorTop)
            Console.ForegroundColor = gOriginalForegroundColour
        End If

    End Sub

    Private Sub UpdateInventoryForHyperV()

        ' the purpose of this routine is to clean-up and consolodate the many entries added by Hyper-V 
        ' 

        If gCommandLineSwitch.Hyper_V Then

            UpdateTableForHyperVCommon(ExistingDeviceTable, True)
            UpdateTableForHyperVCommon(NewDeviceTable, False)

        End If

    End Sub

    Private Sub UpdateTableForHyperVCommon(ByRef sometable() As DeviceTableRowStructure, ByVal ExistingTablePass As Boolean)

        Dim ARepersenativeEntryHasAlreadyBeenIncluded As Boolean = False

        Const DummyIP As String = cHyperVPrefix & "-.-.-"

        Dim y As Integer = 0

        If (sometable IsNot Nothing) Then

            Dim IncludeThisEntry As Boolean

            For x = 0 To sometable.Count - 1

                IncludeThisEntry = True

                If IsHyperV(sometable(x).IPAddress) Then

                    If ExistingTablePass Then

                        IncludeThisEntry = False

                    Else

                        sometable(x).IPAddress = DummyIP
                        sometable(x).Description = "* Consolodated Hyper-V entries"

                        If ARepersenativeEntryHasAlreadyBeenIncluded Then
                            IncludeThisEntry = False
                        Else
                            ARepersenativeEntryHasAlreadyBeenIncluded = True
                        End If

                    End If

                End If

                If IncludeThisEntry Then

                    sometable(y) = sometable(x)
                    y += 1

                End If

            Next

            If ExistingTablePass Then
            Else
                If y > 0 Then y -= 1
                ReDim Preserve sometable(y)
            End If

        End If

    End Sub

    Private Function IsHyperV(ByVal IPAddress As String) As Boolean

        'the criteria for consolodating hyper-v entries, assuming ip address is a.b.c.d, is as follows:

        ' a = cHyperVPrefix
        ' and 
        '    b = c = d
        '    or
        '    (b + 1 = c) and (c = d)
        '    or
        '    (b = c) and (c + 1 = d)


        Dim ReturnValue As Boolean = False

        If IPAddress.StartsWith(cHyperVPrefix) Then

            Dim Subnet() As String = Split(IPAddress, ".")

            If Subnet.Count = 4 Then

                If ((Subnet(1) = Subnet(2)) AndAlso (Subnet(2) = Subnet(3))) OrElse
                   ((Int(Subnet(1)) + 1 = Int(Subnet(2))) AndAlso (Subnet(2) = Subnet(3))) OrElse
                   ((Subnet(1) = Subnet(2)) AndAlso ((Int(Subnet(2)) + 1) = Int(Subnet(3)))) Then

                    ReturnValue = True

                End If

            End If

        End If

        Return ReturnValue

    End Function

    Private Sub UpdateInventoryForExtender()

        ' update the description of a device connected via NetGear to match the description of the actual device

        Const NETGearExtenderPrefix As String = "02:0F:B5"
        Const cViaExtender As String = " (via extender)"

        'First, confirm a NetGear device is being used

        Dim NetGearExtenderInUse As Boolean = False

        For Each entry In ExistingDeviceTable
            If entry.MAC.StartsWith(NETGearExtenderPrefix) Then
                NetGearExtenderInUse = True
                Exit For
            End If
        Next

        If NetGearExtenderInUse Then
        Else
            For Each entry In NewDeviceTable
                If entry.MAC.StartsWith(NETGearExtenderPrefix) Then
                    NetGearExtenderInUse = True
                    Exit For
                End If
            Next
        End If

        If NetGearExtenderInUse Then
        Else
            'a NetGear extender is not being used
            Exit Sub
        End If

        'Second, scan the Existing Device Table, comparing it with with info form other Existing Device Table entries 

        Dim UpdateMade As Boolean = False

        For x = 0 To ExistingDeviceTable.Count - 1

            If ExistingDeviceTable(x).MAC.StartsWith(NETGearExtenderPrefix) Then


                Dim ExtenderMac As String = ExistingDeviceTable(x).MAC
                Dim OriginalEnding As String = ExistingDeviceTable(x).MAC.Remove(0, 9)

                For y = 0 To ExistingDeviceTable.Count - 1

                    If x <> y Then

                        If ExistingDeviceTable(y).MAC.EndsWith(OriginalEnding) Then

                            'This is the original device in the existing device table


                            ' Try to set the description of the extended device to correspond to that of the original device

                            If ExistingDeviceTable(y).Description = cUnknown Then
                                ' can't do anything with an unknown description
                            Else

                                Dim OrigionalEntryx As String = ExistingDeviceTable(x).Description
                                Dim OrigionalEntryy As String = ExistingDeviceTable(y).Description

                                'Make sure the original device is not tagged with '(via extender)'
                                ExistingDeviceTable(y).Description = ExistingDeviceTable(y).Description.Replace(cViaExtender, "").Trim

                                'Set the extendor desciption to end in '(via extender)'
                                ExistingDeviceTable(x).Description = ExistingDeviceTable(y).Description.Trim & cViaExtender

                                If UpdateMade Then
                                    ' no need to check again, all that needs to be flagged is one update 
                                Else
                                    UpdateMade = (OrigionalEntryx <> ExistingDeviceTable(x).Description) OrElse (OrigionalEntryy <> ExistingDeviceTable(y).Description)
                                End If

                            End If


                            If UpdateMade Then
                            Else

                                ' Try to set the description of the original device to correspond to that of the extended device

                                If ExistingDeviceTable(x).Description = cUnknown Then
                                    ' can't do anything with and unknown description
                                Else

                                    Dim OrigionalEntryx As String = ExistingDeviceTable(x).Description
                                    Dim OrigionalEntryy As String = ExistingDeviceTable(y).Description

                                    'make sure extended device ends with '(via extender)'
                                    ExistingDeviceTable(y).Description = ExistingDeviceTable(x).Description.Replace(cViaExtender, "").Trim
                                    ExistingDeviceTable(x).Description = ExistingDeviceTable(y).Description & cViaExtender

                                    UpdateMade = (OrigionalEntryx <> ExistingDeviceTable(x).Description) OrElse (OrigionalEntryy <> ExistingDeviceTable(y).Description)

                                End If

                            End If
                            'Thats it for this particular device

                            Exit For

                        End If

                    End If

                Next

                If ExistingDeviceTable(x).Description.EndsWith(cViaExtender) Then
                Else
                    UpdateMade = True
                    ExistingDeviceTable(x).Description &= cViaExtender
                End If

            End If

        Next

        'if an update was made set all desciptions in the New Device Table to match those in the ExistingDeviceTable

        If UpdateMade Then
            For x = 0 To NewDeviceTable.Count - 1

                For y = 0 To ExistingDeviceTable.Count - 1

                    If NewDeviceTable(x).MAC = ExistingDeviceTable(y).MAC Then
                        NewDeviceTable(x).Description = ExistingDeviceTable(y).Description
                        Exit For
                    End If

                Next

            Next
        End If


    End Sub

    Private Sub UpdateInventoryForAutomaticPrivateIPAddressing()

        InventoryForAutomaticPrivateIPAddressingCommon(ExistingDeviceTable)
        InventoryForAutomaticPrivateIPAddressingCommon(NewDeviceTable)

    End Sub

    Private Sub InventoryForAutomaticPrivateIPAddressingCommon(ByRef sometable() As DeviceTableRowStructure)

        If (sometable IsNot Nothing) Then
            For x = 0 To sometable.Count - 1

                If sometable(x).IPAddress.StartsWith(cAPIAPrefix) Then

                    If sometable(x).Description.EndsWith(cUsingAPIA) Then
                    Else
                        sometable(x).Description &= cUsingAPIA
                    End If
                End If

            Next

        End If

    End Sub

    Private Function IsPhsyicalAddressInDatabase(ByVal DeviceTable() As DeviceTableRowStructure, ByVal Mac As String) As Boolean

        Dim ReturnValue As Boolean = False

        If DeviceTable Is Nothing Then
        Else
            For x = 0 To DeviceTable.Count - 1

                If Mac = DeviceTable(x).MAC Then
                    ReturnValue = True
                    Exit For
                End If

            Next
        End If

        Return ReturnValue

    End Function

    Private Function UpdateFromExistingDeviceInfo(ByVal MACAddress As String, ByVal IPAddress As String, ByVal HostQuerry As Boolean) As DeviceTableRowStructure

        ' for performance reasons first look in the existing database to find the device name
        ' if not there, second look up based on IP address (unless the command line option says not to refresh device names)

        Dim ReturnValue As DeviceTableRowStructure

        Dim MatchFound As Boolean = False

        ReturnValue.Active = "Active"
        ReturnValue.MAC = MACAddress
        ReturnValue.IPAddress = IPAddress
        ReturnValue.SortableIPAddress = GetNumericIP(IPAddress)
        ReturnValue.DeviceName = cUnknown
        ReturnValue.Description = cUnknown
        ReturnValue.LastSeen = Now.ToString("yyyy-MM-dd HH:mm:ss")
        ReturnValue.StaticOrDynamic = cUnknown

        For x As Integer = 0 To ExistingDeviceTableIndex

            'search for exact match based on MACAddress
            If ExistingDeviceTable(x).MAC = MACAddress Then
                ExistingDeviceTable(x).Active = "Processed"
                ReturnValue.DeviceName = ExistingDeviceTable(x).DeviceName
                ReturnValue.Description = ExistingDeviceTable(x).Description
                ReturnValue.StaticOrDynamic = ExistingDeviceTable(x).StaticOrDynamic
                MatchFound = True
                Exit For
            End If

        Next

        If HostQuerry Then

            ReturnValue.DeviceName = System.Net.Dns.GetHostName() ' Host querry is fast, always get name for host

        Else

            If MatchFound Then

                'Network querry is really slow, only re-querry existing names if the CommandLineRefreshOptions is set

                If gCommandLineSwitch.Refresh Then
                    ReturnValue.DeviceName = GetDeviceName(IPAddress)
                End If

            Else

                ReturnValue.DeviceName = GetDeviceName(IPAddress)

            End If

        End If

        Return ReturnValue

    End Function

    Private Function GetNumericIP(IP As String) As Long

        On Error Resume Next

        Dim ReturnValue As Long = 0

        Dim sections() As String = Split(IP, ".")

        ReturnValue = sections(0) * 16777216 + sections(1) * 65536 + sections(2) * 256 + sections(3)

        Return ReturnValue

    End Function
    Private Function GetDeviceName(ByVal IP_Address As String) As String

        Dim ReturnValue As String = cUnknown

        Dim myIPs As System.Net.IPHostEntry

        Try

            If (gCommandLineSwitch.Hyper_V) AndAlso (IsHyperV(IP_Address)) Then

                'Console.WriteLine(IP_Address & " is a Hyper-V address")

            Else

                Console.Write(IP_Address & " device name is ")
                myIPs = System.Net.Dns.GetHostByAddress(IP_Address) ' System.Net.Dns.GetHostEntry(IP_Address) 'v1.1
                ReturnValue = myIPs.HostName
                Console.WriteLine(ReturnValue)

            End If

        Catch ex As Exception

            Console.WriteLine("unavailable")

        End Try

        myIPs = Nothing

        If ReturnValue = IP_Address Then ReturnValue = cUnknown

        Return ReturnValue

    End Function

    Private Sub WriteHeaderLine(ByVal Active As Boolean)

        If gCommandLineSwitch.NoActive AndAlso Active Then Exit Sub

        If gCommandLineSwitch.NoInactive AndAlso (Not Active) Then Exit Sub

        Dim HeaderEntry As New DeviceTableRowStructure

        If Active Then
            'Console.BackgroundColor = ConsoleColor.Blue
            Console.ForegroundColor = ConsoleColor.Green
        Else
            Console.ForegroundColor = ConsoleColor.Yellow
        End If

        'Console.Clear()
        Console.WriteLine()

        With HeaderEntry
            .IPAddress = IIf(Active, "ACTIVE:", "INACTIVE:")
            .StaticOrDynamic = " "
            .MAC = " "
            .DeviceName = " "
            .Description = " "
            .LastSeen = " "
        End With
        Console_WriteFormatedLine(HeaderEntry, True)

        Console.WriteLine()

        With HeaderEntry
            .IPAddress = IIf(Active, "IP", "Last known IP")
            .StaticOrDynamic = "Type"
            .MAC = "Physical Address"
            .DeviceName = "Device name"
            .Description = "Description"
            .LastSeen = IIf(Active, " ", "Last seen")
        End With
        Console_WriteFormatedLine(HeaderEntry, False)

        With HeaderEntry
            .IPAddress = Dashes
            .StaticOrDynamic = Dashes
            .MAC = Dashes
            .DeviceName = Dashes
            .Description = Dashes
            .LastSeen = IIf(Active, " ", Dashes)
        End With
        Console_WriteFormatedLine(HeaderEntry, Active)

        Console.ForegroundColor = ConsoleColor.White

    End Sub

    Private Sub WriteActiveRecords()

        Dim RecordsFound As Boolean = False

        If NewDeviceTable IsNot Nothing Then

            System.Array.Sort(Of DeviceTableRowStructure)(NewDeviceTable, Function(x, y) x.SortableIPAddress.CompareTo(y.SortableIPAddress))

            If gCommandLineSwitch.NoActive Then Exit Sub

            For Each Entry As DeviceTableRowStructure In NewDeviceTable

                If (gCommandLineSwitch.NoDynamic AndAlso Entry.StaticOrDynamic = cDynamic) OrElse (gCommandLineSwitch.NoStatic AndAlso Entry.StaticOrDynamic = cStatic) Then
                    'skip writing to console
                Else
                    Console_WriteFormatedLine(Entry, True)
                    RecordsFound = True
                End If

            Next

            System.Threading.Thread.Sleep(500) ' provide time to finish writing entries

        End If

        If RecordsFound Then
        Else

            Dim HeaderEntry As New DeviceTableRowStructure
            With HeaderEntry
                .IPAddress = "(none)"
                .MAC = " "
                .DeviceName = " "
                .Description = " "
            End With
            Console_WriteFormatedLine(HeaderEntry, True)

        End If

    End Sub

    Private Sub WriteInactiveRecords()

        Dim RecordsFound As Boolean = False

        If ExistingDeviceTable IsNot Nothing Then

            System.Array.Sort(Of DeviceTableRowStructure)(ExistingDeviceTable, Function(x, y) x.SortableIPAddress.CompareTo(y.SortableIPAddress))

            If gCommandLineSwitch.NoInactive Then Exit Sub

            For Each Entry As DeviceTableRowStructure In ExistingDeviceTable

                If Entry.Active = "Processed" Then
                Else
                    RecordsFound = True

                    If (gCommandLineSwitch.NoDynamic AndAlso Entry.StaticOrDynamic = cDynamic) OrElse (gCommandLineSwitch.NoStatic AndAlso Entry.StaticOrDynamic = cStatic) Then
                        'skip writing to console
                    Else
                        RecordsFound = True
                        Console_WriteFormatedLine(Entry, False)
                    End If

                End If

            Next

            System.Threading.Thread.Sleep(500) ' provide time to finish writing entries

        End If

        If RecordsFound Then
        Else

            Dim HeaderEntry As New DeviceTableRowStructure
            With HeaderEntry
                .IPAddress = "(none)"
                .MAC = " "
                .DeviceName = " "
                .Description = " "
            End With
            Console_WriteFormatedLine(HeaderEntry, True)

        End If

    End Sub

#End Region

#Region "Wrapup"

    Private Sub Wrapup()

        ' hold if pause was selected

        If gCommandLineSwitch.NotQuite Then

            If gCommandLineSwitch.Pause Then
                Console_PressAnyKeyToContinue()
            End If

        End If

        Console.ForegroundColor = gOriginalForegroundColour

        Dim CloseCode As CtrlTypes = CtrlTypes.CTRL_CLOSE_EVENT
        WindowClosingHandler(CloseCode) ' used to save Window size

    End Sub

#End Region

#Region "Write to Console"

    Private Sub Console_PressAnyKeyToContinue()

        AddHandler System.AppDomain.CurrentDomain.UnhandledException, AddressOf UnhandledExceptionTrapper

        Console.ForegroundColor = ConsoleColor.Cyan
        Console.WriteLine("Press the enter key to continue ...")
        Console.ForegroundColor = ConsoleColor.Gray

        While Console.ReadKey.Key <> ConsoleKey.Enter
            'loop forever
            Console_Backspace()
            System.Threading.Thread.Sleep(50)
        End While

    End Sub

    Private Function Console_PromptForYesorNo(ByVal Message As String) As Boolean

        Dim ReturnValue As Boolean = False

        Console.ForegroundColor = ConsoleColor.White

        Console.WriteLine(Message & " (Y/N) ")

        Dim ckey As ConsoleKeyInfo = Console.ReadKey

        While True

            Select Case ckey.Key

                Case Is = ConsoleKey.Y
                    Console_Backspace()
                    ReturnValue = True
                    Exit While

                Case Is = ConsoleKey.N
                    Console_Backspace()
                    ReturnValue = False
                    Exit While

                Case Is = ConsoleKey.Enter
                    ' no backspace required

                Case Else
                    Console_Backspace()

            End Select

            ckey = Console.ReadKey

        End While

        Console.ForegroundColor = ConsoleColor.Gray

        Return ReturnValue

    End Function

    Private Sub WriteAColouredConsoleLine(ByVal Message As String, ByVal Colour As ConsoleColor, ByVal Optional AddExtraLineBeforeMessage As Boolean = True, ByVal Optional AddExtraLineAfterMessage As Boolean = True)

        If gCommandLineSwitch.NotQuite Then

            Console.ForegroundColor = Colour

            If AddExtraLineBeforeMessage Then Console.WriteLine(" ")

            Console.WriteLine(Message)

            If AddExtraLineAfterMessage Then Console.WriteLine(" ")

            Console.ForegroundColor = gOriginalForegroundColour

        End If

    End Sub

    Private Sub Console_Backspace()

        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop)
        Console.Write(" ")
        Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop)

    End Sub

    Private Sub Console_WriteFormatedLine(ByVal entry As DeviceTableRowStructure, ByVal Active As Boolean)

        Dim a, b, c, d, e, f As String

        With entry
            a = CreateFixedWidthString(.IPAddress, 15) & " "
            b = CreateFixedWidthString(.StaticOrDynamic, 7) & " "
            c = CreateFixedWidthString(.MAC, 17) & " "
            d = CreateFixedWidthString(.DeviceName, 15) & " "
            e = CreateFixedWidthString(.Description, 64) & " "
            f = CreateFixedWidthString(.LastSeen, 19)
        End With

        If Active Then
            Console.WriteLine(a & b & c & d & e)
        Else
            Console.WriteLine(a & b & c & d & e & f)
        End If

    End Sub

    Private Function CreateFixedWidthString(ByVal ws As String, ww As Integer) As String

        If (ws Is Nothing) OrElse (ws.Length = 0) Then ws = " "
        Return (ws & StrDup(ww, " ")).Remove(ww)

    End Function

#End Region

#Region "Common IP Addresses"

    'ref https://en.wikipedia.org/wiki/Multicast_address
    Const LookupTable As String =
"
224.0.0.0       Base address (reserved)
224.0.0.1       All Hosts multicast group 
224.0.0.2       All Routers multicast group 
224.0.0.4       Distance Vector Multicast Routing Protocol
224.0.0.5       Open Shortest Path First to send Hello packets 
224.0.0.6       All Designated Routers send OSPF routing info
224.0.0.9       Routing Information Protocol (RIP) v2 group 
224.0.0.10      Enhanced Interior Gateway Routing Protocol group 
224.0.0.13      Protocol Independent Multicast (PIM) v2
224.0.0.18      Virtual Router Redundancy Protocol
224.0.0.19      IS-IS over IP
224.0.0.20      IS-IS over IP
224.0.0.21      IS-IS over IP
224.0.0.22      Internet Group Management Protocol v3[2]
224.0.0.102     Hot Standby Router Protocol v2 / Gateway Load Balancing Protocol
224.0.0.107     Precision Time Protocol v2
224.0.0.251     Multicast DNS address
224.0.0.252     Link-local Multicast Name Resolution address
224.0.0.253     Teredo tunneling client discovery address[3]
224.0.1.1       Network Time Protocol clients listener
224.0.1.22      Service Location Protocol v1 general
224.0.1.35      Service Location Protocol v1 directory agent
224.0.1.39      Cisco multicast router AUTO-RP-ANNOUNCE 
224.0.1.40      Cisco multicast router AUTO-RP-DISCOVERY 
224.0.1.41      H.323 Gatekeeper discovery address
224.0.1.129     Precision Time Protocol v1 
224.0.1.130     Precision Time Protocol v1 
224.0.1.131     Precision Time Protocol v1 
224.0.1.132     Precision Time Protocol v1 
224.0.1.129     Precision Time Protocol v2 
239.255.255.250 Simple Service Discovery Protocol 
239.255.255.253 Service Location Protocol v2
"
    Private Function LookupIPSpecialStaticAddresses(ByVal IPAddress As String) As String

        Static Dim LookupTableEntries() As String
        If LookupTableEntries Is Nothing Then
            LookupTableEntries = LookupTable.Split(vbCrLf)
            For x = 0 To LookupTableEntries.Count - 1
                LookupTableEntries(x) = LookupTableEntries(x).Trim
            Next
        End If

        Dim ReturnValue As String = "Unknown"

        If IPAddress.EndsWith(".255") Then

            ReturnValue = "* Broadcast on " & IPAddress.Remove(IPAddress.Length - 4)

        Else

            For Each entry In LookupTableEntries

                If entry.StartsWith(IPAddress) Then
                    ReturnValue = "* " & entry.Replace(IPAddress, " ").TrimStart
                    Exit For
                End If

            Next

        End If

        Return ReturnValue

    End Function

#End Region

    Private Sub PingDevicesLoadedInMemory()

        If gCommandLineSwitch.NoPing Then Exit Sub

        If ExistingDeviceTableIndex = 0 Then Exit Sub

        If gCommandLineSwitch.NotQuite Then
            Console.ForegroundColor = ConsoleColor.White
            Console.Write("Pinging known devices ")
        End If

        Dim UniqueListOfIPAddresses As New List(Of String)
        For j = 0 To ExistingDeviceTable.Count - 1
            UniqueListOfIPAddresses.Add(ExistingDeviceTable(j).IPAddress)
        Next
        UniqueListOfIPAddresses = UniqueListOfIPAddresses.Distinct().ToList

        'use parallel processing to speed things up
        Parallel.ForEach(UniqueListOfIPAddresses, Sub(entry)

                                                      Try
                                                          Dim PingSender As New NetworkInformation.Ping
                                                          PingSender.Send(entry, 1500)  ' 5000 > 1500 v1.1
                                                      Catch ex As Exception
                                                      End Try

                                                  End Sub)


        If gCommandLineSwitch.NotQuite Then
            'Erase the pinging line
            Console.SetCursorPosition(0, Console.CursorTop)
            Console.Write(StrDup(99, " "))
            Console.SetCursorPosition(0, Console.CursorTop)
            Console.ForegroundColor = gOriginalForegroundColour
        End If

        Console.Clear()

    End Sub

#Region "Database functions"

    Private Sub SetDatabaseName()

        gProgramName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name

        Dim WorkingPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & gProgramName

        If Directory.Exists(WorkingPath) Then
        Else
            Directory.CreateDirectory(WorkingPath)
        End If

        gDatabasePathAndFilename = WorkingPath & "\Database.xml"
        gBackupDatabasePathAndFilename = WorkingPath & "\Database.backup"

        Try
            If File.Exists(gDatabasePathAndFilename) Then
            Else
                File.Copy(AppDomain.CurrentDomain.BaseDirectory & "\Database.xml", gDatabasePathAndFilename)
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString)
        End Try

    End Sub

    Private Sub LoadDatabaseInMemory()

        If File.Exists(gDatabasePathAndFilename) Then

            Dim wdb = My.Computer.FileSystem.ReadAllText(gDatabasePathAndFilename)

            Dim xmlDoc As New XmlDocument()
            xmlDoc.Load(gDatabasePathAndFilename)
            Dim nodes As XmlNodeList = xmlDoc.DocumentElement.SelectNodes("/Devices/Device")

            For Each node As XmlNode In nodes

                With ExistingDeviceTable(ExistingDeviceTableIndex)
                    .Active = node.SelectSingleNode("Active").InnerText
                    .IPAddress = node.SelectSingleNode("IP_Address").InnerText
                    .SortableIPAddress = GetNumericIP(.IPAddress)
                    .MAC = node.SelectSingleNode("Physical_address").InnerText
                    .DeviceName = node.SelectSingleNode("Device_name").InnerText
                    .Description = node.SelectSingleNode("Description").InnerText
                    .LastSeen = node.SelectSingleNode("Last_seen").InnerText
                    .StaticOrDynamic = node.SelectSingleNode("Static_or_Dynamic").InnerText
                End With

                ExistingDeviceTableIndex += 1

            Next

            If ExistingDeviceTableIndex > 0 Then ExistingDeviceTableIndex -= 1

        Else

            gCommandLineSwitch.Refresh = True
            ExistingDeviceTableIndex = 0

        End If

        ReDim Preserve ExistingDeviceTable(ExistingDeviceTableIndex)

    End Sub

    Private Sub WriteDatabaseFromMemory()

        Dim ws As String = String.Empty

        Dim writer As New XmlTextWriter(gDatabasePathAndFilename, System.Text.Encoding.ASCII)

        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2

        writer.WriteStartElement("Devices")

        If NewDeviceTable IsNot Nothing Then

            For Each entry In NewDeviceTable
                CreateNode(entry, writer)
            Next

        End If

        For Each entry In ExistingDeviceTable

            If entry.Active <> "Processed" Then

                If IsPhsyicalAddressInDatabase(NewDeviceTable, entry.MAC) Then
                    ' if already in the New DeviceTable don't repeat it in the list of inactive devices
                Else
                    CreateNode(entry, writer)
                End If

            End If

        Next

        writer.WriteEndElement()

        writer.WriteEndDocument()
        writer.Close()

    End Sub

    Private Sub CreateNode(ByVal entry As DeviceTableRowStructure, ByVal writer As XmlTextWriter)

        With entry

            writer.WriteStartElement("Device")

            writer.WriteStartElement("Active")
            writer.WriteString(.Active)
            writer.WriteEndElement()

            writer.WriteStartElement("Static_or_Dynamic")
            writer.WriteString(.StaticOrDynamic)
            writer.WriteEndElement()

            writer.WriteStartElement("IP_Address")
            writer.WriteString(.IPAddress)
            writer.WriteEndElement()

            writer.WriteStartElement("Physical_address")
            writer.WriteString(.MAC)
            writer.WriteEndElement()

            writer.WriteStartElement("Device_name")
            writer.WriteString(.DeviceName)
            writer.WriteEndElement()

            writer.WriteStartElement("Description")
            writer.WriteString(.Description)
            writer.WriteEndElement()

            writer.WriteStartElement("Last_seen")
            writer.WriteString(.LastSeen)
            writer.WriteEndElement()

            writer.WriteEndElement()

        End With

    End Sub

#End Region

#Region "Save window size when done"

    ' the following will save the window size if the user exits the application early,
    ' for example by clicking the red 'X' in the top right hand corner of the window.
    ' it is also called by the Wrapup subroutine as the program ends naturally

    Public Declare Auto Function SetConsoleCtrlHandler Lib "kernel32.dll" (ByVal Handler As HandlerRoutine, ByVal Add As Boolean) As Boolean

    Public Delegate Function HandlerRoutine(ByVal CtrlType As CtrlTypes) As Boolean

    Public Enum CtrlTypes
        CTRL_C_EVENT = 0
        CTRL_BREAK_EVENT
        CTRL_CLOSE_EVENT
        CTRL_LOGOFF_EVENT = 5
        CTRL_SHUTDOWN_EVENT
    End Enum

    Public Function WindowClosingHandler(ByVal ctrlType As CtrlTypes) As Boolean

        If (My.Settings.Height <> Console.WindowHeight) OrElse (My.Settings.Width <> Console.WindowWidth) Then
            My.Settings.Height = Console.WindowHeight
            My.Settings.Width = Console.WindowWidth
            My.Settings.Save()
        End If

        System.Threading.Thread.Sleep(1000)

        Return True

    End Function
    Private Sub UnhandledExceptionTrapper(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)

        Dim CloseCode As CtrlTypes = CtrlTypes.CTRL_CLOSE_EVENT
        WindowClosingHandler(CloseCode) ' used to save Window size
        'Console.WriteLine(e.ExceptionObject.ToString())
        System.Threading.Thread.Sleep(1000)
        Environment.[Exit](0)

    End Sub

#End Region

End Module
