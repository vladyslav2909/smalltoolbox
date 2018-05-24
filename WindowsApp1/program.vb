Imports System.IO
Imports System.Environment
Imports System.Management
Imports System.Threading
Imports Microsoft.Win32



Public Class aranjare
    'temp files 
    Dim tempclean As Thread
    Dim tempFolderPath As String = System.IO.Path.GetTempPath()
    'prefetch
    Dim prefectchPath As String = "%systemroot%/Prefetch"
    'recycle bin
    Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Int32, ByVal pszRootPath As String, ByVal dwFlags As Int32) As Int32
    Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Int32
    Private Const SHREB_NOPROGRESSUI = &H2 'userinferface
    Private Const SHREB_NOSOUND = &H4 'fara sunet
    'sub recycle - sterge recylce fara sunet
    Private Sub EmptyRecycleBin()
        SHEmptyRecycleBin(Me.Handle.ToInt32, "", SHREB_NOSOUND)
        SHUpdateRecycleBinIcon()

    End Sub


    'time
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        MetroLabel6.Text = DateTime.Now.ToString("HH:mm:ss")

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Timer1.Enabled = True 'timer data exacta
        MetroRadioButton1.Checked = True
        Timer3.Start() 'timer alarma



        Me.StyleManager = MetroStyleManager1
        'system componente
        Width.Text = Screen.PrimaryScreen.WorkingArea.Width
        Height.Text = Screen.PrimaryScreen.WorkingArea.Height
        WINLabel.Text = My.Computer.Info.OSFullName
        CPUlabel.Text = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\CentralProcessor\0", "ProcessorNameString", Nothing)
        MBLabel.Text = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardProduct", Nothing)
        'butoane fara margine
        CloseBTN.Region = New Region(New Rectangle(4, 4, CloseBTN.Width - 8, CloseBTN.Height - 8))
        BtnTAB1.Region = New Region(New Rectangle(4, 4, BtnTAB1.Width - 8, BtnTAB1.Height - 8))
        BtnTAB2.Region = New Region(New Rectangle(4, 4, BtnTAB2.Width - 8, BtnTAB2.Height - 8))
        BtnTAB3.Region = New Region(New Rectangle(4, 4, BtnTAB3.Width - 8, BtnTAB3.Height - 8))
        BtnTAB4.Region = New Region(New Rectangle(4, 4, BtnTAB4.Width - 8, BtnTAB4.Height - 8))
        SortBTN.Region = New Region(New Rectangle(4, 4, SortBTN.Width - 8, SortBTN.Height - 8))
        UpdateBTN.Region = New Region(New Rectangle(4, 4, UpdateBTN.Width - 8, UpdateBTN.Height - 8))
        BTNTab5.Region = New Region(New Rectangle(4, 4, BTNTab5.Width - 8, BTNTab5.Height - 8))
        AskBTN.Region = New Region(New Rectangle(4, 4, AskBTN.Width - 8, AskBTN.Height - 8))




        'preia hwid

        Dim hw As New clsComputerInfo

        Dim hdd As String
        Dim cpu As String
        Dim mb As String
        Dim mac As String
        Dim hwid As String

        cpu = hw.GetProcessorId()
        hdd = hw.GetVolumeSerial("C")
        mb = hw.GetMotherBoardID()
        mac = hw.GetMACAddress()
        hwid = cpu + hdd + mb + mac

        Dim hwidEncrypted As String = Strings.UCase(hw.getMD5Hash(cpu & hdd & mb & mac))

        HwidLabel.Text = hwidEncrypted
    End Sub
    'cautare id componente
    Public Class clsComputerInfo

        Friend Function GetProcessorId() As String
            Dim strProcessorId As String = String.Empty
            Dim query As New SelectQuery("Win32_processor")
            Dim search As New ManagementObjectSearcher(query)
            Dim info As ManagementObject

            For Each info In search.Get()
                strProcessorId = info("processorId").ToString()
            Next
            Return strProcessorId

        End Function

        Friend Function GetMACAddress() As String

            Dim mc As ManagementClass = New ManagementClass("Win32_NetworkAdapterConfiguration")
            Dim moc As ManagementObjectCollection = mc.GetInstances()
            Dim MACAddress As String = String.Empty
            For Each mo As ManagementObject In moc

                If (MACAddress.Equals(String.Empty)) Then
                    If CBool(mo("IPEnabled")) Then MACAddress = mo("MacAddress").ToString()

                    mo.Dispose()
                End If
                MACAddress = MACAddress.Replace(":", String.Empty)

            Next
            Return MACAddress
        End Function

        Friend Function GetVolumeSerial(Optional ByVal strDriveLetter As String = "C") As String

            Dim disk As ManagementObject = New ManagementObject(String.Format("win32_logicaldisk.deviceid=""{0}:""", strDriveLetter))
            disk.Get()
            Return disk("VolumeSerialNumber").ToString()
        End Function

        Friend Function GetMotherBoardID() As String

            Dim strMotherBoardID As String = String.Empty
            Dim query As New SelectQuery("Win32_BaseBoard")
            Dim search As New ManagementObjectSearcher(query)
            Dim info As ManagementObject
            For Each info In search.Get()

                strMotherBoardID = info("product").ToString()

            Next
            Return strMotherBoardID

        End Function


        Friend Function getMD5Hash(ByVal strToHash As String) As String
            Dim md5Obj As New Security.Cryptography.MD5CryptoServiceProvider
            Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(strToHash)

            bytesToHash = md5Obj.ComputeHash(bytesToHash)

            Dim strResult As String = ""

            For Each b As Byte In bytesToHash
                strResult += b.ToString("x2")
            Next

            Return strResult
        End Function


    End Class

    'tab1
    Private Sub BtnTAB1_Click(sender As Object, e As EventArgs) Handles BtnTAB1.Click
        MetroTabControl1.SelectedTab = MetroTabPage1
    End Sub
    'tab2
    Private Sub BtnTAB2_Click(sender As Object, e As EventArgs) Handles BtnTAB2.Click
        MetroTabControl1.SelectedTab = MetroTabPage2
    End Sub
    'tab3
    Private Sub BtnTAB3_Click(sender As Object, e As EventArgs) Handles BtnTAB3.Click
        MetroTabControl1.SelectedTab = MetroTabPage3
    End Sub
    'tab4
    Private Sub BtnTAB4_Click(sender As Object, e As EventArgs) Handles BtnTAB4.Click
        MetroTabControl1.SelectedTab = MetroTabPage4
    End Sub
    'tab5
    Private Sub BtnTAB5_Click(sender As Object, e As EventArgs) Handles BTNTab5.Click
        MetroTabControl1.SelectedTab = MetroTabPage5
    End Sub

    'aranjare
    Private Sub SortBTN_Click(sender As Object, e As EventArgs) Handles SortBTN.Click
        'porneste aranjarea
        Dim dir As String = My.Computer.FileSystem.SpecialDirectories.Temp
        Dim filename As String = dir + "af.exe"
        IO.File.WriteAllBytes(filename, My.Resources.af)
        Process.Start(filename)
    End Sub
    'quit button
    Private Sub CloseBTN_Click_1(sender As Object, e As EventArgs) Handles CloseBTN.Click
        Close()
        Try
            Dim aranjare1234() As Process = Process.GetProcessesByName("aranjarefisiere")
            For Each Process As Process In aranjare1234
                Process.Kill()
            Next
        Catch ex As Exception
        End Try

    End Sub
    'toggle dark/light mode
    Private Sub ToggleBTN_CheckedChanged(sender As Object, e As EventArgs) Handles ToggleBTN.CheckedChanged
        'schimba culoarea dark/light

        If Me.MetroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark Then
            Me.MetroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Light
        Else
            Me.MetroStyleManager1.Theme = MetroFramework.MetroThemeStyle.Dark
        End If


    End Sub
    'alarma buton start
    Private Sub MetroButton1_Click(sender As Object, e As EventArgs) Handles MetroButton1.Click
        Timer4.Start()
        AlarmTimeLabel.Text = AlarmTextBox1.Text
        MetroLabel8.Text = "Alarma este pornita"

    End Sub
    'timer4 alarma
    Private Sub Timer4_Tick(sender As Object, e As EventArgs) Handles Timer4.Tick
        If AlarmTimeLabel.Text = MetroLabel6.Text Then
            MetroLabel9.Text = AlarmTimeLabel.Text

        End If
        If MetroLabel9.Text = AlarmTimeLabel.Text And MetroRadioButton1.Checked = True Then

            Timer4.Stop()

            MessageBox.Show(TextBox1.Text)
        ElseIf MetroLabel9.Text = AlarmTimeLabel.Text And MetroRadioButton2.Checked = True Then
            Console.Beep()


        End If
    End Sub
    ' alarma buton stop
    Private Sub MetroButton2_Click(sender As Object, e As EventArgs) Handles MetroButton2.Click
        Timer4.Stop()
        MetroLabel8.Text = "Alarma este oprita"

    End Sub
    'alarm btn autodelete 00:00:00
    Private Sub AlarmTextBox1_Click(sender As Object, e As EventArgs) Handles AlarmTextBox1.Click

        AlarmTextBox1.Text = Nothing
    End Sub

    'text box alarma 
    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        TextBox1.Text = Nothing

    End Sub
    'textbox alarma in caz ca nu e nimic pune mesaj custom
    Private Sub TextBox1_LostFocus(sender As Object, e As EventArgs) Handles TextBox1.LostFocus
        If TextBox1.Text = Nothing Then
            TextBox1.Text = "Mesaj Custom"
        End If
    End Sub
    'updatebtn

    Private Sub UpdateBTN_Click(sender As Object, e As EventArgs) Handles UpdateBTN.Click
        CheckForUpdates()
    End Sub
    'declarare checkforupdates
    Public Sub CheckForUpdates()
        Dim request As System.Net.HttpWebRequest = System.Net.HttpWebRequest.Create("https://dl.dropbox.com/s/ih9hsm5k4rjqxuf/Version.txt?dl=0")
        Dim response As System.Net.HttpWebResponse = request.GetResponse()
        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(response.GetResponseStream())
        Dim newestversion As String = sr.ReadToEnd()
        Dim currentversion As String = Application.ProductVersion
        If newestversion.Contains(currentversion) Then
            MsgBox("Aveti ultima versiune.")
        Else
            MsgBox("Este un update nou, se va auto-descarca in cateva momente.")
            WebBrowser1.Navigate("https://dl.dropbox.com/s/787t8xwonxyqkoh/SmallToolBox.exe?dl=0")
        End If
    End Sub
    'sub clean()
    Sub clean()
        For Each filePath In Directory.GetFiles(tempFolderPath)
            Try
                File.Delete(filePath)
            Catch
            End Try
        Next
        For Each filePath In Directory.GetFiles("\Windows\Temp")
            Try
                File.Delete(filePath)
            Catch

            End Try
        Next
    End Sub
    'temp files
    Private Sub MetroToggle1_CheckedChanged(sender As Object, e As EventArgs) Handles MetroToggle1.CheckedChanged
        If MetroToggle1.Checked = True Then

            tempclean = New System.Threading.Thread(AddressOf clean)
            tempclean.Start()
            MsgBox("Fisierele temporale au fost sterse cu succes.", MsgBoxStyle.Information)
        Else
            MsgBox("Ai oprit stergerea fisierelor temporale.")

        End If

    End Sub
    'prefetch
    Private Sub MetroToggle2_CheckedChanged(sender As Object, e As EventArgs) Handles MetroToggle2.CheckedChanged
        If MetroToggle2.Checked = True Then
            Dim prefetchPath As String
            prefetchPath = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine) & "\Prefetch"

            For Each file In IO.Directory.GetFiles(prefetchPath)
                If file.Contains(".pf") Then IO.File.Delete(file)
            Next
            MsgBox("Ai sters cu succes fisierele din Prefetch!")

        Else
            MsgBox("Ai oprit stergerea fisierelor.")
        End If

    End Sub

    Private Sub MetroToggle3_CheckedChanged(sender As Object, e As EventArgs) Handles MetroToggle3.CheckedChanged
        If MetroToggle3.Checked = True Then
            EmptyRecycleBin()
            MsgBox("Ai sters Recycle Bin cu succes.")
        Else

            MsgBox("Ai oprit stergerea Recycle Bin-ului.")

        End If
    End Sub
    'questionbtn winvsita
    Private Sub AskBTN_Click(sender As Object, e As EventArgs) Handles AskBTN.Click
        MsgBox("Daca incerci sa stergi fisierele TEMP si ai Windows Vista, te rog ruleaza acest program cu admin rights.")
    End Sub
End Class
