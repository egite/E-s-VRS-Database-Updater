Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Module Main_Program
    Public dbPath As String, Want_Exit As Long, NoSilsFound As Integer, FirstRun As Integer, BackupVRSdb As String
    Public UserName As String, FAA_db_Download As String, Run_on_Start As String, AutoStartWait As Integer
    Public VRSdbPath As String, TimeStarted As String, Complete As String, OpenSky_Download As String
    Sub Main_Loop()
        'If Not File.Exists(dbPath & "Sils.csv") Then
        '    MessageBox.Show("Warning:  I can't find the ""Sils.csv"" file." & vbCrLf & vbCrLf &
        '                    "Please place this file in the working folder location (shown in the program window) and run the program again." & vbCrLf & vbCrLf &
        '                    "Exiting.", "E's VRS Updater")
        '    Application.Exit()
        '    End
        'End If
        TimeStarted = DateTime.Now.ToString("HHmm")
        Form1.TextBox1.AppendText("Started at:  " & TimeStarted & "." & vbCrLf)
        Form1.TextBox1.Update()
        If VRSdbPath = dbPath Then
            MessageBox.Show("Working folder and VRS folder cannot be the same location." & vbCrLf & "Fix this and run again." & vbCrLf & "Exiting.", "E's VRS Updater")
            Exit Sub
        End If
        If FAA_db_Download = "Y" Then
            Form1.TextBox1.Text = "Downloading FAA database (unresponsive until complete)..."
            Form1.TextBox1.Update()
            Application.DoEvents()
            If File.Exists(dbPath & "ReleasableAircraft.zip") Then
                My.Computer.FileSystem.DeleteFile(dbPath & "ReleasableAircraft.zip")
            End If
            My.Computer.Network.DownloadFile("https://registry.faa.gov/database/ReleasableAircraft.zip", dbPath & "ReleasableAircraft.zip")
            Form1.TextBox1.Text = "Downloading FAA database...Done."
            Form1.TextBox1.AppendText(vbCrLf & "Unzipping FAA database...")
            Form1.TextBox1.Update()
            Application.DoEvents()
            If Directory.Exists(dbPath & "fsdc3245092fm2") Then
                System.IO.Directory.Delete(dbPath & "fsdc3245092fm2", True)
            End If
            System.IO.Directory.CreateDirectory(dbPath & "fsdc3245092fm2")
            System.IO.Directory.Delete(dbPath & "fsdc3245092fm2", True)
            System.IO.Compression.ZipFile.ExtractToDirectory(dbPath & "ReleasableAircraft.zip", dbPath & "fsdc3245092fm2", overwriteFiles:=1)
            Application.DoEvents()
            Form1.TextBox1.Text = "Downloading FAA database...Done." & vbCrLf & "Unzipping FAA database...Done."
            Form1.TextBox1.Update()
            Application.DoEvents()
            Call FAA_to_SQL()
        End If
        Call CCAR_to_SQL()
        Call GetOpenSky()
        Call Update_VRS()
        Form1.TextBox1.AppendText(vbCrLf & "Started:  " & TimeStarted & ".  Ended:  " & DateTime.Now.ToString("HHmm"))
        Form1.TextBox1.Update()
        Form1.TextBox1.Text = Form1.TextBox1.Text & vbCrLf & "Done."
        Form1.TextBox1.Update()
        Form1.TextBox4.AppendText(vbCrLf & "Done.")
        Form1.TextBox4.Update()
        Form1.TextBox6.Text = "100%"
        Form1.TextBox6.Update()
        Form1.TextBox7.Text = " - "
        Form1.TextBox7.Update()
        If File.Exists("C:\Users\" & UserName & "\AppData\Local\VirtualRadar\AircraftOnlineLookupCache.sqb") Then
            File.Copy(dbPath & "AircraftOnlineLookupCache.sqb", VRSdbPath & "\AircraftOnlineLookupCache.sqb", True)
        End If
        If Run_on_Start = "y" Then
            Application.Exit()
            End
        End If
    End Sub
End Module