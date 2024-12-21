
Imports System.ComponentModel
Imports System.Data.Common
Imports System.Data.SQLite
Imports System.Threading
Imports System.Diagnostics
Imports System.Data.Entity.Core.Common.CommandTrees
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1
    Private Sub Form1_Start(sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String, i As Integer
        UserName = Environment.UserName
        TextBox10.Text = Environment.CurrentDirectory
        If Dir(Environment.CurrentDirectory & "\settings.sqb") = vbNullString Then
            connection.Open()
            command.CommandText = "CREATE TABLE Settings (Name Char, Information Char)"
            command.ExecuteNonQuery()
            ToEnter = "('dbPath','" & Replace(Environment.CurrentDirectory, "'", "''") & "\" & "')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('FAA_db_Download','Y')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('BackupVRSdb','Y')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            ToEnter = "('VRSdbPath','C:\Users\" & UserName & "\AppData\Local\VirtualRadar\')"
            command.ExecuteNonQuery()
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('Run_on_Start','N')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('Complete','N')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('OpenSky_Download','Y')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('FAA_URL','https://registry.faa.gov/database/ReleasableAircraft.zip')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            ToEnter = "('OpenSky_URL','https://s3.opensky-network.org/data-samples/metadata/aircraftDatabase.zip')"
            command.CommandText = "INSERT INTO Settings (Name, Information) VALUES " & ToEnter
            command.ExecuteNonQuery()
            connection.Close()
        End If
        connection.Open()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'dbPath';"
        dbPath = command.ExecuteScalar()
        If Strings.Right(dbPath, 1) <> "\" Then
            dbPath = dbPath & "\"
        End If
        TextBox10.Text = dbPath
        TextBox10.Update()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'FAA_URL';"
        FAA_URL = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'OpenSky_URL';"
        OpenSky_URL = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'VRSdbPath';"
        VRSdbPath = command.ExecuteScalar()
        If Strings.Right(VRSdbPath, 1) <> "\" Then
            VRSdbPath = VRSdbPath & "\"
        End If
        TextBox11.Text = VRSdbPath
        TextBox11.Update()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'FAA_db_Download';"
        FAA_db_Download = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'VRSdbPath';"
        VRSdbPath = command.ExecuteScalar()
        TextBox11.Text = VRSdbPath
        TextBox11.Update()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'Run_on_Start';"
        Run_on_Start = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'BackupVRSdb';"
        BackupVRSdb = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'Complete';"
        Complete = command.ExecuteScalar()
        command.CommandText = "SELECT Information FROM Settings WHERE Name = 'OpenSky_Download';"
        OpenSky_Download = command.ExecuteScalar()
        connection.Close()
        If Complete = "Y" Then
            CheckBox4.Checked = True
        Else
            CheckBox4.Checked = False
        End If
        CheckBox4.Update()
        If OpenSky_Download = "Y" Then
            CheckBox5.Checked = True
        Else
            CheckBox5.Checked = False
        End If
        CheckBox5.Update()
        If BackupVRSdb = "Y" Then
            CheckBox3.Checked = True
        Else
            CheckBox3.Checked = False
        End If
        CheckBox3.Update()
        If FAA_db_Download = "Y" Then
            CheckBox2.Checked = True
        Else
            CheckBox2.Checked = False
        End If
        CheckBox2.Update()
        If Run_on_Start = "Y" Then
            CheckBox1.Checked = True
        Else
            CheckBox1.Checked = False
        End If
        CheckBox1.Update()
        TextBox10.Text = dbPath
        TextBox10.Update()
        If FirstRun = 0 Then
            FirstRun = 1
        End If
        If Run_on_Start = "Y" Then
            i = 10
            AutoStartWait = 1
            While i <> 0
                TextBox1.Text = "Starting in " & i & " seconds..."
                TextBox1.Update()
                Application.DoEvents()
                Threading.Thread.Sleep(1000)
                Application.DoEvents()
                i = i - 1
                If Run_on_Start = "N" Then
                    MessageBox.Show("Autostart terminated." & vbCrLf & vbCrLf & "When the program starts next time, you must hit the green ""Start"" button to begin processing the databases.", "E's VRS Updater")
                    i = 0
                End If
            End While
            TextBox1.Text = ""
            TextBox1.Update()
            If Run_on_Start = "Y" Then
                Call Main_Loop()
                End
            End If
        End If
    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        If MessageBox.Show("Exit E's VRS Updater?", "Question", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Want_Exit = 1
        Else
            e.Cancel = True
        End If
    End Sub
    Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Main_Loop()
    End Sub
    Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim webAddress As String = "https://wwwapps.tc.gc.ca/saf-sec-sur/2/ccarcs-riacc/DDZip.aspx"
        MessageBox.Show("Visit this address in your browser to obtain the CCAR registration database.  It's been copied into your clipboard." & vbCrLf & vbCrLf &
                        webAddress & vbCrLf & vbCrLf &
                        "After downloading the file ""ccarcsdb.zip"" from the website, be sure to place it in the working folder that you've chosen (as shown in the program window).",
                        "E's VRS Updater")
        My.Computer.Clipboard.SetText(webAddress)
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged

    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged

    End Sub

    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs)

    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click
        MsgBox("Version 1.0." & vbCrLf & vbCrLf & "Please visit the git page for additional information.", vbOKOnly, “E's VRS Updater”)
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
        Application.Exit()
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String
        connection.Open()
        If CheckBox1.Checked = True Then
            ToEnter = "Information = 'Y'"
            Run_on_Start = "Y"
            If FirstRun = 1 Then
                MessageBox.Show("When the program starts next time, it will immediately begin processing the databases.", "E's VRS Updater")
            End If
        Else
            ToEnter = "Information = 'N'"
            Run_on_Start = "N"
            If FirstRun = 1 And AutoStartWait = 0 Then
                MessageBox.Show("When the program starts next time, you must hit the green ""Start"" button to begin processing the databases.", "E's VRS Updater")
            End If
        End If
        command.CommandText = "UPDATE Settings SET " & ToEnter & " WHERE Name='Run_on_Start'"
        command.ExecuteNonQuery()
        connection.Close()
        CheckBox1.Update()
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String
        connection.Open()
        If CheckBox2.Checked = True Then
            ToEnter = "Information = 'Y'"
            FAA_db_Download = "Y"
        Else
            ToEnter = "Information = 'N'"
            FAA_db_Download = "N"
        End If
        command.CommandText = "UPDATE Settings SET " & ToEnter & " WHERE Name='FAA_db_Download'"
        command.ExecuteNonQuery()
        connection.Close()
        CheckBox2.Update()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim dialog As New FolderBrowserDialog()
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection)
        If DialogResult.OK = dialog.ShowDialog Then
            dbPath = dialog.SelectedPath
            TextBox10.Text = dbPath
            TextBox10.Update()
            connection.Open()
            command.CommandText = "UPDATE Settings Set Information = '" & dbPath & "' WHERE Name='dbPath'"
            command.ExecuteNonQuery()
            connection.Close()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim dialog As New FolderBrowserDialog()
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection)
        If DialogResult.OK = dialog.ShowDialog Then
            VRSdbPath = dialog.SelectedPath
            TextBox11.Text = VRSdbPath
            TextBox11.Update()
            connection.Open()
            command.CommandText = "UPDATE Settings SET Information = '" & VRSdbPath & "' WHERE Name='dbPath'"
            command.ExecuteNonQuery()
            connection.Close()
        End If
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String
        connection.Open()
        If CheckBox3.Checked = True Then
            ToEnter = "Information = 'Y'"
            BackupVRSdb = "Y"
        Else
            ToEnter = "Information = 'N'"
            BackupVRSdb = "N"
        End If
        command.CommandText = "UPDATE Settings SET " & ToEnter & " WHERE Name='BackupVRSdb'"
        command.ExecuteNonQuery()
        connection.Close()
        CheckBox3.Update()
    End Sub


    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String
        connection.Open()
        If CheckBox4.Checked = True Then
            ToEnter = "Information= 'Y'"
            Complete = "Y"
        Else
            ToEnter = "Information = 'N'"
            Complete = "N"
        End If
        command.CommandText = "UPDATE Settings SET " & ToEnter & " WHERE Name='Complete'"
        command.ExecuteNonQuery()
        connection.Close()
        CheckBox4.Update()
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        Dim connection As New SQLiteConnection("DataSource=" & Environment.CurrentDirectory & "\settings.sqb")
        Dim command As New SQLiteCommand("", connection), ToEnter As String
        connection.Open()
        If CheckBox5.Checked = True Then
            ToEnter = "Information= 'Y'"
            OpenSky_Download = "Y"
        Else
            ToEnter = "Information = 'N'"
            OpenSky_Download = "N"
        End If
        command.CommandText = "UPDATE Settings SET " & ToEnter & " WHERE Name='OpenSky_Download'"
        command.ExecuteNonQuery()
        connection.Close()
        CheckBox4.Update()
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged

    End Sub

End Class
