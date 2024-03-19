Imports System.Data.SQLite
Imports System.IO

Module OpenSky
    Sub GetOpenSky()
        If OpenSky_Download = "N" Then
            Form1.TextBox1.AppendText(vbCrLf & "Not downloading OpenSky database.")
            Form1.TextBox1.Update()
            GoTo NoDownload
        End If
        Form1.TextBox1.AppendText(vbCrLf & "Downloading OpenSky database (unresponsive until complete)...")
        Form1.TextBox1.Update()
        Application.DoEvents()
        If File.Exists(dbPath & "aircraftDatabase.zip") Then
            My.Computer.FileSystem.DeleteFile(dbPath & "aircraftDatabase.zip")
        End If
        My.Computer.Network.DownloadFile("https://opensky-network.org/datasets/metadata/aircraftDatabase.zip", dbPath & "aircraftDatabase.zip")
        Form1.TextBox1.AppendText("...Done.")
        Form1.TextBox1.AppendText(vbCrLf & "Unzipping OpenSky database...")
        Form1.TextBox1.Update()
        Application.DoEvents()
        System.IO.Compression.ZipFile.ExtractToDirectory(dbPath & "aircraftDatabase.zip", dbPath, overwriteFiles:=1)
        Application.DoEvents()
        If File.Exists(dbPath & "aircraftDatabase.csv") Then Kill(dbPath & "aircraftDatabase.csv")
        My.Computer.FileSystem.MoveFile(dbPath & "media\data\samples\metadata\aircraftDatabase.csv", dbPath & "aircraftDatabase.csv")
        System.IO.Directory.Delete(dbPath & "media", True)
        If File.Exists(dbPath & "aircraftDatabase.zip") Then Kill(dbPath & "aircraftDatabase.zip")
        Form1.TextBox1.AppendText(vbCrLf & "Downloading OpenSky database...Done." & vbCrLf & "Unzipping OpenSky database...Done.")
        Form1.TextBox1.Update()
        Application.DoEvents()

        Dim Num_Registrations = File.ReadAllLines(dbPath & "aircraftDatabase.csv").Length

        Form1.TextBox1.AppendText(vbCrLf & "Creating OpenSky SQL database...")

        On Error Resume Next
        Kill(dbPath & "OpenSkyDatabase.sqb")
        On Error GoTo 0
        Threading.Thread.Sleep(1000)
        If File.Exists(dbPath & "OpenSkyDatabase.sqb") Then
            On Error Resume Next
            Kill(dbPath & "OpenSkyDatabase.sqb")
            On Error GoTo 0
            Threading.Thread.Sleep(1000)
            If File.Exists(dbPath & "OpenSkyDatabase.sqb") Then
                MessageBox.Show("I can't seem to delete the ""OpenSkyDatabase.sqb"" file. in the working folder." & vbCrLf & " Please delete this manually and run the program again." & vbCrLf & "Exiting.", "E's VRS Updater")
                Application.Exit()
                End
            End If
        End If

        Dim connection As New SQLiteConnection("DataSource=" & dbPath & "OpenSkyDatabase.sqb")
        Dim command As New SQLiteCommand("", connection)
        Dim Manufacturer As String

        connection.Open()
        If connection.State = ConnectionState.Closed Then
            MessageBox.Show("Connection To db closed still." & vbCrLf & "Exiting.", "E's VRS Updater")
            Application.Exit()
            End
        End If

        command.CommandText = "CREATE TABLE Aircraft (ICAO Char, Registration Char, Manufacturer Char, Model Char, ModelIcao Char, Operator Char, OperatorIcao Char, Serial Char, YearBuilt Char)"
        command.ExecuteNonQuery()

        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(dbPath & "aircraftDatabase.csv")
        Dim a As String, Split_String() As String
        Dim i As Long, j As Long, k As Long, m As Long, ThePercent As Integer
        Dim ToEnter As String, strSQL As String, Year_Built As String, OperatorIcao As String, Serial_Number As String
        Dim Registration As String, ICAO As String, ICAO_Type As String, Owner_Name As String, Model_String As String, Type_Kit As String
        Dim TimeEstimator1 As String, TimeEstimator2 As Double, Time_String As String

        i = 0
        k = 0
        m = 0
        j = 0
        ToEnter = vbNullString
        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        Do
            a = reader.ReadLine
            i = i + 1
            Split_String = Split(a, """,")
            If Split_String(0) <> vbNullString And i > 1 Then
                ICAO = PrepNull(Replace(Split_String(0), """", ""))
                If ICAO <> "NULL" And Split_String.Length >= 18 Then
                    ICAO = UCase(ICAO)
                    Registration = PrepNull(Replace(Replace(Split_String(1), "'", "''"), """", ""))
                    Manufacturer = PrepNull(Replace(Replace(Split_String(3), "'", "''"), """", ""))
                    Model_String = PrepNull(Replace(Replace(Split_String(4), "'", "''"), """", ""))
                    'If ICAO = "'404E83'" Or ICAO = "'405F86'" Then Stop  'I don't know why I have to do this for these models
                    'If ICAO = "'40434F'" Then Stop
                    Model_String = Replace(Model_String, "X''AIR", "XAIR")
                    Model_String = Replace(Model_String, "X''air", "Xair")
                    ICAO_Type = PrepNull(Replace(Replace(Split_String(5), "'", "''"), """", ""))
                    Serial_Number = PrepNull(Replace(Replace(Split_String(6), "'", "''"), """", ""))
                    OperatorIcao = PrepNull(Replace(Replace(Split_String(11), "'", "''"), """", ""))
                    Owner_Name = Replace(Split_String(13), """", "")
                    Year_Built = PrepNull(Replace(Replace(Split_String(18), "'", "''"), """", ""))
                    If Year_Built <> "NULL" Then
                        Year_Built = Left(Year_Built, 5) & "'"
                    End If
                    Owner_Name = Remove_Trailing_Spaces(Owner_Name)
                    Owner_Name = Remove_All_Caps(Owner_Name)
                    Owner_Name = Replace(Owner_Name, "'", "''")
                    Owner_Name = Replace(Owner_Name, " Llc", " LLC")
                    Owner_Name = Replace(Owner_Name, " Llp", " LLP")
                    Owner_Name = Replace(Owner_Name, " Of ", " of ")
                    If Owner_Name = "Private" Then
                        Owner_Name = vbNullString
                    End If
                    Owner_Name = PrepNull(Owner_Name)

                    If ICAO = "'404E83'" Or ICAO = "'405F86'" Or ICAO = "'C0890F'" Or ICAO = "'C08819'" Or ICAO = "'C062EC'" Then
                        Model_String = "'Emarit'"
                    End If
                    ToEnter = "(" & ICAO & "," & Registration & "," & Manufacturer & "," & Model_String & "," & ICAO_Type & "," & Owner_Name & "," & OperatorIcao & "," & Serial_Number & "," & Year_Built & ")," & ToEnter

                    j = j + 1
                    If j = 1000 Then    'rather than insert every time, insert 1,000 at-a-time to speed things up
                        command.CommandText = "INSERT INTO Aircraft (ICAO, Registration, Manufacturer, Model, ModelIcao, Operator, OperatorIcao, Serial, YearBuilt) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";"
                        command.ExecuteNonQuery()
                        j = 0
                        ToEnter = vbNullString
                        'Stop
                    End If

                    ThePercent = Math.Round(100 * i / Num_Registrations, 0)
                    If ThePercent - k >= 1 Then
                        TimeEstimator2 = CDbl(DateDiff("s", TimeEstimator1, DateTime.Now.ToString("HH:mm:ss")))
                        TimeEstimator2 = CDbl((TimeEstimator2 / (ThePercent / 100)) * (100 - ThePercent) / 100)
                        If TimeEstimator2 <> 0 Then
                            If TimeEstimator2 > 60 Then
                                Time_String = CStr(Math.Round(TimeEstimator2 / 60, 1))
                                Form1.TextBox8.Text = "Minutes Remain"
                            Else
                                Time_String = CStr(Math.Round(TimeEstimator2, 1))
                                Form1.TextBox8.Text = "Seconds Remain"
                            End If
                        Else
                            Time_String = vbNullString
                        End If
                        Form1.TextBox8.Update()
                        Form1.TextBox7.Text = Time_String
                        Form1.TextBox7.Update()
                        Form1.TextBox6.Text = Math.Round(ThePercent, 0) & "%"
                        Form1.TextBox6.Update()
                        Application.DoEvents()
                        k = ThePercent
                    End If
                    If Want_Exit = 1 Then
                        connection.Close()
                        End
                    End If

                End If
            End If
        Loop Until a Is Nothing

        reader.Close()
NoDownload:
    End Sub

    Function PrepNull(x As String) As String
        If x <> vbNullString Then
            PrepNull = "'" & x & "'"
        Else
            PrepNull = "NULL"
        End If
    End Function
End Module



