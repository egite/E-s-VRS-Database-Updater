Imports System.IO.Enumeration
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Data.SQLite 'execute this from package manger console:  Install-Package System.Data.SqLite 
Imports System.Data.Common
Imports System.Data.SqlTypes
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Data.Entity.Migrations
Imports System.Windows.Input
Imports System.IO
Imports System.Text.Json.Nodes
Imports System.DirectoryServices.ActiveDirectory

Module VRS
    Public Type_Data() As String, FAA_Model_Data() As String, FAA_Manufacturer_Data() As String, Remap_Data() As String
    Public Found_Type As Integer, Manufacturer As String, Model As String, Emitter_Type As String
    Sub Update_VRS()
        Dim dbPath1 As String, dbPath2 As String, Num_Records1 As Long, Num_Records2 As Long, i As Long, k As Double
        Dim ToEnter As String, ThePercent As Single, TheOperator As String, UTC_Time As String, OperatorIcao As String
        Dim ModelIcao As String, Registration As String, Update_YN As Long, Temp_Split() As String, Rules_Number As Integer
        Dim Update_Step As Long, Silhouettes As Long, Year_Built As String, TheRules() As String, ChangeOp As Integer
        Dim TimeEstimator1 As String, TimeEstimator2 As Double, Time_String As String, UTC_Enter As String, Change_Flag As Integer
        Dim dbPath3 As String, Num_Records3 As Long, Manufacturer_Orig As String, m As Long, Country As String, ICAO As String
        Dim Rule(0 To 7) As String, Change(0 To 7) As String, Change_MSG As String, Split_String() As String, Rule_Flag As Integer
        Dim Not_Change_Flag As Integer, Registration_in_Operator As Integer, q As Double, Have_Rules As Integer, n As Integer

        dbPath1 = dbPath & "FAADatabase.sqb"
        dbPath2 = dbPath & "AircraftOnlineLookupCache.sqb"
        dbPath3 = dbPath & "CCARDatabase.sqb"

        If Not File.Exists(dbPath1) Then
            MessageBox.Show("Warning:  I can't find the ""FAADatabase.sqb"" file anymore." & vbCrLf & vbCrLf &
                            "Please place this file in the working folder location (shown in the program window) or re-run the program to download the FAA database (ensure that ""Always re-download FAA database"" is enabled)." & vbCrLf & vbCrLf &
                            "Exiting.", "E's VRS Updater")
            Application.Exit()
            End
        End If

        If File.Exists("C:\Users\" & UserName & "\AppData\Local\VirtualRadar\AircraftOnlineLookupCache.sqb") And BackupVRSdb = "Y" Then
            Form1.TextBox1.AppendText(vbCrLf & "Copying VirtualRadar database...")
            Form1.TextBox1.Update()
            File.Copy(VRSdbPath & "AircraftOnlineLookupCache.sqb", dbPath & "AircraftOnlineLookupCache.sqb", True)
            File.Copy(VRSdbPath & "AircraftOnlineLookupCache.sqb", dbPath & "AircraftOnlineLookupCache-" & DateTime.Now.ToString("yyMMdd-HHmm") & ".sqb", True)
            Form1.TextBox1.AppendText("Done.")
            Form1.TextBox1.Update()
        Else
            Form1.TextBox1.AppendText(vbCrLf & """Virtual Radar Server"" is not installed in given location.")
            Form1.TextBox1.Update()
            If File.Exists(dbPath & "AircraftOnlineLookupCache.sqb") = False Then
                MsgBox(vbCrLf & "I can't find VRS's ""AircraftOnlineLookupCache.sqb"" file." & vbCrLf & vbCrLf & "Exiting.", "E's VRS updater")
                Application.Exit()
                End
            End If
        End If

        If File.Exists(dbPath & "Rules.csv") Then
            Have_Rules = 1
            Form1.TextBox1.AppendText(vbCrLf & """Rules.csv"" file found.")
            Form1.TextBox1.Update()
            TheRules = GetFAA_Data(dbPath & "Rules.csv")
            Rules_Number = UBound(TheRules)
        End If

        Form1.TextBox1.AppendText(vbCrLf & "Updating VirtualRadar database (FAA)...")
        Form1.TextBox1.Update()


        If File.Exists(dbPath & "Sils.csv") Then
            i = File.ReadAllLines(dbPath & "Sils.csv").Length
            ReDim FAA_Manufacturer_Data(0 To i - 1)
            ReDim FAA_Model_Data(0 To i - 1)
            ReDim Remap_Data(0 To i - 1)
            ReDim Type_Data(0 To i - 1)
            i = 0

            Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(dbPath & "Sils.csv")
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()
                        FAA_Manufacturer_Data(i) = currentRow(0).ToString
                        FAA_Model_Data(i) = currentRow(1).ToString
                        Remap_Data(i) = currentRow(2).ToString
                        Type_Data(i) = currentRow(3).ToString
                        i = i + 1
                    Catch ex As Microsoft.VisualBasic.
                    FileIO.MalformedLineException
                        'MsgBox("Line " & ex.Message & "is not valid and will be skipped.")
                    End Try
                End While
            End Using
        Else
            Call Sils()
        End If

        Dim conn0 As New SQLiteConnection("DataSource=" & dbPath1)
        Dim command0 As New SQLiteCommand("", conn0)
        Dim rst0 As DbDataReader

        Dim conn1 As New SQLiteConnection("DataSource=" & dbPath1)
        Dim command1 As New SQLiteCommand("", conn1)
        Dim rst1 As DbDataReader

        Dim conn2 As New SQLiteConnection("DataSource=" & dbPath2)
        Dim command2 As New SQLiteCommand("", conn2)
        Dim rst2 As DbDataReader

        conn0.Open()
        conn1.Open()
        conn2.Open()

        command1.CommandText = "SELECT Count(*) FROM Master"
        Num_Records1 = command1.ExecuteScalar()
        command2.CommandText = "SELECT * FROM AircraftDetail"
        Num_Records2 = command2.ExecuteScalar()

        Update_YN = 0
        Update_Step = 1
        command1.CommandText = "SELECT * FROM Master"
        rst1 = command1.ExecuteReader()

        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        k = 0
        i = 0
        q = 0.01
        'GoTo OverHere  'skip FAA db and do Canadian instead
        While rst1.Read()
            i = i + 1
            ThePercent = 100 * i / Num_Records1
            If ThePercent - k >= q Then
                If q = 0.01 And ThePercent >= 0.5 Then
                    q = 0.1
                ElseIf q = 0.1 And ThePercent >= 1 Then
                    q = 1
                End If
                TimeEstimator2 = CDbl(DateDiff("s", TimeEstimator1, DateTime.Now.ToString("HH:mm:ss")))
                TimeEstimator2 = CDbl((TimeEstimator2 / (ThePercent / 100)) * (100 - ThePercent) / 100)
                If TimeEstimator2 <> 0 Then
                    If TimeEstimator2 > 60 Then
                        If TimeEstimator2 > 10 Then
                            Time_String = Custom_Format(CStr(Math.Round(TimeEstimator2 / 60, 0)))
                        Else
                            Time_String = CStr(Math.Round(TimeEstimator2 / 60, 1))
                        End If

                        Form1.TextBox8.Text = "Minutes Remain"
                    Else
                        Time_String = CStr(Math.Round(TimeEstimator2, 0))
                        Form1.TextBox8.Text = "Seconds Remain"
                    End If
                Else
                    Time_String = vbNullString
                End If
                Form1.TextBox8.Update()
                Form1.TextBox7.Text = Time_String
                Form1.TextBox7.Update()
                If q = 0.01 Then
                    Form1.TextBox6.Text = Math.Round(ThePercent, 2) & "%"
                ElseIf q = 0.1 Then
                    Form1.TextBox6.Text = Math.Round(ThePercent, 2) & "%"
                ElseIf q = 0.5 Then
                    Form1.TextBox6.Text = Math.Round(ThePercent, 1) & "%"
                Else
                    Form1.TextBox6.Text = Math.Round(ThePercent, 1) & "%"
                End If
                Form1.TextBox6.Update()
                Application.DoEvents()
                k = ThePercent
            End If
            command2.CommandText = "SELECT * FROM AircraftDetail WHERE Icao = '" & rst1(1) & "';"
            If command2.ExecuteScalar() <> 0 Then
                rst2 = command2.ExecuteReader()
                rst2.Read()
                TheOperator = vbNullString
                Manufacturer = vbNullString
                Registration = vbNullString
                Model = vbNullString
                Year_Built = vbNullString
                Year_Built = rst1(6)
                ICAO = rst1(1).ToString
                Registration = rst1(0).ToString
                If Year_Built = vbNullString Then
                    Year_Built = System.DBNull.Value.ToString
                End If
                'update operator
                TheOperator = Replace(rst1(3).ToString, "'", "''")
                If UCase(Right(TheOperator, 3)) = "LLC" Then TheOperator = Left(TheOperator, Len(TheOperator) - 3) & "LLC"  'this catches if we see something xxxxxxllc where x is not a space
                TheOperator = Replace(TheOperator, " Llc", " LLC")
                TheOperator = Replace(TheOperator, " Llp", " LLP")
                TheOperator = Replace(TheOperator, " Of ", " of ")
                TheOperator = Replace(TheOperator, " Rv7a ", " RV7A ")
                If Left(TheOperator, Len(rst1(0))) = rst1(0) Then
                    TheOperator = UCase(rst1(0)) & Right$(TheOperator, Len(TheOperator) - Len(rst1(0)))
                End If
                If Len(TheOperator) >= Len(Registration) And Left(TheOperator, Len(Registration)) <> Registration Then  'look for the N-number in the Operator and swap it to all uppercase
                    For n = 1 To Len(TheOperator) - Len(Registration)
                        If UCase(Mid(TheOperator, n, Len(Registration))) = Registration Then
                            Registration_in_Operator = 1
                            TheOperator = Left(TheOperator, n - 1) & Registration & Right(TheOperator, Len(TheOperator) - n - Len(Registration) + 1)
                            'ToEnter = UTC_Enter & "Operator = '" & TheOperator
                            'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                            'command2.ExecuteNonQuery()
                            ChangeOp = ChangeOp + 1
                            Form1.TextBox4.AppendText(vbCrLf & "Operator has registration:  " & TheOperator & ".")
                        End If
                    Next
                End If
                If Len(TheOperator) >= Len(Registration) - 1 And Registration_in_Operator = 0 Then
                    For n = 1 To Len(TheOperator) - Len(Registration) - 1 'look for the N-number without the "N" in the Operator and swap it to all uppercase
                        If UCase(Mid(TheOperator, n, Len(Registration) - 1)) = Right(Registration, Len(Registration) - 1) Then
                            TheOperator = Left(TheOperator, n - 1) & Right(Registration, Len(Registration) - 1) & Right(TheOperator, Len(TheOperator) - n - Len(Registration) + 1 + 1)
                            'ToEnter = UTC_Enter & "Operator = '" & TheOperator
                            'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                            'command2.ExecuteNonQuery()
                            ChangeOp = ChangeOp + 1
                            Form1.TextBox4.AppendText(vbCrLf & "Operator has registration w/o N:  " & TheOperator & ".")
                        End If
                    Next
                End If
                If Mid(TheOperator, 2, 2) = " &" And Mid(TheOperator, 4, 1) = " " And Mid(TheOperator, 6, 1) = " " Then 'if it's like "x & x ", then capitalize the second letter
                    TheOperator = Left(TheOperator, 2) & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'ToEnter = UTC_Enter & "Operator = '" & Left(TheOperator, 2) & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                    'command2.ExecuteNonQuery()
                    ChangeOp = ChangeOp + 1
                    Form1.TextBox4.AppendText(vbCrLf & "Operator format X&x:  " & TheOperator & ".")
                End If
                If Mid(TheOperator, 2, 1) = "&" And Mid(TheOperator, 4, 1) = " " Then 'if it's  like "x&x "
                    TheOperator = Left(TheOperator, 2) & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'ToEnter = UTC_Enter & "Operator = '" & Left(TheOperator, 2) & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                    'command2.ExecuteNonQuery()
                    ChangeOp = ChangeOp + 1
                    Form1.TextBox4.AppendText(vbCrLf & "Operator format X&x:  " & TheOperator & ".")
                End If
                If Mid(TheOperator, 2, 1) = " " And (Mid(TheOperator, 3, 4) = "and " Or Mid(TheOperator, 3, 4) = "And ") And Mid(TheOperator, 8, 1) = " " Then 'if it's  like "x and x ", then capitalize the second letter
                    TheOperator = UCase(Mid(TheOperator, 6, 1)) & Right(TheOperator, Len(TheOperator) - 6)
                    'ToEnter = UTC_Enter & "Operator = '" & UCase(Mid(TheOperator, 6, 1)) & Right(TheOperator, Len(TheOperator) - 6)
                    'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                    'command2.ExecuteNonQuery()
                    ChangeOp = ChangeOp + 1
                    Form1.TextBox4.AppendText(vbCrLf & "Operator format X and x:  " & TheOperator & ".")
                End If
                If Left(TheOperator, 2) = "Mc" And Len(TheOperator) > 3 And Mid(TheOperator, 3, 1) <> " " Then 'if it's like "Mcxx" then assume it needs to be "McXx"
                    TheOperator = "Mc" & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'ToEnter = UTC_Enter & "Operator = '" & "Mc" & UCase(Mid(TheOperator, 3, 1)) & Right(TheOperator, Len(TheOperator) - 3)
                    'command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                    'command2.ExecuteNonQuery()
                    ChangeOp = ChangeOp + 1
                    Form1.TextBox4.AppendText(vbCrLf & "Operator format Mcxx..." & TheOperator & ".")
                End If
                'update manufacturer
                If rst1(4).ToString <> vbNullString Then
                    Manufacturer = Replace(rst1(4).ToString, "'", "''")
                    Model = Replace(rst1(5).ToString, "'", "''")
                Else
                    command0.CommandText = "SELECT * FROM Aircraft_Reference WHERE Type = '" & rst1(2).ToString & "';"
                    Manufacturer = command0.ExecuteScalar()
                    If command0.ExecuteScalar() <> vbNullString Then
                        rst0 = command0.ExecuteReader()
                        rst0.Read()
                        Manufacturer = Replace(rst0(1).ToString, "'", "''")
                        Model = Replace(rst0(2).ToString, "'", "''")
                        rst0.Close()
                    End If
                End If
                Emitter_Type = ""
                Found_Type = 0
                ModelIcao = Determine_Silhouette(Model)
                Silhouettes = 0
                If Found_Type = 0 Then
                    ModelIcao = System.DBNull.Value.ToString
                End If
                OperatorIcao = rst2(8).ToString
                Country = rst2(3).ToString
                Model = Replace(rst2(5).ToString, "'", "''")

                'make sure the plane is ID'd as from USA
                UTC_Time = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") & ".0000000Z"
                ToEnter = "Registration = '" & Registration & "', Country = 'United States', " & "Manufacturer = '" & Manufacturer & "', Model = '" & Model & "', ModelIcao = '" & ModelIcao & "', Operator = '" & TheOperator & "', YearBuilt = '" & Year_Built & "', CreatedUtc = '" & UTC_Time & "', UpdatedUtc = '" & UTC_Time & "'"
                rst2.Close()
                'Form1.TextBox4.AppendText(vbCrLf & ToEnter)
                command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & " WHERE ICAO='" & rst1(1).ToString & "'"
                command2.ExecuteNonQuery()
                UTC_Enter = "UpdatedUtc = '" & UTC_Time & "', "
                ToEnter = vbNullString
                If ChangeOp > 0 Then
                    ToEnter = UTC_Enter & "Operator = '" & TheOperator
                    command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                    command2.ExecuteNonQuery()
                End If
                ChangeOp = 0
                'If Registration = "10-0105" Then Stop
                'If Registration = "N736QG" Then Stop
                'If TheOperator = "United States Air Force" Then Stop
                If Have_Rules = 1 Then
                    For m = 2 To Rules_Number
                        'If m >= 45 Then Stop
                        Registration_in_Operator = 0
                        Not_Change_Flag = 0
                        Change_Flag = 0
                        Rule_Flag = 0
                        Split_String = Split(TheRules(m - 1), ",")
                        Rule(0) = Split_String(1)    'ICAO
                        Rule(1) = Split_String(2)    'Registration
                        Rule(2) = Split_String(3)    'Country
                        Rule(3) = Split_String(4)    'Manufacturer
                        Rule(4) = Split_String(5)    'Model
                        Rule(5) = Split_String(6)    'ModelICAO
                        Rule(6) = Split_String(7)    'Operator
                        Rule(7) = Split_String(8)    'Operator_ICA0

                        Change(0) = Split_String(12)    'ICAO
                        Change(1) = Split_String(13)    'Registration
                        Change(2) = Split_String(14)    'Country
                        Change(3) = Split_String(15)    'Manufacturer
                        Change(4) = Split_String(16)    'Model
                        Change(5) = Split_String(17)    'ModelICAO
                        Change(6) = Split_String(18)    'Operator
                        Change(7) = Split_String(19)    'Operator_ICA0

                        If ICAO = Rule(0) And Rule(0) <> vbNullString Then Rule_Flag = 1
                        If Registration = Rule(1) And Rule(1) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If Country = Rule(2) And Rule(2) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If Manufacturer = Rule(3) And Rule(3) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If Model = Rule(4) And Rule(4) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If ModelIcao = Rule(5) And Rule(5) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If TheOperator = Rule(6) And Rule(6) <> vbNullString Then Rule_Flag = 1 + Rule_Flag
                        If OperatorIcao = Rule(7) And Rule(7) <> vbNullString Then Rule_Flag = 1 + Rule_Flag

                        For n = 0 To 7
                            If n = 0 And "!" & ICAO = Rule(0) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(0)
                            End If
                            If n = 1 And "!" & Registration = Rule(1) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(1)
                            End If
                            If n = 2 And "!" & Country = Rule(2) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(2)
                            End If
                            If n = 3 And "!" & Manufacturer = Rule(3) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(3)
                            End If
                            If n = 4 And "!" & Model = Rule(4) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(4)
                            End If
                            If n = 5 And "!" & ModelIcao = Rule(5) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(5)
                            End If
                            If n = 6 And "!" & TheOperator = Rule(6) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(6)
                            End If
                            If n = 7 And "!" & OperatorIcao = Rule(7) Then
                                Not_Change_Flag = Not_Change_Flag + 1
                                Change_MSG = Rule(7)
                            End If
                            If Rule(n) <> vbNullString And Left(Rule(n), 1) <> "!" Then
                                Change_Flag = Change_Flag + 1
                            End If
                        Next

                        If Rule_Flag = Change_Flag And Rule_Flag <> 0 And Not_Change_Flag = 0 Then
                            ToEnter = vbNullString
                            If Change(0) <> vbNullString Then ToEnter = "Icao = '" & Change(0) & "', "
                            If Change(1) <> vbNullString Then ToEnter = ToEnter & " Registration = '" & Change(1) & "', "
                            If Change(2) <> vbNullString Then ToEnter = ToEnter & " Country = '" & Change(2) & "', "
                            If Change(3) <> vbNullString Then ToEnter = ToEnter & " Manufacturer = '" & Change(3) & "', "
                            If Change(4) <> vbNullString Then ToEnter = ToEnter & " Model = '" & Change(4) & "', "
                            If Change(5) <> vbNullString Then ToEnter = ToEnter & " ModelIcao = '" & Change(5) & "', "
                            If Change(6) <> vbNullString Then ToEnter = ToEnter & " Operator = '" & Change(6) & "', "
                            If Change(7) <> vbNullString Then ToEnter = ToEnter & " OperatorIcao = '" & Change(7)
                            If Right(ToEnter, 3) = "', " Then ToEnter = Left(ToEnter, Len(ToEnter) - 3)
                            If ToEnter <> vbNullString Then
                                ToEnter = UTC_Enter & ToEnter
                                command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & "' WHERE Icao='" & rst1(1).ToString & "'"
                                command2.ExecuteNonQuery()
                                Change_MSG = Split_String(21)
                                If Split_String(20) = "ICAO" Then Change_MSG = Change_MSG & ICAO & "."
                                If Split_String(20) = "Registration" Then Change_MSG = Change_MSG & Registration & "."
                                If Split_String(20) = "Country" Then Change_MSG = Change_MSG & Country & "."
                                If Split_String(20) = "Manufacturer" Then Change_MSG = Change_MSG & Manufacturer & "."
                                If Split_String(20) = "Model" Then Change_MSG = Change_MSG & Model & "."
                                If Split_String(20) = "ModelIcao" Then Change_MSG = Change_MSG & ModelIcao & "."
                                If Split_String(20) = "Operator" Then Change_MSG = Change_MSG & TheOperator & "."
                                If Split_String(20) = "Operator ICAO" Then Change_MSG = Change_MSG & OperatorIcao & "."
                                If Change_MSG <> vbNullString Then
                                    Form1.TextBox4.AppendText(vbCrLf & Change_MSG & " " & Registration & " -- From rules.")
                                    Form1.TextBox4.Update()
                                End If
                            End If
                        ElseIf Not_Change_Flag <> 0 Then
                            'Form1.TextBox4.AppendText(vbCrLf & "Not updating " & Change_MSG & "  --  From rules.")
                            'Form1.TextBox6.Update()
                        End If
                    Next
                End If
            End If
            If Want_Exit = 1 Then
                rst1.Close()
                conn0.Close()
                conn1.Close()
                conn2.Close()
                End
            End If
        End While
OverHere:
        rst1.Close()
        conn0.Close()
        conn1.Close()

        If Dir(dbPath3) = vbNullString Then
            Form1.TextBox1.AppendText(vbCrLf & "Updating VirtualRadar database...CCAR skipped.")
            Form1.TextBox1.Update()
            GoTo No_CCAR
        End If

        'now let's setup to do the CCAR database
        Dim OutputFile As System.IO.StreamWriter
        If NoSilsFound = 1 Then
            If System.IO.File.Exists(dbPath & "No Silhouette Found.txt") = True Then
                System.IO.File.Delete(dbPath & "No Silhouette Found.txt")
                Threading.Thread.Sleep(500)
            End If
            OutputFile = My.Computer.FileSystem.OpenTextFileWriter(dbPath & "No Silhouette Found.txt", True)
        End If
        Dim conn3 As New SQLiteConnection("DataSource=" & dbPath3)
        Dim command3 As New SQLiteCommand("", conn3)
        Dim rst3 As DbDataReader
        conn3.Open()
        command3.CommandText = "SELECT Count(*) FROM Master"
        Num_Records3 = command3.ExecuteScalar()

        Form1.TextBox1.AppendText(vbCrLf & "Updating VirtualRadar database (CCAR)...")
        Form1.TextBox1.Update()
        Update_YN = 0
        Update_Step = 1
        command3.CommandText = "SELECT * FROM Master"
        rst3 = command3.ExecuteReader()

        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        k = 0
        i = 0

        While rst3.Read()
            i = i + 1
            ThePercent = 100 * i / Num_Records3
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
            command2.CommandText = "SELECT * FROM AircraftDetail WHERE Icao = '" & rst3(1).ToString & "';"
            If command2.ExecuteScalar() <> 0 Then
                rst2 = command2.ExecuteReader()
                rst2.Read()
                Registration = rst3(0).ToString

                TheOperator = Replace(rst3(3).ToString, "'", "''")
                Manufacturer = Replace(rst3(4), "'", "''")
                Model = Replace(rst3(2).ToString, "'", "''")
                'Year_Built = rst3(6).ToString
                'update year
                m = 0
                If rst3(6).ToString = vbNullString Then
                    If rst2(10).ToString <> vbNullString Then
                        Year_Built = rst2(10).ToString
                        m = 1
                    End If
                ElseIf rst3(6) = System.DbNull.Value.ToString Then
                    If rst2(10).ToString <> vbNullString Then
                        Year_Built = rst2(10).ToString
                        m = 1
                    End If
                Else
                    Year_Built = rst3(6).ToString
                    m = 1
                End If
                'update operator
                TheOperator = Replace(rst3(3), "'", "''")
                'If UCase(Right(TheOperator)) = "LLC" Then TheOperator = Left(TheOperator, Len(TheOperator) - 3) & "LLC"  'this catches if we see something xxxxxxllc where x is not a space
                TheOperator = Replace(TheOperator, " Llc", " LLC")
                TheOperator = Replace(TheOperator, " Llp", " LLP")
                TheOperator = Replace(TheOperator, " Of ", " of ")
                If Left(TheOperator, Len(rst3(0))) = rst3(0).ToString Then
                    TheOperator = UCase(rst3(0).ToString) & Right$(TheOperator, Len(TheOperator) - Len(rst3(0).ToString))
                End If
                'update manufacturer
                m = 0
                If rst2(4).ToString = vbNullString Then
                    If rst3(4).ToString <> vbNullString Then
                        m = 1
                    End If
                ElseIf rst2(4) = System.DbNull.Value.tostring Then
                    If rst3(4).ToString <> vbNullString Then
                        m = 1
                    End If
                End If
                If m = 1 Then
                    Manufacturer = Replace(rst3(4), "'", "''")
                End If
                'update model
                m = 0
                If rst2(6).ToString = vbNullString Then
                    If rst3(2).ToString <> vbNullString Then
                        m = 1
                    End If
                ElseIf rst2(6) = System.DbNull.Value.ToString Then
                    If rst3(2).ToString <> vbNullString Then
                        m = 1
                    End If
                End If
                If m = 1 Then
                    Model = rst3(2)
                    Model = Replace(Model, "'", "''")
                End If
                'update silhouette
                m = 0
                Manufacturer_Orig = Manufacturer
                'this is not a complete remapping
                If Left(Manufacturer, 6) = "Cessna" Then
                    Manufacturer = "Cessna"
                ElseIf Left(Manufacturer, 6) = "Boeing" Or Left(Manufacturer, 10) = "The Boeing" Then
                    Manufacturer = "Boeing"
                ElseIf Left(Manufacturer, 6) = "Piper" Then
                    Manufacturer = "Piper"
                ElseIf Manufacturer = "Texas Engineering and Manufacturing Co. Inc." Then
                    Manufacturer = "Temco"
                ElseIf Left(Manufacturer, 8) = "Champion" Then
                    Manufacturer = "Champion"
                ElseIf Left(Manufacturer, 9) = "Schweizer" Then
                    Manufacturer = "Schweizer"
                ElseIf Left(Manufacturer, 7) = "Aeronca" Then
                    Manufacturer = "Aeronca"
                ElseIf Left(Manufacturer, 14) = "Aero Commander" Then
                    Manufacturer = "Aero Commander"
                ElseIf Left(Manufacturer, 5) = "Beech" Then
                    Manufacturer = "Beech"
                ElseIf Left(Manufacturer, 6) = "Airbus" Then
                    Manufacturer = "Airbus"
                ElseIf Left(Manufacturer, 4) = "Bell" Then
                    Manufacturer = "Bell"
                End If
                If rst2(6).ToString = vbNullString Or rst2(6).ToString = System.DBNull.Value.ToString Or rst2(6).ToString = "NULL" Then
                    m = 1
                End If
                If m = 1 Then
                    Emitter_Type = vbNullString
                    Found_Type = 0
                    Temp_Split = Split(Manufacturer, " ")
                    Manufacturer = Temp_Split(0)                'Canada uses a much more verbose description of the manufacturer than the FAA, so just use the first word to help me find the ICAO model code
                    ModelIcao = Determine_Silhouette(Model)
                    Manufacturer = Manufacturer_Orig
                    Silhouettes = 0
                    If Found_Type = 0 Then
                        ModelIcao = System.DBNull.Value.ToString
                        'OutputFile.WriteLine(Registration & "," & Manufacturer & "," & Model)
                    End If
                Else
                    ModelIcao = rst2(6).ToString
                End If
                OperatorIcao = rst2(8).ToString
                UTC_Time = DateTime.UtcNow.ToString("yyyy-MM-dd HH:mm:ss") & ".0000000Z"
                ToEnter = "Registration = '" & Registration & "', Country = 'Canada', " & "Manufacturer = '" & Manufacturer & "', Model = '" & Model & "', ModelIcao = '" & ModelIcao & "', Operator = '" & TheOperator & "', YearBuilt = '" & Year_Built & "', CreatedUtc = '" & UTC_Time & "', UpdatedUtc = '" & UTC_Time & "'"
                rst2.Close()
                command2.CommandText = "UPDATE AircraftDetail SET " & ToEnter & " WHERE ICAO='" & rst3(1).ToString & "'"
                command2.ExecuteNonQuery()
                If Want_Exit = 1 Then
                    rst1.Close()
                    conn0.Close()
                    conn1.Close()
                    conn2.Close()
                    conn3.Close()
                    End
                End If
            End If
        End While
        conn3.Close()
        If NoSilsFound = 1 Then
            If OutputFile.BaseStream.CanWrite = True Then
                OutputFile.Close()
            End If
        End If
No_CCAR:
        rst1.Close()
        conn0.Close()
        conn1.Close()
        conn2.Close()
    End Sub
End Module

