Imports System.IO.Enumeration
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Data.SQLite
Imports System.Data.Common
Imports System.Data.SqlTypes
' right click on project, manage NuGEt, get system.data.sqlite
' https://www.youtube.com/watch?v=1ucw719e4j8
Module Create_SQL
    Sub FAA_to_SQL()
        Dim i As Long, j As Long, k As Long, m As Long, ThePercent As Integer
        Dim ToEnter As String, strSQL As String, FileNumber As Integer, Year_Built As String
        Dim Registration As String, ICAO As String, FAA_Type As String, Owner_Name As String, Manu_Kit As String, Type_Kit As String
        Dim Registration_Data() As String, Num_Registrations As Long, Aircraft_Reference_Data() As String, Num_References As Long
        Dim TimeEstimator1 As String, TimeEstimator2 As Double, Time_String As String, WTC As String

        Form1.TextBox1.AppendText(vbCrLf & "Creating FAA SQL database..." & vbCrLf & "  'Master' table...")
        On Error Resume Next
        Kill(dbPath & "FAADatabase.sqb")
        On Error GoTo 0
        Threading.Thread.Sleep(1000)

        Dim connection As New SQLiteConnection("DataSource=" & dbPath & "FAADatabase.sqb")
        Dim command As New SQLiteCommand("", connection)

        connection.Open()
        If connection.State = ConnectionState.Closed Then MsgBox("Connection To db closed still.")
        command.CommandText = "CREATE TABLE Master (Registration Char, ICAO Char, Type Char, Owner_Name Char, Manufacturer_Kit Char, Type_Kit Char, Year Char)"
        command.ExecuteNonQuery()
        command.CommandText = "CREATE TABLE Aircraft_Reference (Type Char, Manufacturer Char, Model Char, Engine_Type Char, Species Char, Engine_Num Char, WTC Char)"
        command.ExecuteNonQuery()
        'connection.Close()

        Registration_Data = GetFAA_Data(dbPath & "fsdc3245092fm2\MASTER.txt")
        Num_Registrations = UBound(Registration_Data)

        k = 0
        m = 0
        j = 0
        ToEnter = vbNullString

        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        For i = 1 To Num_Registrations - 1
            Year_Built = Mid$(Registration_Data(i), 52, 4)
            ICAO = Mid$(Registration_Data(i), Len(Registration_Data(i)) - 10, 6)
            Registration = Remove_Trailing_Spaces(Left("N" & Registration_Data(i), 6))
            FAA_Type = Mid$(Registration_Data(i), 38, 7)

            Manu_Kit = Mid$(Registration_Data(i), 550, 30)
            Manu_Kit = Remove_Trailing_Spaces(Manu_Kit)
            If Len(Manu_Kit) = 0 Then
                Manu_Kit = "NULL"
            Else
                Manu_Kit = Remove_All_Caps(Manu_Kit)
                Manu_Kit = RemoveNonASCII(Manu_Kit)
                Manu_Kit = Replace(Manu_Kit, "'", "''")
            End If


            Type_Kit = Mid$(Registration_Data(i), 581, 20)
            Type_Kit = Remove_Trailing_Spaces(Type_Kit)
            If Len(Type_Kit) = 0 Then
                Type_Kit = "NULL"
            Else
                Type_Kit = RemoveNonASCII(Type_Kit)
                Type_Kit = Replace(Type_Kit, "'", "''")
            End If

            Owner_Name = Mid$(Registration_Data(i), 59, 50)
            Owner_Name = Remove_Trailing_Spaces(Owner_Name)
            Owner_Name = Remove_All_Caps(Owner_Name)
            Owner_Name = Replace(Owner_Name, "'", "''")
            Owner_Name = Replace(Owner_Name, " Llc", " LLC")
            Owner_Name = Replace(Owner_Name, " Llp", " LLP")
            Owner_Name = Replace(Owner_Name, " Of ", " of ")

            If Type_Kit = "NULL" And Manu_Kit = "NULL" Then
                ToEnter = "('" & ICAO & "','" & Registration & "','" & FAA_Type & "','" & Owner_Name & "'," & Manu_Kit & "," & Type_Kit & ",'" & Year_Built & "')," & ToEnter
            ElseIf Type_Kit = "NULL" Then
                ToEnter = "('" & ICAO & "','" & Registration & "','" & FAA_Type & "','" & Owner_Name & "','" & Manu_Kit & "'," & Type_Kit & ",'" & Year_Built & "')," & ToEnter
            ElseIf Manu_Kit = "NULL" Then
                ToEnter = "('" & ICAO & "','" & Registration & "','" & FAA_Type & "','" & Owner_Name & "'," & Manu_Kit & ",'" & Type_Kit & "','" & Year_Built & "')," & ToEnter
            Else
                ToEnter = "('" & ICAO & "','" & Registration & "','" & FAA_Type & "','" & Owner_Name & "','" & Manu_Kit & "','" & Type_Kit & "','" & Year_Built & "')," & ToEnter
            End If

            j = j + 1
            If j = 1000 Then    'rather than insert every time, insert 1,000 at-a-time to speed things up
                command.CommandText = "INSERT INTO Master (ICAO, Registration, Type, Owner_Name, Manufacturer_Kit, Type_Kit, Year) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";"
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
        Next
        If j <> 0 Then
            command.CommandText = "INSERT INTO Master (ICAO, Registration, Type, Owner_Name, Manufacturer_Kit, Type_Kit, Year) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";"
            command.ExecuteNonQuery()
            ToEnter = vbNullString
        End If
        Erase Registration_Data
        ICAO = vbNullString
        Registration = vbNullString
        Owner_Name = vbNullString
        Manu_Kit = vbNullString
        Type_Kit = vbNullString
        Form1.TextBox6.Text = " - %"
        Form1.TextBox6.Update()
        Form1.TextBox1.AppendText("Done.")
        Form1.TextBox1.AppendText(vbCrLf & "  'Reference' table...")
        Form1.TextBox1.Update()
        Application.DoEvents()


        Aircraft_Reference_Data = GetFAA_Data(dbPath & "fsdc3245092fm2\ACFTREF.txt")
        Num_References = UBound(Aircraft_Reference_Data)
        k = 0
        m = 0
        j = 0
        ToEnter = vbNullString
        Dim Manufacturer As String, Model_Name As String, Engine_Type As String, Species As String, Engine_Num As String

        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        For i = 1 To Num_References - 1
            Manufacturer = Mid$(Aircraft_Reference_Data(i), 9, 30)
            Manufacturer = Remove_Trailing_Spaces(Manufacturer)
            Manufacturer = Remove_All_Caps(Manufacturer)
            Manufacturer = Replace(Manufacturer, "'", "''")
            Manufacturer = Replace(Manufacturer, """", """""")

            Model_Name = Mid$(Aircraft_Reference_Data(i), 40, 20)
            Model_Name = Remove_Trailing_Spaces(Model_Name)
            Model_Name = Replace(Model_Name, "'", "''")

            Engine_Type = Mid$(Aircraft_Reference_Data(i), 63, 2)
            Engine_Type = Remove_Trailing_Spaces(Engine_Type)

            FAA_Type = Left(Aircraft_Reference_Data(i), 7)
            Species = Mid$(Aircraft_Reference_Data(i), 61, 1)
            Engine_Num = Mid$(Aircraft_Reference_Data(i), 70, 2)
            WTC = Mid$(Aircraft_Reference_Data(i), 77, 7)
            If WTC = "CLASS 1" Then
                WTC = "L"
            ElseIf WTC = "CLASS 2" Then
                WTC = "M"
            ElseIf WTC = "CLASS 3" Then
                WTC = "L"
            ElseIf WTC = "CLASS 4" Then
                WTC = "UAV"
            Else
                WTC = "?"
            End If
            If Asc(Engine_Num) = 48 Then    '0
                Engine_Num = Right$(Engine_Num, 1)
            End If
            ToEnter = "('" & FAA_Type & "','" & Manufacturer & "','" & Model_Name & "','" & Engine_Type & "','" & Species & "','" & Engine_Num & "','" & WTC & "')," & ToEnter
            j = j + 1
            If j = 1000 Then    'rather than insert every time, insert 1,000 at-a-time to speed things up
                command.CommandText = ("INSERT INTO Aircraft_Reference (Type, Manufacturer, Model, Engine_Type, Species, Engine_Num, WTC) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";")
                command.ExecuteNonQuery()
                j = 0
                ToEnter = vbNullString
            End If
            ThePercent = Math.Round(100 * i / Num_References, 0)
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
        Next
        If j <> 0 Then
            command.CommandText = ("INSERT INTO Aircraft_Reference (Type, Manufacturer, Model, Engine_Type, Species, Engine_Num, WTC) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";")
            command.ExecuteNonQuery()
            ToEnter = vbNullString
        End If
        ToEnter = vbNullString
        Form1.TextBox6.Text = " - %"
        Form1.TextBox6.Update()
        Form1.TextBox1.AppendText("Done.")
        Form1.TextBox1.Update()
        Application.DoEvents()

        Kill(dbPath & "ReleasableAircraft.zip")
        connection.Close()
    End Sub


    Function GetFAA_Data(FilePath As String) As String()

        Dim MyData As String, FileNumber As Integer

        'MyData = Space$(LOF(FileNumber))        'size of file
        MyData = My.Computer.FileSystem.ReadAllText(FilePath)

        FileSystem.FileClose(FileNumber)
        GetFAA_Data = Split(MyData, vbCrLf)
        MyData = vbNullString
    End Function


    Function Custom_Format(InputValue As Long) As String
        InputValue = CDbl(InputValue)
        If InputValue = 0 Then
            Custom_Format = 0
        Else
            Custom_Format = Format(InputValue, "#,###.##")
            If (Right(Custom_Format, 1) = ".") Then
                Custom_Format = Left(Custom_Format, Len(Custom_Format) - 1)
            End If
        End If
    End Function

    Function Remove_Trailing_Spaces(String_To_Adjust As String) As String
        Dim n As Integer
        'remove trailing spaces from a string
        If Len(String_To_Adjust) = 0 Or Len(String_To_Adjust) = 1 Then
        Else
            While n <> 1
                If Right$(String_To_Adjust, 1) = " " Then
                    String_To_Adjust = Left(String_To_Adjust, Len(String_To_Adjust) - 1)
                Else
                    n = 1
                End If
            End While
        End If
        Remove_Trailing_Spaces = String_To_Adjust
    End Function

    Function Remove_Leading_Spaces(String_To_Adjust As String) As String
        Dim n As Integer
        'remove leading spaces from a string
        If Len(String_To_Adjust) = 0 Or Len(String_To_Adjust) = 1 Then
        Else
            While n <> 1
                If Left(String_To_Adjust, 1) = " " Then
                    String_To_Adjust = Right$(String_To_Adjust, Len(String_To_Adjust) - 1)
                Else
                    n = 1
                End If
            End While
        End If
        Remove_Leading_Spaces = String_To_Adjust
    End Function

    Function Remove_All_Caps(String_To_Adjust As String) As String
        Dim n As Integer
        If Len(String_To_Adjust) = 0 Or Len(String_To_Adjust) = 1 Then
        Else
            String_To_Adjust = LCase(String_To_Adjust)
            String_To_Adjust = UCase$(Left(String_To_Adjust, 1)) & Mid$(String_To_Adjust, 2, Len(String_To_Adjust) - 1)
            For n = 2 To Len(String_To_Adjust)
                If Mid$(String_To_Adjust, n, 1) = " " Then
                    String_To_Adjust = Left(String_To_Adjust, n) & UCase$(Mid$(String_To_Adjust, n + 1, 1)) & Mid$(String_To_Adjust, n + 2, Len(String_To_Adjust) - n)
                End If
            Next
        End If
        Remove_All_Caps = String_To_Adjust
    End Function

    '-----------this will remove non ASCII characters from a string
    Function RemoveNonASCII(str As String) As String
        Dim i As Integer
        RemoveNonASCII = vbNullString
        For i = 1 To Len(str)
            If Asc(Mid$(str, i, 1)) = 38 Then                    'It's an ampersand, so replace it with "and" and skip to the next character
                RemoveNonASCII = RemoveNonASCII & "and"
            ElseIf Asc(Mid$(str, i, 1)) < 127 And Asc(Mid$(str, i, 1)) > 31 Then              'Then 'It's an ASCII character or an ampersand
                RemoveNonASCII = RemoveNonASCII & Mid$(str, i, 1) 'Append it
            End If
        Next i
    End Function

End Module

