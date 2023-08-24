Imports System.ComponentModel
Imports System.Data.SQLite
Imports System.IO

Module Creat_SQL_CCAR
    Sub CCAR_to_SQL()
        Dim i As Long, j As Long, k As Long, m As Long, ThePercent As Integer, Num_Owners As Long, Registration_Orig As String
        Dim ToEnter As String, strSQL As String, FileNumber As Integer, Year_Built As String, Have_Owner As Long
        Dim Registration As String, ICAO As String, CCAR_Type As String, Owner_Name As String, Serial_Number As String
        Dim Registration_Data() As String, Num_Registrations As Long, Owner_Data() As String, Num_References As Long
        Dim TimeEstimator1 As String, TimeEstimator2 As Double, Time_String As String, Split_String() As String


        If File.Exists(dbPath & "ccarcsdb.zip") = False Then
            Form1.TextBox1.AppendText(vbCrLf & "Skipping CCAR SQL database creation.")
            Form1.TextBox1.Update()
            MsgBox("I don't see the CCAR database." & vbCrLf & "You can download the CCAR database from here:" & vbCrLf & vbCrLf & "https://wwwapps.tc.gc.ca/saf-sec-sur/2/ccarcs-riacc/DDZip.aspx" & vbCrLf & vbCrLf & "Place the zipfile in the same folder as this program.")
            GoTo No_CCAR_SQL
        End If

        If Directory.Exists(dbPath & "fsdc3245092fm2") Then
            System.IO.Directory.Delete(dbPath & "fsdc3245092fm2", True)
        End If
        System.IO.Compression.ZipFile.ExtractToDirectory(dbPath & "ccarcsdb.zip", dbPath & "fsdc3245092fm2", overwriteFiles:=1)
        'Kill(dbPath & "ccarcsdb.zip")

        Registration_Data = GetFAA_Data(dbPath & "fsdc3245092fm2\carscurr.txt")
        Num_Registrations = UBound(Registration_Data)
        Owner_Data = GetFAA_Data(dbPath & "fsdc3245092fm2\carsownr.txt")
        Num_Owners = UBound(Owner_Data)

        System.IO.Directory.Delete(dbPath & "fsdc3245092fm2", True)
        If FAA_db_Download = "Y" Then
            Form1.TextBox1.AppendText(vbCrLf & "Creating CCAR SQL database...")
        Else
            Form1.TextBox1.AppendText("Creating CCAR SQL database...")
        End If
        Form1.TextBox1.Update()
        Application.DoEvents()
        On Error Resume Next
        Kill(dbPath & "CCARDatabase.sqb")
        On Error GoTo 0
        Threading.Thread.Sleep(1000)

        Dim connection As New SQLiteConnection("DataSource=" & dbPath & "CCARDatabase.sqb")
        Dim command As New SQLiteCommand("", connection)

        connection.Open()
        If connection.State = ConnectionState.Closed Then MsgBox("Connection To db closed still.")
        command.CommandText = "CREATE TABLE Master (Registration Char, ICAO Char, CCAR_Type Char, Owner_Name Char, Manufacturer Char, Serial Char, Year Char)"
        command.ExecuteNonQuery()
        'connection.Close()


        k = 0
        m = 0
        j = 0
        ToEnter = vbNullString

        TimeEstimator1 = DateTime.Now.ToString("HH:mm:ss")
        For i = 0 To Num_Registrations - 3
            Split_String = Split(Registration_Data(i), """,""")     'the indices here are 2 less than the indices in the "carslayout.txt" file
            Registration_Orig = Split_String(0)
            Registration = Strings.Replace(Registration_Orig, """", vbNullString)
            If Left(Registration, 1) = " " Then                     'if the left character is a space, it's a "CF-" airplane, otherwise its a "C-" 
                Registration = "CF-" & Strings.Replace(Registration, " ", vbNullString)
            Else
                Registration = "C-" & Registration
            End If
            ICAO = Split_String(42)
            ICAO = Bin2Hex(Left(ICAO, 4)) & Bin2Hex(Mid(ICAO, 5, 4)) & Bin2Hex(Mid(ICAO, 9, 4)) & Bin2Hex(Mid(ICAO, 13, 4)) & Bin2Hex(Mid(ICAO, 17, 4)) & Bin2Hex(Mid(ICAO, 21, 4))
            CCAR_Type = Strings.Replace(Split_String(4), """", vbNullString)
            CCAR_Type = Replace(CCAR_Type, "'", "''")
            Manufacturer = Strings.Replace(Split_String(7), """", vbNullString)
            Manufacturer = Replace(Manufacturer, "'", "''")
            Manufacturer = Globalization.CultureInfo.CurrentCulture.TextInfo.ToTitleCase(Manufacturer.ToLower)  'capitlizes only the first letter of each word (at least for my locale)
            Manufacturer = Replace(Manufacturer, " Of ", " of ")
            Manufacturer = Replace(Manufacturer, " And ", " and ")
            Serial_Number = Strings.Replace(Split_String(5), """", vbNullString)
            Serial_Number = Replace(Serial_Number, "'", "''")
            Year_Built = Left(Split_String(31), 4)

            Have_Owner = 0              'looks like the owner table is longer than the registered table (maybe because they keep a list of owners for out-of-date registrations)
            While Have_Owner = 0
                Split_String = Split(Owner_Data(m), """,""")
                If Registration_Orig = Split_String(0) Then
                    Owner_Name = Split_String(1)
                    Owner_Name = Remove_Trailing_Spaces(Owner_Name)
                    Owner_Name = Replace(Owner_Name, "'", "''")
                    Owner_Name = Replace(Owner_Name, " Llc", " LLC")
                    Owner_Name = Replace(Owner_Name, " Llp", " LLP")
                    Owner_Name = Replace(Owner_Name, " Of ", " of ")
                    Have_Owner = 1
                End If
                m = m + 1
                If m = Num_Owners Then Stop
            End While

            If Owner_Name = vbNullString Then Stop

            ToEnter = "('" & ICAO & "','" & Registration & "','" & CCAR_Type & "','" & Owner_Name & "','" & Manufacturer & "','" & Serial_Number & "','" & Year_Built & "')," & ToEnter
            Owner_Name = vbNullString
            j = j + 1
            If j = 1000 Then    'rather than insert every time, insert 1,000 at-a-time to speed things up
                command.CommandText = "INSERT INTO Master (ICAO, Registration, CCAR_Type, Owner_Name, Manufacturer, Serial, Year) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";"
                command.ExecuteNonQuery()
                j = 0
                ToEnter = vbNullString
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
            command.CommandText = "INSERT INTO Master (ICAO, Registration, CCAR_Type, Owner_Name, Manufacturer, Serial, Year) VALUES " & Mid$(ToEnter, 1, Len(ToEnter) - 1) & ";"
            command.ExecuteNonQuery()
            ToEnter = vbNullString
        End If
        Erase Registration_Data
        Erase Owner_Data
        ICAO = vbNullString
        Registration = vbNullString
        Owner_Name = vbNullString
        ToEnter = vbNullString
        Form1.TextBox6.Text = " - %"
        Form1.TextBox6.Update()
        Form1.TextBox1.AppendText("Done.")
        Form1.TextBox1.Update()
        Application.DoEvents()

        connection.Close()
        'Kill(dbPath & "ccarcsdb.zip")
No_CCAR_SQL:
    End Sub
    Function Bin2Hex(ICAO As String) As String      'you'd think vb.net would have a function for this
        If ICAO = "0000" Then
            Bin2Hex = "0"
        ElseIf ICAO = "0001" Then
            Bin2Hex = "1"
        ElseIf ICAO = "0010" Then
            Bin2Hex = "2"
        ElseIf ICAO = "0011" Then
            Bin2Hex = "3"
        ElseIf ICAO = "0100" Then
            Bin2Hex = "4"
        ElseIf ICAO = "0101" Then
            Bin2Hex = "5"
        ElseIf ICAO = "0110" Then
            Bin2Hex = "6"
        ElseIf ICAO = "0111" Then
            Bin2Hex = "7"
        ElseIf ICAO = "1000" Then
            Bin2Hex = "8"
        ElseIf ICAO = "1001" Then
            Bin2Hex = "9"
        ElseIf ICAO = "1010" Then
            Bin2Hex = "A"
        ElseIf ICAO = "1011" Then
            Bin2Hex = "B"
        ElseIf ICAO = "1100" Then
            Bin2Hex = "C"
        ElseIf ICAO = "1101" Then
            Bin2Hex = "D"
        ElseIf ICAO = "1110" Then
            Bin2Hex = "E"
        Else
            Bin2Hex = "F"
        End If
    End Function
End Module