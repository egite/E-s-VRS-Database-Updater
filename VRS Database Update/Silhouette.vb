Module Silhouette
    Function Determine_Silhouette(ICAO_Type_Text_Input As String)
        Dim ICAO_Type_Text As String, ICAO_Type_Text_Orig As String

        ICAO_Type_Text_Orig = ICAO_Type_Text_Input
        ICAO_Type_Text = ICAO_Type_Text_Input

        If ICAO_Type_Text = vbNullString Then
            Dim Emitter_Text() As String
            Emitter_Text = Split(",Light,Small,Large,High vortex,Heavy,Highly maneuverable,Rotorcaft,Unassigned," &
                             "Glider/Sailplane,Lighter than air,Parachutist/Sky diver,Ultralight/hang glider/paraglider,Unassigned,UAV,Space/Transatmospheric,Unassigned,Surface vehicle - emergency," &
                             "Surface - service,Point obstacle,Cluster obstacle,Line obstacle,,,,,,,,,,,,,,,,,,,,,,,,,,,", ",")
            If Emitter_Type = "Surface vehicle - emergency" Then
                ICAO_Type_Text = "FIRE"
                Found_Type = 1
            ElseIf Emitter_Type = "Surface - service" Then
                ICAO_Type_Text = "TUG"
                Found_Type = 1
            End If
            If Found_Type = 0 Then Exit Function
        ElseIf ICAO_Type_Text = "B--B" Or ICAO_Type_Text = "----" Then
            Exit Function
        End If

        Dim First_Two As String, First_Three As String, First_Four As String, First_Five As String, First_Six As String
        Dim i As Long, j As Integer, k As Integer
        Dim Split_String() As String, Split_String2() As String

        Manufacturer = Replace(Manufacturer, ",", vbNullString)

        ICAO_Type_Text = Replace(ICAO_Type_Text, "-", vbNullString)
        ICAO_Type_Text = Replace(ICAO_Type_Text, " ", vbNullString)

        First_Two = Left(ICAO_Type_Text, 2)
        First_Three = Left(ICAO_Type_Text, 3)
        First_Four = Left(ICAO_Type_Text, 4)
        First_Five = Left(ICAO_Type_Text, 5)
        First_Six = Left(ICAO_Type_Text, 6)

        If Manufacturer = "Airbus Industrie" Or Left(Manufacturer, 6) = "Airbus" Or Manufacturer = "Airbus Sas" Then
            If First_Four = "A330" Or First_Four = "A350" Then
                ICAO_Type_Text_Orig = First_Three & Right$(First_Five, 1)
            ElseIf First_Four = "A319" Or First_Four = "A318" Or First_Four = "A320" Or First_Four = "A321" Then
                ICAO_Type_Text_Orig = First_Four
            End If
        ElseIf Left(Manufacturer, 6) = "Boeing" Then
            If ICAO_Type_Text = "78710" Then
                ICAO_Type_Text = "B789"
                Found_Type = 1
            ElseIf ICAO_Type_Text_Orig = "777F" Or First_Four = "777F" Then
                ICAO_Type_Text = "B77L"
                Found_Type = 1
            ElseIf First_Two = "MD" Then
                Manufacturer = "Mcdonnell Douglas"
            ElseIf ICAO_Type_Text = "E3TF" Then
                ICAO_Type_Text = "E3TF"
                Found_Type = 1
            ElseIf ICAO_Type_Text = "E6" Then
                Found_Type = 1
            ElseIf ICAO_Type_Text = "A75N1(PT17)" Or ICAO_Type_Text = "A75" Then
                ICAO_Type_Text = "ST75"
                Found_Type = 1
            ElseIf ICAO_Type_Text = "K35R" Or ICAO_Type_Text = "C135" Then
                Found_Type = 1
            ElseIf Left(ICAO_Type_Text, 1) <> "B" Then
                ICAO_Type_Text = "B" & First_Two & Right$(First_Four, 1)
                'Found_Type = 1
            End If
        ElseIf Left(Manufacturer, 6) = "Cessna" Or Manufacturer = "Textron Aviation Inc" Or Manufacturer = "Textron Aviation Inc." Then
            Manufacturer = "Cessna"
            If First_Four = "T182" Then
                ICAO_Type_Text = "C182"
                Found_Type = 1
            ElseIf First_Four = "T206" Then
                ICAO_Type_Text = "C210"
                Found_Type = 1
            ElseIf First_Four = "T210" Then
                ICAO_Type_Text = "C210"
                Found_Type = 1
            ElseIf ICAO_Type_Text = "TU206A" Then
                ICAO_Type_Text = "C206"
                Found_Type = 1
            ElseIf ICAO_Type_Text = "530" Then
                ICAO_Type_Text = "E530"
                Found_Type = 0
            ElseIf ICAO_Type_Text = "P172D" Then
                ICAO_Type_Text = "C172"
                Found_Type = 1
            ElseIf ICAO_Type_Text = "G36" Then
                Found_Type = 0
            Else
                ICAO_Type_Text = "C" & First_Three
            End If
        ElseIf Manufacturer = "Eiriavion Oy" Then
            ICAO_Type_Text = "AS25"
        ElseIf Manufacturer = "Glasflugel" Then
            ICAO_Type_Text = "AS25"
            Found_Type = 1
        ElseIf left(Manufacturer, 7) = "Pilatus" Or Manufacturer = "Pilatus Aircraft Ltd" Then
            Manufacturer = "Pilatus"
            Split_String = Split(ICAO_Type_Text_Orig, "/")
            ICAO_Type_Text_Orig = Split_String(0)
        ElseIf Manufacturer = "Piper" Then
            ICAO_Type_Text = First_Four
        ElseIf Manufacturer = "Mooney Aircraft Corp." Then
            ICAO_Type_Text = Replace(ICAO_Type_Text, "-", vbNullString)
            Split_String = Split(ICAO_Type_Text, " ")
            ICAO_Type_Text = Split_String(0)
        ElseIf ICAO_Type_Text_Orig = "BALL" Then
            ICAO_Type_Text = "Ball"
            Found_Type = 1
        ElseIf Manufacturer = vbNullString Then
            If ICAO_Type_Text = "BE3D" Then
                ICAO_Type_Text = "BE33"
                Found_Type = 1
            End If
        End If

        If Found_Type = 0 Then
            For i = 1 To UBound(Type_Data, 1)
                If FAA_Manufacturer_Data(i) <> vbNullString Then
                    Split_String = Split(FAA_Manufacturer_Data(i), ",")
                    For k = 0 To UBound(Split_String)
                        If Manufacturer = Split_String(k) Or Split_String(k) = "*" Then
                            If FAA_Model_Data(i) <> vbNullString Then
                                Split_String2 = Split(CStr(FAA_Model_Data(i)), ",")
                                For j = 0 To UBound(Split_String2)
                                    If ICAO_Type_Text_Orig = Split_String2(j) Or Split_String2(j) = "*" Then
                                        ICAO_Type_Text = Remap_Data(i)
                                        If ICAO_Type_Text = vbNullString Then
                                            ICAO_Type_Text = Type_Data(i)
                                        End If
                                        Found_Type = 1
                                        GoTo Got_ICAO
                                    End If
                                Next
                            End If
                        End If
                    Next
                Else
                    Found_Type = 0
                End If
            Next
        End If

Got_ICAO:

        If Manufacturer = "Boeing" Then
            Found_Type = 1
        End If

        Determine_Silhouette = ICAO_Type_Text

        Erase Split_String
        Erase Split_String2
        Emitter_Type = "Nada"      'indicates that we haven't defined the emitter yet for silhouette determination
    End Function

    Function Get_Data_From_File(FilePath As String) As String()

        Dim MyData As String, FileNumber As Integer
        MyData = My.Computer.FileSystem.ReadAllText(FilePath)

        FileSystem.FileClose(FileNumber)
        Get_Data_From_File = Split(MyData, vbCrLf)
        MyData = vbNullString
    End Function

End Module