Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button4"
                Dim RVLID As Long = 0
                Dim Found As Boolean = False
                Try
                    RVLID = InputBox("Please input RVL ID", "RVL ID", "123456")
                Catch ex As Exception
                    MsgBox("Please input a correct RVL ID")
                End Try

                For Each row As DataGridViewRow In Form1.DataGridView1.Rows
                    Dim TestRVL As Long = CLng(row.Cells("SubjectID").Value.ToString)
                    If RVLID = TestRVL Then
                        row.Cells("SubjectID").Selected = True
                        Form1.DataGridView1.FirstDisplayedScrollingRowIndex = row.Index
                        Found = True
                    End If
                Next

                If Found = False Then MsgBox("RVL ID not found. Please ensure 'All' records are searched")

            Case "Button1"
                Call Saver(Form1.DataGridView1)

            Case "Button2"
                Call UpdateFromMeddbase()
                Call Form1.Specifics(Form1.DataGridView1)
                MsgBox("Update Complete")

            Case "Button3"
                Dim LocationAtSite As String = ""
                Dim SiteAt As String = ""
                Dim FromString As String = ""
                Dim ToString As String = ""

                Try
                    LocationAtSite = Form1.ComboBox4.SelectedValue
                Catch ex As Exception
                End Try
                Try
                    SiteAt = Form1.ComboBox1.SelectedValue
                Catch ex As Exception
                End Try
                Try
                    FromString = Form1.ComboBox3.SelectedValue
                Catch ex As Exception
                End Try
                Try
                    ToString = Form1.ComboBox2.SelectedValue
                Catch ex As Exception
                End Try

                If LocationAtSite <> "" And LocationAtSite <> "In Transit" And SiteAt = "" Then
                    MsgBox("A Site must be chosen if a location is selected")
                    Exit Sub
                End If

                If LocationAtSite = "In Transit" And (FromString = "" Or ToString = "") Then
                    MsgBox("A 'From' and 'To' Site must be selected")
                    Exit Sub
                End If

                If LocationAtSite <> "In Transit" And (FromString <> "" Or ToString <> "") Then
                    MsgBox("'From' and 'To' can only be selected with location 'In Transit'")
                    Exit Sub
                End If

                If SiteAt <> "" And LocationAtSite = "" Then
                    MsgBox("A Location must be selected if a site has been chosen")
                    Exit Sub
                End If

                If SiteAt <> "" And LocationAtSite = "In Transit" Then
                    MsgBox("A site must not be selected if location is 'In Transit'")
                    Exit Sub
                End If

                If ToString = FromString And ToString <> "" And FromString <> "" Then
                    MsgBox("Sites cannot be the same")
                    Exit Sub
                End If

                If LocationAtSite = "" Then
                    MsgBox("A Location must be chosen")
                    Exit Sub
                End If

                For Each txtbox In Form1.SplitContainer2.Panel2.Controls

                    If TypeOf txtbox IsNot TextBox Then Continue For

                    Dim RVLID As Long = 0

                    Try
                        If txtbox.Text.ToString = "" Then Continue For
                        RVLID = CLng(txtbox.Text.ToString)
                    Catch ex As Exception
                        MsgBox("Volunteer ID does not appear to be a number")
                        Exit Sub
                    End Try

                    If RVLID = 0 Then Continue For

                    Dim CheckString As String = "SELECT 1 FROM [Notes] WHERE SubjectID=" & RVLID
                    If OverClass.SELECTCount(CheckString) = 0 Then
                        Call UpdateFromMeddbase()
                        If OverClass.SELECTCount(CheckString) = 0 Then
                            Call UpdateFromMeddbase()
                            If OverClass.SELECTCount(CheckString) = 0 Then
                                MsgBox("Error - Volunteer " & RVLID & " missing from Meddbase Link.")
                                Exit Sub
                            End If
                        End If
                    End If

                Next

                LocationAtSite = "'" & LocationAtSite & "'"
                SiteAt = "'" & SiteAt & "'"
                ToString = "'" & ToString & "'"
                FromString = "'" & FromString & "'"

                If SiteAt = "''" Then SiteAt = "Null"
                If ToString = "''" Then ToString = "Null"
                If FromString = "''" Then FromString = "Null"

                OverClass.CmdList.Clear()

                For Each txtbox In Form1.SplitContainer2.Panel2.Controls

                    If TypeOf txtbox IsNot TextBox Then Continue For

                    Dim RVLID As Long = 0

                    Try
                        If txtbox.Text.ToString = "" Then Continue For
                        RVLID = CLng(txtbox.Text.ToString)
                    Catch ex As Exception
                        MsgBox("Volunteer ID does not appear to be a number")
                        Exit Sub
                    End Try

                    Dim UpdateString As String = "UPDATE [NOTES] SET [SiteAt]=" & SiteAt &
                        ", [LocationAtSite]=" & LocationAtSite &
                        ", [FROM]=" & FromString &
                        ", [To]=" & ToString &
                        ", [Timestamp]=" & OverClass.SQLDate(DateAndTime.Now) &
                        ", [Person]='" & OverClass.GetUserName & "'" &
                        " WHERE [SubjectID]=" & RVLID


                    OverClass.AddToMassSQL(UpdateString)

                Next

                OverClass.ExecuteMassSQL()

                For Each txtbox In Form1.SplitContainer2.Panel2.Controls
                    If TypeOf txtbox IsNot TextBox Then Continue For
                    txtbox.Text = ""
                Next

                MsgBox("Update Complete")

        End Select



    End Sub


End Module
