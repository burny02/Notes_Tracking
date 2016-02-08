Option Explicit On

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Maximized

        Call StartUp(Me)

        Try
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside" & vbNewLine & "Version: " & System.Deployment.Application.ApplicationDeployment.CurrentDeployment.CurrentVersion.ToString
        Catch
            Me.Label2.Text = SolutionName & vbNewLine & "Developed by David Burnside"
        End Try

        Me.Text = SolutionName

    End Sub

    Private Sub TabControl1_Selecting(sender As Object, e As TabControlCancelEventArgs) Handles TabControl1.Selecting

        Dim SQLCode As String = vbNullString
        Dim Bind As BindingSource = BindingSource1

        If OverClass.UnloadData() = True Then
            e.Cancel = True
            Exit Sub
        End If

        OverClass.ResetCollection()

        Select Case e.TabPage.Text

            Case "Spreadsheet"

                Call Specifics(DataGridView1)

            Case "Bulk Update"

                Call Specifics(ComboBox1)

        End Select

    End Sub

    Public Sub Specifics(ctl As Object)

        If IsNothing(ctl) Then Exit Sub

        OverClass.ResetCollection()

        Select Case ctl.name

            Case "DataGridView1"

                Dim TopString As String = vbNullString

                If CheckBox1.Checked = True Then
                    TopString = "TOP 10 PERCENT "
                ElseIf CheckBox2.Checked = True Then
                    TopString = "TOP 50 PERCENT "
                ElseIf CheckBox3.Checked = True Then
                    TopString = vbNullString
                End If

                ctl.AutoGenerateColumns = True
                OverClass.CreateDataSet("SELECT " & TopString & "* FROM Notes ORDER BY SubjectID DESC", BindingSource1, ctl)
                OverClass.SetupFilterCombo(ComboBox5, "SiteAt", "SiteAt")
                OverClass.SetupFilterCombo(ComboBox6, "LocationAtSite", "LocationAtSite")

                Dim dt As DataTable = OverClass.TempDataTable("SELECT DISTINCT '' AS Site FROM SITE " &
                                                              "UNION ALL SELECT Site FROM Site ORDER BY Site ASC")

                Dim clm1 As New DataGridViewComboBoxColumn()
                clm1.DataSource = dt
                clm1.DataPropertyName = "PSPSite"
                clm1.ValueMember = "Site"
                clm1.DisplayMember = "Site"
                clm1.HeaderText = "PSP Site"
                ctl.Columns.Add(clm1)

                Dim clm2 As New DataGridViewComboBoxColumn()
                clm2.DataSource = dt
                clm2.DataPropertyName = "SiteAt"
                clm2.ValueMember = "Site"
                clm2.DisplayMember = "Site"
                clm2.HeaderText = "Site At"
                ctl.Columns.Add(clm2)

                Dim clm5 As New DataGridViewComboBoxColumn()
                clm5.DataSource = OverClass.TempDataTable("SELECT DISTINCT '' As LocationAtSite FROM SiteLocation " &
                                                          "UNION ALL SELECT LocationAtSite FROM SiteLocation ORDER BY LocationAtSite ASC")
                clm5.DataPropertyName = "LocationAtSite"
                clm5.ValueMember = "LocationAtSite"
                clm5.DisplayMember = "LocationAtSite"
                clm5.HeaderText = "Location At Site"
                ctl.Columns.Add(clm5)

                Dim clm3 As New DataGridViewComboBoxColumn()
                clm3.DataSource = dt
                clm3.DataPropertyName = "From"
                clm3.ValueMember = "Site"
                clm3.DisplayMember = "Site"
                clm3.HeaderText = "From"
                ctl.Columns.Add(clm3)

                Dim clm4 As New DataGridViewComboBoxColumn()
                clm4.DataSource = dt
                clm4.DataPropertyName = "To"
                clm4.ValueMember = "Site"
                clm4.DisplayMember = "Site"
                clm4.HeaderText = "To"
                ctl.Columns.Add(clm4)


                Dim clm6 As New DataGridViewImageColumn()
                clm6.Image = My.Resources.Preview
                clm6.ImageLayout = DataGridViewImageCellLayout.Zoom
                clm6.HeaderText = "View History"
                clm6.Name = "History"
                ctl.Columns.Add(clm6)

                Dim clm7 As New DataGridViewImageColumn
                clm7.Name = "PickColour"
                clm7.HeaderText = "Pick Colour"
                clm7.Image = My.Resources.Art_512
                clm7.ImageLayout = DataGridViewImageCellLayout.Zoom
                ctl.columns.add(clm7)

                ctl.AllowUserToAddRows = False
                ctl.AutoGenerateColumns = False

                ctl.Columns("Colour").Visible = False
                ctl.Columns("PSPSite").Visible = False
                ctl.Columns("SiteAt").Visible = False
                ctl.Columns("LocationAtSite").Visible = False
                ctl.Columns("From").Visible = False
                ctl.Columns("To").Visible = False

                ctl.Columns("Timestamp").ReadOnly = True
                ctl.Columns("Person").ReadOnly = True
                ctl.Columns("SubjectID").ReadOnly = True

                ctl.Columns("Timestamp").DefaultCellStyle.Format = "dd-MMM-yyyy"

                ctl.Columns("SubjectID").DisplayIndex = 0
                ctl.Columns("Timestamp").DisplayIndex = 1
                ctl.Columns("Person").DisplayIndex = 2
                clm1.DisplayIndex = 3
                clm2.DisplayIndex = 4
                clm5.DisplayIndex = 5
                clm3.DisplayIndex = 6
                clm4.DisplayIndex = 7
                ctl.Columns("Comments").DisplayIndex = 8
                clm7.DisplayIndex = 9
                clm6.DisplayIndex = 10

            Case "ComboBox1"

                ComboBox1.DataSource = OverClass.TempDataTable("SELECT DISTINCT '' AS Site FROM SITE " &
                                                              "UNION ALL SELECT Site FROM Site ORDER BY Site ASC")
                ComboBox1.DisplayMember = "Site"
                ComboBox1.ValueMember = "Site"

                ComboBox2.DataSource = OverClass.TempDataTable("SELECT DISTINCT '' AS Site FROM SITE " &
                                                              "UNION ALL SELECT Site FROM Site ORDER BY Site ASC")
                ComboBox2.DisplayMember = "Site"
                ComboBox2.ValueMember = "Site"

                ComboBox3.DataSource = OverClass.TempDataTable("SELECT DISTINCT '' AS Site FROM SITE " &
                                                              "UNION ALL SELECT Site FROM Site ORDER BY Site ASC")
                ComboBox3.DisplayMember = "Site"
                ComboBox3.ValueMember = "Site"

                ComboBox4.DataSource = OverClass.TempDataTable("SELECT DISTINCT '' As LocationAtSite FROM SiteLocation " &
                                                          "UNION ALL SELECT LocationAtSite FROM SiteLocation ORDER BY LocationAtSite ASC")
                ComboBox4.ValueMember = "LocationAtSite"
                ComboBox4.DisplayMember = "LocationAtSite"

        End Select

    End Sub

    Private Sub CheckBox1_Click(sender As Object, e As EventArgs) Handles CheckBox1.Click
        CheckBox1.CheckState = CheckState.Checked
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        ComboBox5.SelectedValue = ""
        ComboBox6.SelectedValue = ""
        Call Specifics(DataGridView1)
    End Sub

    Private Sub CheckBox2_Click(sender As Object, e As EventArgs) Handles CheckBox2.Click
        CheckBox2.CheckState = CheckState.Checked
        CheckBox1.Checked = False
        CheckBox3.Checked = False
        ComboBox5.SelectedValue = ""
        ComboBox6.SelectedValue = ""
        Call Specifics(DataGridView1)
    End Sub

    Private Sub CheckBox3_Click(sender As Object, e As EventArgs) Handles CheckBox3.Click
        CheckBox3.CheckState = CheckState.Checked
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        ComboBox5.SelectedValue = ""
        ComboBox6.SelectedValue = ""
        Call Specifics(DataGridView1)
    End Sub

    Private Sub DataGridView1_Paint(sender As Object, e As PaintEventArgs) Handles DataGridView1.Paint

        For Each row As DataGridViewRow In DataGridView1.Rows
            If IsDBNull(row.Cells("Colour").Value) = True Then Continue For
            If row.Cells("Colour").Value = vbNullString Then Continue For
            row.DefaultCellStyle.BackColor = ColorTranslator.FromHtml(row.Cells("Colour").Value)
        Next

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        If e.RowIndex < 0 Then Exit Sub


        If e.ColumnIndex = DataGridView1.Columns("PickColour").Index Then
            If ColourEdit = False Then Exit Sub
            Dim cDialog As New ColorDialog()
            If (cDialog.ShowDialog() = DialogResult.OK) Then
                DataGridView1.Item("Colour", e.RowIndex).Value = ColorTranslator.ToHtml(cDialog.Color)
                DataGridView1.Rows(e.RowIndex).DefaultCellStyle.BackColor = cDialog.Color
            End If
        ElseIf e.ColumnIndex = DataGridView1.Columns("History").Index Then
            Dim SQLString As String = "SELECT Person & ' (' & format(Timestamp,'dd-MMM-yyyy') & ') ' & chr(10) & chr(13) " &
                "& iif(LocationAtSite='In Transit','',' Site: ' & SiteAt) " &
                "& ' Location: ' & LocationAtSite & ' ' & iif(LocationAtSite='In Transit',[FROM] & '>' & To & '.','') " &
                "FROM History WHERE SubjectID=" & DataGridView1.Item("SubjectID", e.RowIndex).Value &
                " ORDER BY Timestamp DESC"
            Dim CSVString = OverClass.CreateCSVString(SQLString)
            MsgBox(Trim(Replace(CSVString, ",", vbNewLine & vbNewLine)),, DataGridView1.Item("SubjectID", e.RowIndex).Value)

        End If

    End Sub

    Private Sub DataGridView1_RowValidating(sender As Object, e As DataGridViewCellCancelEventArgs) Handles DataGridView1.RowValidating

        Dim LocationAtSite As String = ""
        Dim SiteAt As String = ""
        Dim FromString As String = ""
        Dim ToString As String = ""

        Try
            LocationAtSite = DataGridView1.Item("LocationAtSite", e.RowIndex).Value
        Catch ex As Exception
        End Try
        Try
            SiteAt = DataGridView1.Item("Siteat", e.RowIndex).Value
        Catch ex As Exception
        End Try
        Try
            FromString = DataGridView1.Item("From", e.RowIndex).Value
        Catch ex As Exception
        End Try
        Try
            ToString = DataGridView1.Item("To", e.RowIndex).Value
        Catch ex As Exception
        End Try

        If LocationAtSite <> "" And LocationAtSite <> "In Transit" And SiteAt = "" Then
            MsgBox("A Site must be chosen if a location is selected")
            e.Cancel = True
        End If

        If LocationAtSite = "In Transit" And (FromString = "" Or ToString = "") Then
            MsgBox("A 'From' and 'To' Site must be selected")
            e.Cancel = True
        End If

        If LocationAtSite <> "In Transit" And (FromString <> "" Or ToString <> "") Then
            MsgBox("'From' and 'To' can only be selected with location 'In Transit'")
            e.Cancel = True
        End If

        If SiteAt <> "" And LocationAtSite = "" Then
            MsgBox("A Location must be selected if a site has been chosen")
            e.Cancel = True
        End If

        If SiteAt <> "" And LocationAtSite = "In Transit" Then
            MsgBox("A site must not be selected if location is 'In Transit'")
            e.Cancel = True
        End If

        If ToString = FromString And ToString <> "" And FromString <> "" Then
            MsgBox("Sites cannot be the same")
            e.Cancel = True
        End If

    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged

        If ComboBox4.SelectedValue.ToString = "In Transit" Then

            ComboBox1.Visible = False
            Label4.Visible = False
            ComboBox1.SelectedValue = ""
            ComboBox3.Visible = True
            ComboBox2.Visible = True
            Label5.Visible = True
            Label6.Visible = True
        Else
            ComboBox1.Visible = True
            Label4.Visible = True
            ComboBox3.Visible = False
            ComboBox2.Visible = False
            ComboBox3.SelectedValue = ""
            ComboBox2.SelectedValue = ""
            Label5.Visible = False
            Label6.Visible = False
        End If

    End Sub
End Class
