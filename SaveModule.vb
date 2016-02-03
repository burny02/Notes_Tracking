Module SaveModule
    Public Sub Saver(ctl As Object)

        Dim DisplayMessage As Boolean = True

        'Get a generic command list first - Ignore errors (Multi table)
        Dim cb As New OleDb.OleDbCommandBuilder(OverClass.CurrentDataAdapter)

        Try
            OverClass.CurrentDataAdapter.UpdateCommand = cb.GetUpdateCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.InsertCommand = cb.GetInsertCommand()
        Catch
        End Try
        Try
            OverClass.CurrentDataAdapter.DeleteCommand = cb.GetDeleteCommand()
        Catch
        End Try


        'Create and overwrite a custom one if needed (More than 1 table) ...OLEDB Parameters must be added in the order they are used
        Select Case ctl.name


            Case "DataGridView1"

                Dim Person As String = "'" & OverClass.GetUserName & "'"
                Dim Timestamp As String = OverClass.SQLDate(DateAndTime.Now)

                OverClass.CurrentDataAdapter.UpdateCommand = New OleDb.OleDbCommand("UPDATE Notes " &
                                                                   "SET [PSPSite]=@P1, [SiteAt]=@P2, [LocationAtSite]=@P3, " &
                                                                    "[From]=@P4, [TO]=@P5, [Timestamp]=" & Timestamp & "," &
                                                                    "[Person]=" & Person & ", [Comments]=@P6, [Colour]=@P7 " &
                                                                    "WHERE [SubjectID]=@P8")


                With OverClass.CurrentDataAdapter.UpdateCommand.Parameters
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "PSPSite")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "SiteAt")
                    .Add("@P3", OleDb.OleDbType.VarChar, 255, "LocationAtSite")
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "FROM")
                    .Add("@P2", OleDb.OleDbType.VarChar, 255, "To")
                    .Add("@P3", OleDb.OleDbType.VarChar, 255, "Comments")
                    .Add("@P1", OleDb.OleDbType.VarChar, 255, "Colour")
                    .Add("@P2", OleDb.OleDbType.Double, 255, "SubjectID")
                End With

        End Select


        Call OverClass.SetCommandConnection()
        Call OverClass.UpdateBackend(ctl, DisplayMessage)
        If DisplayMessage = False Then MsgBox("Table Updated")


    End Sub


End Module
