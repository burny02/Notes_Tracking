Module ButtonModule

    Public Sub ButtonSpecifics(sender As Object, e As EventArgs)

        Dim ctl As Object = Nothing

        Select Case sender.name.ToString

            Case "Button1"
                Call Saver(Form1.DataGridView1)

            Case "Button2"
                Call UpdateFromMeddbase()
                Call Form1.Specifics(Form1.DataGridView1)
                MsgBox("Update Complete")

        End Select



    End Sub


End Module
