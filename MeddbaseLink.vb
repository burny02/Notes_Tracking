Option Explicit On
Imports System.IO
Imports System.Threading

Module MeddbaseLink

    Public Sub UpdateFromMeddbase()



        Dim appExcel As Object
        Dim objWorkSheet As Object
        Dim objQueryTable As Object

        Dim ExcelLocation = "M:\VOLUNTEER SCREENING SERVICES\Systems\Notes_Tracking\MeddBase.xlsx"
        Dim SheetName = "Sheet1"
        Dim QueryName = "Meddbase"
        Dim Attempt As Integer = 0

        Do While Attempt < 10
            If IsWorkBookOpen(ExcelLocation) = True Then
                Attempt = Attempt + 1
                Thread.Sleep(2000)
            Else
                Attempt = 20
            End If
        Loop

        If Attempt = 10 Then
            MsgBox("Meddbase link currently in use." & vbNewLine & vbNewLine & "For newest RVL ID's try again in 10 minutes.")
            End
        End If

        appExcel = GetObject(ExcelLocation)
        objWorkSheet = appExcel.Worksheets(SheetName)

        objQueryTable = objWorkSheet.QueryTables(QueryName)

        objQueryTable.Refresh(False)

        While objQueryTable.Refreshing
            'Do Nothing
        End While

        objQueryTable = Nothing
        objWorkSheet = Nothing

        appExcel.Close(True)

        appExcel = Nothing


        OverClass.ExecuteSQL("DELETE FROM Meddbase WHERE ID IS NOT NULL")

        Dim SQLString As String = "INSERT INTO Meddbase ( Id, DOB, FName, MName, SName )
        SELECT Link.Id, Link.[Date of Birth], Link.Name, Link.[Middle Name], Link.Surname
        FROM Link"

        OverClass.ExecuteSQL(SQLString)

        SQLString = "INSERT INTO Notes ([SubjectID]) SELECT LINK.[ID] FROM LINK WHERE LINK.ID > (SELECT max([SubjectID]) FROM [Notes])"

        OverClass.ExecuteSQL(SQLString)

    End Sub


    Function IsWorkBookOpen(ByRef sName As String) As Boolean
        Dim fs As FileStream
        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
            IsWorkBookOpen = False
        Catch ex As Exception
            IsWorkBookOpen = True
        End Try
    End Function

End Module
