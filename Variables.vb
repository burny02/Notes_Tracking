Imports TemplateDB

Module Variables
    Public OverClass As OverClass
    Private Const TablePath As String = "M:\VOLUNTEER SCREENING SERVICES\Systems\Notes_Tracking\Backend.accdb"
    Private Const PWord As String = "RetroRetro*1"
    Private Const Connect2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & TablePath & ";Jet OLEDB:Database Password=" & PWord
    Private Const UserTable As String = "[ApprovedUsers]"
    Private Const UserField As String = "User"
    Private Const LockTable As String = "[Locker]"
    Private Const AuditTable As String = "[Audit2]"
    Private Contact As String = "Michal Sieracki"
    Public Const SolutionName As String = "Notes Tracking"
    Public ColourEdit As Boolean = False


    Public Sub StartUp(WhichForm As Form)

        OverClass = New TemplateDB.OverClass
        OverClass.SetPrivate(UserTable,
                           UserField,
                           LockTable,
                           Contact,
                           Connect2,
                           AuditTable)

        OverClass.LockCheck()

        OverClass.LoginCheck()

        OverClass.AddAllDataItem(WhichForm)

        ColourEdit = OverClass.TempDataTable("SELECT ColourEdit FROM " & UserTable & " WHERE " & UserField & "='" & OverClass.GetUserName & "'").Rows(0).Item(0)

        For Each ctl In OverClass.DataItemCollection
            If (TypeOf ctl Is Button) Then
                Dim But As Button = ctl
                AddHandler But.Click, AddressOf ButtonSpecifics
            End If
        Next


    End Sub

End Module
