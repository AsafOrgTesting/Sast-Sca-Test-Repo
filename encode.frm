Imports System.Data.SqlClient

Private Sub cmdSafe_Click()
Dim user_name As String
Dim password As String
Dim query As String
Dim rs As DAO.Recordset

    ' Get the user name and password.
    user_name = Replace$(txtUserName.Text, "'", "''")
    password = Replace$(txtPassword.Text, "'", "''")

    ' Compose the query.
    query = "SELECT COUNT (*) FROM Passwords " & _
        "WHERE UserName='" & user_name & "'" & _
        "  AND Password='" & password & "'"
    txtQuery.Text = query

    ' Execute the query.
    On Error Resume Next
    Set rs = m_DB.OpenRecordset(query, dbOpenSnapshot)
    If Err.Number <> 0 Then
        lblValid.Caption = "Invalid Query"
    ElseIf (CInt(rs.Fields(0)) > 0) Then
        lblValid.Caption = "Valid"
    Else
        lblValid.Caption = "Invalid"
    End If

    rs.Close
End Sub


Private Sub cmdUnsafe_Click()
Dim user_name As String
Dim password As String
Dim conn As ADODB.Connection
Dim cmd As ADODB.Command
Dim rs As ADODB.Recordset

    ' Get the user name and password.
    user_name = txtUserName.Text
    password = txtPassword.Text

    ' Display the parameterized query template (no user data embedded).
    txtQuery.Text = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

    ' Execute the query using parameterized ADODB Command to prevent SQL Injection.
    On Error GoTo Error_Handler

    Set conn = New ADODB.Connection
    conn.Open m_DB.Name

    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "SELECT COUNT (*) FROM Passwords WHERE UserName=? AND Password=?"

    ' Bind parameters — user input is passed as data, never interpreted as SQL.
    cmd.Parameters.Append cmd.CreateParameter("UserName", adVarChar, adParamInput, 255, user_name)
    cmd.Parameters.Append cmd.CreateParameter("Password", adVarChar, adParamInput, 255, password)

    Set rs = cmd.Execute

    If (CInt(rs.Fields(0)) > 0) Then
        lblValid.Caption = "Valid"
    Else
        lblValid.Caption = "Invalid"
    End If

    rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing
    Set cmd = Nothing
    Exit Sub

Error_Handler:
    lblValid.Caption = "Invalid Query"
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
    End If
    Set rs = Nothing
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
    End If
    Set conn = Nothing
    Set cmd = Nothing
End Sub


p = txtP.Text
Dim conn As New ADODB.Connection
conn.Open "connection string"

Dim cmd As New ADODB.Command
With cmd
    .ActiveConnection = conConnection
    .CommandText = "SELECT fields FROM table WHERE condition = ?"
    .CommandType = adCmdText
End With

Dim param As New ADODB.Parameter
Set param = cmd.CreateParameter("condition", adVarChar, adParamInput, 5, "value")
cmd.Parameters.Append p

Dim rs As New ADODB.Recordset
rs.CursorLocation = adUseClient
rs.Open cmd, , adOpenStatic, adLockOptimistic

Dim temp
Do While Not rs.EOF
    temp = rs("field")
    rs.MoveNext
Loop

rs.Close
conn.Close



Public Class Form1
    Private Sub Button1_Click(ByVal sender As System.Object, _
                        ByVal e As System.EventArgs) Handles Button1.Click
        Dim con As SqlConnection = New SqlConnection( _
                        "Data Source=.;Integrated Security=True;AttachDbFilename=D:\myDB.mdf")
        con.Open()
        Dim cmdText As String = _
                        "INSERT INTO Customer(UserName, [Password]) VALUES (@UserName,@Password)"
        Dim cmd As SqlCommand = New SqlCommand(cmdText, con)
        With cmd.Parameters
            .Add(New SqlParameter("@UserName", txtUserName.Text))
            .Add(New SqlParameter("@Password", txtPassword.Text))
        End With
        cmd.ExecuteNonQuery()
        con.Close()
        con = Nothing
    End Sub
End Class