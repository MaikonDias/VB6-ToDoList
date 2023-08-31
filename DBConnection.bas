Attribute VB_Name = "DBConnection"
Public conn As New ADODB.Connection
Public rs As New ADODB.Recordset

Public index As Integer

Public Sub Connect()
    conn.CursorLocation = adUseClient
    conn.Open "provider=microsoft.jet.oledb.4.0;persist security info = false; data source = " & App.Path & "\ToDoList.mdb;Data"
End Sub

Public Sub LoadDataGrid()
    On Error GoTo ErrorDescription

    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM Tarefas", conn, adOpenStatic, adLockOptimistic
    
    If Not rs.EOF Then
        Set DataGrid1.DataSource = rs
    Else
        MsgBox "No records found."
    End If

    Exit Sub

ErrorDescription:
    MsgBox ("Error loading data: " & Err.Description)
End Sub




Public Sub CloseConnection()
    If conn.State = adStateOpen Then
        conn.Close
    End If
End Sub

