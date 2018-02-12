Attribute VB_Name = "mdBaseDatos"
Option Explicit

Public adoConnection As ADODB.Connection

Public Sub ConectarBaseDatos()
On Error GoTo ErrorHandle
    
    Set adoConnection = New ADODB.Connection
    With adoConnection
        '###
'       .Provider = "Microsoft.Jet.OLEDB.4.0"
'       .Provider = "Microsoft.Jet.OLEDB.3.51"
'       .ConnectionString = App.Path & "\..\Servidor\Tegnet.tnt"
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Tegnet.tnt;Jet OLEDB:Database Password=koala;"
    End With
    adoConnection.Open
    
    Exit Sub

ErrorHandle:
    ReportErr "ConectarBaseDatos", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
End Sub

Public Sub DesconectarBaseDatos()
On Error GoTo ErrorHandle
    
    adoConnection.Close
    Set adoConnection = Nothing
    Exit Sub

ErrorHandle:
    ReportErr "DesconectarBaseDatos", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
End Sub

Public Function EjecutarConsulta(strSQL As String, Optional intCursorType As ADODB.CursorTypeEnum = adOpenKeyset) As ADODB.Recordset
    'On Error GoTo ErrorHandle
    
    Set EjecutarConsulta = New ADODB.Recordset
    
    With EjecutarConsulta
        .CursorType = intCursorType
        .ActiveConnection = adoConnection
        .Open strSQL
    End With

    Exit Function
ErrorHandle:
    ReportErr "EjecutarConsulta", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
    
End Function

Public Function EjecutarConsultaValor(strSQL As String) As Variant
    'On Error GoTo ErrorHandle
    Dim rsRecordset As ADODB.Recordset
    
    Set rsRecordset = New ADODB.Recordset
    
    With rsRecordset
        .CursorType = adOpenForwardOnly
        .ActiveConnection = adoConnection
        .Open strSQL
    End With
    
    If Not rsRecordset.EOF Then
        EjecutarConsultaValor = rsRecordset.Fields(0).Value
    Else
        EjecutarConsultaValor = Null
    End If
    
    rsRecordset.Close
    Set rsRecordset = Nothing
    
    Exit Function
ErrorHandle:
    ReportErr "EjecutarConsultaValor", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
    Set rsRecordset = Nothing
End Function


Public Sub EjecutarComando(strSQL As String)
    'On Error GoTo ErrorHandle
    
    adoConnection.Execute strSQL
    
    Exit Sub
ErrorHandle:
    ReportErr "EjecutarComando", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
    
End Sub

Public Sub RecordsetAVector(rsRecordset As Recordset, intCol As Integer, ByRef vecVector() As String)
    On Error GoTo ErrorHandle
    'Dado un recordset y una columna devuelve un vector que los contiene
    Dim i As Integer
    ReDim vecVector(0)
    i = 0
    'Si esta vacio no hace nada
    If Not rsRecordset.EOF Then
        'Cuenta la cantidad de registros del recordset
        rsRecordset.MoveLast
        ReDim vecVector(rsRecordset.RecordCount - 1)
        rsRecordset.MoveFirst
        While Not rsRecordset.EOF
            vecVector(i) = rsRecordset.Fields(intCol).Value
            i = i + 1
            rsRecordset.MoveNext
        Wend
    End If
    
    Exit Sub
ErrorHandle:
    ReportErr "RecordsetAVector", "mdlBaseDatos", Err.Description, Err.Number, Err.Source
End Sub


