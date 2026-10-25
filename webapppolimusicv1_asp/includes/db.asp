<%
' ============================================================================
' PoliMusic - Classic ASP Database Connection & Utilities
' SQL Server 2025 Compatible
' ============================================================================

Function GetDBConnection()
    On Error Resume Next
    Dim conn, connStr, success
    Set conn = Server.CreateObject("ADODB.Connection")
    conn.CommandTimeout = 30
    success = False

    ' 1. Try default localhost / MSSQLSERVER with MSOLEDBSQL
    connStr = "Provider=MSOLEDBSQL;Server=.;Database=BDD_PoliMusic;UID=usr_polimusic;PWD=Politecnica1;TrustServerCertificate=True;"
    conn.Open connStr
    If Err.Number = 0 Then
        success = True
    Else
        Err.Clear
        ' 2. Try default localhost with Windows Integrated Security
        connStr = "Provider=MSOLEDBSQL;Server=.;Database=BDD_PoliMusic;Integrated Security=SSPI;TrustServerCertificate=True;"
        conn.Open connStr
        If Err.Number = 0 Then
            success = True
        Else
            Err.Clear
            ' 3. Try SQL Server LocalDB with MSOLEDBSQL
            connStr = "Provider=MSOLEDBSQL;Server=(localdb)\MSSQLLocalDB;Database=BDD_PoliMusic;UID=usr_polimusic;PWD=Politecnica1;TrustServerCertificate=True;"
            conn.Open connStr
            If Err.Number = 0 Then
                success = True
            Else
                Err.Clear
                ' 4. Try ODBC Driver 17 for SQL Server with LocalDB
                connStr = "Driver={ODBC Driver 17 for SQL Server};Server=(localdb)\MSSQLLocalDB;Database=BDD_PoliMusic;Uid=usr_polimusic;Pwd=Politecnica1;"
                conn.Open connStr
                If Err.Number = 0 Then
                    success = True
                Else
                    Err.Clear
                    ' 5. Try SQLEXPRESS
                    connStr = "Provider=MSOLEDBSQL;Server=localhost\SQLEXPRESS;Database=BDD_PoliMusic;UID=usr_polimusic;PWD=Politecnica1;TrustServerCertificate=True;"
                    conn.Open connStr
                    If Err.Number = 0 Then
                        success = True
                    End If
                End If
            End If
        End If
    End If

    If Not success Then
        Response.Write "<div style='background:#f8d7da;color:#721c24;padding:20px;border-radius:8px;font-family:sans-serif;margin:20px;'>"
        Response.Write "<h3>Error de Conexión a la Base de Datos</h3>"
        Response.Write "<p>No se pudo conectar a SQL Server BDD_PoliMusic. Verifique que la instancia de SQL Server esté iniciada y las credenciales 'usr_polimusic' sean válidas.</p>"
        Response.Write "<p>Detalle: " & Server.HTMLEncode(Err.Description) & "</p>"
        Response.Write "</div>"
        Response.End
    End If

    Set GetDBConnection = conn
End Function

Sub CloseDBConnection(ByRef conn)
    On Error Resume Next
    If IsObject(conn) Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
End Sub

Sub CloseRecordset(ByRef rs)
    On Error Resume Next
    If IsObject(rs) Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
End Sub

' Sanitiza cadenas para evitar SQL Injection básico
Function SafeSql(val)
    If IsNull(val) Or IsEmpty(val) Then
        SafeSql = ""
    Else
        SafeSql = Replace(Trim(val), "'", "''")
    End If
End Function

' Resuelve rutas de medios que inician con ~/ a rutas web relativas correctas
Function ResolveMediaUrl(path, defaultFallback)
    Dim cleaned
    If IsNull(path) Or Trim(path) = "" Then
        ResolveMediaUrl = defaultFallback
        Exit Function
    End If
    
    cleaned = Trim(path)
    ' Si comienza con ~/, removerlo
    If Left(cleaned, 2) = "~/" Then
        cleaned = Mid(cleaned, 3)
    ElseIf Left(cleaned, 1) = "/" Or Left(cleaned, 1) = "\" Then
        cleaned = Mid(cleaned, 2)
    End If
    
    ' Reemplazar barras invertidas
    cleaned = Replace(cleaned, "\", "/")
    
    ResolveMediaUrl = cleaned
End Function

' Formatea nombre de artista
Function FormatArtistFullName(name, lastname)
    Dim n, l
    n = Trim("" & name)
    l = Trim("" & lastname)
    If l <> "" Then
        FormatArtistFullName = n & " " & l
    Else
        FormatArtistFullName = n
    End If
End Function
%>
