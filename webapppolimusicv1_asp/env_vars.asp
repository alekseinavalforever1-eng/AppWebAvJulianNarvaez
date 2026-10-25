<%@ Language="VBScript" %>
<!--#include file="includes/db.asp"-->
<!--#include file="includes/auth.asp"-->
<%
' ============================================================================
' PoliMusic - Variables de Entorno del Servidor (Classic ASP / VBScript)
' ----------------------------------------------------------------------------
' Fuentes de datos realmente disponibles en Classic ASP:
'   1. PROCESS -> WScript.Shell.Environment("PROCESS")  (hereda maquina + usuario)
'   2. MAQUINA -> Registro de Windows HKLM\SYSTEM\CurrentControlSet\Control\
'                 Session Manager\Environment  (WScript.Shell.RegRead)
'   3. USUARIO -> WScript.Shell.Environment("USER")
'   4. SV      -> Request.ServerVariables
'   5. HEADERS -> Request.ServerVariables("ALL_RAW") (peticion HTTP cruda)
'   6. APP     -> Application.Contents
'   7. SESSION -> Session.Contents
' ============================================================================
RequireAdmin()

Const SEP = "|~|~|"                                    ' Separador interno de cadenas
Const MACHINE_ENV_KEY = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\"

' ---------------------------------------------------------------------------
' UTILIDADES
' ---------------------------------------------------------------------------

' Convierte un valor a cadena de forma segura (evita errores con objetos/Null)
Function SafeStr(v)
    Dim s
    s = ""
    On Error Resume Next
    If IsObject(v) Then
        s = "[Objeto]"
    ElseIf IsArray(v) Then
        s = "[Arreglo]"
    ElseIf IsNull(v) Then
        s = ""
    ElseIf IsEmpty(v) Then
        s = ""
    Else
        s = CStr(v)
    End If
    If Err.Number <> 0 Then
        Err.Clear
        s = "[valor no convertible]"
    End If
    On Error GoTo 0
    SafeStr = s
End Function

' Agrega un nombre a una cadena acumulada (evita arrays O(n^2))
Sub AppendName(ByRef acc, name)
    If Len(acc) > 0 Then acc = acc & SEP
    acc = acc & CStr(name)
End Sub

' Convierte la cadena acumulada en un array de nombres
Function SplitNames(acc)
    If Len(acc) = 0 Then
        SplitNames = Array()
    Else
        SplitNames = Split(acc, SEP)
    End If
End Function

' Ordena un array de nombres (case-insensitive)
Function SortNames(arr)
    Dim i, j, tmp
    If UBound(arr) < 1 Then
        SortNames = arr
        Exit Function
    End If
    For i = 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= 0
            If StrComp(arr(j), tmp, 1) > 0 Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next
    SortNames = arr
End Function

' ---------------------------------------------------------------------------
' 1) VARIABLES DE ENTORNO DEL PROCESO Y DEL USUARIO
'    OJO: WshEnvironment no expone .Keys en VBScript. La enumeracion devuelve
'    entradas con el formato "NOMBRE=VALOR", por lo que se separa en el 1er "=".
' ---------------------------------------------------------------------------
Function GetEnvScope(scopeName)
    Dim d, shell, env, entry, sepPos, nm, vl
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    On Error Resume Next
    Err.Clear
    Set shell = Server.CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        Set env = shell.Environment(scopeName)
        If Err.Number = 0 Then
            For Each entry In env
                Err.Clear
                sepPos = InStr(entry, "=")
                If sepPos > 1 Then
                    nm = Left(entry, sepPos - 1)
                    vl = Mid(entry, sepPos + 1)
                    If Not d.Exists(nm) Then d.Add nm, vl
                End If
            Next
        End If
    End If
    Err.Clear
    On Error GoTo 0
    Set GetEnvScope = d
End Function

' ---------------------------------------------------------------------------
' 2) VARIABLES DE ENTORNO DEL SISTEMA (REGISTRO DE WINDOWS)
'    El ambito "MACHINE" de WScript.Shell no es accesible desde el sandbox de
'    ASP, por lo que se lee la clave del registro que las define.
' ---------------------------------------------------------------------------
Function GetMachineEnvVars()
    Dim d, shell, i, nm, v
    Dim stdMachineNames
    stdMachineNames = Array( _
        "Path", "PATHEXT", "SystemRoot", "windir", "TEMP", "TMP", _
        "ProgramFiles", "ProgramFiles(x86)", "ProgramW6432", _
        "CommonProgramFiles", "CommonProgramFiles(x86)", "CommonProgramW6432", _
        "ProgramData", "ALLUSERSPROFILE", "SystemDrive", "HOMEDRIVE", "HOMEPATH", _
        "COMPUTERNAME", "OS", "PROCESSOR_ARCHITECTURE", "PROCESSOR_ARCHITEW6432", _
        "PROCESSOR_IDENTIFIER", "PROCESSOR_LEVEL", "NUMBER_OF_PROCESSORS", _
        "ComSpec", "PSModulePath", "DriverData", "JAVA_HOME", "GOPATH", _
        "NODE_PATH", "PYTHONPATH", "PYTHONHOME", "SESSIONNAME", "USERPROFILE", _
        "LOGONSERVER", "SESSIONNAME", "NUMBER_OF_PROCESSORS")

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    On Error Resume Next
    Err.Clear
    Set shell = Server.CreateObject("WScript.Shell")
    If Err.Number = 0 Then
        For i = 0 To UBound(stdMachineNames)
            nm = stdMachineNames(i)
            Err.Clear
            v = shell.RegRead(MACHINE_ENV_KEY & nm)
            If Err.Number = 0 Then
                If Len(SafeStr(v)) > 0 And Not d.Exists(nm) Then d.Add nm, SafeStr(v)
            End If
            Err.Clear
        Next
    End If
    Err.Clear
    On Error GoTo 0
    Set GetMachineEnvVars = d
End Function

' ---------------------------------------------------------------------------
' 5) CABECERAS HTTP DE LA PETICION
'    Request.Headers no es utilizable en ASP clasico, pero la peticion cruda si
'    esta disponible en Request.ServerVariables("ALL_RAW").
' ---------------------------------------------------------------------------
Function GetRawHttpHeaders()
    Dim d, raw, lines, i, ln, p, nm, vl
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    On Error Resume Next
    Err.Clear
    raw = Request.ServerVariables("ALL_RAW")
    If Err.Number = 0 And Len(raw) > 0 Then
        lines = Split(raw, vbCrLf)
        For i = 0 To UBound(lines)
            ln = Trim(lines(i))
            If Len(ln) > 0 Then
                If Left(ln, 4) <> "GET " And Left(ln, 5) <> "POST " And Left(ln, 5) <> "HEAD " _
                   And Left(ln, 4) <> "PUT " And Left(ln, 7) <> "DELETE " And Left(ln, 8) <> "OPTIONS " Then
                    p = InStr(ln, ":")
                    If p > 1 Then
                        nm = Trim(Left(ln, p - 1))
                        vl = Trim(Mid(ln, p + 1))
                        If Len(nm) > 0 Then
                            If d.Exists(nm) Then
                                d(nm) = d(nm) & ", " & vl
                            Else
                                d.Add nm, vl
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    Set GetRawHttpHeaders = d
End Function

' ---------------------------------------------------------------------------
' 6-7) NOMBRES DE VARIABLES DE APPLICATION / SESSION
'     (For Each sobre Contents devuelve los nombres de las variables)
' ---------------------------------------------------------------------------
Function GetContentsNames(coll)
    Dim acc, item
    acc = ""
    On Error Resume Next
    For Each item In coll
        Err.Clear
        If Len(SafeStr(item)) > 0 Then AppendName acc, SafeStr(item)
    Next
    Err.Clear
    On Error GoTo 0
    GetContentsNames = SplitNames(acc)
End Function

Function GetAppValue(name)
    Dim v
    v = ""
    On Error Resume Next
    Err.Clear
    v = Application.Contents(name)
    If Err.Number <> 0 Then
        Err.Clear
        v = Application(name)
    End If
    Err.Clear
    On Error GoTo 0
    GetAppValue = SafeStr(v)
End Function

Function GetSessionValue(name)
    Dim v
    v = ""
    On Error Resume Next
    Err.Clear
    v = Session.Contents(name)
    If Err.Number <> 0 Then
        Err.Clear
        v = Session(name)
    End If
    Err.Clear
    On Error GoTo 0
    GetSessionValue = SafeStr(v)
End Function

' ---------------------------------------------------------------------------
' RECOLECCION DE DATOS
' ---------------------------------------------------------------------------
Dim envError
envError = ""
On Error Resume Next
Err.Clear
Set shellTest = Server.CreateObject("WScript.Shell")
If Err.Number <> 0 Then
    envError = "No se pudo crear el objeto WScript.Shell (" & Err.Description & "). "
    envError = envError & "Verifique que el componente est&eacute; registrado y que el proceso de IIS tenga permisos para leerlo."
    Err.Clear
End If
Set shellTest = Nothing
On Error GoTo 0

Dim dictProc, dictMachine, dictUser, dictHeaders, dictApp, dictSession
Set dictProc    = GetEnvScope("PROCESS")
Set dictUser    = GetEnvScope("USER")
Set dictMachine = GetMachineEnvVars()
Set dictHeaders = GetRawHttpHeaders()

' --- Application / Session ---
Dim appNames, sessNames, i, nm
appNames  = GetContentsNames(Application.Contents)
sessNames = GetContentsNames(Session.Contents)

Set dictApp = CreateObject("Scripting.Dictionary")
dictApp.CompareMode = 1
For i = 0 To UBound(appNames)
    If appNames(i) <> "" Then
        If Not dictApp.Exists(appNames(i)) Then dictApp.Add appNames(i), GetAppValue(appNames(i))
    End If
Next

Set dictSession = CreateObject("Scripting.Dictionary")
dictSession.CompareMode = 1
For i = 0 To UBound(sessNames)
    If sessNames(i) <> "" Then
        If Not dictSession.Exists(sessNames(i)) Then dictSession.Add sessNames(i), GetSessionValue(sessNames(i))
    End If
Next

' --- ServerVariables ---
' Request.ServerVariables no expone .Keys en ASP clasico, por lo que se usa la
' lista estandar de IIS + los nombres HTTP_ deducidos de la peticion cruda.
Dim stdSvNames
stdSvNames = Array( _
    "ALL_RAW", "AUTH_TYPE", "CONTENT_LENGTH", "CONTENT_TYPE", "GATEWAY_INTERFACE", _
    "HTTPS", "LOGON_USER", "PATH_INFO", "PATH_TRANSLATED", "QUERY_STRING", _
    "REMOTE_ADDR", "REMOTE_HOST", "REMOTE_IDENT", "REMOTE_USER", "REQUEST_METHOD", _
    "SCRIPT_NAME", "SERVER_NAME", "SERVER_PORT", "SERVER_PORT_SECURE", _
    "SERVER_PROTOCOL", "SERVER_SOFTWARE", "URL", "APPL_PHYSICAL_PATH", "APPL_MD_PATH", _
    "HTTP_ACCEPT", "HTTP_ACCEPT_CHARSET", "HTTP_ACCEPT_ENCODING", "HTTP_ACCEPT_LANGUAGE", _
    "HTTP_AUTHORIZATION", "HTTP_CACHE_CONTROL", "HTTP_CONNECTION", "HTTP_COOKIE", _
    "HTTP_DNT", "HTTP_EXPECT", "HTTP_FORWARDED", "HTTP_FROM", "HTTP_HOST", _
    "HTTP_IF_MATCH", "HTTP_IF_MODIFIED_SINCE", "HTTP_IF_NONE_MATCH", "HTTP_IF_UNMODIFIED_SINCE", _
    "HTTP_MAX_FORWARDS", "HTTP_PRAGMA", "HTTP_PROXY_AUTHORIZATION", "HTTP_RANGE", _
    "HTTP_REFERER", "HTTP_TE", "HTTP_TRAILER", "HTTP_TRANSFER_ENCODING", _
    "HTTP_UPGRADE", "HTTP_USER_AGENT", "HTTP_VIA", "HTTP_WARNING")

Dim svDictTmp, svNames
Set svDictTmp = CreateObject("Scripting.Dictionary")
svDictTmp.CompareMode = 1
For i = 0 To UBound(stdSvNames)
    If Not svDictTmp.Exists(stdSvNames(i)) Then svDictTmp.Add stdSvNames(i), 1
Next
' Cabeceras HTTP reales de la peticion -> nombres HTTP_ de IIS
Dim hdrKeys, hk
hdrKeys = dictHeaders.Keys
For i = 0 To UBound(hdrKeys)
    hk = "HTTP_" & UCase(Replace(CStr(hdrKeys(i)), "-", "_"))
    If Not svDictTmp.Exists(hk) Then svDictTmp.Add hk, 1
Next
svNames = svDictTmp.Keys
Set svDictTmp = Nothing

Dim dictSv
Set dictSv = CreateObject("Scripting.Dictionary")
dictSv.CompareMode = 1
For i = 0 To UBound(svNames)
    nm = CStr(svNames(i))
    If nm <> "" Then
        If Not dictSv.Exists(nm) Then dictSv.Add nm, SvValue(nm)
    End If
Next

' ---------------------------------------------------------------------------
' RESUMEN / DATOS CLAVE
' ---------------------------------------------------------------------------
Function SvValue(name)
    Dim v
    v = ""
    On Error Resume Next
    Err.Clear
    v = Request.ServerVariables(name)
    Err.Clear
    On Error GoTo 0
    SvValue = SafeStr(v)
End Function

Dim infoServer, infoScript, infoOs, infoApp, infoRemote, infoSoftware, infoProtocol, infoMethod
infoServer   = SvValue("SERVER_NAME")
infoScript   = SvValue("SCRIPT_NAME")
infoOs       = SvValue("OS")
infoApp      = SvValue("APPL_PHYSICAL_PATH")
infoRemote   = SvValue("REMOTE_ADDR")
infoSoftware = SvValue("SERVER_SOFTWARE")
infoProtocol = SvValue("SERVER_PROTOCOL")
infoMethod   = SvValue("REQUEST_METHOD")

' Arreglos ordenados de claves para cada seccion
Dim keysProc, keysMachine, keysUser, keysSv, keysHeaders, keysApp, keysSession
keysProc    = SortNames(dictProc.Keys)
keysMachine = SortNames(dictMachine.Keys)
keysUser    = SortNames(dictUser.Keys)
keysSv      = SortNames(dictSv.Keys)
keysHeaders = SortNames(dictHeaders.Keys)
keysApp     = SortNames(dictApp.Keys)
keysSession = SortNames(dictSession.Keys)
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic - Variables de Entorno</title>
  <link rel="stylesheet" href="css/style.css">
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap" rel="stylesheet">
  <style>
    body { font-family: 'Plus Jakarta Sans', sans-serif; }

    .env-summary-grid {
      display: grid;
      grid-template-columns: repeat(auto-fill, minmax(230px, 1fr));
      gap: 14px;
      margin-bottom: 24px;
    }
    .env-summary-card {
      background: var(--bg-card);
      border: 1px solid var(--border-color);
      border-radius: var(--radius-md);
      padding: 14px 16px;
      backdrop-filter: blur(12px);
    }
    .env-summary-card .env-summary-label {
      display: block;
      font-size: 10px;
      letter-spacing: 0.8px;
      text-transform: uppercase;
      color: var(--text-dim);
      font-weight: 700;
      margin-bottom: 6px;
    }
    .env-summary-card .env-summary-value {
      font-size: 13px;
      color: var(--primary-light);
      font-weight: 600;
      word-break: break-all;
      font-family: ui-monospace, Consolas, "Courier New", monospace;
    }

    .env-toolbar {
      display: flex;
      flex-wrap: wrap;
      gap: 12px;
      align-items: center;
      justify-content: space-between;
      margin-bottom: 18px;
    }
    .env-search-wrap { position: relative; flex: 1; min-width: 240px; max-width: 420px; }
    .env-search-wrap input { padding-left: 34px; }
    .env-search-wrap svg {
      position: absolute;
      left: 11px; top: 50%;
      transform: translateY(-50%);
      color: var(--text-dim);
      pointer-events: none;
    }

    .env-tabs {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
      margin-bottom: 18px;
    }
    .env-tab {
      display: inline-flex;
      align-items: center;
      gap: 8px;
      padding: 9px 14px;
      border-radius: var(--radius-full);
      border: 1px solid var(--border-color);
      background: var(--bg-card);
      color: var(--text-muted);
      font-size: 12.5px;
      font-weight: 600;
      transition: var(--transition);
    }
    .env-tab:hover { border-color: var(--border-hover); color: var(--text-main); }
    .env-tab.active {
      background: var(--primary-accent);
      border-color: var(--primary-accent);
      color: #fff;
      box-shadow: 0 4px 18px var(--primary-glow);
    }
    .env-tab-count {
      background: rgba(0,0,0,0.28);
      border-radius: var(--radius-full);
      padding: 1px 8px;
      font-size: 11px;
      font-weight: 700;
    }

    .env-panel { display: none; }
    .env-panel.active { display: block; }

    .env-var {
      font-family: ui-monospace, Consolas, "Courier New", monospace;
      font-size: 12.5px;
      color: var(--accent-teal);
      font-weight: 600;
      white-space: nowrap;
      vertical-align: top;
    }
    .env-val {
      font-family: ui-monospace, Consolas, "Courier New", monospace;
      font-size: 12.5px;
      color: var(--text-main);
      white-space: pre-wrap;
      word-break: break-all;
      max-width: 900px;
    }
    .env-val.empty { color: var(--text-dim); font-style: italic; }
    .env-row:hover { background: var(--bg-card-hover); }
    .env-empty-row td { text-align: center; padding: 30px; color: var(--text-muted); }
    .env-footnote {
      font-size: 11.5px;
      color: var(--text-dim);
      margin-top: 10px;
      line-height: 1.6;
    }
    .env-footnote code {
      background: var(--bg-elevated);
      padding: 1px 5px;
      border-radius: 4px;
      color: var(--primary-light);
    }
  </style>
</head>
<body>
  <div class="app-container">
    <!--#include file="includes/sidebar.asp"-->

    <div class="main-wrapper">
      <!--#include file="includes/topbar.asp"-->

      <div class="content-body">

        <div class="section-header">
          <div class="section-title-wrap">
            <span class="section-subtitle">Diagn&oacute;stico del servidor</span>
            <h1 class="section-title">Variables de Entorno</h1>
          </div>
          <a href="env_vars.asp" class="btn btn-secondary btn-sm">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor"><path d="M17.65 6.35A7.958 7.958 0 0 0 12 4c-4.42 0-7.99 3.58-7.99 8s3.57 8 7.99 8c3.73 0 6.84-2.55 7.73-6h-2.08A5.99 5.99 0 0 1 12 18c-3.31 0-6-2.69-6-6s2.69-6 6-6c1.66 0 3.14.69 4.22 1.78L13 11h7V4l-2.35 2.35z"/></svg>
            <span>Actualizar</span>
          </a>
        </div>

        <% If envError <> "" Then %>
        <div class="alert alert-danger">
          <svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
          <span><%=envError%></span>
        </div>
        <% End If %>

        <!-- RESUMEN DEL ENTORNO -->
        <div class="env-summary-grid">
          <div class="env-summary-card">
            <span class="env-summary-label">Servidor</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoServer)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Software</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoSoftware)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Sistema Operativo</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoOs)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Protocolo / Metodo</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoProtocol)%> &middot; <%=Server.HTMLEncode(infoMethod)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Script Actual</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoScript)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Ruta Fisica App</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoApp)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">IP del Cliente</span>
            <span class="env-summary-value"><%=Server.HTMLEncode(infoRemote)%></span>
          </div>
          <div class="env-summary-card">
            <span class="env-summary-label">Total de Variables</span>
            <span class="env-summary-value"><%=(dictProc.Count + dictMachine.Count + dictUser.Count + dictSv.Count + dictHeaders.Count + dictApp.Count + dictSession.Count)%></span>
          </div>
        </div>

        <!-- BARRA DE HERRAMIENTAS -->
        <div class="env-toolbar">
          <div class="env-search-wrap">
            <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor">
              <path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/>
            </svg>
            <input type="text" id="envSearch" class="form-control" placeholder="Filtrar variables por nombre o valor...">
          </div>
          <div class="env-search-wrap" style="max-width:none;">
            <span class="env-footnote" id="envFilterInfo">Sin filtro activo</span>
          </div>
        </div>

        <!-- PESTANAS -->
        <div class="env-tabs" id="envTabs">
          <button type="button" class="env-tab active" data-tab="tab-process">
            <span>Proceso (PROCESS)</span>
            <span class="env-tab-count"><%=dictProc.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-machine">
            <span>M&aacute;quina (MACHINE)</span>
            <span class="env-tab-count"><%=dictMachine.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-user">
            <span>Usuario (USER)</span>
            <span class="env-tab-count"><%=dictUser.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-sv">
            <span>ServerVariables</span>
            <span class="env-tab-count"><%=dictSv.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-headers">
            <span>Headers HTTP</span>
            <span class="env-tab-count"><%=dictHeaders.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-app">
            <span>Application</span>
            <span class="env-tab-count"><%=dictApp.Count%></span>
          </button>
          <button type="button" class="env-tab" data-tab="tab-session">
            <span>Session</span>
            <span class="env-tab-count"><%=dictSession.Count%></span>
          </button>
        </div>

        <%
        ' Helper de renderizado de cada seccion
        Sub RenderRows(keys, dict)
            Dim i, nm, vl
            If UBound(keys) < 0 Then
                Response.Write "<tr class='env-empty-row'><td colspan='2'>No hay variables en esta secci&oacute;n.</td></tr>"
            Else
                For i = 0 To UBound(keys)
                    nm = CStr(keys(i))
                    vl = dict(nm)
                    Response.Write "<tr class='env-row'>"
                    Response.Write "<td class='env-var' title='" & Server.HTMLEncode(nm) & "'>" & Server.HTMLEncode(nm) & "</td>"
                    If Len(vl) = 0 Then
                        Response.Write "<td class='env-val empty'>(vac&iacute;o)</td>"
                    Else
                        Response.Write "<td class='env-val' title='" & Server.HTMLEncode(vl) & "'>" & Server.HTMLEncode(vl) & "</td>"
                    End If
                    Response.Write "</tr>"
                Next
            End If
        End Sub
        %>

        <!-- PANEL: VARIABLES DEL PROCESO -->
        <div class="env-panel active" id="tab-process">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Variable de entorno del proceso</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysProc, dictProc %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Obtenidas con <code>WScript.Shell.Environment("PROCESS")</code>: son las variables que tiene asignadas
            el proceso web (IIS) en este momento.Incluyen las heredadas del sistema y del usuario.
          </p>
        </div>

        <!-- PANEL: VARIABLES DE MAQUINA -->
        <div class="env-panel" id="tab-machine">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Variable de entorno del sistema</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysMachine, dictMachine %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Leidas del registro de Windows con <code>WScript.Shell.RegRead</code> sobre la clave
            <code>HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment</code>.
            Nota: el ambito <code>Environment("MACHINE")</code> de WSH no es accesible desde el sandbox de ASP clasico.
          </p>
        </div>

        <!-- PANEL: VARIABLES DE USUARIO -->
        <div class="env-panel" id="tab-user">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Variable de entorno del usuario</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysUser, dictUser %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Obtenidas con <code>WScript.Shell.Environment("USER")</code>: variables del perfil del usuario con el que
            se ejecuta el proceso web.
          </p>
        </div>

        <!-- PANEL: SERVERVARIABLES -->
        <div class="env-panel" id="tab-sv">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">ServerVariable</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysSv, dictSv %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Colecci&oacute;n <code>Request.ServerVariables</code>: informaci&oacute;n del servidor, del script y de la
            petici&oacute;n HTTP (incluye las cabeceras <code>HTTP_*</code>).
          </p>
        </div>

        <!-- PANEL: HEADERS -->
        <div class="env-panel" id="tab-headers">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Cabecera HTTP</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysHeaders, dictHeaders %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Cabeceras HTTP de la petici&oacute;n, obtenidas de <code>Request.ServerVariables("ALL_RAW")</code>.
            La colecci&oacute;n <code>Request.Headers</code> no est&aacute; disponible en ASP cl&aacute;sico.
          </p>
        </div>

        <!-- PANEL: APPLICATION -->
        <div class="env-panel" id="tab-app">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Variable de Application</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysApp, dictApp %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Colecci&oacute;n <code>Application.Contents</code>: variables compartidas por toda la aplicaci&oacute;n web
            (visibles para todos los usuarios).
          </p>
        </div>

        <!-- PANEL: SESSION -->
        <div class="env-panel" id="tab-session">
          <div class="song-table-container">
            <table class="song-table">
              <thead>
                <tr><th style="width:340px;">Variable de Session</th><th>Valor</th></tr>
              </thead>
              <tbody>
                <% RenderRows keysSession, dictSession %>
              </tbody>
            </table>
          </div>
          <p class="env-footnote">
            Colecci&oacute;n <code>Session.Contents</code>: variables de la sesi&oacute;n actual del usuario.
          </p>
        </div>

      </div>
    </div>
  </div>

  <!--#include file="includes/player_bar.asp"-->

  <script>
  (function () {
    var tabs = document.querySelectorAll('.env-tab');
    var panels = document.querySelectorAll('.env-panel');
    var search = document.getElementById('envSearch');
    var info = document.getElementById('envFilterInfo');

    function activateTab(tabId) {
      var i;
      for (i = 0; i < tabs.length; i++) {
        tabs[i].className = tabs[i].className.replace(/\s*active/g, '');
      }
      for (i = 0; i < panels.length; i++) {
        panels[i].className = panels[i].className.replace(/\s*active/g, '');
      }
      var target = document.getElementById(tabId);
      if (target) target.className += ' active';
      for (i = 0; i < tabs.length; i++) {
        if (tabs[i].getAttribute('data-tab') === tabId) {
          tabs[i].className += ' active';
        }
      }
    }

    function applyFilter() {
      var term = (search.value || '').toLowerCase().trim();
      var totalVisible = 0, totalAll = 0, p, r;
      for (p = 0; p < panels.length; p++) {
        var rows = panels[p].querySelectorAll('.env-row');
        var visible = 0;
        for (r = 0; r < rows.length; r++) {
          var text = (rows[r].textContent || '').toLowerCase();
          if (term === '' || text.indexOf(term) !== -1) {
            rows[r].style.display = '';
            visible++;
          } else {
            rows[r].style.display = 'none';
          }
        }
        totalVisible += visible;
        totalAll += rows.length;
      }
      if (term === '') {
        info.textContent = 'Mostrando ' + totalAll + ' variables en total';
      } else {
        info.textContent = 'Mostrando ' + totalVisible + ' de ' + totalAll + ' variables que coinciden con "' + search.value + '"';
      }
    }

    for (var i = 0; i < tabs.length; i++) {
      tabs[i].onclick = function () {
        activateTab(this.getAttribute('data-tab'));
      };
    }

    if (search) {
      search.oninput = applyFilter;
      search.onkeyup = applyFilter;
    }

    applyFilter();
  })();
  </script>
</body>
</html>
