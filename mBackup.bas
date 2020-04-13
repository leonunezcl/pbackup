Attribute VB_Name = "mBackup"
Option Explicit


Private cTLI As TypeLibInfo
Public cRegistro As New cRegistry
Private nFreeFile As Long
Private PathProyecto As String
Private KeyRegistro As String
Private PathRegistro As String
Private sGUID As String
Private sArchivo As String
Private nLinea As Integer
Private Linea As String

Private REF_DLL As Integer
Private REF_OCX As Integer
Private REF_RES As Integer

Private MayorV As Variant 'As Integer
Private MenorV As Variant 'As Integer
Private P1 As Integer
Private P2 As Integer
Private LineaPaso As String

Public Enum eTipoDepencia
    TIPO_DLL = 1
    TIPO_OCX = 2
    TIPO_RES = 3
    TIPO_PAGE = 4
End Enum

Enum eTipoArchivo
    TIPO_ARCHIVO_FRM = 1
    TIPO_ARCHIVO_BAS = 2
    TIPO_ARCHIVO_CLS = 3
    TIPO_ARCHIVO_OCX = 4
    TIPO_ARCHIVO_PAG = 5
    TIPO_ARCHIVO_REL = 6
    TIPO_ARCHIVO_FRX = 7
    TIPO_ARCHIVO_RPT = 8
    TIPO_ARCHIVO_ICO = 9
    TIPO_ARCHIVO_GIF = 10
    TIPO_ARCHIVO_INI = 11
    TIPO_ARCHIVO_HLP = 12
    TIPO_ARCHIVO_BMP = 13
End Enum

Public Enum eTipoProyecto
    PRO_TIPO_NONE = 0
    PRO_TIPO_EXE = 1
    PRO_TIPO_DLL = 2
    PRO_TIPO_OCX = 3
    PRO_TIPO_EXE_X = 4
End Enum

Public Type eDependencias
    Tipo As eTipoDepencia
    Archivo As String
    GUID As String
    KeyNode As String
    Name As String
    ContainingFile As String
    HelpString As String
    HelpFile As String
    MajorVersion As Long
    MinorVersion As Long
    FileSize As Long
    FILETIME As String
End Type

Private Type eDatos
    OptionExplicit As Boolean
    Explorar As Boolean
    Nombre As String
    PathFisico As String
    FileSize As Double
    FILETIME As String
    ObjectName As String
    Descripcion As String
    TipoDeArchivo As eTipoArchivo
    KeyNode As String
End Type

Public Type eProyecto
    Nombre As String
    Archivo As String
    Icono As Integer
    Version As Integer
    PathFisico As String
    TipoProyecto As eTipoProyecto
    FileSize As Double
    FILETIME As String
    aArchivos() As eDatos
    aDepencias() As eDependencias
End Type
Public Proyecto As eProyecto

'Private Const MAXDWORD = &HFFFF
Private Const INVALID_HANDLE_VALUE = -1
'Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
'Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
'Private Const FILE_ATTRIBUTE_HIDDEN = &H2
'Private Const FILE_ATTRIBUTE_NORMAL = &H80
'Private Const FILE_ATTRIBUTE_READONLY = &H1
'Private Const FILE_ATTRIBUTE_SYSTEM = &H4
'Private Const FILE_ATTRIBUTE_TEMPORARY = &H100

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
'Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Const MAX_PATH = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'busca archivos en un directorio segun criterio
Private Sub BuscarArchivos(k As Integer, ByVal Path As String, ByVal SearchStr As String, ByVal Icono As Integer)
    
    Dim FileName As String
    Dim DirName As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    
    If Right$(Path, 1) <> "\" Then Path = Path & "\"
    
    ' Search for subdirectories.
    nDir = 0
    
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                
                If Icono = 1 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_RPT)
                ElseIf Icono = 2 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_ICO)
                ElseIf Icono = 3 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_GIF)
                ElseIf Icono = 4 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_INI)
                ElseIf Icono = 5 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_BMP)
                ElseIf Icono = 18 Then
                    Call AgregaArchivoDeProyecto(k, FileName, TIPO_ARCHIVO_HLP)
                End If
                                
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    
End Sub
'agrega la referencias
Private Sub AgregaReferencias(d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2 + 2, Linea, "#") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
    If Right$(MenorV, 1) = "#" Then
        MenorV = Left$(MenorV, Len(MenorV) - 1)
    End If
    
    KeyRegistro = Mid$(Linea, InStr(1, Linea, "G") + 1)
    KeyRegistro = Left$(KeyRegistro, InStr(1, KeyRegistro, "}"))
                    
    cRegistro.ClassKey = HKEY_CLASSES_ROOT
    cRegistro.ValueType = REG_SZ
    cRegistro.SectionKey = "TypeLib\" & KeyRegistro & "\" & Val(MayorV) & "\" & Val(MenorV) & "\win32"
    sArchivo = cRegistro.Value
    
    If sArchivo = "" Then
        sArchivo = NombreArchivo(Linea, 1)
    End If
            
    Set cTLI = TLI.TypeLibInfoFromRegistry(KeyRegistro, Val(MayorV), Val(MenorV), 0)
    
    If Err.Number <> 0 Then
        Err = 0
        Set cTLI = TLI.TypeLibInfoFromFile(sArchivo)
    
        If Err.Number <> 0 Then
            MsgBox "Error al cargar información de referencia : " & vbNewLine & sArchivo, vbCritical
        Else
            ReDim Preserve Proyecto.aDepencias(d)
        
            Proyecto.aDepencias(d).Archivo = sArchivo
            
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = cTLI.GUID
            Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
            Proyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
            REF_DLL = REF_DLL + 1
            d = d + 1
        End If
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        If Proyecto.Version > 3 Or Proyecto.Version = 0 Then
            Proyecto.aDepencias(d).Archivo = sArchivo
            Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
            Proyecto.aDepencias(d).HelpString = cTLI.HelpString
            Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
            Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
            Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = cTLI.GUID
        Else
            Proyecto.aDepencias(d).Archivo = Linea
            Proyecto.aDepencias(d).ContainingFile = Linea
            Proyecto.aDepencias(d).HelpString = ""
            Proyecto.aDepencias(d).HelpFile = 0
            Proyecto.aDepencias(d).MajorVersion = 0
            Proyecto.aDepencias(d).MinorVersion = 0
            Proyecto.aDepencias(d).Tipo = TIPO_DLL
            Proyecto.aDepencias(d).GUID = ""
        End If
        
        Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).KeyNode = "REFDLL" & REF_DLL
        REF_DLL = REF_DLL + 1
        d = d + 1
    End If
    
    Err = 0
                
End Sub

Private Function NombreArchivo(ByVal sLinea As String, ByVal Leer As Integer) As String

    Dim k As Integer
    Dim ret As String
    Dim Inicio As Integer
    
    Inicio = 0
    
    If Leer = 1 Then        'REFERENCIAS
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = "#" Then
                If Inicio = 0 Then
                    Inicio = k
                Else
                    ret = Mid$(sLinea, k + 1, Inicio - (k + 1))
                    Exit For
                End If
            End If
        Next k
    ElseIf Leer = 2 Then    'CONTROLES
        For k = Len(sLinea) To 1 Step -1
            If Mid$(sLinea, k, 1) = ";" Then
                Inicio = k
                ret = Trim$(Mid$(sLinea, Inicio + 1))
                Exit For
            End If
        Next k
    End If
    
    NombreArchivo = ret
    
End Function

Public Function CargaProyecto() As Boolean

    Dim ret As Boolean
    Dim Archivo As String
    
    Dim f As Integer
    Dim M As Integer
    Dim c As Integer
    Dim k As Integer
    Dim r As Integer
    Dim i As Integer
    Dim j As Integer
    Dim d As Integer
        
    Dim Formulario As String
    Dim Modulo As String
    Dim ControlUsuario As String
    Dim Clase As String
    Dim Referencia As String
    Dim RefRes As String
    Dim PagPropiedades As String
    Dim DoctosRelacionados As String
    Dim Path As String
    
    ret = True
    
    frmMain.lvwArchivos.ListItems.Clear
            
    Archivo = StripNulls(glbArchivo)
    
    PathProyecto = PathArchivo(Archivo)
    
    nFreeFile = FreeFile
        
    'determinar el tipo de proyecto
    If Not DeterminaTipoDeProyecto(Archivo) Then
        ret = False
        GoTo SalirCargaProyecto
    End If
    
    Proyecto.PathFisico = Archivo
    Proyecto.FILETIME = VBGetFileTime(Archivo)
    
    frmMain.Caption = App.Title & " - " & VBArchivoSinPath(Proyecto.PathFisico)
    
    ReDim Proyecto.aArchivos(0)
    ReDim Proyecto.aDepencias(0)
    
    nFreeFile = FreeFile
    
    REF_DLL = 1
    REF_OCX = 1
    REF_RES = 1
    k = 1
    
    'determinar los diferentes archivos que componen el proyecto
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If InStr(Linea, "Form=") <> 0 Then          'FORMULARIOS
                If InStr(Linea, "IconForm=") = 0 Then
                    Formulario = Mid$(Linea, InStr(Linea, "=") + 1)
                    
                    Call AgregaArchivoDeProyecto(k, Formulario, TIPO_ARCHIVO_FRM)
                End If
            ElseIf InStr(Linea, "Module=") <> 0 Then    'MODULOS
                Modulo = Mid$(Linea, InStr(Linea, "=") + 1)
                Modulo = Trim$(Mid$(Modulo, InStr(Modulo, ";") + 1))
                
                Call AgregaArchivoDeProyecto(k, Modulo, TIPO_ARCHIVO_BAS)
            ElseIf InStr(Linea, "UserControl=") <> 0 Then   'CONTROLES
                ControlUsuario = Mid$(Linea, InStr(Linea, "=") + 1)
                Call AgregaArchivoDeProyecto(k, ControlUsuario, TIPO_ARCHIVO_OCX)
            ElseIf InStr(Linea, "Class=") <> 0 Then         'MODULOS DE CLASE
                Clase = Mid$(Linea, InStr(Linea, "=") + 1)
                Clase = Trim$(Mid$(Clase, InStr(Clase, ";") + 1))
                                                
                Call AgregaArchivoDeProyecto(k, Clase, TIPO_ARCHIVO_CLS)
            ElseIf InStr(Linea, "Reference=") <> 0 Then     'REFERENCIAS
                Call AgregaReferencias(d, Linea)
            ElseIf InStr(Linea, "Object=") <> 0 Then        'CONTROLES
                If Left$(Linea, 6) = "Object" Then
                    Call AgregaComponentes(d, Linea)
                End If
            ElseIf InStr(Linea, "ResFile32=") <> 0 Then
                RefRes = Trim$(Mid$(Linea, InStr(Linea, """") + 1))
                RefRes = Left$(RefRes, Len(RefRes) - 1)
                
                ReDim Preserve Proyecto.aDepencias(d)
                
                'CHEQUEAR \
                If PathArchivo(RefRes) = "" Then
                    Proyecto.aDepencias(d).Archivo = PathProyecto & RefRes
                Else
                    Proyecto.aDepencias(d).Archivo = PathArchivo(RefRes)
                End If
                
                Proyecto.aDepencias(d).Tipo = TIPO_RES
                Proyecto.aDepencias(d).KeyNode = "REFRES" & REF_RES
                Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
                Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
                REF_RES = REF_RES + 1
                d = d + 1
            ElseIf InStr(Linea, "PropertyPage=") <> 0 Then  'Pagina de propiedades
                PagPropiedades = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDeProyecto(k, PagPropiedades, TIPO_ARCHIVO_PAG)
                
            ElseIf InStr(Linea, "RelatedDoc=") <> 0 Then  'Documentos Relacionados
                DoctosRelacionados = Mid$(Linea, InStr(Linea, "=") + 1)
                
                Call AgregaArchivoDeProyecto(k, DoctosRelacionados, TIPO_ARCHIVO_REL)
            ElseIf Right$(Linea, 3) = "FRM" Then 'para versiones anteriores de VB3
                Formulario = Linea
                Call AgregaArchivoDeProyecto(k, Formulario, TIPO_ARCHIVO_FRM)
            ElseIf Right$(Linea, 3) = "BAS" Then 'para versiones anteriores de VB3
                Modulo = Linea
                Call AgregaArchivoDeProyecto(k, Modulo, TIPO_ARCHIVO_BAS)
            ElseIf Right$(Linea, 3) = "VBX" Then 'para versiones anteriores de VB3
                Call AgregaReferencias(d, Linea)
            End If
        Loop
    Close #nFreeFile
    
    'verificar x archivos anexos como .frx , .vbw
    Archivo = Left$(glbArchivo, Len(Archivo) - 3)
    Archivo = Archivo & "vbw"
    If VBOpenFile(Archivo) Then
        Archivo = VBArchivoSinPath(Archivo)
        Call AgregaArchivoDeProyecto(k, Archivo, TIPO_ARCHIVO_REL)
    End If
            
    For j = 1 To UBound(Proyecto.aArchivos)
        If Proyecto.aArchivos(j).TipoDeArchivo = TIPO_ARCHIVO_FRM Then
            Archivo = Left$(Proyecto.aArchivos(j).PathFisico, Len(Proyecto.aArchivos(j).PathFisico) - 3)
            Archivo = Archivo & "frx"
            'verificar si existe el archivo
            If VBOpenFile(Archivo) Then
                Archivo = VBArchivoSinPath(Archivo)
                Call AgregaArchivoDeProyecto(k, Archivo, TIPO_ARCHIVO_FRX)
            End If
        End If
    Next j
    
    'buscar archivos
    Path = PathArchivo(glbArchivo)
    If glbRpt Then Call BuscarArchivos(k, Path, "*.rpt", 1)
    If glbIma Then Call BuscarArchivos(k, Path, "*.ico", 2)
    If glbIma Then Call BuscarArchivos(k, Path, "*.gif", 3)
    If glbIma Then Call BuscarArchivos(k, Path, "*.jpg", 3)
    If glbIma Then Call BuscarArchivos(k, Path, "*.jpeg", 3)
    If glbIni Then Call BuscarArchivos(k, Path, "*.ini", 4)
    If glbIma Then Call BuscarArchivos(k, Path, "*.bmp", 5)
    If glbHlp Then Call BuscarArchivos(k, Path, "*.hlp", 18)
    If glbHlp Then Call BuscarArchivos(k, Path, "*.chm", 18)
    
    GoTo SalirCargaProyecto
    
SalirCargaProyecto:
    CargaProyecto = ret
    
End Function


'agrega los componentes al arbol de proyecto
Private Sub AgregaComponentes(d As Integer, ByVal Linea As String)

    On Local Error Resume Next
    
    'BUSCAR MAYOR
    P1 = 0: P2 = 0
    P1 = InStr(1, Linea, "#")
    P2 = InStr(P1 + 1, Linea, "#") - 1
    MayorV = Mid$(Linea, P1 + 1, P2 - P1)
    
    'BUSCAR MENOR
    P1 = InStr(P2, Linea, ";") - 1
    MenorV = Mid$(Linea, P2 + 2, P1 - P2)
    If Right$(MenorV, 1) = ";" Then MenorV = Left$(MenorV, Len(MenorV) - 1)

    sGUID = Left$(Linea, InStr(1, Linea, "}"))
    sGUID = Mid$(sGUID, 8)
    
    If InStr(1, MayorV, ".") Then
        MenorV = Mid$(MayorV, InStr(1, MayorV, ".") + 1)
        MayorV = Left$(MayorV, InStr(1, MayorV, ".") - 1)
    End If
    
    Set cTLI = TLI.TypeLibInfoFromRegistry(sGUID, Val(MayorV), Val(MenorV), 0)
    
    If Err <> 0 Then
        MsgBox "Error al cargar información de referencia : " & vbNewLine & sArchivo, vbCritical
    Else
        ReDim Preserve Proyecto.aDepencias(d)
        
        Proyecto.aDepencias(d).Archivo = cTLI.ContainingFile
        Proyecto.aDepencias(d).ContainingFile = cTLI.ContainingFile
        Proyecto.aDepencias(d).HelpString = cTLI.HelpString
        Proyecto.aDepencias(d).HelpFile = cTLI.HelpFile
        Proyecto.aDepencias(d).MajorVersion = cTLI.MajorVersion
        Proyecto.aDepencias(d).MinorVersion = cTLI.MinorVersion
        Proyecto.aDepencias(d).GUID = cTLI.GUID
        Proyecto.aDepencias(d).Tipo = TIPO_OCX
        Proyecto.aDepencias(d).FileSize = VBGetFileSize(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).FILETIME = VBGetFileTime(Proyecto.aDepencias(d).Archivo)
        Proyecto.aDepencias(d).KeyNode = "REFOCX" & REF_OCX
        REF_OCX = REF_OCX + 1
        d = d + 1
    End If
    
    Err = 0
                    
End Sub

'agrega el archivo de proyecto a estructura
Private Sub AgregaArchivoDeProyecto(ByRef k As Integer, ByVal Archivo As String, _
                                    ByVal Tipo As eTipoArchivo)

    ReDim Preserve Proyecto.aArchivos(k)
                    
    'CHEQUEAR \
    If PathArchivo(Archivo) = "" Then
        Proyecto.aArchivos(k).Nombre = Archivo
        Proyecto.aArchivos(k).PathFisico = PathProyecto & Archivo
    Else
        Proyecto.aArchivos(k).Nombre = Mid$(Archivo, InStr(Archivo, "\") + 1)
        Proyecto.aArchivos(k).PathFisico = PathProyecto & Archivo
    End If
    
    Proyecto.aArchivos(k).TipoDeArchivo = Tipo
    
    If Tipo = TIPO_ARCHIVO_FRM Then
        Proyecto.aArchivos(k).KeyNode = "FRM" & k
    ElseIf Tipo = TIPO_ARCHIVO_BAS Then
        Proyecto.aArchivos(k).KeyNode = "BAS" & k
    ElseIf Tipo = TIPO_ARCHIVO_CLS Then
        Proyecto.aArchivos(k).KeyNode = "CLS" & k
    ElseIf Tipo = TIPO_ARCHIVO_OCX Then
        Proyecto.aArchivos(k).KeyNode = "OCX" & k
    ElseIf Tipo = TIPO_ARCHIVO_PAG Then
        Proyecto.aArchivos(k).KeyNode = "PAG" & k
    ElseIf Tipo = TIPO_ARCHIVO_REL Then
        Proyecto.aArchivos(k).KeyNode = "REL" & k
    ElseIf Tipo = TIPO_ARCHIVO_FRX Then
        Proyecto.aArchivos(k).KeyNode = "FRX" & k
    ElseIf Tipo = TIPO_ARCHIVO_HLP Then
        Proyecto.aArchivos(k).KeyNode = "HLP" & k
    ElseIf Tipo = TIPO_ARCHIVO_RPT Then
        Proyecto.aArchivos(k).KeyNode = "RPT" & k
    ElseIf Tipo = TIPO_ARCHIVO_ICO Then
        Proyecto.aArchivos(k).KeyNode = "ICO" & k
    ElseIf Tipo = TIPO_ARCHIVO_INI Then
        Proyecto.aArchivos(k).KeyNode = "INI" & k
    ElseIf Tipo = TIPO_ARCHIVO_BMP Then
        Proyecto.aArchivos(k).KeyNode = "BMP" & k
    End If
    
    Proyecto.aArchivos(k).FILETIME = VBGetFileTime(Proyecto.aArchivos(k).PathFisico)
    Proyecto.aArchivos(k).FileSize = VBGetFileSize(Proyecto.aArchivos(k).PathFisico)
    
    Proyecto.aArchivos(k).Explorar = True
    
    k = k + 1
                    
End Sub

'determina el tipo de proyecto
Private Function DeterminaTipoDeProyecto(ByVal Archivo As String) As Boolean

    On Local Error GoTo ErrorDeterminaTipoDeProyecto
    
    Dim ret As Boolean
    Dim Icono As Integer
    Dim sProyecto As String
    Dim Linea As String
    Dim sNombreArchivo As String
    Dim nFreeFile As Long
    
    Icono = C_ICONO_VBP
    
    sNombreArchivo = VBArchivoSinPath(Archivo)
    
    nFreeFile = FreeFile
    
    ret = True
    
    Proyecto.TipoProyecto = PRO_TIPO_NONE
    Proyecto.Version = 0
    
    Open Archivo For Input Shared As #nFreeFile
        Do While Not EOF(nFreeFile)
            Line Input #nFreeFile, Linea
            If Left$(Linea, 4) = "Type" Then
                If InStr(Linea, "Exe") Then
                    Icono = C_ICONO_VBP
                    Proyecto.TipoProyecto = PRO_TIPO_EXE
                    Proyecto.Icono = Icono
                ElseIf InStr(Linea, "Control") Then
                    Icono = C_ICONO_OCX
                    Proyecto.TipoProyecto = PRO_TIPO_OCX
                    Proyecto.Icono = Icono
                ElseIf InStr(Linea, "OleDll") Then
                    Icono = C_ICONO_DLL
                    Proyecto.TipoProyecto = PRO_TIPO_DLL
                    Proyecto.Icono = Icono
                End If
            ElseIf Left$(Linea, 4) = "Name" Then
                sProyecto = Mid$(Linea, 6)
                sProyecto = Mid$(sProyecto, 2)
                sProyecto = Left$(sProyecto, Len(sProyecto) - 1)
                Proyecto.Nombre = sProyecto
                Proyecto.Archivo = sNombreArchivo
            End If
        Loop
    Close #nFreeFile
            
    'para versiones de visual basic que no tienen el name
    If Proyecto.TipoProyecto = PRO_TIPO_NONE Then
        Proyecto.TipoProyecto = PRO_TIPO_EXE
        Proyecto.Icono = C_ICONO_VBP
        Proyecto.Version = 3
    End If
    
    If Proyecto.Nombre = "" Then
        Proyecto.Nombre = Left$(sNombreArchivo, InStr(1, sNombreArchivo, ".") - 1)
        Proyecto.Archivo = sNombreArchivo
        Proyecto.Version = 3
    End If
    
    GoTo SalirDeterminaTipoDeProyecto
    
ErrorDeterminaTipoDeProyecto:
    ret = False
    'MsgBox "DeterminaTipoDeProyecto : " & Err & " " & Error$, vbCritical
    Resume SalirDeterminaTipoDeProyecto
    
SalirDeterminaTipoDeProyecto:
    DeterminaTipoDeProyecto = ret
    Err = 0
    
End Function

