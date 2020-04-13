Attribute VB_Name = "mMain"
Option Explicit

Public Ini As New CIni
Public gbInicio As Boolean
Public glbArchivo As String
Public glbArchivoZIP As String
Public gsLastPath As String
Public glbArchivoLng As String

'opciones de configuracion
Public glbRecursive As Boolean
Public glbSavePath As Boolean

'opciones anexas
Public glbRpt As Boolean
Public glbIma As Boolean
Public glbIni As Boolean
Public glbHlp As Boolean

Public Const C_RELEASE = "12/01/2002"
Public Const C_WEB_PAGE = "http://www.vbsoftware.cl/"
Public Const C_WEB_PAGE_PE = "http://www.vbsoftware.cl/pbackup.html"
Public Const C_EMAIL = "lnunez@vbsoftware.cl"

Public Const WM_SETREDRAW = &HB
Private Const OF_EXIST = &H4000
Private Const OFS_MAXPATHNAME = 256
Private Const IDC_WAIT = 32514&
Private Const IDC_ARROW = 32512&
'Private Const GWL_WNDPROC = (-4)
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const C_INI = "PBACKUP.INI"

Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type LOGFONT
  lfHeight As Long
  lfWidth As Long
  lfEscapement As Long
  lfOrientation As Long
  lfWeight As Long
  lfItalic As Byte
  lfUnderline As Byte
  lfStrikeOut As Byte
  lfCharSet As Byte
  lfOutPrecision As Byte
  lfClipPrecision As Byte
  lfQuality As Byte
  lfPitchAndFamily As Byte
  lfFaceName As String * 33
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum SysMet
    SM_CXSCREEN = 0
    SM_CYSCREEN = 1
    SM_CXVSCROLL = 2
    SM_CYHSCROLL = 3
    SM_CYCAPTION = 4
    SM_CXBORDER = 5
    SM_CYBORDER = 6
    SM_CXDLGFRAME = 7
    SM_CYDLGFRAME = 8
    SM_CYVTHUMB = 9
    SM_CXHTHUMB = 10
    SM_CXICON = 11
    SM_CYICON = 12
    SM_CXCURSOR = 13
    SM_CYCURSOR = 14
    SM_CYMENU = 15
    SM_CXFULLSCREEN = 16
    SM_CYFULLSCREEN = 17
    SM_CYKANJIWINDOW = 18
    SM_MOUSEPRESENT = 19
    SM_CYVSCROLL = 20
    SM_CXHSCROLL = 21
    SM_DEBUG = 22
    SM_SWAPBUTTON = 23
    SM_RESERVED1 = 24
    SM_RESERVED2 = 25
    SM_RESERVED3 = 26
    SM_RESERVED4 = 27
    SM_CXMIN = 28
    SM_CYMIN = 29
    SM_CXSIZE = 30
    SM_CYSIZE = 31
    SM_CXFRAME = 32
    SM_CYFRAME = 33
    SM_CXMINTRACK = 34
    SM_CYMINTRACK = 35
    SM_CXDOUBLECLK = 36
    SM_CYDOUBLECLK = 37
    SM_CXICONSPACING = 38
    SM_CYICONSPACING = 39
    SM_MENUDROPALIGNMENT = 40
    SM_PENWINDOWS = 41
    SM_DBCSENABLED = 42
    SM_CMOUSEBUTTONS = 43
    SM_CMETRICS = 44
End Enum

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

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, lpCursorName As Any) As Long
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function IsDebuggerPresent Lib "kernel32" () As Long

Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long 'Optional parameter
    lpClass As String 'Optional parameter
    hkeyClass As Long 'Optional parameter
    dwHotKey As Long 'Optional parameter
    hIcon As Long 'Optional parameter
    hProcess As Long 'Optional parameter
End Type

Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long
Public Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Public Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Const MF_BYPOSITION = &H400&

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Function FindFilesAPI(Path As String, SearchStr As String, FileCount As Integer, DirCount As Integer)

    Dim FileName As String
    Dim DirName As String
    Dim dirNames() As String
    Dim nDir As Integer
    Dim i As Integer
    Dim hSearch As Long
    Dim WFD As WIN32_FIND_DATA
    Dim Cont As Integer
    
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    ' Search for subdirectories.
    nDir = 0
    ReDim dirNames(nDir)
    Cont = True
    hSearch = FindFirstFile(Path & "*", WFD)
    If hSearch <> INVALID_HANDLE_VALUE Then
        Do While Cont
        DirName = StripNulls(WFD.cFileName)
        ' Ignore the current and encompassing directories.
        If (DirName <> ".") And (DirName <> "..") Then
            ' Check for directory with bitwise comparison.
            If GetFileAttributes(Path & DirName) And FILE_ATTRIBUTE_DIRECTORY Then
                dirNames(nDir) = DirName
                DirCount = DirCount + 1
                nDir = nDir + 1
                ReDim Preserve dirNames(nDir)
            End If
        End If
        Cont = FindNextFile(hSearch, WFD) 'Get next subdirectory.
        Loop
        Cont = FindClose(hSearch)
    End If
    ' Walk through this directory and sum file sizes.
    hSearch = FindFirstFile(Path & SearchStr, WFD)
    Cont = True
    If hSearch <> INVALID_HANDLE_VALUE Then
        While Cont
            FileName = StripNulls(WFD.cFileName)
            If (FileName <> ".") And (FileName <> "..") Then
                FindFilesAPI = FindFilesAPI + (WFD.nFileSizeHigh * MAXDWORD) + WFD.nFileSizeLow
                FileCount = FileCount + 1
                
                If Len(FileName) > 0 Then
                    frmMain.List1.AddItem Path & FileName
                    frmMain.stbMain.Panels(1).Text = "Agregando : " & FileName
                End If
            End If
            Cont = FindNextFile(hSearch, WFD) ' Get next file
        Wend
        Cont = FindClose(hSearch)
    End If
    ' If there are sub-directories...
    If nDir > 0 Then
        ' Recursively walk into them...
        For i = 0 To nDir - 1
            FindFilesAPI = FindFilesAPI + FindFilesAPI(Path & dirNames(i) & "\", SearchStr, FileCount, DirCount)
        Next i
    End If
End Function


Public Sub GrabaInfoFormularios(frm As Form)

    Dim nFreeFile As Integer
    Dim k As Integer
    Dim lv As Integer
    
    nFreeFile = FreeFile
    
    Open App.Path & "\spanish.lng" For Append As #nFreeFile
        Print #nFreeFile, "[" & frm.Name & "]"
        Print #nFreeFile, frm.Name & "=" & frm.Caption
        
        For k = 0 To frm.Controls.Count - 1
            If TypeOf frm.Controls(k) Is Label Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is Frame Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is Menu Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is CommandButton Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is CheckBox Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is OptionButton Then
                Print #nFreeFile, frm.Controls(k).Name & "=" & frm.Controls(k).Caption
            ElseIf TypeOf frm.Controls(k) Is ListView Then
                For lv = 1 To frm.Controls(k).ColumnHeaders.Count '- 1
                    Print #nFreeFile, frm.Controls(k).ColumnHeaders(lv).Key & "=" & frm.Controls(k).ColumnHeaders(lv).Text
                Next lv
            ElseIf TypeOf frm.Controls(k) Is Toolbar Then
                For lv = 1 To frm.Controls(k).Buttons.Count
                    Print #nFreeFile, frm.Controls(k).Buttons(lv).Key & "=" & frm.Controls(k).Buttons(lv).ToolTipText
                Next lv
            End If
        Next k
    Close #nFreeFile
    
End Sub
'carga la informacion del archivo de lenguaje a la pantalla
Public Sub CargarPantalla(frm As Form)

    Dim k As Integer
    Dim lv As Integer
    Dim Texto As String
    Dim Seccion As String
    
    Seccion = frm.Name
        
    frm.Caption = Ini.Leer(Seccion, frm.Name, glbArchivoLng)
    
    For k = 0 To frm.Controls.Count - 1
        If TypeOf frm.Controls(k) Is Label Then
            frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
        ElseIf TypeOf frm.Controls(k) Is Frame Then
            frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
        ElseIf TypeOf frm.Controls(k) Is Menu Then
            If frm.Controls(k).Caption <> "-" Then
                frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
            End If
        ElseIf TypeOf frm.Controls(k) Is CommandButton Then
            frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
        ElseIf TypeOf frm.Controls(k) Is CheckBox Then
            frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
        ElseIf TypeOf frm.Controls(k) Is OptionButton Then
            frm.Controls(k).Caption = Ini.Leer(Seccion, frm.Controls(k).Name, glbArchivoLng)
        ElseIf TypeOf frm.Controls(k) Is ListView Then
            For lv = 1 To frm.Controls(k).ColumnHeaders.Count
                frm.Controls(k).ColumnHeaders(lv).Text = Ini.Leer(Seccion, frm.Controls(k).ColumnHeaders(lv).Key, glbArchivoLng)
            Next lv
        ElseIf TypeOf frm.Controls(k) Is Toolbar Then
            For lv = 1 To frm.Controls(k).Buttons.Count
                frm.Controls(k).Buttons(lv).ToolTipText = Ini.Leer(Seccion, frm.Controls(k).Buttons(lv).Key, glbArchivoLng)
            Next lv
        End If
    Next k
        
End Sub

'remueve la x
Public Sub RemoveMenus(frm As Form, _
    remove_restore As Boolean, _
    remove_move As Boolean, _
    remove_size As Boolean, _
    remove_minimize As Boolean, _
    remove_maximize As Boolean, _
    remove_seperator As Boolean, _
    remove_close As Boolean)
Dim hMenu As Long
    
    ' Get the form's system menu handle.
    hMenu = GetSystemMenu(frm.hWnd, False)
    
    If remove_close Then DeleteMenu hMenu, 6, MF_BYPOSITION
    If remove_seperator Then DeleteMenu hMenu, 5, MF_BYPOSITION
    If remove_maximize Then DeleteMenu hMenu, 4, MF_BYPOSITION
    If remove_minimize Then DeleteMenu hMenu, 3, MF_BYPOSITION
    If remove_size Then DeleteMenu hMenu, 2, MF_BYPOSITION
    If remove_move Then DeleteMenu hMenu, 1, MF_BYPOSITION
    If remove_restore Then DeleteMenu hMenu, 0, MF_BYPOSITION
End Sub

'verifica si un archivo existe
Public Function VBOpenFile(ByVal archivo As String) As Boolean

    Dim ret As Boolean
    Dim lret As Long
    Dim of As OFSTRUCT
    
    ret = False
    
    lret = OpenFile(archivo, of, OF_EXIST)
    
    If of.nErrCode = 0 Then ret = True
    
    VBOpenFile = ret
    
End Function
Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Public Function PathArchivo(ByVal archivo As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(archivo) To 1 Step -1
        If Mid$(archivo, k, 1) = "\" Then
            ret = Mid$(archivo, 1, k)
            Exit For
        End If
    Next k
    
    PathArchivo = ret
    
End Function

Public Function VBArchivoSinPath(ByVal ArchivoConPath As String) As String

    Dim k As Integer
    
    Dim ret As String
    
    ret = ""
    
    For k = Len(ArchivoConPath) To 1 Step -1
        If Mid$(ArchivoConPath, k, 1) = "\" Then
            ret = Mid$(ArchivoConPath, k + 1)
            Exit For
        End If
    Next k
    
    VBArchivoSinPath = ret
    
End Function

Public Function Confirma(ByVal Msg As String) As Integer
    Confirma = MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2)
End Function
Public Sub Hourglass(hWnd As Long, fOn As Boolean)

    If fOn Then
        Call SetCapture(hWnd)
        Call SetCursor(LoadCursor(0, ByVal IDC_WAIT))
    Else
        Call ReleaseCapture
        Call SetCursor(LoadCursor(0, IDC_ARROW))
    End If
    DoEvents
    
End Sub
Sub CenterWindow(ByVal hWnd As Long)

    Dim wRect As RECT
    
    Dim x As Integer
    Dim y As Integer

    Dim ret As Long
    
    ret = GetWindowRect(hWnd, wRect)
    
    x = (GetSystemMetrics(SM_CXSCREEN) - (wRect.Right - wRect.Left)) / 2
    y = (GetSystemMetrics(SM_CYSCREEN) - (wRect.Bottom - wRect.Top + GetSystemMetrics(SM_CYCAPTION))) / 2
    
    ret = SetWindowPos(hWnd, vbNull, x, y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER)
    
End Sub
'obtener tamaño del archivo
Public Function VBGetFileSize(ByVal archivo As String) As Double

    Dim lngHandle As Long
    Dim lret As Double
    Dim ret As Long
    Dim of As OFSTRUCT
    
    lngHandle = OpenFile(archivo, of, 0&)
    lret = GetFileSize(lngHandle, ret)
    CloseHandle lngHandle
    
    VBGetFileSize = Round((lret / 1024), 2)
    
End Function
'obtener la fecha de creacion del archivo
Public Function VBGetFileTime(ByVal archivo As String) As String

    Dim ret As String
    Dim lngHandle As Long
    Dim Ft1 As FILETIME, Ft2 As FILETIME, SysTime As SYSTEMTIME
    Dim Fecha As String
    Dim Hora As String
    
    Dim of As OFSTRUCT
    
    lngHandle = OpenFile(archivo, of, 0&)
    
    GetFileTime lngHandle, Ft1, Ft1, Ft2
    
    FileTimeToLocalFileTime Ft2, Ft1
    
    FileTimeToSystemTime Ft1, SysTime
    
    CloseHandle lngHandle
    
    Fecha = Format(Trim(Str$(SysTime.wDay)), "00") & "/" & Format(Trim$(Str$(SysTime.wMonth)), "00") + "/" + LTrim(Str$(SysTime.wYear))
    Hora = Format(Trim(Str$(SysTime.wHour)), "00") & ":" & Format(Trim$(Str$(SysTime.wMinute)), "00") + ":" + LTrim(Str$(SysTime.wSecond))
    
    VBGetFileTime = Fecha & " " & Hora
    
End Function

Public Sub Shell_Email()

    On Local Error Resume Next
    ShellExecute frmMain.hWnd, vbNullString, "mailto:lnunez@vbsoftware.cl", vbNullString, "C:\", 3
    Err = 0
    
End Sub

Public Sub Shell_PaginaWeb()

    On Local Error Resume Next
    ShellExecute frmMain.hWnd, vbNullString, "http://www.vbsoftware.cl/", vbNullString, "C:\", 3
    Err = 0
    
End Sub



Public Function ShowProperties(FileName As String, OwnerhWnd As Long) As Long
        
    '     'open a file properties property page for specified file if return value
    '     '<=32 an error occurred
    '     'From: Delphi code provided by "Ian Land" (iml@dircon.co.uk)
    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
     
    '     'Fill in the SHELLEXECUTEINFO structure
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hWnd = OwnerhWnd
        .lpVerb = "properties"
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 0
        .hInstApp = 0
        .lpIDList = 0
    End With
       
    '     'call the API
    r = ShellExecuteEX(SEI)
 
    '     'return the instance handle as a sign of success
    ShowProperties = SEI.hInstApp
       
End Function

Public Sub FontStuff(ByVal Titulo As String, picDraw As PictureBox)
    
    On Error GoTo GetOut
    Dim f As LOGFONT, hPrevFont As Long, hFont As Long, FontName As String
    Dim FONTSIZE As Integer
    FONTSIZE = 10 'Val(txtSize.Text)
    
    f.lfEscapement = 10 * 90 'Val(txtDegree.Text) 'rotation angle, in tenths
    FontName = "Tahoma" + Chr$(0) 'null terminated
    f.lfFaceName = FontName
    f.lfHeight = (FONTSIZE * -20) / Screen.TwipsPerPixelY
    hFont = CreateFontIndirect(f)
    hPrevFont = SelectObject(picDraw.hDC, hFont)
    
    picDraw.CurrentX = 3
    'picDraw.CurrentY = 310
    
    picDraw.CurrentY = picDraw.Height - 10
    picDraw.Print Titulo
    
    '  Clean up, restore original font
    hFont = SelectObject(picDraw.hDC, hPrevFont)
    DeleteObject hFont
    
    Exit Sub
GetOut:
    Exit Sub

End Sub

