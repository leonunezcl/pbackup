VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Project Backup"
   ClientHeight    =   5715
   ClientLeft      =   2580
   ClientTop       =   1980
   ClientWidth     =   7200
   HelpContextID   =   10
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   381
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   Begin VB.PictureBox pixSmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6600
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   495
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.ListBox List1 
      Height          =   450
      Left            =   5505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   4710
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   4770
      Left            =   45
      ScaleHeight     =   316
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   405
      Width           =   360
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   5910
      Top             =   1725
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lvwArchivos 
      Height          =   5010
      Left            =   645
      TabIndex        =   2
      Top             =   360
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   8837
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgFiles"
      SmallIcons      =   "imgFiles"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "cmdlvwsel"
         Text            =   "Nº"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "cmdlvwarchivo"
         Text            =   "Archivo"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "cmdlvwtipo"
         Text            =   "Tipo"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "cmdlvwtamano"
         Text            =   "Tamaño"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Key             =   "cmdlvwfecha"
         Text            =   "Fecha y Hora"
         Object.Width           =   2910
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Key             =   "cmdlvwpath"
         Text            =   "Path"
         Object.Width           =   6615
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   2985
      Top             =   2220
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":08CA
            Key             =   "kAbrir"
            Object.Tag             =   "&Abrir"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12BA
            Key             =   "kSalir"
            Object.Tag             =   "&Salir"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":170E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2A
            Key             =   ""
            Object.Tag             =   "&Propiedades de archivo"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B86
            Key             =   "kIndice"
            Object.Tag             =   "&Indice"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CE2
            Key             =   "kEliminar"
            Object.Tag             =   "&Eliminar"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25BE
            Key             =   "kAgregar"
            Object.Tag             =   "&Agregar"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E9A
            Key             =   "kRespaldar"
            Object.Tag             =   "&Respaldar"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3776
            Key             =   ""
            Object.Tag             =   "&Ir a VBSoftware"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3DAA
            Key             =   ""
            Object.Tag             =   "&Seleccionar todos ..."
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F0A
            Key             =   ""
            Object.Tag             =   "&Invertir selección"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":406A
            Key             =   ""
            Object.Tag             =   "&Quitar todos ..."
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":41CA
            Key             =   ""
            Object.Tag             =   "&Tip del dia ..."
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":44E6
            Key             =   ""
            Object.Tag             =   "&Email a VBSoftware"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4642
            Key             =   ""
            Object.Tag             =   "&Configurar"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A96
            Key             =   ""
            Object.Tag             =   "&Archivos desde path"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar stbMain 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5460
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13229
            MinWidth        =   13229
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAbrir"
            Object.ToolTipText     =   "Abrir proyecto visual basic"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRespaldar"
            Object.ToolTipText     =   "Respaldar proyecto"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdRemover"
            Object.ToolTipText     =   "Eliminar archivo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAgregar"
            Object.ToolTipText     =   "Agregar archivo"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPath"
            Object.ToolTipText     =   "Agregar archivos desde path"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdPropiedades"
            Object.ToolTipText     =   "Propiedades del archivo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdOpc"
            Object.ToolTipText     =   "Opciones"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdNet"
            Object.ToolTipText     =   "Ir al sitio web de VBSoftware"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdEmail"
            Object.ToolTipText     =   "Enviar email"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdTip"
            Object.ToolTipText     =   "Tips"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdAyuda"
            Object.ToolTipText     =   "Ayuda de la aplicación"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "cmdSalir"
            Object.ToolTipText     =   "Salir de la aplicación"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.Image imgIco 
      Height          =   240
      Left            =   6615
      Stretch         =   -1  'True
      Top             =   780
      Width           =   240
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      HelpContextID   =   20
      Begin VB.Menu mnuArchivo_Abrir 
         Caption         =   "|Abre un proyecto Visual Basic|&Abrir proyecto ..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuArchivo_Respaldar 
         Caption         =   "|Respaldar proyecto seleccionado|&Respaldar proyecto"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuArchivo_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Resumen 
         Caption         =   "|Muestra un resumen de los archivos a respaldar|Resumen de archivos"
      End
      Begin VB.Menu mnuArchivo_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArchivo_Salir 
         Caption         =   "|Salir de la aplicación|&Salir"
      End
      Begin VB.Menu mnuArchivo_sep3 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuArchivo_Proyecto 
         Caption         =   "XXX"
         Index           =   0
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "&Edición"
      HelpContextID   =   30
      Begin VB.Menu mnuEdicion_Agregar 
         Caption         =   "|Agregar archivo al proyecto seleccionado|&Agregar archivo"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEdicion_Eliminar 
         Caption         =   "|Eliminar archivo del proyecto seleccionado|&Eliminar archivo"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuEdicion_AgregarPath 
         Caption         =   "|Agregar todos los archivos desde path especificado|&Archivos desde path"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuEdicion_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_Propiedades 
         Caption         =   "|Propiedades del archivo a respaldar|&Propiedades de archivo"
      End
      Begin VB.Menu mnuEdicion_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEdicion_seleccionar 
         Caption         =   "|Seleccionar todos los archivos|&Seleccionar todos ..."
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuEdicion_Invertir 
         Caption         =   "|Invertir selección de archivos|&Invertir selección"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuEdicion_Quitar 
         Caption         =   "|Quitar selección de archivos|&Quitar todos ..."
         Shortcut        =   {F9}
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      HelpContextID   =   220
      Begin VB.Menu mnuOpciones_ConfRespaldo 
         Caption         =   "|Configurar opciones a respaldar|&Configurar opciones de respaldo"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuOpciones_sep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuOpciones_Idioma 
         Caption         =   "|Seleccionar archivo de lenguaje|Seleccionar idioma ..."
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&Ayuda"
      Begin VB.Menu mnuAyuda_Indice 
         Caption         =   "|Indice de ayuda de la aplicación|&Indice"
         HelpContextID   =   10
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAyuda_Busqueda 
         Caption         =   "|Buscar en archivo de ayuda|B&usqueda ..."
      End
      Begin VB.Menu mnuAyuda_sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_WebSite 
         Caption         =   "|Ir al sitio WWW de VBSoftware|&Ir a VBSoftware"
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuAyuda_Email 
         Caption         =   "|Escribe un email a VBSoftware|&Email a VBSoftware"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuAyuda_sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_Tips 
         Caption         =   "|Mostrar tips del dia|&Tip del dia ..."
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuAyuda_sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyuda_Acercade 
         Caption         =   "|Información de Copyright de la aplicación y sobre el Autor|Acerca de ..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents MyHelpCallBack As HelpCallBack
Attribute MyHelpCallBack.VB_VarHelpID = -1
Private clsXmenu As New CXtremeMenu

Private Cdlg As New GCommonDialog
Private Itmx As ListItem
Private mGradient As New clsGradient
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1
Private total As Long
Private bProcese As Boolean
Private glSmallIcons() As Long
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Const BIF_RETURNONLYFSDIRS = 1
Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you many not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source on
'               any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const SHGFI_DISPLAYNAME = &H200
Private Const SHGFI_EXETYPE = &H2000
Private Const SHGFI_SYSICONINDEX = &H4000 'system icon index
Private Const SHGFI_LARGEICON = &H0 'large icon
Private Const SHGFI_SMALLICON = &H1 'small icon
Private Const SHGFI_SHELLICONSIZE = &H4
Private Const SHGFI_TYPENAME = &H400
Private Const ILD_TRANSPARENT = &H1 'display transparent
Private Const BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or _
                                 SHGFI_SHELLICONSIZE Or _
                                 SHGFI_SYSICONINDEX Or _
                                 SHGFI_DISPLAYNAME Or _
                                 SHGFI_EXETYPE

Private Type SHFILEINFO
   hIcon As Long
   iIcon As Long
   dwAttributes As Long
   szDisplayName As String * MAX_PATH
   szTypeName As String * 80
End Type

Private Declare Function SHGetFileInfo Lib "shell32" _
   Alias "SHGetFileInfoA" _
   (ByVal pszPath As String, _
    ByVal dwFileAttributes As Long, _
    psfi As SHFILEINFO, _
    ByVal cbSizeFileInfo As Long, _
    ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32" _
   (ByVal himl As Long, ByVal i As Long, _
    ByVal hDCDest As Long, ByVal x As Long, _
    ByVal y As Long, ByVal flags As Long) As Long

Private shinfo As SHFILEINFO

'agregar archivo al listview
Private Sub AgregarArchivoAlvw(ByVal File As String, ByRef Tipo As String)

    Dim Ext As String
    Dim k As Integer
    Dim Indice As Integer
    
    stbMain.Panels(1).Text = "Agregando : " & File
        
    'extension del archivo
    Ext = ExtensionArchivo(File)
        
    'llave
    k = lvwArchivos.ListItems.Count + 1
                    
    'ver si imagen esta en imagelist
    Call ExtraerIconoAsociado(File, Ext)
    
    'indice de la imagen
    Indice = IndiceImagen(Ext, File)
    
    'cargar imagen
    Call lvwArchivos.ListItems.Add(, "k" & k, Format(k, "000"), Indice, Indice)
                
    Tipo = Ext
    
End Sub
Private Sub AgregarArchivosPath()

    Dim iNull As Integer, lpIDList As Long, lResult As Long
    Dim sPath As String, udtBI As BrowseInfo
    Dim File As String
    Dim Path As String
    Dim k As Integer
    Dim Ext As String
    Dim Tipo As String
    Dim Tamaño As Double
    
    If glbArchivo = "" Then
        MsgBox "Debe seleccionar un proyecto.", vbCritical
        Exit Sub
    End If
    
    Call Hourglass(hWnd, True)
    Call Habilitar(False)
    
    List1.Clear
    
    With udtBI
        'Set the owner window
        .hWndOwner = Me.hWnd
        'lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat("C:\", "")
        'Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    'Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        'Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
    End If

    'verificar si se selecciono un path
    If Len(sPath) > 0 Then
        'buscar archivos desde el path señalado
        If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
            
        'cargar archivos desde path
        Call CargarArchivosDesdePath(sPath)
    End If
    
    Call TotalItemes
    Call Hourglass(hWnd, False)
    Call Habilitar(True)

End Sub
'verificar que el archivo no exista en listview
Private Function ArchivoEnProyecto(ByVal archivo As String) As Boolean

    Dim itmFound As ListItem
    Dim ret As Boolean
    
    Set itmFound = lvwArchivos.FindItem(VBArchivoSinPath(archivo), lvwSubItem, , lvwPartial)
      
    If itmFound Is Nothing Then
        ret = False
    Else
        ret = True
    End If
    
    ArchivoEnProyecto = ret
    
End Function

'carga el proyecto a respaldar
Private Sub CargaProyectoARespaldar()

    Dim imgX As ListImage
    
    Call Hourglass(hWnd, True)
    Call Habilitar(False)
                        
    'limpiar las imagenes
    Call LimpiaImagenes
        
    'iniciar listimage ?
    Set imgX = imgFiles.ListImages.Add(, , Me.Icon)
   
    'cargar archivos del proyecto
    If CargaProyecto() Then
        'setear nueva imagen
        lvwArchivos.View = lvwReport
        lvwArchivos.Icons = imgFiles
        lvwArchivos.SmallIcons = imgFiles
        
        'cargar archivos en listview
        Call CargaArchivosLista
        
        'grabar historial en .ini
        Call GrabarProyectoINI(glbArchivo)
        
        'cargar en menu de proyectos abiertos
        Call CargarProyectosExplorados
        
        'actualizar archivos
        Call mnuEdicion_seleccionar_Click
        
        MsgBox "Proyecto cargado con éxito!", vbInformation
    Else
        MsgBox "Proyecto no fue cargado.", vbCritical
    End If
    
    Set imgX = Nothing
    
    Call Habilitar(True)
    Call Hourglass(hWnd, False)
        
End Sub

'cargar los archivos desde el path especificado
Public Sub CargarArchivosDesdePath(ByVal sPath As String)

    Dim k As Integer
    Dim File As String
    Dim Tipo As String
    Dim Tamaño As Double
    
    Call FindFilesAPI(sPath, "*.*", 1, 1)
        
    SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, False, 0
    
    'agregar los archivos a la lista
    For k = 0 To List1.ListCount - 1
        ValidateRect lvwArchivos.hWnd, 0&
        
        'archivo
        File = List1.List(k)
        
        'verificar si el archivo no esta en el proyecto
        If Not ArchivoEnProyecto(File) Then
            'agregar tipo de archivo
            Call AgregarArchivoAlvw(File, Tipo)
                    
            Set Itmx = lvwArchivos.ListItems(lvwArchivos.ListItems.Count)
            
            Itmx.SubItems(1) = VBArchivoSinPath(File)
            Itmx.SubItems(2) = Tipo
            
            Tamaño = VBGetFileSize(File)
            If Tamaño > 0 Then
                Itmx.SubItems(3) = Tamaño & " KB"
            Else
                Itmx.SubItems(3) = Tamaño & " Bytes"
            End If
            
            Itmx.SubItems(4) = VBGetFileTime(File)
            Itmx.SubItems(5) = PathArchivo(File)
        End If
        
        If (k Mod 10) = 0 Then
            InvalidateRect lvwArchivos.hWnd, 0&, 0&
            DoEvents
        End If
    Next k
    
    SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, True, 0
    
    ValidateRect lvwArchivos.hWnd, 0&
                
    'total de archivos a respaldar
    Call TotalItemes
    
End Sub

Private Sub ConfiguraMenus()

    Dim k As Integer
    
    clsXmenu.Uninstall hWnd
    DoEvents
    
    For k = 1 To imgMain.ListImages.Count
        If imgMain.ListImages(k).Key = "kAbrir" Then
            imgMain.ListImages(k).Tag = MenuName(mnuArchivo_Abrir.Caption)
        ElseIf imgMain.ListImages(k).Key = "kSalir" Then
            imgMain.ListImages(k).Tag = MenuName(mnuArchivo_Salir.Caption)
        ElseIf imgMain.ListImages(k).Key = "kRespaldar" Then
            imgMain.ListImages(k).Tag = MenuName(mnuArchivo_Respaldar.Caption)
        ElseIf imgMain.ListImages(k).Key = "kAgregar" Then
            imgMain.ListImages(k).Tag = MenuName(mnuEdicion_Agregar.Caption)
        ElseIf imgMain.ListImages(k).Key = "kEliminar" Then
            imgMain.ListImages(k).Tag = MenuName(mnuEdicion_Eliminar.Caption)
        ElseIf imgMain.ListImages(k).Key = "kIndice" Then
            imgMain.ListImages(k).Tag = MenuName(mnuAyuda_Indice.Caption)
        End If
    Next k
    
    Set MyHelpCallBack = New HelpCallBack
    
    Call clsXmenu.Install(hWnd, MyHelpCallBack, Me.imgMain)
    Call clsXmenu.FontName(hWnd, "Tahoma")
        
End Sub

'devuelve la extension del archivo
Public Function ExtensionArchivo(ByVal File As String) As String

    Dim ret As String
    
    ret = ""
    If InStr(1, File, ".") > 0 Then
        ret = Mid$(File, InStr(1, File, ".") + 1)
    End If
    
    ExtensionArchivo = ret
    
End Function

'extraer el icono asociado al archivo
Private Sub ExtraerIconoAsociado(ByVal archivo As String, ByVal Ext As String)
    
    Dim k As Integer
    Dim hImgSmall As Long
    Dim found As Boolean
        
    found = False
    
    'verificar que no sea un archivo de icono
    If LCase$(Ext) <> "ico" Then
        'verificar si imagen ya existe en listimage
        For k = 1 To imgFiles.ListImages.Count
            If UCase$(imgFiles.ListImages(k).Key) = UCase$(Ext) Then
                found = True
                Exit For
            End If
        Next k
    Else
        'limpiar picture
        pixSmall.Picture = LoadPicture()
        pixSmall.AutoRedraw = True
        
        'cargar icono
        imgIco.Picture = LoadPicture(archivo)
                
        pixSmall.Picture = imgIco.Picture
        
        On Local Error Resume Next
        'agregar a listimgage
        imgFiles.ListImages.Add , archivo, pixSmall.Image
        
        Err = 0
        
        Exit Sub
    End If
    
    'se encontro imagen ?
    If Not found Then
         'get the system icon associated with that file
         hImgSmall& = SHGetFileInfo(archivo, 0&, _
                                   shinfo, Len(shinfo), _
                                   BASIC_SHGFI_FLAGS Or SHGFI_SMALLICON)
        
         'set a picture box to receive the small icon
         'its size must be 16x16 pixels (240x240 twips),
         'with no 3d or border!
         'clear any existing image
         pixSmall.Picture = LoadPicture()
         pixSmall.AutoRedraw = True
                 
         'draw the associated icon into the picturebox
         Call ImageList_Draw(hImgSmall&, shinfo.iIcon, _
                            pixSmall.hDC, 0, 0, ILD_TRANSPARENT)
                            
         'realize the image by assigning its image property
         '(where the icon was drawn) to the actual picture property
         pixSmall.Picture = pixSmall.Image
        
         'esperar hasta que se dibuje
         Do
             If Not pixSmall.Picture Is Nothing Then
                 Exit Do
             End If
         Loop

        'agregar a listimgage
        imgFiles.ListImages.Add , Ext, pixSmall.Image
    End If
            
End Sub

'graba los
Private Sub GrabarProyectoINI(ByVal archivo As String)

    Dim k
    Dim j As Integer
    Dim sProyecto As String
    Dim sArchivos()
    Dim sProyectos()
    
    Dim n As Integer
    
    k = Ini.Leer("proyectos", "archivos", C_INI)
    If k = "" Then k = 1
    
    ReDim sArchivos(4)
    ReDim sProyectos(0)
    
    sArchivos(1) = archivo
    
    'leer anteriores proyectos
    n = 4
    For j = 4 To 1 Step -1
        sProyecto = Ini.Leer("proyectos", "archivo" & j, C_INI)
        'si proyecto leido es distinto al que tengo que grabar
        sArchivos(n) = sProyecto
        n = n - 1
    Next j
            
    'ciclo de 1 a 4. max 4 proyectos a analizados.
    'queda como el 1 el ultimo analizado.
    ReDim Preserve sProyectos(1)
    
    sProyectos(1) = archivo
    
    For k = 1 To 4
        If sArchivos(k) <> "" Then 'si no esta en blanco
            If sArchivos(k) <> archivo Then
                ReDim Preserve sProyectos(UBound(sProyectos) + 1)
                sProyectos(UBound(sProyectos)) = sArchivos(k)
            End If
        End If
    Next k
        
    For k = 1 To UBound(sProyectos)
        Call Ini.Grabar(C_INI, "proyectos", "archivo" & k, sProyectos(k))
    Next k
    
    If UBound(sProyectos) < 4 Then
        n = UBound(sProyectos)
    Else
        n = 4
    End If
    
    'grabar los n proyectos analizados
    Call Ini.Grabar(C_INI, "proyectos", "archivos", n)
    
End Sub

Private Sub CargarProyectosExplorados()

    On Local Error Resume Next
    
    Dim k
    Dim j As Integer
    Dim sProyecto As String
    
    k = Ini.Leer("proyectos", "archivos", C_INI)
    
    If k <> "" And Val(k) > 0 Then
        mnuArchivo_sep3.Visible = True
        'descargar todos los menus cargados dinamicamente
        For j = 3 To 1 Step -1
            Unload mnuArchivo_Proyecto(j)
        Next j
        
        For j = 1 To Val(k)
            sProyecto = Ini.Leer("proyectos", "archivo" & j, C_INI)
            If sProyecto <> "" Then
                If j > 1 Then
                    Load mnuArchivo_Proyecto(j - 1)
                End If
                mnuArchivo_Proyecto(j - 1).Caption = sProyecto
                mnuArchivo_Proyecto(j - 1).Visible = True
            End If
        Next j
    End If
    
    Err = 0
    
End Sub
'agrega un archivo a la lista
Private Sub AgregarArchivo()

    Dim Glosa As String
    Dim archivo As String
    Dim File As String
    Dim Path As String
    Dim k As Integer
    Dim Tipo As String
    Dim Tamaño As Double
    
    'verificar si se selecciono un proyecto
    If glbArchivo = "" Then
        MsgBox "Debe seleccionar un proyecto.", vbCritical
        Exit Sub
    End If
    
    If gsLastPath = "" Then gsLastPath = App.Path
    
    Glosa = "Todos los archivos (*.*)|*.*"
        
    If Cdlg.VBGetOpenFileName(archivo, , , True, , , Glosa, , gsLastPath, "Agregar archivo ...", "TXT") Then
        'comenzar a agregar archivos
        Call Hourglass(hWnd, True)
        Call Habilitar(False)
        
        Path = ""
        If InStr(1, archivo, Chr(32)) > 0 Then
            Path = Left$(archivo, InStr(1, archivo, Chr(32)) - 1)
            If Right$(Path, 1) <> "\" Then
                Path = Path & "\"
            End If
            archivo = Mid$(archivo, InStr(1, archivo, Chr(32)) + 1)
        End If
        
        SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, False, 0
        
        'agregar los archivos
        k = 1
        Do While archivo <> ""
            
            ValidateRect lvwArchivos.hWnd, 0&
            
            If InStr(1, archivo, Chr(32)) > 0 Then
                File = Path & Left$(archivo, InStr(1, archivo, Chr(32)) - 1)
            Else
                File = Path & archivo
            End If
            
            'verificar si el archivo no esta en el proyecto
            If Not ArchivoEnProyecto(File) Then
                'agregar tipo de archivo
                Call AgregarArchivoAlvw(File, Tipo)
                
                Set Itmx = lvwArchivos.ListItems(lvwArchivos.ListItems.Count)
                Itmx.SubItems(1) = VBArchivoSinPath(File)
                
                stbMain.Panels(1).Text = "Agregando : " & Itmx.SubItems(1)
                Itmx.SubItems(2) = Tipo
                
                Tamaño = VBGetFileSize(File)
                
                If Tamaño > 0 Then
                    Itmx.SubItems(3) = Tamaño & " KB"
                Else
                    Itmx.SubItems(3) = Tamaño & " Bytes"
                End If
                
                Itmx.SubItems(4) = VBGetFileTime(File)
                Itmx.SubItems(5) = PathArchivo(File)
            End If
            
            If (k Mod 10) = 0 Then
                DoEvents
                InvalidateRect lvwArchivos.hWnd, 0&, 0&
            End If
    
            If InStr(1, archivo, Chr(32)) > 0 Then
                archivo = Mid$(archivo, InStr(1, archivo, Chr(32)) + 1)
            Else
                archivo = ""
            End If
            k = k + 1
        Loop
                
        SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, True, 0
        
        Call TotalItemes
        
        Call Habilitar(True)
        Call Hourglass(hWnd, False)
        
        MsgBox "Archivos agregados con éxito!", vbInformation
    End If
    
End Sub

'carga los archivos del proyecto
Private Sub CargaArchivosLista()

    Dim k As Integer
    Dim Tipo As String
    Dim c As Integer
    Dim archivo As String
    
    SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, False, 0
    
    'cargar dependencias
    For k = 1 To UBound(Proyecto.aDepencias)
        ValidateRect lvwArchivos.hWnd, 0&
        
        'extraer nombre del archivo
        archivo = VBArchivoSinPath(Proyecto.aDepencias(k).archivo)
        
        stbMain.Panels(1).Text = archivo
        
        'agregar archivo a listview
        Call AgregarArchivoAlvw(Proyecto.aDepencias(k).archivo, Tipo)
                        
        Set Itmx = lvwArchivos.ListItems(k)
        
        Itmx.SubItems(1) = archivo
        Itmx.SubItems(2) = Tipo
        
        If Proyecto.aDepencias(k).FileSize > 0 Then
            Itmx.SubItems(3) = Proyecto.aDepencias(k).FileSize & " KB"
        Else
            Itmx.SubItems(3) = Proyecto.aDepencias(k).FileSize & " Bytes"
        End If
        
        Itmx.SubItems(4) = Proyecto.aDepencias(k).FILETIME
        Itmx.SubItems(5) = PathArchivo(Proyecto.aDepencias(k).archivo)
        
        If (k Mod 10) = 0 Then
            DoEvents
            InvalidateRect lvwArchivos.hWnd, 0&, 0&
        End If
    Next k
    
    ValidateRect lvwArchivos.hWnd, 0&
    
    c = k
    
    'cargar archivos del proyecto
    For k = 1 To UBound(Proyecto.aArchivos)
        ValidateRect lvwArchivos.hWnd, 0&
        
        'nombre fisico del archivo
        archivo = VBArchivoSinPath(Proyecto.aArchivos(k).PathFisico)
        
        'agregar al listview segun el tipo de archivo
        Call AgregarArchivoAlvw(Proyecto.aArchivos(k).PathFisico, Tipo)
                
        'setear datos opcionales
        Set Itmx = lvwArchivos.ListItems(c)
        Itmx.SubItems(1) = archivo
        Itmx.SubItems(2) = Tipo
        
        If Proyecto.aArchivos(k).FileSize > 0 Then
            Itmx.SubItems(3) = Proyecto.aArchivos(k).FileSize & " KB"
        Else
            Itmx.SubItems(3) = Proyecto.aArchivos(k).FileSize & " Bytes"
        End If
        
        Itmx.SubItems(4) = Proyecto.aArchivos(k).FILETIME
        Itmx.SubItems(5) = PathArchivo(Proyecto.aArchivos(k).PathFisico)
        
        If (k Mod 10) = 0 Then
            DoEvents
            InvalidateRect lvwArchivos.hWnd, 0&, 0&
        End If
        c = c + 1
    Next k
    
    ValidateRect lvwArchivos.hWnd, 0&
    
    SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, True, 0
    
    'cargar archivos desde path interiores ?
    If glbRecursive Then
        Call CargarArchivosDesdePath(PathArchivo(Proyecto.PathFisico))
    End If
    
    Call TotalItemes
    
    ValidateRect lvwArchivos.hWnd, 0&
    
End Sub

Private Sub Habilitar(ByVal Estado As Boolean)

    On Local Error Resume Next
    
    Dim k As Integer
    
    For k = 1 To tbrMain.Buttons.Count
        DoEvents
        tbrMain.Buttons(k).Enabled = Estado
    Next k
    
    For k = 1 To Controls.Count
        DoEvents
        If TypeOf Controls(k) Is Menu Then
            Controls(k).Enabled = Estado
        End If
    Next k
    
    Err = 0
    
End Sub

'indice de la imagen en picture
Private Function IndiceImagen(ByVal Ext As String, ByVal File As String) As Integer

    Dim ret As Integer
    Dim k As Integer
        
    'ciclar x las imagenes
    For k = 1 To imgFiles.ListImages.Count
        If UCase$(imgFiles.ListImages(k).Key) = UCase$(Ext) Or UCase$(imgFiles.ListImages(k).Key) = UCase$(File) Then
            ret = k
            Exit For
        End If
    Next k
        
    IndiceImagen = ret
    
End Function

Private Sub LeerInfoIni()

    Dim valor As Variant
    
    valor = Ini.Leer("opciones", "recursive", C_INI)
    
    If valor = "" Then
        glbRecursive = False
    ElseIf valor = "1" Then
        glbRecursive = True
    Else
        glbRecursive = False
    End If
    
    valor = Ini.Leer("opciones", "savepath", C_INI)
    
    If valor = "" Then
        glbSavePath = False
    ElseIf valor = "1" Then
        glbSavePath = True
    Else
        glbSavePath = False
    End If
    
    valor = Ini.Leer("opciones", "reportes", C_INI)
    
    If valor = "" Then
        glbRpt = False
    ElseIf valor = "1" Then
        glbRpt = True
    Else
        glbRpt = False
    End If
    
    valor = Ini.Leer("opciones", "imagenes", C_INI)
    
    If valor = "" Then
        glbIma = False
    ElseIf valor = "1" Then
        glbIma = True
    Else
        glbIma = False
    End If
    
    valor = Ini.Leer("opciones", "configuracion", C_INI)
    
    If valor = "" Then
        glbIni = False
    ElseIf valor = "1" Then
        glbIni = True
    Else
        glbIni = False
    End If
    
    valor = Ini.Leer("opciones", "ayuda", C_INI)
    
    If valor = "" Then
        glbHlp = False
    ElseIf valor = "1" Then
        glbHlp = True
    Else
        glbHlp = False
    End If
    
End Sub
'limpiar las imagenes
Private Sub LimpiaImagenes()

    List1.Clear
    
    lvwArchivos.ListItems.Clear
    Set lvwArchivos.SmallIcons = Nothing
    Set lvwArchivos.Icons = Nothing
    
    imgFiles.ListImages.Clear
    imgFiles.ImageHeight = 16
    imgFiles.ImageWidth = 16
    
End Sub

Private Sub Logo()

    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picMain
    End With
        
    Call FontStuff(App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision, picMain)
    
    picMain.Refresh
    
End Sub

'devuelve el nombre del menu
Private Function MenuName(ByVal Key As String) As String

    Dim k As Integer
    Dim ret As String
    
    For k = Len(Key) To 1 Step -1
        If Mid$(Key, k, 1) = "|" Then
            ret = Mid$(Key, k + 1)
            Exit For
        End If
    Next k
    
    MenuName = ret
    
End Function

'muestra las propiedades del archivo
Private Sub MostrarPropiedades()

    Dim Path As String
    Dim sFile As String
    Dim k As Integer
    
    If glbArchivo = "" Then
        MsgBox "Debe seleccionar un proyecto.", vbCritical
        Exit Sub
    End If
    
    If Not lvwArchivos.SelectedItem Is Nothing Then
    
        k = lvwArchivos.SelectedItem.Index
        sFile = lvwArchivos.ListItems(k).SubItems(1)
        Path = lvwArchivos.ListItems(k).SubItems(5)
        
        If Right$(Path, 1) <> "\" Then
            Path = Path & "\"
        End If

        sFile = Path & sFile
        
        Call ShowProperties(sFile, hWnd)
    End If
    
End Sub

'remueve el archivo de la lista
Private Sub RemoverArchivo()

    Dim Msg As String
    Dim k As Integer
    
    If glbArchivo = "" Then
        MsgBox "Debe seleccionar un proyecto.", vbCritical
        Exit Sub
    End If
    
    'verificar si hay un archivo seleccionado
    If Not lvwArchivos.SelectedItem Is Nothing Then
        Msg = "Confirma eliminar archivo."
        If Confirma(Msg) = vbYes Then
            Call Habilitar(False)
            Call Hourglass(hWnd, True)
            
            SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, False, 0
            
            'ciclar x los archivos
            For k = lvwArchivos.ListItems.Count To 1 Step -1
                ValidateRect lvwArchivos.hWnd, 0&
                If lvwArchivos.ListItems(k).Selected Then
                    stbMain.Panels(1).Text = "Eliminando : " & lvwArchivos.ListItems(k).SubItems(1)
                    lvwArchivos.ListItems.Remove lvwArchivos.ListItems(k).Index
                End If
                If (k Mod 10) = 0 Then
                    DoEvents
                    InvalidateRect lvwArchivos.hWnd, 0&, 0&
                End If
            Next k
            
            SendMessageLong lvwArchivos.hWnd, WM_SETREDRAW, True, 0
            
            Call TotalItemes
            Call Habilitar(True)
            Call Hourglass(hWnd, False)
        End If
    Else
        MsgBox "Debe seleccionar un archivo.", vbCritical
    End If
        
End Sub

'verifica cuantos archivos hay a respaldar
Private Sub TotalItemes()
    stbMain.Panels(1).Text = lvwArchivos.ListItems.Count & " archivos a respaldar."
End Sub

'respalda los archivos del proyecto
Private Function RespaldaProyecto() As Boolean

    Dim Msg As String
    Dim ret As Boolean
    Dim k As Integer
    Dim First As Boolean
    Dim Path As String
    Dim sFile As String
    Dim sFile2 As String
    Dim c As Integer
    Dim e As Long
    Dim Glosa As String
    
    ret = True
    First = True
    
    Call Hourglass(hWnd, True)
    Call Habilitar(False)
    
    Glosa = "Archivos ZIP (*.ZIP)|*.ZIP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"

    'seleccionar el path donde se guarda el archivo
    If Not Cdlg.VBGetSaveFileName(glbArchivoZIP, , , Glosa, , gsLastPath, "Guardar respaldo como ...", "ZIP") Then
        ret = False
    End If
           
    'verificar si archivo existe
    If VBOpenFile(glbArchivoZIP) Then
        Msg = "El archivo ya existe." & vbNewLine & vbNewLine
        Msg = Msg & "Confirma eliminar archivo existente."
        If Confirma(Msg) = vbYes Then
            DeleteFile glbArchivoZIP
        Else
            MsgBox "Se agregaran los archivos al archivo de respaldo", vbInformation
            First = False
        End If
    End If
    
    frmAccion.total = total
    frmAccion.Show
    c = 1
    
    'ciclar x los archivos del proyecto
    If glbRecursive Then
        m_cZ.BasePath = PathArchivo(glbArchivoZIP)
    End If
    
    'ciclar x los archivos seleccionados
    For k = 1 To lvwArchivos.ListItems.Count
        e = DoEvents()
            
        sFile = lvwArchivos.ListItems(k).SubItems(1)
        sFile2 = sFile
        Path = lvwArchivos.ListItems(k).SubItems(5)
        
        If Right$(Path, 1) <> "\" Then
            Path = Path & "\"
        End If

        sFile = Path & sFile
        
        If First Then
            m_cZ.AllowAppend = False
            First = False
        Else
            m_cZ.AllowAppend = True
        End If
        
        frmAccion.Label1.Caption = sFile2
        frmAccion.pgb.Value = c
        
        'zipear archivos ...
        Call Zipear(sFile)
        c = c + 1
    Next k
    
    Call Zipear(glbArchivo)
    
    Unload frmAccion
    
    Call Habilitar(True)
    Call Hourglass(hWnd, False)
        
    stbMain.Panels(1).Text = "Listo!"
    
    RespaldaProyecto = ret
    
End Function
'valida que existan archivos seleccionados
Private Function ValidaArchivos() As Boolean
        
    total = lvwArchivos.ListItems.Count
        
    ValidaArchivos = True
    
End Function

'agregar archivo al arhivo .zip
Private Sub Zipear(ByVal sFile As String)

    With m_cZ
        .ZipFile = glbArchivoZIP
        .StoreFolderNames = False
        .ClearFileSpecs
        .AddFileSpec sFile
       .StoreFolderNames = glbSavePath
       .Zip
    End With
            
End Sub

Private Sub Form_Activate()

    If Len(Command) > 0 And Not bProcese Then
        bProcese = True
        glbArchivo = Command
        
        Call CargaProyectoARespaldar
    End If
    
End Sub

Private Sub Form_Load()

    If IsDebuggerPresent <> 0 Then End
               
    CenterWindow hWnd
            
    Set m_cZ = New cZip
    stbMain.Panels(1).Text = "VBSoftware 2000-2002"
    
    Call ConfiguraMenus
            
    Call LeerInfoIni
    Call CargarProyectosExplorados
        
    RemoveMenus Me, False, False, _
        False, False, False, True, True
        
    SetAppHelp hWnd
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Dim Msg As String
    
    Msg = "Confirma salir de la aplicación."
    
    If Confirma(Msg) = vbNo Then
        Cancel = 1
        Exit Sub
    End If
    
    Set lvwArchivos.Icons = Nothing
    Set lvwArchivos.SmallIcons = Nothing
        
    QuitHelp
    
End Sub


Private Sub Form_Resize()

    If WindowState <> vbMinimized Then
        picMain.Left = 0
        picMain.Top = tbrMain.Height + 1
        picMain.Height = ScaleHeight - tbrMain.Height - stbMain.Height
        
        lvwArchivos.Left = picMain.Width '+ 1
        lvwArchivos.Top = tbrMain.Height + 1
        lvwArchivos.Height = ScaleHeight - tbrMain.Height - stbMain.Height
        lvwArchivos.Width = ScaleWidth - picMain.Width
        
        Call Logo
    End If
    
End Sub


Private Sub lvwArchivos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Call Hourglass(hWnd, True)
    Call Habilitar(False)
    
    If lvwArchivos.SortOrder = lvwAscending Then
        lvwArchivos.SortOrder = lvwDescending
    Else
        lvwArchivos.SortOrder = lvwAscending
    End If
    
    lvwArchivos.SortKey = ColumnHeader.Index - 1
    
    lvwArchivos.Sorted = True
    
    Call Hourglass(hWnd, False)
    Call Habilitar(True)
    
End Sub


Private Sub lvwArchivos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        PopupMenu mnuEdicion
    End If
End Sub


Private Sub m_cZ_Progress(ByVal lCount As Long, ByVal sMsg As String)
    stbMain.Panels(1).Text = sMsg
End Sub


Private Sub mnuArchivo_Abrir_Click()
    
    Dim Glosa As String
    Dim ArchivoTmp As String
    
    If gsLastPath = "" Then gsLastPath = App.Path
    If Len(glbArchivo) > 0 Then ArchivoTmp = glbArchivo
    
    Glosa = "Visual Basic 3.0 (*.MAK)|*.MAK|"
    Glosa = Glosa & "Visual Basic 4,5,6 (*.VBP)|*.VBP|"
    Glosa = Glosa & "Todos los archivos (*.*)|*.*"
        
    If Cdlg.VBGetOpenFileName(glbArchivo, , , , , , Glosa, , gsLastPath, "Abrir proyecto VB ...", "VBP") Then
        Call CargaProyectoARespaldar
    End If
    
    If Len(ArchivoTmp) > 0 Then glbArchivo = ArchivoTmp
    
End Sub
Private Sub mnuArchivo_Proyecto_Click(Index As Integer)

    glbArchivo = mnuArchivo_Proyecto(Index).Caption
    
    Call CargaProyectoARespaldar
    
End Sub

Private Sub mnuArchivo_Respaldar_Click()
    Call tbrMain_ButtonClick(tbrMain.Buttons("cmdRespaldar"))
End Sub

Private Sub mnuArchivo_Resumen_Click()

    If glbArchivo = "" Then
        MsgBox "Debe seleccionar un proyecto.", vbCritical
        Exit Sub
    End If
            
    frmResumen.Show vbModal
    
End Sub

Private Sub mnuArchivo_Salir_Click()
    Unload Me
End Sub

Private Sub mnuAyuda_Acercade_Click()
    frmAcerca.Show vbModal
End Sub

Private Sub mnuAyuda_Busqueda_Click()
    Call SearchHelp
End Sub

Private Sub mnuAyuda_Email_Click()
    Shell_Email
End Sub

Private Sub mnuAyuda_Indice_Click()
    Call ShowHelpContents
End Sub

Private Sub mnuAyuda_Tips_Click()
    frmTip.Show vbModal
End Sub

Private Sub mnuAyuda_WebSite_Click()
    Shell_PaginaWeb
End Sub

Private Sub mnuEdicion_Agregar_Click()
    Call tbrMain_ButtonClick(tbrMain.Buttons("cmdAgregar"))
End Sub

Private Sub mnuEdicion_AgregarPath_Click()
    Call AgregarArchivosPath
End Sub

Private Sub mnuEdicion_Eliminar_Click()
    Call tbrMain_ButtonClick(tbrMain.Buttons("cmdRemover"))
End Sub


Private Sub mnuEdicion_Invertir_Click()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwArchivos.ListItems.Count
        lvwArchivos.ListItems(k).Selected = Not lvwArchivos.ListItems(k).Selected
    Next k
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuEdicion_Propiedades_Click()
    Call MostrarPropiedades
End Sub

Private Sub mnuEdicion_Quitar_Click()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwArchivos.ListItems.Count
        lvwArchivos.ListItems(k).Selected = False
    Next k
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuEdicion_seleccionar_Click()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    For k = 1 To lvwArchivos.ListItems.Count
        lvwArchivos.ListItems(k).Selected = True
    Next k
    
    Call Hourglass(hWnd, False)
    
End Sub

Private Sub mnuOpciones_ConfRespaldo_Click()
    
    Dim OldglbRecursive As Boolean
    
    OldglbRecursive = glbRecursive
    
    frmConfOpc.Show vbModal
    
    'si se cambio a recursivo then
    If Not OldglbRecursive And glbRecursive Then
        If glbArchivo <> "" Then
            Call Hourglass(hWnd, True)
            Call Habilitar(False)
            Call CargarArchivosDesdePath(PathArchivo(Proyecto.PathFisico))
            Call Habilitar(True)
            Call Hourglass(hWnd, False)
        End If
    End If
    
End Sub

Private Sub mnuOpciones_Idioma_Click()

    Dim archivo As String
    Dim Glosa As String
    
    Glosa = Ini.Leer(Me.Name, "Msg1", glbArchivoLng)
    Glosa = Glosa & Ini.Leer(Me.Name, "Msg2", glbArchivoLng)
        
    If Cdlg.VBGetOpenFileName(archivo, , , , , , Glosa, , App.Path, Ini.Leer(Me.Name, "Msg3", glbArchivoLng), "LNG", Me.hWnd) Then
        If archivo <> "" Then
            Call Ini.Grabar(C_INI, "general", "lenguaje", VBArchivoSinPath(archivo))
            glbArchivoLng = archivo
            
            'Call CargarPantalla(Me)
            
            Call ConfiguraMenus
            
            Call CargarProyectosExplorados
        End If
    End If
    
End Sub

Private Sub MyHelpCallBack_MenuHelp(ByVal MenuText As String, ByVal MenuHelp As String, ByVal Enabled As Boolean)
    stbMain.Panels(1).Text = MenuHelp
End Sub

Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim Msg As String
    
    Select Case Button.Key
        Case "cmdAbrir"
            Call mnuArchivo_Abrir_Click
        Case "cmdRespaldar"
            If glbArchivo = "" Then
                MsgBox "Debe seleccionar un proyecto.", vbCritical
                Exit Sub
            End If
            
            If ValidaArchivos() Then
                Msg = "Confirma respaldar proyecto."
                If Confirma(Msg) = vbYes Then
                    If RespaldaProyecto() Then
                        MsgBox "Proyecto respaldado con éxito!", vbInformation
                    End If
                End If
            Else
                MsgBox "Debe seleccionar archivos a respaldar.", vbCritical
            End If
        Case "cmdRemover"
            Call RemoverArchivo
        Case "cmdAgregar"
            Call AgregarArchivo
        Case "cmdPath"
            Call AgregarArchivosPath
        Case "cmdPropiedades"
            Call MostrarPropiedades
        Case "cmdOpc"
            frmConfOpc.Show vbModal
        Case "cmdNet"
            mnuAyuda_WebSite_Click
        Case "cmdEmail"
            mnuAyuda_Email_Click
        Case "cmdTip"
            frmTip.Show vbModal
        Case "cmdAyuda"
            
        Case "cmdSalir"
            Call mnuArchivo_Salir_Click
    End Select
    
End Sub

