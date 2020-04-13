VERSION 5.00
Begin VB.Form frmConfOpc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configurar opciones de respaldo"
   ClientHeight    =   2565
   ClientLeft      =   3540
   ClientTop       =   3465
   ClientWidth     =   6375
   HelpContextID   =   220
   Icon            =   "frmConfOpc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   425
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   2565
      Left            =   30
      ScaleHeight     =   169
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   10
      Top             =   0
      Width           =   360
   End
   Begin VB.Frame fra 
      Caption         =   "Opciones Anexas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Index           =   1
      Left            =   435
      TabIndex        =   5
      Top             =   1065
      Width           =   4515
      Begin VB.CheckBox chkHlp 
         Caption         =   "Agregar archivos de ayuda (.hlp , .chm)"
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   975
         Width           =   4095
      End
      Begin VB.CheckBox chkIni 
         Caption         =   "Agregar archivos de configuración (.ini)"
         Height          =   255
         Left            =   135
         TabIndex        =   8
         Top             =   735
         Width           =   4095
      End
      Begin VB.CheckBox chkImg 
         Caption         =   "Agregar archivos de imagenes (.ico , .bmp , .gif , .jpg)"
         Height          =   255
         Left            =   135
         TabIndex        =   7
         Top             =   480
         Width           =   4095
      End
      Begin VB.CheckBox chkRpt 
         Caption         =   "Comprobar existencia de reportes Cristal Reports"
         Height          =   255
         Left            =   135
         TabIndex        =   6
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5070
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5070
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Frame fra 
      Caption         =   "Opciones de Compresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Index           =   0
      Left            =   420
      TabIndex        =   0
      Top             =   60
      Width           =   4530
      Begin VB.CheckBox chkSavInfo 
         Caption         =   "Guardar información de directorio"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   585
         Width           =   3150
      End
      Begin VB.CheckBox chkIncSub 
         Caption         =   "Incluir información de subdirectorios"
         Height          =   255
         Left            =   165
         TabIndex        =   2
         Top             =   330
         Width           =   3150
      End
   End
End
Attribute VB_Name = "frmConfOpc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient
Private Sub cmdAceptar_Click()

    Call Ini.Grabar(C_INI, "opciones", "recursive", chkIncSub.Value)
    Call Ini.Grabar(C_INI, "opciones", "savepath", chkSavInfo.Value)
    Call Ini.Grabar(C_INI, "opciones", "reportes", chkRpt.Value)
    Call Ini.Grabar(C_INI, "opciones", "imagenes", chkImg.Value)
    Call Ini.Grabar(C_INI, "opciones", "configuracion", chkIni.Value)
    Call Ini.Grabar(C_INI, "opciones", "ayuda", chkHlp.Value)
    
    If chkIncSub.Value = 1 Then glbRecursive = True Else glbRecursive = False
    If chkSavInfo.Value = 1 Then glbSavePath = True Else glbSavePath = False
    
    If chkRpt.Value = 1 Then glbRpt = True Else glbRpt = False
    If chkImg.Value = 1 Then glbIma = True Else glbIma = False
    If chkIni.Value = 1 Then glbIni = True Else glbIni = False
    If chkHlp.Value = 1 Then glbHlp = True Else glbHlp = False
    
    Unload Me
        
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    CenterWindow hWnd
    
    If glbRecursive Then chkIncSub.Value = 1 Else chkIncSub.Value = 0
    If glbSavePath Then chkSavInfo.Value = 1 Else chkSavInfo.Value = 0
    
    If glbRpt Then chkRpt.Value = 1 Else chkRpt.Value = 0
    If glbIma Then chkImg.Value = 1 Else chkImg.Value = 0
    If glbIni Then chkIni.Value = 1 Else chkIni.Value = 0
    If glbHlp Then chkHlp.Value = 1 Else chkHlp.Value = 0
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff("Opciones", picDraw)
    
    picDraw.Refresh
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    Set mGradient = Nothing
    Set frmConfOpc = Nothing
    
End Sub


