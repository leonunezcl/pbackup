VERSION 5.00
Begin VB.Form frmAcerca 
   BackColor       =   &H80000006&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de ..."
   ClientHeight    =   5940
   ClientLeft      =   1710
   ClientTop       =   1905
   ClientWidth     =   5535
   Icon            =   "Acerca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   369
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   5955
      Left            =   0
      ScaleHeight     =   395
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   3
      Top             =   -15
      Width           =   360
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   4365
      TabIndex        =   0
      Top             =   5475
      Width           =   915
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Explorer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   3180
      MouseIcon       =   "Acerca.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Tag             =   "http://www.vbsoftware.cl/vbpexplorer.html"
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Project Backup fue analizado con :"
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   600
      TabIndex        =   7
      Top             =   4995
      Width           =   2490
   End
   Begin VB.Label lblGlosa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Respalda proyectos creados con Microsoft Visual Basic."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1050
      TabIndex        =   2
      Top             =   375
      Width           =   4245
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   540
      Picture         =   "Acerca.frx":0BD4
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      Caption         =   "Project Backup Home Page"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   630
      MouseIcon       =   "Acerca.frx":149E
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Tag             =   "http://www.vbsoftware.cl/pbackup.html"
      Top             =   5265
      Width           =   2355
   End
   Begin VB.Label lblURL 
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.vbsoftware.cl"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   630
      MouseIcon       =   "Acerca.frx":17A8
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Tag             =   "http://www.vbsoftware.cl/"
      Top             =   5715
      Width           =   2370
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2000-2002 Luis Núñez Ibarra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   630
      MouseIcon       =   "Acerca.frx":1AB2
      MousePointer    =   99  'Custom
      TabIndex        =   4
      Tag             =   "http://www.vbsoftware.cl/autor.html"
      Top             =   5490
      Width           =   3105
   End
   Begin VB.Label lblDescrip 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Explora , Documenta , Respalda , Visualiza , Limpia , Optimiza aplicaciones creadas con Visual Basic 3,4,5,6."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   585
      Left            =   705
      TabIndex        =   1
      Top             =   990
      Width           =   4770
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAcerca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient






Private Sub cmd_Click()
    
    Unload Me
    
End Sub


Private Sub Form_Load()

    Dim Msg As String
    Dim Texto As String
    Dim k As Integer
    Dim td As Variant
    
    CenterWindow hWnd
       
    Msg = "Creado por Luis Núñez Ibarra." & vbNewLine
    Msg = Msg & "Todos los derechos reservados." & vbNewLine
    Msg = Msg & "Santiago de Chile 2000-2002" & vbNewLine
    Msg = Msg & "" & vbNewLine
    Msg = Msg & "Respalda todos los archivos relacionados de un proyecto visual basic." & vbNewLine
    Msg = Msg & "" & vbNewLine
    Msg = Msg & "Se distribuye libre de cargo alguno bajo el término de distribución postcardware." & vbNewLine
    Msg = Msg & "" & vbNewLine
    Msg = Msg & "Si le gusta este software apreciaria mucho que me enviara una postal de su ciudad a la siguiente dirección :" & vbNewLine
    Msg = Msg & "" & vbNewLine
    Msg = Msg & "        Avda Vicuña Mackenna 7000" & vbNewLine
    Msg = Msg & "        Depto 204-B" & vbNewLine
    Msg = Msg & "        Santiago de Chile" & vbNewLine
    Msg = Msg & "" & vbNewLine
    Msg = Msg & "VBSoftware no se hace responsable por algún daño ocasionado por el uso de esta aplicación." & vbNewLine
        
    lblDescrip.Caption = Msg
        
    lblURL.Tag = C_WEB_PAGE
            
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(App.Title & " " & App.Major & "." & App.Minor & "." & App.Revision, picDraw)
    
    picDraw.Refresh
                
End Sub


Private Sub Form_Unload(Cancel As Integer)

    If Not gbInicio Then
        gbInicio = True
        frmMain.Show
    End If
        
    Set frmAcerca = Nothing
    
End Sub


Private Sub lblCopyright_Click()
    pShell lblCopyright.Tag, hWnd
End Sub

Private Sub lblProduct_Click()
    pShell C_WEB_PAGE_PE, hWnd
End Sub


Private Sub lblURL_Click()
    pShell C_WEB_PAGE, hWnd
End Sub


