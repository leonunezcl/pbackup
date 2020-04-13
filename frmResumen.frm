VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmResumen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resumen de archivos"
   ClientHeight    =   2655
   ClientLeft      =   3240
   ClientTop       =   5400
   ClientWidth     =   6225
   Icon            =   "frmResumen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   177
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtTotal 
      Height          =   255
      Left            =   3390
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   60
      Width           =   1425
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
      Left            =   4905
      TabIndex        =   3
      Top             =   345
      Width           =   1215
   End
   Begin VB.PictureBox picDraw 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H00C0FFFF&
      Height          =   2640
      Left            =   0
      ScaleHeight     =   174
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   22
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin MSComctlLib.ListView lvwResumen 
      Height          =   2310
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   4075
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nº"
         Object.Width           =   1058
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tipo de archivo"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tamaño"
         Object.Width           =   2646
      EndProperty
   End
   Begin MSComctlLib.ImageList imgFiles 
      Left            =   4980
      Top             =   1170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   105
      Width           =   450
   End
   Begin VB.Label lblSumario 
      AutoSize        =   -1  'True
      Caption         =   "Sumario del respaldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   420
      TabIndex        =   1
      Top             =   105
      Width           =   1785
   End
End
Attribute VB_Name = "frmResumen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGradient As New clsGradient

Private Type eExt
    Ext As String
    Size As Double
End Type

Private arr_ext() As eExt
'cargar resumen de respaldo
Private Sub CargaResumen()

    Dim k As Integer
    Dim j As Integer
    Dim total As Double
    Dim archivo As String
    Dim Size As Double
    Dim Ext As String
    Dim found As Boolean
    Dim Path As String
    Dim Indice As Integer
    
    'ciclar x los archivos a respaldar
    With frmMain.lvwArchivos
        For k = 1 To .ListItems.Count
            'archivo a respaldar
            archivo = .ListItems(k).SubItems(1)
            'path del archivo
            Path = .ListItems(k).SubItems(5)
            'extension del archivo
            Ext = UCase$(frmMain.ExtensionArchivo(archivo))
            'tamaño
            Size = VBGetFileSize(Path & archivo)
            
            'total de tamaño
            total = total + Size
            
            'verificar si archivo existe
            found = False
            For j = 1 To UBound(arr_ext)
                If arr_ext(j).Ext = Ext Then
                    arr_ext(j).Size = arr_ext(j).Size + Size
                    found = True
                    Exit For
                End If
            Next j
            
            'se encontro el archivo ?
            If Not found Then
                ReDim Preserve arr_ext(UBound(arr_ext) + 1)
                arr_ext(UBound(arr_ext)).Ext = Ext
                arr_ext(UBound(arr_ext)).Size = Size
            End If
        Next k
    End With
    
    'cargar los archivos en el listview
    For k = 1 To UBound(arr_ext)
        Indice = IndiceImagen(arr_ext(k).Ext, "")
        lvwResumen.ListItems.Add , "k" & k, k, Indice, Indice
        lvwResumen.ListItems("k" & k).SubItems(1) = arr_ext(k).Ext
        lvwResumen.ListItems("k" & k).SubItems(2) = arr_ext(k).Size
    Next k
    
    txtTotal.Text = total & " Kbytes"
    
End Sub

Private Sub cmdAceptar_Click()
    Unload Me
End Sub


Private Sub Form_Load()

    Dim k As Integer
    
    Call Hourglass(hWnd, True)
    
    Call CenterWindow(hWnd)
    
    ReDim arr_ext(0)
    
    'iniciar listimage ?
    imgFiles.ListImages.Clear
    imgFiles.ImageHeight = 16
    imgFiles.ImageWidth = 16
    
    For k = 1 To frmMain.imgFiles.ListImages.Count
        Call imgFiles.ListImages.Add(, frmMain.imgFiles.ListImages(k).Key, frmMain.imgFiles.ListImages(k).Picture)
    Next k
    
    'setear nueva imagen
    lvwResumen.View = lvwReport
    lvwResumen.Icons = imgFiles
    lvwResumen.SmallIcons = imgFiles
        
    'cargar resumen
    Call CargaResumen
    
    With mGradient
        .Angle = 90 '.Angle
        .Color1 = 16744448
        .Color2 = 0
        .Draw picDraw
    End With
        
    Call FontStuff(Me.Caption, picDraw)
    
    picDraw.Refresh
    
    Call Hourglass(hWnd, False)
    
End Sub


'indice de la imagen en picture
Private Function IndiceImagen(ByVal Ext As String, ByVal File As String) As Integer

    Dim ret As Integer
    Dim k As Integer
        
    'ciclar x las imagenes
    For k = 1 To imgFiles.ListImages.Count
        If UCase$(imgFiles.ListImages(k).Key) = UCase$(Ext) Then
            ret = k
            Exit For
        End If
    Next k
        
    IndiceImagen = ret
    
End Function
Private Sub Form_Unload(Cancel As Integer)
    Set frmResumen = Nothing
End Sub


