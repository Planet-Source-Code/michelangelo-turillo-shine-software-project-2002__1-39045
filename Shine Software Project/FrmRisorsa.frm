VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmRisorsa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Assegna Risorse"
   ClientHeight    =   3030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4065
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3030
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList 
      Left            =   3090
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRisorsa.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdChiudi 
      Caption         =   "&Chiudi"
      Height          =   315
      Left            =   2925
      TabIndex        =   2
      Top             =   795
      Width           =   1140
   End
   Begin VB.TextBox TxtRisorsa 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   45
      TabIndex        =   1
      Top             =   135
      Width           =   2850
   End
   Begin VB.CommandButton CmdAssegna 
      Caption         =   "&Assegna"
      Height          =   315
      Left            =   2925
      TabIndex        =   0
      Top             =   420
      Width           =   1140
   End
   Begin vbAcceleratorGrid6.vbalGrid GrdRisorse 
      Height          =   2535
      Left            =   45
      TabIndex        =   3
      Top             =   435
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   4471
      BackgroundPictureHeight=   16
      BackgroundPictureWidth=   16
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   2
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
      DefaultRowHeight=   30
   End
End
Attribute VB_Name = "FrmRisorsa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Row As Long
Public SelectedRow As Integer
Private Sub CmdAssegna_Click()
    Dim Indice As Long
    Indice = 0
    For conta = 1 To GrdRisorse.Rows
        If GrdRisorse.CellIcon(conta, 1) = 0 Then
            ReDim Preserve Temp_Risorse(Indice)
            Temp_Risorse(Indice) = GrdRisorse.CellText(conta, 2)
            Indice = Indice + 1
        End If
    Next
    Unload Me
End Sub

Private Sub CmdChiudi_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GrdRisorse.Redraw = False
    GrdRisorse.MultiSelect = True
    GrdRisorse.DefaultRowHeight = 18
    GrdRisorse.HeaderFlat = True
    GrdRisorse.GridLines = True
    GrdRisorse.ImageList = ImageList
    GrdRisorse.AddColumn "icona", , , , 22
    GrdRisorse.AddColumn "cognomenome", "Cognome & Nome", , , 160
    GrdRisorse.SetHeaders
    GrdRisorse.Redraw = False
    
    For CicloRisorse = 1 To Risorse.Count
        GrdRisorse.CellDetails CicloRisorse, 2, Risorse(CicloRisorse)
    Next
    GrdRisorse.Redraw = True
End Sub

Public Sub ResourceList(ByRef Risorsa As String)
    Row = Row + 1
    GrdRisorse.CellDetails Row, 2, Risorsa
    GrdRisorse.Redraw = True
End Sub

Private Sub GrdRisorse_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If GrdRisorse.CellIcon(lRow, 1) = -1 Then
        GrdRisorse.CellDetails GrdRisorse.SelectedRow, 1, , DT_RIGHT, 0
      Else
        GrdRisorse.CellDetails GrdRisorse.SelectedRow, 1, , DT_RIGHT
    End If
    GrdRisorse.Redraw = True
    TxtRisorsa.Text = GrdRisorse.CellText(lRow, 2)
End Sub
