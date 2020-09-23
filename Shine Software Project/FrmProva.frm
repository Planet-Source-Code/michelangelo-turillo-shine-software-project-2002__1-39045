VERSION 5.00
Object = "{073261CE-02EE-4E9A-9529-D6C3250AC308}#7.0#0"; "OutGrid.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Out Grid"
   ClientHeight    =   9570
   ClientLeft      =   465
   ClientTop       =   2235
   ClientWidth     =   12705
   LinkTopic       =   "Form1"
   ScaleHeight     =   9570
   ScaleWidth      =   12705
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin PrjOutGrid.OutGrid OutGrid1 
      Height          =   4995
      Left            =   -15
      TabIndex        =   14
      Top             =   105
      Width           =   12615
      _extentx        =   22251
      _extenty        =   8811
   End
   Begin VB.Frame Frame1 
      Height          =   3690
      Left            =   75
      TabIndex        =   0
      Top             =   5775
      Width           =   4995
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4410
         TabIndex        =   12
         Text            =   "0"
         Top             =   255
         Width           =   390
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Nuova Attività"
         Height          =   330
         Left            =   3300
         TabIndex        =   11
         Top             =   3150
         Width           =   1605
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   1830
         Left            =   825
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "FrmProva.frx":0000
         Top             =   1680
         Width           =   2325
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1260
         TabIndex        =   7
         Text            =   "0"
         Top             =   1320
         Width           =   390
      End
      Begin MSComCtl2.DTPicker data 
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   975
         Width           =   1485
         _ExtentX        =   2619
         _ExtentY        =   503
         _Version        =   393216
         Format          =   22806529
         CurrentDate     =   37465
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         TabIndex        =   3
         Text            =   "2"
         Top             =   615
         Width           =   405
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   780
         TabIndex        =   1
         Text            =   "Progetto"
         Top             =   285
         Width           =   1470
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Livello :"
         Height          =   195
         Left            =   3840
         TabIndex        =   13
         Top             =   300
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Risorse :"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   1695
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Predecessore :"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Data inizio :"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Durata :"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   630
         Width           =   570
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Attività :"
         Height          =   195
         Left            =   165
         TabIndex        =   2
         Top             =   300
         Width           =   570
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
    If Check1.Value Then
        
    End If
End Sub

Private Sub Command1_Click()
    OutGrid1.AddTask False, Text1.Text, Text2.Text, data.Value, , Text4.Text, Text5.Text, Text3.Text, True
End Sub

Private Sub Form_Load()
    OutGrid1.AddResource "1", "Turillo Donatella"
    OutGrid1.AddResource "2", "Turillo Michelangelo"
    OutGrid1.AddResource "3", "Turillo Caterina"
    OutGrid1.AddResource "4", "Turillo Federica"
    OutGrid1.AddResource "5", "Turillo Andrea"
    OutGrid1.AddResource "6", "Arcidiacono Maria Grazia"
    
End Sub

Private Sub Form_Resize()
    OutGrid1.Width = Me.Width - 200
End Sub

