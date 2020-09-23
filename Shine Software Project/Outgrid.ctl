VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.0#0"; "vbalGrid6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl OutGrid 
   BackColor       =   &H80000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   PropertyPages   =   "Outgrid.ctx":0000
   ScaleHeight     =   2940
   ScaleWidth      =   6825
   ToolboxBitmap   =   "Outgrid.ctx":002C
   Begin VB.ComboBox cmbRisorse 
      Height          =   315
      ItemData        =   "Outgrid.ctx":033E
      Left            =   4890
      List            =   "Outgrid.ctx":0340
      TabIndex        =   1
      Top             =   1380
      Visible         =   0   'False
      Width           =   1770
   End
   Begin VB.TextBox TxtSelezione 
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   15
      TabIndex        =   0
      Top             =   360
      Width           =   6630
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   165
      Top             =   2025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   63
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0342
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":049C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":05F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0750
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":08AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0B5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0E12
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":0F6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":10C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1220
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":137A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":14D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":162E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1788
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":18E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1A3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1B96
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1CF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":1FA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":20FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2258
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":23B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":250C
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2666
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":27C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":291A
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2BCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2D28
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2E82
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":2FDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":3136
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":3290
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":33EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":3544
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":369E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":3E70
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":3FCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4124
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":427E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":43D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4532
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":468C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":47E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4940
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4A9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4BF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4D4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":4EA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":5002
            Key             =   ""
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":515C
            Key             =   ""
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":52B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":5410
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":556A
            Key             =   ""
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":56C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":581E
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":5978
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":5AD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":5F24
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Outgrid.ctx":6376
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox TxtValore 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4890
      TabIndex        =   3
      Top             =   825
      Visible         =   0   'False
      Width           =   1215
   End
   Begin MSComCtl2.DTPicker Calendario 
      Height          =   300
      Left            =   4920
      TabIndex        =   2
      Top             =   1080
      Visible         =   0   'False
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   529
      _Version        =   393216
      Format          =   22806529
      CurrentDate     =   37454
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   62
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   63
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   50
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin vbAcceleratorGrid6.vbalGrid OutGrid 
      Height          =   2250
      Left            =   15
      TabIndex        =   5
      Top             =   660
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   3969
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
End
Attribute VB_Name = "OutGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'########################################################
'                       OutGrid                         #
'########################################################
'########################################################

'-----------------------------------------------------------
'Il componente simula il funzionamento di MsProject  e  come
'tutti i componenti  OLE di  Microsoft sono pieni  di  bug e
'sono oltremodo fortemente collegati all'applicazione.Questo
'componente non avrà mai tutte le  potenzialità dell'oggetto
'della Microsoft ma può essere reputato un inizio.
'-----------------------------------------------------------
Option Explicit
Public Enum Durata
    giorno
    Mese
    Trimestre
    Anno
End Enum

'Valori predefiniti proprietà:
Const Const_Debug = True
Const def_TipoDurata = giorno
Const def_ColoreSfondo = &H80000009
Const def_ConsideraPrimoGiorno = True
Const def_QuickIns = False

'Variabili Proprietà
Dim Pr_TipoDurata As Durata             'Seleziona se il conteggio deve avvenire per giorno, mese oppure anni.
Dim Pr_ConsideraPrimoGiorno As Boolean  'Imposta il conteggio della durata includendo il giorno di partenza Es: dal 11/04/2002 al 12/04/2002 - Giorni Considerati = 2
Dim Pr_QuickIns As Boolean              'Mostra / Nasconde la barra del testo per l'immissione semplificata delle attività


'Variabili interne


Private CodiceUnivoco As Long 'Codice Univoco
Private RowSelected As Long
Private ColSelected As Long

Private m_eSelOrder() As cShellSortOrderCOnstants
Private m_iSelCount As Long



'Inizio delle Proprietà
Public Property Get ColoreSfondo() As OLE_COLOR
    ColoreSfondo = OutGrid.BackColor
End Property

Public Property Let ColoreSfondo(ByVal New_ColoreSfondo As OLE_COLOR)
    OutGrid.BackColor() = New_ColoreSfondo
    PropertyChanged "ColoreSfondo"
End Property

Public Property Get TipoDurata() As Durata
    TipoDurata = Pr_TipoDurata
End Property

Public Property Let TipoDurata(ByVal New_TipoDurata As Durata)
    Pr_TipoDurata = New_TipoDurata
    PropertyChanged "TipoDurata"
End Property

Public Property Get ConsideraPrimoGiorno() As Boolean
Attribute ConsideraPrimoGiorno.VB_ProcData.VB_Invoke_Property = "Impostazioni_Generali"
    ConsideraPrimoGiorno = Pr_ConsideraPrimoGiorno
End Property

Public Property Let ConsideraPrimoGiorno(ByVal New_ConsideraPrimoGiorno As Boolean)
    Pr_ConsideraPrimoGiorno = New_ConsideraPrimoGiorno
    PropertyChanged "ConsideraPrimoGiorno"
End Property

Public Property Get QuickIns() As Boolean
    QuickIns = Pr_QuickIns
End Property

Public Property Let QuickIns(ByVal New_QuickIns As Boolean)
    Pr_QuickIns = New_QuickIns
    PropertyChanged "QuickIns"
    UserControl_Resize
End Property

Private Sub Calendario_CloseUp()
    Dim lRow As Long
    Dim lCol As Integer
      With OutGrid
         lRow = .SelectedRow
         lCol = .SelectedCol
         .CellText(lRow, lCol) = Calendario.Value
         .SetFocus
      End With
    Calendario.Visible = False


End Sub

'Fine delle Proprietà
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Pr_TipoDurata = PropBag.ReadProperty("TipoDurata", def_TipoDurata)
    Pr_ConsideraPrimoGiorno = PropBag.ReadProperty("ConsideraPrimoGiorno", def_ConsideraPrimoGiorno)
    Pr_QuickIns = PropBag.ReadProperty("QuickIns", def_QuickIns)
End Sub

Private Sub UserControl_Terminate()
     
    Set Risorse = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("TipoDurata", Pr_TipoDurata, def_TipoDurata)
    Call PropBag.WriteProperty("ConsideraPrimoGiorno", Pr_ConsideraPrimoGiorno, def_ConsideraPrimoGiorno)
    Call PropBag.WriteProperty("QuickIns", Pr_QuickIns, def_QuickIns)
End Sub



'Inizializzazione Controllo OutGrid:
Private Sub UserControl_InitProperties()
    TipoDurata = giorno
    OutGrid.BackColor = def_ColoreSfondo
    ConsideraPrimoGiorno = def_ConsideraPrimoGiorno
    QuickIns = def_QuickIns
    UserControl_Resize 'Dopo il QuickIns reimposto l'oggetto
End Sub



'Ridimensionamento OutGrid:
Private Sub UserControl_Resize()
    If UserControl.Height < 1500 Then
       UserControl.Height = 1500
    End If
    
    If UserControl.Width < 12620 Then
       UserControl.Width = 12620
    End If
    
    If Pr_QuickIns Then
        OutGrid.Top = Toolbar.Height + TxtSelezione.Height + 50
        OutGrid.Height = UserControl.Height - (Toolbar.Height + TxtSelezione.Height + 90)
        TxtSelezione.Width = UserControl.Width
      Else
        OutGrid.Top = Toolbar.Height + 50
        OutGrid.Height = UserControl.Height - (Toolbar.Height + 90)
    End If
    TxtSelezione.Visible = Pr_QuickIns
    OutGrid.Width = UserControl.Width
    

End Sub


Private Sub UserControl_Initialize()

    Dim sFnt As New StdFont
    sFnt.Bold = True

    
    Set Risorse = New Collection 'Inizializzo la collezione Risorse

    CodiceUnivoco = 1
    OutGrid.Redraw = False
    OutGrid.RowHeight(1) = 21
    OutGrid.MultiSelect = True
    OutGrid.DefaultRowHeight = 18
    OutGrid.HeaderFlat = True
    OutGrid.GridLines = True
    TxtSelezione.Visible = Pr_QuickIns
    
    OutGrid.AddColumn "IDAttività", "ID", , , 30, True, , , False
    OutGrid.AddColumn "info", "Info", , , 35
    OutGrid.AddColumn "nome_attivita", "Nome Attività", , , 300
    OutGrid.AddColumn "durata " & TipoDurata, "Durata", , , 58
    OutGrid.AddColumn "inizio", "Inizio", , , 85
    OutGrid.AddColumn "fine", "Fine", , , 85
    OutGrid.AddColumn "predecessori", "Predecessori", , , 90
    OutGrid.AddColumn "risorse", "Risorse", , , 100
    OutGrid.AddColumn "Livello", "Liv.", , , 40, Const_Debug
    OutGrid.AddColumn "Relazione", "Rel.", , , 40, Const_Debug
    OutGrid.AddColumn "CodiceUnivoco", "Cod.", , , 40, Const_Debug
    OutGrid.AddColumn "Riferimento", "Rif.", , , 40, Const_Debug
    OutGrid.AddColumn "Proprietà", "Prop.", , , 44, Const_Debug
    OutGrid.SetHeaders
    
    OutGrid.CellDetails 1, 1, "1", DT_CENTER, , &HE0E0E0, , sFnt     'IDAttività
    OutGrid.CellDetails 1, 2, "" 'Info
    OutGrid.CellDetails 1, 3, "" 'Nome Attività
    OutGrid.CellDetails 1, 4, "" 'durata
    OutGrid.CellDetails 1, 5, Date 'Inizio
    OutGrid.CellDetails 1, 6, Date 'fine
    OutGrid.CellDetails 1, 7, "" 'Predecessori
    OutGrid.CellDetails 1, 8, "" 'Risorse
    OutGrid.CellDetails 1, 9, "0" 'Livello
    OutGrid.CellDetails 1, 10, "" 'Relazione
    OutGrid.CellDetails 1, 11, 1 'Codice Univoco della Riga
    OutGrid.CellDetails 1, 12, "" 'Riferimento
    OutGrid.CellDetails 1, 13, "" 'Proprietà
    OutGrid.Redraw = True
    
    
End Sub
'Sezione Metodi:
Public Sub AddResource(ByVal id As String, ByVal cognomenome As String)
    Risorse.Add cognomenome, id
    cmbRisorse.AddItem cognomenome
End Sub

Public Sub AddTask(Optional ByRef BottoneINS As Boolean, Optional ByRef Attivita As String, Optional ByRef Durata As Integer, Optional ByRef Inizio As Date, Optional ByRef Fine As Date, Optional ByRef Predecessore As Integer, Optional ByRef Risorse, Optional ByRef Livello As Integer, Optional ByRef ImmissioneManuale As Boolean)
    Dim RowsNum As String
    Dim sFnt As New StdFont
    sFnt.Bold = True
    OutGrid.Redraw = False
    If ImmissioneManuale Then
        If OutGrid.CellText(OutGrid.Rows, 3) = "" Then
            RowsNum = OutGrid.Rows
          Else
            RowsNum = OutGrid.Rows + 1
            CodiceUnivoco = CodiceUnivoco + 1
        End If
        OutGrid.CellDetails RowsNum, 1, "" 'IDAttività
        OutGrid.CellDetails RowsNum, 2, "" 'Info
        OutGrid.CellDetails RowsNum, 3, Attivita  'Nome Attività
        OutGrid.CellDetails RowsNum, 4, Durata 'durata
        OutGrid.CellDetails RowsNum, 5, Inizio 'Inizio
        OutGrid.CellDetails RowsNum, 6, Fine 'fine
        OutGrid.CellDetails RowsNum, 7, Predecessore 'Predecessori
        OutGrid.CellDetails RowsNum, 8, Risorse 'Risorse
        OutGrid.CellDetails RowsNum, 9, Livello 'Livello
        OutGrid.CellDetails RowsNum, 10, ""  'Relazione
        OutGrid.CellDetails RowsNum, 11, CodiceUnivoco  'Codice Univoco della Riga
        OutGrid.CellDetails RowsNum, 12, ""  'Riferimento
        OutGrid.CellDetails 1, 13, "" 'Proprietà
        OutGrid.RowHeight(RowsNum) = 21 'Altezza riga
        OutGrid.CellIndent(RowsNum, 3) = Livello * 10
      Else
        If OutGrid.CellText(OutGrid.Rows, 3) <> "" Or BottoneINS Then
            RowsNum = OutGrid.Rows + 1
            CodiceUnivoco = CodiceUnivoco + 1
            OutGrid.CellDetails RowsNum, 1, RowsNum, DT_CENTER, , &HE0E0E0, , sFnt      'IDAttività
            OutGrid.CellDetails RowsNum, 2, "" 'Info
            OutGrid.CellDetails RowsNum, 3, "" 'Nome Attività
            OutGrid.CellDetails RowsNum, 4, "" 'durata
            OutGrid.CellDetails RowsNum, 5, Date 'Inizio
            OutGrid.CellDetails RowsNum, 6, Date 'fine
            OutGrid.CellDetails RowsNum, 7, "" 'Predecessori
            OutGrid.CellDetails RowsNum, 8, "" 'Risorse
            OutGrid.CellDetails RowsNum, 9, "0" 'Livello
            OutGrid.CellDetails RowsNum, 10, "" 'Relazione
            OutGrid.CellDetails RowsNum, 11, CodiceUnivoco 'Codice Univoco della Riga
            OutGrid.CellDetails RowsNum, 12, "" 'Riferimento
            OutGrid.CellDetails RowsNum, 13, "" 'Proprietà
            OutGrid.RowHeight(RowsNum) = 21 'Altezza riga
            
        End If
        
    End If
    
    OutGrid.Redraw = True

    
End Sub


Private Sub Calendario_LostFocus()
    Calendario.Visible = False
End Sub

Private Sub Calendario_Validate(Cancel As Boolean)

    
'    OutGrid.CellText(lrow, 4) = DateDiff("d", OutGrid.CellText(lrow, 5), OutGrid.CellText(lrow, 6)) & " gg."
End Sub

Private Sub cmbRisorse_Click()
    Dim lRow As Long
   If (cmbRisorse.ListIndex > -1) Then
      With OutGrid
         lRow = .SelectedRow
         .CellText(lRow, 8) = cmbRisorse.Text
         .SetFocus
      End With
      cmbRisorse.Visible = False
   End If
End Sub

Private Sub cmbRisorse_LostFocus()
    OutGrid.CellText(RowSelected, ColSelected) = cmbRisorse.Text
    cmbRisorse.Visible = False
End Sub

Private Sub OutGrid_CancelEdit()
    TxtValore.Visible = False
End Sub

Private Sub OutGrid_ColumnWidthChanged(ByVal lCol As Long, ByVal lWidth As Long, bCancel As Boolean)
    TxtValore.Visible = False
End Sub

Private Sub OutGrid_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    If lRow <> 0 Or lCol <> 0 Then
        TxtSelezione.Text = OutGrid.CellText(lRow, lCol)
    End If
End Sub

Private Sub OutGrid_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    Select Case KeyCode
        Case 116
            FrmRisorsa.Show 1
        Case 45 'INS
            Call InserimentoNuovaAttività
    End Select
End Sub
Private Sub FatherControl(ByRef Row As Long)
    Dim sFnt As New StdFont
    
    Select Case OutGrid.CellText(Row, 13)
        Case "P" 'Padre
            sFnt.Bold = True
            OutGrid.CellFont(Row, 3) = sFnt
            OutGrid.CellFont(Row, 4) = sFnt
            OutGrid.CellFont(Row, 5) = sFnt
            OutGrid.CellFont(Row, 6) = sFnt
            OutGrid.CellText(Row, 8) = ""
        Case Else
            sFnt.Bold = False
            OutGrid.CellFont(Row, 3) = sFnt
            OutGrid.CellFont(Row, 4) = sFnt
            OutGrid.CellFont(Row, 5) = sFnt
            OutGrid.CellFont(Row, 6) = sFnt
    End Select
    
End Sub
Private Sub InserimentoNuovaAttività()
    Dim Rows As Long
    Dim RowSel As Long
    Dim Contatore As Long
    Dim sFnt As New StdFont
    
    RowSel = OutGrid.SelectedRow + 1
    Rows = OutGrid.Rows
    OutGrid.Redraw = False
    For Contatore = Rows To (Rows - (Rows - RowSel)) Step -1
        OutGrid.CellText(Contatore, 3) = OutGrid.CellText(Contatore - 1, 3) 'Descrizione
        OutGrid.CellIndent(Contatore, 3) = OutGrid.CellIndent(Contatore - 1, 3) 'Indentazione della descrizione
        
        OutGrid.CellText(Contatore, 4) = OutGrid.CellText(Contatore - 1, 4) 'Durata
        OutGrid.CellText(Contatore, 5) = OutGrid.CellText(Contatore - 1, 5) 'Data d'inizio
        OutGrid.CellText(Contatore, 6) = OutGrid.CellText(Contatore - 1, 6) 'Data di fine
        OutGrid.CellText(Contatore, 7) = OutGrid.CellText(Contatore - 1, 7) 'Predecessore
        OutGrid.CellText(Contatore, 8) = OutGrid.CellText(Contatore - 1, 8) 'Risorse
        OutGrid.CellText(Contatore, 9) = OutGrid.CellText(Contatore - 1, 9) 'Livello
        OutGrid.CellText(Contatore, 10) = OutGrid.CellText(Contatore - 1, 10) 'Relazione
        OutGrid.CellText(Contatore, 11) = OutGrid.CellText(Contatore - 1, 11) 'Codice Univoco
        OutGrid.CellText(Contatore, 12) = OutGrid.CellText(Contatore - 1, 12) 'Riferimento
        OutGrid.CellText(Contatore, 13) = OutGrid.CellText(Contatore - 1, 13) 'Proprietà
        Call FatherControl(Contatore)
    Next
    OutGrid.CellText(RowSel - 1, 3) = ""
    OutGrid.CellText(RowSel - 1, 9) = OutGrid.CellIndent(Contatore, 3) 'Indentazione della descrizione
    OutGrid.CellText(RowSel - 1, 7) = ""
    OutGrid.CellText(RowSel - 1, 8) = ""
    OutGrid.CellText(RowSel - 1, 10) = ""
    OutGrid.CellText(RowSel - 1, 11) = CodiceUnivoco
    OutGrid.CellText(RowSel - 1, 12) = "" 'Riferimento
    OutGrid.CellText(RowSel - 1, 13) = "" 'Proprietà
    
    sFnt.Bold = False
    OutGrid.CellFont(RowSel - 1, 3) = sFnt
    OutGrid.CellFont(RowSel - 1, 4) = sFnt
    OutGrid.CellFont(RowSel - 1, 5) = sFnt
    OutGrid.CellFont(RowSel - 1, 6) = sFnt
    
    Call AddTask(True)
    Call RiordinaRiferimenti
    
    OutGrid.Redraw = True
    
      
    'Call RiordinaRighe
    
    
    
End Sub
Private Sub Outgrid_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
    Dim iArt As Long, iRow As Long, iType As Long, iArticle As Long, iLink As Long
    Dim Top As Integer
    OutGrid.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
    RowSelected = lRow
    ColSelected = lCol
    
    'Impostazione della posizione dell'altezza dell'oggetto Txtvalore:
    If Pr_QuickIns Then
        Top = TxtSelezione.Height + 90
       Else
        Top = 90
    End If
    
    Select Case lCol
        Case 3 'Descrizione
            With TxtValore
              .Move lLeft + 45, lTop + Top + Toolbar.Height, lWidth, lHeight - 40
              .Text = OutGrid.CellText(lRow, 3)
              .Visible = True
              .ZOrder
              .SelStart = 0
              .SelLength = Len(.Text)
              .SetFocus
            End With
        Case 4 'Durata
            If OutGrid.CellFont(lRow, 3).Bold = False Then 'Sfrutto il tipo di formattazione per individuare se l'attività è padre!
                With TxtValore
                  .Move lLeft + 45, lTop + Top + Toolbar.Height, lWidth, lHeight - 40
                  .Text = Val(OutGrid.CellText(lRow, 4))
                  .Visible = True
                  .ZOrder
                  .SelStart = 0
                  .SelLength = Len(.Text)
                  .SetFocus
                End With
            End If
        Case 5
            If OutGrid.CellFont(lRow, 3).Bold = False Then 'Sfrutto il tipo di formattazione per individuare se l'attività è padre!
                With Calendario
                  .Move lLeft + 40, lTop + Top + Toolbar.Height, lWidth
                  .Value = OutGrid.CellText(lRow, 5)
                  .Visible = True
                  .ZOrder
                  .SetFocus
                End With
            End If
        Case 6
            If OutGrid.CellFont(lRow, 3).Bold = False Then 'Sfrutto il tipo di formattazione per individuare se l'attività è padre!
                With Calendario
                  .Move lLeft + 80, lTop + Top + Toolbar.Height, lWidth
                  .Value = OutGrid.CellText(lRow, 5)
                  .Visible = True
                  .ZOrder
                  .SetFocus
                End With
            End If
        Case 7
            If OutGrid.CellFont(lRow, 3).Bold = False Then 'Sfrutto il tipo di formattazione per individuare se l'attività è padre!
                With TxtValore
                  .Move lLeft + 45, lTop + Top + Toolbar.Height, lWidth - 50, lHeight - 40
                  .Text = Val(OutGrid.CellText(lRow, 7))
                  .Visible = True
                  .ZOrder
                  .SelStart = 0
                  .SelLength = Len(.Text)
                  .SetFocus
                End With
            End If
        Case 8
            If OutGrid.CellFont(lRow, 3).Bold = False Then 'Sfrutto il tipo di formattazione per individuare se l'attività è padre!
                With cmbRisorse
                  .Move lLeft + 80, lTop + Top + Toolbar.Height, lWidth
                  .Text = OutGrid.CellText(lRow, 8)
                  .Visible = True
                  .ZOrder
                  .SetFocus
                End With
            End If
    End Select

End Sub
Private Sub InOut(ByRef InOut As Boolean)
    Dim Contatore As Long
    Dim Livello As Integer
    Dim A As Long
    OutGrid.Redraw = False
    For Contatore = 1 To OutGrid.Rows
        If OutGrid.CellText(Contatore, 3) <> "" Then 'Controllo se esiste un attività nella riga da indentare
            If OutGrid.CellSelected(Contatore, 3) Then 'Controllo se quale riga ho selezionato
                Livello = 0
                If InOut Then 'controllo l'indentazione dell'attività
                    If OutGrid.CellIndent(Contatore, 3) - OutGrid.CellIndent(Contatore - 1, 3) < 10 Then 'Controllo l'indentazione superflua
                        Livello = OutGrid.CellIndent(Contatore, 3) + 10
                      Else
                        Livello = OutGrid.CellIndent(Contatore, 3)
                    End If
                    Call DatePredecessori(Contatore)
                  Else
                    If OutGrid.CellIndent(Contatore, 3) > 0 Then 'Controllo se il livello è posto a zero
                        If OutGrid.CellIndent(Contatore, 3) > 0 Then 'Controllo l'indentazione superflua
                            Livello = OutGrid.CellIndent(Contatore, 3) - 10
                          Else
                            Livello = 0
                        End If
                    End If
                End If
                OutGrid.CellIndent(Contatore, 3) = Livello
                OutGrid.CellText(Contatore, 9) = Livello 'Registro il livello di indentazione in una colonna nascosta
                OutGrid.CellText(Contatore, 10) = RicercaPadre(Contatore, Livello) 'Registro la relazione in una colonna nascosta
                CaratteristicheAttivita Livello, Contatore
                Debug.Print "L'elemento " & OutGrid.CellText(Contatore, 3) & " ha come padre " & OutGrid.CellText(Contatore, 10)
                Call DatePredecessori(Contatore)
            End If
        End If
    Next
    OutGrid.Redraw = True

End Sub
Private Function RicercaPadre(ByRef lRow As Long, ByRef Livello As Integer) As Long
    Dim conta As Long
    For conta = lRow To 1 Step -1
        If CInt(Val(OutGrid.CellText(conta, 9))) < Livello Then
            RicercaPadre = OutGrid.CellText(conta, 11)
            Exit Function
        End If
    Next
End Function

'Riordina l'idAttività:
Private Sub RiordinaRighe()
    Dim ContaRighe As Long
    For ContaRighe = 1 To OutGrid.Rows
        OutGrid.CellText(ContaRighe, 1) = ContaRighe
    Next
End Sub

Private Function RiordinaRiferimenti()
    Dim ContaRighe As Long
    Dim CodiceUnivoco As Long
    Dim Riferimento As Long
    For ContaRighe = 1 To OutGrid.Rows
        If OutGrid.CellText(ContaRighe, 7) <> 0 And OutGrid.CellText(ContaRighe, 7) <> "" Then
            Riferimento = Val(OutGrid.CellText(ContaRighe, 12))
            OutGrid.CellText(ContaRighe, 7) = CercaCodiceUnivoco(Riferimento)
            If OutGrid.CellText(ContaRighe, 7) = 0 Then 'Elimino lo zero
                OutGrid.CellText(ContaRighe, 7) = ""
                OutGrid.CellText(ContaRighe, 12) = ""
            End If
        End If
    Next
End Function
Private Function CercaCodiceUnivoco(ByRef Riferimento As Long) As Long
    Dim ContaRighe As Long
    For ContaRighe = 1 To OutGrid.Rows
        If Riferimento = OutGrid.CellText(ContaRighe, 11) Then
            CercaCodiceUnivoco = OutGrid.CellText(ContaRighe, 1)
            Exit Function
        End If
    Next
End Function
Private Sub CaratteristicheAttivita(ByRef Livello As Integer, ByRef Row As Long)
    Dim sFnt As New StdFont
    Dim TotRow As Long
    TotRow = OutGrid.Rows
    
    
    If Row <> 1 Then
        If CInt(Val(OutGrid.CellText(Row, 9))) > CInt(Val(OutGrid.CellText(Row - 1, 9))) Then
            sFnt.Bold = True
            OutGrid.CellFont(Row - 1, 3) = sFnt
            OutGrid.CellFont(Row - 1, 3) = sFnt
            OutGrid.CellFont(Row - 1, 4) = sFnt
            OutGrid.CellFont(Row - 1, 5) = sFnt
            OutGrid.CellFont(Row - 1, 6) = sFnt
            OutGrid.CellText(Row - 1, 8) = "" 'Cancello le risorse poichè possono essere inserite solamente nelle attività!
            OutGrid.CellText(Row - 1, 13) = "P" 'Proprietà Padre
          Else
            sFnt.Bold = False
            OutGrid.CellFont(Row - 1, 3) = sFnt
            OutGrid.CellFont(Row - 1, 3) = sFnt
            OutGrid.CellFont(Row - 1, 4) = sFnt
            OutGrid.CellFont(Row - 1, 5) = sFnt
            OutGrid.CellFont(Row - 1, 6) = sFnt
            OutGrid.CellText(Row - 1, 13) = ""
        End If
        If TotRow > Row + 1 Then
            If CInt(Val(OutGrid.CellText(Row, 9))) < CInt(Val(OutGrid.CellText(Row + 1, 9))) Then
              sFnt.Bold = True
              OutGrid.CellFont(Row, 3) = sFnt
              OutGrid.CellFont(Row, 4) = sFnt
              OutGrid.CellFont(Row, 5) = sFnt
              OutGrid.CellFont(Row, 6) = sFnt
              OutGrid.CellText(Row, 8) = "" 'Cancello le risorse poichè possono essere inserite solamente nelle attività!
              OutGrid.CellText(Row, 13) = "P"  'Proprietà Padre
            Else
              sFnt.Bold = False
              OutGrid.CellFont(Row, 3) = sFnt
              OutGrid.CellFont(Row, 4) = sFnt
              OutGrid.CellFont(Row, 5) = sFnt
              OutGrid.CellFont(Row, 6) = sFnt
              OutGrid.CellText(Row, 13) = ""   'Proprietà Padre
            End If
        End If
    End If
  
End Sub

Private Sub OutGrid_SelectionChange(ByVal lRow As Long, ByVal lCol As Long)
    ControllaAbilitazioniBottoni lRow
End Sub
Private Sub ControllaAbilitazioniBottoni(ByRef Row As Long)
    If OutGrid.CellText(Row, 3) = "" Or OutGrid.CellText(Row, 9) = 0 Or OutGrid.CellText(Row, 9) = "" Then
        Toolbar.Buttons(1).Enabled = False
       Else
        Toolbar.Buttons(1).Enabled = True
        Toolbar.Buttons(2).Enabled = True
    End If
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    
    Dim lRow As Long
    Dim NumRisorse As Integer
    Dim RisorseAssegnate As String
    Dim AssegnaRisorse As Long
    
    
    lRow = OutGrid.SelectedRow
    Select Case Button.Index
        Case 1 'Indietro
            If OutGrid.CellText(lRow, 3) <> "" Then
                Call InOut(False)
                ControllaAbilitazioniBottoni lRow
            End If
        Case 2 'Avanti
            If OutGrid.CellText(lRow, 3) <> "" Then
                Call InOut(True)
                ControllaAbilitazioniBottoni lRow
            End If
        Case 3
            If OutGrid.Rows > 1 Then
                TxtValore.Visible = False
                If OutGrid.CellFont(lRow, 3).Bold Then
                    If MsgBox("L'attività " & OutGrid.CellText(lRow, 3) & " è un'attività di riepilogo." & vbCrLf & "Eliminandola saranno eliminate anche tutte le sue sottoattività!", vbInformation + vbYesNo, "Attenzione") = vbYes Then
                        Call CancellaRighe(OutGrid.CellText(lRow, 11))
                        OutGrid.RemoveRow CercaRigaDaCodiceUni(CLng(OutGrid.CellText(lRow, 11)))
                    End If
                  Else
                    OutGrid.RemoveRow lRow
                End If
                
                Call RiordinaRighe
                Call RiordinaRiferimenti
                Call CaratteristicheAttivita(OutGrid.CellIndent(lRow, 3), lRow)
                
                
              Else
                MsgBox "Attenzione non puoi cancellare la riga selezionata", vbInformation, "Messaggio"
            End If
        Case 4
            If lRow <> 0 Then
                FrmRisorsa.Show 1
                OutGrid.CellText(lRow, 8) = ""

                For AssegnaRisorse = 0 To UBound(Temp_Risorse)
                   OutGrid.CellText(lRow, 8) = OutGrid.CellText(lRow, 8) & Temp_Risorse(AssegnaRisorse) & ";"
                Next
                
            End If
            
        Case 5
            Dim m_sSelKey() As String
            Dim I As Long
            For I = 1 To OutGrid.Rows
                If OutGrid.RowVisible(I) Then
                    m_iSelCount = m_iSelCount + 1
                    ReDim Preserve m_sSelKey(1 To m_iSelCount) As String
                    m_sSelKey(m_iSelCount) = "Relazione"
                    ReDim Preserve m_eSelOrder(1 To m_iSelCount) As cShellSortOrderCOnstants
                    If (OutGrid.CellText(1, 2) = "Descending") Then
                       m_eSelOrder(m_iSelCount) = CCLOrderDescending
                    Else
                       m_eSelOrder(m_iSelCount) = CCLOrderAscending
                    End If
                End If
            Next
    End Select
End Sub
Private Function QRelazione(ByRef idrelazione As Long) As String
    Dim ContaRighe As Long
    Dim CodiceUni As String
    CodiceUni = ""
    For ContaRighe = 1 To OutGrid.Rows
        If idrelazione = OutGrid.CellText(ContaRighe, 10) Then
            CodiceUni = CodiceUni & OutGrid.CellText(ContaRighe, 11) & ";"
        End If
    Next
    If CodiceUni <> "" Then
        QRelazione = Left(CodiceUni, Len(CodiceUni) - 1)
      Else
        QRelazione = ""
    End If
End Function
Private Function CercaRigaDaCodiceUni(ByRef CodiceUnivoco As Long) As String
    Dim ContaRighe As Long
    For ContaRighe = 1 To OutGrid.Rows
        If CodiceUnivoco = OutGrid.CellText(ContaRighe, 11) Then
            CercaRigaDaCodiceUni = ContaRighe
        End If
    Next
End Function

Private Sub CancellaRighe(ByRef CodiceUnivoco As Long)
    Dim ListaCodiciUni
    Dim CodiceUni, ArrayCodici
    OutGrid.Redraw = False
    ListaCodiciUni = QRelazione(CodiceUnivoco)
    
    If ListaCodiciUni <> "" Then
        ArrayCodici = Split(ListaCodiciUni, ";")
        For Each CodiceUni In ArrayCodici
            Call CancellaRighe(CLng(CodiceUni))
            OutGrid.RemoveRow CercaRigaDaCodiceUni(CLng(CodiceUni))
        Next
    End If
    
    OutGrid.Redraw = True
End Sub
Private Sub TxtSelezione_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If OutGrid.SelectedCol = 3 Then
            OutGrid.CellText(OutGrid.SelectedRow, OutGrid.SelectedCol) = TxtSelezione.Text
            If OutGrid.CellText(OutGrid.SelectedRow, OutGrid.SelectedCol) <> "" Then
                Call AddTask
            End If
            TxtSelezione.Text = ""
        End If
    End If
End Sub



Private Sub TxtValore_KeyPress(KeyAscii As Integer)
    Dim Tipo As String
    Dim Riferimento As Long
    Select Case Pr_TipoDurata
        Case 0
            Tipo = "g"
        Case 1
            Tipo = "m"
        Case 2
            Tipo = "t"
        Case 3
            Tipo = "a"
    End Select
    
    Select Case OutGrid.SelectedCol
        Case 1
        Case 2
        Case 3 'Descrizione
            If KeyAscii = 13 Then
                If OutGrid.CellText(RowSelected, 1) = "" Then 'Se l'indice non è stato assegnato assegnalo!
                    OutGrid.CellText(RowSelected, 1) = OutGrid.Rows
                End If
                OutGrid.CellText(RowSelected, ColSelected) = TxtValore.Text
                OutGrid.CellText(RowSelected, 4) = "1 " & Tipo
                If OutGrid.CellText(RowSelected, ColSelected) <> "" Then
                    Call AddTask
                End If
                TxtSelezione.Text = TxtValore.Text
                TxtSelezione.SelStart = 0
                TxtSelezione.SelLength = Len(TxtSelezione.Text)
                TxtValore.Visible = False
            End If
        Case 4 'Conteggio Data
            If KeyAscii = 13 Then
                Call ConteggiaData
            End If
        Case 5
        Case 6
        Case 7 'Predecessore
            If KeyAscii = 13 Then
                If Val(TxtValore.Text) = 0 Then
                    OutGrid.CellText(RowSelected, 12) = ""
                End If
                OutGrid.CellText(RowSelected, ColSelected) = TxtValore.Text
                TxtValore.Visible = False
                If IsNumeric(TxtValore.Text) Then
                    If CInt(TxtValore.Text) < OutGrid.SelectedRow And CInt(TxtValore.Text) <> 0 Then
                        OutGrid.CellText(OutGrid.SelectedRow, 12) = OutGrid.CellText(TxtValore.Text, 11)
                      Else
                        MsgBox "Attenzione il valore non è consentito", vbExclamation, "Messaggio"
                        OutGrid.CellText(OutGrid.SelectedRow, 7) = ""
                    End If
                End If
            End If
    End Select
    
End Sub
Private Sub TxtValore_LostFocus()
    Dim Tipo As String
    Dim Riferimento As Long
    Select Case Pr_TipoDurata
        Case 0
            Tipo = "g"
        Case 1
            Tipo = "m"
        Case 2
            Tipo = "t"
        Case 3
            Tipo = "a"
    End Select
    
    Select Case ColSelected
        Case 3 'Descrizione
                If OutGrid.CellText(RowSelected, 1) = "" Then 'Se l'indice non è stato assegnato assegnalo!
                    OutGrid.CellText(RowSelected, 1) = OutGrid.Rows
                End If
                OutGrid.CellText(RowSelected, ColSelected) = TxtValore.Text
                OutGrid.CellText(RowSelected, 4) = "1 " & Tipo
                If OutGrid.CellText(RowSelected, ColSelected) <> "" Then
                    Call AddTask
                End If
                TxtSelezione.Text = TxtValore.Text
                TxtSelezione.SelStart = 0
                TxtSelezione.SelLength = Len(TxtSelezione.Text)
                TxtValore.Visible = False
        Case 4 'Conteggio Data
            Call ConteggiaData
        Case 7
            Call ControllaPredecessori(RowSelected)
        Case Else
            OutGrid.CellText(RowSelected, ColSelected) = TxtValore.Text
            If OutGrid.CellText(RowSelected, ColSelected) <> "" Then
                Call AddTask
            End If
            TxtValore.Visible = False
    End Select
End Sub
Private Sub ConteggiaData()
    Dim ValoreAggiunto As Integer
    Dim T_Durata As String
    Dim Tipo As String
    'Controllo se il valore è un numero:
    If IsNumeric(TxtValore.Text) = False Then
        TxtValore.Visible = False
        Exit Sub
      Else
        ValoreAggiunto = CInt(TxtValore.Text)
    End If
    
    If Pr_ConsideraPrimoGiorno Then
        ValoreAggiunto = ValoreAggiunto - 1
    End If
    
    'Controllo il tipo di conteggio che devo effettuare:
    Select Case Pr_TipoDurata
        Case giorno
            T_Durata = "d"
            Tipo = "g"
        Case Mese
            T_Durata = "m"
            Tipo = "m"
        Case Trimestre
            T_Durata = "q"
            Tipo = "t"
        Case Anno
            T_Durata = "yyyy"
            Tipo = "a"
    End Select

    
    OutGrid.CellText(RowSelected, 6) = DateAdd(T_Durata, ValoreAggiunto, OutGrid.CellText(RowSelected, 5))
    OutGrid.CellText(RowSelected, ColSelected) = TxtValore.Text & " " & Tipo
    TxtValore.Visible = False
End Sub

Private Sub DatePredecessori(ByRef Row As Long)
    Dim conta As Long
    Dim Padre As Long
    
    Padre = RicercaPadre(Row, OutGrid.CellIndent(Row, 3))
    
    If OutGrid.CellText(Row, 7) <> "" And (Val(OutGrid.CellText(Row - 1, 9)) < Val(OutGrid.CellText(Row, 9))) Then
        MsgBox "Impossibile collegare un'attività di riepilogo ad una delle sue sottoattività." & vbCrLf & vbCrLf & "Per collegare le due attività, annullare il rientro delle sottoattività rispetto all'attività di riepilogo, quindi procedere con il collegamento." & vbCrLf & vbCrLf & "Verrà cancellato automaticamente il collegamento!", vbExclamation
        OutGrid.CellText(Row, 7) = ""
        OutGrid.CellText(Row, 12) = ""
        Call InOut(False)
        Exit Sub
    End If
    
    'Controllo sui Predecessori
    'Tutti i figli annidati di una determinata attività non possono
    'assumere come predecessore un valore <= al valore indicato nel campo relazioni.
    
    If Padre = 0 Then
        Exit Sub
    End If
    OutGrid.CellText(Padre, 5) = OutGrid.CellText(Row, 5)
    OutGrid.CellText(Padre, 6) = OutGrid.CellText(Row, 6)

End Sub

Private Function ControllaPredecessori(ByRef Row As Long)

    If Val(OutGrid.CellText(Row, 10)) >= Val(TxtValore.Text) Then
        MsgBox "Impossibile collegare un'attività di riepilogo ad una delle sue sottoattività." & vbCrLf & vbCrLf & "Per collegare le due attività, annullare il rientro delle sottoattività rispetto all'attività di riepilogo, quindi procedere con il collegamento." & vbCrLf & vbCrLf & "Verrà cancellato automaticamente il collegamento!", vbExclamation
        OutGrid.CellText(Row, 7) = ""
        OutGrid.CellText(Row, 12) = ""
        TxtValore.Text = ""
        TxtValore.Visible = False
    End If

End Function


