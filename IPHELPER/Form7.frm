VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIFTable 
   Caption         =   "IF Table"
   ClientHeight    =   690
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form7"
   MDIChild        =   -1  'True
   ScaleHeight     =   690
   ScaleWidth      =   1725
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1296
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmIFTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const DATE_SECONDS = "s"

Dim t As NetworkTables

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
'   Sets the information in the current cell!
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(315)
    Call Init ' Initialize the grid!
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (rougly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub
Private Sub Init()
    Dim AddRow() As MIB_IFROW, i As Integer
    
    ' Create the Network Tables Class
    Set t = New NetworkTables
    
    ' Ge the rows of the If Table
    AddRow = t.IfTable

    If IsEmpty(AddRow) = False Then
        With MSFlexGrid1
            ' Stop redrawing the grid!
            .Redraw = False
            
            .Cols = 22
            
            .Rows = UBound(AddRow) + 2
            
            .Row = 0
            
            ' Set the fixed column information
            Call SetColumnText(0, LoadResString(128))
            Call SetColumnText(1, LoadResString(129))
            Call SetColumnText(2, LoadResString(130))
            Call SetColumnText(3, LoadResString(131))
            Call SetColumnText(4, LoadResString(132))
            Call SetColumnText(5, LoadResString(133))
            Call SetColumnText(6, LoadResString(134))
            Call SetColumnText(7, LoadResString(135))
            Call SetColumnText(8, LoadResString(136))
            Call SetColumnText(9, LoadResString(137))
            Call SetColumnText(10, LoadResString(138))
            Call SetColumnText(11, LoadResString(139))
            Call SetColumnText(12, LoadResString(140))
            Call SetColumnText(13, LoadResString(141))
            Call SetColumnText(14, LoadResString(142))
            Call SetColumnText(15, LoadResString(143))
            Call SetColumnText(16, LoadResString(144))
            Call SetColumnText(17, LoadResString(145))
            Call SetColumnText(18, LoadResString(146))
            Call SetColumnText(19, LoadResString(147))
            Call SetColumnText(20, LoadResString(148))
            Call SetColumnText(21, LoadResString(149))
        
            For i = 0 To UBound(AddRow) ' For each row of information
                ' Change the row!
                .Row = i + 1
                
                ' Set the column information with the information from the If Table
                Call SetColumnText(0, TrimNull(StrConv(AddRow(i).bDescr, vbUnicode)))
                Call SetColumnText(1, t.ConvAddress(AddRow(i).bPhysAddr, AddRow(i).dwPhysAddrLen))
                Call SetColumnText(2, IIf(AddRow(i).dwAdminStatus, True, False))
                Call SetColumnText(3, AddRow(i).dwIndex)
                Call SetColumnText(4, AddRow(i).dwInDiscards)
                Call SetColumnText(5, AddRow(i).dwInErrors)
                Call SetColumnText(6, AddRow(i).dwInNUcastPkts)
                Call SetColumnText(7, AddRow(i).dwInOctets)
                Call SetColumnText(8, AddRow(i).dwInUcastPkts)
                Call SetColumnText(9, AddRow(i).dwInUnknownProtos)
                Call SetColumnText(10, AddRow(i).dwLastChange)
                Call SetColumnText(11, AddRow(i).dwMtu)
                Call SetColumnText(12, GetOperStatusString(AddRow(i).dwOperStatus))
                Call SetColumnText(13, AddRow(i).dwOutDiscards)
                Call SetColumnText(14, AddRow(i).dwOutErrors)
                Call SetColumnText(15, AddRow(i).dwOutNUcastPkts)
                Call SetColumnText(16, AddRow(i).dwOutOctets)
                Call SetColumnText(17, AddRow(i).dwOutQLen)
                Call SetColumnText(18, AddRow(i).dwOutUcastPkts)
                Call SetColumnText(19, AddRow(i).dwSpeed)
                Call SetColumnText(20, GetIfTypeString(AddRow(i).dwType))
                Call SetColumnText(21, TrimNull(StrConv(AddRow(i).wszName, vbUnicode)))
            
            Next i
            
            ' Auto size the grid columns
            modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
            
            ' Start redrawing the grid again
            .Redraw = True
        End With
    End If
    
    ' Release the network tables class
    Set t = Nothing
End Sub
