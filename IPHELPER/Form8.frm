VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIPForwardTable 
   Caption         =   "IP Forward Table"
   ClientHeight    =   1200
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   2190
   LinkTopic       =   "Form8"
   MDIChild        =   -1  'True
   ScaleHeight     =   1200
   ScaleWidth      =   2190
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   2143
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuCreateNew 
      Caption         =   "&Create New"
   End
   Begin VB.Menu mnuDelete 
      Caption         =   "&Delete"
   End
End
Attribute VB_Name = "frmIPForwardTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As New NetworkTables
Dim currentMouseRow As Long

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
'   Sets the column specified value on the current row!
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(318)
    Me.mnuCreateNew.Caption = LoadResString(364)
    Me.mnuDelete.Caption = LoadResString(365)
    Call Init ' Initializes the grid
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub mnuCreateNew_Click()
    ' Load the form to create a new forward entry
    Load frmNewIpForwardEntry
    ' Show the form to create a new forward entry
    frmNewIpForwardEntry.Show
End Sub

Private Sub mnuDelete_Click()
    If currentMouseRow > 0 Then ' If the current selected row isn't a fixed row!
        Dim AddRow() As MIB_IPFORWARDROW
        Dim RetVal As Long
        
        ' Get the IP forward Table rows
        AddRow = t.IpForwardTable
        
        ' Delete the row with the same index as the currently selected row
        RetVal = DeleteIpForwardEntry(AddRow(currentMouseRow - 1))
        
        If RetVal = 0 Then ' If it succeeds then
            Call Init ' Refresh the grid with the new rows
        Else
            ' Inform the user that it didn't work
            MsgBox LoadResString(156) & Chr(vbKeySpace) & GetDLLMessage(RetVal)
        End If
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Store the current selected row!
    currentMouseRow = MSFlexGrid1.MouseRow
End Sub

Private Sub Init()
    Dim AddRow() As MIB_IPFORWARDROW, i As Integer
    
    ' Get the IP forward table rows
    AddRow = t.IpForwardTable

    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        .Cols = 14
        
        .Rows = UBound(AddRow) + 2
        
        .Row = 0
        
        ' Set the fixed column headers
        Call SetColumnText(0, LoadResString(157))
        Call SetColumnText(1, LoadResString(158))
        Call SetColumnText(2, LoadResString(159))
        Call SetColumnText(3, LoadResString(160))
        Call SetColumnText(4, LoadResString(161))
        Call SetColumnText(5, LoadResString(162))
        Call SetColumnText(6, LoadResString(163))
        Call SetColumnText(7, LoadResString(164))
        Call SetColumnText(8, LoadResString(165))
        Call SetColumnText(9, LoadResString(166))
        Call SetColumnText(10, LoadResString(167))
        Call SetColumnText(11, LoadResString(168))
        Call SetColumnText(12, LoadResString(169))
        Call SetColumnText(13, LoadResString(170))

        ' For each row
        For i = 0 To UBound(AddRow)
    
            .Row = i + 1
            
            ' Set the column data with the data from the IP forward Table row
            Call SetColumnText(0, AddRow(i).dwForwardAge)
            Call SetColumnText(1, modWinsock.getascip(AddRow(i).dwForwardDest))
            Call SetColumnText(2, AddRow(i).dwForwardIfIndex)
            Call SetColumnText(3, modWinsock.getascip(AddRow(i).dwForwardMask))
            Call SetColumnText(4, AddRow(i).dwForwardMetric1)
            Call SetColumnText(5, AddRow(i).dwForwardMetric2)
            Call SetColumnText(6, AddRow(i).dwForwardMetric3)
            Call SetColumnText(7, AddRow(i).dwForwardMetric4)
            Call SetColumnText(8, AddRow(i).dwForwardMetric5)
            Call SetColumnText(9, modWinsock.getascip(AddRow(i).dwForwardNextHop))
            Call SetColumnText(10, AddRow(i).dwForwardNextHopAS)
            Call SetColumnText(11, AddRow(i).dwForwardPolicy)
            Call SetColumnText(12, GetIPForwardProtocolString(AddRow(i).dwForwardProto))
            Call SetColumnText(13, GetIPForwardTypeString(AddRow(i).dwForwardType))
        Next i
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
        
        ' Start redrawing the grid again!
        .Redraw = True
    End With
End Sub
