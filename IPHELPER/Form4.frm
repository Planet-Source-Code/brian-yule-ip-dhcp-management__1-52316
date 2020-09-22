VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIpAddressTable 
   Caption         =   "Ip Address Table"
   ClientHeight    =   765
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   765
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
Attribute VB_Name = "frmIpAddressTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As NetworkTables

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
'   Sets the information in the current cell!
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(317)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Init()
    Dim AddRow() As MIB_IPADDRROW, i As Integer
    
    ' Create the network tables class
    Set t = New NetworkTables
    
    ' Get the rows of the IP Address Table
    AddRow = t.IpAddressTable
    
    If IsEmpty(AddRow) = False Then
        With MSFlexGrid1
            ' Stop redrawing the grid
            .Redraw = False
            .Cols = 5
            .Rows = UBound(AddRow) + 2
            
            .Row = 0
            
            ' Fill in the fixed column headers
            Call SetColumnText(0, LoadResString(151))
            Call SetColumnText(1, LoadResString(152))
            Call SetColumnText(2, LoadResString(153))
            Call SetColumnText(3, LoadResString(154))
            Call SetColumnText(4, LoadResString(155))
        
            ' For each row
            For i = 0 To UBound(AddRow)
        
                .Row = i + 1
                
                ' Set the IP Address details
                Call SetColumnText(0, modWinsock.getascip(AddRow(i).dwBCastAddr))
                Call SetColumnText(1, modWinsock.getascip(AddRow(i).dwAddr))
                Call SetColumnText(2, AddRow(i).dwIndex)
                Call SetColumnText(3, modWinsock.getascip(AddRow(i).dwMask))
                Call SetColumnText(4, AddRow(i).dwReasmSize)
            Next i
            
            ' Auto size the grid columns
            modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
        
            ' Start redrawing the grid again!
            .Redraw = True
        End With
    End If
    
    ' Release the network tables class
    Set t = Nothing
End Sub
