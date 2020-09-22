VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmInterfaceAdapters 
   Caption         =   "Interface adapters"
   ClientHeight    =   705
   ClientLeft      =   3255
   ClientTop       =   2970
   ClientWidth     =   2805
   LinkTopic       =   "Form10"
   MDIChild        =   -1  'True
   ScaleHeight     =   705
   ScaleWidth      =   2805
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
Attribute VB_Name = "frmInterfaceAdapters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private z As NetworkInterfaces

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
'   Sets the information in the current cell!
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(316)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (rougly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Init()
    Dim i As Integer
    
    ' Create the network interfaces class
    Set z = New NetworkInterfaces
    
    ' If there is more than one network interface then!
    If z.Count > 0 Then
        With MSFlexGrid1
            ' Stop redrawing the grid
            .Redraw = False
            
            .Cols = 2
            .Rows = z.Count + 1
    
            .Row = 0
    
            ' Set the fixed column headers
            Call SetColumnText(0, LoadResString(368))
            Call SetColumnText(1, LoadResString(150))
    
            ' For each network interface
            For i = 0 To z.Count - 1
                .Row = i + 1
    
                ' Output the network interface information
                Call SetColumnText(0, z.Item(Trim(Str(i))).Index)
                Call SetColumnText(1, z.Item(Trim(Str(i))).Name)
            Next i
    
            ' Auto size the grid columns.
            modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
    
            ' Start redrawing the grid again
            .Redraw = True
        End With
        
        ' Release the network interfaces class
        Set z = Nothing
    End If
End Sub
