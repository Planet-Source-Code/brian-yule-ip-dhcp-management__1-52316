VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTCPTable 
   Caption         =   "TCP Table"
   ClientHeight    =   975
   ClientLeft      =   60
   ClientTop       =   645
   ClientWidth     =   2040
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   2040
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   1720
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuDeleteTCB 
      Caption         =   "&Delete"
   End
End
Attribute VB_Name = "frmTCPTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim t As New NetworkTables
Dim currentMouseRow As Long

Private Enum COLUMNS
    LOCAL_ADDRESS = 0
    LOCAL_PORT
    REMOTE_ADDRESS
    REMOTE_PORT
    STATE
End Enum

Private Sub SetColumnText(ByVal Column As Integer, ByRef colVal As Variant)
    ' Set the column value of the current row
    MSFlexGrid1.Col = Column
    MSFlexGrid1.Text = colVal
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(325)
    Me.mnuDeleteTCB.Caption = LoadResString(366)
    Call Init ' Initialize the grid
End Sub

Private Sub Form_Activate()
    Call Init
End Sub


Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the size of the form (roughly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub mnuDeleteTCB_Click()
    Dim RetVal As Long
    If currentMouseRow > 0 Then ' If the current selected row is not a fixed row then
        ' Kill the connection
        RetVal = t.DeleteTCB(currentMouseRow - 1)
        If RetVal = 0 Then ' If the connection was killed then
            Call Init ' Refresh the grid
            currentMouseRow = 0
        Else
            ' Inform the user that this item cannot be killed
            MsgBox LoadResString(240) & Chr(vbKeySpace) & GetDLLMessage(RetVal)
        End If
    End If
End Sub

Private Sub MSFlexGrid1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If the state of the current mouse row is not listening and the row is not fixed then
    If MSFlexGrid1.TextMatrix(currentMouseRow, COLUMNS.STATE) <> GetState(MIB_TCP_STATE_LISTEN) Then
        ' Save the current selected row
        currentMouseRow = MSFlexGrid1.MouseRow
    Else
        currentMouseRow = 0
    End If
End Sub

Private Function Init()
    Dim AddRow() As MIB_TCPROW, i As Integer
    
    ' Get the TCP Table rows
    AddRow = t.TcpTable

    With MSFlexGrid1
        ' Stop redrawing the grid
        .Redraw = False
        .Cols = 5
        .Rows = UBound(AddRow) + 2
        .Row = 0
        
        ' Se the fixed column headers
        Call SetColumnText(COLUMNS.LOCAL_ADDRESS, LoadResString(241))
        Call SetColumnText(COLUMNS.LOCAL_PORT, LoadResString(242))
        Call SetColumnText(COLUMNS.REMOTE_ADDRESS, LoadResString(243))
        Call SetColumnText(COLUMNS.REMOTE_PORT, LoadResString(244))
        Call SetColumnText(COLUMNS.STATE, LoadResString(245))
    
        ' For each TCP Table row
        For i = 0 To UBound(AddRow)
    
            .Row = i + 1
            
            ' Set the column data with the TCP Table row data
            Call SetColumnText(COLUMNS.LOCAL_ADDRESS, modWinsock.getascip(AddRow(i).dwLocalAddr)) ' Convert the long network address to a text address
            Call SetColumnText(COLUMNS.LOCAL_PORT, modWinsock.GetPort(AddRow(i).dwLocalPort)) ' Conver the port number from a network port to an application port number
            Call SetColumnText(COLUMNS.REMOTE_ADDRESS, modWinsock.getascip(AddRow(i).dwRemoteAddr)) ' Convert the long network address to a text address
            Call SetColumnText(COLUMNS.REMOTE_PORT, modWinsock.GetPort(AddRow(i).dwRemotePort)) ' Conver the port number from a network port to an application port number
            Call SetColumnText(COLUMNS.STATE, GetState(AddRow(i).dwState)) ' Get the string representing the current state
            
        Next i
        
        ' Auto size the grid columns
        modGridAutoSize.AutoSizeFlexGrid MSFlexGrid1
    
        ' Start redrawing the grid
        .Redraw = True
    End With
End Function
