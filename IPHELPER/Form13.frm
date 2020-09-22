VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmICMPStatistics 
   Caption         =   "ICMP Statistics"
   ClientHeight    =   750
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   1965
   LinkTopic       =   "Form13"
   MDIChild        =   -1  'True
   ScaleHeight     =   750
   ScaleWidth      =   1965
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   1296
      _Version        =   393216
      AllowUserResizing=   1
   End
End
Attribute VB_Name = "frmICMPStatistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim u As ICMPStats

Private Sub AddRow(ByVal RowHeader As String, ByRef RowText As Variant)
' Adds rows to the grid and places the information into the row created
    With MSFlexGrid1
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .Col = 0
        .Text = RowHeader
        .Col = 1
        .Text = RowText
    End With
End Sub

Private Sub Form_Activate()
    Call Init
End Sub

Private Sub Form_Load()
    Me.Caption = LoadResString(314) ' Set the forms caption
    Call Init ' Initialize the grid!
End Sub

Private Sub Form_Resize()
    If WindowState <> vbMinimized Then
        ' Resize the grid to the same size as the form (rougly)!
        MSFlexGrid1.Width = Width - 105
        MSFlexGrid1.Height = Height - 465
    End If
End Sub

Private Sub Init()
    ' Create the ICMP statistics class
    Set u = New ICMPStats
    
    With MSFlexGrid1
        ' Stop redraw until we have added all the info.  This makes the grid load faster!
        .Redraw = False
        
        .Cols = 2
        .Rows = 1
        .Col = 1
        .Row = 0
        
        .Text = LoadResString(101)
        
        ' Add the details from the ICMP statistics class to the form!
        Call AddRow(LoadResString(102), u.InAddrMaskReps)
        Call AddRow(LoadResString(103), u.InAddrMasks)
        Call AddRow(LoadResString(104), u.InDestUnreachs)
        Call AddRow(LoadResString(105), u.InEchoReps)
        Call AddRow(LoadResString(106), u.InEchos)
        Call AddRow(LoadResString(107), u.InErrors)
        Call AddRow(LoadResString(108), u.InMsgs)
        Call AddRow(LoadResString(109), u.InParmProbs)
        Call AddRow(LoadResString(110), u.InRedirects)
        Call AddRow(LoadResString(111), u.InSrcQuenchs)
        Call AddRow(LoadResString(112), u.InTimeExcds)
        Call AddRow(LoadResString(113), u.InTimestampReps)
        Call AddRow(LoadResString(114), u.InTimestamps)
        Call AddRow(LoadResString(115), u.OutAddrMaskReps)
        Call AddRow(LoadResString(116), u.OutAddrMasks)
        Call AddRow(LoadResString(117), u.OutDestUnreachs)
        Call AddRow(LoadResString(118), u.OutEchoReps)
        Call AddRow(LoadResString(119), u.OutEchos)
        Call AddRow(LoadResString(120), u.OutErrors)
        Call AddRow(LoadResString(121), u.OutMsgs)
        Call AddRow(LoadResString(122), u.OutParmProbs)
        Call AddRow(LoadResString(123), u.OutRedirects)
        Call AddRow(LoadResString(124), u.OutSrcQuenchs)
        Call AddRow(LoadResString(125), u.OutTimeExcds)
        Call AddRow(LoadResString(126), u.OutTimestampReps)
        Call AddRow(LoadResString(127), u.OutTimestamps)

        ' Auto size the flex grid so you can see all the information without having to resize the columns
        Call AutoSizeFlexGrid(MSFlexGrid1)

        ' Start redrawing the grid again!
        .Redraw = True
    
    End With
    
    Set u = Nothing ' Unload the ICMP class
End Sub
