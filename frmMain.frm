VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "LAN Test - Powersoft Programming"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4995
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   4995
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   3300
      Width           =   3615
   End
   Begin MSComctlLib.TreeView tvLAN 
      Height          =   3135
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   5530
      _Version        =   393217
      Indentation     =   0
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilsTreeview"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsTreeview 
      Left            =   60
      Top             =   4260
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0842
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRefresh_Click()
    Dim cComputers As New clsComputers
    Dim cDomains As New clsDomains
    Dim lCurrentNode As Long
    Dim lK As Long
    Dim lX As Long
        
    '// Add LAN node
    tvLAN.Nodes.Clear
    tvLAN.Nodes.Add , , "LAN", "LAN", 1
    
    '// Enumerate Domains
    For lK = 1 To cDomains.GetCount
        tvLAN.Nodes.Add "LAN", tvwChild, cDomains.GetItem(lK), cDomains.GetItem(lK), 2
        
        '// Save Node Position (always the last, since
        '// sorting is disabled)
        lCurrentNode = tvLAN.Nodes.Count
        
        '// Enumerate Computers in Domain
        cComputers.Domain = cDomains.GetItem(lK)
        cComputers.Refresh
        
        For lX = 1 To cComputers.GetCount
            tvLAN.Nodes.Add cDomains.GetItem(lK), tvwChild, cComputers.GetItem(lX), cComputers.GetItem(lX), 3
        Next lX
        
        '// Expand Domain view
        tvLAN.Nodes.Item(lCurrentNode).Expanded = True
    Next lK
    
    '// Expand LAN view
    tvLAN.Nodes.Item(1).Expanded = True
End Sub

Private Sub Form_Load()
    '// Refresh
    cmdRefresh_Click
End Sub

Private Sub Form_Resize()
    On Local Error Resume Next
    
    '// Resize Treeview
    With tvLAN
        .Top = 60
        .Left = 60
        .Width = frmMain.ScaleWidth - 120
        .Height = frmMain.ScaleHeight - cmdRefresh.Height - 180
    End With
    
    '// Resize Refresh Button
    With cmdRefresh
        .Top = tvLAN.Top + tvLAN.Height + 60
        .Left = 60
        .Width = tvLAN.Width
    End With
End Sub


