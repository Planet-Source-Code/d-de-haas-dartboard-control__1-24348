VERSION 5.00
Object = "{26D0F692-856C-40D3-8F5F-7696CF070D47}#1.0#0"; "DartBoard.ocx"
Begin VB.Form frmDartBord 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dartboard .ocx Demo"
   ClientHeight    =   7260
   ClientLeft      =   1275
   ClientTop       =   930
   ClientWidth     =   9900
   Icon            =   "frmDartBord.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin DartBoard.Board Board1 
      Height          =   5775
      Left            =   240
      TabIndex        =   22
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   10186
   End
   Begin VB.Timer tmrSlide 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   4440
   End
   Begin VB.Frame fraSliders 
      Caption         =   "Sliders"
      Enabled         =   0   'False
      Height          =   975
      Left            =   7080
      TabIndex        =   0
      Top             =   4200
      Width           =   2535
      Begin VB.CommandButton cmdThrow 
         Caption         =   "Next"
         Height          =   375
         Index           =   1
         Left            =   1560
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.CommandButton cmdThrow 
         Caption         =   "Throw"
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox picVer 
      Height          =   5775
      Left            =   6120
      ScaleHeight     =   5715
      ScaleWidth      =   195
      TabIndex        =   21
      Top             =   240
      Width           =   255
   End
   Begin VB.PictureBox picHor 
      Height          =   255
      Left            =   240
      ScaleHeight     =   195
      ScaleWidth      =   5715
      TabIndex        =   20
      Top             =   6120
      Width           =   5775
   End
   Begin VB.Frame fraManual 
      Caption         =   "Manual"
      Enabled         =   0   'False
      Height          =   1695
      Left            =   7080
      TabIndex        =   17
      Top             =   5400
      Width           =   2535
      Begin VB.TextBox txtManual 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   600
         MaxLength       =   1
         TabIndex        =   7
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txtManual 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   1320
         MaxLength       =   2
         TabIndex        =   8
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdManual 
         Caption         =   "Throw"
         Height          =   375
         Left            =   720
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Value"
         Height          =   195
         Left            =   1320
         TabIndex        =   19
         Top             =   360
         Width           =   405
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ring"
         Height          =   195
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Control"
      Height          =   1455
      Left            =   7080
      TabIndex        =   16
      Top             =   240
      Width           =   2535
      Begin VB.OptionButton optControl 
         Caption         =   "Enter Ring and Value"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optControl 
         Caption         =   "Use Sliders"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton optControl 
         Caption         =   "MouseClick on Board"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Score Board"
      Height          =   2055
      Left            =   7080
      TabIndex        =   11
      Top             =   1920
      Width           =   2535
      Begin VB.CommandButton cmdClear 
         Caption         =   "&Clear"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   1440
         Width           =   735
      End
      Begin VB.Line Line1 
         X1              =   1320
         X2              =   240
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label lblTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label lblDart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   14
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblDart 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmDartBord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private bytSlide As Byte
Private X As Single
Private Y As Single
Private D As Integer

Private Sub cmdClear_Click()
    Board1.ResetBoard
    lblDart(0).Caption = ""
    lblDart(1).Caption = ""
    lblDart(2).Caption = ""
    lblTotal.Caption = ""
End Sub

Private Sub cmdManual_Click()
    ' Ring can be 1 - 3 for Values 1 to 20
    ' Ring can be 1 - 2 for Value 25 (bull)
    
    Board1.PlaceDart Val(txtManual(0).Text), Val(txtManual(1).Text)
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdThrow_Click(Index As Integer)
    Select Case Index
        Case 0
            If bytSlide = 2 Then
                tmrSlide.Enabled = False
                cmdThrow(0).Enabled = False
                cmdThrow(1).Enabled = True
                
                ' X - 0 to board.Width (in twips)
                ' Y - 0 to board.Heigth (in twips)
                Board1.PlaceDartXY X, Y
            Else
                bytSlide = bytSlide + 1
            End If
        
        Case 1
            cmdThrow(0).Enabled = True
            cmdThrow(1).Enabled = False
            tmrSlide.Enabled = True
            bytSlide = 1
            X = 0
            Y = 0
            
    End Select
End Sub

Private Sub Board1_DartThrown()
    Dim N As Integer
    Dim M As Integer
    
    Dim Nr As Integer
    Dim Ring As Byte
    Dim Value As Byte
    
    With Board1
        Nr = .DartNumber
        Ring = .Dart(Nr).Ring
        Value = .Dart(Nr).Value
        
        Select Case Value
            Case 1 To 20
                Select Case Ring
                    Case 1
                        lblDart(Nr - 1).Caption = "Single " & Value
                    Case 2
                        lblDart(Nr - 1).Caption = "Dubble " & Value
                    Case 3
                        lblDart(Nr - 1).Caption = "Tripple " & Value
                End Select
                
            Case 25
                Select Case Ring
                    Case 1
                        lblDart(Nr - 1).Caption = "Single Bull"
                    Case 2
                        lblDart(Nr - 1).Caption = "Dubble Bull"
                End Select
                
            Case Else
                lblDart(Nr - 1).Caption = "Nothing"
        End Select
                
        M = 0
        For N = 1 To Nr
            M = M + .Dart(N).TotalValue
        Next N
        lblTotal.Caption = M
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub optControl_Click(Index As Integer)
    Select Case Index
        Case 0
            Board1.MouseEnabled = True
            fraSliders.Enabled = False
            fraManual.Enabled = False
            
            picHor.Cls
            picVer.Cls
            picHor.Enabled = False
            picVer.Enabled = False
            tmrSlide.Enabled = False
            
        Case 1
            Board1.MouseEnabled = False
            fraSliders.Enabled = True
            fraManual.Enabled = False
            
            picHor.Enabled = True
            picVer.Enabled = True
            cmdThrow(0).Enabled = True
            cmdThrow(1).Enabled = False
            bytSlide = 1
            X = 0
            Y = 0
            tmrSlide.Enabled = True
            
        Case 2
            Board1.MouseEnabled = False
            fraSliders.Enabled = False
            fraManual.Enabled = True
            
            picHor.Cls
            picVer.Cls
            picHor.Enabled = False
            picVer.Enabled = False
            tmrSlide.Enabled = False
        
        End Select
End Sub

Private Sub tmrSlide_Timer()
    If D = 0 Then D = 50
    
    Select Case bytSlide
        Case 1 ' Ver
            If Y >= picVer.Height - D Then D = -D
            If Y <= 0 Then D = Abs(D)
            
            Y = Y + D
            
            picVer.Cls
            picVer.Line (0, Y)-(picVer.Width, Y + D), RGB(0, 0, 255), BF
        
        Case 2 ' Hor
            If X >= picHor.Width - D Then D = -D
            If X <= 0 Then D = Abs(D)
            
            X = X + D
            
            picHor.Cls
            picHor.Line (X, 0)-(X + D, picHor.Height), RGB(0, 0, 255), BF
                    
    End Select
    
    tmrSlide.Enabled = True
End Sub

Private Sub txtManual_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case 8, 48 To 57
            ' Leave KeyAscii for what it is
        Case Else
            KeyAscii = 0
    End Select
End Sub
