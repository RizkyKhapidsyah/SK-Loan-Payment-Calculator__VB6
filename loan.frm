VERSION 5.00
Begin VB.Form frmLoan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Calculator"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "loan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPayment 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2655
      MaxLength       =   3
      TabIndex        =   9
      Text            =   "295"
      Top             =   855
      Width           =   1185
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   180
      TabIndex        =   6
      Top             =   1305
      Width           =   5820
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "&Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4140
      TabIndex        =   5
      Top             =   5310
      Width           =   1860
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "&Calculate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   225
      TabIndex        =   4
      Top             =   5310
      Width           =   1860
   End
   Begin VB.TextBox txtInterest 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2655
      TabIndex        =   3
      Text            =   "21"
      Top             =   495
      Width           =   735
   End
   Begin VB.TextBox txtPrincipal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2655
      TabIndex        =   1
      Text            =   "7500"
      Top             =   135
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT AMOUNT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   180
      TabIndex        =   8
      Top             =   855
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   3465
      TabIndex        =   7
      Top             =   540
      Width           =   210
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INTEREST RATE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   180
      TabIndex        =   2
      Top             =   495
      Width           =   2430
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "LOAN AMOUNT:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   2430
   End
End
Attribute VB_Name = "frmLoan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()
    Dim total As Double
    Dim principal As Double
    Dim payment As Double
    Dim totalPrincipal As Double
    Dim interest As Double
    Dim interestRate As Double
    Dim totalInterest As Double
    Dim balance As Double
    Dim count As Long
    
    If txtPrincipal.Text = "" Then
        MsgBox "Enter the loan amount"
        Exit Sub
    End If
    
    If txtInterest.Text = "" Then
        MsgBox "Enter the interest rate"
        Exit Sub
    End If
    
    List1.Clear
    List1.AddItem "Month" & Chr$(9) & "Principal" & Chr$(9) & "Interest" & Chr$(9) & "Balance"
    List1.AddItem ""
    
    balance = Val(txtPrincipal.Text)
    interestRate = Val(txtInterest.Text) / 100
    payment = Val(txtPayment.Text)
    count = 0
    totalPrincipal = 0
    totalInterest = 0
    
    While balance > 0
        count = count + 1
        interest = Round(interestRate * balance / 12, 2)
        If balance < payment Then payment = balance
        principal = payment - interest
        balance = balance - principal
        
        List1.AddItem Format$(count, "0") & Chr$(9) _
            & Format$(principal, "standard") & Chr$(9) & Chr$(9) _
            & Format$(interest, "standard") & Chr$(9) & Chr$(9) _
            & Format$(balance, "standard")
        
        totalPrincipal = totalPrincipal + principal
        totalInterest = totalInterest + interest
    Wend
    
    List1.AddItem ""
    List1.AddItem "Total principal paid:    " & Format$(totalPrincipal, "currency")
    List1.AddItem "Total interest paid:     " & Format$(totalInterest, "currency")
    List1.AddItem "Total amount paid:       " & Format$(totalPrincipal + totalInterest, "currency")
    List1.AddItem "Compound interest rate:  " & Format$(100 - (((Val(txtPrincipal.Text) / (totalPrincipal + totalInterest)) * 100)), "standard") & "%"
    List1.AddItem ""
End Sub

Private Sub cmdQuit_Click()
    End
End Sub
