VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Simultaneous Equation Solver  [SVD Demo]"
   ClientHeight    =   8385
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   8385
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDisplay 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   59
      Top             =   120
      Width           =   6015
   End
   Begin VB.Frame Frame2 
      Caption         =   "Equations"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   6015
      Begin VB.CommandButton cmdSolveEchelon 
         Caption         =   "Solve Echelon"
         Height          =   375
         Left            =   3000
         TabIndex        =   61
         Top             =   3240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSolveSVD 
         Caption         =   "SVD Solve"
         Height          =   375
         Left            =   4440
         TabIndex        =   60
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   7
         Left            =   3240
         TabIndex        =   58
         Text            =   "0"
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   6
         Left            =   3240
         TabIndex        =   57
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   5
         Left            =   3240
         TabIndex        =   56
         Text            =   "0"
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   4
         Left            =   3240
         TabIndex        =   55
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   3
         Left            =   3240
         TabIndex        =   54
         Text            =   "0"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   2
         Left            =   3240
         TabIndex        =   53
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   1
         Left            =   3240
         TabIndex        =   52
         Text            =   "0"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   7
         Left            =   1800
         TabIndex        =   51
         Text            =   "0"
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   6
         Left            =   1800
         TabIndex        =   50
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   5
         Left            =   1800
         TabIndex        =   49
         Text            =   "0"
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   4
         Left            =   1800
         TabIndex        =   48
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   47
         Text            =   "0"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   46
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   45
         Text            =   "0"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   7
         Left            =   360
         TabIndex        =   44
         Text            =   "0"
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   6
         Left            =   360
         TabIndex        =   43
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   5
         Left            =   360
         TabIndex        =   42
         Text            =   "0"
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   4
         Left            =   360
         TabIndex        =   41
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   3
         Left            =   360
         TabIndex        =   40
         Text            =   "0"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   2
         Left            =   360
         TabIndex        =   39
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   38
         Text            =   "0"
         Top             =   780
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   7
         Left            =   4800
         TabIndex        =   37
         Text            =   "0"
         Top             =   2580
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   6
         Left            =   4800
         TabIndex        =   36
         Text            =   "0"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   5
         Left            =   4800
         TabIndex        =   35
         Text            =   "0"
         Top             =   1980
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   4
         Left            =   4800
         TabIndex        =   34
         Text            =   "0"
         Top             =   1680
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   3
         Left            =   4800
         TabIndex        =   33
         Text            =   "0"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   32
         Text            =   "0"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   1
         Left            =   4800
         TabIndex        =   31
         Text            =   "0"
         Top             =   780
         Width           =   975
      End
      Begin VB.ComboBox cboNumEquations 
         Height          =   315
         ItemData        =   "Form1.frx":0000
         Left            =   2040
         List            =   "Form1.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txtSolution 
         Height          =   285
         Index           =   0
         Left            =   4800
         TabIndex        =   6
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTerm2 
         Height          =   285
         Index           =   0
         Left            =   3240
         TabIndex        =   4
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTerm1 
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtTerm0 
         Height          =   285
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Text            =   "0"
         Top             =   480
         Width           =   975
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   5760
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label lblNumEquations 
         Caption         =   "Number of Equations:"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   7
         Left            =   4200
         TabIndex        =   28
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   6
         Left            =   4200
         TabIndex        =   27
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   5
         Left            =   4200
         TabIndex        =   26
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   4
         Left            =   4200
         TabIndex        =   25
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   3
         Left            =   4200
         TabIndex        =   24
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   23
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   22
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblZ 
         Caption         =   "* Z ="
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   21
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   7
         Left            =   2760
         TabIndex        =   20
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   19
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   5
         Left            =   2760
         TabIndex        =   18
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   3
         Left            =   2760
         TabIndex        =   16
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   15
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   14
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   7
         Left            =   1320
         TabIndex        =   13
         Top             =   2580
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   6
         Left            =   1320
         TabIndex        =   12
         Top             =   2280
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   11
         Top             =   1980
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   4
         Left            =   1320
         TabIndex        =   10
         Top             =   1680
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   9
         Top             =   1380
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   8
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   7
         Top             =   780
         Width           =   375
      End
      Begin VB.Label lblY 
         Caption         =   "* Y +"
         Height          =   255
         Index           =   0
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblX 
         Caption         =   "* X +"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 0

Private Const NumericalFormat        As String = "0.#####0"



'======================================================================================
'                         VALIDATE TEXTBOX VALUES ARE NUMERIC
'======================================================================================
Private Sub txtTerm0_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtTerm0(Index).Text) Then
        Call MsgBox(Chr$(34) & txtTerm0(Index).Text & Chr$(34) & vbCrLf & "This is not a valid numeric value", vbInformation, "Non-Numeric Value")
        txtTerm0(Index).SelStart = 0
        txtTerm0(Index).SelLength = Len(txtTerm0(Index).Text)
        Cancel = True
    End If
End Sub
Private Sub txtTerm1_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtTerm1(Index).Text) Then
        Call MsgBox(Chr$(34) & txtTerm1(Index).Text & Chr$(34) & vbCrLf & "This is not a valid numeric value", vbInformation, "Non-Numeric Value")
        txtTerm1(Index).SelStart = 0
        txtTerm1(Index).SelLength = Len(txtTerm1(Index).Text)
        Cancel = True
    End If
End Sub
Private Sub txtTerm2_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtTerm2(Index).Text) Then
        Call MsgBox(Chr$(34) & txtTerm2(Index).Text & Chr$(34) & vbCrLf & "This is not a valid numeric value", vbInformation, "Non-Numeric Value")
        txtTerm2(Index).SelStart = 0
        txtTerm2(Index).SelLength = Len(txtTerm2(Index).Text)
        Cancel = True
    End If
End Sub
Private Sub txtSolution_Validate(Index As Integer, Cancel As Boolean)
    If Not IsNumeric(txtSolution(Index).Text) Then
        Call MsgBox(Chr$(34) & txtSolution(Index).Text & Chr$(34) & vbCrLf & "This is not a valid numeric value", vbInformation, "Non-Numeric Value")
        txtSolution(Index).SelStart = 0
        txtSolution(Index).SelLength = Len(txtSolution(Index).Text)
        Cancel = True
    End If
End Sub



'======================================================================================
'                      TRANSFER TEXTBOX VALUES TO/FROM DOUBLE ARRAYS
'======================================================================================
Private Sub GetTextValues(A() As Double, B() As Double)
    Dim i As Long 'A * x = B
    Dim m As Long
    Dim n As Long

    m = cboNumEquations.ListIndex
    n = 2
    ReDim A(m, n)
    ReDim B(m)

    For i = 0 To m
        A(i, 0) = Val(txtTerm0(i).Text)
        A(i, 1) = Val(txtTerm1(i).Text)
        A(i, 2) = Val(txtTerm2(i).Text)
        B(i) = Val(txtSolution(i).Text)
    Next i
End Sub
Private Sub PutTextValues(A() As Double, B() As Double)
    Dim i As Long
    Dim m As Long
    Dim n As Long

    m = UBound(A, 1) ' also = UBound(B)
    n = UBound(A, 2)
    cboNumEquations.ListIndex = m

    For i = 0 To m
        txtTerm0(i).Text = CStr(A(i, 0))
        txtTerm1(i).Text = CStr(A(i, 1))
        txtTerm2(i).Text = CStr(A(i, 2))
        txtSolution(i).Text = CStr(B(i))
    Next i
End Sub



'======================================================================================
'                            SOLUTION TEXT FORMATTING
'======================================================================================
Private Sub DisplayEquations(A() As Double, B() As Double)
    Dim i       As Long 'Formatting text is never pretty
    Dim j       As Long
    Dim m       As Long
    Dim n       As Long
    Dim s       As String
    Dim t       As String
    Dim Value   As Double
    Dim Sign    As String
    Dim l       As Boolean

    m = UBound(A, 1)
    n = UBound(A, 2)

    For i = 0 To m
        s = ""
        l = False
        For j = 0 To n
            t = Chr$(90 + j - n) ' ...X, Y, Z
            Value = Abs(A(i, j))
            If A(i, j) < 0# Then
                Sign = "-"
            Else
                If l Then
                    Sign = "+"
                Else
                    Sign = " "
                End If
            End If
            Select Case Value
            Case 0#
            Case 1#
                If l Then
                    s = s & " " & Sign & " " & t
                Else
                    s = s & " " & Sign & " " & t
                End If
                l = True
            Case Else
               s = s & " " & Sign & " " & Format(Value) & "*" & t
               l = True
            End Select
        Next j
        If Len(s) = 0 Then s = "   0"
        s = s & " = " & Format(B(i))
        AddItem s
    Next i
    AddItem ""

End Sub
Private Sub DisplaySolutionVector(W() As Double, V() As Double, x() As Double)
    Dim i       As Long
    Dim j       As Long
    Dim n       As Long
    Dim t       As Long
    Dim ss()    As String

    n = UBound(V, 1)

    ReDim ss(n)
    For i = 0 To n
        ss(i) = "   " & Chr$(90 - n + i) & " = " & _
        IIf(x(i) < 0, "- ", "  ") & _
        Format(Abs(x(i)), NumericalFormat)
    Next i
    For i = 0 To n
        If W(i) = 0# Then
            t = t + 1
            For j = 0 To n
                ss(j) = ss(j) & IIf(V(j, i) < 0, " - ", " + ") & _
                Format(Abs(V(j, i)), NumericalFormat) & " * t" & CStr(t)
            Next j
        End If
    Next i
    AddItem "Solution Vector:"
    For i = 0 To n
        AddItem ss(i)
    Next i
    AddItem ""

End Sub
Public Sub DisplayResults(A() As Double, B() As Double, x() As Double)
    Dim i As Long
    Dim j As Long
    Dim m As Long
    Dim n As Long
    Dim E1 As Double
    Dim E2 As Double
    Dim Sum As Double
    Dim s As String

    m = UBound(A, 1) ' Also, m = UBound(B)
    n = UBound(A, 2)

    AddItem "Equat" + vbTab + "Expected" + vbTab + "Actual" + vbTab + vbTab + "Error"
    For i = 0 To m
        s = ""
        Sum = 0#
        For j = 0 To n
            Sum = Sum + A(i, j) * x(j)
        Next j
        E1 = Sum - B(i)
        E2 = E2 + E1 * E1
        s = Format(i + 1) + vbTab + Format(B(i), NumericalFormat) _
                        + vbTab + Format(Sum, NumericalFormat) _
                        + vbTab + Format(E1, NumericalFormat)
        AddItem s
    Next i
    AddItem "Sum Square Error: " + Format(E2)
    AddItem ""

End Sub
Private Sub DisplayMatrix(A() As Double)
    Dim i       As Long
    Dim j       As Long
    Dim m       As Long
    Dim n       As Long
    Dim s       As String

    m = UBound(A, 1)
    n = UBound(A, 2)

    For i = 0 To m
        s = ""
        For j = 0 To n
            s = s & Format(A(i, j)) & "    "
        Next j
        AddItem s
    Next i
    AddItem ""
End Sub


'======================================================================================
'                                INITIALIZATION / MISC
'======================================================================================
Private Sub Form_Load()
    Dim A() As Double ' Some default equations
    Dim B() As Double

    ReDim A(2, 2)
    ReDim B(2)

    A(0, 0) = 1#       '   X -  Y +  3Z = 4
    A(0, 1) = -1#
    A(0, 2) = 3#
    B(0) = 4#

    A(1, 0) = -3#      '- 3X + 4Y - 10Z = 10
    A(1, 1) = 4#
    A(1, 2) = -10#
    B(1) = 10#

    A(2, 0) = 2#       '  2X +  Y +   Z = 2
    A(2, 1) = 1#
    A(2, 2) = 1#
    B(2) = 2#

    PutTextValues A, B

End Sub
Public Sub AddItem(Item As String)
    txtDisplay.Text = txtDisplay.Text + vbCrLf + Item
    txtDisplay.Refresh
End Sub
Private Sub cboNumEquations_Click()
    Dim i As Long
    Dim e As Boolean
    For i = 0 To 7
        e = cboNumEquations.ListIndex >= i
        txtTerm0(i).Visible = e
        txtTerm1(i).Visible = e
        txtTerm2(i).Visible = e
        lblX(i).Visible = e
        lblY(i).Visible = e
        lblZ(i).Visible = e
        txtSolution(i).Visible = e
    Next i
End Sub



'======================================================================================
'                                COMMAND BUTTONS
'======================================================================================
Private Sub cmdSolveSVD_Click()
    Dim c       As Double
    Dim i       As Long
    Dim A()     As Double
    Dim B()     As Double
    Dim x()     As Double
    Dim U()     As Double
    Dim V()     As Double
    Dim W()     As Double

    GetTextValues A, B
    U = A 'Note for VB5 users: This makes U a copy of the 2D array A
    SVD U, W, V
    SVDEDIT W
    SVDBACKSUBSTITUTE U, W, V, B, x

    txtDisplay.Text = ""
    AddItem "Working Equations:"
    DisplayEquations A, B
    DisplaySolutionVector W, V, x
    DisplayResults A, B, x
    AddItem ""
    Erase A
End Sub

Private Sub cmdSolveEchelon_Click()
    Dim A() As Double
    Dim B() As Double

    GetTextValues A, B
    txtDisplay.Text = ""
    AddItem "Working Equations:"
    DisplayEquations A, B

    MatrixEchelon A, B
    AddItem "Reduced Equations:"
    DisplayEquations A, B
End Sub
