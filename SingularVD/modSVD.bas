Attribute VB_Name = "modSVD"

Option Explicit
Option Base 0

'======================================================================================
'                        SINGULAR VALUE DECOMPOSITION MODULE
'
'  Author:  John Korejwa  <korejwa@tiac.net>
'
'  Version: 1.0  [14 / January / 2003]
'           [planetsourcecode.com submission]
'
'  Resubmitted to PSC on [8 / November / 2003]
'======================================================================================



Private Const MaxDouble As Double = 1.79769313486231E+308 'Maximum Double Value
Private Const MinDouble As Double = 4.94065645841247E-324 'Minimum _Positive_ Double Value
Private Const TOL       As Double = 10 ^ 13               'Tolerance for editing singular values



'======================================================================================
'                           SINGULAR VALUE DECOMPOSITION
'======================================================================================
Public Sub SVD(U() As Double, W() As Double, V() As Double)
    Dim i       As Long    'Given a matrix A(0 to m, 0 to n)
    Dim j       As Long    'Compute singular value decomposition
    Dim k       As Long    'A = U * W * V'             (where V' is the Transpose of V)
    Dim l       As Long
    Dim m       As Long    'The diagonal matrix of singular values W are output as vector W(0 ... n)
    Dim n       As Long    'The matrix V (not V') is output as matrix V(0 ... n, 0 ... n)
    Dim p       As Long    'Matrix A is input as 'U', and overwritten with U data
    Dim q       As Long
    Dim t       As Long

    Dim c       As Double  'OK to declare all variables Static if desired
    Dim d       As Double  '  [all are explicitly initialized]
    Dim f       As Double
    Dim g       As Double
    Dim h       As Double
    Dim s       As Double
    Dim x       As Double
    Dim y       As Double
    Dim z       As Double

    Dim e       As Boolean
    Dim r()     As Double

    m = UBound(U, 1) 'Rows - 1
    n = UBound(U, 2) 'Columns - 1
    ReDim r(n)
    ReDim W(n)
    ReDim V(n, n)

    g = 0# 'Householder Reduction to Bidiagonal form
    x = 0#
    d = 0#
    For i = 0 To n
        l = i + 1
        r(i) = x * g
        g = 0#
        s = 0#
        x = 0#
        If i <= m Then
            For k = i To m
                x = x + Abs(U(k, i))
            Next k
            If x <> 0# Then
                For k = i To m
                    U(k, i) = U(k, i) / x
                    s = s + U(k, i) * U(k, i)
                Next k
                f = U(i, i)
                If f < 0# Then
                    g = s ^ 0.5
                Else
                    g = -s ^ 0.5
                End If
                h = f * g - s
                U(i, i) = f - g
                For j = l To n
                    s = 0#
                    For k = i To m
                        s = s + U(k, i) * U(k, j)
                    Next k
                    f = s / h
                    For k = i To m
                        U(k, j) = U(k, j) + f * U(k, i)
                    Next k
                Next j
                For k = i To m
                    U(k, i) = U(k, i) * x
                Next k
            End If
        End If
        W(i) = x * g
        g = 0#
        s = 0#
        x = 0#
        If i <= m And i <> n Then
            For k = l To n
                x = x + Abs(U(i, k))
            Next k
            If x <> 0# Then
                For k = l To n
                    U(i, k) = U(i, k) / x
                    s = s + U(i, k) * U(i, k)
                Next k
                f = U(i, l)
                If f < 0# Then
                    g = s ^ 0.5
                Else
                    g = -s ^ 0.5
                End If
                h = f * g - s
                U(i, l) = f - g
                For k = l To n
                    r(k) = U(i, k) / h
                Next k
                For j = l To m
                    s = 0#
                    For k = l To n
                        s = s + U(j, k) * U(i, k)
                    Next k
                    For k = l To n
                        U(j, k) = U(j, k) + s * r(k)
                    Next k
                Next j
                For k = l To n
                    U(i, k) = U(i, k) * x
                Next k
            End If
        End If
        y = Abs(W(i)) + Abs(r(i))
        If y > d Then d = y
    Next i

    For i = n To 0 Step -1 'Accumulation of right-hand transformations
        If i < n Then
            If g <> 0# Then
                For j = l To n
                    V(j, i) = (U(i, j) / U(i, l)) / g
                Next j
                For j = l To n
                    s = 0#
                    For k = l To n
                        s = s + U(i, k) * V(k, j)
                    Next k
                    For k = l To n
                        V(k, j) = V(k, j) + s * V(k, i)
                    Next k
                Next j
            End If
            For j = l To n
                V(i, j) = 0#
                V(j, i) = 0#
            Next j
        End If
        V(i, i) = 1#
        g = r(i)
        l = i
    Next i

    If n < m Then t = n Else t = m 'Accumulation of left-hand transformation
    For i = t To 0 Step -1
        l = i + 1
        g = W(i)
        For j = l To n
            U(i, j) = 0#
        Next j
        If g <> 0# Then
            g = 1# / g
            For j = l To n
                s = 0#
                For k = l To m
                    s = s + U(k, i) * U(k, j)
                Next k
                f = (s / U(i, i)) * g
                For k = i To m
                    U(k, j) = U(k, j) + f * U(k, i)
                Next k
            Next j
            For j = i To m
                U(j, i) = U(j, i) * g
            Next j
        Else
            For j = i To m
                U(j, i) = 0#
            Next j
        End If
        U(i, i) = U(i, i) + 1
    Next i

    For k = n To 0 Step -1 'Diagonalization of the bidirectional form
        For p = 1 To 30
            e = True
            For l = k To 0 Step -1 'Test for Splitting
                t = l - 1
                If Abs(r(l)) + d = d Then
                    e = False
                    Exit For
                End If
                If Abs(W(t)) + d = d Then Exit For
            Next l
            If e Then
                c = 0# 'Cancelation of r(l) if l>1
                s = 1#
                For i = l To k
                    f = s * r(i)
                    r(i) = r(i) * c
                    If Abs(f) + d = d Then Exit For
                    g = W(i)
                    Select Case Abs(f) - Abs(g) 'h = (f^2 + g^2) ^ 0.5  '[without destructive overflow or underflow]
                    Case Is > 0#: h = Abs(f) * (1# + (g / f) * (g / f)) ^ 0.5
                    Case Is < 0#: h = Abs(g) * (1# + (f / g) * (f / g)) ^ 0.5
                    Case Else:    h = Abs(f) * 2# ^ 0.5
                    End Select
                    W(i) = h
                    h = 1# / h
                    c = g * h
                    s = -f * h
                    For j = 0 To m
                        y = U(j, t)
                        z = U(j, i)
                        U(j, t) = y * c + z * s
                        U(j, i) = z * c - y * s
                    Next j
                Next i
            End If
            z = W(k)
            If l = k Then 'Test for Convergence
                If z < 0# Then
                    W(k) = -z
                    For j = 0 To n
                        V(j, k) = -V(j, k)
                    Next j
                End If
                Exit For
            End If
            If p = 30 Then Err.Raise 1, , "No Convergence after 30 iterations"
            x = W(l) 'Shift from bottom 2-by-2 minor
            t = k - 1
            y = W(t)
            g = r(t)
            h = r(k)
            f = ((y - z) * (y + z) + (g - h) * (g + h)) / (2# * h * y)
            If f > 1# Then 'g=(f^2 + 1)^0.5  [without destructive overflow or underflow]
                g = Abs(f) * (1# + (1# / f) * (1# / f)) ^ 0.5
            Else
                g = (f * f + 1#) ^ 0.5
            End If
            If f < 0# Then
                f = ((x - z) * (x + z) + h * ((y / (f - Abs(g))) - h)) / x
            Else
                f = ((x - z) * (x + z) + h * ((y / (f + Abs(g))) - h)) / x
            End If
            c = 1# 'Next QR Transformation
            s = 1#
            For j = l To t
                i = j + 1
                g = r(i)
                y = W(i)
                h = s * g
                g = g * c
                Select Case Abs(f) - Abs(h) 'z = (f^2 + h^2) ^ 0.5  '[without destructive overflow or underflow]
                Case Is > 0#: z = Abs(f) * (1# + (h / f) * (h / f)) ^ 0.5
                Case Is < 0#: z = Abs(h) * (1# + (f / h) * (f / h)) ^ 0.5
                Case Else:    z = Abs(f) * 2# ^ 0.5
                End Select
                r(j) = z
                c = f / z
                s = h / z
                f = x * c + g * s
                g = g * c - s * x
                h = y * s
                y = y * c
                For q = 0 To n
                    x = V(q, j)
                    z = V(q, i)
                    V(q, j) = x * c + z * s
                    V(q, i) = z * c - s * x
                Next q
                Select Case Abs(f) - Abs(h) 'z = (f^2 + h^2) ^ 0.5  '[without destructive overflow or underflow]
                Case Is > 0#: z = Abs(f) * (1# + (h / f) * (h / f)) ^ 0.5
                Case Is < 0#: z = Abs(h) * (1# + (f / h) * (f / h)) ^ 0.5
                Case Else:    z = Abs(f) * 2# ^ 0.5
                End Select
                W(j) = z
                If z <> 0# Then 'Rotation can be arbitrary if z=0
                    z = 1# / z
                    c = f * z
                    s = h * z
                End If
                f = c * g + s * y
                x = c * y - s * g
                For q = 0 To m
                    y = U(q, j)
                    z = U(q, i)
                    U(q, j) = y * c + z * s
                    U(q, i) = z * c - y * s
                Next q
            Next j
            r(l) = 0#
            r(k) = f
            W(k) = x
        Next p
    Next k
    Erase r

End Sub


Public Sub SVDEDIT(W() As Double)
    Dim i       As Long     'Edit singular values by eliminating [set to zero]
    Dim n       As Long     '  the ones that are likely to cause
    Dim MaxW    As Double   '  destructive overflow/underflow errors
'    Dim MinW    As Double
'    Dim cn      As Double
    Dim Thresh  As Double

    n = UBound(W)
'    MinW = MaxDouble
    MaxW = -1#

    For i = 0 To n               'Determine maximum, minimum singular values
        If Abs(W(i)) > MaxW Then MaxW = Abs(W(i))
'        If W(i) < MinW Then MinW = W(i)
    Next i

    Thresh = MaxW / TOL
    For i = 0 To n
        If W(i) < Thresh Then W(i) = 0#
    Next i

'    If MinW = 0 Then             'Calculate Condition Number
'        cn = -1#    '  if condition number in infinate, set cn = -1.0
'    Else
'        If Log(MaxW) - Log(MinW) > Log(MaxDouble) Then 'Division will cause overflow
'            cn = -1#
'        Else
'            cn = MaxW / MinW
'        End If
'    End If

End Sub


Public Sub SVDBACKSUBSTITUTE(U() As Double, W() As Double, V() As Double, B() As Double, x() As Double)
    Dim i       As Long    'SVD Backsubstitution
    Dim j       As Long    'Solves  A * x = B  for vector x
    Dim m       As Long    'A is specified by return values of SVD():
    Dim n       As Long    '                  U(0 to m, 0 to n),
    Dim s       As Double  '                  W(0 to n),
    Dim UB()    As Double  '                  V(0 to n, 0 to n)

    m = UBound(U, 1)
    n = UBound(U, 2)
    ReDim UB(n)
    ReDim x(n)

    For j = 0 To n         ' Calculate U' * B
        s = 0#
        If W(j) <> 0# Then ' Nonzero result only if W(j) is nonzero
            For i = 0 To m
                s = s + U(i, j) * B(i)
            Next i
            s = s / W(j)
        End If
        UB(j) = s
    Next j
    For j = 0 To n         ' X = U' * B * V
        s = 0#
        For i = 0 To n
            s = s + V(j, i) * UB(i)
        Next i
        x(j) = s
    Next j

End Sub


Public Sub SVDSORT(U() As Double, W() As Double, V() As Double)
    Dim i       As Long    ' Sort U, V, W  by Singular Values in decending order
    Dim j       As Long    '   [meaning W(0) will be greatest singular value]
    Dim k       As Long
    Dim m       As Long
    Dim n       As Long
    Dim s       As Double

    m = UBound(U, 1)
    n = UBound(U, 2)

    For i = 0 To n - 1
        k = i                    'Find next highest singular value index k
        s = W(k)
        For j = i + 1 To n
            If W(j) > s Then
                k = j
                s = W(k)
            End If
        Next j
        If k <> i Then
            W(k) = W(i)          'Swap W(k), W(i)
            W(i) = s
            For j = 0 To n       'Swap V(Row i), V(Row k)
                s = V(j, i)
                V(j, i) = V(j, k)
                V(j, k) = s
            Next j
            For j = 0 To m       'Swap U(Row i), U(Row k)
                s = U(j, i)
                U(j, i) = U(j, k)
                U(j, k) = s
            Next j
        End If
    Next i

End Sub


Public Sub MatrixEchelon(U() As Double, B() As Double)
    Dim i       As Long
    Dim j       As Long 'Attempt to put in a form like this:
    Dim m       As Long
    Dim n       As Long '       1 | 0 | 0  || B(0)
    Dim x       As Long '       0 | 1 | 0  || B(1)
    Dim y       As Long '       0 | 0 | 1  || B(2)
    Dim c       As Long
    Dim r       As Long

    Dim g       As Double

    m = UBound(U, 1) ' and ubound(b)
    n = UBound(U, 2)

    For c = 0 To n
        For i = r To m
            If U(i, c) <> 0 Then
                If i <> r Then         'Swap Rows i,r, so U(r,c) is non-zero
                    For j = c To n
                        g = U(i, j)
                        U(i, j) = U(r, j)
                        U(r, j) = g
                    Next j
                    g = B(i)
                    B(i) = B(r)
                    B(r) = g
                End If

                g = 1 / U(r, c)       'Divide Row r by term U(r,c)
                U(r, c) = 1#          '  so U(r,c) = 1.0
                For y = c + 1 To n
                    U(r, y) = U(r, y) * g
                Next y
                B(r) = B(r) * g

                For x = 0 To m       'Multiply the rest of the rows by
                    If x <> r Then   '  row r * first term to make the first
                        g = U(x, c)  '  term zero
                        U(x, c) = 0#
                        For y = c + 1 To n
                            U(x, y) = U(x, y) - g * U(r, y)
                        Next y
                        B(x) = B(x) - g * B(r)
                    End If
                Next x
                r = r + 1
                Exit For
            End If
        Next i
    Next c

End Sub
