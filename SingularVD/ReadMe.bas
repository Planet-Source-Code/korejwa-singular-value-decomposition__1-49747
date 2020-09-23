Attribute VB_Name = "ReadMe"
'
'This module contains no code.  VB5, VB4 users see note at bottom
'
'
'Singular Value Decomposition is a technique for solving systems of linear equations
'in the form:
'
'    a(0,0)*x(0) + a(0,1)*x(1) + a(0,2)*x(2)  ...  a(0,n)*x(n) = b(0)
'    a(1,0)*x(0) + a(1,1)*x(1) + a(1,2)*x(2)  ...  a(1,n)*x(n) = b(1)
'    a(2,0)*x(0) + a(2,1)*x(1) + a(2,2)*x(2)  ...  a(2,n)*x(n) = b(2)
'
'                                ....
'
'    a(m,0)*x(0) + a(m,1)*x(1) + a(m,2)*x(2)  ...  a(m,n)*x(n) = b(m)
'
'
'Where it is desired to compute values of x(0)~x(n) based on known values of
'a(0,0)~a(m,n) and b(0)~b(m).
'
'======================================================================================
'
'if n<m, then there are more equations than unknowns, so this is an overdetermined
'set of linear equations.  SVD will find the linear least squares solution to the set of
'equations.
'
'======================================================================================
'
'if n>m, then it is mathematically impossible to calculate an exact solution set.
'However, SVD can give you a "best solution" that describes all possible solutions:
'
'Given:
'     X -   Y +  3*Z = 4
' - 3*X + 4*Y - 10*Z = 10
'
'SVD Solution:
'   X =   16 - 2*t1
'   Y =   27 +   t1
'   Z =    5 +   t1
'
'This describes a 1 dimensional family of solutions, where t1 is an unknown value.
'Any value of t1 will make the given equations true.
'
'======================================================================================
'
'if n=m, then the number of equations is equal to the number of unknowns.  SVD will
'solve for the unknowns.
'
'======================================================================================
'
'Why use SVD?
'
'SVD solves sets of linear equations, and gives linear least-squares solutions, but
'what makes SVD special is that it diagnoses possible problems that other algorithms
'might not see.
'
'Accumulated round-off error can produce a completely inaccurate solution to the set
'of equations without warning.
'
'SVD decomposes a matrix into three pieces, the most important of which are the
'singular values [W()].  If a singular value is 0, the matrix is singular and can not
'be solved for.  If a singular value is relatively small, [compared to the rest of
'the singular values] it represents a row in the matrix whose values are so close to
'singular, it will cause significant roundoff error.  You can eliminate this roundoff
'error problem by setting any relatively small singular values to zero.
'
'Then you can backsubstitute the (possibly modified) SVD return values to get a
'trustworthy solution.
'
'If any singular value is zero, the coresponding V() column defines the unknown vector.
'(as with the case where n>m)
'
'======================================================================================
'
'Note for VB4, VB5 users:
'
'  VB6 allows matrix assign statements.
'  The following code makes B a copy of A:
'    Dim A() As Double
'    Dim B() As Double
'    ReDim A(9)
'      ...
'    B = A
'
'  Although there are no array assignments in the SVD module, there are some in the
'    Form file.
'
'  If you are using a version of VB older than VB6, you'll need to replace any
'    array assign statements and copy arrays manually, element by element.
'    (or use CopyMemory API, if you are savy enough)
'
'
'  MatrixAssign1D(B, A)
'
'
'Public Sub MatrixAssign2D(U() As Double, V() As Double)
'    Dim i As Long 'For pre-VB6 users
'    Dim j As Long '  Make U a copy of the 2D array V
'    Dim m As Long
'    Dim n As Long
'
'    m = UBound(V, 1)
'    n = UBound(V, 2)
'
'    ReDim U(LBound(V, 1) To m, LBound(V, 2) To n)
'    For i = LBound(V, 1) To m
'        For j = LBound(V, 2) To n
'            U(i, j) = V(i, j)
'        Next j
'    Next i
'
'End Sub
'
'Public Sub MatrixAssign1D(U() As Double, V() As Double)
'    Dim i As Long 'For pre-VB6 users
'    Dim n As Long '  Make U a copy of the 1D array V
'
'    n = UBound(V)
'
'    ReDim U(LBound(V) To n)
'    For i = LBound(V) To n
'        U(i) = V(i)
'    Next i
'
'End Sub
'
'
'-korejwa

