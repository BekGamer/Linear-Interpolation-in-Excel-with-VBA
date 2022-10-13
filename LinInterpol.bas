Attribute VB_Name = "Module1"
Function LinInterpol(X, Xs As Range, Ys As Range)
    'For info https://github.com/BekGamer/Linear-Interpolation-in-Excel-with-VBA.git
    Dim Ks() As Double
    Dim Cs() As Double
    Dim Razn() As Double
    Dim RaznX() As Double
    Dim Proizv() As Double
    Dim ProizvX() As Double
    Dim a As Double
    Dim b As Double
    Dim n As Double
    Dim summpr As Double
    Dim summprX As Double
    Dim counter As Integer
    summpr = 0
    summprX = 0
    n = Application.CountA(Xs)
    If (n = 1) Then LinInterpol = Ys(1)
    If (n = 2) Then LinInterpol = Ys(1) + (X - Xs(1)) / (Xs(2) - Xs(1)) * (Ys(2) - Ys(1))
    If (n > 2) Then
        ReDim Ks(1 To n - 1) As Double
        ReDim Cs(1 To n - 2) As Double
        ReDim Razn(1 To n - 2) As Double
        ReDim Proizv(1 To n - 2) As Double
        ReDim RaznX(1 To n - 2) As Double
        ReDim ProizvX(1 To n - 2) As Double
        For counter = 2 To (n)
            Ks(counter - 1) = (Ys(counter) - Ys(counter - 1)) / (Xs(counter) - Xs(counter - 1))
            Next
        For counter = 1 To (n - 2)
            Cs(counter) = (-Ks(counter) + Ks(counter + 1)) / 2
            Razn(counter) = Abs(Xs(2) - Xs(counter + 1))
            RaznX(counter) = Abs(X - Xs(counter + 1))
            Proizv(counter) = Razn(counter) * Cs(counter)
            ProizvX(counter) = RaznX(counter) * Cs(counter)
            summpr = summpr + Proizv(counter)
            summprX = summprX + ProizvX(counter)
            Next
        a = (Ks(1) + Ks(n - 1)) / 2
        b = Ys(2) - (a * Xs(2) + summpr)
        LinInterpol = a * X + b + summprX
    End If
End Function
