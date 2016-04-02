Attribute VB_Name = "MdlVector"
Option Explicit

'Vector

Public Vec0 As Vector

Public Type Vector
    X As Double
    Y As Double
    Z As Double
End Type

Function VecAdd(Vec1 As Vector, Vec2 As Vector) As Vector
    With VecAdd
        .X = Vec1.X + Vec2.X
        .Y = Vec1.Y + Vec2.Y
        .Z = Vec1.Z + Vec2.Z
    End With
End Function

Function VecSub(Vec1 As Vector, Vec2 As Vector) As Vector
    With VecSub
        .X = Vec1.X - Vec2.X
        .Y = Vec1.Y - Vec2.Y
        .Z = Vec1.Z - Vec2.Z
    End With
End Function

Function VecNum(Vec1 As Vector, Num As Double) As Vector
    With VecNum
        .X = Vec1.X * Num
        .Y = Vec1.Y * Num
        .Z = Vec1.Z * Num
    End With
End Function

Function VecMul(Vec1 As Vector, Vec2 As Vector) As Vector
    With VecMul
        .X = Vec1.Y * Vec2.Z - Vec1.Z * Vec2.Y
        .Y = Vec1.Z * Vec2.X - Vec1.X * Vec2.Z
        .Z = Vec1.X * Vec2.Y - Vec1.Y * Vec2.X
    End With
End Function

Function VecDot(Vec1 As Vector, Vec2 As Vector) As Double
    VecDot = Vec1.X * Vec2.X + Vec1.Y * Vec2.Y + Vec1.Z * Vec2.Z
End Function

Function Vec(X As Double, Y As Double, Z As Double) As Vector
    With Vec
        .X = X
        .Y = Y
        .Z = Z
    End With
End Function

Function VecLen(Vec1 As Vector) As Double
    VecLen = Sqr(VecDot(Vec1, Vec1))
End Function

Function VecDis(Vec1 As Vector, Vec2 As Vector) As Double
    Dim VecDiff As Vector
    VecDiff = VecSub(Vec1, Vec2)
    VecDis = VecLen(VecDiff)
End Function

Function VecDir(Vec1 As Vector, Length As Double) As Vector
    Dim Len1 As Double
    Len1 = VecLen(Vec1)
    If Len1 <> 0 Then
        VecDir = VecNum(Vec1, Length / Len1)
    Else
        VecDir = Vec0
    End If
End Function

Function VecInRange(Vec1 As Vector, RangeStart As Vector, RangeEnd As Vector) As Boolean
    VecInRange = False
    If (Vec1.X >= RangeStart.X And Vec1.X <= RangeEnd.X) Or (Vec1.X <= RangeStart.X And Vec1.X >= RangeEnd.X) Then
        If (Vec1.Y >= RangeStart.Y And Vec1.Y <= RangeEnd.Y) Or (Vec1.Y <= RangeStart.Y And Vec1.Y >= RangeEnd.Y) Then
            If (Vec1.Z >= RangeStart.Z And Vec1.Z <= RangeEnd.Z) Or (Vec1.Z <= RangeStart.Z And Vec1.Z >= RangeEnd.Z) Then
                VecInRange = True
            End If
        End If
    End If
End Function
