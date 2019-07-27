Attribute VB_Name = "Matematicas"
Option Explicit

Public Function Porcentaje(ByVal Total As Long, ByVal Porc As Long) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        Porcentaje = (Total * Porc) / 100

End Function

Function Distancia(ByRef wp1 As WorldPos, ByRef wp2 As WorldPos) As Long
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Encuentra la distancia entre dos WorldPos
        Distancia = Abs(wp1.X - wp2.X) + Abs(wp1.Y - wp2.Y) + (Abs(wp1.Map - wp2.Map) * 100)

End Function

Function Distance(ByVal X1 As Integer, _
                  ByVal Y1 As Integer, _
                  ByVal X2 As Integer, _
                  ByVal Y2 As Integer) As Double
        '***************************************************
        'Author: Unknown
        'Last Modification: -
        '
        '***************************************************

        'Encuentra la distancia entre dos puntos

        Distance = Sqr(((Y1 - Y2) ^ 2 + (X1 - X2) ^ 2))

End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
        '**************************************************************
        'Author: Juan Martín Sotuyo Dodero
        'Last Modify Date: 3/06/2006
        'Generates a random number in the range given - recoded to use longs and work properly with ranges
        '**************************************************************
        RandomNumber = Fix(Rnd * (UpperBound - LowerBound + 1)) + LowerBound

End Function
