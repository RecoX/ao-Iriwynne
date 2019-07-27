Attribute VB_Name = "MercadoAo"
Option Explicit

Public Sub CargarListadoMAO(ByVal UserIndex As Integer)
    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Mercado.MAO", "DATOS", "POSTEOS"))

    For LoopC = 1 To N
        UserList(UserIndex).Name = GetVar(App.Path & "\Dat\Mercado.MAO", "LIST", "NOMBRE" & LoopC)
    Next LoopC
End Sub

Public Sub PosteoListado(ByVal UserIndex As Integer)
    Dim N As Integer, LoopC As Integer

    If UserList(UserIndex).flags.Posteado = 0 Then
        N = val(GetVar(App.Path & "\Dat\Mercado.MAO", "DATOS", "POSTEOS"))
        N = N + 1
        Call WriteVar(App.Path & "\Dat\Mercado.MAO", "DATOS", "POSTEOS", N)

        Call WriteVar(App.Path & "\Dat\Mercado.MAO", "LIST", "NOMBRE" & N, UserList(UserIndex).Name)

        Call WriteConsoleMsg(UserIndex, "Se ha posteado tu personaje con exito!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.Posteado = 1
    Else
        Call WriteConsoleMsg(UserIndex, "Este personaje ya se encuentra en el mercadoAO!", FontTypeNames.FONTTYPE_INFO)
    End If

End Sub

Public Sub DesposteoListado(ByVal UserIndex As Integer)
Dim N As Integer, LoopC As Integer, i As Integer

If UserList(UserIndex).flags.Posteado = 1 Then
N = val(GetVar(App.Path & "\Dat\Mercado.MAO", "DATOS", "POSTEOS"))
For LoopC = 1 To N
        If GetVar(App.Path & "\Dat\Mercado.MAO", "LIST", "NOMBRE" & LoopC) = UserList(UserIndex).Name Then
            Call WriteVar(App.Path & "\Dat\Mercado.MAO", "LIST", "NOMBRE" & N, "")
            N = N - 1
            Call WriteVar(App.Path & "\Dat\Mercado.MAO", "DATOS", "POSTEOS", N)
        End If
Next LoopC
Call WriteConsoleMsg(UserIndex, "Se ha Desposteado tu personaje con exito!", FontTypeNames.FONTTYPE_INFO)
UserList(UserIndex).flags.Posteado = 0
Else
Call WriteConsoleMsg(UserIndex, "Este personaje no se encuentra en el mercadoAO!", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

