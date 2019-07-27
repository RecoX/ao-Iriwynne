Attribute VB_Name = "ListEvent"
Option Explicit
 
Private Type tList
 
    Quotas        As Byte      'Cantidad de cupos para la lista
    Active        As Boolean   'Lista activada?
    UsersInList   As Byte      'Usuarios en la lista
    UserIndex()   As Integer   'ID
 
End Type
 
Private List As tList
'_
 
Public Sub Start_List(ByVal UserIndex As Integer, ByVal Amount As Byte)
 
    '@@ Iniciamos la lista.
 
    With List
 
        If Amount < 1 Then Exit Sub
 
        If Not esGM(UserIndex) Then
            Call WriteConsoleMsg(UserIndex, "Tienes que ser GM para hacer esta operación.", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
 
        Call Clean_List
 
        .Quotas = Amount
        .Active = True
 
        '@@ Redimensionamos el array
        ReDim .UserIndex(1 To .Quotas) As Integer
 
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Lista activada para " & .Quotas & " usuarios, tipee /ENTRAR para ingresar al evento.", FontTypeNames.FONTTYPE_INFOBOLD))
 
    End With
 
End Sub
 
Public Sub Access_List(ByVal UserIndex As Integer)
 
    '@@ Meter usuario a la lista.
 
    With UserList(UserIndex)
 
        If .flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡¡Estás muerto!!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
 
       ' If EsGm(UserIndex) = True Or LCase(.Name) = "varawel" Then
       '     Call WriteConsoleMsg(UserIndex, "Tienes que ser una persona normal para ingresar a la lista.", FontTypeNames.FONTTYPE_INFOBOLD)
       '     Exit Sub
       ' End If

        If .flags.List <> 0 Then
            Call WriteConsoleMsg(UserIndex, "Ya estás en la lista!", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
 
        If List.Active = False Then
            Call WriteConsoleMsg(UserIndex, "No hay lista en curso.", FontTypeNames.FONTTYPE_INFOBOLD)
            Exit Sub
        End If
 
        '@@ Le damos el flag para que esté en lista
        .flags.List = 1
        '@@ Aumentamos la variable de usuarios en lista
        List.UsersInList = List.UsersInList + 1
        '@@ Guardamos el ID del usuario
        List.UserIndex(List.UsersInList) = UserIndex
 
        Call WriteNameList(UserIndex)
 
        Call WriteConsoleMsg(UserIndex, "Has ingresado a lista.", FontTypeNames.FONTTYPE_INFOBOLD)
 
        If List.UsersInList = List.Quotas Then Call Full_List
 
    End With
 
End Sub
 
Private Sub Full_List()
 
    '@@ Se completa la lista.
 
        List.Active = False
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("La lista ha sido completada.", FontTypeNames.FONTTYPE_INFOBOLD))
 
End Sub
 
Public Sub Clean_List()
 
    '@ Limpiamos la lista.
 
    Dim LoopC As Long
 
    With List
 
        '@@ Canceló la lista.
        'If .UsersInList < .Quotas Then _
        '    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Lista cancelada.", FontTypeNames.FONTTYPE_INFOBOLD))
 
        If .UsersInList <> 0 Then
            
            Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Lista cancelada.", FontTypeNames.FONTTYPE_INFOBOLD))
 
            For LoopC = 1 To .UsersInList
            
               If .UserIndex(LoopC) > 0 Then
                    UserList(.UserIndex(LoopC)).flags.List = 0
                End If
            Next LoopC
     
        End If
 
        .Active = False
        .Quotas = 0
        .UsersInList = 0
        
        
    End With
 
End Sub
 
Public Sub Remove_User(ByVal UserIndex As Integer)
 
    '@@ Por si se desconecta
 
    Dim LoopX As Long
    Dim LoopC As Long
 
    UserList(UserIndex).flags.List = 0
 
    For LoopX = 1 To List.UsersInList
 
        If List.UserIndex(LoopX) = UserList(UserIndex).ID Then
     
            For LoopC = LoopX To (List.UsersInList - 1)
     
                List.UserIndex(LoopC) = List.UserIndex(LoopC + 1)
         
            Next LoopC
     
            List.UsersInList = List.UsersInList - 1
     
            Exit For
     
        End If
 
    Next LoopX
 
End Sub
 

