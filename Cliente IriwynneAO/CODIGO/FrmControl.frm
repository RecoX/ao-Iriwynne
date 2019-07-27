VERSION 5.00
Begin VB.Form FrmControl 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   5295
   ClientLeft      =   -60
   ClientTop       =   -30
   ClientWidth     =   4125
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   4800
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   2400
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Captions y Intervalos De :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "FrmControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Esta función Api devuelve un valor  Boolean indicando si la ventana es una ventana visible
Private Declare Function IsWindowVisible _
   Lib "user32" (ByVal hWnd As Long) As Long

'Esta función retorna el número de caracteres del caption de la ventana
Private Declare Function GetWindowTextLength _
                Lib "user32" _
                Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long

'Esta devuelve el texto. Se le pasa el hwnd de la ventana, un buffer donde se
'almacenará el texto devuelto, y el Lenght de la cadena en el último parámetro
'que obtuvimos con el Api GetWindowTextLength
Private Declare Function GetWindowText _
                Lib "user32" _
                Alias "GetWindowTextA" (ByVal hWnd As Long, _
                                        ByVal lpString As String, _
                                        ByVal cch As Long) As Long

'Esta es la función Api que busca las ventanas y retorna su handle o Hwnd
Private Declare Function GetWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wFlag As Long) As Long

'Constantes para buscar las ventanas mediante el Api GetWindow
Private Const GW_HWNDFIRST = 0&

Private Const GW_HWNDNEXT = 2&

Public Function ListarCaptions(ByRef Cant As Integer) As String
 
        Dim Handle   As Long, Titulo As String, lenT As Long, Ret As Long
        
        Dim SepAnssi As String
        
        SepAnssi = "#" ' @@ Lo usamos es mas optimo setearlo aca, que ir haciendolo en el bucle
        
        'Obtenemos el Hwnd de la primera ventana, usando la constante GW_HWNDFIRST
        Handle = GetWindow(hWnd, GW_HWNDFIRST)

        'Este bucle va a recorrer todas las ventanas.
        'cuando GetWindow devielva un 0, es por que no hay mas
        Do While Handle <> 0

                'Tenemos que comprobar que la ventana es una de tipo visible
                If IsWindowVisible(Handle) Then
                
                        'Obtenemos el número de caracteres de la ventana
                        lenT = GetWindowTextLength(Handle)

                        'si es el número anterior es mayor a 0
                        If lenT > 0 Then
                        
                                'Creamos un buffer. Este buffer tendrá el tamaño con la variable LenT
                                Titulo = String$(lenT, 0)
                                
                                'Ahora recuperamos el texto de la ventana en el buffer que le enviamos
                                'y también debemos pasarle el Hwnd de dicha ventana
                                Ret = GetWindowText(Handle, Titulo, lenT + 1)
                                
                                Titulo$ = Left$(Titulo, Ret)
                                
                                'La agregamos al ListBox
                                ListarCaptions = Titulo & SepAnssi & ListarCaptions
                                
                                Cant = Cant + 1

                        End If

                End If

                'Buscamos con GetWindow la próxima ventana usando la constante GW_HWNDNEXT
                Handle = GetWindow(Handle, GW_HWNDNEXT)
        Loop

End Function

Private Sub Command1_Click()
        Unload Me

End Sub

