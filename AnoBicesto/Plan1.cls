VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Ano()

    Dim AnoAt As Integer
    Dim MesAt As Integer
    Dim DiaAt As Integer
    Dim AnoNac As Integer
    Dim MesNac As Integer
    Dim DiaNac As Integer
    Dim Idade As Integer
    Dim Valido As Integer
    
    
    AnoNac = InputBox(Prompt:="Insira o ano do seu nascimento: ")
    MesNac = InputBox(Prompt:="Insira o mês do seu nascimento: ")
    
     Select Case (MesNac)
        Case 1
         DiaNac = InputBox(Prompt:="Insira o dia do seu nascimento: ")
        Case 2
            Valido = AnoNac Mod 4
            If (Valido = 0) Then
                MsgBox Prompt:="Você nasceu em um ano bicesto"
                DiaNac = I
        Case 3
            
        
        
    
    
    

End Sub
