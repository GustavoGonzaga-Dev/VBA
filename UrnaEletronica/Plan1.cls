VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub a()

    Dim Tom As Integer
    Dim Tomas As Integer
    Dim Jerry As Integer
    Dim Votos As Integer
    Dim Ctd As Integer
    Dim Nulo As Integer
    
    
    
    
    Tom = 0
    Tomas = 0
    Jerry = 0
    Ctd = 0
    Nulo = 0
    
    
    
    
    
        While Ctd <= 4
    
            MsgBox Prompt:=" Bem Vindo"
        
            Votos = InputBox(Prompt:=" Por favor, Digite '1' Para votar no candidato Tom, Digite '2' para votar no candidato Tomas ou digite '3' para votar no candidato Jerry: ")
        
            If (Votos = 1) Then
                Tom = Tom + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tom'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 2) Then
                Tomas = Tomas + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tomas'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos >= 4 Or Votos <= 0) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Wend
        
        MsgBox Prompt:="O candidato Tom teve ao total " & Tom & " Votos, o candidato Tomas teve ao total " & Tomas & " votos, o candidato Jerry teve ao total " & Jerry & " votos, e o numero de votos nulos foi de " & Nulo
        
        If (Tom > Tomas And Tomas > Jerry) Then
            MsgBox Prompt:=" os candidatos Tom e Tomas passaram para o segundo turno"
            Votos = 0
            Tom = 0
            Tomas = 0
        
        ElseIf (Tomas > Tom And Tom > Jerry) Then
            MsgBox Prompt:=" os candidatos Tomas e Tom passaram para o segundo turno"
            Votos = 0
            Tomas = 0
            Tom = 0
        
        ElseIf (Tom > Tomas And Tomas < Jerry) Then
            MsgBox Prompt:=" os candidatos Tom e Jerry passaram para o segundo turno"
            Votos = 0
            Tom = 0
            Jerry = 0
        
        ElseIf (Tomas > Tom And Tom < Jerry) Then
            MsgBox Prompt:=" os candidatos Tomas e Jerry passaram para o segundo turno"
            Votos = 0
            Tomas = 0
            Jerry = 0
            
        ElseIf (Jerry > Tom And Tom > Tomas) Then
            MsgBox Prompt:=" os candidatos Jerry e Tom passaram para o segundo turno"
            Votos = 0
            Jerry = 0
            Tom = 0
            
        ElseIf (Jerry > Tom And Tom < Tomas) Then
            MsgBox Prompt:=" os candidatos Jerry e Tomas passaram para o segundo turno"
            Votos = 0
            Jerry = 0
            Tomas = 0
            
        End If
    
       
End Sub



