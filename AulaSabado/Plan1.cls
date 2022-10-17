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
    Dim Branco As Integer
    Dim Sair As Integer
    
    
    
    
    Tom = 0
    Tomas = 0
    Jerry = 0
    Ctd = 0
    Nulo = 0
    Branco = 0
    
    
    
    
    
    
        Do
        
        
         MsgBox Prompt:=" Bem Vindo"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "2 - Tomas" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
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
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Primeiro turno."
             Voto = Votos - Votos
             
        
            
        MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Tomas teve ao total " & Tomas & " votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
        
        If (Tom > Tomas And Tom > Jerry) Then
            MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            Votos = 0
            Tom = 0
            Tomas = 0
        
        ElseIf (Tomas > Tom And Tomas > Jerry) Then
            MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            Votos = 0
            Tomas = 0
            Tom = 0
        
        ElseIf (Jerry > Tomas And Jerry > Tom) Then
            MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            Votos = 0
            Tom = 0
            Jerry = 0
        
        ElseIf (Tomas > Tom And Tomas = Jerry) Then
            MsgBox Prompt:=" os candidatos Tomas e Jerry empataram com " & Tomas & " votos cada um."
            Votos = 0
            Tomas = 0
            Jerry = 0
            
            Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 2 - Tomas" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 2) Then
                Tomas = Tomas + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tomas'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 1) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
             MsgBox Prompt:="o candidato Tomas teve ao total " & Tomas & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Jerry > Tomas) Then
                MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Jerry) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Jerry And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Jerry And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
            
        ElseIf (Tom > Tomas And Tom = Jerry) Then
            MsgBox Prompt:=" os candidatos Tom e Jerry empataram com " & Tom & " votos cada um."
            Votos = 0
            Jerry = 0
            Tom = 0
            
            Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "3- Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 1) Then
                Tom = Tom + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tom'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 2) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
            
            MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Jerry) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Jerry > Tom) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
        ElseIf (Tomas > Jerry And Tomas = Tom) Then
            MsgBox Prompt:=" os candidatos Tomas e Tom empataram com " & Tomas & " votos cada um."
            Votos = 0
            Tom = 0
            Tomas = 0
            
             Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "2 - Tomas" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
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
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
            
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
            MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Tomas teve ao total " & Tomas & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Tomas) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Tom) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
            
            Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 2 - Tomas" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 2) Then
                Tomas = Tomas + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tomas'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 1) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
             MsgBox Prompt:="o candidato Tomas teve ao total " & Tomas & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Jerry > Tomas) Then
                MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Jerry) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Jerry And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Jerry And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
        
        ElseIf (Nulo > Tom And Nulo > Jerry And Nulo > Tomas) Then
            MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
        ElseIf (Branco > Tom And Branco > Jerry And Branco > Tomas) Then
            MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
        
        
        
        
                            'Segundo turno'
        
        
        '1'
        
        ElseIf (Tom > Tomas And Tomas > Jerry) Then
            MsgBox Prompt:=" os candidatos Tom e Tomas passaram para o segundo turno"
            Votos = 0
            Tom = 0
            Tomas = 0
            
            
         Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "2 - Tomas" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
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
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
            
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
            MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Tomas teve ao total " & Tomas & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Tomas) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Tom) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
        '2'
        
        ElseIf (Tomas > Tom And Tom > Jerry) Then
            MsgBox Prompt:=" os candidatos Tomas e Tom passaram para o segundo turno"
            Votos = 0
            Tomas = 0
            Tom = 0
            
            Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "2 - Tomas" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
        
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
                ElseIf (Votos = 3) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
             MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Tomas teve ao total " & Tomas & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Tomas) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Tom) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
        End If
            
            '3'
        
        ElseIf (Tom > Tomas And Tomas < Jerry) Then
            MsgBox Prompt:=" os candidatos Tom e Jerry passaram para o segundo turno"
            Votos = 0
            Tom = 0
            Jerry = 0
            
         Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "3- Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 1) Then
                Tom = Tom + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tom'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 2) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
            
            MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Jerry) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Jerry > Tom) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
            '4'
        
        ElseIf (Tomas > Tom And Tom < Jerry) Then
            MsgBox Prompt:=" os candidatos Tomas e Jerry passaram para o segundo turno"
            Votos = 0
            Tomas = 0
            Jerry = 0
            
        Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 2 - Tomas" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 2) Then
                Tomas = Tomas + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tomas'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 1) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
             MsgBox Prompt:="o candidato Tomas teve ao total " & Tomas & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Jerry > Tomas) Then
                MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Jerry) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Jerry And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Jerry And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
            
        '5'
            
        ElseIf (Jerry > Tom And Tom > Tomas) Then
            MsgBox Prompt:=" os candidatos Jerry e Tom passaram para o segundo turno"
            Votos = 0
            Jerry = 0
            Tom = 0
            
            Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 1 - Tom" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 1) Then
                Tom = Tom + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tom'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 2) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
        Loop While (Votos <> 6)
             MsgBox Prompt:="Finalizando Segundo turno."
             Voto = Votos - Votos
             
             MsgBox Prompt:="o candidato Tom teve ao total " & Tom & " Votos" & vbCrLf & "o candidato Jerry teve ao total " & Jerry & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Tom > Jerry) Then
                MsgBox Prompt:=" o Candidato Tom foi eleito com " & Tom & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Jerry > Tom) Then
                MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Tom And Nulo > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Tom And Branco > Jerry) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
            
        '6'
            
        ElseIf (Jerry > Tom And Tom < Tomas) Then
            MsgBox Prompt:=" os candidatos Jerry e Tomas passaram para o segundo turno"
            Votos = 0
            Jerry = 0
            Tomas = 0
        
        Do
        
        
         MsgBox Prompt:=" Bem Vindo ao segundo turno"
        
         Votos = InputBox(Prompt:=" 2 - Tomas" & vbCrLf & "3 - Jerry" & vbCrLf & "4 - Nulo" & vbCrLf & "5 - Branco" & vbCrLf & "6 - Sair")
         
        If (Votos = 2) Then
                Tomas = Tomas + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Tomas'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
        
            ElseIf (Votos = 3) Then
                Jerry = Jerry + 1
                MsgBox Prompt:=" Voto atribuido ao candidato 'Jerry'. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            
            ElseIf (Votos = 1) Then
                MsgBox Prompt:=" Numero de candidato não permitido. "
                Votos = Votos - Votos
                
            
            ElseIf (Votos = 4) Then
                Nulo = Nulo + 1
                MsgBox Prompt:=" Voto aceito como nulo. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
                
            ElseIf (Votos = 5) Then
                Branco = Branco + 1
                MsgBox Prompt:=" Voto aceito como em Branco. "
                Votos = Votos - Votos
                Ctd = Ctd + 1
            End If
            
        Loop While (Votos <> 6)
            MsgBox Prompt:="Finalizando Segundo turno."
            Voto = Votos - Votos
            
            MsgBox Prompt:="o candidato Jerry teve ao total " & Jerry & " Votos" & vbCrLf & "o candidato Tomas teve ao total " & Tomas & " votos" & vbCrLf & "o numero de votos nulos foi de " & Nulo & vbCrLf & "o numero de votos em Branco foi de " & Branco
             
            If (Jerry > Tomas) Then
                MsgBox Prompt:=" o Candidato Jerry foi eleito com " & Jerry & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Tomas > Jerry) Then
                MsgBox Prompt:=" o Candidato Tomas foi eleito com " & Tomas & "votos, " & Branco & " votos brancos e " & Nulo & " votos nulos"
            ElseIf (Nulo > Jerry And Nulo > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos nulo foi maior que os votos atribuidos ao candidatos, tendo " & Nulo & " votos nulos, de " & Ctd & " votos no total."
            
            ElseIf (Branco > Jerry And Branco > Tomas) Then
                MsgBox Prompt:="Nenhum candidato foi eleito pois o numero de votos em Branco foi maior que os votos atribuidos ao candidatos, tendo " & Branco & " votos Brancos, de " & Ctd & " votos no total."
            
        End If
        
        End If
    
       
End Sub




