VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub CaixaEletronico()

    Dim ValorSaque, Saldo As Double
    Dim Senha, Ctd As Integer
    
    Ctd = 0
    
    Saldo = 12000
    
    Senha = InputBox(Prompt:=" Insira sua senha: ")
    
    While Ctd < 5 'iniciando contador'
        
        If (Senha = 123) Then 'validação de senha'
        
        MsgBox Prompt:=" Seu saldo Atual é de: R$" & Saldo
        ValorSaque = InputBox(Prompt:=" Insira o valor que deseja retirar: ")
        
        
        
        Else
        
        MsgBox Prompt:=" Senha incorreta, tente outra vez."
        Ctd = Ctd + 1
        End If
        
     Wend
        
            If (ValorSaque > 50) Then 'validação do valor do saque'
                MsgBox Prompt:=" Valor de saque não  permitido."
                
            Application.Quit
                
            Else
            MsgBox Prompt:=" Processando saque..." 'iniciando saque'
                
            Saldo = Saldo - ValorSaque
        
            End If
                
                                'verificando notas'
                
                If (ValorSaque = 50) Then
                
                    MsgBox Prompt:=" Nota de R$50"
        
                    MsgBox Prompt:=" Seu saldo Atual é de: R$ " & Saldo
            
                        ElseIf (ValorSaque = 20) Then
                            MsgBox Prompt:=" Nota de R$20"
                    
                            MsgBox Prompt:=" Seu saldo Atual é de: R$" & Saldo
                    
                                ElseIf (ValorSaque = 10) Then
                                    MsgBox Prompt:=" Nota de R$10"
                    
                                    MsgBox Prompt:=" Seu saldo Atual é de: R$" & Saldo
                                    
                                End If
                        
                

End Sub
