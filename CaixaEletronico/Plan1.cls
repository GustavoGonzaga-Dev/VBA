VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub CommandButton1_Click()
    Dim ValorSaque As Double
    Dim ValorDeposito As Double
    Dim Saldo As Double
    Dim Saque As Double
    
    
    Dim NumConta As Integer
    Dim NumCadas As Integer
    Dim Ctd As Integer
    
    Dim valido, Resposta As Integer
    
    
    Ctd = 0
    
    Saldo = Cells(1, 1) 'linha/coluna'
    NumCadas = Cells(1, 2)
    
    
    NumConta = InputBox(Prompt:=" Insira o numero da conta: ")
    
    While Ctd < 5 'iniciando contador'
        
        If (NumConta = NumCadas) Then 'validação de senha'
        
            MsgBox Prompt:=" Seu saldo Atual é de: R$" & Saldo
            
                                'tipo de transição
            
            Resposta = InputBox(Prompt:=" Digite '1' para sacar ou '2' para depositar: ")
            
            If (Resposta > 2 Or Resposta <= 0) Then
            
                MsgBox Prompt:=" Ação invalida, tente outra vez."
            
            
            ElseIf (Resposta = 1) Then
            
 
                MsgBox Prompt:=" Este caixa não trabalha com moedas."
                ValorSaque = InputBox(Prompt:=" Insira o valor que deseja retirar: R$")
                Ctd = 5
                valido = 1
            
            ElseIf (Resposta = 2) Then
            
                MsgBox Prompt:=" Este caixa não trabalha com moedas."
                ValorDeposito = InputBox(Prompt:=" Insira o valor que deseja Depositar: R$")
            
                Ctd = 5
                valido = 2
            End If
            
            
        
        Else
        
        MsgBox Prompt:=" Senha incorreta, tente outra vez."
        Ctd = Ctd + 1
        NumConta = InputBox(Prompt:=" Insira o numero da conta: ")
        End If
        
        
        Wend
        
        
        
        
        
        
        
                                'operação 1'
        
        If (valido = 1) Then
            
            
                                'verificando valor do saque'
            If (ValorSaque > 300 Or ValorSaque <= 0 Or ValorSaque > Saldo) Then
                MsgBox Prompt:=" Valor de saque não  permitido."
      
                
            Else
            MsgBox Prompt:=" Processando saque..." 'iniciando saque'
                
            Saldo = Saldo - ValorSaque
            
            Cells(1, 1) = Saldo
        
            
             'verificando notas'
              
                
                Dim Contador50 As Integer
                Dim Contador20 As Integer
                Dim Contador10 As Integer
                Dim Contador5 As Integer
                
                
                
                Saque = 0
                Contador50 = 0
                Contador20 = 0
                Contador10 = 0
                Contador5 = 0
                
                
            While Saque + 1 <= ValorSaque
            
                If (Saque + 50 <= ValorSaque) Then
                    Saque = Saque + 50
                    Contador50 = Contador + 1
                    MsgBox Prompt:=" Nota de R$50"
                    
                ElseIf (Saque + 20 <= ValorSaque) Then
                    Saque = Saque + 20
                    Contador20 = Contador + 1
                    MsgBox Prompt:=" Nota de R$20"
                    
                ElseIf (Saque + 10 <= ValorSaque) Then
                    Saque = Saque + 10
                    Contador10 = Contador + 1
                    MsgBox Prompt:=" Nota de R$10"
                    
                ElseIf (Saque + 5 <= ValorSaque) Then
                    Saque = Saque + 5
                    Contador5 = Contador + 1
                    MsgBox Prompt:=" Nota de R$5"
                    
                ElseIf (Saque + 2 <= ValorSaque) Then
                    Saque = Saque + 2
                    Contador5 = Contador + 1
                    MsgBox Prompt:=" Nota de R$2"
                    
                ElseIf (Saque + 1 <= ValorSaque) Then
                    Saque = Saque + 1
                    Contador5 = Contador + 1
                    MsgBox Prompt:=" Nota de R$1"
                
                    
                End If
                
            
            Wend
            
            
            
            End If        'finalizando saque'
                
                               
                    
                        
                        
                         
        End If 'fim da operação 1'
        
        
        
        
                        'operação 2'
        
        If (valido = 2) Then
                        If (ValorDeposito <= 0) Then
                        MsgBox Prompt:=" Valor de deposito não  permitido."
                        
                        Else
                        MsgBox Prompt:=" Processando deposito..." 'iniciando saque'
                
                        Saldo = Saldo + ValorDeposito
                        MsgBox Prompt:=" Seu saldo atual é de: R$" & Saldo
                        
                        Cells(1, 1) = Saldo
                
                        End If
        End If                  'fim da operação2'
        
        
End Sub
