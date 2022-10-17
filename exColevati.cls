VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub colevati()

    Dim N As Integer
    Dim M As Integer
    Dim Ctd As Integer
    Dim Cto As Integer
    Dim P As Integer
    Dim H As Integer
    
    Ctd = 0
    Do
    
        N = InputBox(Prompt:="Insira um numero natural: ")
        M = InputBox(Prompt:="Insira um numero natural que seja maior que o numero insirido anteriormente: ")
        
        If (M > N) Then
            
            Cto = N
            
            While Cto < M
                Cto = Cto + 1
                P = 2
                While P < Cto
                    H = Cto Mod P
                    
                    If (H = 0) Then
                        MsgBox Prompt:="" & Cto & " não é primo."
                        P = M
                    End If
                    P = P + 1
                Wend
                     
            Wend
        Else
            MsgBox Prompt:="Insira os numeros novamente."
            Ctd = Ctd + 1
        End If
        
    Loop While M > N Or Ctd <= 1
    
        MsgBox Prompt:="Finalizando operação"
    

End Sub
