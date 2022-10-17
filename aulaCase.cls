VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Plan1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Bonus()

    Dim Cargo As Integer
    Dim Salario As Currency
    Dim Bonus As Currency
    
    Bonus = 0
    
    Cargo = InputBox(Prompt:="Insira '1' para pão, " & vbCrLf & "'2' para batata, " & vbCrLf & "'3' para macarrão, " & vbCrLf & "'4' miojo, " & vbCrLf & "'5' para salada, " & vbCrLf & "'6' feijão, " & vbCrLf & "'7' para pizza, " & vbCrLf & "'8' farofa: ")
    Salario = InputBox(Prompt:="Insira o valor do seu salario: ")
    
    Select Case Cargo
        Case 1
            Bonus = (Salario * 0.15)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 2
            Bonus = (Salario * 0.1)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 3
            Bonus = (Salario * 0.05)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 4
            Bonus = (Salario * 0.02)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 5
            Bonus = (Salario * 0.02)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 6
            Bonus = (Salario * 0.01)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 7
            Bonus = (Salario * 0.01)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case 8
            Bonus = (Salario * 0.01)
            MsgBox Prompt:="Cargo: " & Cargo & ", Bonus do salario: R$" & Bonus
        Case Else
            MsgBox Prompt:=" Valor invalido"
    End Select
        
End Sub
