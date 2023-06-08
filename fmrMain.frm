VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5265
   ClientLeft      =   150
   ClientTop       =   195
   ClientWidth     =   5625
   Icon            =   "fmrMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "fmrMain.frx":169B2
   ScaleHeight     =   5265
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image 
      Height          =   5220
      Left            =   0
      Picture         =   "fmrMain.frx":17A7C
      ToolTipText     =   "Click na fórmula para entrada de dados"
      Top             =   0
      Width           =   5595
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim value0 As Integer
Dim value1, value2, result As Single

Private Sub Form_Load()
   value0 = 250 '250 x 250
   Me.Caption = App.Title & " - " & "Version " & App.Major & "." & App.Minor & "." & App.Revision
   
End Sub

Private Sub Image_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   'Coordenadas
   'Para criar uma área de ação, posicionei o cursor no centro do objetivo para colher os valores de X, Y.
   'Adicionei na função estes valores, criando assim uma quadrante com o cáculo de value0. Ex: value0=250, área=250x250
   'Me.Caption = "X: " & X & " Y: " & Y

   'CORRENTE
   '-------------------------------------------------------------------------------------------
   If X >= 2950 - value0 And X <= 2950 + value0 And Y >= 1050 - value0 And Y <= 1050 + value0 Then
      value1 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Corrente", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESISTÊNCIA (R)", "Cálculo de Corrente", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 / value2
      MsgBox Format(result, "0.000") & " A", , "Resultado de CORRENTE"
      Me.Caption = value1 & " / " & value2 & " = " & Format(result, "0.000") & " A"
      
   ElseIf X >= 3975 - value0 And X <= 3975 + value0 And Y >= 1450 - value0 And Y <= 1450 + value0 Then
      value1 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Corrente", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Corrente", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 / value2
      MsgBox Format(result, "0.000") & " A", , "Resultado de CORRENTE"
      Me.Caption = value1 & " / " & value2 & " = " & Format(result, "0.000") & " A"
      
   ElseIf X >= 4320 - value0 And X <= 4320 + value0 And Y >= 2325 - value0 And Y <= 2325 + value0 Then
      value1 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Corrente", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESISTÊNCIA (R)", "Cálculo de Corrente", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = Math.Sqr(value1 / value2)
      MsgBox Format(result, "0.000") & " A", , "Resultado de CORRENTE"
      Me.Caption = "Raiz( " & value1 & " / " & value2 & " )" & " = " & Format(result, "0.000") & " A"
   
   'RESISTÊNCIA
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 4320 - value0 And X <= 4320 + value0 And Y >= 3120 - value0 And Y <= 3120 + value0 Then
      value1 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Resistência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Resistência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      result = (value1 * value1) / value2
      MsgBox Format(result, "0.000") & " R", , "Resultado de RESISTÊNCIA"
      Me.Caption = "( " & value1 & " * " & value2 & " )" & " / " & value2 & " = " & Format(result, "0.000") & " R"
   
   ElseIf X >= 3825 - value0 And X <= 3825 + value0 And Y >= 3885 - value0 And Y <= 3885 + value0 Then
      value1 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Resistência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Resistência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 / value2
      MsgBox Format(result, "0.000") & " R", , "Resultado de RESISTÊNCIA"
      Me.Caption = value1 & " / " & value2 & " = " & Format(result, "0.000") & " R"
   
   ElseIf X >= 3090 - value0 And X <= 3090 + value0 And Y >= 4305 - value0 And Y <= 4305 + value0 Then
      value1 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Resistência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Resistência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 / (value2 * value2)
      MsgBox Format(result, "0.000") & " R", , "Resultado de RESISTÊNCIA"
      Me.Caption = value1 & " / " & "( " & value2 & " * " & value2 & " )" & " = " & Format(result, "0.000") & " R"
   
   'TENSÃO
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 2220 - value0 And X <= 2220 + value0 And Y >= 4290 - value0 And Y <= 4290 + value0 Then
      value1 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Tensão", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Tensão", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 / value2
      MsgBox Format(result, "0.000") & " V", , "Resultado de TENSÃO"
      Me.Caption = value1 & " / " & value2 & " = " & Format(result, "0.000") & " V"
   
   ElseIf X >= 1395 - value0 And X <= 1395 + value0 And Y >= 3945 - value0 And Y <= 3945 + value0 Then
      value1 = Replace(InputBox("Digite o valor de POTÊNCIA (W)", "Cálculo de Tensão", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESISTÊNCIA (R)", "Cálculo de Tensão", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = Math.Sqr(value1 * value2)
      MsgBox Format(result, "0.000") & " V", , "Resultado de TENSÃO"
      Me.Caption = "Raiz( " & value1 & " * " & value2 & " )" & " = " & Format(result, "0.000") & " V"
   
   ElseIf X >= 1000 - value0 And X <= 1000 + value0 And Y >= 3135 - value0 And Y <= 3135 + value0 Then
      value1 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Tensão", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESITÊNCIA (R)", "Cálculo de Tensão", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 * value2
      MsgBox Format(result, "0.000") & " V", , "Resultado de TENSÃO"
      Me.Caption = value1 & " * " & value2 & " = " & Format(result, "0.000") & " V"
   
   'POTÊNCIA
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 1020 - value0 And X <= 1020 + value0 And Y >= 2340 - value0 And Y <= 2340 + value0 Then
      value1 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Potência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESISTÊNCIA (R)", "Cálculo de Potência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = (value1 * value1) * value2
      MsgBox Format(result, "0.000") & " W", , "Resultado de POTÊNCIA"
      Me.Caption = "( " & value1 & " * " & value1 & " )" & " * " & value2 & " = " & Format(result, "0.000") & " W"
   
   ElseIf X >= 1335 - value0 And X <= 1335 + value0 And Y >= 1395 - value0 And Y <= 1395 + value0 Then
      value1 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Potência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de RESISTÊNCIA (R)", "Cálculo de Potência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = (value1 * value1) / value2
      MsgBox Format(result, "0.000") & " W", , "Resultado de POTÊNCIA"
      Me.Caption = "( " & value1 & " * " & value1 & " )" & " / " & value2 & " = " & Format(result, "0.000") & " W"
   
   ElseIf X >= 2200 - value0 And X <= 2200 + value0 And Y >= 1020 - value0 And Y <= 1020 + value0 Then
      value1 = Replace(InputBox("Digite o valor de TENSÃO (V)", "Cálculo de Potência", ""), ".", ",")
      If value1 = Empty Then Exit Sub
      If Not IsNumeric(value1) Then Exit Sub
      value2 = Replace(InputBox("Digite o valor de CORRENTE (A)", "Cálculo de Potência", ""), ".", ",")
      If value2 = Empty Then Exit Sub
      If Not IsNumeric(value2) Then Exit Sub
      result = value1 * value2
      MsgBox Format(result, "0.000") & " W", , "Resultado de POTÊNCIA"
      Me.Caption = value1 & " * " & value2 & " = " & Format(result, "0.000") & " W"
   
   End If
   
End Sub

Private Sub Image_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   'CORRENTE
   '-------------------------------------------------------------------------------------------
   If X >= 2950 - value0 And X <= 2950 + value0 And Y >= 1050 - value0 And Y <= 1050 + value0 Then
      Me.MousePointer = 99
      
   ElseIf X >= 3975 - value0 And X <= 3975 + value0 And Y >= 1450 - value0 And Y <= 1450 + value0 Then
      Me.MousePointer = 99
      
   ElseIf X >= 4320 - value0 And X <= 4320 + value0 And Y >= 2325 - value0 And Y <= 2325 + value0 Then
      Me.MousePointer = 99
   
   'RESISTÊNCIA
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 4320 - value0 And X <= 4320 + value0 And Y >= 3120 - value0 And Y <= 3120 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 3825 - value0 And X <= 3825 + value0 And Y >= 3885 - value0 And Y <= 3885 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 3090 - value0 And X <= 3090 + value0 And Y >= 4305 - value0 And Y <= 4305 + value0 Then
      Me.MousePointer = 99
   
   'TENSÃO
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 2220 - value0 And X <= 2220 + value0 And Y >= 4290 - value0 And Y <= 4290 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 1395 - value0 And X <= 1395 + value0 And Y >= 3945 - value0 And Y <= 3945 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 1000 - value0 And X <= 1000 + value0 And Y >= 3135 - value0 And Y <= 3135 + value0 Then
      Me.MousePointer = 99
   
   'POTÊNCIA
   '-------------------------------------------------------------------------------------------
   ElseIf X >= 1020 - value0 And X <= 1020 + value0 And Y >= 2340 - value0 And Y <= 2340 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 1335 - value0 And X <= 1335 + value0 And Y >= 1395 - value0 And Y <= 1395 + value0 Then
      Me.MousePointer = 99
   
   ElseIf X >= 2200 - value0 And X <= 2200 + value0 And Y >= 1020 - value0 And Y <= 1020 + value0 Then
      Me.MousePointer = 99
      
   Else
      Me.MousePointer = 0 'Default

   End If

End Sub





