# Atualização: Bônus Fidelidade

A cada 12 meses, o aluno recebe o bônus de 1% de acréscimo nos descontos.

## Sub Desconto(ByVal rowTbDCad As Integer)

```diff
Sub Desconto(ByVal rowTbDCad As Integer)
' Argumento: Linha da matrícula na tabela Dados Cadastro


    Dim planInsPag As Worksheet
    Set planInsPag = Sheets("INSERIR PAGAMENTO")
    Dim tbDCad As ListObject
    Set tbDCad = Sheets("DADOS CADASTRO").ListObjects("TabelaDadosCadastro")
    

    ' Desconto Fixo:
    If tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Desconto Fixo").Index).Value <> "" Then
        planInsPag.Range("InsPag_18") = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("Desconto Fixo").Index).Value
        planInsPag.Range("InsPag_19") = "(fixo)"
    End If


-    If dataHoje <= (5 & "/" & planInsPag.Range("InsPag_MesRef") & "/" & planInsPag.Range("InsPag_AnoRef")) Then
-        
-            planInsPag.Range("InsPag_18") = "5%"
-            
-    End If

+    ' Descontos somente até o dia 05 do mês de referência, caso seja o atual ou futuro
+    Dim dataHoje As Date
+    dataHoje = Date
+
+    If Not (dataHoje <= (5 & "/" & planInsPag.Range("InsPag_MesRef") & "/" & planInsPag.Range("InsPag_AnoRef"))) Then Exit Sub
+     
+     
+    ' Há desconto
+    planInsPag.Range("InsPag_18").FormulaLocal = "=Cel_Desc+Cel_Bonus"
+    planInsPag.Range("Cel_Desc") = 5 / 100
+
+    
+    ' Bônus fidelidade: A cada 12 meses acumulado, ganha 1% de desconto além dos 5%
+    
+    Dim dataMatr As Date
+    dataMatr = tbDCad.DataBodyRange.Cells(rowTbDCad, tbDCad.ListColumns("DataMatricula").Index).Value
+    
+    Dim diasMatr As Integer
+    
+    diasMatr = DateDiff("d", dataMatr, dataHoje)
+    
+    If diasMatr >= 365 Then
+    ' Há bônus
+    
+        Dim difAnos As Variant
+        
+        difAnos = diasMatr / 365
+        
+        difAnos = WorksheetFunction.RoundDown(diasMatr / 365, 0)
+        
+        planInsPag.Range("Cel_Bonus") = difAnos / 100
+        
+        planInsPag.Range("Cel_Bonus_Text") = "Bônus"
+        
+        With planInsPag.Range("Cel_Bonus").Borders
+            .LineStyle = xlContinuous
+            .Color = -5791991
+            .Weight = xlThin
+        End With
+        
+    End If
    
End Sub
```

## Sub LimparInsPag()

```diff
    ' Apaga "Parc."
    planInsPag.Range("Cel_InsPag_CobParc").Offset(-2, 0) = ""
    
    
+    ' Apaga bônus
+    planInsPag.Range("Cel_Bonus_Text").Value = ""
+    
+    planInsPag.Range("Cel_Bonus").Borders.LineStyle = xlNone
+    planInsPag.Range("Cel_Bonus").Value = ""
```