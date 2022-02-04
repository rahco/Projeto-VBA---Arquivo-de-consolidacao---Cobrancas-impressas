Attribute VB_Name = "Módulo1"
Sub Consolidar()

    'Tipo Var
    Dim valor As String
    valor = MsgBox("Consolidar todos os dados atualizados?", vbOKCancel, "VALIDAÇÃO DE ATIVAÇÃO DE MACROS")
    If valor = 1 Then

    Application.ScreenUpdating = False

    Sheets("CARREGAR").Select
    Range("E2").Select
    Selection.Copy
    Range("H2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("B4").Select

    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B2").Select
    

' Carregamento 53.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\53.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C4").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[53.1.xlsx]53.1'!C4)"
    Range("C4").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Windows("53.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("53.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close
    

' Carregamento 53.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\53.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C5").Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[53.2.xlsx]53.2'!C4)"
    Range("C5").Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Range("C6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("B4").Select
     
    Windows("53.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("53.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 53.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\53.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[53.3.xlsx]53.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("53.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("53.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 54.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\54.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[54.1.xlsx]54.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("54.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("54.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 54.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\54.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[54.2.xlsx]54.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("54.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("54.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 54.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\54.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[54.3.xlsx]54.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("54.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("54.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 55.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\55.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[55.1.xlsx]55.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("55.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("55.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 55.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\55.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[55.2.xlsx]55.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("55.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("55.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 55.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\55.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[55.3.xlsx]55.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("55.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("55.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 67.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\67.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[67.1.xlsx]67.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("67.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("67.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 67.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\67.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[67.2.xlsx]67.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("67.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("67.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 67.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\67.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[67.3.xlsx]67.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("67.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("67.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 81.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\81.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[81.1.xlsx]81.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("81.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("81.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 81.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\81.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[81.2.xlsx]81.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("81.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("81.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 81.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\81.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[81.3.xlsx]81.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("81.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("81.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 82.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\82.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[82.1.xlsx]82.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("82.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("82.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 82.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\82.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[82.2.xlsx]82.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("82.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("82.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 82.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\82.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[82.3.xlsx]82.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("82.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("82.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' ---------------------------------------------
' ---------------------------------------------
' ---------------------------------------------


' Carregamento 87.1
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\87.1.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[87.1.xlsx]87.1'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("87.1.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("87.1.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 87.2
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\87.2.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[87.2.xlsx]87.2'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("87.2.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("87.2.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


' Carregamento 87.3
    
    Workbooks.Open Filename:= _
    ActiveWorkbook.Path & "\87.3.xlsx"
    
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("CARREGAR").Select
    Range("C3").Select
    Selection.End(xlDown).Offset(1, 0).Select
    ActiveCell.FormulaR1C1 = "=COUNTA('[87.3.xlsx]87.3'!C4)"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
     
    Windows("87.3.xlsx").Activate
    Range("G1").Select
    Selection.End(xlDown).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Range("A1").Select
    Windows("Arquivo de consolidação - Cobranças impressas.xlsm").Activate
    Sheets("DADOS CONSOLIDADOS").Select
    Range("B2").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Windows("87.3.xlsx").Activate
    ActiveWorkbook.Save
    ActiveWindow.Close


    Else
    End If
    
    Sheets("CARREGAR").Select
    Range("B4").Select
    
    Application.ScreenUpdating = True
    
End Sub


