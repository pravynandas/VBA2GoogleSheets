Sub HelloWorld()
'
' HelloWorld Macro
'

'
    ActiveCell.FormulaR1C1 = "Hello"
    ActiveSheet.Range("B1").Select
	ActiveCell.FormulaR1C1 = "World"
    		 Range("C1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(R[-4]C,R[5]C,RC[-3],RC[4])"
	    Range("D1").Select
    ActiveCell.FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-2],RC[-1])"
Range("A2").Select
End Sub
