'-----------------------------------------------------------
'|  SCRIPT PARA PREENCHER EXCEL - TRELLO EXPORT    			|
'|															|
'|  DESENVOLVEDOR: JOSE ARTEIRO TEIXEIRA (JOSE.CAVALCANTI)	|
'------------------------------------------------------------

Sub Criar_colunas(Col_Name, Col_Orig, Col_Number)
	Set objRange = objExcel.Range(Col_Orig).EntireColumn
	objRange.Insert(xlShiftToRight)
	objExcel.Cells(1, Col_Number).Value = Col_Name
End Sub

Function Ajust_Cols(Range_cell_border, Num_Cols)
	objExcel.Range(Range_cell_border).Select
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range(Range_cell_border).Borders.colorindex = 1
	
	For Aux_loop = 1 to Num_Cols
		objExcel.Cells(1, Aux_loop).Interior.Color = RGB(141, 180, 226)
	Next
End Function

Function Ajust_Cols_Red(Range_cell_border)
	objExcel.Range(Range_cell_border).Select
	objExcel.Cells.EntireColumn.AutoFit
	objExcel.Range(Range_cell_border).Borders.colorindex = 1
	
	For Aux_loop = 16 to 19
		objExcel.Cells(1, Aux_loop).Interior.Color = RGB(248, 203, 173)
	Next
End Function

Function Contar_Linhas(Col_range, Sheet_Name)
	Cont_rows = 0
	Count_White = 0
	objExcel.Cells(1, 1).Select
	For Each Cell In objWorkbook.Worksheets(Sheet_Name).Range(Col_range).Cells
		'If Cell.Value = "" Then Exit For
		If Cell.Value = "" Then 
			Count_White = Count_White + 1
		Else
			Cont_rows = Cont_rows + 1
			Count_White = 0
		End If
		If Count_White = 10 Then Exit For
	Next
	Contar_Linhas = Cont_rows
End Function

Function Create_Sheet(Sheet_Name)
	objExcel.ActiveWorkbook.Sheets.Add 
	objExcel.ActiveSheet.name = Sheet_Name
End Function

Function Copy_Sheet(Orign_Sheet, Dest_Sheet)
	objExcel.Worksheets(Orign_Sheet).Copy , objExcel.Worksheets(Orign_Sheet)
	objExcel.ActiveSheet.name = Dest_Sheet
End Function

Function WhatEver(num)
    If(Len(num)=1) Then
        WhatEver="0"&num
    Else
        WhatEver=num
    End If
End Function

Function myDateFormat(myDate)
    d = WhatEver(Day(myDate)-1)
    m = WhatEver(Month(myDate))    
    y = Year(myDate)
	If d < 1 Then
		d = 30
		m = m - 1
	End If
	If m < 1 Then
		m = 12
		y = y - 1
	End If
    myDateFormat= d & "/" & m & "/" & y
End Function

Sub Delete_Column(P_Val, P_Column)
	If objExcel.Cells(1, P_Column).Value = P_Val Then
		objExcel.Range(P_Column & ":" & P_Column).Delete
	Else
		Msgbox "Não encontrado Coluna '" & P_Val &"' em " & P_Column
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		WScript.Quit
	End If
End Sub

Sub Move_Column(P_Val, P_Column, T_Column)
	If objExcel.Cells(1, P_Column).Value = P_Val Then
		
		objExcel.Range(T_Column & ":" & T_Column).Value = objExcel.Range(P_Column & ":" & P_Column).Value
		objExcel.Range(P_Column & ":" & P_Column).Delete
	Else
		Msgbox "Não encontrado Coluna '" & P_Val &"' em " & P_Column
		objExcel.ActiveWorkbook.Close
		objExcel.Application.Quit
		WScript.Quit
	End If
End Sub

Function WorksheetExists(wsName, objWorkbook)
    
    ret = False
    wsName = UCase(wsName)
    For Each ws In objWorkbook.Sheets
        If UCase(ws.Name) = wsName Then
            ret = True
            Exit For
        End If
    Next
    WorksheetExists = ret
End Function

Sub Remover_Sheet(STR_VAR)
	'Stopping Application Alerts
	objExcel.DisplayAlerts=FALSE
	
	objExcel.Worksheets(STR_VAR).delete 
	'objWorkbook.sheets(STR_VAR).delete
	
	'Enabling Application alerts once we are done with our task
	objExcel.DisplayAlerts=TRUE
End Sub

Function Str_Month(Aux_Cell)
	If Aux_Cell = "1" Then
		Str_Month = "Janeiro"
	ElseIf Aux_Cell = "2" Then
		Str_Month = "Fevereiro"
	ElseIf Aux_Cell = "3" Then
		Str_Month = "Março"
	ElseIf Aux_Cell = "4" Then
		Str_Month = "Abril"
	ElseIf Aux_Cell = "5" Then
		Str_Month = "Maio"
	ElseIf Aux_Cell = "6" Then
		Str_Month = "Junho"
	ElseIf Aux_Cell = "7" Then
		Str_Month = "Julho"
	ElseIf Aux_Cell = "8" Then
		Str_Month = "Agosto"
	ElseIf Aux_Cell = "9" Then
		Str_Month = "Setembro"
	ElseIf Aux_Cell = "10" Then
		Str_Month = "Outubro"
	ElseIf Aux_Cell = "11" Then
		Str_Month = "Novembro"
	ElseIf Aux_Cell = "12" Then
		Str_Month = "Dezembro"
	Else
		Str_Month = "Unknow"
	End If
End Function

Function Release_Month(Aux_Cell)
	If InStr(Aux_Cell, "- 01 -") > 0 Then
		Release_Month = "1"
	ElseIf InStr(Aux_Cell, "- 02 -") > 0 Then
		Release_Month = "2"
	ElseIf InStr(Aux_Cell, "- 03 -") > 0 Then
		Release_Month = "3"
	ElseIf InStr(Aux_Cell, "- 04 -") > 0 Then
		Release_Month = "4"
	ElseIf InStr(Aux_Cell, "- 05 -") > 0 Then
		Release_Month = "5"
	ElseIf InStr(Aux_Cell, "- 06 -") > 0 Then
		Release_Month = "6"
	ElseIf InStr(Aux_Cell, "- 07 -") > 0 Then
		Release_Month = "7"
	ElseIf InStr(Aux_Cell, "- 08 -") > 0 Then
		Release_Month = "8"
	ElseIf InStr(Aux_Cell, "- 09 -") > 0 Then
		Release_Month = "9"
	ElseIf InStr(Aux_Cell, "- 10 -") > 0 Then
		Release_Month = "10"
	ElseIf InStr(Aux_Cell, "- 11 -") > 0 Then
		Release_Month = "11"
	ElseIf InStr(Aux_Cell, "- 12 -") > 0 Then
		Release_Month = "12"
	Else
		Release_Month = "0"
	End If
End Function

Function Func_Column_Excel(Aux)
	Dim Var_Column_Excel(26)
	Var_Column_Excel(0) = ""
	Var_Column_Excel(1) = "A"
	Var_Column_Excel(2) = "B"
	Var_Column_Excel(3) = "C"
	Var_Column_Excel(4) = "D"
	Var_Column_Excel(5) = "E"
	Var_Column_Excel(6) = "F"
	Var_Column_Excel(7) = "G"
	Var_Column_Excel(8) = "H"
	Var_Column_Excel(9) = "I"
	Var_Column_Excel(10) = "J"
	Var_Column_Excel(11) = "K"
	Var_Column_Excel(12) = "L"
	Var_Column_Excel(13) = "M"
	Var_Column_Excel(14) = "N"
	Var_Column_Excel(15) = "O"
	Var_Column_Excel(16) = "P"
	Var_Column_Excel(17) = "Q"
	Var_Column_Excel(18) = "R"
	Var_Column_Excel(19) = "S"
	Var_Column_Excel(20) = "T"
	Var_Column_Excel(21) = "U"
	Var_Column_Excel(22) = "V"
	Var_Column_Excel(23) = "W"
	Var_Column_Excel(24) = "X"
	Var_Column_Excel(25) = "Y"
	Var_Column_Excel(26) = "Z"
	Func_Column_Excel = Var_Column_Excel(Aux)
End Function

Function Remover_Etiqueta(Func_Var_Stat)
	Func_Aux_Stat = Split(Func_Var_Stat, "(")
	If IsArray(Func_Aux_Stat) And UBound(Func_Aux_Stat) >= 0 Then
		Remover_Etiqueta = Func_Aux_Stat(0)
	Else
		Remover_Etiqueta = Func_Var_Stat
	End If
End Function

Function Validar_TrelloExport()

	Validar_TrelloExport = True

	Dim Column_Name(34)
	Column_Name(1) = "Organization"
	Column_Name(2) = "Board"
	Column_Name(3) = "List"
	Column_Name(4) = "Card #"
	Column_Name(5) = "Title"
	Column_Name(6) = "Link"
	Column_Name(7) = "Description"
	Column_Name(8) = "Total Checklist items"
	Column_Name(9) = "Completed Checklist items"
	Column_Name(10) = "Checklists"
	Column_Name(11) = "NumberOfComments"
	Column_Name(12) = "Comments"
	Column_Name(13) = "Attachments"
	Column_Name(14) = "Votes"
	Column_Name(15) = "Spent"
	Column_Name(16) = "Estimate"
	Column_Name(17) = "Points Estimate"
	Column_Name(18) = "Points Consumed"
	Column_Name(19) = "Created"
	Column_Name(20) = "CreatedBy"
	Column_Name(21) = "LastActivity"
	Column_Name(22) = "Due"
	Column_Name(23) = "Done"
	Column_Name(24) = "DoneBy"
	Column_Name(25) = "DoneTime"
	Column_Name(26) = "Members"
	Column_Name(27) = "Labels"
	Column_Name(28) = "Período ET"
	Column_Name(29) = "Período CTU"
	Column_Name(30) = "Migracao TS"
	Column_Name(31) = "Andamento CTU (%)"
	Column_Name(32) = "Andamento ET (%)"
	Column_Name(33) = "Líder da Demanda"
	Column_Name(34) = "Migracao TI"
	
	For Column = 1 to uBound(Column_Name)
		If objExcel.Cells(1, Column).Value <> Column_Name(Column) Then
			Msgbox "Erro: " & VbCrLf & "Campo " & objExcel.Cells(1, Column).Value & " não esperado nessa sequencia. " & VbCrLf & "Esperado: " & Column_Name(Column)
			Validar_TrelloExport = False
		End If
	Next
	
End Function


'------ VOID MAIN ------
VAR_TRELLO_EMANAGER = "-= Trello Export Manager V3.2=-"

MsgBox(VAR_TRELLO_EMANAGER)

Dim Var_Stat_Name()
Dim Var_Stat_Risc()
Dim Var_Status()
Dim Var_Count_Month(13)
Dim DEBUG

DEBUG = MsgBox("Debug?", "36", VAR_TRELLO_EMANAGER)

For Each Count_Month in Var_Count_Month
	Count_Month = 0
Next

Set FSO = WScript.CreateObject("Scripting.FileSystemObject")
	Set objExcel = CreateObject("Excel.Application")
		vFileName = objExcel.GetOpenFilename ("ExcelFiles (*.xl*), *.xl*")
		
		If vFileName = "False" Or vFileName = "FALSE" Or Len(vFileName) < 1 Then
			MsgBox "Caminho de Entrada Incorreto. Aplicativo estará sendo fechado."
			objExcel.Application.Quit
			Set objExcel = Nothing
			Set FSO = Nothing
			WScript.Quit
		End If

		Set objWorkbook = objExcel.Workbooks.Open(vFileName)
			
			If DEBUG <> 7 Then
				objExcel.Application.Visible = True
			Else
				objExcel.Application.Visible = False
			End If
			
			If Validar_TrelloExport = False Then
				MsgBox "Campos Incorretos. Aplicativo estará sendo fechado."
				objExcel.Application.Quit
				Set objWorkbook = Nothing
				Set objExcel = Nothing
				Set FSO = Nothing
				WScript.Quit
			End If
			
			If WorksheetExists("TrelloExport", objWorkbook) = "True" Then

				Remover_Sheet("Archive")
			
				Name_Sheet   = "TrelloExport"
			
				Cont_rows = Contar_Linhas("A:A", Name_Sheet)
				
				objExcel.Worksheets(Name_Sheet).select

				Call Delete_Column("Organization", "A")				'Período ET' em AA
				Call Delete_Column("Período ET", "AA")				'Período ET' em AA
				Call Delete_Column("Members", "Y")					'Members' em Y
				Call Delete_Column("DoneTime", "X")					'DoneTime' em X
				Call Delete_Column("DoneBy", "W")					'DoneBy' em W
				Call Delete_Column("Done", "V")						'Done' em V
				Call Delete_Column("Due", "U")						'Due' em U
				Call Delete_Column("LastActivity", "T")				'LastActivity' em T
				Call Delete_Column("CreatedBy", "S")				'CreatedBy' em S
				Call Delete_Column("Created", "R")					'Created' em R
				Call Delete_Column("Points Consumed", "Q")			'Points Consumed' em Q
				Call Delete_Column("Points Estimate", "P")			'Points Estimate' em P
				Call Delete_Column("Estimate", "O")					'Estimate' em O
				Call Delete_Column("Spent", "N")					'Spent' em N
				Call Delete_Column("Votes", "M")					'Votes' em M
				Call Delete_Column("Attachments", "L")				'Attachments' em L
				Call Delete_Column("Comments", "K")					'Comments' em K
				Call Delete_Column("NumberOfComments", "J")			'NumberOfComments' em J
				Call Delete_Column("Checklists", "I")				'Checklists' em I
				Call Delete_Column("Completed Checklist items", "H")'Completed Checklist items' em H
				Call Delete_Column("Total Checklist items", "G")	'Total Checklist items' em G
				Call Delete_Column("Description", "F")				'Description' em F
				Call Delete_Column("Card #", "C") 					'Card #' em C
				
				Call Criar_colunas("Title", "B:B", "B")
				Call Move_Column("Title", "D", "B")
				
				Call Criar_colunas("Link", "J:J", "J")
				Call Move_Column("Link", "D", "J")
				
				Call Criar_colunas("Andamento CTU (%)", "E:E", "E")
				Call Move_Column("Andamento CTU (%)", "H", "E")
				
				Call Criar_colunas("Andamento ET (%)", "F:F", "F")
				Call Move_Column("Andamento ET (%)", "I", "F")
				
				Call Criar_colunas("Migracao TI", "I:I", "I")
				Call Move_Column("Migracao TI", "L", "I")
				
				Call Criar_colunas("Líder da Demanda", "J:J", "J")
				Call Move_Column("Líder da Demanda", "L", "J")
				
				objExcel.Cells(1, "A").Value = "Release"				'Board -> Release
				objExcel.Cells(1, "B").Value = "PRJ"					'Title -> PRJ
				objExcel.Cells(1, "C").Value = "Fase atual do Projeto"	'List -> Fase atual do Projeto
				objExcel.Cells(1, "D").Value = "Status"					'Labels -> Status
				objExcel.Cells(1, "K").Value = "URL Trello"				'Link -> URL Trello
				
				objExcel.Worksheets(Name_Sheet).Range("A1:K1").Interior.Color = RGB(191, 191, 191)
				objExcel.Worksheets(Name_Sheet).Range("A1:K1").Font.Color = RGB(0, 0, 0)
				
				For AuxCount = 2 to Cont_rows
				
					Rel_Month = Release_Month(objExcel.Cells(AuxCount, "A").Value)
					
					Var_Count_Month(Rel_Month) = Var_Count_Month(Rel_Month) + 1
					
					Aux_Stat = Split(objExcel.Cells(AuxCount, "D").Value, ",")
					
					For Each Arr_Stat in Aux_Stat
						Arr_Stat = Remover_Etiqueta(Arr_Stat)
						Arr_Stat = Trim(Arr_Stat)
						Bool_Stat = False
						Count_For_each = 0
						For Each Var_Stat in Var_Stat_Name
							If Var_Stat = Arr_Stat & "|" & Rel_Month Then
								Var_Stat_Risc(Count_For_each) = Var_Stat_Risc(Count_For_each) + 1
								Bool_Stat = True
							End If
							Count_For_each = Count_For_each + 1
						Next
						
						If Bool_Stat = False Then
							'MsgBox "->" & Count_For_each & ", " & Arr_Stat & ", " & Rel_Month
							ReDim Preserve Var_Stat_Name(Count_For_each)
							Var_Stat_Name(Count_For_each) = Arr_Stat & "|" & Rel_Month
							
							ReDim Preserve Var_Stat_Risc(Count_For_each)
							Var_Stat_Risc(Count_For_each) = 1
						End If
					Next
					
					'If objExcel.Cells(AuxCount, "H").Value <> "" Then
					'	objExcel.Cells(AuxCount, "H").Value = Replace(objExcel.Cells(AuxCount, "H").Value, "T", " ")
					'	objExcel.Cells(AuxCount, "H").Value = Mid(objExcel.Cells(AuxCount, "H").Value, 9, 2) & "/" & Mid(objExcel.Cells(AuxCount, "H").Value, 6, 2)& "/" & Mid(objExcel.Cells(AuxCount, "H").Value, 1, 4) & Mid(objExcel.Cells(AuxCount, "H").Value, 11)
					'End If
					
					'If objExcel.Cells(AuxCount, "I").Value <> "" Then
					'	objExcel.Cells(AuxCount, "I").Value = Replace(objExcel.Cells(AuxCount, "I").Value, "T", " ")
					'	objExcel.Cells(AuxCount, "I").Value = Mid(objExcel.Cells(AuxCount, "I").Value, 9, 2) & "/" & Mid(objExcel.Cells(AuxCount, "I").Value, 6, 2)& "/" & Mid(objExcel.Cells(AuxCount, "I").Value, 1, 4) & Mid(objExcel.Cells(AuxCount, "I").Value, 11)
					'End If
				Next
				
				For Each Var_Stat in Var_Stat_Name
					Aux_Stat = Split(Var_Stat, "|")
					Bool_Stat = False
					Count_For_each = 0
					For Each Pri_Stat in Var_Status
						If Pri_Stat = Aux_Stat(0) Then Bool_Stat = True
						Count_For_each = Count_For_each + 1
					Next
					If Bool_Stat = False Then
						ReDim Preserve Var_Status(Count_For_each)
						Var_Status(Count_For_each) = Aux_Stat(0)
					End If
				Next
				
				objExcel.Cells(Cont_rows + 3, "B").Value = "TOTAL demandas"
				
				Count_For_each = 2
				For Each Var_Stat in Var_Status
					Count_For_each = Count_For_each + 1
					objExcel.Cells(Cont_rows + 3, Count_For_each).Value = Var_Stat
				Next
				
				For Aux = 1 to 12
					objExcel.Cells(Cont_rows + 3 + Aux, "A").Value = Str_Month(Aux)
					objExcel.Cells(Cont_rows + 3 + Aux, "B").Value = Var_Count_Month(Aux)
					Count_For_each = 2
					For Each Var_Stat in Var_Status
						Count_For_each2 = 0
						Count_For_each = Count_For_each + 1
						For Each Var_Stat2 in Var_Stat_Name
							Aux_Stat = Split(Var_Stat2, "|")
							'MsgBox "->" & Aux_Stat(0) & " (" & Aux_Stat(1) & "), " & Var_Stat & "(" & Aux & ")"
							If Aux_Stat(0) = Var_Stat And "0" & Aux_Stat(1) = "0" & Aux Then
								objExcel.Cells(Cont_rows + 3 + Aux, Count_For_each).Value = Var_Stat_Risc(Count_For_each2)
							End If
							Count_For_each2 = Count_For_each2 + 1
						Next
					Next
				Next

				objExcel.Range("A1:K" & Cont_rows).Select
				objExcel.Cells.EntireColumn.AutoFit
				objExcel.Range("A1:K1").HorizontalAlignment = -4108
				objExcel.Range("A1:K" & Cont_rows).Borders.colorindex = 1
				
				objExcel.Range("A1:K" & Cont_rows).NumberFormat = "General"
				objExcel.Range("E1:F" & Cont_rows).NumberFormat = "0.0%"
				'objExcel.Range("M2:M" & Cont_rows).NumberFormat = "DD/MM/YYYY"
				'objExcel.Range("O2:O" & Cont_rows).NumberFormat = "DD/MM/YYYY"
				'objExcel.Range("Q2:Q" & Cont_rows).NumberFormat = "DD/MM/YYYY"
				'objExcel.Range("R2:R" & Cont_rows).NumberFormat = "DD/MM/YYYY"
				
				objExcel.Worksheets(Name_Sheet).Range("A" & Cont_rows + 3 & ":" & Func_Column_Excel(Count_For_each) & Cont_rows + 3).Interior.Color = RGB(191, 191, 191)
				objExcel.Worksheets(Name_Sheet).Range("A" & Cont_rows + 3 & ":" & Func_Column_Excel(Count_For_each) & Cont_rows + 3).Font.Color = RGB(0, 0, 0)

				objExcel.Range("A" & Cont_rows + 3 & ":" & Func_Column_Excel(Count_For_each) & Cont_rows + 3 + Aux).Select
				objExcel.Cells.EntireColumn.AutoFit
				objExcel.Range("A" & Cont_rows + 3 & ":" & Func_Column_Excel(Count_For_each) & Cont_rows + 3 + Aux).HorizontalAlignment = -4108
				objExcel.Range("A" & Cont_rows + 3 & ":" & Func_Column_Excel(Count_For_each) & Cont_rows + 3 + Aux).Borders.colorindex = 1
			
			Else
				Msgbox "Planilha Desconhecida."
			End If
			
			objExcel.ActiveWorkbook.Save
			objExcel.ActiveWorkbook.Close
			objExcel.Application.Quit
			'WScript.Quit

		Set objWorkbook = Nothing
		
		MsgBox("Finished with " & Cont_rows & " Lines")

		VAR_STR_MENU2 = MsgBox("Abrir Planilha?", "36", VAR_TRELLO_EMANAGER)

		If VAR_STR_MENU2 <> 7 Then
			'Set objShell = CreateObject("WScript.Shell")
			'	Set objExecObject = objShell.Exec("excel.exe " & vFileName)
			'		'varTMP = objExecObject.StdOut.ReadAll
			'		'DEBUG(varTMP)
			'	Set objExecObject = Nothing
			'Set objShell = Nothing
			Set objWorkbook = objExcel.Workbooks.Open(vFileName)
				objExcel.Application.Visible = True
			Set objWorkbook = Nothing
		End If
	Set objExcel = Nothing
Set FSO = Nothing

WScript.Quit