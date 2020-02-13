PUBLIC oform1

SET PROCEDURE TO vfpxworkbookxlsx ADDITIVE

oform1=NEWOBJECT("form1")
oform1.Show()
RETURN


	**************************************************
*-- Form:         form1 (d:\work\xlsx-workbook-class\testform.scx)
*-- ParentClass:  form
*-- BaseClass:    form
*-- Time Stamp:   05/27/17 06:16:04 PM
*
DEFINE CLASS form1 AS form


	DataSession = 2
	Height = 250
	Width = 761
	DoCreate = .T.
	AutoCenter = .T.
	Caption = "Test"
	MaxButton = .F.
	MinButton = .F.
	Name = "Form1"


	ADD OBJECT clsvfpxworkbookxlsx AS vfpxworkbookxlsx WITH ;
		Height = 17, ;
		Left = 617, ;
		Top = 9, ;
		Width = 129, ;
		Name = "clsVFPxWorkbookXLSX"


	ADD OBJECT command1 AS commandbutton WITH ;
		Top = 6, ;
		Left = 7, ;
		Height = 30, ;
		Width = 183, ;
		Caption = "Export Northwind Customers", ;
		Name = "Command1"


	ADD OBJECT command2 AS commandbutton WITH ;
		Top = 37, ;
		Left = 7, ;
		Height = 30, ;
		Width = 183, ;
		Caption = "Class Demo", ;
		Name = "Command2"


	ADD OBJECT grid1 AS grid WITH ;
		DeleteMark = .F., ;
		Height = 168, ;
		Left = 7, ;
		ReadOnly = .T., ;
		Top = 75, ;
		Width = 749, ;
		Name = "Grid1"


	ADD OBJECT command3 AS commandbutton WITH ;
		Top = 6, ;
		Left = 193, ;
		Height = 30, ;
		Width = 183, ;
		Caption = "Export Grid to Excel", ;
		Name = "Command3"


	PROCEDURE Load
		SET SAFETY OFF
	ENDPROC


	PROCEDURE Resize
		thisform.grid1.Width  = thisform.Width - 14
		thisform.grid1.Height = thisform.Height - thisform.grid1.Top - 7
	ENDPROC


	PROCEDURE command1.Click
		LOCAL lcTable, lcExcel
		lcExcel = SYS(5) + ADDBS(SYS(2003)) + "Northwind Customers.xlsx"
		lcTable = ADDBS(SYS(2004)) + "Samples\Northwind\customers.dbf"
		lcTable = LOCFILE(lcTable)
		IF !ISNULL(lcTable) .AND. FILE(lcTable)
			thisform.clsVFPxWorkbookXLSX.SaveTabletoWorkbook(lcTable, lcExcel, .T., .T.)
			WAIT WINDOW "Saved To Excel" NOWAIT
		ENDIF
	ENDPROC


	PROCEDURE command2.Click
		thisform.clsVFPxWorkbookXLSX.Demo()
		WAIT WINDOW "Saved To Excel" NOWAIT
	ENDPROC


	PROCEDURE grid1.Init
		LOCAL lcTable, llFailed
		lcTable = ADDBS(SYS(2004)) + "Samples\Northwind\employees.dbf"
		lcTable = LOCFILE(lcTable)
		TRY
			USE (lcTable) IN 0 ALIAS employees SHARED
			llFailed = .F.
		CATCH TO loException
			llFailed = .T.
		ENDTRY
		IF llFailed
			RETURN
		ENDIF
		WITH this
			.ColumnCount  = 7
			.RecordSource = 'employees'
			WITH .Column1
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "employees.employeeid"
				.Header1.Caption   = "Employee Id"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column2
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "ALLTRIM(employees.lastname) + ', ' + ALLTRIM(employees.firstname)"
				.Header1.Caption   = "Employee Name"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column3
				.Resizable = .T.
				.Alignment = 2
				.ControlSource = "employees.hiredate"
				.Header1.Caption   = "Hire Date"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column4
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "employees.address"
				.Header1.Caption   = "Address"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column5
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "employees.city"
				.Header1.Caption   = "City"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
				.Text1.Alignment   = 1
			ENDWITH
			WITH .Column6
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "employees.region"
				.Header1.Caption   = "State"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
				.Text1.Alignment   = 1
			ENDWITH
			WITH .Column7
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "employees.postalcode"
				.Header1.Caption   = "Postal Code"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
		ENDWITH
	ENDPROC


	PROCEDURE command3.Click
		LOCAL lcExcel
		lcExcel = SYS(5) + ADDBS(SYS(2003)) + "Northwind Employees.xlsx"
		thisform.clsVFPxWorkbookXLSX.SaveGridToWorkbookEX(thisform.Grid1, lcExcel, .T., "Employees")
		WAIT WINDOW "Saved To Excel" NOWAIT
	ENDPROC


ENDDEFINE
*
*-- EndDefine: form1
**************************************************
