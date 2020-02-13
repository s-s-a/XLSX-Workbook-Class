***************************************************
* Class Browser modifications by Arto Toikka 2005 *
* =============================================== *
*                                                 *
* Mode 4                                          *
* View Code generated: 02/13/20 01:11:31 PM       *
***************************************************

PUBLIC oform1_5P40S9WGX

SET CLASSLIB TO d:\work\xlsx-workbook-class\vfpxworkbookxlsx.vcx ADDITIVE

oform1_5P40S9WGX=NEWOBJECT("form1")

oform1_5P40S9WGX.Show
RETURN


**************************************************
*-- Form:            form1 (d:\work\xlsx-workbook-class\multigriddemo.scx)
*-- ParentClass:     form
*-- BaseClass:       form
*-- Time Stamp:      06/22/18 07:10:09 PM
*
DEFINE CLASS form1 AS form


	DataSession = 2
	Height = 379
	Width = 761
	DoCreate = .T.
	AutoCenter = .T.
	Caption = "Test"
	MaxButton = .F.
	MinButton = .F.
	*-- XML Metadata for customizable properties
	_memberdata = ""
	Name = "Form1"
	Visible = .T.


	ADD OBJECT clsvfpxworkbookxlsx AS vfpxworkbookxlsx WITH ;
		Height = 17, ;
		Left = 617, ;
		Top = 9, ;
		Width = 129, ;
		Name = "clsVFPxWorkbookXLSX", ;
		Visible = .T.


	ADD OBJECT grid1 AS grid_L1_C2_TC3


	ADD OBJECT command3 AS commandbutton_L1_C3_TC4


	ADD OBJECT grid2 AS grid_L1_C4_TC5


	PROCEDURE Load
		SET SAFETY OFF
	ENDPROC


	PROCEDURE Resize
		LOCAL lnHeight
		lnHeight = INT((thisform.Height - 44) / 2)
		thisform.grid1.Width  = thisform.Width - 14
		thisform.grid1.Height = lnHeight 
		thisform.grid2.Top    = lnHeight + 40
		thisform.grid2.Width  = thisform.Width - 14
		thisform.grid2.Height = lnHeight
	ENDPROC


ENDDEFINE
*
*-- EndDefine: form1
**************************************************


**************************************************
*-- Grid:            grid1 (d:\work\xlsx-workbook-class\multigriddemo.scx)
*-- ParentClass:     grid
*-- BaseClass:       grid
*-- Time Stamp:      06/22/18 06:47:10 PM
*
DEFINE CLASS grid_L1_C2_TC3 AS grid


	ColumnCount = 7
	DeleteMark = .F.
	Height = 168
	Left = 5
	ReadOnly = .T.
	Top = 34
	Width = 749
	Name = "Grid1"
	Column1.ReadOnly = .T.
	Column1.Name = "Column1"
	Column2.ReadOnly = .T.
	Column2.Name = "Column2"
	Column3.ReadOnly = .T.
	Column3.Name = "Column3"
	Column4.FontName = "Times New Roman"
	Column4.FontSize = 10
	Column4.ReadOnly = .T.
	Column4.Name = "Column4"
	Column5.ReadOnly = .T.
	Column5.Name = "Column5"
	Column6.ReadOnly = .T.
	Column6.Name = "Column6"
	Column7.ReadOnly = .T.
	Column7.Name = "Column7"
	Visible = .T.


	PROCEDURE Init && 1 
		


		Try
			This.column1.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column1.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column1.ADDOBJECT("header1","header")
			With This.column1.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column1.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column1.ADDOBJECT("text1","textbox")
			With This.column1.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column2.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column2.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column2.ADDOBJECT("header1","header")
			With This.column2.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column2.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column2.ADDOBJECT("text1","textbox")
			With This.column2.text1
				.FontName = "Arial"
				.FontSize = 9
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column3.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column3.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column3.ADDOBJECT("header1","header")
			With This.column3.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column3.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column3.ADDOBJECT("text1","textbox")
			With This.column3.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column4.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column4.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column4.ADDOBJECT("header1","header")
			With This.column4.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column4.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column4.ADDOBJECT("text1","textbox")
			With This.column4.text1
				.FontName = "Times New Roman"
				.FontSize = 10
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column5.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column5.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column5.ADDOBJECT("header1","header")
			With This.column5.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column5.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column5.ADDOBJECT("text1","textbox")
			With This.column5.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column6.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column6.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column6.ADDOBJECT("header1","header")
			With This.column6.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column6.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column6.ADDOBJECT("text1","textbox")
			With This.column6.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column7.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column7.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column7.ADDOBJECT("header1","header")
			With This.column7.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column7.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column7.ADDOBJECT("text1","textbox")
			With This.column7.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		LOCAL lcTable, llFailed
		lcTable = ADDBS(SYS(2004)) + "Samples\Northwind\customers.dbf"
		lcTable = LOCFILE(lcTable)
		TRY
			USE (lcTable) IN 0 ALIAS customers SHARED
			llFailed = .F.
		CATCH TO loException
			llFailed = .T.
		ENDTRY
		IF llFailed
			RETURN
		ENDIF
		WITH this
			.ColumnCount  = 7
			.RecordSource = 'customers'
			WITH .Column1
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "customers.customerid"
				.Header1.Caption   = "Customer Id"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column2
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "ALLTRIM(customers.companyname)"
				.Header1.Caption   = "Company Name"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column3
				.Resizable = .T.
				.Alignment = 2
				.ControlSource = "customers.contactname"
				.Header1.Caption   = "Contact Name"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column4
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "customers.address"
				.Header1.Caption   = "Address"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
			WITH .Column5
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "customers.city"
				.Header1.Caption   = "City"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
				.Text1.Alignment   = 1
			ENDWITH
			WITH .Column6
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "customers.region"
				.Header1.Caption   = "State"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
				.Text1.Alignment   = 1
			ENDWITH
			WITH .Column7
				.Resizable = .T.
				.Alignment = 0
				.ControlSource = "customers.postalcode"
				.Header1.Caption   = "Postal Code"
				.Header1.FontBold  = .T.
				.Header1.Alignment = 0
			ENDWITH
		ENDWITH
	ENDPROC


ENDDEFINE
*
*-- EndDefine: grid1
**************************************************


**************************************************
*-- Commandbutton:   command3 (d:\work\xlsx-workbook-class\multigriddemo.scx)
*-- ParentClass:     commandbutton
*-- BaseClass:       commandbutton
*-- Time Stamp:      06/22/18 07:10:09 PM
*
DEFINE CLASS commandbutton_L1_C3_TC4 AS commandbutton


	Top = 1
	Left = 5
	Height = 30
	Width = 183
	Caption = "Export Grid to Excel"
	Name = "Command3"
	Visible = .T.


	PROCEDURE Click
		LOCAL lcExcel, loGrids
		lcExcel = SYS(5) + ADDBS(SYS(2003)) + "NorthWindMultiDemo.xlsx"
		loGrids = CREATEOBJECT("Empty")
		ADDPROPERTY(loGrids, "Count", 2)
		ADDPROPERTY(loGrids, "List[2, 4]", "")
		loGrids.List[1, 1] = thisform.Grid1           && Grid object reference
		loGrids.List[1, 2] = "Customers"              && Sheet name
		loGrids.List[1, 3] = .T.                      && Freeze top row indicator
		loGrids.List[1, 4] = .F.                      && Include hidden columns indicator

		loGrids.List[2, 1] = thisform.Grid2
		loGrids.List[2, 2] = "Employees"
		loGrids.List[2, 3] = .T.
		loGrids.List[2, 4] = .F.

		thisform.clsVFPxWorkbookXLSX.SaveMultiGridToWorkbookEx(loGrids, lcExcel)
		WAIT WINDOW "Saved To Excel" NOWAIT
	ENDPROC


ENDDEFINE
*
*-- EndDefine: command3
**************************************************


**************************************************
*-- Grid:            grid2 (d:\work\xlsx-workbook-class\multigriddemo.scx)
*-- ParentClass:     grid
*-- BaseClass:       grid
*-- Time Stamp:      06/22/18 06:47:10 PM
*
DEFINE CLASS grid_L1_C4_TC5 AS grid


	ColumnCount = 7
	DeleteMark = .F.
	Height = 168
	Left = 5
	ReadOnly = .T.
	Top = 208
	Width = 749
	Name = "Grid2"
	Column1.ReadOnly = .T.
	Column1.Name = "Column1"
	Column2.ReadOnly = .T.
	Column2.Name = "Column2"
	Column3.ReadOnly = .T.
	Column3.Name = "Column3"
	Column4.FontName = "Times New Roman"
	Column4.FontSize = 10
	Column4.ReadOnly = .T.
	Column4.Name = "Column4"
	Column5.ReadOnly = .T.
	Column5.Name = "Column5"
	Column6.ReadOnly = .T.
	Column6.Name = "Column6"
	Column7.ReadOnly = .T.
	Column7.Name = "Column7"
	Visible = .T.


	PROCEDURE Init && 1 
		

		Try
			This.column1.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column1.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column1.ADDOBJECT("header1","header")
			With This.column1.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column1.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column1.ADDOBJECT("text1","textbox")
			With This.column1.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column2.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column2.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column2.ADDOBJECT("header1","header")
			With This.column2.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column2.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column2.ADDOBJECT("text1","textbox")
			With This.column2.text1
				.FontName = "Arial"
				.FontSize = 9
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column3.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column3.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column3.ADDOBJECT("header1","header")
			With This.column3.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column3.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column3.ADDOBJECT("text1","textbox")
			With This.column3.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column4.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column4.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column4.ADDOBJECT("header1","header")
			With This.column4.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column4.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column4.ADDOBJECT("text1","textbox")
			With This.column4.text1
				.FontName = "Times New Roman"
				.FontSize = 10
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column5.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column5.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column5.ADDOBJECT("header1","header")
			With This.column5.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column5.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column5.ADDOBJECT("text1","textbox")
			With This.column5.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column6.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column6.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column6.ADDOBJECT("header1","header")
			With This.column6.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column6.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column6.ADDOBJECT("text1","textbox")
			With This.column6.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


		Try
			This.column7.REMOVEOBJECT("text1")
		Catch
		Endtry

		Try
			This.column7.REMOVEOBJECT("header1")
		Catch
		Endtry
		Try
			This.column7.ADDOBJECT("header1","header")
			With This.column7.header1
				.Caption = "Header1"
				.Name = "Header1"
			Endwith
		Catch
		Endtry


		Try
			This.column7.REMOVEOBJECT("text1")
		Catch
		Endtry
		Try
			This.column7.ADDOBJECT("text1","textbox")
			With This.column7.text1
				.BorderStyle = 0
				.Margin = 0
				.ForeColor = RGB(0,0,0)
				.BackColor = RGB(255,255,255)
				.Name = "Text1"
				.Visible = .T.
			Endwith
		Catch
		Endtry


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


ENDDEFINE
*
*-- EndDefine: grid2
**************************************************

