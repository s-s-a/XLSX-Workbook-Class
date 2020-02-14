Public oform1

Set Procedure To vfpxworkbookxlsx Additive

oform1 = Newobject("form1")
oform1.Show()

Define Class form1 As Form

  DataSession = 2
  Height = 379
  Width = 761
  DoCreate = .T.
  AutoCenter = .T.
  Caption = "Test"
  MaxButton = .F.
  MinButton = .F.
*-- XML Metadata for customizable properties
  _MemberData = ""
  Name = "Form1"
  Visible = .T.

  Add Object clsvfpxworkbookxlsx As vfpxworkbookxlsx With ;
    Height = 17, ;
    Left = 617, ;
    Top = 9, ;
    Width = 129, ;
    Name = "clsVFPxWorkbookXLSX", ;
    Visible = .T.

  Add Object grid1 As grid_L1_C2_TC3

  Add Object command3 As commandbutton_L1_C3_TC4

  Add Object grid2 As grid_L1_C4_TC5

  Procedure Load
  Set Safety Off

  Procedure Resize
  Local lnHeight
  With This
    lnHeight = Int((.Height - 44) / 2)
    With .grid1
      .Width  = This.Width - 14
      .Height = lnHeight
    Endwith
    With .grid2
      .Top    = lnHeight + 40
      .Width  = This.Width - 14
      .Height = lnHeight
    Endwith
  Endwith

Enddefine
*-- EndDefine: form1
**************************************************

Define Class grid_L1_C2_TC3 As Grid

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

  Procedure Init && 1

  Try
    This.Column1.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column1.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column1.AddObject("header1","header")
    With This.Column1.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column1.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column1.AddObject("text1","textbox")
    With This.Column1.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column2.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column2.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column2.AddObject("header1","header")
    With This.Column2.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column2.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column2.AddObject("text1","textbox")
    With This.Column2.text1
      .FontName = "Arial"
      .FontSize = 9
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column3.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column3.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column3.AddObject("header1","header")
    With This.Column3.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column3.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column3.AddObject("text1","textbox")
    With This.Column3.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column4.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column4.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column4.AddObject("header1","header")
    With This.Column4.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column4.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column4.AddObject("text1","textbox")
    With This.Column4.text1
      .FontName = "Times New Roman"
      .FontSize = 10
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column5.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column5.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column5.AddObject("header1","header")
    With This.Column5.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column5.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column5.AddObject("text1","textbox")
    With This.Column5.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column6.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column6.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column6.AddObject("header1","header")
    With This.Column6.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column6.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column6.AddObject("text1","textbox")
    With This.Column6.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column7.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column7.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column7.AddObject("header1","header")
    With This.Column7.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column7.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column7.AddObject("text1","textbox")
    With This.Column7.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Local lcTable, llFailed
  lcTable = Addbs(Sys(2004)) + "Samples\Northwind\customers.dbf"
  lcTable = Locfile(lcTable)
  Try
    Use (lcTable) In 0 Alias customers Shared
    llFailed = .F.
  Catch To loException
    llFailed = .T.
  Endtry
  If llFailed
    Return
  Endif
  With This
    .ColumnCount  = 7
    .RecordSource = 'customers'
    With .Column1
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "customers.customerid"
      .header1.Caption   = "Customer Id"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
    Endwith
    With .Column2
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "ALLTRIM(customers.companyname)"
      .header1.Caption   = "Company Name"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
    Endwith
    With .Column3
      .Resizable = .T.
      .Alignment = 2
      .ControlSource = "customers.contactname"
      .header1.Caption   = "Contact Name"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
    Endwith
    With .Column4
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "customers.address"
      .header1.Caption   = "Address"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
    Endwith
    With .Column5
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "customers.city"
      .header1.Caption   = "City"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
      .text1.Alignment   = 1
    Endwith
    With .Column6
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "customers.region"
      .header1.Caption   = "State"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
      .text1.Alignment   = 1
    Endwith
    With .Column7
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "customers.postalcode"
      .header1.Caption   = "Postal Code"
      .header1.FontBold  = .T.
      .header1.Alignment = 0
    Endwith
  Endwith
  Endproc


Enddefine
*-- EndDefine: grid1
**************************************************

Define Class commandbutton_L1_C3_TC4 As CommandButton

  Top = 1
  Left = 5
  Height = 30
  Width = 183
  Caption = "Export Grid to Excel"
  Name = "Command3"
  Visible = .T.


  Procedure Click
  Local lcExcel, loGrids
  lcExcel = Sys(5) + Addbs(Sys(2003)) + "NorthWindMultiDemo.xlsx"
  loGrids = Createobject("Empty")
  AddProperty(loGrids, "Count", 2)
  AddProperty(loGrids, "List[2, 4]", "")
  With loGrids
    .List[1, 1] = Thisform.grid1           && Grid object reference
    .List[1, 2] = "Customers"              && Sheet name
    .List[1, 3] = .T.                      && Freeze top row indicator
    .List[1, 4] = .F.                      && Include hidden columns indicator

    .List[2, 1] = Thisform.grid2
    .List[2, 2] = "Employees"
    .List[2, 3] = .T.
    .List[2, 4] = .F.
  Endwith

  Thisform.clsvfpxworkbookxlsx.SaveMultiGridToWorkbookEx(loGrids, lcExcel)
  Wait Window "Saved To Excel" Nowait

Enddefine
*-- EndDefine: command3
**************************************************

Define Class grid_L1_C4_TC5 As Grid

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

  Procedure Init && 1

  Try
    This.Column1.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column1.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column1.AddObject("header1","header")
    With This.Column1.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry


  Try
    This.Column1.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column1.AddObject("text1","textbox")
    With This.Column1.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry


  Try
    This.Column2.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column2.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column2.AddObject("header1","header")
    With This.Column2.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column2.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column2.AddObject("text1","textbox")
    With This.Column2.text1
      .FontName = "Arial"
      .FontSize = 9
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column3.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column3.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column3.AddObject("header1","header")
    With This.Column3.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column3.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column3.AddObject("text1","textbox")
    With This.Column3.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column4.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column4.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column4.AddObject("header1","header")
    With This.Column4.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column4.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column4.AddObject("text1","textbox")
    With This.Column4.text1
      .FontName = "Times New Roman"
      .FontSize = 10
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column5.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column5.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column5.AddObject("header1","header")
    With This.Column5.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column5.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column5.AddObject("text1","textbox")
    With This.Column5.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column6.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column6.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column6.AddObject("header1","header")
    With This.Column6.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column6.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column6.AddObject("text1","textbox")
    With This.Column6.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Try
    This.Column7.RemoveObject("text1")
  Catch
  Endtry

  Try
    This.Column7.RemoveObject("header1")
  Catch
  Endtry
  Try
    This.Column7.AddObject("header1","header")
    With This.Column7.header1
      .Caption = "Header1"
      .Name = "Header1"
    Endwith
  Catch
  Endtry

  Try
    This.Column7.RemoveObject("text1")
  Catch
  Endtry
  Try
    This.Column7.AddObject("text1","textbox")
    With This.Column7.text1
      .BorderStyle = 0
      .Margin = 0
      .ForeColor = Rgb(0,0,0)
      .BackColor = Rgb(255,255,255)
      .Name = "Text1"
      .Visible = .T.
    Endwith
  Catch
  Endtry

  Local lcTable, llFailed
  lcTable = Addbs(Sys(2004)) + "Samples\Northwind\employees.dbf"
  lcTable = Locfile(lcTable)
  Try
    Use (lcTable) In 0 Alias employees Shared
    llFailed = .F.
  Catch To loException
    llFailed = .T.
  Endtry
  If llFailed
    Return
  Endif
  With This
    .ColumnCount  = 7
    .RecordSource = 'employees'
    With .Column1
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.employeeid"
      With .header1
        .Caption   = "Employee Id"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column2
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "ALLTRIM(employees.lastname) + ', ' + ALLTRIM(employees.firstname)"
      With .header1
        .Caption   = "Employee Name"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column3
      .Resizable = .T.
      .Alignment = 2
      .ControlSource = "employees.hiredate"
      With .header1
        .Caption   = "Hire Date"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column4
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.address"
      With .header1
        .Caption   = "Address"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column5
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.city"
      With .header1
        .Caption   = "City"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
      .text1.Alignment   = 1
    Endwith
    With .Column6
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.region"
      With .header1
        .Caption   = "State"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
      .text1.Alignment   = 1
    Endwith
    With .Column7
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.postalcode"
      With .header1
        .Caption   = "Postal Code"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
  Endwith

Enddefine
*-- EndDefine: grid2
**************************************************
