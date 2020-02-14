Public oform1

Set Procedure To vfpxworkbookxlsx Additive

oform1 = Newobject("form1")
oform1.Show()
Return

Define Class form1 As Form

  DataSession = 2
  Height = 250
  Width = 761
  DoCreate = .T.
  AutoCenter = .T.
  Caption = "Test"
  MaxButton = .F.
  MinButton = .F.
  Name = "Form1"

  Add Object clsvfpxworkbookxlsx As vfpxworkbookxlsx With ;
    Height = 17, ;
    Left = 617, ;
    Top = 9, ;
    Width = 129, ;
    Name = "clsVFPxWorkbookXLSX"

  Add Object command1 As CommandButton With ;
    Top = 6, ;
    Left = 7, ;
    Height = 30, ;
    Width = 183, ;
    Caption = "Export Northwind Customers", ;
    Name = "Command1"

  Add Object command2 As CommandButton With ;
    Top = 37, ;
    Left = 7, ;
    Height = 30, ;
    Width = 183, ;
    Caption = "Class Demo", ;
    Name = "Command2"

  Add Object grid1 As Grid With ;
    DeleteMark = .F., ;
    Height = 168, ;
    Left = 7, ;
    ReadOnly = .T., ;
    Top = 75, ;
    Width = 749, ;
    Name = "Grid1"

  Add Object command3 As CommandButton With ;
    Top = 6, ;
    Left = 193, ;
    Height = 30, ;
    Width = 183, ;
    Caption = "Export Grid to Excel", ;
    Name = "Command3"

  Procedure Load
  Set Safety Off

  Procedure Resize
  With This
    .grid1.Width  = .Width - 14
    .grid1.Height = .Height - .grid1.Top - 7
  Endwith

  Procedure command1.Click
  Local lcTable, lcExcel
  lcExcel = Sys(5) + Addbs(Sys(2003)) + "Northwind Customers.xlsx"
  lcTable = Addbs(Sys(2004)) + "Samples\Northwind\customers.dbf"
  lcTable = Locfile(lcTable)
  If !Isnull(lcTable) .And. File(lcTable)
    Thisform.clsvfpxworkbookxlsx.SaveTabletoWorkbook(lcTable, lcExcel, .T., .T.)
    Wait Window "Saved To Excel" Nowait
  Endif

  Procedure command2.Click
  Thisform.clsvfpxworkbookxlsx.Demo()
  Wait Window "Saved To Excel" Nowait

  Procedure grid1.Init
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
      With .Header1
        .Caption   = "Employee Id"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column2
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "ALLTRIM(employees.lastname) + ', ' + ALLTRIM(employees.firstname)"
      With .Header1
        .Caption   = "Employee Name"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column3
      .Resizable = .T.
      .Alignment = 2
      .ControlSource = "employees.hiredate"
      With .Header1
        .Caption   = "Hire Date"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column4
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.address"
      With .Header1
        .Caption   = "Address"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
    With .Column5
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.city"
      With .Header1
        .Caption   = "City"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
      .Text1.Alignment   = 1
    Endwith
    With .Column6
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.region"
      With .Header1
        .Caption   = "State"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
      .Text1.Alignment   = 1
    Endwith
    With .Column7
      .Resizable = .T.
      .Alignment = 0
      .ControlSource = "employees.postalcode"
      With .Header1
        .Caption   = "Postal Code"
        .FontBold  = .T.
        .Alignment = 0
      Endwith
    Endwith
  Endwith

  Procedure command3.Click
  Local lcExcel
  lcExcel = Sys(5) + Addbs(Sys(2003)) + "Northwind Employees.xlsx"
  Thisform.clsvfpxworkbookxlsx.SaveGridToWorkbookEX(Thisform.grid1, lcExcel, .T., "Employees")
  Wait Window "Saved To Excel" Nowait

Enddefine
