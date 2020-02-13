**************************************************
*-- Class:        vfpxworkbookxlsx (vfpxworkbookxlsx.vcx)
*-- ParentClass:  custom
*-- BaseClass:    custom
*-- Time Stamp:   12.02.20 16:16:04
*
#INCLUDE "vfpxworkbookxlsx.h"
*
Define Class vfpxworkbookxlsx As Custom

*-- Indicates to auto-trim sheet name to max length ***Boolean
  AutoTrimSheetName = .T.
*-- CodePage to use for the Strings cursor
  Codepage = 0
*-- Company Name
  CompanyName = "VFPxWorkbookXLSX"
*-- Creator Name
  CreatorName = "VFPxWorkbookXLSX"
*-- Sets debugging mode ***Boolean
  Debug = .F.
*-- Flag to execute the DeclareDLL method in the Init() ***Boolean
  DeclareWinApi = .T.
*-- Workbook Default Font
  DefaultFont = "Calibri"
*-- Default font size ***Integer
  DefaultFontSize = 11
*-- Default sheet name (language specific)
  DefaultSheetName = "Sheet"
*-- Error level Id ***Integer
  ErrorLevelId = 0
*-- Release version of Class ***Integer
  ExcelXlsxRelease = 31
*-- XML DOM Object
  oXDom = ""
*-- Flag to save a currency value as a numeric rather than currency in the spreadsheet; will convert the value to four decimal places ***Boolean
  SaveCurrencyAsNumeric = .F.
*-- Subject of Workbook (stored in properties)
  Subject = ""
*-- Title of workbook (stored in properties)
  Title = ""
*-- The boolean value as displayed in Excel delimited by a pipe symbol
  TrueFalseValue = "Yes|No"
*-- User Name
  UserName = "VFPxWorkbookXLSX"
*-- XML Metadata for customizable properties
  _MemberData = [<VFPData>]+;
    [<memberdata name="autotrimsheetname" type="property" display="AutoTrimSheetName"/>]+;
    [<memberdata name="defaultfontsize" type="property" display="DefaultFontSize"/>]+;
    [<memberdata name="cellrefasciitoindex" type="method" display="CellRefAsciiToIndex"/>]+;
    [<memberdata name="getchecksum" type="method" display="GetCheckSum"/>]+;
    [<memberdata name="getstringrecord" type="method" display="GetStringRecord"/>]+;
    [<memberdata name="initalizeidvalues" type="method" display="InitalizeIdValues"/>]+;
    [<memberdata name="insertcell" type="method" display="InsertCell"/>]+;
    [<memberdata name="addclassdefinednumericformats" type="method" display="AddClassDefinedNumericFormats"/>]+;
    [<memberdata name="addcustomnumericformat" type="method" display="AddCustomNumericFormat"/>]+;
    [<memberdata name="addindexcolor" type="method" display="AddIndexColor"/>]+;
    [<memberdata name="addinlinefontobject" type="method" display="AddInlineFontObject"/>]+;
    [<memberdata name="addmrucolor" type="method" display="AddMruColor"/>]+;
    [<memberdata name="addnamedrange" type="method" display="AddNamedRange"/>]+;
    [<memberdata name="addnumericformat" type="method" display="AddNumericFormat"/>]+;
    [<memberdata name="addsheet" type="method" display="AddSheet"/>]+;
    [<memberdata name="addstringvalue" type="method" display="AddStringValue"/>]+;
    [<memberdata name="addstyleborders" type="method" display="AddStyleBorders"/>]+;
    [<memberdata name="addstylefill" type="method" display="AddStyleFill"/>]+;
    [<memberdata name="addstylefont" type="method" display="AddStyleFont"/>]+;
    [<memberdata name="addstylehorizalignment" type="method" display="AddStyleHorizAlignment"/>]+;
    [<memberdata name="addstyleindent" type="method" display="AddStyleIndent"/>]+;
    [<memberdata name="addstylenumericformat" type="method" display="AddStyleNumericFormat"/>]+;
    [<memberdata name="addstyletextrotation" type="method" display="AddStyleTextRotation"/>]+;
    [<memberdata name="addstylevertalignment" type="method" display="AddStyleVertAlignment"/>]+;
    [<memberdata name="addstylewordwrap" type="method" display="AddStyleWordWrap"/>]+;
    [<memberdata name="cellformatpainter" type="method" display="CellFormatPainter"/>]+;
    [</VFPData>]
  Name = "vfpxworkbookxlsx"

*-- Stores the last used border Id for a workbook
  Dimension LastBorderId[1,2]

*-- Stores the last used CellXfs Id for a workbook
  Dimension LastCellXfsId[1,2]

*-- Stores the last used fill Id for a workbook
  Dimension LastFillId[1,2]

*-- Stores the last used font Id for a workbook
  Dimension LastFontId[1,2]

*-- Last used IndexedColor Id
  Dimension LastIndexColorId[1,2]

*-- Last used mruColor Id value
  Dimension LastMruColorId[1,2]

*-- Last used NumFmt Id
  Dimension LastNumFmtId[1,2]

*-- Last used Relationship Id
  Dimension LastRelationId[1,2]

*-- Last used sheet Id
  Dimension LastSheetId[1,2]

*-- Stores the last used string Id for a workbook
  Dimension LastStringId[1,2]

*-- Returns the next string index value by workbook ***Integer
  Dimension StringNdx[1]

*-- Adds pre-defined numeric cell formats
  Protected Procedure AddClassDefinedNumericFormats
    Lparameters tnWB, tnFormatId
    Local llAdded, lcTextFormat, lcXMLFormat

    llAdded = Inlist(tnFormatId, CELL_FORMAT_CURRENCY_RED, CELL_FORMAT_ACC_CURR_POUNDS, CELL_FORMAT_ACC_CURR_EURO,;
      CELL_FORMAT_CURR_POUNDS_RED, CELL_FORMAT_CURR_EURO_RED)

    Do Case
    Case tnFormatId = CELL_FORMAT_CURRENCY_RED
      lcTextFormat = '"$"#,##0.00;[Red]("$"#,##0.00)'
      lcXMLFormat  = This.GetXMLString(lcTextFormat)

    Case tnFormatId = CELL_FORMAT_ACC_CURR_POUNDS
      lcXMLFormat  = '_-[$&#163;-809]* #,##0.00_-;-[$&#163;-809]* #,##0.00_-;_-[$&#163;-809]* &quot;-&quot;??_-;_-@_-'
      lcTextFormat = '"£"#,##0.00;[Red]("£"#,##0.00)'

    Case tnFormatId = CELL_FORMAT_ACC_CURR_EURO
      lcXMLFormat  = '_-[$&#8364;-2]* .00_-;-[$&#8364;-2]* .00_-;_-[$&#8364;-2]* &quot;-&quot;??_-;_-@_-'
      lcTextFormat = '"А"#,##0.00;[Red]("А"#,##0.00)'

    Case tnFormatId = CELL_FORMAT_CURR_POUNDS_RED
      lcXMLFormat  = '[$&#163;-809]#,##0.00;[Red]\-[$&#163;-809]#,##0.00'
      lcTextFormat = '"£"#,##0.00;[Red]("£"#,##0.00)'

    Case tnFormatId = CELL_FORMAT_CURR_EURO_RED
      lcXMLFormat  = '[$&#8364;-2]\ #,##0.00_);[Red]\([$&#8364;-2]\ #,##0.00\)'
      lcTextFormat = '"А"#,##0.00;[Red]("А"#,##0.00)'
    Endcase

    If llAdded
      Insert Into xl_numFmts (workbook, Id, formatxml, formatcode, applydec) ;
        VALUES (tnWB, tnFormatId, lcXMLFormat, lcTextFormat, False)
    Endif
    Return llAdded

*-- Adds a new custom defined numeric format
  Procedure AddCustomNumericFormat
  Lparameters tnWB, tcPosFormat, tcNegFormat, tcZeroFormat, tcTextFormat, tlApplyDec
  Local lnId, lcFormatCode, lnNdx, lcSetPoint, lcSeparator
  If Vartype(tcPosFormat) != "C"
    Return 0
  Endif
  Do Case
  Case Pcount() = 1
    tcNegFormat  = ""
    tcZeroFormat = ""
    tcTextFormat = ""
    tlApplyDec   = False

  Case Pcount() = 2
    tcZeroFormat = ""
    tcTextFormat = ""
    tlApplyDec   = False

  Case Pcount() = 3
    tcTextFormat = ""
    tlApplyDec   = False

  Case Pcount() = 4
    tlApplyDec = False
  Endcase
  If Vartype(tcNegFormat) != "C"
    tcNegFormat = ""
  Endif
  If Vartype(tcZeroFormat) != "C"
    tcZeroFormat = ""
  Endif
  If Vartype(tcTextFormat) != "C"
    tcTextFormat = ""
  Endif
  If Occurs(";", tcPosFormat) > 0 Or Occurs(";", tcNegFormat) > 0 Or Occurs(";", tcZeroFormat) > 0 Or Occurs(";", tcTextFormat) > 0
    Return 0
  Endif
*-*	If not US standard for separator and decimal point, then change to US standard
  lcSetPoint  = Set("POINT")
  lcSeparator = Set("SEPARATOR")
  If lcSetPoint != "." Or lcSeparator != ","
    tcPosFormat = Chrtran(Chrtran(tcPosFormat, lcSetPoint + lcSeparator, Tab + CR), Tab + CR, ".,")
    tcNegFormat = Chrtran(Chrtran(tcNegFormat, lcSetPoint + lcSeparator, Tab + CR), Tab + CR, ".,")
    tcZeroFormat = Chrtran(Chrtran(tcZeroFormat, lcSetPoint + lcSeparator, Tab + CR), Tab + CR, ".,")
    tcTextFormat = Chrtran(Chrtran(tcTextFormat, lcSetPoint + lcSeparator, Tab + CR), Tab + CR, ".,")
  Endif
*-*	Build expression from parts
  lcFormatCode = tcPosFormat
  If !Empty(tcNegFormat)
    lcFormatCode = lcFormatCode + ";" + tcNegFormat
  Endif
  If !Empty(tcZeroFormat)
    lcFormatCode = lcFormatCode + ";" + tcZeroFormat
  Endif
  If !Empty(tcTextFormat)
    lcFormatCode = lcFormatCode + ";" + tcTextFormat
  Endif
  lnId = This.GetNextId(tnWB, "xl_numfmts")
  Insert Into xl_numFmts (workbook, Id, formatxml, formatcode, applydec) ;
    VALUES (tnWB, lnId, This.GetXMLString(lcFormatCode), lcFormatCode, tlApplyDec)
  Return lnId

*-- Adds a new indexed color definition
  Procedure AddIndexColor
  Lparameters tnWB, tnRGBColor
  Local lnIndexId, lcHexColor
  If Pcount() != 2 Or Vartype(tnRGBColor) != "N"
    Return 0
  Endif
  lcHexColor = This.ConvertColorToHex(tnRGBColor)
  If Seek(BinToC(tnWB)+lcHexColor, "xl_ndxcolors", "rgbcolor")
    lnIndexId = xl_ndxcolors.indexid
  Else
    lnIndexId = This.GetNextId(tnWB, "xl_ndxcolors")
    Insert Into xl_ndxcolors (workbook, indexid, rgbcolor) Values (tnWB, lnIndexId, lcHexColor)
  Endif
  Return lnIndexId

*-- Adds an object for setting the inline formatting of cell text
  Procedure AddInlineFontObject
  Lparameters toInline, tnBeg, tnLen, tcFontName, tnFontSize, tnFontColor, tlFontBold, tlFontItalic, tcULine, tlStrkThru, tlSubscript, tlSuperscript
  Local loCharacter
  If Pcount() < 4 Or Vartype(toInline) != "O"
    Return Null
  Endif
  Do Case
  Case Pcount() = 4
    tnFontSize    = 10
    tnFontColor   = Rgb(0,0,0)
    tlFontBold    = False
    tlFontItalic  = False
    tcULine       = UNDERLINE_NONE
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 5
    tnFontColor   = Rgb(0,0,0)
    tlFontBold    = False
    tlFontItalic  = False
    tcULine       = UNDERLINE_NONE
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 6
    tlFontBold    = False
    tlFontItalic  = False
    tcULine       = UNDERLINE_NONE
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 7
    tlFontItalic  = False
    tcULine       = UNDERLINE_NONE
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 8
    tcULine       = UNDERLINE_NONE
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 9
    tlStrkThru    = False
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 10
    tlSubscript   = False
    tlSuperscript = False

  Case Pcount() = 11
    tlSuperscript = False
  Endcase
  loCharacter = Createobject("Empty")
  AddProperty(loCharacter, "BegPos", tnBeg)
  AddProperty(loCharacter, "Length", tnLen)
  AddProperty(loCharacter, "FontName", tcFontName)
  AddProperty(loCharacter, "FontSize", tnFontSize)
  AddProperty(loCharacter, "FontBold", tlFontBold)
  AddProperty(loCharacter, "FontItalic", tlFontItalic)
  AddProperty(loCharacter, "FontColor", tnFontColor)
  AddProperty(loCharacter, "Underline", tcULine)
  AddProperty(loCharacter, "StrikeThru", tlStrkThru)
  AddProperty(loCharacter, "SubScript", tlSubscript)
  AddProperty(loCharacter, "SuperScript", tlSuperscript)

  toInline.Count = toInline.Count + 1
  Dimension toInline.Characters[toInline.Count]
  toInline.Characters[toInline.Count] = loCharacter

  Return toInline

*-- Adds a custom defined MRU color to the workbook
  Procedure AddMruColor
  Lparameters tnWB, tnRGBColor
  Local lnIndexId, lcHexColor
  If Pcount() != 2 Or Vartype(tnRGBColor) != "N"
    Return 0
  Endif
  With This
    lcHexColor = .ConvertColorToHex(tnRGBColor)
    If Seek(BinToC(tnWB)+lcHexColor, "xl_mrucolors", "rgbcolor")
      lnIndexId = xl_mrucolors.indexid
    Else
      lnIndexId = .GetNextId(tnWB, "xl_mrucolors")
      Insert Into xl_mrucolors (workbook, indexid, rgbcolor) Values (tnWB, lnIndexId, lcHexColor)
    Endif
  Endwith
  Return lnIndexId

*-- Adds a name range to the workbook
  Procedure AddNamedRange
  Lparameters tnWB, tnSheet, tcName, tnScope, tcComment, tnBegRow, tnBegCol, tnEndRow, tnEndCol
  Local lcRangeName
  If Pcount() < 8 Or Vartype(tcName) != "C" Or Len(tcName) > 50 Or Len(tcComment) > 254 Or tnBegRow > tnEndRow Or tnBegCol > tnEndCol
    Return ""
  Endif
  If Vartype(tnScope) != "N"
    tnScope = 0
  Endif
  If Vartype(tcComment) != "C"
    tcComment = ""
  Endif
  lcRangeName = Chrtran(Alltrim(tcName), " !@#$%^&*()+={}[]|<>,.'?/" + Chr(34), Replicate("_", 25))
  If !Isalpha(lcRangeName) And !Inlist(Left(lcRangeName, 1), "_", "\")
    lcRangeName = "_" + lcRangeName
  Endif
  If Seek(BinToC(tnWB), "xl_workbooks", "workbook") And !Seek(BinToC(tnWB)+Padl(lcRangeName, 254, " "), "xl_namerange", "wbrname")
    Insert Into xl_namerange (workbook, Sheet, rname, scope, Comment, begrow, begcol, endrow, endcol) ;
      VALUES (tnWB, tnSheet, lcRangeName, tnScope, tcComment, tnBegRow, tnBegCol, tnEndRow, tnEndCol)
  Else
    lcRangeName = ""
  Endif
  Return lcRangeName

*-- Adds a numeric format expression to be used
  Procedure AddNumericFormat
  Lparameters tnWB, tcFormatCode
  Local lcPosCode, lcNegCode, lcZerCode
  If Pcount() < 2 Or (Vartype(tnWB) != "N" And tnWB > 0) Or Vartype(tcFormatCode) != "C"
    Return 0
  Endif
  With This
    lcPosCode = Getwordnum(tcFormatCode, 1, ";")
    lcNegCode = Getwordnum(tcFormatCode, 2, ";")
    lcZerCode = Getwordnum(tcFormatCode, 3, ";")
    Return .AddCustomNumericFormat(tnWB, lcPosCode, lcNegCode, lcZerCode)
  Endwith

*-- Adds a new sheet to the workbook
  Procedure AddSheet
  Lparameters tnWB, tcSheetName, tnState
  Local lnShId, lnRelId
  lnShId = 0
  With This
    If Pcount() > 0 And Seek(BinToC(tnWB), "xl_workbooks", "workbook")
      lnShId = .GetNextId(tnWB, "xl_sheets")
      Do Case
      Case Pcount() = 0

      Case Pcount() = 1
        tcSheetName = "Sheet" + Transform(lnShId + 1)        && Bug fix identified by Dan Lauer
        tnState = VISIBLE_SHEET_STATE

      Case Pcount() = 2
        If Vartype(tcSheetName) != "C"
          tcSheetName = "Sheet" + Transform(lnShId + 1)
        Endif
        tnState = VISIBLE_SHEET_STATE

      Otherwise
        If Vartype(tcSheetName) != "C"
          tcSheetName = "Sheet" + Transform(lnShId + 1)
        Endif
        If Vartype(tnState) != "N"
          tnState = VISIBLE_SHEET_STATE
        Endif
      Endcase
      If Empty(tcSheetName)
        tcSheetName = "Sheet" + Transform(lnShId + 1)
      Endif
      tcSheetName = Chrtran(tcSheetName, ":\/?*", "_____")
      If Len(tcSheetName) > LIMITS_MAX_SH_NAME
        If .AutoTrimSheetName
          tcSheetName = Left(Alltrim(tcSheetName), LIMITS_MAX_SH_NAME)
        Else
          .SetLastId(tnWB, lnShId-1, "xl_sheets")
          Return 0
        Endif
      Endif
      If Seek(BinToC(tnWB)+Upper(Padr(tcSheetName, Len(xl_sheets.shname))), "xl_sheets", "shname")    && Change recommendation by Doug Hennig
        lnShId = 0
      Else
        Insert Into xl_sheets (workbook, Sheet, shname, state, mleft, mright, mtop, mbot, mheader, mfooter, shdeleted, xsplit, ysplit, prnorient, ;
          papersize, Paperwidth, paperheight, Scale, fittowidth, fittoheight) ;
          VALUES (tnWB, lnShId, tcSheetName, tnState, 0.75, 0.75, 0.75, 0.75, 0.3, 0.3, False, 0, 0, PORTRAIT_PRINT_ORIENTATION, 0, 0, 0, 100, 0, 0)

        lnRelId = This.GetNextId(tnWB, "xl_relationships")
        Insert Into xl_relationships (workbook, Sheet, relid, reltype, Target) Values (tnWB, lnShId, lnRelId, "worksheet", "worksheets/sheet" + Transform(lnShId) + ".xml")
      Endif
    Endif
  Endwith
  Return lnShId

*-- Adds a string value to the internal cursor
  Protected Procedure AddStringValue
    Lparameters tnWB, tcString, tlInLine
    Local lcStringXml, lnStringId, llPresrvSp, lcCheckSum
    With This
      If tlInLine
        llPresrvSp  = Iif(Left(tcString, 1) = " ", True, False)
        lcStringXml = .GetXMLString(tcString)
        lnStringId  = .GetNextId(tnWB, "xl_strings")
        lcCheckSum  = .GetCheckSum(tcString)
        Insert Into xl_strings (Id, workbook, checksum, stringxml, stringval, presvspace, Formatted) ;
          VALUES (lnStringId, tnWB, lcCheckSum, lcStringXml, tcString, llPresrvSp, True)
      Else
        If .GetStringRecord(tnWB, tcString) And xl_strings.Formatted = False
          lnStringId = xl_strings.Id
        Else
          llPresrvSp  = Iif(Left(tcString, 1) = " " Or Right(tcString, 1) = " ", True, False)
          lcStringXml = .GetXMLString(tcString)
          lnStringId  = .GetNextId(tnWB, "xl_strings")
          lcCheckSum  = .GetCheckSum(tcString)
          Insert Into xl_strings (Id, workbook, checksum, stringxml, stringval, presvspace, Formatted) ;
            VALUES (lnStringId, tnWB, lcCheckSum, lcStringXml, tcString, llPresrvSp, False)
        Endif
      Endif
    Endwith
    Return lnStringId

*-- Adds to the cell style a border definition
  Procedure AddStyleBorders
  Lparameters tnWB, tnCellXfsId, tnBorders, tcStyle, tnColor
  Local lnDiagDn, lnDiagUp, lcLStyle, lnLColor, lcRStyle, lnRColor, lcTStyle, lnTColor, lcBStyle, lnBColor
  Local lcDStyle, lnDColor, lnBorderId
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tnBorders) != 'N' Or Empty(tnBorders) Or !Between(tnBorders, 0, 63)
      tnBorders = 0
    Endif
    If !This.IsValidBorderStyle(tcStyle)
      tcStyle = BORDER_STYLE_NONE
    Endif
    If Vartype(tnColor) != 'N' Or !Between(tnColor, 0, Rgb(255,255,255)) Or Empty(tnColor)
      tnColor = Rgb(0, 0, 0)
    Endif
    Store 0 To lnDiagDn, lnDiagUp
    Store Rgb(0,0,0) To lnLColor, lnRColor, lnTColor, lnBColor, lnDColor
    Store BORDER_STYLE_NONE To lcLStyle, lcRStyle, lcTStyle, lcBStyle, lcDStyle
    If Bittest(tnBorders, 0)
      lcLStyle = tcStyle
      lnLColor = tnColor
    Endif
    If Bittest(tnBorders, 1)
      lcRStyle = tcStyle
      lnRColor = tnColor
    Endif
    If Bittest(tnBorders, 2)
      lcTStyle = tcStyle
      lnTColor = tnColor
    Endif
    If Bittest(tnBorders, 3)
      lcBStyle = tcStyle
      lnBColor = tnColor
    Endif
    If Bittest(tnBorders, 4)
      lcDStyle = tcStyle
      lnDColor = tnColor
      lnDiagDn = 1
    Endif
    If Bittest(tnBorders, 5)
      lcDStyle = tcStyle
      lnDColor = tnColor
      lnDiagUp = 1
    Endif
    If This.GetBordersRecord(tnWB, lcLStyle, lnLColor, lcRStyle, lnRColor, lcTStyle, lnTColor, lcBStyle, lnBColor, lcDStyle, lnDColor, lnDiagDn, lnDiagUp)
      lnBorderId = xl_borders.Id
    Else
      lnBorderId = This.GetNextId(tnWB, 'xl_borders')
      Insert Into xl_borders (workbook, Id, lstyle, lcolor, rstyle, rcolor, tstyle, tcolor, bstyle, bcolor, dstyle, dcolor, diagdn, diagup) ;
        VALUES (tnWB, lnBorderId, lcLStyle, lnLColor, lcRStyle, lnRColor, lcTStyle, lnTColor, lcBStyle, lnBColor, lcDStyle, lnDColor, lnDiagDn, lnDiagUp)
    Endif
*-*	Set the border id to the style definition
    Replace xl_cellxfs.borderid With lnBorderId In xl_cellxfs
    Return True
  Else
    Return False
  Endif

*-- Adds to the cell style a fill definition
  Procedure AddStyleFill
  Lparameters tnWB, tnCellXfsId, tnFColor, tnBColor, tcPatternType
  Local lnFillId, lnFIndex, lnBIndex
  Do Case
  Case Pcount() < 3
    Return False

  Case Pcount() = 3
    tnBColor = tnFColor
    tcPatternType = FILL_STYLE_SOLID

  Case Pcount() = 4
    tcPatternType = FILL_STYLE_SOLID
  Endcase
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tnFColor) = 'N' And Between(tnFColor, 0, Rgb(255, 255, 255))
      tnFColor = Floor(tnFColor)
      lnFIndex = 0
    Else
      tnFColor = 0
      lnFIndex = 64
    Endif
    If Vartype(tnBColor) = 'N' And Between(tnBColor, 0, Rgb(255, 255, 255))
      tnBColor = Floor(tnBColor)
      lnBIndex = 0
    Else
      tnBColor = tnFColor
      lnBIndex = lnFIndex
    Endif
    If Vartype(tcPatternType) != "C" Or Empty(tcPatternType)
      tcPatternType = FILL_STYLE_SOLID
    Endif
*-*	Get the fill record id
    If Seek(BinToC(tnWB)+BinToC(tnFColor)+BinToC(tnBColor)+tcPatternType, "xl_fills", "fillcolor")
      lnFillId = xl_fills.Id
      Replace xl_fills.Theme     With 0, ;
        xl_fills.tint      With 0, ;
        xl_fills.fgindexed With lnFIndex, ;
        xl_fills.bgindexed With lnBIndex In xl_fills
    Else
      lnFillId = This.GetNextId(tnWB, 'xl_fills')
      Insert Into xl_fills (workbook, Id, fgcolor, bgcolor, patttype, Theme, tint, fgindexed, bgindexed) ;
        VALUES (tnWB, lnFillId, tnFColor, tnBColor, tcPatternType, 0, 0, lnFIndex, lnBIndex)
    Endif
*-*	Set the fill id to the style definition
    Replace xl_cellxfs.fillid With lnFillId In xl_cellxfs
    Return True
  Else
    Return False
  Endif

*-- Adds to the cell style a font definition
  Procedure AddStyleFont
  Lparameters tnWB, tnCellXfsId, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos
  Local lnFontId
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tcFName) != 'C' Or Empty(tcFName)
      tcFName = 'Calibri'
    Endif
    tcFName = Proper(Alltrim(tcFName))
    tnFSize = Iif(Vartype(tnFSize) != 'N' Or Empty(tnFSize) Or tnFSize < 0.5, 11, Int(tnFSize))
    If Vartype(tlBold) != 'L'
      tlBold = False
    Endif
    If Vartype(tlItalic) != 'L'
      tlItalic = False
    Endif
    tnFColor = Iif(Vartype(tnFColor) = 'N' And Between(tnFColor, 0, Rgb(255, 255, 255)), Floor(tnFColor), Rgb(0, 0, 0))
    If Vartype(tlStrikThr) != 'L'
      tlStrikThr = False
    Endif
    If Vartype(tcULine) != 'C' Or !Inlist(tcULine, UNDERLINE_SINGLE, UNDERLINE_DOUBLE, UNDERLINE_SINGLEACCOUNTING, UNDERLINE_DOUBLEACCOUNTING, UNDERLINE_NONE)
      tcULine = UNDERLINE_NONE
    Endif
    If Vartype(tcVPos) != 'C' Or !Inlist(tcVPos, FONT_VERTICAL_BASELINE, FONT_VERTICAL_SUBSCRIPT, FONT_VERTICAL_SUPERSCRIPT)
      tcVPos = FONT_VERTICAL_BASELINE
    Endif
*-*	Get the font record id
    If Seek(BinToC(tnWB)+Padr(tcFName, 100) + Str(tnFSize, 5, 1) + Transform(tlBold) + Transform(tlItalic) + Padl(tnFColor, 15) + Padr(tcULine, 16) + ;
        TRANSFORM(tlStrikThr) + Padr(tcVPos, 11), "xl_fonts", "cellformat")

      lnFontId = xl_fonts.Id
      Replace xl_fonts.Theme   With 0, ;
        xl_fonts.tint    With 0, ;
        xl_fonts.indexed With 0 In xl_fonts
    Else
      lnFontId = This.GetNextId(tnWB, 'xl_fonts')
      Insert Into xl_fonts (workbook, Id, fname, Fsize, fbold, fitalic, fcolor, uline, strkthr, fvpos) ;
        VALUES (tnWB, lnFontId, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos)
    Endif
*-*	Set the font id to the style definition
    Replace xl_cellxfs.fontid With lnFontId In xl_cellxfs
    Return True
  Else
    Return False
  Endif

*-- Adds to the cell style horizontal cell alignment
  Procedure AddStyleHorizAlignment
  Lparameters tnWB, tnCellXfsId, tcHorizAlign
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tcHorizAlign) = 'C' And Inlist(tcHorizAlign, CELL_HORIZ_ALIGN_LEFT, CELL_HORIZ_ALIGN_RIGHT, CELL_HORIZ_ALIGN_CENTER)
      Replace xl_cellxfs.halign With tcHorizAlign In xl_cellxfs
      Return True
    Else
      Return False
    Endif
  Else
    Return False
  Endif

*-- Adds to the style indent definition
  Procedure AddStyleIndent
  Lparameters tnWB, tnCellXfsId, tnIndent
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tnIndent) = 'N'
      Replace xl_cellxfs.indent With tnIndent In xl_cellxfs
      Return True
    Else
      Return False
    Endif
  Else
    Return False
  Endif

*-- Adds to the style definition numeric format
  Procedure AddStyleNumericFormat
  Lparameters tnWB, tnCellXfsId, tnNumFmtId
  Local lnNumFmtId
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    Do Case
    Case Vartype(tnNumFmtId) != 'N'
      Return False

    Case tnNumFmtId < START_NUMERIC_FORMAT_ID
      lnNumFmtId = tnNumFmtId

    Case tnNumFmtId >= START_NUMERIC_FORMAT_ID
      If Seek(BinToC(tnWB)+BinToC(tnNumFmtId), "xl_numfmts", "id")
        lnNumFmtId = xl_numFmts.Id
      Else
        If !This.AddClassDefinedNumericFormats(tnWB, tnNumFmtId)
          Return False
        Endif
        lnNumFmtId = tnNumFmtId
      Endif

    Otherwise
      Return False
    Endcase
    Replace xl_cellxfs.numfmtid With lnNumFmtId In xl_cellxfs
    Return True
  Else
    Return False
  Endif

*-- Adds to the style definition text rotation
  Procedure AddStyleTextRotation
  Lparameters tnWB, tnCellXfsId, tnRotation
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tnRotation) = 'N' And Between(tnRotation, -90, 90)
      Replace xl_cellxfs.Rotation With tnRotation In xl_cellxfs
      Return True
    Else
      Return False
    Endif
  Else
    Return False
  Endif

*-- Adds to the cell style vertical cell alignment
  Procedure AddStyleVertAlignment
  Lparameters tnWB, tnCellXfsId, tcVertAlign
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tcVertAlign) = 'C' And Inlist(tcVertAlign, CELL_VERT_ALIGN_TOP, CELL_VERT_ALIGN_BOTTOM, CELL_VERT_ALIGN_CENTER)
      Replace xl_cellxfs.valign With tcVertAlign In xl_cellxfs
      Return True
    Else
      Return False
    Endif
  Else
    Return False
  Endif

*-- Adds to the style definition word wrap
  Procedure AddStyleWordWrap
  Lparameters tnWB, tnCellXfsId, tlWordWrap
  If Pcount() < 3
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If Vartype(tlWordWrap) = 'L'
      Replace xl_cellxfs.wraptext With Iif(tlWordWrap, 1, 0) In xl_cellxfs
      Return True
    Else
      Return False
    Endif
  Else
    Return False
  Endif

*-- Copies the formatting of a source cell to other cells
  Procedure CellFormatPainter
  Lparameters tnWB, tnSheet, tnSrcRow, tnSrcCol, tnBegRow, tnBegCol, tnEndRow, tnEndCol
  Local lnCellXfs, lnNumDec, llReturn, lnRow, lnCol
  Do Case
  Case Pcount() < 6
    Return False

  Case Pcount() = 6
    tnEndRow = tnBegRow
    tnEndCol = tnBegCol

  Case Pcount() = 7
    tnEndCol = tnBegCol
  Endcase
  If This.GetCellRecord(tnWB, tnSheet, tnSrcRow, tnSrcCol)
    lnCellXfs = xl_cells.cellxfs
    lnNumDec  = xl_cells.numdec
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If This.GetCellRecord(tnWB, tnSheet, lnRow, lnCol)
          Replace xl_cells.cellxfs With lnCellXfs, ;
            xl_cells.numdec  With lnNumDec In xl_cells
        Else
          Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
            VALUES (tnWB, tnSheet, lnRow, lnCol, DATA_TYPE_NONE, lnCellXfs, False, lnNumDec, 0)
        Endif
      Endfor
    Endfor
    llReturn = True
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Converts a 'AA444' cell reference to the row and column index values
  Procedure CellRefAsciiToIndex
  Lparameters tcCellRef
  Local lnCellCol, lnNdx, lnCellRow, loCellRef
  lnCellCol = This.ColumnAsciiToIndex(tcCellRef)
  lnCellRow = 0
  For lnNdx=1 To Len(tcCellRef)
    If Isdigit(Substr(tcCellRef, lnNdx))
      lnCellRow = Int(Val(Substr(tcCellRef, lnNdx)))
      Exit
    Endif
  Endfor
  loCellRef = Createobject("Empty")
  AddProperty(loCellRef, "Column", lnCellCol)
  AddProperty(loCellRef, "Row", lnCellRow)
  Return loCellRef

*-- Clears/deletes the named range from the workbook
  Procedure ClearCellValidation
  Lparameters tnWB, tcName
  If Pcount() < 2
    Return False
  Endif
  If Vartype(tcName) != "C"
    Return False
  Endif
  If Seek(BinToC(tnWB)+Padl(tcName, 254, " "), "xl_namerange", "wbrname")
    Delete In xl_namerange
    llReturn = True
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Clears the cell value
  Procedure ClearCellValue
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local llReturn
  If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
    Replace xl_cells.celldeleted With True In xl_cells
    llReturn = True
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Converts Excel Ascii column value to index value
  Procedure ColumnAsciiToIndex
  Lparameters tcColIndex
  Local lnIndex, lnNdx, lnChr
  lnIndex = 0
  For lnNdx=1 To Len(tcColIndex)
    lnChr = Asc(Upper(Substr(tcColIndex, lnNdx, 1)))
    If Between(lnChr, 65, 90)
      lnIndex = (lnChr - 64) + (lnIndex * 26)
    Else
      Exit
    Endif
  Endfor
  Return lnIndex

*-- Gets the column ASCII code for the column index
  Procedure ColumnIndexToAscii
  Lparameters tnCol
  Local lnRemCol, lcColumn, lnPartCol
  If Pcount() = 0 Or Vartype(tnCol) != "N"
    Return ""
  Endif
  Try
    lnRemCol = tnCol
    lcColumn = ""
    Do While lnRemCol > 0
      lnPartCol = Mod(lnRemCol, 26)
      If lnPartCol = 0
        lnPartCol = 26
      Endif
      lnRemCol = (lnRemCol - lnPartCol) / 26
      lcColumn = Chr(lnPartCol + 64)  + lcColumn
    Enddo

  Catch To loException
    Set Step On
  Endtry
  Return lcColumn

*-- Converts a color value (integer) to Hex representation
  Protected Procedure ConvertColorToHex
    Lparameters tnColor
    Local lnRed, lnGreen, lnBlue, lcHexRed, lcHexGreen, lcHexBlue
    With This
      lnRed   = .GetRGBValues(tnColor, 'R')
      lnGreen = .GetRGBValues(tnColor, 'G')
      lnBlue  = .GetRGBValues(tnColor, 'B')
      lcHexRed   = Right(Transform(lnRed, '@0'), 2)
      lcHexGreen = Right(Transform(lnGreen, '@0'), 2)
      lcHexBlue  = Right(Transform(lnBlue, '@0'), 2)
    Endwith
    Return "FF" + Upper(lcHexRed + lcHexGreen + lcHexBlue)

*-- ConvertHexStringToNumeric
  Protected Procedure ConvertHexStringToNumeric
    Lparameters tcHexNum
    Local lcValue
    If Empty(tcHexNum)
      Return 0
    Endif
    If Len(tcHexNum) > 6
      tcHexNum = Substr(tcHexNum, 3)
    Endif
    tcHexNum = Upper(Alltrim(tcHexNum))
    tcHexNum = Right(tcHexNum, 2) + Substr(tcHexNum, 3, 2) + Left(tcHexNum, 2)
    lcValue  = "0x" + tcHexNum
    Return Eval(lcValue)

*-- Converts Pixel measurement to Excel measurement
  Procedure ConvertPixelsToExcelUnits
  Lparameters tnPixels
  Return tnPixels * 0.152542

*-- Creates the Excel file from the components
  Protected Procedure CreateExcelFile
    Lparameters tcTempPath, tcFilePath, tcWBName
    Local loShell, loFolder, lcZipFile, loException, loFile, lnCountBefore, lcFile, lcZipPath, lnCount, lhFile
    lcZipPath = Alltrim(tcFilePath)
    If Empty(lcZipPath)
      lcZipPath = Sys(5) + Sys(2003)
    Endif
    lcZipFile = Addbs(lcZipPath) + Forceext(Alltrim(tcWBName), "zip")
*-*	Delete any prior file with same Excel name
    If File(Forceext(lcZipFile, "xlsx"))
      Try
        Erase (Forceext(lcZipFile, "xlsx"))
        llError = False

      Catch To loException
        This.ErrorLevelId = 10
        Raiseevent(This, "OnShowErrorMessage", 10, "Error - unable to delete existing file" + CR + loException.Message)
        llError = True
      Endtry
      If llError
        Return False
      Endif
    Endif
*-*	Validate that the file name and path are valid
    Try
      lhFile = Fcreate(lcZipFile, 0)
      Fclose(lhFile)

    Catch To loException
      This.ErrorLevelId = 11
      Raiseevent(This, "OnShowErrorMessage", 11, "CreateExcelFile - Failed to create Zip file"+ CR + loException.Message)
      Return False
    Endtry

*-*	Create base zip file
    Strtofile(Chr(80)+Chr(75)+Chr(5)+Chr(6)+Replicate(Chr(0), 18), lcZipFile, 0)
    loShell  = Createobject("shell.application")
    loFolder = loShell.NameSpace(tcTempPath).Items

*-*	Add the files to the zip
    If Os(3)<'6' Or Os(3)='6' And Os(4)<'1'
      Try
        For Each loFile In loFolder
          loShell.NameSpace(lcZipFile).MoveHere(loFile, FOF_SILENT)
          apiSleep(100)
        Endfor

      Catch To loException
        Set Step On
      Endtry

      llErr = True
      Do While llErr
        Try
          apiMoveFile(lcZipFile, Forceext(lcZipFile, "xlsx"))
          llErr = False
        Catch
          apiSleep(100)
        Endtry
      Enddo
    Else
      Try
        For Each loFile In loFolder
          lnCountBefore = loShell.NameSpace(tcTempPath).Items.Count
          loShell.NameSpace(lcZipFile).MoveHere(loFile, FOF_SILENT)
          Do While lnCountBefore = loShell.NameSpace(tcTempPath).Items.Count
            apiSleep(50)
          Enddo
        Endfor

      Catch To loException
        This.ErrorLevelId = 12
        Raiseevent(This, "OnShowErrorMessage", 12, "CreateExcelFile - Failed to add contents to Zip file"+ CR + loException.Message)
        Set Step On
        Erase (lcZipFile)
        Do Case
        Case loException.ErrorNo = 1943

        Otherwise
          Set Step On
        Endcase
      Endtry
      Try
        apiMoveFile(lcZipFile, Forceext(lcZipFile, "xlsx"))

      Catch To loException
        This.ErrorLevelId = 13
        Raiseevent(This, "OnShowErrorMessage", 13, "CreateExcelFile - Rename failed"+ CR + loException.Message)
        Set Step On
      Endtry
    Endif
    Return True

*-- Creates a new format style definition
  Procedure CreateFormatStyle
  Lparameters tnWB
  If Pcount() > 0 And Seek(BinToC(tnWB), "xl_workbooks", "workbook")
    lnId = This.GetNextId(tnWB, "xl_cellxfs")
    Insert Into xl_cellxfs (workbook, Id, numfmtid, fontid, fillid, borderid, halign, valign, indent, wraptext, Rotation) ;
      VALUES (tnWB, lnId, 0, 0, 0, 0, "", "", 0, 0, 0)
    Return lnId
  Else
    Return -1
  Endif

*-- Creates an in-line text formatted definition
  Procedure CreateInlineFormatText
  Lparameters tnWB, tcCellText
  Local loInline
  loInline = Createobject("Empty")
  AddProperty(loInline, "Workbook", tnWB)
  AddProperty(loInline, "StringId", Null)
  AddProperty(loInline, "StringValue", tcCellText)
  AddProperty(loInline, "Count", 0)
  AddProperty(loInline, "Characters[1]")
  loInline.Characters[1] = Null
  Return loInline

*-- Creates a new workbook object
  Procedure CreateWorkbook
  Lparameters tcName
  Local lcWBName, lcPath, lnWB
  If Pcount() = 1
    tlLoadDefaults = True
  Endif
  lcWBName = Juststem(tcName)
  lcPath   = Justpath(tcName)
  If Empty(lcPath)
    lcPath = Sys(5) + Sys(2003)
    tcName = Addbs(lcPath) + lcWBName + "." + Justext(tcName)
  Endif
  With This
    If Seek(Padr(Upper(tcName), 240, "_"), "xl_workbooks", "fullname")
      lnWB = 0
    Else
      Insert Into xl_workbooks (wbname, filepath, FullName) Values (lcWBName, lcPath, Padr(Upper(tcName), 240, "_"))

      lnWB = xl_workbooks.workbook
      .SetDefaultFont(lnWB, This.DefaultFont, This.DefaultFontSize)        && Change recommended by Doug Hennig; 2017-06-12
      .SetDefaultBorder(lnWB, BORDER_STYLE_NONE)

      Insert Into xl_relationships (workbook, relid, reltype, Target) Values (lnWB, 1, "styles", "styles.xml")
      Insert Into xl_relationships (workbook, relid, reltype, Target) Values (lnWB, 2, "sharedStrings", "sharedStrings.xml")
      .SetLastId(lnWB, 2, "xl_relationships")

      Insert Into xl_cellxfs (workbook, Id, numfmtid, fontid, fillid, borderid, halign, valign, indent, wraptext) ;
        VALUES (lnWB, 0, 0, 0, 0, 0, "", "", 0, 0)

      Insert Into xl_fills (workbook, Id, fgcolor, bgcolor, patttype, Theme, tint, fgindexed, bgindexed) ;
        VALUES (lnWB, 0, 0, 0, FILL_STYLE_NONE, 0, 0, 0, 0)

      Insert Into xl_fills (workbook, Id, fgcolor, bgcolor, patttype, Theme, tint, fgindexed, bgindexed) ;
        VALUES (lnWB, 1, 0, 0, FILL_STYLE_GRAY125, 0, 0, 0, 0)
    Endif
  Endwith
  Return lnWB

*-- Internal method for creating a workbook
  Protected Procedure CreateBorkbookEx
    Lparameters tcName
    Local lcWBName, lcPath, lnWB
    lcWBName = Juststem(tcName)
    lcPath   = Justpath(tcName)
    If Empty(lcPath)
      lcPath = Sys(5) + Sys(2003)
      tcName = Addbs(lcPath) + lcWBName + "." + Justext(tcName)
    Endif
    If Seek(Padr(Upper(tcName), 240, "_"), "xl_workbooks", "fullname")
      lnWB = 0
    Else
      Insert Into xl_workbooks (wbname, filepath, FullName) Values (lcWBName, lcPath, Padr(Upper(tcName), 240, "_"))
      lnWB = xl_workbooks.workbook
    Endif
    Return lnWB

*-- Creates the working cursors for worksheet definitions
  Protected Procedure CreateWorkingCursors
    If Used('xl_strings')
      Return
    Endif

    Create Cursor xl_strings Codepage = (This.Codepage) (Id I, workbook I, checksum C(230), stringxml M, stringval M, presvspace L, Formatted L)

    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+checksum Tag StringNdx
    Index On workbook Tag workbook

    Create Cursor xl_strformat (workbook I, Id I, Index I, stringxml M, stringval M, fbold L, fitalic L, fcolor I, fname C(100), Fsize N(5,1), uline C(16), strkthr L, ;
      fvpos C(11), Theme I, tint F(20,15), indexed I, presvspace L)
    Index On BinToC(workbook)+BinToC(Id)+Padl(Index, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(Id)+BinToC(Index) Tag Id For !Deleted()

    Create Cursor xl_workbooks (workbook I Autoinc Nextvalue 1, wbname C(200), filepath C(200), FullName C(240))
    Index On BinToC(workbook) Tag workbook
    Index On Upper(FullName) Tag FullName
    Index On Padr(Upper(Alltrim(wbname)), 200, "_") Tag wbname

    Create Cursor xl_sheets (workbook I, Sheet I, shname C(30), state I, mleft N(6,3), mright N(6,3), mtop N(6,3), mbot N(6,3), mheader N(6,3), ;
      mfooter N(6,3), shdeleted L, xsplit I, ysplit I, prnorient I, Papersize I, Paperwidth I, paperheight I, ;
      paperdimen C(2), Scale I, fittowidth I, fittoheight I, tabcolorndx I, tabcolorrgb C(8))
    Index On Sheet Tag Sheet
    Index On workbook Tag workbook
    Index On BinToC(workbook)+Upper(Alltrim(shname)) Tag shname
    Index On BinToC(workbook)+BinToC(Sheet) Tag sheetndx For !Deleted()

    Create Cursor xl_cells (workbook I, Sheet I, cellrow I, cellcol I, cellvalue C(25), DataType C(1), cellformula M, stringid I, cellxfs I, numdec I, celldeleted L, validndx I)
    Index On Sheet Tag Sheet
    Index On cellrow Tag cellrow
    Index On cellcol Tag cellcol
    Index On workbook Tag workbook
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(cellrow)+BinToC(cellcol) Tag reverse Descending
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(cellrow)+BinToC(cellcol) Tag cellindex Candidate

    Create Cursor xl_cellxfs (workbook I, Id I, numfmtid I, fontid I, fillid I, borderid I, halign C(8), valign C(8), indent I, wraptext I, Rotation I)
    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(numfmtid)+BinToC(fontid)+BinToC(fillid)+BinToC(borderid)+BinToC(indent)+BinToC(wraptext)+BinToC(Rotation)+Alltrim(halign)+Alltrim(valign) Tag cellxfs For !Deleted()
    Index On BinToC(workbook)+BinToC(fontid)+BinToC(fillid)+BinToC(borderid)+BinToC(indent)+BinToC(wraptext)+BinToC(Rotation)+Alltrim(halign)+Alltrim(valign) Tag cellformat For !Deleted()

    Create Cursor xl_fonts (workbook I, Id I, fname C(100), Fsize N(5,1), fbold L, fitalic L, fcolor I, uline C(16), strkthr L, ;
      fvpos C(11), Theme I, tint F(20,15), indexed I)
    Index On BinToC(workbook)+fname + Str(Fsize, 5, 1) + Transform(fbold) + Transform(fitalic) + Padl(fcolor, 15) + uline + Transform(strkthr) + fvpos Tag cellformat
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()

    Create Cursor xl_fills (workbook I, Id I, fgcolor I, bgcolor I, patttype C(25), Theme I, tint F(20,15), fgindexed I, bgindexed I)
    Index On BinToC(workbook)+BinToC(fgcolor)+BinToC(bgcolor)+patttype Tag FillColor
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()

    Create Cursor xl_borders (workbook I, Id I, lstyle C(20), lcolor I, ltheme I, ltint I, lindexed I, rstyle C(20), rcolor I, rtheme I, rtint I, rindexed I, ;
      tstyle C(20), tcolor I, ttheme I, ttint I, tindexed I, bstyle C(20), bcolor I, btheme I, btint I, bindexed I, ;
      dstyle C(20), dcolor I, dtheme I, dtint I, dindexed I, diagdn I, diagup I, Theme I, tint F(20,15))
    Index On BinToC(workbook) + Alltrim(lstyle) + BinToC(lcolor) + Alltrim(rstyle) + BinToC(rcolor) + Alltrim(tstyle) + BinToC(tcolor) + Alltrim(bstyle) + BinToC(bcolor) + ;
      ALLTRIM(dstyle) + BinToC(dcolor) + BinToC(diagdn) + BinToC(diagup) Tag cellborder
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()

    Create Cursor xl_mergecells (workbook I, Sheet I, begrow I, begcol I, endrow I, endcol I)
    Index On workbook Tag workbook
    Index On Sheet Tag Sheet
    Index On begrow Tag begrow
    Index On begcol Tag begcol
    Index On endrow Tag endrow
    Index On endcol Tag endcol
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(begrow)+BinToC(begcol) Tag mergecell For !Deleted()

    Create Cursor xl_colwidths (workbook I, Sheet I, Column I, Width N(16,9), bestfit L)
    Index On Sheet Tag Sheet
    Index On Column Tag Column
    Index On workbook Tag workbook
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(Column) Tag shcolumn For !Deleted()

    Create Cursor xl_hdrfooterdefn (workbook I, Sheet I, alignmargin L, difffirstpg L, diffoddeven L, scalewdoc L)
    Index On workbook Tag workbook
    Index On Sheet Tag Sheet
    Index On BinToC(workbook)+BinToC(Sheet) Tag wbsheet For !Deleted()

    Create Cursor xl_hdrfootertext (workbook I, Sheet I, Page I, section I, Text C(200), FontName C(50), FontSize I, fontstyle I, fontcolor I)
    Index On workbook Tag workbook
    Index On Sheet Tag Sheet
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(Page)+BinToC(section) Tag hdrfootr For !Deleted()

    Create Cursor xl_rowheights (workbook I, Sheet I, Row I, Height N(14,7))
    Index On workbook Tag workbook
    Index On Sheet Tag Sheet
    Index On Row Tag Row
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(Row) Tag shrow For !Deleted()

    Create Cursor xl_validation (workbook I, Sheet I, validndx I Autoinc Nextvalue 1, vtype I, vstyle I, voperator I, errmsg C(100), errtitle C(100), ;
      allowblank L, showinpmsg L, showerrmsg L, vprompt C(100), formula L, formula1 C(254), formula2 C(254))
    Index On BinToC(validndx) Tag validndx For !Deleted()
    Index On BinToC(workbook)+BinToC(Sheet) Tag wbshindex For !Deleted()
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(validndx) Tag validindex For !Deleted()
    Index On Left(formula1, 240) Tag formula1 For !Deleted()
    Index On Left(formula2, 240) Tag formula2 For !Deleted()

    Create Cursor xl_namerange (workbook I, rname C(50), scope I, Comment C(254), Sheet I, begrow I, begcol I, endrow I, endcol I)
    Index On BinToC(workbook)+rname Tag wbrname For !Deleted()

    Create Cursor xl_numFmts (workbook I, Id I, tempid I, formatxml M, formatcode M, applydec L)
    Index On BinToC(workbook)+Padl(Id, 10, '0') Tag sortorder For !Deleted()
    Index On BinToC(workbook)+BinToC(Id) Tag Id For !Deleted()

    Create Cursor xl_rows (workbook I, Sheet I, Row I, maxcol I)
    Index On BinToC(workbook)+BinToC(Sheet)+BinToC(Row) Tag Id For !Deleted()

    Create Cursor xl_ndxcolors (workbook I, indexid I, rgbcolor C(8))
    Index On BinToC(workbook)+rgbcolor Tag rgbcolor For !Deleted()
    Index On BinToC(workbook)+BinToC(indexid) Tag indexid For !Deleted()

    Create Cursor xl_mrucolors (workbook I, indexid I, rgbcolor C(8))
    Index On BinToC(workbook)+rgbcolor Tag rgbcolor For !Deleted()
    Index On BinToC(workbook)+BinToC(indexid) Tag indexid For !Deleted()

    Create Cursor xl_relationships (workbook I, Sheet I, relid I, reltype C(20), Target C(240))
    Index On BinToC(workbook)+BinToC(relid) Tag relid For !Deleted()
    Index On BinToC(workbook)+BinToC(Sheet) Tag Sheet For !Deleted()
    Index On BinToC(workbook)+reltype Tag reltype For !Deleted()

    Create Cursor xl_extsheets (workbook I, extid I, relid I, sheetname C(30))
    Index On BinToC(workbook)+BinToC(extid) Tag extid For !Deleted()
    Index On BinToC(workbook)+BinToC(extid)+BinToC(relid) Tag relid For !Deleted()

    Create Cursor xl_extcells (workbook I, extid I, cellrow I, cellcol I, cellvalue M)
    Index On BinToC(workbook)+BinToC(extid) Tag extid For !Deleted()
*INDEX ON BINTOC(workbook)+BINTOC(extid)+BINTOC(relid) TAG relid FOR !DELETED()

*-- Creates the working directories for the XML files
  Protected Procedure CreateWorkingDirectories
    Local lcDir, loException
    lcDir = Addbs(Sys(2023)) + Sys(2015)
    Try
      Mkdir (lcDir)
      Mkdir (lcDir + "\_rels")
      Mkdir (lcDir + "\docProps")
      Mkdir (lcDir + "\xl")
      Mkdir (lcDir + "\xl\_rels")
      Mkdir (lcDir + "\xl\worksheets")

    Catch To loException
      lcDir = Null
    Endtry
    Return Addbs(lcDir)

*-- Debugging use for saving the internal cursors to a text file
  Procedure DebugCursorsToFile
  Lparameters tcStemName
  Local lnNumFlds, lcField, lcLineText, lnFld, lcFileName
  Local Array laFields[1, 18]
  Select xl_workbooks
  Set Order To Tag workbook
  lnNumFlds = Afields(laFields, "xl_workbooks")
  lcFileText = "Cursor: xl_workbooks" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_workbooks.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_workbooks.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_sheets
  Set Order To Tag sheetndx
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_sheets")
  lcFileText = lcFileText + "Cursor: xl_sheets" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_sheets.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_sheets.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_cells
  Set Order To Tag cellindex
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_cells")
  lcFileText = lcFileText + "Cursor: xl_cells" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_cells.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_cells.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_strings
  Set Order To Tag stringid
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_strings")
  lcFileText = lcFileText + "Cursor: xl_strings" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_strings.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_strings.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_mergecells
  Set Order To Tag mergecell
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_mergecells")
  lcFileText = lcFileText + "Cursor: xl_mergecells" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_mergecells.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_mergecells.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_colwidths
  Set Order To Tag shcolumn
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_colwidths")
  lcFileText = lcFileText + "Cursor: xl_colwidths" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_colwidths.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_colwidths.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_fonts
  Set Order To Tag Id
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_fonts")
  lcFileText = lcFileText + "Cursor: xl_fonts" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_fonts.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_fonts.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_fills
  Set Order To Tag Id
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_fills")
  lcFileText = lcFileText + "Cursor: xl_fills" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_fills.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_fills.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_numFmts
  Set Order To Tag Id
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_numfmts")
  lcFileText = lcFileText + "Cursor: xl_numfmts" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_numFmts.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_numFmts.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_borders
  Set Order To Tag Id
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_borders")
  lcFileText = lcFileText + "Cursor: xl_borders" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_borders.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_borders.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_rowheights
  Set Order To Tag shrow
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_rowheights")
  lcFileText = lcFileText + "Cursor: xl_rowheights" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_rowheights.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_rowheights.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_hdrfooterdefn
  Set Order To Tag wbsheet
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_hdrfooterdefn")
  lcFileText = lcFileText + "Cursor: xl_hdrfooterdefn" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_hdrfooterdefn.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_hdrfooterdefn.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  Select xl_hdrfootertext
  Set Order To Tag hdrfootr
  lcFileName = tcStemName + " " + Alias() + ".txt"
  List To File (lcFileName) Noconsole
  lnNumFlds = Afields(laFields, "xl_hdrfootertext")
  lcFileText = lcFileText + "Cursor: xl_hdrfootertext" + CRLF
  Scan
    lcField = Lower(laFields[1, 1])
    lcLineText = lcField + ": " + Transform(xl_hdrfootertext.&lcField)
    For lnFld=2 To lnNumFlds
      lcField = Lower(laFields[lnFld, 1])
      lcLineText = lcLineText + ";  " + lcField + ": " + Transform(xl_hdrfootertext.&lcField)
    Endfor
    lcFileText = lcFileText + lcLineText + CRLF
  Endscan
  lcFileText = lcFileText + CRLF + CRLF + CRLF
  Set Order To

  lcFileName = tcStemName + " Cursors.txt"
  Strtofile(lcFileText, lcFileName)

*-- Limited functions for debugging
  Procedure DebugOutput
  Local lnWB, lnSh
  With This
    lnWB = .CreateWorkbook("ExcelDebug.xlsx")

    lnSh = .AddSheet(lnWB, "Debug Output")

    .SetCellValue(lnWB, lnSh, 1, 1, "Test string 1")
    .SetCellValue(lnWB, lnSh, 2, 1, "Test string 2")
    .SetCellValue(lnWB, lnSh, 3, 1, "Test string 3")
    .SetCellValue(lnWB, lnSh, 4, 1, "Test string 4")
    .SetCellValue(lnWB, lnSh, 5, 1, "Test string 5")

    .SetCellValue(lnWB, lnSh, 8, 1, "This is a long Test string for word-wrapping to occur This is a long Test string for word-wrapping to occur")
    .SetCellWordWrap(lnWB, lnSh, 8, 1, True)
    .SetRowHeight(lnWB, lnSh, 8, 75)
    .SetColumnWidth(lnWB, lnSh, 1, 25)

    .SetCellValue(lnWB, lnSh, 10, 1, Date(2015, 5, 6))
    .SetCellValue(lnWB, lnSh, 11, 1, Datetime(2015, 5, 6, 6, 30, 30))

    .SetCellFont(lnWB, lnSh, 1, 1, "Arial",                14,   True,         , Rgb(0,0,255))
    .SetCellFont(lnWB, lnSh, 2, 1, "Times New Roman",        ,       ,         , Rgb(0,0,255))
    .SetCellFont(lnWB, lnSh, 3, 1, "Consolas"       ,      14,   True)
    .SetCellFont(lnWB, lnSh, 4, 1, "Arial",                24,   True,     True, Rgb(0,0,255))
    .SetCellFont(lnWB, lnSh, 5, 1, "Cooper Black"   ,      12,       ,         ,             )

    .SetCellBorder(lnWB, lnSh, 2, 1, BORDER_TOP+BORDER_BOTTOM, BORDER_STYLE_THICK,  Rgb(100,150,200))
    .SetCellBorder(lnWB, lnSh, 5, 1, BORDER_LEFT+BORDER_RIGHT, BORDER_STYLE_DOUBLE, Rgb(255,0,0))

    .SetCellFill(lnWB, lnSh, 1, 1, Rgb(255,204,153))
    .SetCellFill(lnWB, lnSh, 3, 1, Rgb(204,255,255))

    .SaveWorkbook(lnWB)
  Endwith

  Procedure Debug_Assign
  Lparameters tlDebug
  If tlDebug
    Set Debugout To "VFPXWorkbookXLSXDebug.txt"
  Endif
  This.Debug = tlDebug

*-- Declares the required support Win32 API
  Protected Procedure DeclareDLL
    Declare Sleep In kernel32 As apiSleep Integer tnTime
    Declare Integer DeleteFile In kernel32 As apiDeleteFile String lpFileName
    Declare Integer RemoveDirectory In kernel32 As apiRemoveDirectory String lpPathName
    Declare Long MoveFile In Win32API As apiMoveFile String SourceFileName, String DestFileName
    Declare Integer ShellExecute In Shell32 As apiShellExecute Long nhWnd, String lpAction, String lpFile, String lpParams, String lpDirectory, Integer nShowCmd

*-- Deletes all workbooks
  Procedure DeleteAllWorkbooks
  Local loException
  Try
    With This
      .DestroyWorkingCursors()
      .CreateWorkingCursors()
      .InitalizeIdValues()
    Endw

  Catch To loException
    Set Step On
  Endtry

*-- Deletes the sheet from the Workbook
  Procedure DeleteSheet
  Lparameters tnWB, txSheet
  Local lnSheet, lxReturn
  lxReturn = False
  Do Case
  Case Vartype(txSheet) = "C"
    If Seek(BinToC(tnWB)+Upper(txSheet), "xl_sheets", "shname") And Seek(BinToC(tnWB)+BinToC(xl_sheets.Sheet), "xl_relationships", "sheet")
      lnSheet = xl_sheets.Sheet
      Replace xl_sheets.shdeleted With True In xl_sheets
      Delete In xl_sheets
      Delete In xl_relationships
      lxReturn = True
    Endif

  Case Vartype(txSheet) = "N"
    If Seek(BinToC(tnWB)+BinToC(txSheet), "xl_sheets", "sheetndx") And Seek(BinToC(tnWB)+BinToC(xl_sheets.Sheet), "xl_relationships", "sheet")
      lnSheet = txSheet
      Replace xl_sheets.shdeleted With True In xl_sheets
      Delete In xl_sheets
      Delete In xl_relationships
      lxReturn = True
    Endif
  Endcase
  If lxReturn
    Replace xl_cells.celldeleted With True For workbook = tnWB And Sheet = lnSheet In xl_cells
    Delete For workbook = tnWB And Sheet = lnSheet In xl_cells
  Endif
  Return lxReturn

*-- Deletes the selected workbook
  Procedure DeleteWorkbook
  Lparameters tnWB
  Local llSuccess, lnNdx
  Try
    Replace xl_sheets.shdeleted  With True For workbook = tnWB In xl_sheets
    Replace xl_cells.celldeleted With True For workbook = tnWB In xl_cells
    Delete For workbook = tnWB In xl_workbooks
    Delete For workbook = tnWB In xl_sheets
    Delete For workbook = tnWB In xl_cells
    Delete For workbook = tnWB In xl_strformat
    Delete For workbook = tnWB In xl_strings
    Delete For workbook = tnWB In xl_cellxfs
    Delete For workbook = tnWB In xl_fonts
    Delete For workbook = tnWB In xl_fills
    Delete For workbook = tnWB In xl_borders
    Delete For workbook = tnWB In xl_mergecells
    Delete For workbook = tnWB In xl_colwidths
    Delete For workbook = tnWB In xl_hdrfooterdefn
    Delete For workbook = tnWB In xl_hdrfootertext
    Delete For workbook = tnWB In xl_rowheights
    Delete For workbook = tnWB In xl_validation
    Delete For workbook = tnWB In xl_namerange
    Delete For workbook = tnWB In xl_numFmts
    Delete For workbook = tnWB In xl_rows
    Delete For workbook = tnWB In xl_ndxcolors
    Delete For workbook = tnWB In xl_mrucolors
    Delete For workbook = tnWB In xl_relationships
    Delete For workbook = tnWB In xl_extsheets
    Delete For workbook = tnWB In xl_extcells

    With This
      lnNdx = Ascan(.LastBorderId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastBorderId, lnNdx)
        lnNdx = Alen(.LastBorderId, 1) - 1
        If lnNdx > 0
          Dimension .LastBorderId[lnNdx, 2]
        Else
          .LastBorderId[1, 1] = 0                                          && Workbook Id
          .LastBorderId[1, 2] = 0                                          && Last Border Id
        Endif
      Endif

      lnNdx = Ascan(.LastCellXfsId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastCellXfsId, lnNdx)
        lnNdx = Alen(.LastCellXfsId, 1) - 1
        If lnNdx > 0
          Dimension .LastCellXfsId[lnNdx, 2]
        Else
          .LastCellXfsId[1, 1] = 0                                         && Workbook Id
          .LastCellXfsId[1, 2] = 0                                         && Last CellXfs Id
        Endif
      Endif

      lnNdx = Ascan(.LastFillId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastFillId, lnNdx)
        lnNdx = Alen(.LastFillId, 1) - 1
        If lnNdx > 0
          Dimension .LastFillId[lnNdx, 2]
        Else
          .LastFillId[1, 1] = 0                                            && Workbook Id
          .LastFillId[1, 2] = 0                                            && Last Fill Id
        Endif
      Endif

      lnNdx = Ascan(.LastFontId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastFontId, lnNdx)
        lnNdx = Alen(.LastFontId, 1) - 1
        If lnNdx > 0
          Dimension .LastFontId[lnNdx, 2]
        Else
          .LastFontId[1, 1] = 0                                            && Workbook Id
          .LastFontId[1, 2] = 0                                            && Last Font Id
        Endif
      Endif

      lnNdx = Ascan(.LastIndexColorId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastIndexColorId, lnNdx)
        lnNdx = Alen(.LastIndexColorId, 1) - 1
        If lnNdx > 0
          Dimension .LastIndexColorId[lnNdx, 2]
        Else
          .LastIndexColorId[1, 1] = 0                                      && Workbook Id
          .LastIndexColorId[1, 2] = 0                                      && Last IndexColor Id
        Endif
      Endif

      lnNdx = Ascan(.LastMruColorId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastMruColorId, lnNdx)
        lnNdx = Alen(.LastMruColorId, 1) - 1
        If lnNdx > 0
          Dimension .LastMruColorId[lnNdx, 2]
        Else
          .LastMruColorId[1, 1] = 0                                        && Workbook Id
          .LastMruColorId[1, 2] = 0                                        && Last mruColor Id
        Endif
      Endif

      lnNdx = Ascan(.LastNumFmtId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastNumFmtId, lnNdx)
        lnNdx = Alen(.LastNumFmtId, 1) - 1
        If lnNdx > 0
          Dimension Ths.LastNumFmtId[lnNdx, 2]
        Else
          .LastNumFmtId[1, 1] = 0                                          && Workbook Id
          .LastNumFmtId[1, 2] = 0                                          && Last NumFmt Id
        Endif
      Endif

      lnNdx = Ascan(.LastRelationId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastRelationId, lnNdx)
        lnNdx = Alen(.LastRelationId, 1) - 1
        If lnNdx > 0
          Dimension .LastRelationId[lnNdx, 2]
        Else
          .LastRelationId[1, 1] = 0                                        && Workbook Id
          .LastRelationId[1, 2] = 0                                        && Last Relationship Id
        Endif
      Endif

      lnNdx = Ascan(.LastSheetId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastSheetId, lnNdx)
        lnNdx = Alen(.LastSheetId, 1) - 1
        If lnNdx > 0
          Dimension .LastSheetId[lnNdx, 2]
        Else
          .LastSheetId[1, 1] = 0                                           && Workbook Id
          .LastSheetId[1, 2] = 0                                           && Last Sheet Id
        Endif
      Endif

      lnNdx = Ascan(.LastStringId, tnWB, 1, -1, 1, 15)
      If lnNdx > 0
        Adel(.LastStringId, lnNdx)
        lnNdx = Alen(.LastStringId, 1) - 1
        If lnNdx > 0
          Dimension .LastStringId[lnNdx, 2]
        Else
          .LastStringId[1, 1] = 0                                          && Workbook Id
          .LastStringId[1, 2] = 0                                          && Last String Id
        Endif
      Endif
      llSuccess = True
    Endwith

  Catch To loException
    Set Step On
    llSuccess = False
  Endtry
  Return llSuccess

*-- Demo for the features of this class
  Procedure Demo
  Local lnWB, lnSh1, lnRow, lnCol, lnSh2, lnSh3, lnSh4, lnCnt, lnFormat, llBoolean, lcNamedRange
  Local lnSh6, lnSh5, loInline, lnStyle1, lnStyle2, lnBorders, lcString
  lnWB = This.CreateWorkbook("ExcelTest.xlsx")
  If lnWB = 0
    Return False
  Endif

  With This
    lnSh1 = .AddSheet(lnWB, "Test Sheet 1")
    If lnSh1 = 0
      Return False
    Endif
    .FreezePanes(lnWB, lnSh1, 1, 2)

    llBoolean = True
    For lnRow=1 To 10
      For lnCol=1 To 9
        .SetCellValue(lnWB, lnSh1, lnRow, lnCol, Sys(2015))
      Endfor
      llBoolean = !llBoolean
      .SetCellValue(lnWB, lnSh1, lnRow, 10, llBoolean)
    Endfor
    .SetRowHeight(lnWB, lnSh1, 6, 25)

*-*	Example of cell styles
    lnStyle1  = .CreateFormatStyle(lnWB)   && Create the base style definition
    lnBorders = BORDER_LEFT + BORDER_RIGHT + BORDER_TOP + BORDER_BOTTOM
    .AddStyleBorders(lnWB, lnStyle1, lnBorders, BORDER_STYLE_THIN, Rgb(16,100,200))
    .AddStyleFont(lnWB, lnStyle1, "Times New Roman", 14, False, False, Rgb(0,0,255))
    .SetCellStyleRange(lnWB, lnSh1, 2, 1, 2, 9, lnStyle1)

    lnStyle2  = .CreateFormatStyle(lnWB)   && Create the base style definition
    lnBorders = BORDER_BOTTOM
    .AddStyleBorders(lnWB, lnStyle2, lnBorders, BORDER_STYLE_THICK, Rgb(16,100,200))
    .AddStyleFont(lnWB, lnStyle2, "Arial", 16, True, False)
    .SetCellStyleRange(lnWB, lnSh1, 4, 1, 4, 9, lnStyle2)

*-*	* Indentation examples  [OLD WAY - DEPRICATED METHODS]
*-*	this.SetCellIndent(lnWb, lnSh1, 2, 6, 3)
*-*	this.SetCellIndent(lnWb, lnSh1, 4, 6, 3)
*-*	this.SetCellIndent(lnWb, lnSh1, 6, 6, 3)
*-*	this.SetCellIndent(lnWb, lnSh1, 8, 6, 3)
*-*	this.SetCellIndent(lnWb, lnSh1, 10, 6, 3)
*-*
*-*	*                tnWB, tnSheet, tnCellRow, tnCellCol,           tcFName, tnFSize, tlBold, tlItalic,     tnFColor,  tcULine, tlStrikThr, tcVPos
*-*	this.SetCellFont(lnWb, lnSh1,           1,         1, "Arial",                14,   True,     True, RGB(0,0,255))
*-*	this.SetCellFont(lnWb, lnSh1,           2,         1, "Times New Roman",        ,       ,         , RGB(0,0,255))
*-*	this.SetCellFont(lnWb, lnSh1,           3,         1,                  ,      14,   True)
*-*	this.SetCellFont(lnWb, lnSh1,           4,         1, "Arial",                14,   True,     True, RGB(0,0,255))
*-*	this.SetCellFont(lnWb, lnSh1,           6,         1,                  ,        ,       ,         ,             , 'single')
*-*	this.SetCellFont(lnWb, lnSh1,           7,         1,                  ,        ,       ,         ,             , 'double')
*-*	this.SetCellFont(lnWb, lnSh1,           8,         1,                  ,        ,       ,         ,             ,         , True)
*-*	this.SetCellFont(lnWb, lnSh1,           9,         1,                  ,        ,       ,         ,             ,         ,            , "superscript")
*-*
*-*	* New procedure each line is drawed with the same style and color
*-*	this.SetCellBorder(lnWb, lnSh1, 3, 4, 63, BORDER_STYLE_THIN,   RGB(16,100,200))
*-*	this.SetCellBorder(lnWb, lnSh1, 3, 6, 31, BORDER_STYLE_THICK,  RGB(100,150,200))
*-*	this.SetCellBorder(lnWb, lnSh1, 3, 8, 47, BORDER_STYLE_DOUBLE, RGB(200,150,100))
*-*	this.SetCellBorder(lnWb, lnSh1, 3, 4,  8, BORDER_STYLE_THIN,   RGB(16,100,200))
*-*
*-*	this.SetCellWordWrapRange(lnWb, lnSh1, 1, 1, 10, 9, True)
*-*
*-*	* Old procedure - each line can have a different style or color
*-*	this.SetCellBorderEx(lnWb, lnSh1, 5, 2, 'thin',,'thin',,'thick',,'thick')
*-*	this.SetCellBorderEx(lnWb, lnSh1, 9, 2, ,,,,'double',255,'thick',RGB(255,0,255))

* Example of setting inline text formatting
    loInline = .CreateInlineFormatText(lnWB, "This is an in-line formatted text string")
    loInline = .AddInlineFontObject(loInline, 1, 4, "Arial Black", 16)
    loInline = .AddInlineFontObject(loInline, 5, 15, "Times New Roman", 14, Rgb(255,0,0), .T.)
    loInline = .AddInlineFontObject(loInline, 20, 21, "Arial", 12)
    .SetCellInLineFormatText(lnWB, lnSh1, 16, 1, loInline)

* Rotated text
    .SetCellValue(lnWB, lnSh1, 17, 1, "This text is rotated 45 degrees")
    .SetCellTextRotation(lnWB, lnSh1, 17, 1, 45)
    .SetCellValue(lnWB, lnSh1, 18, 1, "This text is rotated -45 degrees")
    .SetCellTextRotation(lnWB, lnSh1, 18, 1, -45)

* Similar text
    .SetCellValue(lnWB, lnSh1, 15, 7, "This text is similar")
    .SetCellValue(lnWB, lnSh1, 16, 7, "This text is similar to the one above but is longer length")

*-* Very long text
    lcString = "Addtionally, there is a new method for reading an existing XLSX file into the working cursors and then supporting methods to extract the individual cell values and formatting.  However, this XLSX read method is still in development "
    .SetCellValue(lnWB, lnSh1, 22, 1, lcString)
    lcString = lcString + "and not yet ready for use.  This is the method OpenXlsxWorkbook() - only use if you want to try it but the method will not necessarily read all the numeric"
    .SetCellValue(lnWB, lnSh1, 23, 1, lcString)
    lcString = lcString + " or date values correctly at this point.  I did not want to hold this release of VFPxWorkbookXlsx class for the completion of the reading of the XLSX file due to the other bug fixes."
    .SetCellValue(lnWB, lnSh1, 24, 1, lcString)

* Text with ASCII > 128
    .SetCellValue(lnWB, lnSh1, 19, 1, "This text that has ASCII chars > 128:" + CRLF + "Trademark symbol: " + Chr(174) + CRLF + "Copyright symbol: " + Chr(169))

* Column Width examples
    .SetColumnWidth(lnWB, lnSh1, 1, 25)
    .SetColumnWidth(lnWB, lnSh1, 2, 15)
    .SetColumnWidth(lnWB, lnSh1, 3, 15)
    .SetColumnWidth(lnWB, lnSh1, 4, 20)
    .SetColumnWidth(lnWB, lnSh1, 5, 15)
    .SetColumnWidth(lnWB, lnSh1, 6, 25)
    .SetColumnWidth(lnWB, lnSh1, 9, 15)

*Row inserted
    .InsertRow(lnWB, lnSh1, 5, INSERT_AFTER)

*Column inserted
    .InsertColumn(lnWB, lnSh1, 5, INSERT_RIGHT)

* Sheet setup examples
    .SetPrintOrientation(lnWB, lnSh1, PORTRAIT_PRINT_ORIENTATION)
    .SetPaperSize(lnWB, lnSh1, PAPERSIZE_LEGAL)
    .SetSheetScale(lnWB, lnSh1, 50)

*******************************************************************************

    lnSh2 = .AddSheet(lnWB, "Test Sheet 2")
    If lnSh2 = 0
      Return False
    Endif
    .FreezePanes(lnWB, lnSh2, 1)

    .SetCellValue(lnWB, lnSh2, 1, 1, "Column Header 1")
    .SetCellValue(lnWB, lnSh2, 1, 2, "Column Header 2")
    .SetCellValue(lnWB, lnSh2, 1, 3, "Column Header 3")
    .SetCellValue(lnWB, lnSh2, 1, 4, "Column Header 4")
    .SetCellValue(lnWB, lnSh2, 1, 5, "Column Header 5")

    .SetCellFont(lnWB, lnSh2, 1, 1, "Arial", 14, True, False, Rgb(255,255,255))
    .SetCellFill(lnWB, lnSh2, 1, 1, Rgb(51,102,255))
    .CellFormatPainter(lnWB, lnSh2, 1, 1, 1, 2, 1, 5)

    .SetColumnWidthRange(lnWB, lnSh2, 1, 5, 24)

* Add integer values to cells
    For lnRow=2 To 6
      For lnCol=1 To 5
        .SetCellValue(lnWB, lnSh2, lnRow, lnCol, lnRow*lnCol)
      Endfor
    Endfor

* Add SUM() formula to columns
    For lnCol=1 To 5
      lcColumn = .ColumnIndexToAscii(lnCol)
      .SetCellFormula(lnWB, lnSh2, 7, lnCol, "=SUM(" + lcColumn + "2:" + lcColumn + "6)")
    Endfor

* Add floating point values to cells
    For lnRow=10 To 15
      For lnCol=3 To 5
        .SetCellValue(lnWB, lnSh2, lnRow, lnCol, lnRow*1.25)
      Endfor
    Endfor

* Add AVERAGE(), MAX(), & MIN() formula to columns
    .SetCellFormula(lnWB, lnSh2, 16, 3, "=AVERAGE(C10:C15)")
    .SetCellFormula(lnWB, lnSh2, 16, 4, "=MAX(D10:D15)")
    .SetCellFormula(lnWB, lnSh2, 16, 5, "=MIN(E10:E15)")

    .SetCellValue(lnWB, lnSh2, 16, 1, "This row contains calculated formulas -->")

* Merge some cells
    .MergeCells(lnWB, lnSh2, 2, 8, 2, 10)
    .MergeCells(lnWB, lnSh2, 3, 8, 3, 10)
    .MergeCells(lnWB, lnSh2, 5, 8, 10, 10)
    .SetCellValue(lnWB, lnSh2, 2, 8, "Merged H2:J2")
    .SetCellValue(lnWB, lnSh2, 3, 8, "Merged H3:J3")
    .SetCellValue(lnWB, lnSh2, 5, 8, "Merged H5:J10")
    .SetCustomPaperSize(lnWB, lnSh2, 841, 1189, "mm")
    .SetHeaderFooterSetup(lnWB, lnSh2, False, True, False, False)
    .SetHeaderFooterText(lnWB, lnSh2, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_HDR_CENTER, "This is First Page", "Arial", 16, HEADERFOOTER_FONT_STYLE_BOLD, Rgb(255,0,0))
    .SetHeaderFooterText(lnWB, lnSh2, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_HDR_CENTER, "These are Additional Pages", "Arial", 10, HEADERFOOTER_FONT_STYLE_BOLD, Rgb(0,128,0))
    .SetHeaderFooterText(lnWB, lnSh2, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_FTR_CENTER, "Page &amp;P of &amp;N", "Arial", 16, HEADERFOOTER_FONT_STYLE_BOLD)

*******************************************************************************

    lnSh3 = .AddSheet(lnWB, "Number Formats")
    If lnSh3 = 0
      Return False
    Endif

    lnFormat = .AddNumericFormat(lnWB, '[Green]"$"#,##0.00;[Red]"$"#,##0.00')
*lnFormat = this.AddCustomNumericFormat(lnWb, '[Green]"$"#,##0.00', '[Red]"$"#,##0.00')
    .SetCellValue(lnWB, lnSh3, 1, 1, 56.558)
    .SetCellValue(lnWB, lnSh3, 1, 2, "Custom defined format: " + .GetCustomNumericFormat(lnWB, lnFormat))
    .SetCellNumberFormat(lnWB, lnSh3, 1, 1, lnFormat)

    .SetColumnWidth(lnWB, lnSh3, 1, 18)
    .SetColumnWidth(lnWB, lnSh3, 2, 60)

    .SetCellValue(lnWB, lnSh3, 3, 1, 14)
    .SetCellValue(lnWB, lnSh3, 3, 2, "Integer format")
    .SetCellNumberFormat(lnWB, lnSh3, 3, 1, CELL_FORMAT_INTEGER)

    .SetCellValue(lnWB, lnSh3, 4, 1, 13355.2)
    .SetCellValue(lnWB, lnSh3, 4, 2, "Float format")
    .SetCellNumberFormat(lnWB, lnSh3, 4, 1, CELL_FORMAT_FLOAT)

    .SetCellValue(lnWB, lnSh3, 5, 1, 16898)
    .SetCellValue(lnWB, lnSh3, 5, 2, "Integer format, comma separated")
    .SetCellNumberFormat(lnWB, lnSh3, 5, 1, CELL_FORMAT_COMMA_INTEGER)

    .SetCellValue(lnWB, lnSh3, 6, 1, 13355.20)
    .SetCellValue(lnWB, lnSh3, 6, 2, "Float format, comma separated")
    .SetCellNumberFormat(lnWB, lnSh3, 6 , 1, CELL_FORMAT_COMMA_FLOAT)

    .SetCellValue(lnWB, lnSh3, 7, 1, -13355.20)
    .SetCellValue(lnWB, lnSh3, 7, 2, "Currency format, negative in parenthesis")
    .SetCellNumberFormat(lnWB, lnSh3, 7, 1, CELL_FORMAT_CURRENCY_PAREN)

    .SetCellValue(lnWB, lnSh3, 8, 1, -13355.20)
    .SetCellValue(lnWB, lnSh3, 8, 2, "Currency format, negative red in parenthesis")
    .SetCellNumberFormat(lnWB, lnSh3, 8, 1, CELL_FORMAT_CURRENCY_RED_PAREN)

    .SetCellValue(lnWB, lnSh3, 10, 1, 0.55)
    .SetCellValue(lnWB, lnSh3, 10, 2, "Percent format, integer")
    .SetCellNumberFormat(lnWB, lnSh3, 10, 1, CELL_FORMAT_PERCENT_INTEGER)

    .SetCellValue(lnWB, lnSh3, 11, 1, 0.5534)
    .SetCellValue(lnWB, lnSh3, 11, 2, "Percent format, float")
    .SetCellNumberFormat(lnWB, lnSh3, 11, 1, CELL_FORMAT_PERCENT_FLOAT)

    .SetCellValue(lnWB, lnSh3, 13, 1, 456800000.00)
    .SetCellValue(lnWB, lnSh3, 13, 2, "Exponential format")
    .SetCellNumberFormat(lnWB, lnSh3, 13, 1, CELL_FORMAT_EXPONENT)

    .SetCellValue(lnWB, lnSh3, 15, 1, 0.25)
    .SetCellValue(lnWB, lnSh3, 15, 2, "Fraction format, single place")
    .SetCellNumberFormat(lnWB, lnSh3, 15, 1, CELL_FORMAT_FRACTION_1)

    .SetCellValue(lnWB, lnSh3, 16, 1, 0.3125)
    .SetCellValue(lnWB, lnSh3, 16, 2, "Fraction format, 2 places")
    .SetCellNumberFormat(lnWB, lnSh3, 16, 1, CELL_FORMAT_FRACTION_2)

    .SetCellValue(lnWB, lnSh3, 18, 1, Date())
    .SetCellValue(lnWB, lnSh3, 18, 2, "Default date format")
    .SetCellNumberFormat(lnWB, lnSh3, 18, 1, CELL_FORMAT_DATE_DMMMYY)

    .SetCellValue(lnWB, lnSh3, 19, 1, Date())
    .SetCellValue(lnWB, lnSh3, 19, 2, "Date format, d-mmm-yy")
    .SetCellNumberFormat(lnWB, lnSh3, 19, 1, CELL_FORMAT_DATE_DMMMYY)

    .SetCellValue(lnWB, lnSh3, 20, 1, Date())
    .SetCellValue(lnWB, lnSh3, 20, 2, "Date format, d-mmm")
    .SetCellNumberFormat(lnWB, lnSh3, 20, 1, CELL_FORMAT_DATE_DMMM)

    .SetCellValue(lnWB, lnSh3, 21, 1, Date())
    .SetCellValue(lnWB, lnSh3, 21, 2, "Date format, mmm-yy")
    .SetCellNumberFormat(lnWB, lnSh3, 21, 1, CELL_FORMAT_DATE_MMMYY)

    .SetCellValue(lnWB, lnSh3, 23, 1, 0.35733796296296294)
    .SetCellValue(lnWB, lnSh3, 23, 2, "Time format, h:mm AM/PM")
    .SetCellNumberFormat(lnWB, lnSh3, 23, 1, CELL_FORMAT_TIME_HMMAMPM)

    .SetCellValue(lnWB, lnSh3, 24, 1, 0.35733796296296294)
    .SetCellValue(lnWB, lnSh3, 24, 2, "Time format, h:mm:ss AM/PM")
    .SetCellNumberFormat(lnWB, lnSh3, 24, 1, CELL_FORMAT_TIME_HMMSSAMPM)

    .SetCellValue(lnWB, lnSh3, 25, 1, 0.35733796296296294)
    .SetCellValue(lnWB, lnSh3, 25, 2, "Time format, h:mm")
    .SetCellNumberFormat(lnWB, lnSh3, 25, 1, CELL_FORMAT_TIME_HMM)

    .SetCellValue(lnWB, lnSh3, 26, 1, 0.35733796296296294)
    .SetCellValue(lnWB, lnSh3, 26, 2, "Time format, h:mm:ss")
    .SetCellNumberFormat(lnWB, lnSh3, 26, 1, CELL_FORMAT_TIME_HMMSS)

    .SetCellValue(lnWB, lnSh3, 28, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 28, 2, "Date-Time format, m/d/yy h:mm")
    .SetCellNumberFormat(lnWB, lnSh3, 28, 1, CELL_FORMAT_DATETIME_MDYYHMM)

    .SetCellValue(lnWB, lnSh3, 29, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 29, 2, "Date-Time format, [$-409]dd/mmm/yyyy\ h:mm\ AM/PM;@")
    .SetCellNumberFormat(lnWB, lnSh3, 29, 1, CELL_FORMAT_DATETIME_DDMMMYYYY_TTAM)

    .SetCellValue(lnWB, lnSh3, 30, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 30, 2, "Date-Time format, dd/mmm/yyyy\ h:mm;@")
    .SetCellNumberFormat(lnWB, lnSh3, 30, 1, CELL_FORMAT_DATETIME_DDMMMYYYY_TT24)

    .SetCellValue(lnWB, lnSh3, 31, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 31, 2, "Date-Time format, [$-409]mmm\ d\,\ yyyy\ h:mm\ AM/PM;@")
    .SetCellNumberFormat(lnWB, lnSh3, 31, 1, CELL_FORMAT_DATETIME_MMMDDYYYY_TTAM)

    .SetCellValue(lnWB, lnSh3, 32, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 32, 2, "Date-Time format, [$-409]mmm\ d\,\ yyyy\ h:mm;@")
    .SetCellNumberFormat(lnWB, lnSh3, 32, 1, CELL_FORMAT_DATETIME_MMMDDYYYY_TT24)

    .SetCellValue(lnWB, lnSh3, 33, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 33, 2, "Date-Time format, m/d/yy\ h:mm\ AM/PM;@")
    .SetCellNumberFormat(lnWB, lnSh3, 33, 1, CELL_FORMAT_DATETIME_MDYY_TTAM)

    .SetCellValue(lnWB, lnSh3, 34, 1, Datetime())
    .SetCellValue(lnWB, lnSh3, 34, 2, "Date-Time format, m/d/yy\ h:mm;@")
    .SetCellNumberFormat(lnWB, lnSh3, 34, 1, CELL_FORMAT_DATETIME_MDYY_TT24)

    .SetCellValue(lnWB, lnSh3, 36, 1, -89235)
    .SetCellValue(lnWB, lnSh3, 36, 2, "Numeric format, #,##0;(#,##0)")
    .SetCellNumberFormat(lnWB, lnSh3, 36, 1, CELL_FORMAT_COMMA_INTEGER_PAREN)

    .SetCellValue(lnWB, lnSh3, 37, 1, -12457)
    .SetCellValue(lnWB, lnSh3, 37, 2, "Numeric format, #,##0;[Red](#,##0)")
    .SetCellNumberFormat(lnWB, lnSh3, 37, 1, CELL_FORMAT_COMMA_INTEGER_RED_PAREN)

    .SetCellValue(lnWB, lnSh3, 38, 1, -1245789.22)
    .SetCellValue(lnWB, lnSh3, 38, 2, "Numeric format, #,##0.00;(#,##0.00)")
    .SetCellNumberFormat(lnWB, lnSh3, 38, 1, CELL_FORMAT_COMMA_FLOAT_PAREN)

    .SetCellValue(lnWB, lnSh3, 39, 1, -125896.33)
    .SetCellValue(lnWB, lnSh3, 39, 2, "Numeric format, #,##0.00;[Red](#,##0.00)")
    .SetCellNumberFormat(lnWB, lnSh3, 39, 1, CELL_FORMAT_COMMA_FLOAT_RED_PAREN)

    .SetCellValue(lnWB, lnSh3, 41, 1, 6.25E-2)
    .SetCellValue(lnWB, lnSh3, 41, 2, "Time format")
    .SetCellNumberFormat(lnWB, lnSh3, 41, 1, CELL_FORMAT_TIME_MMSS)

    .SetCellValue(lnWB, lnSh3, 42, 1, 0.35733796296296294)
    .SetCellValue(lnWB, lnSh3, 42, 2, "Time format, [h]:mm:ss")
    .SetCellNumberFormat(lnWB, lnSh3, 42, 1, CELL_FORMAT_TIME_H_MMSS)

    .SetCellValue(lnWB, lnSh3, 44, 1, -45.55)
    .SetCellValue(lnWB, lnSh3, 44, 2, "CELL_FORMAT_CURRENCY_RED")
    .SetCellNumberFormat(lnWB, lnSh3, 44, 1, CELL_FORMAT_CURRENCY_RED)

    .SetCellValue(lnWB, lnSh3, 45, 1, -45.55)
    .SetCellValue(lnWB, lnSh3, 45, 2, "CELL_FORMAT_ACC_CURR_POUNDS")
    .SetCellNumberFormat(lnWB, lnSh3, 45, 1, CELL_FORMAT_ACC_CURR_POUNDS)

    .SetCellValue(lnWB, lnSh3, 46, 1, -45.55)
    .SetCellValue(lnWB, lnSh3, 46, 2, "CELL_FORMAT_ACC_CURR_EURO")
    .SetCellNumberFormat(lnWB, lnSh3, 46, 1, CELL_FORMAT_ACC_CURR_EURO)

    .SetCellValue(lnWB, lnSh3, 47, 1, -45.55)
    .SetCellValue(lnWB, lnSh3, 47, 2, "CELL_FORMAT_CURR_EURO_RED")
    .SetCellNumberFormat(lnWB, lnSh3, 47, 1, CELL_FORMAT_CURR_EURO_RED)

    .SetCellValue(lnWB, lnSh3, 48, 1, -45.55)
    .SetCellValue(lnWB, lnSh3, 48, 2, "CELL_FORMAT_CURR_POUNDS_RED")
    .SetCellNumberFormat(lnWB, lnSh3, 48, 1, CELL_FORMAT_CURR_POUNDS_RED)

    .SetPrintOrientation(lnWB, lnSh3, LANDSCAPE_PRINT_ORIENTATION)
*-*	this.SetPrintFitToHeight(lnWb, lnSh3, 2)

*******************************************************************************

    lnSh4 = .AddSheet(lnWB, "Color Sheet")
    If lnSh4 = 0
      Return False
    Endif

    .SetCellValue(lnWB, lnSh4, 1, 1, "RGB(51,102,255)")
    .SetCellFill(lnWB, lnSh4, 1, 1, Rgb(51,102,255))

    .SetCellValue(lnWB, lnSh4, 1, 2, "RGB(255,0,0)")
    .SetCellFill(lnWB, lnSh4, 1, 2, Rgb(255,0,0))

    .SetCellAlignmentRange(lnWB, lnSh4, 1, 1, 1, 2, CELL_HORIZ_ALIGN_CENTER, CELL_VERT_ALIGN_CENTER)

    .SetColumnWidthRange(lnWB, lnSh4, 1, 2, 18)
    .SetRowHeight(lnWB, lnSh4, 1, 60)

*******************************************************************************

    lnSh5 = .AddSheet(lnWB, "Blank Sheet")

*******************************************************************************

    lnSh6 = .AddSheet(lnWB, "List Values", HIDDEN_SHEET_STATE)
    If lnSh6 = 0
      Return False
    Endif

    .SetCellValidation(lnWB, lnSh1, 13, 1, LIST_VALID_TYPE, , , , , , , , , "Yes,No,Maybe")
    .SetCellValue(lnWB, lnSh1, 13, 2, "The cell to the left has a droplist validation")

    .SetCellValue(lnWB, lnSh6, 1, 1, "Named Value List")
    .SetCellValue(lnWB, lnSh6, 2, 1, "Named Value 1")
    .SetCellValue(lnWB, lnSh6, 3, 1, "Named Value 2")
    .SetCellValue(lnWB, lnSh6, 4, 1, "Named Value 3")
    .SetCellValue(lnWB, lnSh6, 5, 1, "Named Value 4")
    .SetCellValue(lnWB, lnSh6, 6, 1, "Named Value 5")
    .SetCellValue(lnWB, lnSh6, 7, 1, "Named Value 6")

    lcNamedRange = .AddNamedRange(lnWB, lnSh6, "Named Range", SCOPE_WB_NAMED_RANGE, "This is a comment", 2, 1, 7, 1)

    .SetCellValidation(lnWB, lnSh1, 14, 1, LIST_VALID_TYPE, , , , , , , , , "=" + lcNamedRange)
    .SetCellValue(lnWB, lnSh1, 14, 2, "The cell to the left has a droplist validation based on a Named Range (stored in Hidden Sheet6)")

    .InsertCell(lnWB, lnSh1,  9, 9, INSERT_RIGHT)
    .InsertCell(lnWB, lnSh1, 11, 9, INSERT_BEFORE)

*-*  For testing sheet saving only
*this.DeleteSheet(lnWb, lnSh2)

    .SaveWorkbook(lnWB)
    If .Debug
      .DebugCursorsToFile("ExcelTest")
    Endif
  Endwith

*-- Derives the sheet name
  Protected Procedure DeriveSheetName
    Lparameters txWB
    Do Case
    Case Vartype(txWB) = "C"
      lcSheetName = Juststem(txWB)

    Case Vartype(txWB) = "N"
      lcSheetName = Juststem(This.GetWorkbookFileName(txWB))

    Otherwise
      Return ""
    Endcase
    Return lcSheetName

*-- Destroys the working cursors
  Protected Procedure DestroyWorkingCursors
    Try
      Use In Select('xl_strings')
      Use In Select('xl_workbooks')
      Use In Select('xl_sheets')
      Use In Select('xl_cells')
      Use In Select('xl_fonts')
      Use In Select('xl_fills')
      Use In Select('xl_borders')
      Use In Select('xl_mergecells')
      Use In Select('xl_colwidths')
      Use In Select('xl_hdrfooterdefn')
      Use In Select('xl_hdrfootertext')
      Use In Select('xl_rowheights')
      Use In Select('xl_numfmts')
      Use In Select('xl_namerange')
      Use In Select('xl_validation')
      Use In Select('xl_strformat')
      Use In Select('xl_cellxfs')
      Use In Select('xl_rows')
      Use In Select('xl_ndxcolors')
      Use In Select('xl_mrucolors')
      Use In Select('xl_relationships')
      Use In Select('xl_extsheets')
      Use In Select('xl_extcells')

    Catch To loException
    Endtry

*-- Documentation of the methods and properties for this class
  Protected Procedure Documentation
*-*
*-* Written by:  Vilhelm Praisach
*-*              Gregory Green
*-*
*-* Release History:
*-*
*-*  May 11, 2015     Initial release.
*-*  May 12, 2015     Added new methods: SetCellFormatRange, SetCellBorderRange, SaveGridToWorkbook
*-*                   Added new properties: Title, Subject, CompanyName
*-*                   Added check in SetCellValue that if a value is being set in a merged cell range then the value is set in the upper
*-*                       left cell of the merged range
*-*                   Corrected sheet references when writing the sheet XML
*-*  May 27, 2015     Added new methods: ConvertPixelsToExcelUnits, GetWorkbook, GetWorkbookSheets, DeleteWorkbook, DeleteAllWorkbooks,
*-*                                      SetCellAlignment, SetCellNumberDecimals, SetCellNumberFormat, SetCellNumberFormatRange,
*-*                                      SetColumnWidthRange, SetRowHeight, SetRowHeightRange, SetCellWordWrap
*-*                   Renamed method RemoveSheet to DeleteSheet
*-*                   Corrected string to XML conversion (was performing a double conversion of the string)
*-*                   Corrected issue of generated XLSX file being deleted; increase of Sleep() time for file rename
*-* June 1, 2015      Corrected bug in merge cell and cell value assignment (after a merge cell was performed any subsequent sheets
*-*                        added that had cells in the same merge range with value assignments was not having the values set correctly)
*-*                   Changed property name ExcelXlsxVersion to ExcelXlsxRelease; this property will give the Release number (integer)
*-*                        of this class
*-* June 5, 2015      Added to set the default selected table/cursor (value retuned by ALIAS() function) when the default table/cursor
*-*                        is changed in a method
*-*                   Added new method: SaveTableToWorkbook
*-* June 8, 2015      No new features.  Code was optimized for better performance.
*-*
*-*                   The first optimization I changed the SetCellValue() code to now always attempt to INSERT first. I added the CANDIDATE
*-*                   clause to the primary INDEX so that if a cell is already entered, the INSERT command will fail in a TRY-CATCH block
*-*                   (duplicate KEY); in the CATCH portion I do an update to the cell value. This reduced the time to assign the values
*-*                   significantly. I had been doing a SEEK() for the cell and updating if found otherwise doing an INSERT
*-*
*-*                   The second optimization was how I was assigning a string index value. The original code did a SELECT-SQL WITH a MAX()
*-*                   for the last used index value IN the workbook (strings are universal TO a workbook); I changed to a property value
*-*                   (array indexed by the workbook number) and then just increment this value as it is used. Great amount of savings here too.
*-*
*-*                   The LAST optimization was in how the Sheet.xml were being written. I now just do a SCAN across the cells and add to the
*-*                   XML. I was doing several SQL calls to first get the rows and then SQL calls TO the columns for each row. Again not as
*-*                   efficient as just SCAN-ENDSCAN on the cells cursor table. Had to build in logic to populate rows that did not have a
*-*                   cell value assigned but had row height assigned. The SQL coding was easier for this but at a great cost in time.
*-* June 9, 2015      No new features.  Bug fix for a nasty crash of Excel when the print or print-preview is selected.  Congratulations to
*-*                   Vilhelm for finding this bug.
*-* July 1, 2015      Corrected return value for AddNumericFormat.
*-*                   Changed the parameters on AddNumericFormat(); now just pass the numeric format
*-*
*-* Oct 29, 2015      Added new functions for opening an existing workbook (xlsx).  Added sheet print formatting methods.
*-*
*-************************************************************************************************************************************

*-- Freezes the panes of the sheet
  Procedure FreezePanes
  Lparameters tnWB, tnSheet, tnTopRowCount, tnSideColCount
  Local lxReturn
  Do Case
  Case Pcount() < 3
    Return False

  Case Pcount() = 3
    tnSideColCount = 0
  Endcase
  If Vartype(tnTopRowCount) != 'N'
    tnTopRowCount = 0
  Endif
  If Vartype(tnSideColCount) != 'N'
    tnSideColCount = 0
  Endif
  If This.GetSheetRecord(tnWB, tnSheet)
    Replace xl_sheets.xsplit With tnSideColCount, ;
      xl_sheets.ysplit With tnTopRowCount In xl_sheets
    lxReturn = True
  Else
    lxReturn = False
  Endif
  Return lxReturn

*-- Returns the position of the alpha and numeric split
  Protected Procedure GetAlphaNumericSplit
    Lparameters tcText
    Local lnChr
    For lnChr=1 To Len(tcText)
      If Isdigit(Substr(tcText, lnChr, 1))
        Return lnChr
      Endif
    Endfor
    Return 0

*-- Gets the border record; returns True if found, or False in not found
  Protected Procedure GetBordersRecord
    Lparameters tnWB, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp
    Return Seek(BinToC(tnWB) + tcLStyle + BinToC(tnLColor) + tcRStyle + BinToC(tnRColor) + tcTStyle + BinToC(tnTColor) + tcBStyle + BinToC(tnBColor) + ;
      tcDStyle + BinToC(tnDColor) + BinToC(tnDiagDn) + BinToC(tnDiagUp), "xl_borders", "cellborder")

*-- Returns the cell alignment
  Procedure GetCellAlignment
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loAlignInfo
  If Pcount() < 4
    Return Null
  Endif
  If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And This.GetCellXfsRecord(tnWB, xl_cells.cellxfs)
    loAlignInfo = Createobject("Empty")
    AddProperty(loAlignInfo, "HorzAlign", xl_cellxfs.halign)
    AddProperty(loAlignInfo, "VertAlign", xl_cellxfs.valign)
  Else
    loAlignInfo = Null
  Endif
  Return loAlignInfo

*-- Returns the cell border info
  Procedure GetCellBorders
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loBdrInfo
  If Pcount() < 4
    Return Null
  Endif
  With This
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs) And .GetCursorRecord("xl_borders", tnWB, xl_cellxfs.borderid)
      loBdrInfo = Createobject("Empty")
      AddProperty(loBdrInfo, "LeftStyle",  xl_borders.lstyle)
      AddProperty(loBdrInfo, "LeftColor",  xl_borders.lcolor)
      AddProperty(loBdrInfo, "LeftTheme",  xl_borders.ltheme)
      AddProperty(loBdrInfo, "LeftTint",   xl_borders.ltint)
      AddProperty(loBdrInfo, "LeftIndex",  xl_borders.lindexed)

      AddProperty(loBdrInfo, "RightStyle", xl_borders.rstyle)
      AddProperty(loBdrInfo, "RightColor", xl_borders.rcolor)
      AddProperty(loBdrInfo, "RightTheme", xl_borders.rtheme)
      AddProperty(loBdrInfo, "RightTint",  xl_borders.rtint)
      AddProperty(loBdrInfo, "RightIndex", xl_borders.rindexed)

      AddProperty(loBdrInfo, "TopStyle",   xl_borders.tstyle)
      AddProperty(loBdrInfo, "TopColor",   xl_borders.tcolor)
      AddProperty(loBdrInfo, "TopTheme",   xl_borders.ttheme)
      AddProperty(loBdrInfo, "TopTint",    xl_borders.ttint)
      AddProperty(loBdrInfo, "TopIndex",   xl_borders.tindexed)

      AddProperty(loBdrInfo, "BotStyle",   xl_borders.bstyle)
      AddProperty(loBdrInfo, "BotColor",   xl_borders.bcolor)
      AddProperty(loBdrInfo, "BotTheme",   xl_borders.btheme)
      AddProperty(loBdrInfo, "BotTint",    xl_borders.btint)
      AddProperty(loBdrInfo, "BotIndex",   xl_borders.bindexed)

      AddProperty(loBdrInfo, "DiagStyle",  xl_borders.dstyle)
      AddProperty(loBdrInfo, "DiagColor",  xl_borders.dcolor)
      AddProperty(loBdrInfo, "DiagTheme",  xl_borders.dtheme)
      AddProperty(loBdrInfo, "DiagTint",   xl_borders.dtint)
      AddProperty(loBdrInfo, "DiagIndex",  xl_borders.dindexed)

      AddProperty(loBdrInfo, "DiagDn",     xl_borders.diagdn)
      AddProperty(loBdrInfo, "DiagUp",     xl_borders.diagup)
    Else
      loBdrInfo = Null
    Endif
  Endwith
  Return loBdrInfo

*-- Returns the data type for the selected cell
  Procedure GetCellDataType
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Return Iif(Pcount() >= 4 And This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol), xl_cells.DataType, Null)

*-- Returns the fill info for the cell
  Procedure GetCellFill
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loFillInfo
  If Pcount() < 4
    Return Null
  Endif
  With This
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs) And .GetCursorRecord("xl_fills", tnWB, xl_cellxfs.fillid)
      loFillInfo = Createobject("Empty")
      AddProperty(loFillInfo, "FgColor",   xl_fills.fgcolor)
      AddProperty(loFillInfo, "BgColor",   xl_fills.bgcolor)
      AddProperty(loFillInfo, "PatType",   xl_fills.patttype)
      AddProperty(loFillInfo, "Theme",     xl_fills.Theme)
      AddProperty(loFillInfo, "Tint",      xl_fills.tint)
      AddProperty(loFillInfo, "FgIndexed", xl_fills.fgindexed)
      AddProperty(loFillInfo, "BgIndexed", xl_fills.bgindexed)
    Else
      loFillInfo = Null
    Endif
  Endwith
  Return loFillInfo

*-- Returns the cell font settings
  Procedure GetCellFont
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loFontInfo
  If Pcount() < 4
    Return Null
  Endif
  With This
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs) And .GetCursorRecord("xl_fonts", tnWB, xl_cellxfs.fontid)
      loFontInfo = Createobject("Empty")
      AddProperty(loFontInfo, "FontName",        Alltrim(xl_fonts.fname))
      AddProperty(loFontInfo, "FontSize",        xl_fonts.Fsize)
      AddProperty(loFontInfo, "FontBold",        xl_fonts.fbold)
      AddProperty(loFontInfo, "FontItalic",      xl_fonts.fitalic)
      AddProperty(loFontInfo, "ForeColor",       xl_fonts.fcolor)
      AddProperty(loFontInfo, "FontUnderline",   Alltrim(xl_fonts.uline))
      AddProperty(loFontInfo, "FontStrikeThr",   xl_fonts.strkthr)
      AddProperty(loFontInfo, "FontVerticalPos", Alltrim(xl_fonts.fvpos))
    Else
      loFontInfo = Null
    Endif
  Endwith
  Return loFontInfo

*-- Returns the cell formula expression
  Procedure GetCellFormula
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Return Iif(Pcount() >= 4 And This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol), Alltrim(xl_cells.cellformula), Null)

*-- Gets the cell formula from the XML node
  Protected Procedure GetCellFormulaFromXml
    Lparameters tcFormula, tnCellCol
    Local lcFormula, lcShared, lnId, lcTemp, lnNdx, lcChr, lcAddExpr, lnCol, lcRef, lcRefBeg, lcRefEnd, llColRelative
    Local lnColBase, llRowRelative, lnRowBase
    With This
      lcFormula = .GetNodeElementValue(tcFormula)
      lcShared  = .GetNodeAttributeValue(tcFormula, "t")
      If !Empty(lcShared)
*-*	The formula is a shared formula
        lnId = .GetNodeAttributeValue(tcFormula, "si")
        If Empty(lcFormula)
          If Seek(lnId, "t_shared", "id")
*-*			Shared formula already added; update the column:row relative references
            lcFormula = ""
            lcTemp = Alltrim(t_shared.cellformula)
            Do Case
            Case t_shared.colrelative                && Forumla column is relative
              lnNdx = 1
              Do While lnNdx <= Len(lcTemp)
                lcChr = Substr(lcTemp, lnNdx, 1)
                If Isalpha(lcChr)
                  lcAddExpr = lcChr
                  Do While lnNdx <= Len(lcTemp)
                    lnNdx = lnNdx + 1
                    lcChr = Substr(lcTemp, lnNdx, 1)
                    Do Case
                    Case Isalpha(lcChr)
                      lcAddExpr = lcAddExpr + lcChr

                    Case Isdigit(lcChr)
                      lnCol = .ColumnAsciiToIndex(lcAddExpr)
                      lnCol = lnCol + (tnCellCol - t_shared.colbase)
                      lcAddExpr = .ColumnIndexToAscii(lnCol) + lcChr
                      Exit

                    Otherwise
                      lcAddExpr = lcAddExpr + lcChr
                      Exit
                    Endcase
                  Enddo
                Else
                  lcAddExpr = lcChr
                Endif
                lcFormula = lcFormula + lcAddExpr
                lnNdx = lnNdx + 1
              Enddo

            Case t_shared.rowrelative                && Forumla row is relative
              lnNdx = 1
              Do While lnNdx <= Len(lcTemp)
                lcChr = Substr(lcTemp, lnNdx, 1)
                If Isdigit(lcChr)
                  lcAddExpr = lcChr
                  Do While lnNdx <= Len(lcTemp)
                    lnNdx = lnNdx + 1
                    lcChr = Substr(lcTemp, lnNdx, 1)
                    If Isdigit(lcChr)
                      lcAddExpr = lcAddExpr + lcChr
                    Else
                      If Isalpha(lcChr)
                        lcAddExpr = lcAddExpr + lcChr
                      Else
                        lnRow = Int(Val(lcAddExpr))
                        lnRow = lnRow + (lnCellRow - t_shared.rowbase)
                        lcAddExpr = Transform(lnRow) + lcChr
                      Endif
                      Exit
                    Endif
                  Enddo
                Else
                  lcAddExpr = lcChr
                Endif
                lcFormula = lcFormula + lcAddExpr
                lnNdx = lnNdx + 1
              Enddo
            Endcase
          Else
            lcFormula = ""
          Endif
        Else
          lcRef = .GetNodeAttributeValue(tcFormula, "ref")
          If !Seek(lnId, "t_shared", "id") And Occurs(":", lcRef) > 0
            lcRefBeg = .GetColumnInColRowPair(Getwordnum(lcRef, 1, ":"))
            lcRefEnd = .GetColumnInColRowPair(Getwordnum(lcRef, 2, ":"))
            If lcRefBeg == lcRefEnd
              llColRelative = False
              lnColBase     = 0
              llRowRelative = True
              lnRowBase     = .GetRowInColRowPair(Getwordnum(lcRef, 1, ":"))
            Else
              llColRelative = True
              lnColBase     = .ColumnAsciiToIndex(lcRefBeg)
              llRowRelative = False
              lnRowBase     = 0
            Endif
            Insert Into t_shared (Id, colrelative, colbase, rowrelative, rowbase, cellformula) ;
              VALUES (lnId, llColRelative, lnColBase, llRowRelative, lnRowBase, lcFormula)
          Endif
        Endif
      Endif
    Endwith
    Return lcFormula

*-- Returns the cell indentation
  Procedure GetCellIndent
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local lnIndent
  If Pcount() != 4
    Return -1
  Endif
  With This
    lnIndent = Iif(.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs),;
      xl_cellxfs.indent, 0)
  Endwith
  Return lnIndent

*-- Returns the format code for the selected cell
  Procedure GetCellNumberFormat
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local lnNumFmtId
  If Pcount() != 4
    Return -1
  Endif
  With This
    lnNumFmtId = Iif(.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs),;
      xl_cellxfs.numfmtid, 0)
  Endwith
  Return lnNumFmtId

*-- Returns the selected cell numeric format as text
  Procedure GetCellNumberFormatText
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local lcFormatCode
  If Pcount() != 4
    Return ""
  Endif
  With This
    Return Iif(Pcount() >= 4 And .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs), ;
      xl_cellxfs.formatcode, '')
  Endwith

*-- Positions to the xl_cells record; returns true if found, or false if not found
  Protected Procedure GetCellRecord
    Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
    Return Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnCellRow)+BinToC(tnCellCol), "xl_cells", "cellindex")

*-- Gets the cell style value
  Procedure GetCellStyle
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Return Iif(Pcount() = 4 And This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol), xl_cells.cellxfs, -1)

*-- Gets the cell text rotation value
  Procedure GetCellTextRotation
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local lnRotation
  If Pcount() != 4
    Return 99
  Endif
  With This
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
      lnRotation = xl_cellxfs.Rotation
      If lnRotation > 90
        lnRotation = 90 - lnRotation
      Endif
    Else
      lnRotation = 0
    Endif
  Endwith
  Return lnRotation

*-- Returns the cell validation settings
  Procedure GetCellValidation
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loValidation
  loValidation = Createobject("Empty")
  AddProperty(loValidation, "Type", 0)
  AddProperty(loValidation, "Style", 0)
  AddProperty(loValidation, "Operator", 0)
  AddProperty(loValidation, "AllowBlank", False)
  AddProperty(loValidation, "ShowInputMsg", False)
  AddProperty(loValidation, "ShowErrMsg", False)
  AddProperty(loValidation, "ErrMsg", "")
  AddProperty(loValidation, "ErrTitle", "")
  AddProperty(loValidation, "Prompt", "")
  AddProperty(loValidation, "Formula1", "")
  AddProperty(loValidation, "Formula2", "")
  If Pcount() < 4
    Return loValidation
  Endif
  If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And Seek(BinToC(xl_cells.validndx), "xl_validation", "validndx")
    With loValidation
      .Type         = xl_validation.vtype
      .Style        = xl_validation.vstyle
      .Operator     = xl_validation.voperator
      .allowblank   = xl_validation.allowblank
      .ShowInputMsg = xl_validation.showinpmsg
      .showerrmsg   = xl_validation.showerrmsg
      .errmsg       = xl_validation.errmsg
      .errtitle     = xl_validation.errtitle
      .Prompt       = xl_validation.vprompt
      .formula1     = xl_validation.formula1
      .formula2     = xl_validation.formula2
    Endwith
  Endif
  Return loValiation

*-- Returns the value of the cell
  Procedure GetCellValue
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local lxValue, lnDays, lnDecimals, lnMSec, lnTime, lnSize, lcType, loException
  If Pcount() < 4
    Return Null
  Endif
  lxValue = Null
  With This
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      lnDecimals = Set("DECIMALS")
      Try
        Do Case
        Case xl_cells.DataType = DATA_TYPE_DATE
          lnDays  = Cast(xl_cells.cellvalue As I)
          lxValue = Date(1900, 1, 1) + lnDays - 2

        Case xl_cells.DataType = DATA_TYPE_TIME
          Set Decimals To 12
          lnMSec  = Cast(xl_cells.cellvalue As N(20,12))
          lnTime  = lnMSec * 24 * 60 * 60
          lxValue = Dtot(Date()) + lnTime
          lxValue = Transform(Hour(lxValue)) + ":" + Transform(Minute(lxValue)) + ":" + Transform(Sec(lxValue))

        Case xl_cells.DataType = DATA_TYPE_DATETIME
          Set Decimals To 12
          lnDays  = Cast(xl_cells.cellvalue As I)
          lxValue = Date(1900, 1, 1) + lnDays - 2
          lnMSec  = Cast(xl_cells.cellvalue As N(20,12)) - lnDays
          lnTime  = lnMSec * 24 * 60 * 60
          lxValue = Dtot(lxValue) + lnTime

        Case xl_cells.DataType = DATA_TYPE_CHAR
          lxValue = Iif(.GetCursorRecord("xl_strings", tnWB, xl_cells.stringid), Rtrim(xl_strings.stringval), '')

        Case xl_cells.DataType = DATA_TYPE_INT
          lxValue = Cast(xl_cells.cellvalue As I)
          lxValue = Iif(lxValue != Val(xl_cells.cellvalue), Cast(xl_cells.cellvalue As N(20,0)), lxValue)

        Case xl_cells.DataType = DATA_TYPE_FLOAT
          lnSize = Atc(".", Alltrim(xl_cells.cellvalue))
          lnSize = Iif(lnSize > 0, lnSize + lnDecimals, Len(Alltrim(xl_cells.cellvalue)) + lnDecimals + 1)
          lcType  = "N(" + Transform(lnSize) + "," + Transform(lnDecimals) + ")"
          lxValue = Cast(xl_cells.cellvalue As &lcType)

        Case xl_cells.DataType = DATA_TYPE_CURRENCY
          lxValue = Cast(xl_cells.cellvalue As Y)

        Case xl_cells.DataType = DATA_TYPE_PERCENT
          lxValue = Transform(Cast(xl_cells.cellvalue As N(14,8)) * 100) + "%"

        Case xl_cells.DataType = DATA_TYPE_GENERAL
          lxValue = Alltrim(xl_cells.cellvalue)

        Case xl_cells.DataType = DATA_TYPE_FORMULA

        Case xl_cells.DataType = DATA_TYPE_NONE
          lxValue = Alltrim(xl_cells.cellvalue)

        Otherwise
          lxValue = Alltrim(xl_cells.cellvalue)
        Endcase

      Catch To loException
        Set Step On
      Finally
        Set Decimals To &lnDecimals
      Endtry
    Endif
  Endwith
  Return lxValue

*-- Returns the cell word wrap setting
  Procedure GetCellWordWrap
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Return Iif(Pcount() >= 4 And This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And This.GetCellXfsRecord(tnWB, xl_cells.cellxfs),;
    xl_cellxfs.wraptext, Null)

*-- Returns the Cell XfsId value
  Protected Procedure GetCellXfsId
    Lparameters tnWB, tnNumFmtId, tnFontId, tnFillId, tnBorderId, tcHAlign, tcVAlign, tnIndex, tnWrapText, tnRotation
    Local lnCellXfsId
    If Seek(BinToC(tnWB)+BinToC(tnNumFmtId)+BinToC(tnFontId)+BinToC(tnFillId)+BinToC(tnBorderId)+BinToC(tnIndex)+BinToC(tnWrapText)+BinToC(tnRotation)+Alltrim(tcHAlign)+Alltrim(tcVAlign), "xl_cellxfs", "cellxfs")
      lnCellXfsId = xl_cellxfs.Id
    Else
      lnCellXfsId = This.GetNextId(tnWB, 'xl_cellxfs')
      Insert Into xl_cellxfs (workbook, Id, numfmtid, fontid, fillid, borderid, halign, valign, indent, wraptext, Rotation) ;
        VALUES (tnWB, lnCellXfsId, tnNumFmtId, tnFontId, tnFillId, tnBorderId, tcHAlign, tcVAlign, tnIndex, tnWrapText, tnRotation)
    Endif
    Return lnCellXfsId

*-- Positions the record pointer to xl_cellxfs record; returns true if found, or false if not found
  Protected Procedure GetCellXfsRecord
    Lparameters tnWB, tnId
    Return Seek(BinToC(tnWB)+BinToC(tnId), "xl_cellxfs", "id")

*-- Returns the checksum the string
  Protected Procedure GetCheckSum
    Lparameters tcString
    Local lcString, lcCheckSum
    lcString   = Alltrim(tcString)
    lcCheckSum = Sys(2007, lcString, 0, 1)
    lcCheckSum = lcCheckSum + Padr(lcString, (230-Len(lcCheckSum)), "_")
    Return lcCheckSum

*-- Returns the number of child nodes
  Protected Procedure GetChildNodeCount
    Lparameters tcNode
    Return Occurs("<", tcNode) - 2

*-- Returns the selected child node
  Protected Procedure GetChildNodeElement
    Lparameters tcNode, tnChild
    Local lcElement
    lcElement = Strextract(tcNode, '<', '>', tnChild+1, 1)
    Return Strextract(tcNode, '<'+lcElement+'>', '</'+lcElement+'>', 1, 5)

*-- Returns the column (Ascii value) in a ColRow pair; i.e., returns A from A6 reference
  Protected Procedure GetColumnInColRowPair
    Lparameters tcColRow
    Local lnNdx, lcCol, lcChr
    lcCol = ""
    For lnNdx=1 To Len(tcColRow)
      lcChr = Substr(tcColRow, lnNdx, 1)
      If Isdigit(lcChr)
        Exit
      Endif
      lcCol = lcCol + lcChr
    Endfor
    Return lcCol

*-- Returns the column width
  Procedure GetColumnWidth
  Lparameters tnWB, tnSheet, tnColumn
  Local lxReturn
  If Pcount() < 3
    Return Null
  Endif
  If This.GetSheetRecord(tnWB, tnSheet)
    lxReturn = Iif(Seek(This.GetColWidthKeyExpr(tnWB, tnSheet, tnColumn), "xl_colwidths", "shcolumn"), xl_colwidths.Width, -1)
  Else
    lxReturn = Null
  Endif
  Return lxReturn

*-- Gets the key expression for c_colwidths
  Protected Procedure GetColWidthKeyExpr
    Lparameters tnWB, tnSheet, tnColumn
    Return BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnColumn)

*-- Gets the specified xl_cursor record; returns true if found, false if not found
  Protected Procedure GetCursorRecord
    Lparameters tcCursor, tnWB, tnId
    Return Seek(BinToC(tnWB)+BinToC(tnId), tcCursor, "id")

*-- Returns the format code for the specified custom numeric format
  Procedure GetCustomNumericFormat
  Lparameters tnWB, tnNumFormat
  Return Iif(Seek(BinToC(tnWB)+BinToC(tnNumFormat), "xl_numfmts", "id"), Alltrim(xl_numFmts.formatcode), '')

*-- Gets the values for the custom paper size
  Procedure GetCustomPaperSize
  Lparameters tnWB, tnSheet
  Local loReturn
  If Pcount() = 2 And This.GetSheetRecord(tnWB, tnSheet)
    loReturn = Createobject("Empty")
    AddProperty(loReturn, "PaperWidth",  xl_sheets.Paperwidth)
    AddProperty(loReturn, "PaperHeight", xl_sheets.paperheight)
    AddProperty(loReturn, "PaperDimen",  xl_sheets.paperdimen)
  Else
    loReturn = Null
  Endif
  Return loReturn

*-- Gets the key expression for t_format
  Protected Procedure GetFormatKeyExpr
    Lparameters tnNbrFmtId, tnNumDec, tnFontId, tnFillId, tnBorderId, tcHAlign, tcVAlign, tlWordWrap
    Return BinToC(tnNbrFmtId)+BinToC(tnNumDec)+BinToC(tnFontId)+BinToC(tnFillId)+BinToC(tnBorderId)+tcHAlign+tcVAlign+Transform(tlWordWrap)

*-- Gets the key expression for the header footer table record
  Protected Procedure GetHeaderFooterKeyExpr
    Lparameters tnWB, tnSheet, tnPage, tnSection
    Return BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnPage)+BinToC(tnSection)

*-- Returns the inline formatting definition of cell text for each character group
  Procedure GetInlineFontDefinition
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local loSharedString, loInline, lcAlias, lnBegPos, lnLen, loCharacter
  If Pcount() != 4
    Return Null
  Endif
  loSharedString = This.GetSharedStringId(tnWB, tnSheet, tnCellRow, tnCellCol)
  If Isnull(loSharedString)
    Return Null
  Endif
  lnBegPos = 1
  lcAlias  = Alias()
  loInline = Createobject("Empty")
  AddProperty(loInline, "Workbook", tnWB)
  AddProperty(loInline, "StringId", loSharedString.stringid)
  AddProperty(loInline, "StringValue", loSharedString.StringValue)
  AddProperty(loInline, "Count", 0)
  AddProperty(loInline, "Characters[1]")
  Select xl_strformat
  Set Order To Tag sortorder
  Scan For workbook = tnWB And Id = loSharedString.stringid And !Deleted()
    lnLen = Len(Rtrim(xl_strformat.stringxml))
    loCharacter = Createobject("Empty")
    AddProperty(loCharacter, "BegPos", lnBegPos)
    AddProperty(loCharacter, "Length", lnLen)
    AddProperty(loCharacter, "FontName", xl_strformat.fname)
    AddProperty(loCharacter, "FontSize", xl_strformat.Fsize)
    AddProperty(loCharacter, "FontBold", xl_strformat.fbold)
    AddProperty(loCharacter, "FontItalic", xl_strformat.fitalic)
    AddProperty(loCharacter, "FontColor", xl_strformat.fcolor)
    AddProperty(loCharacter, "Underline", xl_strformat.uline)
    AddProperty(loCharacter, "StrikeThru", xl_strformat.strkthr)
    AddProperty(loCharacter, "SubScript", Iif(xl_strformat.fvpos=FONT_VERTICAL_SUBSCRIPT, True, False))
    AddProperty(loCharacter, "SuperScript", Iif(xl_strformat.fvpos=FONT_VERTICAL_SUPERSCRIPT, True, False))
    loInline.Count = loInline.Count + 1
    Dimension loInline.Characters[loInline.Count]
    loInline.Characters[loInline.Count] = loCharacter
    lnBegPos = lnBegPos + lnLen
  Endscan
  Set Order To
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  If loInline.Count = 0
    loInline = Null
  Endif
  Return loInline

*-- Returns the last column index in the row
  Procedure GetLastColumnInRow
  Lparameters tnWB, tnSheet, tnCellRow
  Local lnMaxCol
  If Pcount() < 3
    Return Null
  Endif
  lnMaxCol = 0
  Select Max(tmp.cellcol) As maxcol From xl_cells As tmp ;
    WHERE workbook = tnWB ;
    AND tmp.Sheet = tnSheet ;
    AND tmp.cellrow = tnCellRow ;
    AND tmp.celldeleted = False ;
    INTO Cursor t_temp

  If _Tally > 0
    lnMaxCol = t_temp.maxcol
  Endif
  If Used('t_temp')
    Use In t_temp
  Endif
  Return lnMaxCol

*-- Returns the last row number in the sheet
  Procedure GetLastRowNumber
  Lparameters tnWB, tnSheet
  Local Array laMaxRow[1]
  If Pcount() < 2
    Return Null
  Endif
  Select Max(cellrow) From xl_cells ;
    WHERE workbook = tnWB ;
    AND Sheet = tnSheet ;
    AND celldeleted = False ;
    INTO Array laMaxRow

  Return Iif(_Tally > 0, laMaxRow[1], 0)

*-- Returns the max column number for a sheet
  Procedure GetMaxColumnNumber
  Lparameters tnWB, tnSheet
  Local Array laMaxCol[1]
  If Pcount() < 2
    Return Null
  Endif
  Select Max(tmp.maxcol) As maxcol From xl_rows As tmp ;
    WHERE tmp.workbook = tnWB ;
    AND tmp.Sheet = tnSheet ;
    INTO Array laMaxCol

  Return Iif(_Tally > 0, laMaxCol[1], 0)

*-- Gets the next sequential Id value
  Protected Procedure GetNextId
    Lparameters tnWB, tcCursor
    Local lnNdx, lnNextId
    With This
      Do Case
      Case tcCursor = "xl_borders"
        lnNdx = Ascan(.LastBorderId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastBorderId[lnNdx, 2] = .LastBorderId[lnNdx, 2] + 1
          lnNextId = .LastBorderId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastBorderId, 1) + 1)
          Dimension .LastBorderId[lnNdx, 2]
          .LastBorderId[lnNdx, 1] = tnWB
          .LastBorderId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_cellxfs"
        lnNdx = Ascan(.LastCellXfsId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastCellXfsId[lnNdx, 2] = .LastCellXfsId[lnNdx, 2] + 1
          lnNextId = .LastCellXfsId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastCellXfsId, 1) + 1)
          Dimension .LastCellXfsId[lnNdx, 2]
          .LastCellXfsId[lnNdx, 1] = tnWB
          .LastCellXfsId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_fills"
        lnNdx = Ascan(.LastFillId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastFillId[lnNdx, 2] = .LastFillId[lnNdx, 2] + 1
          lnNextId = .LastFillId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastFillId, 1) + 1)
          Dimension .LastFillId[lnNdx, 2]
          .LastFillId[lnNdx, 1] = tnWB
          .LastFillId[lnNdx, 2] = 2
          lnNextId = 2
        Endif

      Case tcCursor = "xl_fonts"
        lnNdx = Ascan(.LastFontId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastFontId[lnNdx, 2] = .LastFontId[lnNdx, 2] + 1
          lnNextId = .LastFontId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastFontId, 1) + 1)
          Dimension .LastFontId[lnNdx, 2]
          .LastFontId[lnNdx, 1] = tnWB
          .LastFontId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_ndxcolors"
        lnNdx = Ascan(.LastIndexColorId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastIndexColorId[lnNdx, 2] = .LastIndexColorId[lnNdx, 2] + 1
          lnNextId = .LastIndexColorId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastIndexColorId, 1) + 1)
          Dimension .LastIndexColorId[lnNdx, 2]
          .LastIndexColorId[lnNdx, 1] = tnWB
          .LastIndexColorId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_mrucolors"
        lnNdx = Ascan(.LastMruColorId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastMruColorId[lnNdx, 2] = .LastMruColorId[lnNdx, 2] + 1
          lnNextId = .LastMruColorId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastMruColorId, 1) + 1)
          Dimension .LastMruColorId[lnNdx, 2]
          .LastMruColorId[lnNdx, 1] = tnWB
          .LastMruColorId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_numfmts"
        lnNdx = Ascan(.LastNumFmtId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastNumFmtId[lnNdx, 2] = .LastNumFmtId[lnNdx, 2] + 1
          lnNextId = .LastNumFmtId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastNumFmtId, 1) + 1)
          Dimension .LastNumFmtId[lnNdx, 2]
          .LastNumFmtId[lnNdx, 1] = tnWB
          .LastNumFmtId[lnNdx, 2] = START_NUMERIC_FORMAT_ID
          lnNextId = START_NUMERIC_FORMAT_ID
        Endif

      Case tcCursor = "xl_relationships"
        lnNdx = Ascan(.LastRelationId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastRelationId[lnNdx, 2] = .LastRelationId[lnNdx, 2] + 1
          lnNextId = .LastRelationId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastRelationId, 1) + 1)
          Dimension .LastRelationId[lnNdx, 2]
          .LastRelationId[lnNdx, 1] = tnWB
          .LastRelationId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_sheets"
        lnNdx = Ascan(.LastSheetId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastSheetId[lnNdx, 2] = .LastSheetId[lnNdx, 2] + 1
          lnNextId = .LastSheetId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastSheetId, 1) + 1)
          Dimension .LastSheetId[lnNdx, 2]
          .LastSheetId[lnNdx, 1] = tnWB
          .LastSheetId[lnNdx, 2] = 1
          lnNextId = 1
        Endif

      Case tcCursor = "xl_strings"
        lnNdx = Ascan(.LastStringId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastStringId[lnNdx, 2] = .LastStringId[lnNdx, 2] + 1
          lnNextId = .LastStringId[lnNdx, 2]
        Else
          lnNdx = Max(1, Alen(.LastStringId, 1) + 1)
          Dimension .LastStringId[lnNdx, 2]
          .LastStringId[lnNdx, 1] = tnWB
          .LastStringId[lnNdx, 2] = 0
          lnNextId = 0
        Endif

      Otherwise
        lnNextId = 0
      Endcase
    Endwith
    Return lnNextId

*-- Returns the attribute value
  Protected Procedure GetNodeAttributeValue
    Lparameters tcXf, tcAttrName
    Return Strextract(tcXf, tcAttrName+'="', '"', 1)

*-- Returns the number of nodes to be processed
  Protected Procedure GetNodeCount
    Lparameters tcNode, tcElement
    Local lnCount
    lnCount = Occurs("<"+tcElement+">", tcNode)
    If lnCount = 0
      lnCount = Occurs("<"+tcElement+" ", tcNode)
      If lnCount = 0
        lnCount = Occurs("<"+tcElement, tcNode)
      Endif
    Endif
    Return lnCount

*-- Gets the string content for the node element
  Protected Procedure GetNodeElement
    Lparameters tcNode, tcElement, tnNode, tlStrict, tnFlag
    Local lcContent, lnNdx, loException
    Do Case
    Case Pcount() = 3
      tlStrict = False
      tnFlag   = 5

    Case Pcount() = 4
      tnFlag = 5
    Endcase
    If tlStrict
      lcContent = Strextract(tcNode, "<"+tcElement+">", "</"+tcElement+">", tnNode, tnFlag)
    Else
      lcContent = Strextract(tcNode, "<"+tcElement+" ", ">", tnNode, tnFlag)
      If Right(lcContent, 2) != "/>"
        lcContent = Strextract(tcNode, "<"+tcElement+" ", "</"+tcElement+">", tnNode, tnFlag)
        If Empty(lcContent)
          lcContent = Strextract(tcNode, "<"+tcElement+">", "</"+tcElement+">", tnNode, tnFlag)
        Endif
        If Empty(lcContent)
          lcContent = Strextract(tcNode, "<"+tcElement, "/>", tnNode, tnFlag)
        Endif
      Endif
    Endif
    Return lcContent

*-- Gets the value assigned to the node element
  Protected Procedure GetNodeElementValue
    Lparameters tcElement
    Return Strextract(tcElement, '>', '<', 1, 1)

*-- Returns the number of sheets for the given workbook
  Procedure GetNumberDfSheets
  Lparameters tnWB
  Local lnCount
  Select xl_sheets
  Count For xl_sheets.workbook = tnWB And !xl_sheets.shdeleted To lnCount
  Return lnCount

*-- Gets the paper size for the selected sheet
  Procedure GetPaperSize
  Lparameters tnWB, tnSheet
  Return Iif(Pcount() = 2 And This.GetSheetRecord(tnWB, tnSheet), xl_sheets.Papersize, -1)

*-- Gets the print orientation for the sheet output
  Procedure GetPrintOrientation
  Lparameters tnWB, tnSheet
  Return Iif(Pcount() = 2 And This.GetSheetRecord(tnWB, tnSheet), xl_sheets.prnorient, 0)

*-- Gets the file name for the R:Id value in the relationships XML
  Protected Procedure GetRelationShipFileName
    Lparameters tnWB, tnShRId
    Return Iif(Seek(BinToC(tnWB)+BinToC(tnShRId), "xl_relationships", "relid"), Alltrim(xl_relationships.Target), '')

*-- Gets the specified RGB color value
  Procedure GetRGBValues
  Lparameters tnColorValue, tcRGB
  Local lnValue, lnNdx
  If Pcount() < 2
    Return -1
  Endif
  Do Case
  Case tcRGB = "R"
    lnValue = tnColorValue
    For lnNdx=8 To 31
      lnValue = Bitclear(lnValue, lnNdx)
    Endfor

  Case tcRGB = "G"
    lnValue = Bitrshift(tnColorValue, 8)
    For lnNdx=8 To 31
      lnValue = Bitclear(lnValue, lnNdx)
    Endfor

  Case tcRGB = "B"
    lnValue = Bitrshift(tnColorValue, 16)
    For lnNdx=8 To 31
      lnValue = Bitclear(lnValue, lnNdx)
    Endfor

  Otherwise
    Return -1
  Endcase
  Return Int(lnValue)

*-- Gets the key expression for the c_rowheights
  Protected Procedure GetRowHeightKeyExpr
    Lparameters tnWB, tnSheet, tnRow
    Return BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnRow)

*-- REturns the row (integer value) in a ColRow pair; i.e., return 6 from A6 reference
  Protected Procedure GetRowInColRowPair
    Lparameters tcColRow
    Local lcRow, lnNdx, lcChr
    lcRow = ""
    For lnNdx=1 To Len(tcColRow)
      lcChr = Substr(tcColRow, lnNdx, 1)
      If Isdigit(lcChr)
        lcRow = lcRow + lcChr
      Endif
    Endfor
    Return Int(Val(lcRow))

*-- Returns the selected row maximum column
  Procedure GetRowMaxColumn
  Lparameters tnWB, tnSheet, tnCellRow
  Return Iif(Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnCellRow), "xl_rows", "id"), xl_rows.maxcol, 0)

*-- Gets the Shared Strings Id value for the cell text
  Protected Procedure GetSharedStringId
    Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
    Local loSharedString
    loSharedString = Null
    If Pcount() < 4
      Return loSharedString
    Endif
    If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol) And This.GetCursorRecord("xl_strings", tnWB, xl_cells.stringid)
      loSharedString = Createobject("Empty")
      AddProperty(loSharedString, "StringId", xl_cells.stringid)
      AddProperty(loSharedString, "StringValue", xl_strings.stringval)
    Endif
    Return loSharedString

*-- Gets the sheet name
  Procedure GetSheetName
  Lparameters tnWB, tnSheet
  Return Iif(Pcount() = 2 And This.GetSheetRecord(tnWB, tnSheet), Alltrim(xl_sheets.shname), '')

*-- Positions the record in xl_sheets; returns true if record found, or false if not found
  Protected Procedure GetSheetRecord
    Lparameters tnWB, tnSheet
    Return Seek(BinToC(tnWB)+BinToC(tnSheet), "xl_sheets", "sheetndx")

*-- Returns the cell values for the given row
  Procedure GetSheetRowValues
  Lparameters tnWB, tnSheet, tnCellRow
  Local lnMaxCol, loRow
  If Pcount() < 3
    Return Null
  Endif
  lnMaxCol = This.GetMaxColumnNumber(tnWB, tnSheet)
  With This
    If !Isnull(lnMaxCol) And lnMaxCol > 0
      loRow = Createobject("Empty")
      AddProperty(loRow, "Values[lnMaxCol, 2]", "")
      AddProperty(loRow, "Count", lnMaxCol)
      loRow.Values = Null
      Select tmp.cellcol From xl_cells As tmp ;
        WHERE workbook = tnWB ;
        AND tmp.Sheet = tnSheet ;
        AND tmp.cellrow = tnCellRow ;
        AND tmp.celldeleted = False ;
        INTO Cursor t_temp

      Select t_temp
      Scan
        loRow.Values[t_temp.cellcol, 1] = .GetCellValue(tnWB, tnSheet, tnCellRow, t_temp.cellcol)
        loRow.Values[t_temp.cellcol, 2] = .GetCellDataType(tnWB, tnSheet, tnCellRow, t_temp.cellcol)
      Endscan
      Use In t_temp
    Else
      loRow = Null
    Endif
  Endwith
  Return loRow

*-- Gets the sheet printing scale
  Procedure GetSheetScale
  Lparameters tnWB, tnSheet
  Return Iif(Pcount() = 2 And This.GetSheetRecord(tnWB, tnSheet), xl_sheets.Scale, -1)

*-- Positions the record pointer to the selected string record
  Protected Procedure GetStringRecord
    Lparameters tnWB, tcString
    Return Seek(BinToC(tnWB)+This.GetCheckSum(tcString), "xl_strings", "stringndx")

*-- Returns a converted string from XML to text
  Protected Procedure GetStringXml
    Lparameters tcString
    Local lcString, lcOldChr, lcChr, lnChr, lnNdx
    If Left(tcString, 1) != " " And Empty(tcString)
      Return ""
    Endif
    Try
      If Isnull(This.oXDom)
        lcString = Strtran(tcString, '&amp;', Chr(38))
        lcString = Strtran(lcString, '&gt;', '>')
        lcString = Strtran(lcString, '&lt;', '<')
        lcString = Strtran(lcString, '&quot;', '"')
        lnNdx = 1
        Do While Atc("&#", lcString, lnNdx) > 0
          lcOldChr = Strextract(lcString, "&#", ";", lnNdx)
          lnChr    = Int(Val(lcOldChr))
          If lnChr < 256
            lcChr    = Chr(lnChr)
            lcOldChr = "&#" + lcOldChr + ";"
            lcString = Strtran(lcString, lcOldChr, lcChr)
          Else
            lnNdx = lnNdx + 1            && Increment to skip this embedded value
          Endif
        Enddo
      Else
        This.oXDom.LoadXML("<value>" + tcString + "</value>")
        lcString = This.oXDom.Text
      Endif

    Catch To loException
      Set Step On
    Endtry
    Return lcString

*-- Returns the validation definition
  Procedure GetValidation
  Lparameters tnValidNdx
  Local loValidation
  loValidation = Createobject("Empty")
  AddProperty(loValidation, "Type", 0)
  AddProperty(loValidation, "Style", 0)
  AddProperty(loValidation, "Operator", 0)
  AddProperty(loValidation, "AllowBlank", False)
  AddProperty(loValidation, "ShowInputMsg", False)
  AddProperty(loValidation, "ShowErrMsg", False)
  AddProperty(loValidation, "ErrMsg", "")
  AddProperty(loValidation, "ErrTitle", "")
  AddProperty(loValidation, "Prompt", "")
  AddProperty(loValidation, "Formula1", "")
  AddProperty(loValidation, "Formula2", "")
  If Pcount() < 1
    Return loValidation
  Endif
  If Seek(BinToC(tnValidNdx), "xl_validation", "validndx")
    With loValidation
      .Type         = xl_validation.vtype
      .Style        = xl_validation.vstyle
      .Operator     = xl_validation.voperator
      .allowblank   = xl_validation.allowblank
      .ShowInputMsg = xl_validation.showinpmsg
      .showerrmsg   = xl_validation.showerrmsg
      .errmsg       = xl_validation.errmsg
      .errtitle     = xl_validation.errtitle
      .Prompt       = xl_validation.vprompt
      .formula1     = Iif(xl_validation.formula, "=", "") + Alltrim(xl_validation.formula1)
      .formula2     = Iif(Empty(xl_validation.formula2), "", Iif(xl_validation.formula, "=", "") + Alltrim(xl_validation.formula2))
    Endwith
  Endif
  Return loValidation

*-- Returns a list of validations for the workbook/sheet
  Procedure GetValidationList
  Lparameters tnWB, tnSheet
* TODO: заменить на copy fields vtype, validndx to array loValidation.List
  Local loValiation, lcAlias
  lcAlias = Alias()
  loValidation = Createobject("Empty")
  AddProperty(loValidation, "List[1]")
  AddProperty(loValidation, "Count", 0)
  Select xl_validation
  Scan For BinToC(workbook)+BinToC(Sheet)
    loValidation.Count = loValidation.Count + 1
    Dimension loValidation.List[loValidation.Count, 2]
    loValidation.List[loValidation.Count, 1] = xl_validation.vtype
    loValidation.List[loValidation.Count, 2] = xl_validation.validndx
  Endscan
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return loValidation

*-- Gets the workbook Id from the name
  Procedure GetWorkbook
  Lparameters tcName
  Return Iif(Seek(Padr(Upper(Juststem(tcName)), 200, "_"), "xl_workbooks", "wbname"), xl_workbooks.workbook, 0)

*-- Gets the file name for the workbook index
  Procedure GetWorkbookFileName
  Lparameters tnWB
  Return Iif(Seek(BinToC(tnWB), "xl_workbooks", "workbook"), ;
    FORCEEXT(Addbs(Alltrim(xl_workbooks.filepath)) + Alltrim(xl_workbooks.wbname), "xlsx"), '')

*-- Returns the workbook sheets
  Procedure GetWorkbookSheets
  Lparameters tnWB
  Local loSheets
  loSheets = Createobject("Empty")
  AddProperty(loSheets, "List[1, 2]")
  AddProperty(loSheets, "Count", 0)
  Select xl_sheets
  Set Order To Tag Sheet
* TODO: заменить на copy fields Sheet, shname to array loSheets.List
* TODO: loSheets.Count = RecCount()
  Scan For workbook = tnWB And !shdeleted
    loSheets.Count = loSheets.Count + 1
    Dimension loSheets.List[loSheets.Count, 2]
    loSheets.List[loSheets.Count, 1] = xl_sheets.Sheet
    loSheets.List[loSheets.Count, 2] = Alltrim(xl_sheets.shname)
  Endscan
  Set Order To
  Return loSheets

*-- Returns a XML formatted string for the header and footer text
  Protected Procedure GetXmlHeaderFooterText
    Local lcText
    If Empty(xl_hdrfootertext.FontName)
      lcText = '"-,'
    Else
      lcText = '"' + Alltrim(xl_hdrfootertext.FontName) + ','
    Endif
    Do Case
    Case xl_hdrfootertext.fontstyle = HEADERFOOTER_FONT_STYLE_BOLD
      lcText = lcText + 'Bold"'

    Case xl_hdrfootertext.fontstyle = HEADERFOOTER_FONT_STYLE_ITALIC
      lcText = lcText + 'Italic"'

    Case xl_hdrfootertext.fontstyle = HEADERFOOTER_FONT_STYLE_BOLDITALIC
      lcText = lcText + 'Bold Italic"'

    Otherwise
      lcText = lcText + 'Regular"'
    Endcase
    If xl_hdrfootertext.FontSize > 0
      lcText = lcText + '&amp;' + Transform(xl_hdrfootertext.FontSize)
    Endif
    If xl_hdrfootertext.fontcolor > 0
      lcText = lcText + '&amp;K' + This.ConvertColorToHex(xl_hdrfootertext.fontcolor)
    Endif
    lcText = lcText + Alltrim(xl_hdrfootertext.Text)
    Return lcText

*-- Gets the string as an XML formatted string
  Protected Procedure GetXMLString
    Lparameters tcString
    Local lcString, lcXMLString, lnNdx, lcChar, lcInvalidChars, loNodeText
    If Left(tcString, 1) != " " And Empty(tcString)
      Return ""
    Endif
*-*	TJF 2016-12-12 Strip out invalid XML characters (ASCII 0 to 31 except 9 (TAB), 10 (LF), and 13 (CR))
    lcInvalidChars = Chr(0) + Chr(1) + Chr(2) + Chr(3) + Chr(4) + Chr(5) + Chr(6) + Chr(7) + Chr(8) + Chr(11) + Chr(12) + ;
      CHR(14) + Chr(15) + Chr(16) + Chr(17) + Chr(18) + Chr(19) + Chr(20) + Chr(21) + Chr(22) + Chr(23) + ;
      CHR(24) + Chr(25) + Chr(26) + Chr(27) + Chr(28) + Chr(29) + Chr(30) + Chr(31)
    lcString = Chrtran(tcString, lcInvalidChars, "")
    If Isnull(This.oXDom)
      lcString = Strtran(lcString, Chr(38), '&amp;')
      lcString = Strtran(lcString, '>', '&gt;')
      lcString = Strtran(lcString, '<', '&lt;')
      lcString = Strtran(lcString, '"', '&quot;')

      If This.Codepage = 0
        lcXMLString = ""
        For lnNdx=1 To Len(lcString)
          lcChar = Substr(lcString, lnNdx, 1)
          lnChar = Asc(lcChar)
          lcXMLString = lcXMLString + Iif(lnChar < 128, lcChar, "&#" + Transform(lnChar) + ";")
        Endfor
      Else
        lcXMLString = lcString
      Endif
    Else
      loNodeText = This.oXDom.createTextNode(lcString)
      lcXMLString = loNodeText.XML
      lcXMLString = Strtran(lcXMLString, '"', '&quot;')
    Endif
    Return lcXMLString

*-- Initializes the internal Id values
  Protected Procedure InitalizeIdValues
    With This
      Store 0 To .LastBorderId, .LastCellXfsId, .LastFillId, .LastFontId, .LastIndexColorId, .LastMruColorId, ;
        .LastNumFmtId, .LastRelationId, .LastSheetId, .LastStringId
    Endwith

*-- Inserts a new cell into the sheet
  Procedure InsertCell
  Lparameters tnWB, tnSheet, tnRow, tnCol, tnShift
  Local llReturn, lnRow
  If Pcount() < 5
    Return False
  Endif
  With This
    If .GetSheetRecord(tnWB, tnSheet)
      Select xl_cells
      Set Order To Tag reverse
      Do Case
      Case tnShift = INSERT_LEFT
        Scan For workbook = tnWB And Sheet = tnSheet And cellcol >= tnCol And cellrow = tnRow
          Replace xl_cells.cellcol With xl_cells.cellcol + 1 In xl_cells
          .SetRowMaxColumn(tnWB, tnSheet, xl_cells.cellrow, xl_cells.cellcol)
        Endscan
        llReturn = True

      Case tnShift = INSERT_RIGHT
        Scan For workbook = tnWB And Sheet = tnSheet And cellcol > tnCol And cellrow = tnRow
          Replace xl_cells.cellcol With xl_cells.cellcol + 1 In xl_cells
          .SetRowMaxColumn(tnWB, tnSheet, xl_cells.cellrow, xl_cells.cellcol)
        Endscan
        llReturn = True

      Case tnShift = INSERT_BEFORE
        Scan For workbook = tnWB And Sheet = tnSheet And cellrow >= tnRow And cellcol = tnCol
          Replace xl_cells.cellrow With xl_cells.cellrow + 1 In xl_cells
          .SetRowMaxColumn(tnWB, tnSheet, xl_cells.cellrow, xl_cells.cellcol)
        Endscan
        llReturn = True

      Case tnShift = INSERT_AFTER
        Scan For workbook = tnWB And Sheet = tnSheet And cellrow > tnRow And cellcol = tnCol
          Replace xl_cells.cellrow With xl_cells.cellrow + 1 In xl_cells
          .SetRowMaxColumn(tnWB, tnSheet, xl_cells.cellrow, xl_cells.cellcol)
        Endscan
        llReturn = True

      Otherwise
        llReturn = False
      Endcase
      Set Order To
    Else
      llReturn = False
    Endif
  Endwith
  Return llReturn

*-- Inserts a new column into the sheet
  Procedure InsertColumn
  Lparameters tnWB, tnSheet, tnCol, tnShift
  Local llReturn
  If Pcount() < 4
    Return False
  Endif
  If This.GetSheetRecord(tnWB, tnSheet)
    Select xl_cells
    Set Order To Tag reverse
    Do Case
    Case tnShift = INSERT_LEFT
      Scan For workbook = tnWB And Sheet = tnSheet And cellcol >= tnCol
        Replace xl_cells.cellcol With xl_cells.cellcol + 1 In xl_cells
        If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(xl_cells.cellrow), "xl_rows", "id") And xl_cells.cellcol > xl_rows.maxcol
          Replace xl_rows.maxcol With xl_cells.cellcol In xl_rows
        Endif
      Endscan
      llReturn = True

    Case tnShift = INSERT_RIGHT
      Scan For workbook = tnWB And Sheet = tnSheet And cellcol > tnCol
        Replace xl_cells.cellcol With xl_cells.cellcol + 1 In xl_cells
        If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(xl_cells.cellrow), "xl_rows", "id") And xl_cells.cellcol > xl_rows.maxcol
          Replace xl_rows.maxcol With xl_cells.cellcol In xl_rows
        Endif
      Endscan
      llReturn = True

    Otherwise
      llReturn = False
    Endcase
    Set Order To
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Inserts a new row into the sheet
  Procedure InsertRow
  Lparameters tnWB, tnSheet, tnRow, tnShift
  Local llReturn, lnRow
  If Pcount() < 4
    Return False
  Endif
  If This.GetSheetRecord(tnWB, tnSheet)
    Select xl_cells
    Set Order To Tag reverse
    Do Case
    Case tnShift = INSERT_BEFORE
      Scan For workbook = tnWB And Sheet = tnSheet And cellrow >= tnRow
        lnRow = xl_cells.cellrow
        Replace xl_cells.cellrow With lnRow + 1 In xl_cells
        If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(lnRow), "xl_rows", "id")
          Replace xl_rows.Row With xl_cells.cellrow In xl_rows
        Endif
      Endscan
      llReturn = True

    Case tnShift = INSERT_AFTER
      Scan For workbook = tnWB And Sheet = tnSheet And cellrow > tnRow
        lnRow = xl_cells.cellrow
        Replace xl_cells.cellrow With lnRow + 1 In xl_cells
        If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(lnRow), "xl_rows", "id")
          Replace xl_rows.Row With xl_cells.cellrow In xl_rows
        Endif
      Endscan
      llReturn = True

    Otherwise
      llReturn =  False
    Endcase
    Set Order To
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Returns True if cell contains a formula
  Procedure IsCellFormula
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
  Local llFormula
  If Pcount() < 4
    Return Null
  Endif
  llFormula = False
  If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
    llFormula = !Empty(xl_cells.cellformula)
  Endif
  Return llFormula

*-- Checks if the cell reference is invalid (outside of limits)
  Protected Procedure IsCellReferenceInvalid
    Lparameters tnRow, tnCol
    Return tnRow > LIMITS_MAX_ROWS Or tnCol > LIMITS_MAX_COLUMNS

*-- Determines if the format is defined as a style
  Procedure IsFormatStyleDefined
  Lparameters tnWB, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos
  Do Case
  Case Pcount() < 3
    Return Null

  Case Pcount() = 3
    tlBold     = False
    tlItalic   = False
    tnFColor   = Rgb(0, 0, 0)
    tcULine    = UNDERLINE_NONE
    tlStrikThr = False
    tcVPos     = FONT_VERTICAL_BASELINE

  Case Pcount() = 4
    tlItalic   = False
    tnFColor   = Rgb(0, 0, 0)
    tcULine    = UNDERLINE_NONE
    tlStrikThr = False
    tcVPos     = FONT_VERTICAL_BASELINE

  Case Pcount() = 5
    tnFColor   = Rgb(0, 0, 0)
    tcULine    = UNDERLINE_NONE
    tlStrikThr = False
    tcVPos     = FONT_VERTICAL_BASELINE

  Case Pcount() = 6
    tcULine    = UNDERLINE_NONE
    tlStrikThr = False
    tcVPos     = FONT_VERTICAL_BASELINE

  Case Pcount() = 7
    tlStrikThr = False
    tcVPos     = FONT_VERTICAL_BASELINE

  Case Pcount() = 8
    tcVPos     = FONT_VERTICAL_BASELINE
  Endcase

  If Seek(BinToC(tnWB)+Padr(tcFName, 100) + Str(tnFSize, 5, 1) + Transform(tlBold) + Transform(tlItalic) + Padl(tnFColor, 15) + Padr(tcULine, 16) + ;
      TRANSFORM(tlStrikThr) + Padr(tcVPos, 11), "xl_fonts", "cellformat")

    If Seek(BinToC(tnWB)+BinToC(xl_fonts.Id)+BinToC(0)+BinToC(0)+BinToC(0)+BinToC(0)+BinToC(0), "xl_cellxfs", "cellformat")
      Return xl_cellxfs.Id
    Else
      Return Null
    Endif
  Else
    Return Null
  Endif

*-- Validates the border style value
  Protected Procedure IsValidBorderStyle
    Lparameters tcStyle
    If Vartype(tcStyle) != 'C' Or Empty(tcStyle) Or !Inlist(tcStyle, BORDER_STYLE_THIN, BORDER_STYLE_HAIR, BORDER_STYLE_DOTTED, BORDER_STYLE_DASHDOTDOT, ;
        BORDER_STYLE_DASHDOT, BORDER_STYLE_DASHED, BORDER_STYLE_THIN, BORDER_STYLE_MEDIUMDASHDOTDOT, ;
        BORDER_STYLE_SLANTDASHDOT, BORDER_STYLE_MEDIUMDASHDOT, BORDER_STYLE_MEDIUMDASHED, ;
        BORDER_STYLE_MEDIUM, BORDER_STYLE_THICK, BORDER_STYLE_DOUBLE)
      Return False
    Endif
    Return True

*-- Merges cells together
  Procedure MergeCells
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol
  Local lnRow, lnCol, lcCellValue, lnStringIndex, lcDataType, lcCellFormula, llReturn
  If Pcount() < 6
    Return False
  Endif
  If This.GetSheetRecord(tnWB, tnSheet)
    If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnBegRow)+BinToC(tnBegCol), "xl_mergecells", "mergecell")
      Replace xl_mergecells.endrow With tnEndRow, ;
        xl_mergecells.endcol With tnEndCol In xl_mergecells
    Else
*-*		Check if a cell being merged is already in a merge group
      Select xl_mergecells
      Locate For workbook = tnWB And Sheet = tnSheet And Between(tnBegRow, begrow, endrow) And Between(tnBegCol, begcol, endcol) And !Deleted()
      If Found()
        llReturn = False
      Else
        Locate For workbook = tnWB And Sheet = tnSheet And Between(tnEndRow, begrow, endrow) And Between(tnEndCol, begcol, endcol) And !Deleted()
        If Found()
          llReturn = False
        Else
*-*				Add merged cells
          Insert Into xl_mergecells (workbook, Sheet, begrow, begcol, endrow, endcol) ;
            VALUES (tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol)
*-*				Get upper left most cell value -> will be value of the merged cells (clear all other cell values)
          Select xl_cells
          lcCellValue = Null
          For lnRow=tnBegRow To tnEndRow
            For lnCol=tnBegCol To tnEndCol
              If This.GetCellRecord(tnWB, tnSheet, lnRow, lnCol) And !Empty(xl_cells.cellvalue)
                If Isnull(lcCellValue) And !xl_cells.celldeleted
                  lcDataType    = xl_cells.DataType
                  lcCellValue   = xl_cells.cellvalue
                  lcCellFormula = xl_cells.cellformula
                  lnStringIndex = xl_cells.stringid
                Endif
                Replace xl_cells.cellvalue   With "", ;
                  xl_cells.cellformula With "", ;
                  xl_cells.stringid    With 0, ;
                  xl_cells.DataType    With "" In xl_cells
              Endif
            Endfor
          Endfor
          If !Isnull(lcCellValue)
            If This.GetCellRecord(tnWB, tnSheet, tnBegRow, tnBegCol)
              Replace xl_cells.cellvalue   With lcCellValue, ;
                xl_cells.cellformula With lcCellFormula, ;
                xl_cells.stringid    With lnStringIndex, ;
                xl_cells.DataType    With lcDataType, ;
                xl_cells.celldeleted With False In xl_cells
            Endif
          Endif
*-*				If the tnEndCol cell is not a cell already in the row then add the cell to the cursor for the max cells by row determination
*				SELECT xl_cells
*				FOR lnRow=tnBegRow TO tnEndRow
*					LOCATE FOR workbook = tnWB AND sheet = tnSheet AND cellrow = lnRow AND cellcol >= tnEndCol AND !DELETED()
*					IF !FOUND()
*						INSERT INTO xl_cells (workbook, sheet, cellrow, cellcol, datatype, cellxfs, numdec, celldeleted) ;
*							VALUES (tnWB, tnSheet, lnRow, tnEndCol, DATA_TYPE_NONE, 0, -1, False)
*					ENDIF
*				ENDFOR
          llReturn = True
        Endif
      Endif
    Endif
    For lnRow=tnBegRow To tnEndRow
      This.SetRowMaxColumn(tnWB, tnSheet, lnRow, tnEndCol)
    Endfor
  Else
    llReturn = False
  Endif
  Return llReturn

*-- Called by the Destroy() event - user implementation
  Procedure OnDestroy
*-*
*-*	TODO: Add your code here
*-*

*-- User code for Init() Event
  Procedure OnInit
*-*
*-*	TODO: Add your code here
*-*

*-- Event for displaying an error message to the user
  Procedure OnShowErrorMessage
  Lparameters tnErrorId, tcErrMessage
*-*
*-*	TODO: Use BINDEVENTS to bind to this event and display a message to your users
*-*

*-- Event for displaying a wait message to the user
  Procedure OnShowStatusMessage
  Lparameters tnMode, tnStage, tnTotStages
*-*
*-*	TODO: Use BINDEVENTS to bind to this event and display a message to your users
*-*
*-*	tnTotStages = total number of tnStages (only passed on first call)
*-*
*-*	tnMode = 1; opening an xlsx file
*-* tnStage = 0; start of open
*-* tnStage = 1; reading shared strings XML
*-* tnStage = 2; reading styles XML
*-* tnStage = 3; reading relationships XML
*-* tnStage = 4; reading sheets XML
*-* tnStage = 5; reading named ranges
*-* tnStage = 6; reading external references
*-* tnStage = -1; end of open
*-*
*-*	tnMode = 2; saving an xlsx file
*-* tnStage = 0; start of save
*-* tnStage = 1; indicates saving supporting XMLs
*-* tnStage = 2; indicates saving strings XML
*-* tnStage = 3; indicates saving styles XML
*-* tnStage = 4; indicates saving workbook
*-* tnStage = 5; indicates saving relationship XML
*-* tnStage = 6+; indicates saving sheets
*-* tnStage = -1; end of close
*-*
*-*	tnMode = 3; saving an xlsx file
*-* tnStage = 0; start of write of data to cell values
*-* tnStage = 1-n; indicates saving cell values
*-* tnStage = -1; write of data
*-*

*-- Opens the XLSX file via Win32 API default program
  Procedure OpenCreatedXlsxFile
  Lparameters txWB
  Local lcFileName
  Do Case
  Case Pcount() = 0
    Return False

  Case Vartype(txWB) = "N"
    lcFileName = This.GetWorkbookFileName(txWB)
    If Empty(lcFileName) Or !File(lcFileName)
      Return False
    Endif

  Case Vartype(txWB) = "C"
    If File(txWB)
      lcFileName = txWB
    Else
      Return False
    Endif

  Otherwise
    Return False
  Endcase
  Try
    apiShellExecute(0, "Open", lcFileName, "", "", 5)

  Catch To loException
    Raiseevent(This, "OnShowErrorMessage", 99, "Failure to open file: " + loException.Message)
  Endtry

*-- Opens the passed Xlsx workbook and loads the internal cursors with the content
  Procedure OpenXlsxWorkbook
  Lparameters tcFileName, tlForceTextFormat
  Local lnWB, lcBaseName, lcTempPath, lcZipName, loShell, lnCnt, lnCntTot, lcRelationsXml, lnNode
  Local lnRelId, lcRelType, lcTarget, lcWbXmlPath, lcRelation, lcWorkbook, lcSheet, lcShName, lnShRId, lcShFile
  Local loException, loFS, lcDefinedNames, lcDefinedName, lcRName, lcComment, lnRangeId, lcRange, lnNdx
  Local lnBegRow, lnBegCol, lnEndRow, lnEndCol, lcExternalRefs, lnExRId, lcExtRef, lcCellRange, lcBegRange
  Local lcEndRange, loDir, lnLastRelId, lcSheets, lcRelationships, lcShState
  Local Array laTemp[1]
  With This
    Do Case
    Case Pcount() = 0
      .ErrorLevelId = 1
      Raiseevent(This, "OnShowErrorMessage", 1, "Invalid call to OpenXlsxWorkbook() method; must include file name to open")

    Case Pcount() = 1
      tlForceTextFormat = False
    Endcase
    lnWB = 0
    If Adir(laTemp, tcFileName) > 0
*-*	Open the Workbook as an archive (zip) file and extract contents
      Try
        lcBaseName = Sys(2015)
        lcTempPath = Addbs(Sys(2023)) + lcBaseName
        lcZipName  = lcTempPath + ".zip"
        Copy File (tcFileName) To (lcZipName)
        loShell = Createobject("shell.application")
        Mkdir (lcTempPath)
        loShell.NameSpace(lcTempPath).CopyHere(loShell.NameSpace(lcZipName).Items, FOF_SILENT)
        lnCnt = 0
        lnCntTot = loShell.NameSpace(lcZipName).Items.Count
        Do While loShell.NameSpace(lcTempPath).Items.Count != lnCntTot
          lnCnt = lnCnt + 1
          apiSleep(100)
          If lnCnt > lnCntTot
            Set Step On
          Endif
        Enddo
        Erase (lcZipName)

      Catch To loException
        Set Step On
        .ErrorLevelId = 2
        Raiseevent(This, "OnShowErrorMessage", 2, loException.Message)
        lnWB = Null
      Endtry
      If Isnull(lnWB)
        Return 0
      Endif
*-*	Add to the workbook table
      lnWB = .createworkbookex(tcFileName)
      If lnWB > 0
*-*		Verify all required XML files are present
        lcWbXmlPath = Addbs(lcTempPath)
        If Adir(laTemp, lcWbXmlPath + "xl\workbook.xml") = 0
          .ErrorLevelId = 3
          Raiseevent(This, "OnShowErrorMessage", 3, "Invalid XLSX file - missing workbook.xml")
          .DeleteWorkbook(lnWB)
          Return 0
        Endif
        If Adir(laTemp, lcWbXmlPath + "xl\_rels\workbook.xml.rels") = 0
          .ErrorLevelId = 4
          Raiseevent(This, "OnShowErrorMessage", 4, "Invalid XLSX file - missing workbook.xml.rels")
          .DeleteWorkbook(lnWB)
          Return 0
        Endif
        If Adir(laTemp, lcWbXmlPath + "xl\styles.xml") = 0
          .ErrorLevelId = 5
          Raiseevent(This, "OnShowErrorMessage", 5, "Invalid XLSX file - missing styles.xml")
          .DeleteWorkbook(lnWB)
          Return 0
        Endif

*-*		Process the workbook
        Raiseevent(This, "OnShowStatusMessage", 1, 1, 6)
        Try
*-*			Get the shared strings if present
          If Adir(laTemp, lcWbXmlPath + "xl\sharedStrings.xml") > 0
            If !This.ReadSharedStringsXML(lcWbXmlPath, lnWB)
              .ErrorLevelId = 7
              Raiseevent(This, "OnShowErrorMessage", 7, "Invalid XLSX file - error during shared string loading")
              Throw
            Endif
          Endif

*-*			Get the cell formatting styles
          .ReadStylesXML(lcWbXmlPath, lnWB)

*-*			Get the defined relationship files for this workbook
          Raiseevent(This, "OnShowStatusMessage", 1, 3)
          lcRelationsXml = Filetostr(lcWbXmlPath + "xl\_rels\workbook.xml.rels")
          lcRelationships = .GetNodeElement(lcRelationsXml, 'Relationships', 1, False, 1)

          lnNode = 1
          lnLastRelId = 0
          lcRelation = .GetNodeElement(lcRelationships, 'Relationship', lnNode)
          Do While !Empty(lcRelation)
            lnRelId   = Cast(Substr(.GetNodeAttributeValue(lcRelation, "Id"), 4) As I)
            lcTarget  = .GetNodeAttributeValue(lcRelation, "Target")
            lcRelType = .GetNodeAttributeValue(lcRelation, "Type")
            lcRelType = Substr(lcRelType, Ratc("/", lcRelType)+1)
            If lcRelType != "theme"
              Insert Into xl_relationships (workbook, relid, reltype, Target) Values (lnWB, lnRelId, lcRelType, lcTarget)

              If lnRelId > lnLastRelId
                lnLastRelId = lnRelId
              Endif
            Endif
            lnNode = lnNode + 1
            lcRelation = .GetNodeElement(lcRelationships, 'Relationship', lnNode)
          Enddo
          .SetLastId(lnWB, lnLastRelId, "xl_relationships")

*-*			Get the sheets
          Raiseevent(This, "OnShowStatusMessage", 1, 4)
          lcWorkbook = Filetostr(lcWbXmlPath + "xl\workbook.xml")

          lcSheets   = .GetNodeElement(lcWorkbook, 'sheets', 1, False, 1)
          lnNode     = 1
          lcSheet    = .GetNodeElement(lcSheets, 'sheet', lnNode)
          Do While !Empty(lcSheet)
            lcShName  = .GetNodeAttributeValue(lcSheet, "name")
            lnShRId   = Cast(Substr(.GetNodeAttributeValue(lcSheet, "r:id"), 4) As I)
            lcShFile  = lcWbXmlPath + "xl\" + .GetRelationShipFileName(lnWB, lnShRId)
            lcShState = .GetNodeAttributeValue(lcSheet, "state")
            Do Case
            Case lcShState = "visible"
              lnShState = VISIBLE_SHEET_STATE

            Case lcShState = "hidden"
              lnShState = HIDDEN_SHEET_STATE

            Case lcShState = "veryHidden"
              lnShState = VERYHIDDEN_SHEET_STATE

            Otherwise
              lnShState = VISIBLE_SHEET_STATE
            Endcase

            If Adir(laTemp, lcShFile) = 0 Or !.ReadSheetXML(lcShFile, lnWB, lnNode, lcShName, lnShState, tlForceTextFormat)
              .ErrorLevelId = 8
              Raiseevent(This, "OnShowErrorMessage", 8, "Invalid XLSX file - missing sheet or invalid sheet: " + lcShFile)
              .DeleteWorkbook(lnWB)
              Throw
            Endif
            If Seek(BinToC(lnWB)+BinToC(lnShRId), "xl_relationships", "relid")
              Replace xl_relationships.Sheet With lnNode In xl_relationships
            Else
              Set Step On
            Endif

            lnNode = lnNode + 1
            lcSheet = .GetNodeElement(lcSheets, 'sheet', lnNode)
          Enddo
          .SetLastId(lnWB, lnNode-1, "xl_sheets")

*-*			Get the defined name ranges
          Raiseevent(This, "OnShowStatusMessage", 1, 5)
          lcDefinedNames = .GetNodeElement(lcWorkbook, 'definedNames', 1)
          If !Empty(lcDefinedNames)
            lnNode  = 1
            lcDefinedName = .GetNodeElement(lcDefinedNames, 'definedName', lnNode)
            Do While !Empty(lcDefinedName)
              lcRange = .GetNodeElementValue(lcDefinedName)
              If !Empty(lcRange)
                lnNdx = Atc("!", lcRange)
                lcShName = Left(lcRange, lnNdx-1)
                If !Empty(lcShName)
                  lcShName = Chrtran(lcShName, "'", "")
                  If Seek(BinToC(lnWB)+Upper(Padr(lcShName, Len(xl_sheets.shname))), "xl_sheets", "shname")

                    lcRName   = .GetNodeAttributeValue(lcDefinedName, "name")
                    lcComment = .GetNodeAttributeValue(lcDefinedName, "comment")
                    lnRangeId = Cast(.GetNodeAttributeValue(lcDefinedName, "localSheetId") As I)

                    lcCellRange = Substr(lcRange, lnNdx+1)
                    lcCellRange = Chrtran(lcCellRange, "$", "")
                    lcBegRange  = Getwordnum(lcCellRange, 1, ":")
                    lcEndRange  = Getwordnum(lcCellRange, 2, ":")

                    lnNdx = .GetAlphaNumericSplit(lcBegRange)
                    Do Case
                    Case lnNdx = 0                  && Column only range
                      lnBegCol = This.ColumnAsciiToIndex(lcBegRange)
                      lnBegRow = 0

                    Case lnNdx = 1                  && Row only range
                      lnBegCol = 0
                      lnBegRow = Cast(lcBegRange As I)

                    Otherwise
                      lnBegCol = .ColumnAsciiToIndex(Left(lcBegRange, lnNdx-1))
                      lnBegRow = Cast(Substr(lcBegRange, lnNdx) As I)
                    Endcase

                    lnNdx = .GetAlphaNumericSplit(lcEndRange)
                    Do Case
                    Case lnNdx = 0                  && Column only range
                      lnEndCol = .ColumnAsciiToIndex(lcEndRange)
                      lnEndRow = 0

                    Case lnNdx = 1                  && Row only range
                      lnEndCol = 0
                      lnEndRow = Cast(lcEndRange As I)

                    Otherwise
                      lnEndCol = .ColumnAsciiToIndex(Left(lcEndRange, lnNdx-1))
                      lnEndRow = Cast(Substr(lcEndRange, lnNdx) As I)
                    Endcase

                    Insert Into xl_namerange (workbook, Sheet, rname, scope, Comment, begrow, begcol, endrow, endcol) ;
                      VALUES (lnWB, xl_sheets.Sheet, lcRName, lnRangeId, lcComment, lnBegRow, lnBegCol, lnEndRow, lnEndCol)
                  Endif
                Endif
              Endif
              lnNode = lnNode + 1
              lcDefinedName = .GetNodeElement(lcDefinedNames, 'definedName', lnNode)
            Enddo
          Endif

*-*			Get the external references
          Raiseevent(This, "OnShowStatusMessage", 1, 6)
*			lcExternalRefs = this.GetNodeElement(lcWorkbook, 'externalReferences', 1)
*			IF !EMPTY(lcExternalRefs)
*				lnNode   = 1
*				lcExtRef = this.GetNodeElement(lcExternalRefs, 'externalReference', lnNode)
*				DO WHILE !EMPTY(lcExtRef)
*					lnExRId = CAST(SUBSTR(this.GetNodeAttributeValue(lcExtRef, "r:id"), 4) AS I)
*					lcFile  = lcWbXmlPath + "xl\" + this.GetRelationshipFileName(lnWB, lnExRId)
*					this.ReadExternalRefXML(lnWB, lcFile, lnExRId)
*
*					lnNode = lnNode + 1
*					lcExtRef = this.GetNodeElement(lcExternalRefs, 'externalReference', lnNode)
*				ENDDO
*			ENDIF

        Catch To loException
          Set Step On
          .ErrorLevelId = 9
          Raiseevent(This, "OnShowErrorMessage", 9, "Invalid XLSX file - error reading data; " + loException.Message)
          .DeleteWorkbook(lnWB)
          lnWB = 0

        Finally
          Try
            loFS = Createobject("Scripting.FileSystemObject")
            loDir = loFS.GetFolder(lcTempPath)
            loDir.Delete()

          Catch To loException
            Set Step On
          Endtry
          Raiseevent(This, "OnShowStatusMessage", 1, -1)
        Endtry
        If .Debug
          .DebugCursorsToFile("Extract")
        Endif
      Endif
    Endif
  Endwith
  Return lnWB

*-- Reads the cell value and format for a sheet
  Protected Procedure ReadCellValueFormat
    Lparameters tnWB, tcColumn, tnCellCol
    Local lcCellType, lnNode, lcBaseName, lcFormatCode, lnNdx, loCellFormat, lcWholeNbr, lcDecNbr
    Local lcPosFormat, lcNegFormat, lcZerFormat, lcInLineTxt
    loCellFormat = Createobject("Empty")
    AddProperty(loCellFormat, "CellVal", "")
    AddProperty(loCellFormat, "DataType", DATA_TYPE_NONE)
    AddProperty(loCellFormat, "Formula", "")
    AddProperty(loCellFormat, "StringNdx", 0)
    AddProperty(loCellFormat, "CellXfsId", 0)

    With This
      lcCellType = .GetNodeAttributeValue(tcColumn, "t")
      For lnChild=1 To .GetChildNodeCount(tcColumn)
        lcChild = .GetChildNodeElement(tcColumn, lnChild)
        Do Case
        Case lcChild = "<f"                                   && Formula attribute - cell contains a formula
          lcCellType = "f"
          loCellFormat.formula = .GetCellFormulaFromXml(lcChild, tnCellCol)

        Case lcChild = "<v"                                   && Value assigned to cell
          loCellFormat.CellVal = .GetNodeElementValue(lcChild)

        Case lcChild = "<is"                                  && Value assigned to cell as inline text
          lcInLineTxt = .GetNodeElement(lcChild, 't', 1)
          loCellFormat.CellVal = .GetNodeElementValue(lcInLineTxt)
        Endcase
      Endfor

      loCellFormat.CellXfsId = Cast(.GetNodeAttributeValue(tcColumn, "s") As I)
      Do Case
      Case lcCellType == "f"
        loCellFormat.DataType = DATA_TYPE_FORMULA

      Case lcCellType == "s" Or lcCellType == "str"
        loCellFormat.StringNdx = Cast(loCellFormat.CellVal As I)
        loCellFormat.CellVal   = "StringValue"
        loCellFormat.DataType  = DATA_TYPE_CHAR

      Case lcCellType == "inlineStr"
        loCellFormat.StringNdx = .AddStringValue(tnWB, loCellFormat.CellVal, False)
        loCellFormat.CellVal   = "StringValue"
        loCellFormat.DataType  = DATA_TYPE_CHAR

      Case Empty(loCellFormat.CellVal)
        Return loCellFormat

      Otherwise
*-*		Determine the data type of the value based on assigned format
        If Seek(BinToC(tnWB)+BinToC(loCellFormat.CellXfsId), "xl_cellxfs", "id")
          Do Case
          Case Inlist(xl_cellxfs.numfmtid, 1, 3, 37, 38)                           && Integer format
            loCellFormat.DataType = DATA_TYPE_INT

          Case Inlist(xl_cellxfs.numfmtid, 2, 4, 7, 8, 11, 12, 13, 39, 40)         && Float format
            loCellFormat.DataType = DATA_TYPE_FLOAT

          Case Inlist(xl_cellxfs.numfmtid, 9, 10)                                  && Numeric percent format
            loCellFormat.DataType = DATA_TYPE_FLOAT

          Case Inlist(xl_cellxfs.numfmtid, 14, 15, 16, 17)                         && Date format
            Do Case
            Case Atc(".", loCellFormat.CellVal) = 0 Or Right(loCellFormat.CellVal, 1) = "."
              loCellFormat.DataType = DATA_TYPE_DATE

            Case Atc(".", loCellFormat.CellVal) > 0
              lnNdx = Atc(".", loCellFormat.CellVal)
              lcWholeNbr = Left(loCellFormat.CellVal, lnNdx-1)
              lcDecNbr   = Substr(loCellFormat.CellVal, lnNdx+1)
              If Len(lcDecNbr) > 0
                loCellFormat.CellVal  = lcWholeNbr + "." + Padr(lcDecNbr, 17, "0")
                loCellFormat.DataType = DATA_TYPE_DATETIME
              Else
                loCellFormat.CellVal  = lcWholeNbr
                loCellFormat.DataType = DATA_TYPE_DATE
              Endif
            Endcase

          Case Inlist(xl_cellxfs.numfmtid, 18, 19, 20, 21, 45, 46)                 && Time format
            Do Case
            Case Atc("E-", loCellFormat.CellVal) > 0
              loCellFormat.CellVal  = Padr(Transform(Eval(loCellFormat.CellVal)), 19, "0")
              loCellFormat.DataType = DATA_TYPE_TIME

            Case Atc("E", loCellFormat.CellVal) > 0
              loCellFormat.CellVal = Transform(Eval(loCellFormat.CellVal))
              lnNdx = Atc(".", loCellFormat.CellVal)
              lcWholeNbr = Left(loCellFormat.CellVal, lnNdx-1)
              lcDecNbr   = Substr(loCellFormat.CellVal, lnNdx+1)
              loCellFormat.CellVal  = lcWholeNbr + "." + Padr(lcDecNbr, 17, "0")
              loCellFormat.DataType = DATA_TYPE_DATETIME

            Case Left(loCellFormat.CellVal, 2) = "0."
              lcDecNbr = Substr(loCellFormat.CellVal, 3)
              loCellFormat.CellVal  = "0." + Padr(lcDecNbr, 17, "0")
              loCellFormat.DataType = DATA_TYPE_TIME

            Otherwise
              lnNdx = Atc(".", loCellFormat.CellVal)
              lcWholeNbr = Left(loCellFormat.CellVal, lnNdx-1)
              lcDecNbr   = Substr(loCellFormat.CellVal, lnNdx+1)
              loCellFormat.CellVal  = lcWholeNbr + "." + Padr(lcDecNbr, 17, "0")
              loCellFormat.DataType = DATA_TYPE_DATETIME
            Endcase

          Case Inlist(xl_cellxfs.numfmtid, 22, 29, 30, 31, 32, 33, 34)             && Date-time format
            lnNdx = Atc(".", loCellFormat.CellVal)
            lcWholeNbr = Iif(lnNdx=0, loCellFormat.CellVal, Left(loCellFormat.CellVal, lnNdx-1))     && Bug fix provide by Dan Goodwin from VFPx
            lcDecNbr   = Iif(lnNdx=0, "", Substr(loCellFormat.CellVal, lnNdx+1))                     && Bug fix provide by Dan Goodwin from VFPx
            loCellFormat.CellVal  = lcWholeNbr + "." + Padr(lcDecNbr, 17, "0")
            loCellFormat.DataType = DATA_TYPE_DATETIME

          Case xl_cellxfs.numfmtid = 49                                            && Numeric formatted as text
            loCellFormat.StringNdx = This.AddStringValue(tnWB, loCellFormat.CellVal, False)
            loCellFormat.CellVal   = "StringValue"
            loCellFormat.DataType  = DATA_TYPE_CHAR

          Case Seek(BinToC(tnWB)+BinToC(xl_cellxfs.numfmtid), "xl_numfmts", "id")  && Determine based on custom format
            lcFormatCode = Alltrim(xl_numFmts.formatcode)
            Do Case
            Case Atc("h:m", lcFormatCode) > 0 Or Atc("m:s", lcFormatCode) > 0 Or Atc("AM/PM", lcFormatCode) > 0 Or Atc("A/P", lcFormatCode) > 0
              loCellFormat.DataType = DATA_TYPE_DATETIME

            Case Atc("yy", lcFormatCode) > 0 Or Atc("d\-m", lcFormatCode) > 0
              loCellFormat.DataType = DATA_TYPE_DATE

            Otherwise
              lnNdx = Atc(".", lcFormatCode)
              If lnNdx > 0
                loCellFormat.DataType = DATA_TYPE_FLOAT
                lnNdx = lnNdx + 1
                Do While lnNdx < Len(lcFormatCode)
                  If Substr(lcFormatCode, lnNdx, 1) != "0"
                    Exit
                  Endif
                  lnNdx = lnNdx + 1
                Enddo
              Else
                loCellFormat.DataType = DATA_TYPE_INT
              Endif
            Endcase

          Otherwise
            Do Case
            Case Isnull(loCellFormat.CellVal) Or Empty(loCellFormat.CellVal)
              loCellFormat.DataType = DATA_TYPE_NONE

            Case Atc(".", loCellFormat.CellVal) > 0
              loCellFormat.DataType = DATA_TYPE_FLOAT

            Otherwise
              loCellFormat.DataType = DATA_TYPE_INT
            Endcase
          Endcase
        Else
          Do Case
          Case Isnull(loCellFormat.CellVal) Or Empty(loCellFormat.CellVal)
            loCellFormat.CellVal  = ""
            loCellFormat.DataType = DATA_TYPE_NONE

          Case Atc(".", loCellFormat.CellVal) > 0
            loCellFormat.DataType = DATA_TYPE_FLOAT

          Otherwise
            loCellFormat.DataType = DATA_TYPE_INT
          Endcase
        Endif
      Endcase
    Endwith
    Return loCellFormat

*-- Reads the External References file
  Protected Procedure ReadExternalRefXml
    Lparameters tnWB, tcFileName, tnRelId
    Local lcExternalLink, lcSheetNames, lnNode, lcSheetName, lcShName

*-*	Open the external file for processing
    lcExternalLink = Filetostr(tcFileName)

    With This
*-*	Get the sheet names
      lcSheetNames = .GetNodeElement(lcExternalLink, 'sheetNames', 1)
      If !Empty(lcSheetNames)
        lnNode = 1
        lcSheetName = .GetNodeElement(lcSheetNames, 'sheetName', lnNode)
        Do While !Empty(lcSheetName)
          lcShName = .GetNodeAttributeValue(lcSheetName, "val")

          Insert Into xl_extsheets (workbook, extid, relid, sheetname) Values (tnWB, lnNode, tnRelId, lcShName)

          lnNode = lnNode + 1
          lcSheetName = .GetNodeElement(lcSheetNames, 'sheetName', lnNode)
        Enddo

*-*	Get each external sheet detail
        lnNode = 1
        lcSheetData = .GetNodeElement(lcExternalLink, 'sheetData', lnNode)
        Do While !Empty(lcSheetData)
          lnRow = 1
          lcRowData = .GetNodeElement(lcSheetData, 'row', lnRow)
          Do While !Empty(lcRowData)

          Enddo

          lnNode = lnNode + 1
          lcSheetData = .GetNodeElement(lcExternalLink, 'sheetData', lnNode)
        Enddo
      Endif
    Endwith

*-- Reads the sharedstrings.xml for opening workbooks
  Protected Procedure ReadSharedStringsXML
    Lparameters tcFilePath, tnWB
    Local lcStrings, lcString, lnNode, lcRichTextRun, lcText, lnIndex, llFBold, llFItalic, loException, lcT
    Local lnFColor, lcFName, lnFSize, lcULine, llStrkThr, lcFVPos, lnTheme, lnTint, lnIndexed, lcRunPrp
    Local lnR, lcRText, lcRunProp, lcColor, llPrsvSp, llFormatted, llSuccess, lhFile, lcSharedStringFile, lnSize
    Local lcStringXml, lcCheckSum

    With This
      Raiseevent(This, "OnShowStatusMessage", 1, 1)
      Try
        lcSharedStringFile = Addbs(tcFilePath) + "xl\sharedStrings.xml"
        lcStrings = Filetostr(lcSharedStringFile)
        If .Codepage = 0
          lcStrings = Strconv(lcStrings, 11)
        Else
          lcStrings = Strconv(lcStrings, 11, .Codepage, 1)
        Endif

        lnNode    = 1
        lcString  = .GetNodeElement(lcStrings, 'si', lnNode)
        Do While !Empty(lcString)
          lnR = 1
          lcRichTextRun = .GetNodeElement(lcString, 'r', lnR, True)
          If Empty(lcRichTextRun)
            llFormatted = False
            lcT      = .GetNodeElement(lcString, 't', 1)
            lcText   = .GetNodeElementValue(lcT)
            llPrsvSp = Iif(Empty(.GetNodeAttributeValue(lcT, "xml:space")), False, True)

            If llPrsvSp
              lcRText = Chrtran(lcText, Chr(32), Chr(160))
            Endif
          Else
            llFormatted = True
            lcText   = ""
            lnIndex  = 0
            Do While !Empty(lcRichTextRun)
              lnIndex   = lnIndex + 1
              lcT       = .GetNodeElement(lcRichTextRun, 't', 1)
              llPrsvSp  = Iif(Empty(.GetNodeAttributeValue(lcT, "xml:space")), False, True)
              lcRText   = .GetNodeElementValue(lcT)
              lcText    = lcText + lcRText

              lcRunProp = .GetNodeElement(lcRichTextRun, 'rPr', 1, True)
              llFBold   = Iif(Atc("<b/>", lcRunProp)>0, True, False)
              llFItalic = Iif(Atc("<i/>", lcRunProp)>0, True, False)
              llStrkThr = Iif(Atc("<strike/>", lcRunProp)>0, True, False)
              If Atc("<u/>", lcRunProp) > 0
                lcULine = UNDERLINE_SINGLE
              Else
                lcULine = .GetNodeAttributeValue(.GetNodeElement(lcRunProp, 'u', 1), "val")
                lcULine = Iif(Empty(lcULine), UNDERLINE_NONE, lcULine)
              Endif

              lnFSize   = Cast(.GetNodeAttributeValue(.GetNodeElement(lcRunProp, 'sz', 1), "val") As I)
              lcFName   = .GetNodeAttributeValue(.GetNodeElement(lcRunProp, 'rFont', 1), "val")
              lcFVPos   = .GetNodeAttributeValue(.GetNodeElement(lcRunProp, 'vertAlign', 1), "val")

              lcColor   = .GetNodeElement(lcRunProp, 'color', 1)
              lnFColor  = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
              lnTheme   = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
              lnTint    = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
              lnIndexed = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)

              If llPrsvSp
                lcRText = Chrtran(lcRText, Chr(32), Chr(160))
              Endif

              Insert Into xl_strformat (workbook, Id, Index, stringxml, stringval, fbold, fitalic, fcolor, fname, Fsize, uline, strkthr, fvpos, Theme, tint, indexed, presvspace) ;
                VALUES (tnWB, lnNode-1, lnIndex, lcRText, .GetStringXml(lcRText), llFBold, llFItalic, lnFColor, lcFName, lnFSize, lcULine, llStrkThr, ;
                lcFVPos, lnTheme, lnTint, lnIndexed, llPrsvSp)

              lnR = lnR + 1
              lcRichTextRun = .GetNodeElement(lcString, 'r', lnR, True)
            Enddo
            llPrsvSp = False
          Endif
          lcStringXml = .GetXMLString(lcText)
          lcCheckSum  = .GetCheckSum(lcText)

          Insert Into xl_strings (Id, workbook, checksum, stringxml, stringval, presvspace, Formatted) ;
            VALUES (lnNode-1, tnWB, lcCheckSum, lcText, lcStringXml, llPrsvSp, llFormatted)

          lnNode = lnNode + 1
          lcString = .GetNodeElement(lcStrings, 'si', lnNode)
        Enddo
        .SetLastId(tnWB, lnNode-2, "xl_strings")    && Decrement by 2 to account for empty string not assigned
        llSuccess = True

      Catch To loException
        Set Step On
        llSuccess = False
      Endtry
    Endwith
    Return llSuccess

*-- Reads the sheet.xml for opening a workbook
  Protected Procedure ReadSheetXML
    Lparameters tcShFile, tnWB, tnSheet, tcShName, tnShState, tlForceTextFormat
    Local lcSheet, lcSheetViews, lcSheetView, lcPane, lcMargin, lnXSplit, lnYSplit, lnLeft, lcDimen, lnScale, lnFitToWidth, lnFitToHeight
    Local lnRight, lnTop, lnBottom, lnHeader, lnFooter, lcPageSetup, lnOrientation, lnPaperSize, lnWidth, lnHeight, lnNode, lcMergeCells
    Local lnCnt, lcMergeCell, lcCellRef, lnBegCol, lnBegRow, lnEndCol, lnEndRow, lcTemp, lnNdx, lcSheetData
    Local lcRow, lnRowHt, lnCellRow, lnCellCol, lnCol, loCell, lnTabColNdx, lcTabColRgb, lcTabColor, lcSheetPr
    Local lcValidations, lnValidType, lcValidation, lnValidStyle, lnOperator, lcErrMsg, lcErrTitle, lcPrompt, lcFormula1
    Local lcFormula2, llAllowBlank, llShowInputMsg, llShowErrMsg, llFormula, llCustHt, lcValidType, lcValidStyle
    Local lcValidSqRef, lnIndex

    With This
*-*	Open the sheet for processing
      lcSheet = Filetostr(tcShFile)
      If .Codepage = 0
        lcSheet = Strconv(lcSheet, 11)
      Else
        lcSheet = Strconv(lcSheet, 11, .Codepage, 1)
      Endif

*-*	Get the freeze rows/columns settings
      lcSheetViews = .GetNodeElement(lcSheet, 'sheetViews', 1, True)
      lcSheetView  = .GetNodeElement(lcSheetViews, 'sheetView', 1)
      lcPane       = .GetNodeElement(lcSheetView, 'pane', 1)
      lnXSplit     = Cast(.GetNodeAttributeValue(lcPane, "xSplit") As I)
      lnYSplit     = Cast(.GetNodeAttributeValue(lcPane, "ySplit") As I)

*-*	Get the tab color settings
      lcSheetPr = .GetNodeElement(lcSheet, 'sheetPr', 1)
      If Empty(lcSheetPr)
        lnTabColNdx = 0
        lcTabColRgb = ""
      Else
        lcTabColor  = .GetNodeElement(lcSheetPr, 'tabColor', 1)
        lnTabColNdx = Cast(.GetNodeAttributeValue(lcTabColor, "indexed") As I)
        lcTabColRgb = .GetNodeAttributeValue(lcTabColor, "rgb")
      Endif

*-*	Get the worksheet margins
      lcMargin = .GetNodeElement(lcSheet, 'pageMargins', 1)
      lnLeft   = Cast(.GetNodeAttributeValue(lcMargin, "left") As N(6,3))
      lnLeft   = Iif(lnLeft = 0.000, 0.75, lnLeft)
      lnRight  = Cast(.GetNodeAttributeValue(lcMargin, "right") As N(6,3))
      lnRight  = Iif(lnRight = 0.000, 0.75, lnRight)
      lnTop    = Cast(.GetNodeAttributeValue(lcMargin, "top") As N(6,3))
      lnTop    = Iif(lnTop = 0.000, 0.75, lnTop)
      lnBottom = Cast(.GetNodeAttributeValue(lcMargin, "bottom") As N(6,3))
      lnBottom = Iif(lnBottom = 0.000, 0.75, lnBottom)
      lnFooter = Cast(.GetNodeAttributeValue(lcMargin, "footer") As N(6,3))
      lnFooter = Iif(lnFooter = 0.000, 0.30, lnFooter)
      lnHeader = Cast(.GetNodeAttributeValue(lcMargin, "header") As N(6,3))
      lnHeader = Iif(lnHeader = 0.000, 0.30, lnHeader)

*-*	Get the page setup
      lcPageSetup   = .GetNodeElement(lcSheet, 'pageMargins', 1)
      lnOrientation = Iif(.GetNodeAttributeValue(lcPageSetup, "orientation")='portrait', PORTRAIT_PRINT_ORIENTATION, LANDSCAPE_PRINT_ORIENTATION)
      lnPaperSize   = Cast(.GetNodeAttributeValue(lcPageSetup, "paperSize") As I)
      lnPaperSize   = Iif(lnPaperSize=0, PAPERSIZE_LTR, lnPaperSize)
      lnWidth       = Cast(.GetNodeAttributeValue(lcPageSetup, "paperWidth") As I)
      lnHeight      = Cast(.GetNodeAttributeValue(lcPageSetup, "paperHeight") As I)
      lcDimen       = Right(.GetNodeAttributeValue(lcPageSetup, "paperHeight"), 2)
      lnScale       = Cast(.GetNodeAttributeValue(lcPageSetup, "scale") As I)
      lnScale       = Iif(lnScale=0, 100, lnScale)
      lnFitToWidth  = Cast(.GetNodeAttributeValue(lcPageSetup, "fitToWidth") As I)
      lnFitToHeight = Cast(.GetNodeAttributeValue(lcPageSetup, "fitToHeight") As I)

*-*	Save the worksheet to the cursor
      Insert Into xl_sheets (workbook, Sheet, shname, state, mleft, mright, mtop, mbot, mheader, mfooter, shdeleted, xsplit, ysplit, prnorient, ;
        papersize, Paperwidth, paperheight, paperdimen, Scale, fittowidth, fittoheight, tabcolorndx, tabcolorrgb) ;
        VALUES (tnWB, tnSheet, tcShName, tnShState, lnLeft, lnRight, lnTop, lnBottom, lnHeader, lnFooter, False, lnXSplit, lnYSplit, lnOrientation, ;
        lnPaperSize, lnWidth, lnHeight, lcDimen, lnScale, lnFitToWidth, lnFitToHeight, lnTabColNdx, lcTabColRgb)

*-*	Get the defined column widths
      lcCols = .GetNodeElement(lcSheet, 'cols', 1, True)
      lnNode = 1
      lcCol  = .GetNodeElement(lcCols, 'col', lnNode)
      Do While !Empty(lcCol)
        lnMinCol = Cast(.GetNodeAttributeValue(lcCol, "min") As I)
        lnMaxCol = Cast(.GetNodeAttributeValue(lcCol, "max") As I)
        lnWidth  = Cast(.GetNodeAttributeValue(lcCol, "width") As N(16,9)) - 0.71093750
        For lnCol=lnMinCol To lnMaxCol
          Insert Into xl_colwidths (workbook, Sheet, Column, Width) Values (tnWB, tnSheet, lnCol, lnWidth)
        Endfor
        lnNode = lnNode + 1
        lcCol  = .GetNodeElement(lcCols, 'col', lnNode)
      Enddo

*-*	Get the defined merged cells
      lcMergeCells = .GetNodeElement(lcSheet, 'mergeCells', 1)
      lnCnt = Cast(.GetNodeAttributeValue(lcMergeCells, "count") As I)
      For lnNode=1 To lnCnt
        lcMergeCell = .GetNodeElement(lcMergeCells, 'mergeCell', lnNode)
        If Empty(lcMergeCell)
          Set Step On
          Loop
        Endif
        lcCellRef = .GetNodeAttributeValue(lcMergeCell, "ref")
        lnBegCol = 0
        lnBegRow = 0
        lnEndCol = 0
        lnEndRow = 0
        lcTemp = Getwordnum(lcCellRef, 1, ":")
        lnNdx = 0
        Do While lnNdx < Len(lcTemp)
          lnNdx = lnNdx + 1
          If Isdigit(Substr(lcTemp, lnNdx, 1))
            lnBegCol = .ColumnAsciiToIndex(Left(lcTemp, lnNdx-1))
            lnBegRow = Int(Val(Substr(lcTemp, lnNdx)))
            Exit
          Endif
        Enddo
        If lnBegCol > 0
          lcTemp = Getwordnum(lcCellRef, 2, ":")
          lnNdx = 0
          Do While lnNdx < Len(lcTemp)
            lnNdx = lnNdx + 1
            If Isdigit(Substr(lcTemp, lnNdx, 1))
              lnEndCol = .ColumnAsciiToIndex(Left(lcTemp, lnNdx-1))
              lnEndRow = Int(Val(Substr(lcTemp, lnNdx)))
              Exit
            Endif
          Enddo
          If lnEndCol > 0
            Insert Into xl_mergecells (workbook, Sheet, begrow, begcol, endrow, endcol) ;
              VALUES (tnWB, tnSheet, lnBegRow, lnBegCol, lnEndRow, lnEndCol)
          Endif
        Endif
      Endfor

*-*	Get the sheet cell values and formatting
      lcSheetData = .GetNodeElement(lcSheet, 'sheetData', 1, True)
      lnNode = 1
      lcRow = .GetNodeElement(lcSheetData, 'row', lnNode)
      Do While !Empty(lcRow)                                                   && Process each row in the sheet
        lnCellRow = Cast(.GetNodeAttributeValue(lcRow, "r") As I)
        lnRowHt   = Cast(.GetNodeAttributeValue(lcRow, "ht") As N(14,7))
        llCustHt  = Iif(Cast(.GetNodeAttributeValue(lcRow, "customHeight") As I) = 1, True, False)
        If llCustHt Or (lnRowHt > 0.00 And lnRowHt != 30)
          Insert Into xl_rowheights (workbook, Sheet, Row, Height) Values (tnWB, tnSheet, lnCellRow, lnRowHt)
        Endif

        lnCol     = 1
        lnCellCol = 1
        lcColumn  = .GetNodeElement(lcRow, 'c', lnCol)
        Do While !Empty(lcColumn)                                            && Process each column in a row
          lnCellCol = .ColumnAsciiToIndex(.GetNodeAttributeValue(lcColumn, "r"))
          loCell    = .ReadCellValueFormat(tnWB, lcColumn, lnCellCol)
          If tlForceTextFormat And loCell.DataType != DATA_TYPE_CHAR
            loCell.DataType  = DATA_TYPE_CHAR
            loCell.StringNdx = .AddStringValue(tnWB, Transform(loCell.CellVal), False)
            loCell.CellVal   = "StringValue"
            loCell.formula   = ""
          Endif
          Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, cellvalue, DataType, cellformula, stringid, cellxfs, celldeleted) ;
            VALUES (tnWB, tnSheet, lnCellRow, lnCellCol, loCell.CellVal, loCell.DataType, loCell.formula, loCell.StringNdx, loCell.CellXfsId, False)

          lnCol = lnCol + 1
          lcColumn = .GetNodeElement(lcRow, 'c', lnCol)
        Enddo
        Insert Into xl_rows (workbook, Sheet, Row, maxcol) Values (tnWB, tnSheet, lnCellRow, lnCellCol)

        lnNode = lnNode + 1
        lcRow = .GetNodeElement(lcSheet, 'row', lnNode)
      Enddo

*-*	Get the validations
      lcValidations = .GetNodeElement(lcSheet, 'dataValidations', 1)
      lnCnt = Cast(.GetNodeAttributeValue(lcValidations, "count") As I)
      For lnNode=1 To lnCnt
        lcValidation = .GetNodeElement(lcValidations, 'dataValidation', lnNode)
        If Empty(lcValidation)
          Set Step On
          Loop
        Endif
        lcValidSqRef = .GetNodeAttributeValue(lcValidation, "sqref")
        If Empty(lcValidSqRef)
          Set Step On
          Loop
        Endif
        lcValidType = .GetNodeAttributeValue(lcValidation, "type")
        Do Case
        Case lcValidType = "none"
          lnValidType = NONE_VALID_TYPE

        Case lcValidType = "whole"
          lnValidType = WHOLE_VALID_TYPE

        Case lcValidType = "decimal"
          lnValidType = DECIMAL_VALID_TYPE

        Case lcValidType = "list"
          lnValidType = LIST_VALID_TYPE

        Case lcValidType = "date"
          lnValidType = DATE_VALID_TYPE

        Case lcValidType = "time"
          lnValidType = TIME_VALID_TYPE

        Case lcValidType = "textLength"
          lnValidType = TXTLEN_VALID_TYPE

        Case lcValidType = "custom"
          lnValidType = CUSTOM_VALID_TYPE

        Otherwise
          lnValidType = NONE_VALID_TYPE
        Endcase

        lcValidStyle = .GetNodeAttributeValue(lcValidation, "errorStyle")
        Do Case
        Case lcValidStyle = "stop"
          lnValidStyle = STOP_VALID_STYLE

        Case lcValidStyle = "warning"
          lnValidStyle = WARN_VALID_STYLE

        Case lcValidStyle = "information"
          lnValidStyle = INFO_VALID_STYLE

        Otherwise
          lnValidStyle = 0
        Endcase

        lcOperator = .GetNodeAttributeValue(lcValidation, "operator")
        Do Case
        Case lcOperator = "between"
          lnOperator = BETWEEN_VALID_OPER

        Case lcOperator = "notBetween"
          lnOperator = NOTBETW_VALID_OPER

        Case lcOperator = "equal"
          lnOperator = EQUAL_VALID_OPER

        Case lcOperator = "notEqual"
          lnOperator = NOTEQUAL_VALID_OPER

        Case lcOperator = "lessThan"
          lnOperator = LESSTHAN_VALID_OPER

        Case lcOperator = "lessThanOrEqual"
          lnOperator = LESSOREQUAL_VALID_OPER

        Case lcOperator = "greaterThan"
          lnOperator = GREATTHAN_VALID_OPER

        Case lcOperator = "greaterThanOrEqual"
          lnOperator = GREATOREQUAL_VALID_OPER

        Otherwise
          lnOperator = 0
        Endcase

        lcErrMsg       = .GetNodeAttributeValue(lcValidation, "error")
        lcErrTitle     = .GetNodeAttributeValue(lcValidation, "errorTitle")
        lcPrompt       = .GetNodeAttributeValue(lcValidation, "prompt")
        llAllowBlank   = Iif(Cast(.GetNodeAttributeValue(lcValidation, "allowBlank") As I)=1, True, False)
        llShowInputMsg = Iif(Cast(.GetNodeAttributeValue(lcValidation, "showInputMessage") As I)=1, True, False)
        llShowErrMsg   = Iif(Cast(.GetNodeAttributeValue(lcValidation, "showErrorMessage") As I)=1, True, False)

        lcFormula1     = .GetNodeElementValue(.GetNodeElement(lcValidation, 'formula1', 1, True))
        lcFormula2     = .GetNodeElementValue(.GetNodeElement(lcValidation, 'formula2', 1, True))

        If Left(lcFormula1, 1) = '"' Or Left(lcFormula2, 1) = '"'
          lcFormula1 = Alltrim(lcFormula1, 1, '"')
          lcFormula2 = Alltrim(lcFormula2, 1, '"')
          llFormula  = False
        Else
          llFormula = True
        Endif
        Insert Into xl_validation (workbook, Sheet, vtype, vstyle, voperator, errmsg, errtitle, vprompt, allowblank, showinpmsg, showerrmsg, formula, formula1, formula2) ;
          VALUES (tnWB, tnSheet, lnValidType, lnValidStyle, lnOperator, lcErrMsg, lcErrTitle, lcPrompt, llAllowBlank, llShowInputMsg, llShowErrMsg, llFormula, lcFormula1, lcFormula2)
        lnIndex = xl_validation.validndx

        loCell = .CellRefAsciiToIndex(lcValidSqRef)
        If .GetCellRecord(tnWB, tnSheet, loCell.Row, loCell.Column)
          Replace xl_cells.validndx With lnIndex In xl_cells
        Else
          Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
            VALUES (tnWB, tnSheet, loCell.Row, loCell.Column, DATA_TYPE_NONE, 0, False, -1, lnIndex)

          .SetRowMaxColumn(tnWB, tnSheet, loCell.Row, loCell.Column)
        Endif
      Endfor
    Endwith
    Return True

*-- Reads the styles.xml file for opening a workbook
  Protected Procedure ReadStylesXML
    Lparameters tcFilePath, tnWB
    Local lcCellXfs, lnCnt, lcXf, lnIndent, lnWrapTx, lcAlign, lcStyleSheet, lnNode, lnNumFmtId, lnFontId, lnFillId, lnBorderId
    Local lcHAlign, lcVAlign, lcNumFmts, lcFonts, lcNumFmt, lcFormatCode, lcFont, lnIndexed, lnFSize, lcFName, llFBold, llFItalic, llStrkThr
    Local lcFVPos, lcULine, lnFColor, lnTheme, lnTint, lcColor, lcFills, lcFill, lcPatternFill, lcPatternType, lnFgColor, lnFgIndexed, lcBorder
    Local lnBgColor, lnBgIndexed, lcBorders, lcBLeft, lcLStyle, lnLColor, lnLTheme, lnLTint, lnLIndex, lcBRight, lcRStyle, lnRColor, lnRTheme
    Local lnRTint, lnRIndex, lcBTop, lcTStyle, lnTColor, lnTTheme, lnTTint, lnTIndex, lcBBottm, lcBStyle, lnBColor, lnBTheme, lnBTint, lnBIndex
    Local lnDTint, lnDIndex, lcBDiag, lcDStyle, lnDColor, lnDTheme, lcIndexedColors, lcRgbColor, lcMruColors, lnIdCnt

    Raiseevent(This, "OnShowStatusMessage", 1, 2)

    With This
*-*	Open the styles sheet for processing
      lcStyleSheet = Filetostr(Addbs(tcFilePath) + "xl\styles.xml")

*-*	Retreive the cell format mapping information
      lnIdCnt   = -1
      lcCellXfs = .GetNodeElement(lcStyleSheet, 'cellXfs', 1)
      lnCnt     = Cast(.GetNodeAttributeValue(lcCellXfs, "count") As I)
      For lnNode=1 To lnCnt
        lcXf = .GetNodeElement(lcCellXfs, 'xf', lnNode)
        If Empty(lcXf)
          Set Step On
          Loop
        Endif
        lnIdCnt    = lnIdCnt + 1
        lcAlign    = .GetNodeElement(lcXf, 'alignment', 1)
        lcHAlign   = .GetNodeAttributeValue(lcAlign, "horizontal")
        lcVAlign   = .GetNodeAttributeValue(lcAlign, "vertical")
        lnNumFmtId = Cast(.GetNodeAttributeValue(lcXf, "numFmtId") As I)
        lnFontId   = Cast(.GetNodeAttributeValue(lcXf, "fontId") As I)
        lnFillId   = Cast(.GetNodeAttributeValue(lcXf, "fillId") As I)
        lnBorderId = Cast(.GetNodeAttributeValue(lcXf, "borderId") As I)
        lnIndent   = Cast(.GetNodeAttributeValue(lcAlign, "indent") As I)
        lnWrapTx   = Cast(.GetNodeAttributeValue(lcAlign, "wrapText") As I)
        lnRotation = Cast(.GetNodeAttributeValue(lcAlign, "textRotation") As I)

        Insert Into xl_cellxfs (workbook, Id, numfmtid, fontid, fillid, borderid, halign, valign, indent, wraptext, Rotation) ;
          VALUES (tnWB, lnIdCnt, lnNumFmtId, lnFontId, lnFillId, lnBorderId, lcHAlign, lcVAlign, lnIndent, lnWrapTx, lnRotation)
      Endfor
      .SetLastId(tnWB, lnIdCnt, "xl_cellxfs")

*-*	Retreive the number format information
      lnIdCnt   = -1
      lcNumFmts = .GetNodeElement(lcStyleSheet, 'numFmts', 1)
      lnCnt     = Cast(.GetNodeAttributeValue(lcNumFmts, "count") As I)
      For lnNode=1 To lnCnt
        lcNumFmt = .GetNodeElement(lcNumFmts, 'numFmt', lnNode)
        If Empty(lcNumFmt)
          Set Step On
          Loop
        Endif
        lnNumFmtId   = Cast(.GetNodeAttributeValue(lcNumFmt, "numFmtId") As I)
        lcFormatCode = .GetNodeAttributeValue(lcNumFmt, "formatCode")

        Insert Into xl_numFmts (workbook, Id, formatxml, formatcode, applydec) ;
          VALUES (tnWB, lnNumFmtId, lcFormatCode, .GetStringXml(lcFormatCode), False)
      Endfor

*-*	Retreive the font informaton
      lnIdCnt = -1
      lcFonts = .GetNodeElement(lcStyleSheet, 'fonts', 1)
      lnCnt   = Cast(.GetNodeAttributeValue(lcFonts, "count") As I)
      For lnNode=1 To lnCnt
        lcFont = .GetNodeElement(lcFonts, 'font', lnNode)
        If Empty(lcFont)
          Set Step On
          Loop
        Endif
        lnIdCnt   = lnIdCnt + 1
        lnFSize   = Cast(.GetNodeAttributeValue(.GetNodeElement(lcFont, 'sz', 1), "val") As I)
        lnFSize   = Iif(lnFSize>0, lnFSize, 10)
        lcFName   = .GetNodeAttributeValue(.GetNodeElement(lcFont, 'name', 1), "val")
        llFBold   = Iif(Atc("<b/>", lcFont)>0, True, False)
        llFItalic = Iif(Atc("<i/>", lcFont)>0, True, False)
        llStrkThr = Iif(Atc("<strike/>", lcFont)>0, True, False)
        lcFVPos   = .GetNodeAttributeValue(.GetNodeElement(lcFont, 'vertAlign', 1), "val")

        If Atc("<u/>", lcFont) > 0
          lcULine = UNDERLINE_SINGLE
        Else
          lcULine = .GetNodeAttributeValue(.GetNodeElement(lcFont, 'u', 1), "val")
          lcULine = Iif(Empty(lcULine), UNDERLINE_NONE, lcULine)
        Endif

        lcColor   = .GetNodeElement(lcFont, 'color', 1)
        lnFColor  = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
        lnTheme   = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
        lnTint    = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
        lnIndexed = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)

        Insert Into xl_fonts (workbook, Id, fname, Fsize, fcolor, fbold, fitalic, uline, strkthr, fvpos, Theme, tint, indexed) ;
          VALUES (tnWB, lnIdCnt, lcFName, lnFSize, lnFColor, llFBold, llFItalic, lcULine, llStrkThr, lcFVPos, lnTheme, lnTint, lnIndexed)
      Endfor
      .SetLastId(tnWB, lnIdCnt, "xl_fonts")

*-*	Retreive the cell fill information
      lnIdCnt = -1
      lcFills = .GetNodeElement(lcStyleSheet, 'fills', 1)
      lnCnt   = Cast(.GetNodeAttributeValue(lcFills, "count") As I)
      For lnNode=1 To lnCnt
        lcFill = .GetNodeElement(lcFills, 'fill', lnNode)
        If Empty(lcFill)
          Set Step On
          Loop
        Endif
        lnIdCnt       = lnIdCnt + 1
        lcPatternFill = .GetNodeElement(lcFill, 'patternFill', 1)
        lcPatternType = .GetNodeAttributeValue(lcPatternFill, "patternType")
        lcPatternType = Iif(Empty(lcPatternType), "none", lcPatternType)
        lnTheme       = Cast(.GetNodeAttributeValue(lcPatternFill, "theme") As I)

        lcColor = .GetNodeElement(lcPatternFill, 'fgColor', 1)
        If Empty(lcColor)
          lnFgColor   = 0
          lnFgIndexed = 0
        Else
          lnFgColor   = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnFgIndexed = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif
        lcColor = .GetNodeElement(lcPatternFill, 'bgColor', 1)
        If Empty(lcColor)
          lnBgColor   = 0
          lnBgIndexed = 0
        Else
          lnBgColor   = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnBgIndexed = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif
        Insert Into xl_fills (workbook, Id, patttype, Theme, tint, fgcolor, bgcolor, fgindexed, bgindexed) ;
          VALUES (tnWB, lnIdCnt, lcPatternType, lnTheme, lnTint, lnFgColor, lnBgColor, lnFgIndexed, lnBgIndexed)
      Endfor
      .SetLastId(tnWB, lnIdCnt, "xl_fills")

*-*	Retreive the cell border information
      lnIdCnt   = -1
      lcBorders = .GetNodeElement(lcStyleSheet, 'borders', 1)
      lnCnt     = Cast(.GetNodeAttributeValue(lcBorders, "count") As I)
      For lnNode=1 To lnCnt
        lcBorder = .GetNodeElement(lcBorders, 'border', lnNode)
        If Empty(lcBorder)
          Set Step On
          Loop
        Endif
        lnIdCnt  = lnIdCnt + 1
        lcBLeft  = .GetNodeElement(lcBorder, 'left', 1)
        lcLStyle = .GetNodeAttributeValue(lcBLeft, "style")
        lcColor  = .GetNodeElement(lcBLeft, 'color', 1)
        If Empty(lcColor)
          lnLColor = Rgb(0, 0, 0)
          lnLTheme = 0
          lnLTint  = 0
          lnLIndex = 0
        Else
          lnLColor = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnLTheme = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
          lnLTint  = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
          lnLIndex = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif

        lcBRight = .GetNodeElement(lcBorder, 'right', 1)
        lcRStyle = .GetNodeAttributeValue(lcBRight, "style")
        lcColor  = .GetNodeElement(lcBRight, 'color', 1)
        If Empty(lcColor)
          lnRColor = Rgb(0, 0, 0)
          lnRTheme = 0
          lnRTint  = 0
          lnRIndex = 0
        Else
          lnRColor = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnRTheme = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
          lnRTint  = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
          lnRIndex = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif

        lcBTop   = .GetNodeElement(lcBorder, 'top', 1)
        lcTStyle = .GetNodeAttributeValue(lcBTop, "style")
        lcColor  = .GetNodeElement(lcBTop, 'color', 1)
        If Empty(lcColor)
          lnTColor = Rgb(0, 0, 0)
          lnTTheme = 0
          lnTTint  = 0
          lnTIndex = 0
        Else
          lnTColor = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnTTheme = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
          lnTTint  = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
          lnTIndex = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif

        lcBBottm = .GetNodeElement(lcBorder, 'bottom', 1)
        lcBStyle = .GetNodeAttributeValue(lcBBottm, "style")
        lcColor  = .GetNodeElement(lcBBottm, 'color', 1)
        If Empty(lcColor)
          lnBColor = Rgb(0, 0, 0)
          lnBTheme = 0
          lnBTint  = 0
          lnBIndex = 0
        Else
          lnBColor = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnBTheme = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
          lnBTint  = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
          lnBIndex = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif

        lcBDiag  = .GetNodeElement(lcBorder, 'diagonal', 1)
        lcDStyle = .GetNodeAttributeValue(lcBDiag, "style")
        lnDiagDn = Cast(.GetNodeAttributeValue(lcBDiag, "diagonalDown") As I)
        lnDiagUp = Cast(.GetNodeAttributeValue(lcBDiag, "diagonalUp") As I)
        lcColor  = Cast(.GetNodeElement(lcBDiag, 'color', 1) As I)
        If Empty(lcColor)
          lnDColor = Rgb(0, 0, 0)
          lnDTheme = 0
          lnDTint  = 0
          lnDIndex = 0
        Else
          lnDColor = .ConvertHexStringToNumeric(.GetNodeAttributeValue(lcColor, "rgb"))
          lnDTheme = Cast(.GetNodeAttributeValue(lcColor, "theme") As I)
          lnDTint  = Cast(.GetNodeAttributeValue(lcColor, "tint") As I)
          lnDIndex = Cast(.GetNodeAttributeValue(lcColor, "indexed") As I)
        Endif

        Insert Into xl_borders (workbook, Id, lstyle, lcolor, ltheme, ltint, lindexed, rstyle, rcolor, rtheme, rtint, rindexed, tstyle, tcolor, ttheme, ttint, tindexed, ;
          bstyle, bcolor, btheme, btint, bindexed, dstyle, dcolor, dtheme, dtint, dindexed, diagdn, diagup) ;
          VALUES (tnWB, lnIdCnt, lcLStyle, lnLColor, lnLTheme, lnLTint, lnLIndex, lcRStyle, lnRColor, lnRTheme, lnRTint, lnRIndex, lcTStyle, lnTColor, lnTTheme, ;
          lnTTint, lnTIndex, lcBStyle, lnBColor, lnBTheme, lnBTint, lnBIndex, lcDStyle, lnDColor, lnDTheme, lnDTint, lnDIndex, lnDiagDn, lnDiagUp)
      Endfor
      .SetLastId(tnWB, lnIdCnt, "xl_borders")

*-*	Retreive the indexed color information
      lcColors = .GetNodeElement(lcStyleSheet, 'colors', 1)
      If !Empty(lcColors)
        lcIndexedColors = .GetNodeElement(lcColors, 'indexedColors', 1)
        If !Empty(lcIndexedColors)
          lnNode = 1
          lcRgbColor = .GetNodeElement(lcIndexedColors, 'rgbColor', lnNode)
          Do While !Empty(lcRgbColor)
            lcHexColor = .GetNodeAttributeValue(lcRgbColor, "rgb")

            Insert Into xl_ndxcolors (workbook, indexid, rgbcolor) Values (tnWB, lnNode, lcHexColor)

            lnNode = lnNode + 1
            lcRgbColor = .GetNodeElement(lcIndexedColors, 'rgbColor', lnNode)
          Enddo
          .SetLastId(tnWB, lnNode-1, "xl_ndxcolors")
        Endif

        lcMruColors = .GetNodeElement(lcColors, 'mruColors', 1)
        If !Empty(lcMruColors)
          lnNode = 1
          lcRgbColor = .GetNodeElement(lcMruColors, 'color', lnNode)
          Do While !Empty(lcRgbColor)
            lcHexColor = .GetNodeAttributeValue(lcRgbColor, "rgb")

            Insert Into xl_mrucolors (workbook, indexid, rgbcolor) Values (tnWB, lnNode, lcHexColor)

            lnNode = lnNode + 1
            lcRgbColor = .GetNodeElement(lcMruColors, 'color', lnNode)
          Enddo
          .SetLastId(tnWB, lnNode-1, "xl_mrucolors")
        Endif
      Endif
    Endwith

*-- Removes the working directories
  Protected Procedure RemoveWorkingDirectories
    Lparameters tcDir
    Local llReturn, loException
    Try
      llReturn = apiRemoveDirectory(tcDir) != 0

    Catch To loException
      Set Step On
      llReturn = False
    Endtry
    Return llReturn

*-- Renames the selected sheet
  Procedure RenameSheet
  Lparameters tnWB, txSheet, tcSheetName
  Local llReturn
  llReturn = False
  Do Case
  Case Vartype(txSheet) = "C"
    If Seek(BinToC(tnWB)+Upper(txSheet), "xl_sheets", "shname")
      Replace xl_sheets.shname With tcSheetName In xl_sheets
      llReturn = True
    Endif

  Case Vartype(txSheet) = "N"
    If Seek(BinToC(tnWB)+BinToC(txSheet), "xl_sheets", "sheetndx")
      Replace xl_sheets.shname With tcSheetName In xl_sheets
      llReturn = True
    Endif
  Endcase
  Return llReturn

*-- Resets the column width to default
  Procedure ResetColumnWidth
  Lparameters tnWB, tnSheet, tnColumn
  If Seek(This.GetColWidthKeyExpr(tnWB, tnSheet, tnColumn), "xl_colwidths", "shcolumn")
    Delete In xl_colwidths
  Endif

*-- Saves the file as a UTF-8
  Protected Procedure SaveAsUtf8
    Lparameters tcFileName, tcOutStr
    Strtofile(Strconv(tcOutStr, 9), tcFileName, 4)

*-- Saves the selected grid to a workbook
  Procedure SaveGridToWorkbook
  Lparameters toGrid, txWB, tlFreeze, tlSaveWB, tcSheetName, tlInclHiddenCols
  Local lcAlias, lnCol, lnWB, lnSh, lnRow, lcField, loReturn, loColumn, lnStyle, lnColCount, lnDefStyle
  Local Array laRowCount[1], laColOrder[1]
  loReturn = Createobject("Empty")
  AddProperty(loReturn, "Workbook",  0)
  AddProperty(loReturn, "Sheet", 0)
  With This
    Do Case
    Case Inlist(Pcount(), 0, 1)
      Return loReturn

    Case Pcount() = 2
      tlFreeze     = True
      tlSaveWB     = True
      tcSheetName  = .DeriveSheetName(txWB)
      tlInclHiddenCols = True

    Case Pcount() = 3
      tlSaveWB     = True
      tcSheetName  = .DeriveSheetName(txWB)
      tlInclHiddenCols = True

    Case Pcount() = 4
      tcSheetName  = .DeriveSheetName(txWB)
      tlInclHiddenCols = True

    Case Pcount() = 5
      tlInclHiddenCols = True
    Endcase
    If Vartype(tcSheetName) != 'C' Or Empty(tcSheetName)
      tcSheetName = .DeriveSheetName(txWB)
    Endif
    tcSheetName = Chrtran(tcSheetName, ":\/?*", "_____")
    If Len(tcSheetName) > LIMITS_MAX_SH_NAME
      If .AutoTrimSheetName
        tcSheetName = Left(Alltrim(tcSheetName), LIMITS_MAX_SH_NAME)
      Else
        Return loReturn
      Endif
    Endif
    If Vartype(toGrid) != "O" And Pemstatus(toGrid, "BaseClass", 5) And Lower(toGrid.BaseClass) != "grid"
      Return loReturn
    Endif
    Do Case
    Case Vartype(txWB) = "C"
      lnWB = .CreateWorkbook(txWB)
      If lnWB = 0
        lnWB = .GetWorkbook(txWB)
      Endif

    Case Vartype(txWB) = "N"
      If Seek(BinToC(txWB), "xl_workbooks", "workbook")
        lnWB = txWB
      Else
        Return loReturn
      Endif

    Otherwise
      Return loReturn
    Endcase
    If lnWB > 0
      lnSh = .AddSheet(lnWB, tcSheetName)
      If lnSh > 0
        lcAlias = Juststem(toGrid.RecordSource)

*-*		Get the record count and display a status window
        Select Count(*) From &lcAlias Into Array laRowCount
        Raiseevent(This, "OnShowStatusMessage", 3, 0, laRowCount[1])

*-*		Get the number of columns to process and
*-*		Get the column output order based on grid display order                       && Change requested by Matt Slay to output by grid display order and visible setting
        If tlInclHiddenCols
          lnColCount = toGrid.ColumnCount
          Dimension laColOrder[lnColCount, 2]
          For lnCol=1 To toGrid.ColumnCount
            laColOrder[lnCol, 1] = toGrid.Columns(lnCol).ColumnOrder              && Grid display order
            laColOrder[lnCol, 2] = lnCol                                          && Column property order
          Endfor
        Else
          lnColCount = 0
          For lnCol=1 To toGrid.ColumnCount
            If toGrid.Columns(lnCol).Visible
              lnColCount = lnColCount + 1
              Dimension laColOrder[lnColCount, 2]
              laColOrder[lnColCount, 1] = toGrid.Columns(lnCol).ColumnOrder     && Grid display order
              laColOrder[lnColCount, 2] = lnCol                                 && Column property order
            Endif
          Endfor
        Endif
        If lnColCount = 0
          Return False
        Endif
        Asort(laColOrder)

*-*		Add the header row if defined and set the column widths; get the column font info
        If toGrid.HeaderHeight > 0                                                    && Change recommendation by Doug Hennig (if no headers, start in first row)
          lnRow = 1
          For lnCol=1 To lnColCount
            loColumn = toGrid.Columns(laColOrder[lnCol, 2])
            .SetCellValue(lnWB, lnSh, 1, lnCol, loColumn.Header1.Caption)
            .SetColumnWidth(lnWB, lnSh, lnCol, .ConvertPixelsToExcelUnits(loColumn.Width))
          Endfor
        Else
          lnRow = 0
          For lnCol=1 To lnColCount
            loColumn = toGrid.Columns(laColOrder[lnCol, 2])
            .SetColumnWidth(lnWB, lnSh, lnCol, .ConvertPixelsToExcelUnits(loColumn.Width))
          Endfor
        Endif

*-*		Add the cell data values
        Select (lcAlias)
        Scan
          lnRow = lnRow + 1
          Raiseevent(This, "OnShowStatusMessage", 3, lnRow)
          For lnCol=1 To lnColCount
            loColumn = toGrid.Columns(laColOrder[lnCol, 2])
            lcField  = loColumn.ControlSource
            .SetCellValue(lnWB, lnSh, lnRow, lnCol, &lcField)
          Endfor
        Endscan

*-*		Set the cell style formatting
        lnDefStyle = .IsFormatStyleDefined(lnWB, .DefaultFont, .DefaultFontSize)
        If Isnull(lnDefStyle)
          lnDefStyle = .CreateFormatStyle(lnWB)
          .AddStyleFont(lnWB, lnDefStyle, .DefaultFont, .DefaultFontSize)
        Endif

        For lnCol=1 To lnColCount
          loColumn = toGrid.Columns(laColOrder[lnCol, 2])
          lcField  = loColumn.ControlSource
          If Vartype(&lcField) != DATA_TYPE_DATE
            If .DefaultFont != loColumn.FontName Or .DefaultFontSize != loColumn.FontSize   && Change recommended by Doug Hennig; 2017-06-12
              lnStyle = .IsFormatStyleDefined(lnWB, loColumn.FontName, loColumn.FontSize)
              If Isnull(lnStyle)
                lnStyle = .CreateFormatStyle(lnWB)
                .AddStyleFont(lnWB, lnStyle, loColumn.FontName, loColumn.FontSize)
              Endif
            Else
              lnStyle = lnDefStyle
            Endif
            .SetCellStyleRange(lnWB, lnSh, 1, lnCol, laRowCount[1]+1, lnCol, lnStyle)
          Endif
        Endfor

*-*		Freeze the first row if specified
        If tlFreeze
          .FreezePanes(lnWB, lnSh, 1, 0)
        Endif
        Raiseevent(This, "OnShowStatusMessage", 3, -1)
        If tlSaveWB
          .SaveWorkbook(lnWB)
        Endif
        loReturn.Sheet    = lnSh
        loReturn.workbook = lnWB
      Endif
    Endif
  Endwith
  Return loReturn

*-- Saves the passed grid to a workbook without adding the content to the internal cursors
  Procedure SaveGridToWorkbookEx
  Lparameters toGrid, tcFileName, tlFreeze, tcSheetName, tlInclHiddenCols
  Local llError, lcTempPath, lcAlias, lcFileName, lhFile, lcRowDefHt, lnCol, lnRow, lcSETCentury, lxCellValue
  Local lnColCount, lnColWidth, lcCurAlias
  Local Array laColOrder[1], laSheetNames[1]
  llError    = False
  lcCurAlias = Alias()
  If Vartype(toGrid) != "O" And Pemstatus(toGrid, "BaseClass", 5) And Lower(toGrid.BaseClass) != "grid"
    Return False
  Endif
  With This
    Do Case
    Case Pcount() < 2
      Return False

    Case Pcount() = 2
      tlFreeze     = True
      tcSheetName  = toGrid.Name
      tlInclHiddenCols = True

    Case Pcount() = 3
      tcSheetName  = toGrid.Name
      tlInclHiddenCols = True

    Case Pcount() = 4
      tlInclHiddenCols = True
    Endcase
    If Vartype(tcSheetName) != 'C' Or Empty(tcSheetName)
      tcSheetName = toGrid.Name
    Endif
    tcSheetName = Chrtran(tcSheetName, ":\/?*", "_____")
    If Len(tcSheetName) > LIMITS_MAX_SH_NAME
      If .AutoTrimSheetName
        tcSheetName = Left(Alltrim(tcSheetName), LIMITS_MAX_SH_NAME)
      Else
        Return False
      Endif
    Endif

*-* Set temporary output path
    lcTempPath = .CreateWorkingDirectories()
    If Isnull(lcTempPath)
      Return False
    Endif

*-*	Save current settings and then set to US formats
    lcSetPoint  = Set("POINT")
    lcSeparator = Set("SEPARATOR")
    If lcSetPoint != "."
      Set Point To "."
      Set Separator To ","
    Endif

*-*	Output the workbook
    lcAlias = Juststem(toGrid.RecordSource)

*-*	Get the number of columns to process
*-*	Get the column output order based on grid display order                          && Change requested by Matt Slay to output by grid display order and visible setting
    If tlInclHiddenCols
      lnColCount = toGrid.ColumnCount
      Dimension laColOrder[lnColCount, 2]
      For lnCol=1 To toGrid.ColumnCount
        laColOrder[lnCol, 1] = toGrid.Columns(lnCol).ColumnOrder                     && Grid display order
        laColOrder[lnCol, 2] = lnCol                                                 && Column property order
      Endfor
    Else
      lnColCount = 0
      For lnCol=1 To toGrid.ColumnCount
        If toGrid.Columns(lnCol).Visible
          lnColCount = lnColCount + 1
          Dimension laColOrder[lnColCount, 2]
          laColOrder[lnColCount, 1] = toGrid.Columns(lnCol).ColumnOrder             && Grid display order
          laColOrder[lnColCount, 2] = lnCol                                         && Column property order
        Endif
      Endfor
    Endif
    If lnColCount = 0
      Return False
    Endif
    Asort(laColOrder)

    Dimension laSheetNames[1]
    laSheetNames[1] = tcSheetName
    If .WriteDirectXMLs(lcTempPath, @laSheetNames)
*-*	Write the sheet XML
* TODO: переделать на textmerge
      Try
        lcFileName = lcTempPath + "xl\worksheets\sheet1.xml"
        lhFile = Fcreate(lcFileName)
        If lhFile < 0
          Throw
        Endif
        Fclose(lhFile)

        Set Textmerge On To Memvar lcOutStr Noshow
*-*		Write sheet XML header information
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\ xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        \\ xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"
        \\ xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

*-*		Write the sheet view information which includes the freeze pane information
        \\<dimension ref="A1:<<.ColumnIndexToAscii(lnColCount) + Transform(Reccount(lcAlias)+1)>>"/>
        If tlFreeze
          \\<sheetViews><sheetView workbookViewId="0">
          \\<pane state="frozenSplit" activePane="bottomLeft" topLeftCell="A2" ySplit="1"/>
          \\<selection activeCell="A2" sqref="A2" pane="bottomLeft"/>
          \\</sheetView></sheetViews>')
        Endif
        \\<sheetFormatPr defaultRowHeight="15"/>

*-*		Write sheet Column widths
        \\<cols>
        For lnCol=1 To lnColCount
          lnColWidth = .ConvertPixelsToExcelUnits(toGrid.Columns(laColOrder[lnCol, 2]).Width)
          \\<col min="<<Transform(lnCol)>>" max="<<Transform(lnCol)>>" width="<<Transform(lnColWidth+0.7109375)>>" customWidth="1"/>
        Endfor
        \\</cols>')

*-*		Write sheet cell data values
        lcRowDefHt = ' ht="30"'
        \\<sheetData>

*-*		Place field names in the first row
        If toGrid.HeaderHeight > 0  && Change recommendation by Doug Hennig (if no headers, then no output)
          \\<row r="1" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>> >')
          If .DefaultFont = "Calibri" And .DefaultFontSize = 11
            For lnCol=1 To lnColCount
              .WriteAsInLineString(1, lnCol, toGrid.Columns(laColOrder[lnCol, 2]).Header1.Caption)
            Endfor
          Else
            For lnCol=1 To lnColCount
              .WriteAsInLineFormattedString(1, lnCol, toGrid.Columns(laColOrder[lnCol, 2]).Header1.Caption)
            Endfor
          Endif
          \\</row>
          lnRow = 1
        Else
          lnRow = 0
        Endif

*-*		Write the grid field values to cells
        lcSETCentury = Set("CENTURY")
        Set Century On
        Select (lcAlias)
        If .DefaultFont = "Calibri" And .DefaultFontSize = 11
          Scan
            lnRow = lnRow + 1
            \\<row r="<<Transform(lnRow)>>" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>> >
            For lnCol=1 To lnColCount
              lxCellValue = Evaluate(toGrid.Columns(laColOrder[lnCol, 2]).ControlSource)
              .WriteDirectCellValue(lnRow, lnCol, lxCellValue)
            Endfor
            \\</Row>
          Endscan
        Else
          Scan
            lnRow = lnRow + 1
            \\<Row r="<<Transform(lnRow)>>" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>> >
            For lnCol=1 To lnColCount
              lxCellValue = Evaluate(toGrid.Columns(laColOrder[lnCol, 2]).ControlSource)
              .WriteDirectCellFormattedValue(lnRow, lnCol, lxCellValue)
            Endfor
            \\</Row>
          Endscan
        Endif
        Set Century &lcSETCentury

*-*		End the sheet cell data values and worksheet
       \\</sheetData>
       \\<pageMargins Left="0.7" Right="0.7" Top="0.75" Bottom="0.75" Header="0.3" Footer="0.3"/>
       \\</worksheet>

      Catch To loException
        Set Step On
        llError = True
        .ErrorLevelId = 14
        Raiseevent(This, "OnShowErrorMessage", 14, loException.Message)

      Finally
        .SaveAsUtf8(lcFileName, lcOutStr)
      Endtry
      If llError
        llSuccess = False
      Else
        llSuccess = .CreateExcelFile(lcTempPath, Justpath(tcFileName), Juststem(tcFileName))    && Change recommendation by Doug Hennig [placement of code after ELSE]
      Endif
      .RemoveWorkingDirectories(lcTempPath)

    Else
      llSuccess = False
    Endif
    If !Empty(lcCurAlias)
      Select (lcCurAlias)
    Endif

*-*	Restore default settings
    If lcSetPoint != "."
      Set Point To (lcSetPoint)
      Set Separator To (lcSeparator)
    Endif
  Endwith
  Return llSuccess

*-- Same as SaveGridToWorkbookEx() method but handles multiple grids being passed
  Procedure SaveMultiGridtoWorkbookEx
  Lparameters toGrids, tcFileName
  Local llError, lcTempPath, lcAlias, lcFileName, lhFile, lcRowDefHt, lnCol, lnRow, lcSETCentury, lxCellValue
  Local lnColCount, lnColWidth, lcCurAlias, lnGrid, loGrid, loException, lcSheetName
  Local Array laColOrder[1], laSheetNames[1]
  llError    = False
  lcCurAlias = Alias()
  If Pcount() < 2
    Return False
  Endif
  With This
    If Vartype(toGrids) = "O"
      If Pemstatus(toGrids, "Count", 5) And Vartype(toGrids.Count) != "N"
        Return False
      Else
        If toGrids.Count < 1
          Return False
        Endif
      Endif

      If Pemstatus(toGrids, "List", 5)
        llError = False
        Try
          If Alen(toGrids.List, 2) != 4   && Column 1 Grid; Column 2  SheetName; Column 3 Freeze indicator, Column 4 HiddenCols indicator
            Throw
          Endif
        Catch To loException
          llError = True
        Endtry
        If llError
          Return False
        Endif
        For lnGrid=1 To toGrids.Count
*-*			Validate grid object
          loGrid = toGrids.List[lnGrid, 1]
          If Vartype(loGrid) != "O" And Pemstatus(loGrid, "BaseClass", 5) And Lower(loGrid.BaseClass) != "grid"
            Return False
          Endif
*-*			Validate sheet name
          lcSheetName = toGrids.List[lnGrid, 2]
          If Vartype(lcSheetName) != 'C' Or Empty(lcSheetName)
            lcSheetName = loGrid.Name
          Endif
          lcSheetName = Chrtran(lcSheetName, ":\/?*", "_____")
          If Len(lcSheetName) > LIMITS_MAX_SH_NAME
            If .AutoTrimSheetName
              lcSheetName = Left(Alltrim(lcSheetName), LIMITS_MAX_SH_NAME)
            Else
              Return False
            Endif
          Endif
          toGrids.List[lnGrid, 2] = lcSheetName
*-*			Validate freeze indicator
          If Vartype(toGrids.List[lnGrid, 3]) != 'L'
            toGrids.List[lnGrid, 3] = False
          Endif
*-*			Validate hidden column indicator
          If Vartype(toGrids.List[lnGrid, 4]) != 'L'
            toGrids.List[lnGrid, 4] = False
          Endif
        Endfor
      Endif
    Else
      Return False
    Endif

*-* Set temporary output path
    lcTempPath = .CreateWorkingDirectories()
    If Isnull(lcTempPath)
      Return False
    Endif

*-*	Save current settings and then set to US formats
    lcSetPoint  = Set("POINT")
    lcSeparator = Set("SEPARATOR")
    If lcSetPoint != "."
      Set Point To "."
      Set Separator To ","
    Endif

*-*	Output the workbook
    Dimension laSheetNames[toGrids.Count]
    For lnGrid=1 To toGrids.Count
      laSheetNames[lnGrid] = toGrids.List[lnGrid, 2]
    Endfor
    If !.WriteDirectXMLs(lcTempPath, @laSheetNames)
      Return False
    Endif

*-*	Write the sheet xml for each grid
    For lnGrid=1 To toGrids.Count
      loGrid  = toGrids.List[lnGrid, 1]
      lcAlias = Juststem(loGrid.RecordSource)

*-*	Get the number of columns to process
*-*	Get the column output order based on grid display order       && Change requested by Matt Slay to output by grid display order and visible setting
      If toGrids.List[lnGrid, 4]
        lnColCount = loGrid.ColumnCount
        Dimension laColOrder[lnColCount, 2]
        For lnCol=1 To loGrid.ColumnCount
          laColOrder[lnCol, 1] = loGrid.Columns(lnCol).ColumnOrder     && Grid display order
          laColOrder[lnCol, 2] = lnCol                                 && Column property order
        Endfor
      Else
        lnColCount = 0
        For lnCol=1 To loGrid.ColumnCount
          If loGrid.Columns(lnCol).Visible
            lnColCount = lnColCount + 1
            Dimension laColOrder[lnColCount, 2]
            laColOrder[lnColCount, 1] = loGrid.Columns(lnCol).ColumnOrder   && Grid display order
            laColOrder[lnColCount, 2] = lnCol                        && Column property order
          Endif
        Endfor
      Endif
      If lnColCount = 0
        Return False
      Endif
      Asort(laColOrder)

      Try
        lcFileName = lcTempPath + "xl\worksheets\sheet" + Transform(lnGrid) + ".xml"
        Set Textmerge On To Memvar lcOutStr Noshow
*-*		Write sheet XML header information
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
        \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\ xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        \\ xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"
        \\ xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

*-*		Write the sheet view information which includes the freeze pane information
        \\<dimension ref="A1:' + .ColumnIndexToAscii(lnColCount) + Transform(Reccount(lcAlias)+1) + '"/>
        If toGrids.List[lnGrid, 3]
          \\<sheetViews><sheetView workbookViewId="0">
          \\<pane state="frozenSplit" activePane="bottomLeft" topLeftCell="A2" ySplit="1"/>
          \\<selection activeCell="A2" sqref="A2" pane="bottomLeft"/>
          \\</sheetView></sheetViews>
        Endif
        \\<sheetFormatPr defaultRowHeight="15"/>

*-*		Write sheet Column widths
        \\<cols>
        For lnCol=1 To lnColCount
          lnColWidth = .ConvertPixelsToExcelUnits(loGrid.Columns(laColOrder[lnCol, 2]).Width)
          \\<col min="<<Transform(lnCol)>>" max="<<Transform(lnCol)>>" width="<<Transform(lnColWidth+0.7109375)>>" customWidth="1"/>
        Endfor
        \\</cols>

*-*		Write sheet cell data values
        lcRowDefHt = ' ht="30"'
        \\<sheetData>

*-*		Place field names in the first row
        If loGrid.HeaderHeight > 0                                   && Change recommendation by Doug Hennig (if no headers, then no output)
          \\<row r="1" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>>>
          If .DefaultFont = "Calibri" And .DefaultFontSize = 11
            For lnCol=1 To lnColCount
              .WriteAsInLineString(lhFile, 1, lnCol, loGrid.Columns(laColOrder[lnCol, 2]).Header1.Caption)
            Endfor
          Else
            For lnCol=1 To lnColCount
              .WriteAsInLineFormattedString(lhFile, 1, lnCol, loGrid.Columns(laColOrder[lnCol, 2]).Header1.Caption)
            Endfor
          Endif
          \\</row>
          lnRow = 1
        Else
          lnRow = 0
        Endif

*-*		Write the grid field values to cells
        lcSETCentury = Set("CENTURY")
        Set Century On
        Select (lcAlias)
        If .DefaultFont = "Calibri" And .DefaultFontSize = 11
          Scan
            lnRow = lnRow + 1
            \\<row r="<<Transform(lnRow)>>" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>>>
            For lnCol=1 To lnColCount
              lxCellValue = Evaluate(loGrid.Columns(laColOrder[lnCol, 2]).ControlSource)
              .WriteDirectCellValue(lhFile, lnRow, lnCol, lxCellValue)
            Endfor
            \\</row>
          Endscan
        Else
          Scan
            lnRow = lnRow + 1
            \\<row r="<<Transform(lnRow)>>" spans="1:<<Transform(lnColCount)>>"<<lcRowDefHt>>>
            For lnCol=1 To lnColCount
              lxCellValue = Evaluate(loGrid.Columns(laColOrder[lnCol, 2]).ControlSource)
              .WriteDirectCellFormattedValue(lhFile, lnRow, lnCol, lxCellValue)
            Endfor
            \\</row>
          Endscan
        Endif
        Set Century &lcSETCentury

*-*		End the sheet cell data values and worksheet
        \\</sheetData>
        \\<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
        \\</worksheet>

      Catch To loException
        Set Step On
        llError = True
        .ErrorLevelId = 14
        Raiseevent(This, "OnShowErrorMessage", 14, loException.Message)

      Finally
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)
      Endtry
    Endfor
    If llError
      llSuccess = False
    Else
      llSuccess = .CreateExcelFile(lcTempPath, Justpath(tcFileName), Juststem(tcFileName))   && Change recommendation by Doug Hennig [placement of code after ELSE]
    Endif
*-*	Cleanup
    .RemoveWorkingDirectories(lcTempPath)
    If !Empty(lcCurAlias)
      Select (lcCurAlias)
    Endif

*-*	Restore default settings
    If lcSetPoint != "."
      Set Point To (lcSetPoint)
      Set Separator To (lcSeparator)
    Endif
  Endwith
  Return llSuccess

*-- Saves the table passed (alias) as a workbook sheet
  Procedure SaveTableToWorkbook
  Lparameters tcAlias, tnWB, tlFreeze, tlSaveWB, tcSheetName
  Local lcAlias, lnCol, lnWB, lnSh, lnRow, lcField, lnFldCnt, llError, llOpenedTbl, loException, lcCurAlias, loReturn
  Local lnNdx, lcCaption
  Local Array laFields[1], laRowCount[1]
  loReturn = Createobject("Empty")
  AddProperty(loReturn, "Workbook",  0)
  AddProperty(loReturn, "Sheet", 0)
  Do Case
  Case Pcount() < 2
    Return loReturn

  Case Pcount() = 2
    tlFreeze = True
    tlSaveWB = False
    tcSheetName = Juststem(tcAlias)

  Case Pcount() = 3
    tlSaveWB = False
    tcSheetName = Juststem(tcAlias)

  Case Pcount() = 4
    tcSheetName = Juststem(tcAlias)
  Endcase
  If Vartype(tcSheetName) != 'C' Or Empty(tcSheetName)
    tcSheetName = This.DeriveSheetName(tnWB)
  Endif
  tcSheetName = Chrtran(tcSheetName, ":\/?*", "_____")
  If Len(tcSheetName) > LIMITS_MAX_SH_NAME
    If This.AutoTrimSheetName
      tcSheetName = Left(Alltrim(tcSheetName), LIMITS_MAX_SH_NAME)
    Else
      Return loReturn
    Endif
  Endif
  Do Case
  Case Used(tcAlias)
    lcCurAlias  = Alias()
    llOpenedTbl = False
    lcAlias = tcAlias

  Case Occurs("\", tcAlias) > 0
    Try
      lcAlias = Juststem(tcAlias)
      Use (tcAlias) In 0 Alias (lcAlias) Shared
      llOpenedTbl = True

    Catch To loException
      llError = True
    Endtry
    If llError
      Return loReturn
    Endif

  Otherwise
    Return loReturn
  Endcase
  Do Case
  Case Vartype(tnWB) = "C"
    lnWB = This.CreateWorkbook(tnWB)
    If lnWB = 0
      lnWB = This.GetWorkbook(tnWB)
    Endif

  Case Vartype(tnWB) = "N"
    If Seek(BinToC(tnWB), "xl_workbooks", "workbook")
      lnWB = tnWB
    Else
      Return loReturn
    Endif

  Otherwise
    Return loReturn
  Endcase
  If lnWB > 0
    lnSh = This.AddSheet(lnWB, tcSheetName)
    If lnSh > 0
*-*		Get the record count and display a status window
      Select Count(*) From &lcAlias Into Array laRowCount
      Raiseevent(This, "OnShowStatusMessage", 3, 0, laRowCount[1])

      Select (lcAlias)
      lnFldCnt = Afields(laFields, lcAlias)
      For lnNdx=1 To lnFldCnt
        Try
          lcCaption = DBGetProp(lcAlias + "." + laFields[lnNdx, 1], "FIELD", "CAPTION")
          laFields[lnNdx, 3] = Iif(Empty(lcCaption), laFields[lnNdx, 1], lcCaption)

        Catch To loException
          laFields[lnNdx, 3] = laFields[lnNdx, 1]
        Endtry
      Endfor

*-*		Place field names in the first row
      For lnCol=1 To lnFldCnt
        This.SetCellValue(lnWB, lnSh, 1, lnCol, laFields[lnCol, 3])
      Endfor
*-*		Populate next rows with data
      lnRow = 0
      Scan
        lnRow = lnRow + 1
        Raiseevent(This, "OnShowStatusMessage", 3, lnRow)
        For lnCol=1 To lnFldCnt
          If laFields[lnCol, 2] = "G"
            This.SetCellValue(lnWB, lnSh, 1 + lnRow, lnCol, "General")
          Else
            This.SetCellValue(lnWB, lnSh, 1 + lnRow, lnCol, Evaluate(lcAlias + '.' + laFields[lnCol, 1]))
          Endif
        Endfor
      Endscan

      If tlFreeze
        This.FreezePanes(lnWB, lnSh, 1, 0)
      Endif
      Raiseevent(This, "OnShowStatusMessage", 3, -1)
      If tlSaveWB
        This.SaveWorkbook(lnWB)
      Endif
    Else
      lnWB = 0
      lnSh = 0
    Endif
  Else
    lnSh = 0
    lnWB = 0
  Endif
  If llOpenedTbl
    Use In &lcAlias
  Else
    lnSh = 0
    lnWB = 0
    If !Empty(lcCurAlias)
      Select &lcCurAlias
    Endif
  Endif
  loReturn.Sheet    = lnSh
  loReturn.workbook = lnWB
  Return loReturn

*-- Creates a new workbook from a table; direct write to XLSX file without loading internal cursors
  Procedure SaveTableToWorkbookEx
  Lparameters tcAlias, tcXlsxName, taFields, tlFreeze, tcSheetName
  Local lcCurAlias, llOpenedTbl, lcAlias, llError, loException, lhFile, lnCol, lcRowDefHt
  Local lnFldCnt, lnRow, lxCellValue, lcTempPath, lcFileName, lcSETCentry, lcCaption, llSuccess, lcSetPoint
  Local lcSeparator
  Local Array laSheetNames[1]
  External Array taFields
  llError = False
  Do Case
  Case Pcount() < 2
    Return False

  Case Pcount() = 2
    Dimension taFields[1, 2]
    taFields[1, 1] = Null
    tlFreeze       = True
    tcSheetName    = Juststem(tcAlias)

  Case Pcount() = 3
    tlFreeze    = True
    tcSheetName = Juststem(tcAlias)

  Case Pcount() = 4
    tcSheetName = Juststem(tcAlias)
  Endcase
  If Vartype(tcSheetName) != 'C' Or Empty(tcSheetName)
    tcSheetName = Juststem(tcAlias)
  Endif
  tcSheetName = Chrtran(tcSheetName, ":\/?*", "_____")
  If Len(tcSheetName) > LIMITS_MAX_SH_NAME
    If This.AutoTrimSheetName
      tcSheetName = Left(Alltrim(tcSheetName), LIMITS_MAX_SH_NAME)
    Else
      Return False
    Endif
  Endif
  lcCurAlias = Alias()
  lcAlias    = Juststem(tcAlias)
  Do Case
  Case Used(lcAlias)
    llOpenedTbl = False

  Case Occurs("\", tcAlias) > 0
    Try
      Use (tcAlias) In 0 Alias (lcAlias) Shared
      llOpenedTbl = True

    Catch To loException
      llError = True
    Endtry
    If llError
      Return False
    Endif

  Otherwise
    Return False
  Endcase
*-*	Get the table fields to export if not passed as a parameter
  Select (lcAlias)
  If Isnull(taFields) Or Empty(taFields[1])
    Dimension taFields[1, 5]
    taFields = ""
    lnFldCnt = Afields(taFields, lcAlias)
    For lnCol=1 To lnFldCnt
      Try
        lcCaption = DBGetProp(lcAlias + "." + taFields[lnCol, 1], "FIELD", "CAPTION")
        taFields[lnCol, 2] = Iif(Empty(lcCaption), taFields[lnCol, 1], lcCaption)

      Catch To loException
        taFields[lnCol, 2] = taFields[lnCol, 1]
      Endtry
    Endfor
  Else
    lnFldCnt = Alen(taFields, 1)
    If Alen(taFields, 2) < 2 Or Vartype(taFields[1, 2]) != "C"
      Return False
    Endif
  Endif

*-* Set temporary output path
  lcTempPath = This.CreateWorkingDirectories()
  If Isnull(lcTempPath)
    Return False
  Endif

*-*	Save current settings and then set to US formats
  lcSetPoint  = Set("POINT")
  lcSeparator = Set("SEPARATOR")
  If lcSetPoint != "."
    Set Point To "."
    Set Separator To ","
  Endif

  Dimension laSheetNames[1]
  laSheetNames[1] = tcSheetName
  If This.WriteDirectXMLs(lcTempPath, @laSheetNames)
*-*	Write the sheet XML
    Try
      lcFileName = lcTempPath + "xl\worksheets\sheet1.xml"
      lhFile = Fcreate(lcFileName)
      If lhFile < 0
        Throw
      Endif

*-*		Write sheet XML header information
      Fwrite(lhFile, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
      Fwrite(lhFile, '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"')
      Fwrite(lhFile, ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')
      Fwrite(lhFile, ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"')
      Fwrite(lhFile, ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">')

*-*		Write the sheet view information which includes the freeze pane information
      Fwrite(lhFile, '<dimension ref="A1:' + This.ColumnIndexToAscii(lnFldCnt) + Transform(Reccount(lcAlias)+1) + '"/>')
      If tlFreeze
        Fwrite(lhFile, '<sheetViews><sheetView workbookViewId="0">')
        Fwrite(lhFile, '<pane state="frozenSplit" activePane="bottomLeft" topLeftCell="A2" ySplit="1"/>')
        Fwrite(lhFile, '<selection activeCell="A2" sqref="A2" pane="bottomLeft"/>')
      Endif
      Fwrite(lhFile, '</sheetView></sheetViews>')
      Fwrite(lhFile, '<sheetFormatPr defaultRowHeight="15"/>')

*-*		Write sheet cell data values
      Fwrite(lhFile, '<sheetData>')

*-*		Place field names in the first row
      lcRowDefHt = ' ht="30"'
      Fwrite(lhFile, '<row r="1" spans="1:' + Transform(lnFldCnt) + '"' + lcRowDefHt + '>')
      If This.DefaultFont = "Calibri" And This.DefaultFontSize = 11
        For lnCol=1 To lnFldCnt
          This.WriteAsInLineString(lhFile, 1, lnCol, taFields[lnCol, 2])
        Endfor
      Else
        For lnCol=1 To lnFldCnt
          This.WriteAsInLineFormattedString(lhFile, 1, lnCol, taFields[lnCol, 2])
        Endfor
      Endif
      Fwrite(lhFile, '</row>')

*-*		Write the table field values to cells
      lcSETCentry = Set("CENTURY")
      Set Century On
      lnRow = 1
      If This.DefaultFont = "Calibri" And This.DefaultFontSize = 11
        Scan
          lnRow = lnRow + 1
          Fwrite(lhFile, '<row r="' + Transform(lnRow) + '" spans="1:' + Transform(lnFldCnt) + '"' + lcRowDefHt + '>')
          For lnCol=1 To lnFldCnt
            lxCellValue = Evaluate(lcAlias + '.' + taFields[lnCol, 1])
            This.WriteDirectCellValue(lhFile, lnRow, lnCol, lxCellValue)
          Endfor
          Fwrite(lhFile, '</row>')
        Endscan
      Else
        Scan
          lnRow = lnRow + 1
          Fwrite(lhFile, '<row r="' + Transform(lnRow) + '" spans="1:' + Transform(lnFldCnt) + '"' + lcRowDefHt + '>')
          For lnCol=1 To lnFldCnt
            lxCellValue = Evaluate(lcAlias + '.' + taFields[lnCol, 1])
            This.WriteDirectCellFormattedValue(lhFile, lnRow, lnCol, lxCellValue)
          Endfor
          Fwrite(lhFile, '</row>')
        Endscan
      Endif
      Set Century &lcSETCentry

*-*		End the sheet cell data values and worksheet
      Fwrite(lhFile, '</sheetData>')
      Fwrite(lhFile, '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
      Fwrite(lhFile, '</worksheet>')

    Catch To loException
      Set Step On
      llError = True
      This.ErrorLevelId = 14
      Raiseevent(This, "OnShowErrorMessage", 14, loException.Message)

    Finally
      Fclose(lhFile)
      This.SaveAsUtf8(lcFileName)
    Endtry
    If llError
      llSuccess = False
    Else
*-*		Create the XLSX file
      llSuccess = This.CreateExcelFile(lcTempPath, Justpath(tcXlsxName), Juststem(tcXlsxName))    && Change recommendation by Doug Hennig [placement of code after ELSE]
    Endif
    This.RemoveWorkingDirectories(lcTempPath)
  Else
    llSuccess = False
  Endif
  If !Empty(lcCurAlias)
    Select (lcCurAlias)
  Endif

*-*	Restore default settings
  If lcSetPoint != "."
    Set Point To (lcSetPoint)
    Set Separator To (lcSeparator)
  Endif
  Return llSuccess

*-- Saves the selected workbook as an XLSX file
  Procedure SaveWorkbook
  Lparameters tnWB
  Local lcTempPath, lnSheet, lcDeleted, lcSetPoint, lcSeparator, lcSetDeleted, llReturn, lcAlias
  Local lcSafety, lnTotStages
  With This
    If Seek(BinToC(tnWB), "xl_workbooks", "workbook")
      llReturn     = True
      lcAlias      = Alias()
      lcSafety     = Set("SAFETY")
      lcSetDeleted = Set("DELETED")
      lcSetPoint   = Set("POINT")
      lcSeparator  = Set("SEPARATOR")
      Set Safety Off
      Set Point To "."
      Set Separator To ","
      Set Deleted On

*-*	Write the XLSX Workbook to XML structures
      lcTempPath = .CreateWorkingDirectories()
      If Isnull(lcTempPath)
        llReturn = False
      Else
        Select xl_sheets
        Count For workbook = tnWB And !shdeleted To lnTotStages
        lnTotStages = lnTotStages + 5
        Raiseevent(This, "OnShowStatusMessage", 2, 0, lnTotStages)

        If This.WriteStringsXML(tnWB, lcTempPath) ;
            AND .WriteSupportXMLs(tnWB, lcTempPath) ;
            AND .WriteStylesXML(tnWB, lcTempPath) ;
            AND .WriteWorkbookXML(tnWB, lcTempPath) ;
            AND .WriteRelationshipsXML(tnWB, lcTempPath)

          lnSheet = 0
          Select xl_sheets
          Set Order To Tag Sheet
          Scan For workbook = tnWB And !shdeleted
            lnSheet = lnSheet + 1
            Raiseevent(This, "OnShowStatusMessage", 2, lnSheet+5)
            If !.WriteSheetXMLs(tnWB, xl_sheets.Sheet, lcTempPath)
              llReturn = False
              Exit
            Endif
          Endscan
          Select xl_sheets
          Set Order To

          If llReturn And Seek(BinToC(tnWB), "xl_workbooks", "workbook")
            llReturn = .CreateExcelFile(lcTempPath, xl_workbooks.filepath, xl_workbooks.wbname)
          Endif
        Else
          llReturn = False
        Endif
        .RemoveWorkingDirectories(lcTempPath)
        Raiseevent(This, "OnShowStatusMessage", 2, -1)
      Endif

      Set Safety &lcSafety
      Set Deleted &lcSetDeleted
      Set Point To (lcSetPoint)
      Set Separator To (lcSeparator)
      If !Empty(lcAlias)
        Select (lcAlias)
      Endif
    Else
      llReturn = False
    Endif
  Endwith
  Return llReturn

*-- Saves the selected workbook with the new name
  Procedure SaveWorkbookAs
  Lparameters tnWB, tcWBName
  Local lcName, lcPath
  If Seek(BinToC(tnWB), "xl_workbooks", "workbook") And !Empty(tcWBName)
    lcName = Juststem(tcWBName)
    lcPath = Justpath(tcWBName)
    If Empty(lcPath)
      lcPath = Sys(5) + Sys(2003)
    Endif
    Replace xl_workbooks.wbname   With lcName, ;
      xl_workbooks.filepath With lcPath In xl_workbooks
    This.SaveWorkbook(tnWB)
  Endif

*-- Sets the cell horizontal and vertical alignments
  Procedure SetCellAlignment
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tcHorizAlign, tcVertAlign
  Local llReturn, lnCellXfsId, lcAlias
  With This
    If Pcount() < 4 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      Do Case
      Case Vartype(tcHorizAlign) = 'C' And Vartype(tcVertAlign) = 'C'
        If Inlist(tcHorizAlign, CELL_HORIZ_ALIGN_LEFT, CELL_HORIZ_ALIGN_RIGHT, CELL_HORIZ_ALIGN_CENTER) And ;
            INLIST(tcVertAlign, CELL_VERT_ALIGN_TOP, CELL_VERT_ALIGN_BOTTOM, CELL_VERT_ALIGN_CENTER)

          If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
            If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
              lnCellXfsId = xl_cellxfs.Id
              Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
              If _Tally > 1
                lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, tcHorizAlign, tcVertAlign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
              Else
                Replace xl_cellxfs.halign With tcHorizAlign, ;
                  xl_cellxfs.valign With tcVertAlign In xl_cellxfs
              Endif
              Use In Select('xl_temp')
            Else
              lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, tcHorizAlign, tcVertAlign, 0, 0, 0)
            Endif
            Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
          Else
            lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, tcHorizAlign, tcVertAlign, 0, 0, 0)
            Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted) ;
              VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False)
            This.SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
          Endif
          llReturn = True
        Else
          llReturn = False
        Endif

      Case Vartype(tcHorizAlign) = 'C' And Vartype(tcVertAlign) != 'C'
        If Inlist(tcHorizAlign, CELL_HORIZ_ALIGN_LEFT, CELL_HORIZ_ALIGN_RIGHT, CELL_HORIZ_ALIGN_CENTER)
          If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
            If xl_cells.cellxfs > 0 And This.GetCellXfsRecord(tnWB, xl_cells.cellxfs)
              lnCellXfsId = xl_cellxfs.Id
              Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
              If _Tally > 1
                lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, tcHorizAlign, "", xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
              Else
                Replace xl_cellxfs.halign With tcHorizAlign, ;
                  xl_cellxfs.valign With "" In xl_cellxfs
              Endif
              Use In Select('xl_temp')
            Else
              lnCellXfsId = This.GetCellXfsId(tnWB, 0, 0, 0, 0, tcHorizAlign, "", 0, 0, 0)
            Endif
            Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
          Else
            lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, tcHorizAlign, "", 0, 0, 0)
            Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted) ;
              VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False)
            =.SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
          Endif
          llReturn = True
        Else
          llReturn = False
        Endif

      Case Vartype(tcHorizAlign) != 'C' And Vartype(tcVertAlign) = 'C'
        If Inlist(tcVertAlign, CELL_VERT_ALIGN_TOP, CELL_VERT_ALIGN_BOTTOM, CELL_VERT_ALIGN_CENTER)
          If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
            If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
              lnCellXfsId = xl_cellxfs.Id
              Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
              If _Tally > 1
                lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, "", tcVertAlign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
              Else
                Replace xl_cellxfs.halign With "", ;
                  xl_cellxfs.valign With tcVertAlign In xl_cellxfs
              Endif
              Use In Select('xl_temp')
            Else
              lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", tcVertAlign, 0, 0, 0)
            Endif
            Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
          Else
            lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", tcVertAlign, 0, 0, 0)
            Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted) ;
              VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False)
            SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
          Endif
          llReturn = True
        Else
          llReturn = False
        Endif

      Otherwise
        If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
          Replace xl_cells.halign With "", xl_cells.valign With "" In xl_cells
        Endif
        llReturn = True
      Endcase
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the cell alignment for a range of cells
  Procedure SetCellAlignmentRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tcHorizAlign, tcVertAlign
  Local lnRow, lnCol, llReturn
  With This
    If Pcount() < 7 Or .IsCellReferenceInvalid(tnBegRow, tnBegCol) Or .IsCellReferenceInvalid(tnEndRow, tnEndCol)
      Return False
    Endif
    If tnBegRow > tnEndRow Or tnBegCol > tnEndCol
      Return False
    Endif
    llReturn = True
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If !.SetCellAlignment(tnWB, tnSheet, lnRow, lnCol, tcHorizAlign, tcVertAlign)
          llReturn = False
          Exit
        Endif
      Endfor
    Endfor
  Endwith
  Return llReturn

  Procedure SetCellBorder
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnBorders, tcStyle, tnColor
  Local lcLStyle, lnLColor, lcRStyle, lcTStyle, lnTColor, lcBStyle, lnBColor, lcDStyle, lnDColor, lnDiagDn, lnDiagUp
  Local lnRColor, llReturn, lcAlias
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      If Vartype(tnBorders) != 'N' Or Empty(tnBorders) Or !Between(tnBorders, 0, 63)
        tnBorders = 0
      Endif
      If !.IsValidBorderStyle(tcStyle)
        tcStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnColor) != 'N' Or !Between(tnColor, 0, Rgb(255,255,255)) Or Empty(tnColor)
        tnColor = Rgb(0, 0, 0)
      Endif
      Store 0 To lnDiagDn, lnDiagUp
      Store Rgb(0,0,0) To lnLColor, lnRColor, lnTColor, lnBColor, lnDColor
      Store BORDER_STYLE_NONE To lcLStyle, lcRStyle, lcTStyle, lcBStyle, lcDStyle
      If Bittest(tnBorders, 0)
        lcLStyle = tcStyle
        lnLColor = tnColor
      Endif
      If Bittest(tnBorders, 1)
        lcRStyle = tcStyle
        lnRColor = tnColor
      Endif
      If Bittest(tnBorders, 2)
        lcTStyle = tcStyle
        lnTColor = tnColor
      Endif
      If Bittest(tnBorders, 3)
        lcBStyle = tcStyle
        lnBColor = tnColor
      Endif
      If Bittest(tnBorders, 4)
        lcDStyle = tcStyle
        lnDColor = tnColor
        lnDiagDn = 1
      Endif
      If Bittest(tnBorders, 5)
        lcDStyle = tcStyle
        lnDColor = tnColor
        lnDiagUp = 1
      Endif
      .SetCellBorderEx(tnWB, tnSheet, tnCellRow, tnCellCol, lcLStyle, lnLColor, lcRStyle, lnRColor, ;
        lcTStyle, lnTColor, lcBStyle, lnBColor, lcDStyle, lnDColor, lnDiagDn, lnDiagUp)
      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

  Procedure SetCellBorderEx
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp
  Local lnCellXfsId, lnBorderId
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If .GetSheetRecord(tnWB, tnSheet)
      If !.IsValidBorderStyle(tcLStyle)
        tcLStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnLColor) != 'N' Or Empty(tnLColor) Or !Between(tnLColor, 0, Rgb(255,255,255))
        tnLColor = Rgb(0, 0, 0)
      Endif

      If !.IsValidBorderStyle(tcRStyle)
        tcRStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnRColor) != 'N' Or Empty(tnRColor) Or !Between(tnRColor, 0, Rgb(255,255,255))
        tnRColor = Rgb(0, 0, 0)
      Endif

      If !.IsValidBorderStyle(tcTStyle)
        tcTStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnTColor) != 'N' Or Empty(tnTColor) Or !Between(tnTColor, 0, Rgb(255,255,255))
        tnTColor = Rgb(0, 0, 0)
      Endif

      If !.IsValidBorderStyle(tcBStyle)
        tcBStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnBColor) != 'N' Or Empty(tnBColor) Or !Between(tnBColor, 0, Rgb(255,255,255))
        tnBColor = Rgb(0, 0, 0)
      Endif

      If !.IsValidBorderStyle(tcDStyle)
        tcDStyle = BORDER_STYLE_NONE
      Endif
      If Vartype(tnDColor) != 'N' Or Empty(tnDColor) Or !Between(tnDColor, 0, Rgb(255,255,255))
        tnDColor = Rgb(0, 0, 0)
      Endif
      If Vartype(tnDiagDn) != 'N' Or !Between(tnDiagDn, 0, 3)
        tnDiagDn = 0
      Endif
      If Vartype(tnDiagUp) != 'N' Or !Between(tnDiagUp, 0, 3)
        tnDiagUp = 0
      Endif

*-*	Get the border record Id
      If .GetBordersRecord(tnWB, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp)
        lnBorderId = xl_borders.Id
      Else
        lnBorderId = .GetNextId(tnWB, 'xl_borders')
        Insert Into xl_borders (workbook, Id, lstyle, lcolor, rstyle, rcolor, tstyle, tcolor, bstyle, bcolor, dstyle, dcolor, diagdn, diagup) ;
          VALUES (tnWB, lnBorderId, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp)
      Endif
*-*	Set the XfsId to the cell record
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
          lnCellXfsId = xl_cellxfs.Id
          Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
          If _Tally > 1
            lnCellXfsId = This.GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, lnBorderId, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
          Else
            Replace xl_cellxfs.borderid With lnBorderId In xl_cellxfs
          Endif
          Use In Select('xl_temp')
        Else
          lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, lnBorderId, "", "", 0, 0, 0)
        Endif
        Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
      Else
        lnCellXfsId = This.GetCellXfsId(tnWB, 0, 0, 0, lnBorderId, "", "", 0, 0, 0)
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
        This.SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
  Endwith
  Return llReturn

*-- Sets the cell border for a range of cells
  Procedure SetCellBorderRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tnBorder, tcStyle, tnColor
  Local llReturn, lnRow, lnCol
  With This
    If Pcount() < 7 Or .IsCellReferenceInvalid(tnBegRow, tnBegCol) Or .IsCellReferenceInvalid(tnEndRow, tnEndCol)
      Return False
    Endif
    If tnBegRow > tnEndRow Or tnBegCol > tnEndCol
      Return False
    Endif
    llReturn = True
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If !.SetCellBorder(tnWB, tnSheet, lnRow, lnCol, tnBorder, tcStyle, tnColor)
          llReturn = False
          Exit
        Endif
      Endfor
    Endfor
  Endwith
  Return llReturn

*-- Sets the fill color for the cell
  Procedure SetCellFill
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnFColor, tnBColor, tcPatternType
  Local llReturn, lnFillId, lnCellXfsId, lnFIndex, lnBIndex, lcAlias
  Do Case
  Case Pcount() < 5
    Return False

  Case Pcount() = 5
    tnBColor = Rgb(255,255,255)
    tcPatternType = FILL_STYLE_SOLID

  Case Pcount() = 6
    tcPatternType = FILL_STYLE_SOLID
  Endcase
  With This
    If .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If Pcount() = 5 Or Pcount() = 6 Or Vartype(tcPatternType) != "C" Or Empty(tcPatternType)
      tcPatternType = FILL_STYLE_SOLID
    Endif
    If .GetSheetRecord(tnWB, tnSheet)
      If Vartype(tnBColor) != 'N'
        tnBColor = Rgb(255,255,255)
      Endif
      If Vartype(tnFColor) = 'N' And Between(tnFColor, 0, Rgb(255, 255, 255))
        tnFColor = Floor(tnFColor)
        lnFIndex = 0
      Else
        tnFColor = 0
        lnFIndex = 64
      Endif
      If Vartype(tnBColor) = 'N' And Between(tnBColor, 0, Rgb(255, 255, 255))
        tnBColor = Floor(tnBColor)
        lnBIndex = 0
      Else
        tnBColor = Rgb(0, 0, 0)
        lnBIndex = 64
      Endif
*-*	Get the fill record id
      If Seek(BinToC(tnWB)+BinToC(tnFColor)+BinToC(tnBColor)+tcPatternType, "xl_fills", "fillcolor")
        lnFillId = xl_fills.Id
        Replace xl_fills.Theme     With 0, ;
          xl_fills.tint      With 0, ;
          xl_fills.fgindexed With lnFIndex, ;
          xl_fills.bgindexed With lnBIndex In xl_fills
      Else
        lnFillId = This.GetNextId(tnWB, 'xl_fills')
        Insert Into xl_fills (workbook, Id, fgcolor, bgcolor, patttype, Theme, tint, fgindexed, bgindexed) ;
          VALUES (tnWB, lnFillId, tnFColor, tnBColor, tcPatternType, 0, 0, lnFIndex, lnBIndex)
      Endif
*-*	Set the XfsId to the cell record
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        If xl_cells.cellxfs > 0 And This.GetCellXfsRecord(tnWB, xl_cells.cellxfs)
          lnCellXfsId = xl_cellxfs.Id
          Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
          If _Tally > 1
            lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, lnFillId, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
          Else
            Replace xl_cellxfs.fillid With lnFillId In xl_cellxfs
          Endif
          Use In Select('xl_temp')
        Else
          lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, lnFillId, 0, "", "", 0, 0, 0)
        Endif
        Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, lnFillId, 0, "", "", 0, 0, 0)
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the fill for a range of cells
  Procedure SetCellFillRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tnFColor, tnBColor, tcPatternType
  Local lnRow, lnCol, llReturn
  Do Case
  Case Pcount() < 7
    Return False

  Case Pcount() = 7
    tnBColor = Rgb(255,255,255)
    tcPatternType = FILL_STYLE_SOLID

  Case Pcount() = 8
    tcPatternType = FILL_STYLE_SOLID
  Endcase
  If tnBegRow > tnEndRow Or tnBegCol > tnEndCol Or This.IsCellReferenceInvalid(tnBegRow, tnBegCol) Or This.IsCellReferenceInvalid(tnEndRow, tnEndCol)
    Return False
  Endif
  If Vartype(tnBColor) != 'N'
    tnBColor = Rgb(255,255,255)
  Endif
  llReturn = True
  For lnRow=tnBegRow To tnEndRow
    For lnCol=tnBegCol To tnEndCol
      If !This.SetCellFill(tnWB, tnSheet, lnRow, lnCol, tnFColor, tnBColor, tcPatternType)
        llReturn = False
        Exit
      Endif
    Endfor
  Endfor
  Return llReturn

*-- Sets the selected cell font format
  Procedure SetCellFont
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos
  Local llReturn, lcAlias
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      If Vartype(tcFName) != 'C' Or Empty(tcFName)
        tcFName = 'Calibri'
      Endif
      tcFName = Proper(Alltrim(tcFName))
      If Vartype(tnFSize) != 'N' Or Empty(tnFSize) Or tnFSize < 0.5
        tnFSize = 11
      Else
        tnFSize = Int(tnFSize)
      Endif
      If Vartype(tlBold) != 'L'
        tlBold = False
      Endif
      If Vartype(tlItalic) != 'L'
        tlItalic = False
      Endif
      If Vartype(tnFColor) = 'N' And Between(tnFColor, 0, Rgb(255, 255, 255))
        tnFColor = Floor(tnFColor)
      Else
        tnFColor = Rgb(0, 0, 0)
      Endif
      If Vartype(tlStrikThr) != 'L'
        tlStrikThr = False
      Endif
      If Vartype(tcULine) != 'C' Or !Inlist(tcULine, UNDERLINE_SINGLE, UNDERLINE_DOUBLE, UNDERLINE_SINGLEACCOUNTING, UNDERLINE_DOUBLEACCOUNTING, UNDERLINE_NONE)
        tcULine = UNDERLINE_NONE
      Endif
      If Vartype(tcVPos) != 'C' Or !Inlist(tcVPos, FONT_VERTICAL_BASELINE, FONT_VERTICAL_SUBSCRIPT, FONT_VERTICAL_SUPERSCRIPT)
        tcVPos = FONT_VERTICAL_BASELINE
      Endif
*-*	Get the font record id
      If Seek(BinToC(tnWB)+Padr(tcFName, 100) + Str(tnFSize, 5, 1) + Transform(tlBold) + Transform(tlItalic) + Padl(tnFColor, 15) + Padr(tcULine, 16) + ;
          TRANSFORM(tlStrikThr) + Padr(tcVPos, 11), "xl_fonts", "cellformat")

        lnFontId = xl_fonts.Id
        Replace xl_fonts.Theme   With 0, ;
          xl_fonts.tint    With 0, ;
          xl_fonts.indexed With 0 In xl_fonts
      Else
        lnFontId = .GetNextId(tnWB, 'xl_fonts')
        Insert Into xl_fonts (workbook, Id, fname, Fsize, fbold, fitalic, fcolor, uline, strkthr, fvpos) ;
          VALUES (tnWB, lnFontId, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos)
      Endif
*-*	Set the XfsId to the cell record
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
          lnCellXfsId = xl_cellxfs.Id
          Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
          If _Tally > 1
            lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, lnFontId, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
          Else
            Replace xl_cellxfs.fontid With lnFontId In xl_cellxfs
          Endif
          Use In Select('xl_temp')
        Else
          lnCellXfsId = .GetCellXfsId(tnWB, 0, lnFontId, 0, 0, "", "", 0, 0, 0)
        Endif
        Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, 0, lnFontId, 0, 0, "", "", 0, 0, 0)
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the cell font format for a range of cells
  Procedure SetCellFontRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos
  Local lnRow, lnCol, llReturn
  If Pcount() < 7 Or This.IsCellReferenceInvalid(tnBegRow, tnBegCol) Or This.IsCellReferenceInvalid(tnEndRow, tnEndCol)
    Return False
  Endif
  If tnBegRow > tnEndRow Or tnBegCol > tnEndCol
    Return False
  Endif
  llReturn = True
  For lnRow=tnBegRow To tnEndRow
    For lnCol=tnBegCol To tnEndCol
      If !This.SetCellFont(tnWB, tnSheet, lnRow, lnCol, tcFName, tnFSize, tlBold, tlItalic, tnFColor, tcULine, tlStrikThr, tcVPos)
        llReturn = False
        Exit
      Endif
    Endfor
  Endfor
  Return llReturn

*-- Sets the cell formula expression
  Procedure SetCellFormula
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tcCellFormula
  Local llReturn, lcAlias
  With This
    If Pcount() < 5 Or Vartype(tcCellFormula) != "C" Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        Replace xl_cells.cellvalue   With "", ;
          xl_cells.cellformula With tcCellFormula, ;
          xl_cells.stringid    With 0, ;
          xl_cells.DataType    With DATA_TYPE_FORMULA, ;
          xl_cells.celldeleted With False In xl_cells
      Else
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, cellformula, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, tcCellFormula, DATA_TYPE_FORMULA, 0, False, -1, 0)
        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the cell indentation value
  Procedure SetCellIndent
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnIndent
  Local lnCellXfsId, lcAlias
  With This
    If Pcount() != 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If Vartype(tnIndent) != 'N'
      Return False
    Endif
    lcAlias = Alias()
    tnIndent = Floor(tnIndent)
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
        lnCellXfsId = xl_cellxfs.Id
        Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
        If _Tally > 1
          lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, tnIndent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
        Else
          Replace xl_cellxfs.indent With tnIndent In xl_cellxfs
        Endif
        Use In Select('xl_temp')
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", tnIndent, 0, 0)
      Endif
      Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
    Else
      lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", tnIndent, 0, 0)
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return True

*-- Saves the in-line format definition to a cell value
  Procedure SetCellInLineFormatText
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, toInline
  Local lnNdx, lcRText, llFBold, llFItalic, lnFColor, lcFName, lnFSize, lcULine, llStrkThr, lcFVPos, lcAlias
  With This
    If .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If tnWB = toInline.workbook
      lcAlias = Alias()
      If Isnull(toInline.stringid)
        toInline.stringid = .AddStringValue(toInline.workbook, toInline.StringValue, True)
      Endif
      If .GetCellRecord(toInline.workbook, tnSheet, tnCellRow, tnCellCol)
        Replace xl_cells.cellvalue   With "StringValue", ;
          xl_cells.cellxfs     With 0, ;
          xl_cells.cellformula With "", ;
          xl_cells.stringid    With toInline.stringid, ;
          xl_cells.DataType    With DATA_TYPE_CHAR, ;
          xl_cells.celldeleted With False In xl_cells
      Else
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, cellvalue, stringid, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (toInline.workbook, tnSheet, tnCellRow, tnCellCol, "StringValue", toInline.stringid, DATA_TYPE_CHAR, 0, False, -1, 0)

        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      For lnNdx=1 To toInline.Count
        lcRText   = Substr(toInline.StringValue, toInline.Characters[lnNdx].BegPos, toInline.Characters[lnNdx].Length)
        llFBold   = toInline.Characters[lnNdx].FontBold
        llFItalic = toInline.Characters[lnNdx].FontItalic
        lnFColor  = toInline.Characters[lnNdx].fontcolor
        lcFName   = toInline.Characters[lnNdx].FontName
        lnFSize   = toInline.Characters[lnNdx].FontSize
        lcULine   = toInline.Characters[lnNdx].Underline
        llStrkThr = toInline.Characters[lnNdx].StrikeThru
        lcFVPos   = Icase(toInline.Characters[lnNdx].SubScript, FONT_VERTICAL_SUBSCRIPT, toInline.Characters[lnNdx].SuperScript, FONT_VERTICAL_SUPERSCRIPT, FONT_VERTICAL_BASELINE)

        Insert Into xl_strformat (workbook, Id, Index, stringxml, stringval, fbold, fitalic, fcolor, fname, Fsize, uline, strkthr, fvpos, Theme, tint, indexed) ;
          VALUES (toInline.workbook, toInline.stringid, lnNdx, .GetXMLString(lcRText), lcRText, llFBold, llFItalic, lnFColor, lcFName, lnFSize, lcULine, ;
          llStrkThr, lcFVPos, 0, 0, 0)
      Endfor
      If !Empty(lcAlias)
        Select (lcAlias)
      Endif
      Return True
    Else
      Return False
    Endif
  Endwith

*-- Sets the number of decimals to display
  Procedure SetCellNumberDecimals
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnNumDecimals
  Local lcAlias
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If Vartype(tnNumDecimals) != 'N' Or tnNumDecimals < 0
      Return False
    Endif
    lcAlias = Alias()
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      Replace xl_cells.numdec With tnNumDecimals In xl_cells
    Else
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, nbrfmtid, numdec, celldeleted) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, 0, tnNumDecimals, False)
      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return True

*-- Sets the cell number format
  Procedure SetCellNumberFormat
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnNumFormat
  Local llReturn, lnNumFmtId, lcAlias, loException
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If Vartype(tnNumFormat) != 'N'
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      Do Case
      Case tnNumFormat < START_NUMERIC_FORMAT_ID
        lnNumFmtId = tnNumFormat

      Case tnNumFormat >= START_NUMERIC_FORMAT_ID
        If Seek(BinToC(tnWB)+BinToC(tnNumFormat), "xl_numfmts", "id")
          lnNumFmtId = xl_numFmts.Id
        Else
          If !.AddClassDefinedNumericFormats(tnWB, tnNumFormat)
            Return False
          Endif
          lnNumFmtId = tnNumFormat
        Endif

      Otherwise
        Return False
      Endcase
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
          lnCellXfsId = xl_cellxfs.Id
          Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
          If _Tally > 1
            lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
          Else
            Try
              Replace xl_cellxfs.numfmtid With lnNumFmtId In xl_cellxfs
            Catch To loException
              Set Step On
            Endtry
          Endif
          Use In Select('xl_temp')
        Else
          lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, 0, 0, 0, "", "", 0, 0, 0)
        Endif
        Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, 0, 0, 0, "", "", 0, 0, 0)
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the number format for a range of cells
  Procedure SetCellNumberFormatRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tnNumFormat
  Local lnRow, lnCol, llReturn
  With This
    If Pcount() < 7 Or .IsCellReferenceInvalid(tnBegRow, tnBegCol) Or .IsCellReferenceInvalid(tnEndRow, tnEndCol)
      Return False
    Endif
    If tnBegRow > tnEndRow Or tnBegCol > tnEndCol
      Return False
    Endif
    llReturn = True
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If !.SetCellNumberFormat(tnWB, tnSheet, lnRow, lnCol, tnNumFormat)
          llReturn = False
          Exit
        Endif
      Endfor
    Endfor
  Endwith
  Return llReturn

*-- Sets the cell style
  Procedure SetCellStyle
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnCellXfsId
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    If This.GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      Replace xl_cells.cellxfs With tnCellXfsId In xl_cells
    Else
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, tnCellXfsId, False, -1, 0)
      This.SetRowMaxColumn(tnWB, tnSheet, lnRow, lnCol)
    Endif
    Return True
  Else
    Return False
  Endif

*-- Sets the cell style for a range of cells
  Procedure SetCellStyleRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tnCellXfsId
  Local lnRow, lnCol
  If Pcount() < 7 Or tnBegRow > tnEndRow Or tnBegCol > tnEndCol
    Return False
  Endif
  If Seek(BinToC(tnWB)+BinToC(tnCellXfsId), "xl_cellxfs", "id")
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If This.GetCellRecord(tnWB, tnSheet, lnRow, lnCol)
          Replace xl_cells.cellxfs With tnCellXfsId In xl_cells
        Else
          Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
            VALUES (tnWB, tnSheet, lnRow, lnCol, DATA_TYPE_NONE, tnCellXfsId, False, -1, 0)
          This.SetRowMaxColumn(tnWB, tnSheet, lnRow, lnCol)
        Endif
      Endfor
    Endfor
    Return True
  Else
    Return False
  Endif

*-- Sets the cell text rotation value
  Procedure SetCellTextRotation
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnRotation
  Local lnCellXfsId, lcAlias
  With This
    If Pcount() != 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If Vartype(tnRotation) != 'N'
      Return False
    Endif
    If !Between(tnRotation, -90, 90)
      Return False
    Endif
    lcAlias = Alias()
    tnRotation = Floor(tnRotation)
    If tnRotation < 0
      tnRotation = 90 - tnRotation
    Endif
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
        lnCellXfsId = xl_cellxfs.Id
        Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
        If _Tally > 1
          lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, tnRotation)
        Else
          Replace xl_cellxfs.Rotation With tnRotation In xl_cellxfs
        Endif
        Use In Select('xl_temp')
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", 0, 0, tnRotation)
      Endif
      Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
    Else
      lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", 0, 0, tnRotation)
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return True

*-- Adds a cell validation
  Procedure SetCellValidation
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tnType, tnStyle, tnOperator, tlAllowBlank, tlShowInputMsg, tlShowErrMsg, tcErrMsg, tcErrTitle, tcPrompt, tcFormula1, tcFormula2
  Local lcAlias, lnIndex, llFormula
  lcAlias = Alias()
  With This
    If Pcount() < 14 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    If !Between(tnType, 0, CUSTOM_VALID_TYPE)
      Return False
    Endif
    If Vartype(tcFormula1) != "C"
      tcFormula1 = Transform(tcFormula1)
    Endif
    If Empty(tcFormula1) Or Len(tcFormula1) > 254
      Return False
    Endif
    If Left(tcFormula1, 1) = "="
      llFormula  = True
      tcFormula1 = Alltrim(Substr(tcFormula1, 2))
    Else
      llFormula  = False
    Endif
    If Pcount() < 15
      tcFormula2 = ""
    Endif
    If Vartype(tcFormula2) != "C"
      tcFormula2 = Transform(tcFormula2)
    Endif
    If Len(tcFormula2) > 254
      Return False
    Endif
    If Left(tcFormula2, 1) = "="
      tcFormula2 = Alltrim(Substr(tcFormula2, 2))
    Endif
    If Vartype(tnStyle) = "N"
      If !Between(tnStyle, 0, INFO_VALID_STYLE)
        Return False
      Endif
    Else
      tnStyle = 0
    Endif
    If Vartype(tnOperator) = "N"
      If !Between(tnOperator, 0, GREATOREQUAL_VALID_OPER)
        Return False
      Endif
    Else
      tnOperator = 0
    Endif
    If Vartype(tcErrMsg) != "C"
      tcErrMsg = ""
    Endif
    If Vartype(tcErrTitle) != "C"
      tcErrTitle = ""
    Endif
    If Vartype(tcPrompt) != "C"
      tcPrompt = ""
    Endif
    If Vartype(tlAllowBlank) != "L"
      tlAllowBlank = True
    Endif
    If Vartype(tlShowInputMsg) != "L"
      tlShowInputMsg = True
    Endif
    If Vartype(tlShowErrMsg) != "L"
      tlShowErrMsg = True
    Endif
*-*	Check if the cell is part of a merged range of cells; if so set the cell to the upper left cell of the merged range
    If Reccount('xl_mergecells') > 0
      Select xl_mergecells
      Locate For workbook = tnWB And Sheet = tnSheet And Between(tnCellRow, begrow, endrow) And Between(tnCellCol, begcol, endcol) And !Deleted()
      If Found()
        tnCellRow = xl_mergecells.begrow
        tnCellCol = xl_mergecells.begcol
      Endif
    Endif
    Select xl_validation
    Locate For BinToC(workbook)+BinToC(Sheet) = BinToC(tnWB)+BinToC(tnSheet) ;
      AND Upper(formula1) = Upper(Padr(tcFormula1, 240)) ;
      AND Upper(formula2) = Upper(Padr(tcFormula2, 240)) And !Deleted()
    If Found()
      lnIndex = 0
      Do While Found()
        If xl_validation.vtype = tnType ;
            AND xl_validation.vstyle = tnStyle ;
            AND xl_validation.voperator = tnOperator ;
            AND xl_validation.errmsg = tcErrMsg ;
            AND xl_validation.errtitle = tcErrTitle ;
            AND xl_validation.allowblank = tlAllowBlank ;
            AND xl_validation.showinpmsg = tlShowInputMsg ;
            AND xl_validation.showerrmsg = tlShowErrMsg ;
            AND xl_validation.vprompt = tcPrompt

          lnIndex = xl_validation.validndx
          Exit
        Endif
        Continue
      Enddo
      If lnIndex = 0
        Insert Into xl_validation (workbook, Sheet, vtype, vstyle, voperator, errmsg, errtitle, vprompt, allowblank, showinpmsg, showerrmsg, formula, formula1, formula2) ;
          VALUES (tnWB, tnSheet, tnType, tnStyle, tnOperator, tcErrMsg, tcErrTitle, tcPrompt, tlAllowBlank, tlShowInputMsg, tlShowErrMsg, llFormula, tcFormula1, tcFormula2)
        lnIndex = xl_validation.validndx
      Endif
    Else
      Insert Into xl_validation (workbook, Sheet, vtype, vstyle, voperator, errmsg, errtitle, vprompt, allowblank, showinpmsg, showerrmsg, formula, formula1, formula2) ;
        VALUES (tnWB, tnSheet, tnType, tnStyle, tnOperator, tcErrMsg, tcErrTitle, tcPrompt, tlAllowBlank, tlShowInputMsg, tlShowErrMsg, llFormula, tcFormula1, tcFormula2)
      lnIndex = xl_validation.validndx
    Endif
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      Replace xl_cells.validndx With lnIndex In xl_cells
    Else
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, 0, False, -1, lnIndex)
      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return True

*-- Sets the cell value
  Procedure SetCellValue
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, txCellValue
  Local lcCellValue, lcDataType, lnStringNdx, lnSetDecimals, lnWholePart, lnFractPart, lcAlias, lnNumFmtId, lcHorzAlign
  Local lcSetDate, lcSetPoint, lcSeparator, lcCurrency
  With This
    If Pcount() < 5 Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
*-*	Check if the cell is part of a merged range of cells; if so set the cell to the upper left cell of the merged range
    If Reccount('xl_mergecells') > 0
      Select xl_mergecells
      Locate For workbook = tnWB And Sheet = tnSheet And tnCellRow >= begrow And tnCellRow <= endrow And tnCellCol >= begcol And tnCellCol <= endcol And !Deleted()
      If Found()
        tnCellRow = xl_mergecells.begrow
        tnCellCol = xl_mergecells.begcol
      Endif
    Endif
*-*	Assign the data type based on type of value and add to cell
    lnStringNdx = 0
    lcHorzAlign = ""
    lnNumFmtId  = 0
    lcSetPoint  = Set("POINT")
    lcSeparator = Set("SEPARATOR")
    If lcSetPoint != "."                        && Change recommended by Doug Hennig; 2017-06-12
      Set Point To "."
      Set Separator To ","
    Endif
    Do Case
    Case Vartype(txCellValue) = DATA_TYPE_LOGICAL
      lcDataType  = DATA_TYPE_LOGICAL
      lcCellValue = "StringValue"
      txCellValue = Iif(txCellValue, Getwordnum(.TrueFalseValue, 1, "|"), Getwordnum(.TrueFalseValue, 2, "|"))
      lnStringNdx = .AddStringValue(tnWB, txCellValue, False)

    Case Vartype(txCellValue) = DATA_TYPE_CHAR
      txCellValue = Alltrim(txCellValue)
      If Len(txCellValue) > LIMITS_MAX_CELL_CHARS
        Return False
      Endif
      lcDataType  = DATA_TYPE_CHAR
      lcCellValue = "StringValue"
      lnStringNdx = .AddStringValue(tnWB, txCellValue, False)
      Delete For workbook = tnWB And Id = lnStringNdx In xl_strformat

    Case Vartype(txCellValue) = DATA_TYPE_DATE
      If txCellValue >= Date(1900, 1, 1)
        lcDataType  = DATA_TYPE_DATE
        lcCellValue = Transform(Int(txCellValue - Date(1900, 1, 1) + 2))
      Else
        lcDataType  = DATA_TYPE_CHAR
        lcCellValue = "StringValue"
        lcSetDate   = Set("DATE")
        Set Date To AMERICAN
        lnStringNdx = .AddStringValue(tnWB, Transform(txCellValue), False)
        Set Date To (lcSetDate)
      Endif
      lcHorzAlign = CELL_HORIZ_ALIGN_LEFT
      lnNumFmtId  = CELL_FORMAT_DATE_MMDDYY

    Case Vartype(txCellValue) = DATA_TYPE_DATETIME
      If Ttod(txCellValue) >= Date(1900, 1, 1)
        lnSetDecimals = Set("DECIMALS")
        Set Decimals To 12
        lcDataType  = DATA_TYPE_DATETIME
        lnWholePart = Int(Ttod(txCellValue) - Date(1900, 1, 1) + 2)
        lnFractPart = ((Hour(txCellValue) * 60 + Minute(txCellValue)) * 60 + Sec(txCellValue)) / (24 * 60 * 60)
        lcCellValue = Transform(lnWholePart + lnFractPart)
        Set Decimals To (lnSetDecimals)
      Else
        lcDataType  = DATA_TYPE_CHAR
        lcCellValue = "StringValue"
        lcSetDate   = Set("DATE")
        Set Date To AMERICAN
        lnStringNdx = .AddStringValue(tnWB, Transform(txCellValue), False)
        Set Date To (lcSetDate)
      Endif
      lcHorzAlign = CELL_HORIZ_ALIGN_LEFT
      lnNumFmtId  = CELL_FORMAT_DATETIME_MMMDDYYYY_TTAM

    Case Vartype(txCellValue) = DATA_TYPE_FLOAT
      lcCellValue = Transform(txCellValue)
      If Atc(".", lcCellValue) > 0
        lcDataType = DATA_TYPE_FLOAT
      Else
        lcDataType = DATA_TYPE_INT
      Endif

    Case Vartype(txCellValue) = DATA_TYPE_CURRENCY
      If .SaveCurrencyAsNumeric
        lcDataType  = DATA_TYPE_FLOAT                     && Change requested by Tony Federer
        lcCellValue = Transform(Mton(txCellValue))
      Else
        lcCurrency  = Set("Currency", 1)
        lcCellValue = Chrtran(Transform(txCellValue), lcCurrency+",", "")
        lcDataType  = DATA_TYPE_CURRENCY
        lcHorzAlign = CELL_HORIZ_ALIGN_RIGHT
        Do Case
        Case lcCurrency = "£"
          lnNumFmtId = CELL_FORMAT_ACC_CURR_POUNDS

        Case lcCurrency = "А"
          lnNumFmtId = CELL_FORMAT_ACC_CURR_EURO

        Otherwise
          lnNumFmtId = CELL_FORMAT_CURRENCY_PAREN
        Endcase
      Endif

    Case Vartype(txCellValue) = DATA_TYPE_GENERAL
      lcDataType  = DATA_TYPE_GENERAL
      lcCellValue = ""                                      && Change recommendation by Doug Hennig

    Otherwise
      lcDataType  = DATA_TYPE_GENERAL
      lcCellValue = Transform(txCellValue)
    Endcase
    If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
      If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
        lnCellXfsId = xl_cellxfs.Id
        Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
        If _Tally > 1
          lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, xl_cellxfs.wraptext, xl_cellxfs.Rotation)
        Else
          Replace xl_cellxfs.numfmtid With lnNumFmtId In xl_cellxfs
        Endif
        Use In Select('xl_temp')
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, 0, 0, 0, "", "", 0, 0, 0)
      Endif
      Replace xl_cells.cellvalue   With lcCellValue, ;
        xl_cells.cellxfs     With lnCellXfsId, ;
        xl_cells.cellformula With "", ;
        xl_cells.stringid    With lnStringNdx, ;
        xl_cells.DataType    With lcDataType, ;
        xl_cells.celldeleted With False In xl_cells
    Else
      If lnNumFmtId > 0
        lnCellXfsId = .GetCellXfsId(tnWB, lnNumFmtId, 0, 0, 0, "", "", 0, 0, 0)
      Else
        lnCellXfsId = 0
      Endif
      Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, cellvalue, stringid, DataType, cellxfs, celldeleted, numdec, validndx) ;
        VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, lcCellValue, lnStringNdx, lcDataType, lnCellXfsId, False, -1, 0)

      .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
    If lcSetPoint != "."                        && Change recommended by Doug Hennig; 2017-06-12
      Set Point To (lcSetPoint)
      Set Separator To (lcSeparator)
    Endif
  Endwith
  Return True

*-- Sets the cell word-wrapping value
  Procedure SetCellWordWrap
  Lparameters tnWB, tnSheet, tnCellRow, tnCellCol, tlWordWrap
  Local llReturn, lnWrapText, lnCellXfsId, lcAlias
  With This
    If Pcount() < 5 Or Vartype(tlWordWrap) != "L" Or .IsCellReferenceInvalid(tnCellRow, tnCellCol)
      Return False
    Endif
    lcAlias = Alias()
    If .GetSheetRecord(tnWB, tnSheet)
      lnWrapText = Iif(tlWordWrap, 1, 0)
*-*	Set the XfsId to the cell record
      If .GetCellRecord(tnWB, tnSheet, tnCellRow, tnCellCol)
        If xl_cells.cellxfs > 0 And .GetCellXfsRecord(tnWB, xl_cells.cellxfs)
          lnCellXfsId = xl_cellxfs.Id
          Select workbook From xl_cells Where workbook = tnWB And cellxfs = lnCellXfsId Into Cursor xl_temp
          If _Tally > 1
            lnCellXfsId = .GetCellXfsId(tnWB, xl_cellxfs.numfmtid, xl_cellxfs.fontid, xl_cellxfs.fillid, xl_cellxfs.borderid, xl_cellxfs.halign, xl_cellxfs.valign, xl_cellxfs.indent, lnWrapText, xl_cellxfs.Rotation)
          Else
            Replace xl_cellxfs.wraptext With lnWrapText In xl_cellxfs
          Endif
          Use In Select('xl_temp')
        Else
          lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", 0, lnWrapText, 0)
        Endif
        Replace xl_cells.cellxfs With lnCellXfsId In xl_cells
      Else
        lnCellXfsId = .GetCellXfsId(tnWB, 0, 0, 0, 0, "", "", 0, lnWrapText, 0)
        Insert Into xl_cells (workbook, Sheet, cellrow, cellcol, DataType, cellxfs, celldeleted, numdec, validndx) ;
          VALUES (tnWB, tnSheet, tnCellRow, tnCellCol, DATA_TYPE_NONE, lnCellXfsId, False, -1, 0)
        .SetRowMaxColumn(tnWB, tnSheet, tnCellRow, tnCellCol)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
    If !Empty(lcAlias)
      Select (lcAlias)
    Endif
  Endwith
  Return llReturn

*-- Sets the word wrap for a range of cells
  Procedure SetCellWordWrapRange
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol, tlWordWrap
  Local lnRow, lnCol, llReturn
  With This
    If Pcount() < 7 Or Vartype(tlWordWrap) != "L" Or .IsCellReferenceInvalid(tnBegRow, tnBegCol) Or .IsCellReferenceInvalid(tnEndRow, tnEndCol)
      Return False
    Endif
    If tnBegRow > tnEndRow Or tnBegCol > tnEndCol
      Return False
    Endif
    llReturn = True
    For lnRow=tnBegRow To tnEndRow
      For lnCol=tnBegCol To tnEndCol
        If !.SetCellWordWrap(tnWB, tnSheet, lnRow, lnCol, tlWordWrap)
          llReturn = False
          Exit
        Endif
      Endfor
    Endfor
  Endwith
  Return llReturn

*-- Set the column to best fit
  Procedure SetColumnBestFit
  Lparameters tnWB, tnSheet, tnColumn, tlBestFit
  Local llReturn, lcAlias
  If Pcount() < 4
    Return False
  Endif
  lcAlias = Alias()
  If This.GetSheetRecord(tnWB, tnSheet)
    If Seek(This.GetColWidthKeyExpr(tnWB, tnSheet, tnColumn), "xl_colwidths", "shcolumn")
      Replace xl_colwidths.bestfit With tlBestFit In xl_colwidths
    Else
      Insert Into xl_colwidths (workbook, Sheet, Column, bestfit) Values (tnWB, tnSheet, tnColumn, tlBestFit)
    Endif
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the column width
  Procedure SetColumnWidth
  Lparameters tnWB, tnSheet, tnColumn, tnWidth
  Local llReturn, lcAlias
  If Pcount() < 4 Or Vartype(tnWidth) != "N"
    Return False
  Endif
  lcAlias = Alias()
  If This.GetSheetRecord(tnWB, tnSheet)
    If Seek(This.GetColWidthKeyExpr(tnWB, tnSheet, tnColumn), "xl_colwidths", "shcolumn")
      Replace xl_colwidths.Width With tnWidth In xl_colwidths
    Else
      Insert Into xl_colwidths (workbook, Sheet, Column, Width, bestfit) Values (tnWB, tnSheet, tnColumn, tnWidth, False)
    Endif
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the column width for a range of columns
  Procedure SetColumnWidthRange
  Lparameters tnWB, tnSheet, tnBegCol, tnEndCol, tnWidth
  Local lnCol, llReturn
  If Pcount() < 5
    Return False
  Endif
  If tnBegCol > tnEndCol
    Return False
  Endif
  llReturn = True
  For lnCol=tnBegCol To tnEndCol
    If !This.SetColumnWidth(tnWB, tnSheet, lnCol, tnWidth)
      llReturn = False
      Exit
    Endif
  Endfor
  Return llReturn

*-- Sets the paper size based on custom dimensions
  Procedure SetCustomPaperSize
  Lparameters tnWB, tnSheet, tnWidth, tnHeight, tcDimen
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 5 And This.GetSheetRecord(tnWB, tnSheet)
    Do Case
    Case tnWidth=0 Or tnHeight=0
      Replace xl_sheets.Paperwidth  With 0, ;
        xl_sheets.paperheight With 0, ;
        xl_sheets.paperdimen  With "" In xl_sheets
      llReturn = True

    Case tnWidth>0 And tnHeight>0 And Inlist(tcDimen, "mm", "in")
      Replace xl_sheets.Paperwidth  With tnWidth, ;
        xl_sheets.paperheight With tnHeight, ;
        xl_sheets.paperdimen  With Lower(tcDimen) In xl_sheets
      llReturn = True

    Otherwise
      llReturn = False
    Endcase
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the default border style for the workbook
  Procedure SetDefaultBorder
  Lparameters tnWB, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp
  Local lcAlias
  If Pcount() < 2
    Return False
  Endif
  If Vartype(tcLStyle) != "C"
    tcLStyle = BORDER_STYLE_NONE
  Endif
  lcAlias = Alias()
  Do Case
  Case Pcount() = 2
    tnLColor = 0
    tcRStyle = tcLStyle
    tnRColor = 0
    tcTStyle = tcLStyle
    tnTColor = 0
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 3
    tcRStyle = tcLStyle
    tnRColor = 0
    tcTStyle = tcLStyle
    tnTColor = 0
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 4
    tnRColor = 0
    tcTStyle = tcLStyle
    tnTColor = 0
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 5
    tcTStyle = tcLStyle
    tnTColor = 0
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 6
    tnTColor = 0
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 7
    tcBStyle = tcLStyle
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 8
    tnBColor = 0
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 9
    tcDStyle = tcLStyle
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 10
    tnDColor = 0
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 11
    tnDiagDn = 0
    tnDiagUp = 0

  Case Pcount() = 12
    tnDiagUp = 0

  Endcase
  If Vartype(tnLColor) != "N"
    tnLColor = 0
  Endif
  If Vartype(tcRStyle) != "C"
    tcRStyle = tcLStyle
  Endif
  If Vartype(tnRColor) != "N"
    tnRColor = 0
  Endif
  If Vartype(tcTStyle) != "C"
    tcTStyle = tcLStyle
  Endif
  If Vartype(tnTColor) != "N"
    tnTColor = 0
  Endif
  If Vartype(tcBStyle) != "C"
    tcBStyle = tcLStyle
  Endif
  If Vartype(tnBColor) != "N"
    tnBColor = 0
  Endif
  If Vartype(tcDStyle) != "C"
    tcDStyle = tcLStyle
  Endif
  If Vartype(tnDColor) != "N"
    tnDColor = 0
  Endif
  If Vartype(tnDiagDn) != "N"
    tnDiagDn = 0
  Endif
  If Vartype(tnDiagUp) != "N"
    tnDiagUp = 0
  Endif
  If Seek(BinToC(tnWB)+BinToC(0), "xl_borders", "id")
    Replace xl_borders.lstyle With tcLStyle, ;
      xl_borders.lcolor With tnLColor, ;
      xl_borders.rstyle With tcRStyle, ;
      xl_borders.rcolor With tnRColor, ;
      xl_borders.tstyle With tcTStyle, ;
      xl_borders.tcolor With tnTColor, ;
      xl_borders.bstyle With tcBStyle, ;
      xl_borders.bcolor With tnBColor, ;
      xl_borders.dstyle With tcDStyle, ;
      xl_borders.dcolor With tnDColor, ;
      xl_borders.diagdn With tnDiagDn, ;
      xl_borders.diagup With tnDiagUp In xl_borders
  Else
    Insert Into xl_borders (workbook, Id, lstyle, lcolor, rstyle, rcolor, tstyle, tcolor, bstyle, bcolor, dstyle, dcolor, diagdn, diagup) ;
      VALUES (tnWB, 0, tcLStyle, tnLColor, tcRStyle, tnRColor, tcTStyle, tnTColor, tcBStyle, tnBColor, tcDStyle, tnDColor, tnDiagDn, tnDiagUp)
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return True

*-- Sets the default fill for the workbook
  Procedure SetDefaultFill
  Lparameters tnWB, tcPatternType, tnFColor, tnBColor
  Local lnFIndex, lnBIndex, lcAlias
  If Pcount() < 2
    Return False
  Endif
  lcAlias = Alias()
  If Vartype(tcPatternType) != "C"
    tcPatternType = FILL_STYLE_NONE
  Endif
  Do Case
  Case Pcount() = 2
    tnFColor = 0
    lnFIndex = 0
    tnBColor = 0
    lnBIndex = 0

  Case Pcount() = 3
    tnBColor = 0
    lnBIndex = 64
  Endcase
  If Vartype(tnFColor) != "N"
    tnFColor = 0
  Endif
  If Vartype(tnBColor) != "N"
    tnBColor = Rgb(255,255,255)
  Endif
  If Seek(BinToC(tnWB)+BinToC(0), "xl_fills", "id")
    Replace xl_fills.fgcolor   With tnFColor, ;
      xl_fills.bgcolor   With tnBColor, ;
      xl_fills.patttype  With tcPatternType, ;
      xl_fills.Theme     With 0, ;
      xl_fills.tint      With 0, ;
      xl_fills.fgindexed With 0, ;
      xl_fills.bgindexed With 0 In xl_fills
  Else
    Insert Into xl_fills (workbook, Id, fgcolor, bgcolor, patttype, Theme, tint, fgindexed, bgindexed) ;
      VALUES (tnWB, 0, tnFColor, tnBColor, tcPatternType, 0, 0, lnFIndex, lnBIndex)
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return True

*-- Sets the default font for the workbook
  Procedure SetDefaultFont
  Lparameters tnWB, tcFontName, tnSize, tlBold, tlItalic, tnColor, tcULine, tlStrkthr, tcFVPos, tnTheme, tnTint, tnIndexed
  Local lcAlias
  If Pcount() < 2
    Return False
  Endif
  lcAlias = Alias()
  Do Case
  Case Pcount() = 2
    tnSize    = 11
    tlBold    = False
    tlItalic  = False
    tnColor   = 0
    tcULine   = ""
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 3
    tlBold    = False
    tlItalic  = False
    tnColor   = 0
    tcULine   = ""
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 4
    tlItalic  = False
    tnColor   = 0
    tcULine   = ""
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 5
    tnColor   = 0
    tcULine   = ""
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 6
    tcULine   = ""
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 7
    tlStrkthr = False
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 8
    tcFVPos   = ""
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 9
    tnTheme   = 0
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 10
    tnTint    = 0
    tnIndexed = 0

  Case Pcount() = 11
    tnIndexed = 0
  Endcase
  If Vartype(tcFontName) != "C"
    tcFontName = "Calibri"
  Endif
  If Vartype(tnSize) != "N"
    tnSize = 11
  Endif
  If Vartype(tlBold) != "L"
    tlBold = False
  Endif
  If Vartype(tlItalic) != "L"
    tlItalic = False
  Endif
  If Vartype(tnColor) != "N"
    tnColor = 0
  Endif
  If Vartype(tcULine) != "C"
    tcULine = ""
  Endif
  If Vartype(tlStrkthr) != "L"
    tlStrkthr = False
  Endif
  If Vartype(tcFVPos) != "C"
    tcFVPos = ""
  Endif
  If Vartype(tnTheme) != "N"
    tnTheme = 0
  Endif
  If Vartype(tnTint) != "N"
    tnTint = 0
  Endif
  If Vartype(tnIndexed) != "N"
    tnTint = 0
  Endif
  If Seek(BinToC(tnWB)+BinToC(0), "xl_fonts", "id")
    Replace xl_fonts.fname   With tcFontName, ;
      xl_fonts.Fsize   With tnSize, ;
      xl_fonts.fbold   With tlBold, ;
      xl_fonts.fitalic With tlItalic, ;
      xl_fonts.fcolor  With tnColo , ;
      xl_fonts.uline   With tcULine, ;
      xl_fonts.strkthr With tlStrkthr, ;
      xl_fonts.fvpos   With tcFVPos, ;
      xl_fonts.Theme   With tnTheme, ;
      xl_fonts.tint    With tnTint, ;
      xl_fonts.indexed With tnIndexed In xl_fonts
  Else
    Insert Into xl_fonts (workbook, Id, fname, Fsize, fbold, fitalic, fcolor, uline, strkthr, fvpos, Theme, tint, indexed) ;
      VALUES (tnWB, 0, tcFontName, tnSize, tlBold, tlItalic, tnColor, tcULine, tlStrkthr, tcFVPos, tnTheme, tnTint, tnIndexed)
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return True

*-- Sets the header and footer setup for the sheet
  Procedure SetHeaderFooterSetup
  Lparameters tnWB, tnSheet, tlAlignMargin, tlDiffFirstPg, tlDiffOddEven, tlScaleWDoc
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 6 And This.GetSheetRecord(tnWB, tnSheet)
    If Seek(BinToC(tnWB)+BinToC(tnSheet), "xl_hdrfooterdefn", "wbsheet")
      Replace xl_hdrfooterdefn.alignmargin With tlAlignMargin, ;
        xl_hdrfooterdefn.difffirstpg With tlDiffFirstPg, ;
        xl_hdrfooterdefn.diffoddeven With tlDiffOddEven, ;
        xl_hdrfooterdefn.scalewdoc   With tlScaleWDoc In xl_hdrfooterdefn
    Else
      Insert Into xl_hdrfooterdefn (workbook, Sheet, alignmargin, difffirstpg, diffoddeven, scalewdoc) ;
        VALUES (tnWB, tnSheet, tlAlignMargin, tlDiffFirstPg, tlDiffOddEven, tlScaleWDoc)
    Endif
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the header/footer text for the sheet
  Procedure SetHeaderFooterText
  Lparameters tnWB, tnSheet, tnPage, tnSection, tcText, tcFontName, tnFontSize, tnFontEffect, tnFontColor
  Local llReturn, lcAlias
  Do Case
  Case Pcount() < 5
    Return False

  Case Pcount() = 5
    tcFontName   = ""
    tnFontSize   = 0
    tnFontEffect = 0
    tnFontColor  = 0

  Case Pcount() = 6
    tnFontSize   = 0
    tnFontEffect = 0
    tnFontColor  = 0

  Case Pcount() = 7
    tnFontEffect = 0
    tnFontColor  = 0

  Case Pcount() = 8
    tnFontColor = 0
  Endcase
  If Vartype(tcFontName) != "C"
    tcFontName = ""
  Endif
  If Vartype(tnFontSize) != "N"
    tnFontSize = 0
  Endif
  If Vartype(tnFontEffect) != "N"
    tnFontEffect = 0
  Endif
  If Vartype(tnFontColor) != "N"
    tnFontColor = 0
  Endif
  lcAlias = Alias()
  If This.GetSheetRecord(tnWB, tnSheet)
    If Seek(BinToC(tnWB)+BinToC(tnSheet), "xl_hdrfooterdefn", "wbsheet")
      If Seek(This.GetHeaderFooterKeyExpr(tnWB, tnSheet, tnPage, tnSection), "xl_hdrfootertext", "hdrfootr")
        Replace xl_hdrfootertext.Text      With tcText, ;
          xl_hdrfootertext.FontName  With tcFontName, ;
          xl_hdrfootertext.FontSize  With tnFontSize, ;
          xl_hdrfootertext.fontstyle With tnFontEffect, ;
          xl_hdrfootertext.fontcolor With tnFontColor In xl_hdrfootertext
      Else
        Insert Into xl_hdrfootertext (workbook, Sheet, Page, section, Text, FontName, FontSize, fontstyle, fontcolor) ;
          VALUES (tnWB, tnSheet, tnPage, tnSection, tcText, tcFontName, tnFontSize, tnFontEffect, tnFontColor)
      Endif
      llReturn = True
    Else
      llReturn = False
    Endif
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the last used Id Value for strings, borders, fills, fonts, and cellxfs
  Protected Procedure SetLastId
    Lparameters tnWB, tnLastId, tcCursor
    Local lnNdx
    With This
      Do Case
      Case tcCursor = "xl_borders"
        lnNdx = Ascan(.LastBorderId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastBorderId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1, Alen(.LastBorderId, 1) + 1)
          Dimension .LastBorderId[lnNdx, 2]
          .LastBorderId[lnNdx, 1] = tnWB
          .LastBorderId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_cellxfs"
        lnNdx = Ascan(.LastCellXfsId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastCellXfsId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1, Alen(.LastCellXfsId, 1) + 1)
          Dimension .LastCellXfsId[lnNdx, 2]
          .LastCellXfsId[lnNdx, 1] = tnWB
          .LastCellXfsId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_fills"
        lnNdx = Ascan(.LastFillId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastFillId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastFillId, 1) + 1)
          Dimension .LastFillId[lnNdx, 2]
          .LastFillId[lnNdx, 1] = tnWB
          .LastFillId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_fonts"
        lnNdx = Ascan(.LastFontId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastFontId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastFontId, 1) + 1)
          Dimension .LastFontId[lnNdx, 2]
          .LastFontId[lnNdx, 1] = tnWB
          .LastFontId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_ndxcolors"
        lnNdx = Ascan(.LastIndexColorId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastIndexColorId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1, Alen(.LastIndexColorId, 1) + 1)
          Dimension .LastIndexColorId[lnNdx, 2]
          .LastIndexColorId[lnNdx, 1] = tnWB
          .LastIndexColorId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_mrucolors"
        lnNdx = Ascan(.LastMruColorId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastMruColorId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastMruColorId, 1) + 1)
          Dimension .LastMruColorId[lnNdx, 2]
          .LastMruColorId[lnNdx, 1] = tnWB
          .LastMruColorId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_numfmts"
        lnNdx = Ascan(.LastNumFmtId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastNumFmtId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastNumFmtId, 1) + 1)
          Dimension .LastNumFmtId[lnNdx, 2]
          .LastNumFmtId[lnNdx, 1] = tnWB
          .LastNumFmtId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_relationships"
        lnNdx = Ascan(.LastRelationId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastRelationId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastRelationId, 1) + 1)
          Dimension .LastRelationId[lnNdx, 2]
          .LastRelationId[lnNdx, 1] = tnWB
          .LastRelationId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_sheets"
        lnNdx = Ascan(.LastSheetId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastSheetId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastSheetId, 1) + 1)
          Dimension .LastSheetId[lnNdx, 2]
          .LastSheetId[lnNdx, 1] = tnWB
          .LastSheetId[lnNdx, 2] = tnLastId
        Endif

      Case tcCursor = "xl_strings"
        lnNdx = Ascan(.LastStringId, tnWB, 1, -1, 1, 15)
        If lnNdx > 0
          .LastStringId[lnNdx, 2] = tnLastId
        Else
          lnNdx = Max(1,  Alen(.LastStringId, 1) + 1)
          Dimension .LastStringId[lnNdx, 2]
          .LastStringId[lnNdx, 1] = tnWB
          .LastStringId[lnNdx, 2] = tnLastId
        Endif
      Endcase
    Endwith

*-- Sets the paper size for the selected sheet
  Procedure SetPaperSize
  Lparameters tnWB, tnSheet, tnPaperSize
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    If Between(tnPaperSize, 0, 118)
      Replace xl_sheets.Papersize With tnPaperSize In xl_sheets
      llReturn = True
    Else
      llReturn = False
    Endif
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Number of vertical pages to fit on for printing
  Procedure SetPrintFitToHeight
  Lparameters tnWB, tnSheet, tnFitToHeight
  Local llReturn
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    If Int(tnFitToHeight) >= 0
      Replace xl_sheets.fittoheight With Int(tnFitToHeight) In xl_sheets
      llReturn = True
    Else
      llReturn = False
    Endif
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Number of horizontal pages to fit on for printing
  Procedure SetPrintFitToWidth
  Lparameters tnWB, tnSheet, tnFitToWidth
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    If Int(tnFitToWidth) >= 0
      Replace xl_sheets.fittowidth With Int(tnFitToWidth) In xl_sheets
      llReturn = True
    Else
      llReturn = False
    Endif
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the printer orientation for sheet output
  Procedure SetPrintOrientation
  Lparameters tnWB, tnSheet, tnOrientation
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    Replace xl_sheets.prnorient With tnOrientation In xl_sheets
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the row height
  Procedure SetRowHeight
  Lparameters tnWB, tnSheet, tnRow, tnHeight
  Local llReturn, lcAlias
  If Pcount() < 4 Or Vartype(tnHeight) != "N"
    Return False
  Endif
  lcAlias = Alias()
  If This.GetSheetRecord(tnWB, tnSheet)
    If Seek(This.GetRowHeightKeyExpr(tnWB, tnSheet, tnRow), "xl_rowheights", "shrow")
      Replace xl_rowheights.Height With tnHeight In xl_rowheights
    Else
      Insert Into xl_rowheights (workbook, Sheet, Row, Height) Values (tnWB, tnSheet, tnRow, tnHeight)
    Endif
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the row height for a range of rows
  Procedure SetRowHeightRange
  Lparameters tnWB, tnSheet, tnBegRow, tnEndRow, tnHeight
  Local lnRow, llReturn
  If Pcount() < 5
    Return False
  Endif
  If tnBegRow > tnEndRow
    Return False
  Endif
  llReturn = True
  For lnRow=tnBegRow To tnEndRow
    If !This.SetRowHeight(tnWB, tnSheet, lnRow, tnHeight)
      llReturn = False
      Exit
    Endif
  Endfor
  Return llReturn

*-- Sets the value for the maximum cell column in a row
  Protected Procedure SetRowMaxColumn
    Lparameters tnWB, tnSheet, tnCellRow, tnCellCol
    If Seek(BinToC(tnWB)+BinToC(tnSheet)+BinToC(tnCellRow), "xl_rows", "id")
      If xl_rows.maxcol < tnCellCol
        Replace xl_rows.maxcol With tnCellCol In xl_rows
      Endif
    Else
      Insert Into xl_rows (workbook, Sheet, Row, maxcol) Values (tnWB, tnSheet, tnCellRow, tnCellCol)
    Endif

*-- Sets the sheet margin values
  Procedure SetSheetMargins
  Lparameters tnWB, tnSheet, tnLeft, tnRight, tnTop, tnbot, tnHeader, tnFooter
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 8 And This.GetSheetRecord(tnWB, tnSheet)
    Replace xl_sheets.mleft   With tnLeft, ;
      xl_sheets.mright  With tnRight, ;
      xl_sheets.mtop    With tnTop, ;
      xl_sheets.mbot    With tnbot, ;
      xl_sheets.mheader With tnHeader, ;
      xl_sheets.mfooter With tnFooter In xl_sheets
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the print scale; must be between 10 and 400; i.e. 10=10%, 100=100%, etc.
  Procedure SetSheetScale
  Lparameters tnWB, tnSheet, tnScale
  Local llReturn, lcAlias
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    If Between(Int(tnScale), 10, 400)
      Replace xl_sheets.Scale With Int(tnScale) In xl_sheets
      llReturn = True
    Else
      llReturn = False
    Endif
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Sets the sheet visibility
  Procedure SetSheetVisibility
  Lparameters tnWB, tnSheet, tnVisiblity
  Local llReturn, lcAlias
  llReturn = False
  lcAlias = Alias()
  If Pcount() = 3 And This.GetSheetRecord(tnWB, tnSheet)
    If Between(tnVisiblity, VISIBLE_SHEET_STATE, VERYHIDDEN_SHEET_STATE)
      Replace xl_sheets.state With tnVisiblity In xl_sheets
      llReturn = True
    Endif
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn
*-- Removes the pane freeze from the sheet
  Procedure UnFreezePanes
  Lparameters tnWB, tnSheet
  Local llReturn, lcAlias
  lcAlias = Alias()
  If This.GetSheetRecord(tnWB, tnSheet)
    Replace xl_sheets.xsplit With 0, ;
      xl_sheets.ysplit With 0 In xl_sheets
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Unmerges merged cells
  Procedure UnMergedCells
  Lparameters tnWB, tnSheet, tnBegRow, tnBegCol, tnEndRow, tnEndCol
  Local llReturn, lcAlias
  lcAlias = Alias()
  Select xl_mergecells
  Locate For workbook=tnWB And Sheet=tnSheet And begrow=tnBegRow And begcol=tnBegCol And endrow=tnEndRow And endcol=tnEndCol And !Deleted()
  If Found()
    Delete In xl_mergecells
    llReturn = True
  Else
    llReturn = False
  Endif
  If !Empty(lcAlias)
    Select (lcAlias)
  Endif
  Return llReturn

*-- Writes the string as formatted in-line
  Protected Procedure WriteAsInLineFormattedString
    Lparameters tnRow, tnCol, txCellValue
    With This
      \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="inlineStr">
      \\<is>
      \\<r>
      \\<rPr>
      \\<sz val="<<Transform(.DefaultFontSize)>>"/>
      \\<rFont val="<<.DefaultFont>>"/>
      \\</rPr>
      \\<t><<.GetXMLString(txCellValue)>></t>
      \\</r>
      \\</is>
      \\</c>
    Endwith

*-- Writes the value as an inlineStr value
  Protected Procedure WriteAsInLineString
    Lparameters tnRow, tnCol, txCellValue
    With This
      \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="inlineStr">
      \\<is><t><<.GetXMLString(txCellValue)>></t></is>
      \\</c>
    Endwith

*-- WriteDirectCellFormattedValue
  Protected Procedure WriteDirectCellFormattedValue
    Lparameters tnRow, tnCol, txCellValue
    With This
      Do Case
      Case Isnull(txCellValue)
        Return

      Case Inlist(Vartype(txCellValue), "C", "D", "T") And Empty(txCellValue)
        Return

      Case Vartype(txCellValue) = DATA_TYPE_DATE
        .WriteAsInLineFormattedString(tnRow, tnCol, Transform(txCellValue))

      Case Vartype(txCellValue) = DATA_TYPE_DATETIME
        .WriteAsInLineFormattedString(tnRow, tnCol, Ttoc(txCellValue))

      Case Vartype(txCellValue) = DATA_TYPE_LOGICAL
        txCellValue = Iif(txCellValue, Getwordnum(.TrueFalseValue, 1, "|"), Getwordnum(.TrueFalseValue, 2, "|"))
        .WriteAsInLineFormattedString(tnRow, tnCol, txCellValue)

      Case Vartype(txCellValue) = DATA_TYPE_CURRENCY
        \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="n">
        \\<v><<Transform(Mton(txCellValue))>></v></c>

      Case Vartype(txCellValue) = DATA_TYPE_CHAR
        .WriteAsInLineFormattedString(tnRow, tnCol, Alltrim(txCellValue))

      Case Inlist(Vartype(txCellValue), DATA_TYPE_INT, DATA_TYPE_FLOAT)
        \<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="n">
        \<v><<Transform(txCellValue)>></v></c>

      Otherwise
        \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="inlineStr">
        \\<is><t>Unsupported Field Type</t></is></c>
      Endcase
    Endwith

*-- Writes the cell value directly to the sheets xml
  Protected Procedure WriteDirectCellValue
    Lparameters tnRow, tnCol, txCellValue
    With This
      Do Case
      Case Isnull(txCellValue)
        Return

      Case Inlist(Vartype(txCellValue), "C", "D", "T") And Empty(txCellValue)
        Return

      Case Vartype(txCellValue) = DATA_TYPE_DATE
        .WriteAsInLineString(tnRow, tnCol, Transform(txCellValue))

      Case Vartype(txCellValue) = DATA_TYPE_DATETIME
        .WriteAsInLineString(thFile, tnRow, tnCol, Ttoc(txCellValue))

      Case Vartype(txCellValue) = DATA_TYPE_LOGICAL
        txCellValue = Iif(txCellValue, Getwordnum(.TrueFalseValue, 1, "|"), Getwordnum(.TrueFalseValue, 2, "|"))
        .WriteAsInLineString(thFile, tnRow, tnCol, txCellValue)

      Case Vartype(txCellValue) = DATA_TYPE_CURRENCY
        \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="n">
        \\<v><<Transform(Mton(txCellValue))>></v></c>

      Case Vartype(txCellValue) = DATA_TYPE_CHAR
        .WriteAsInLineString(thFile, tnRow, tnCol, Alltrim(txCellValue))

      Case Inlist(Vartype(txCellValue), DATA_TYPE_INT, DATA_TYPE_FLOAT)
        \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="n">
        \\<v><<Transform(txCellValue)>></v></c>

      Otherwise
        \\<c r="<<.ColumnIndexToAscii(tnCol) + Transform(tnRow)>>" t="inlineStr">')
        \\<is><t>Unsupported Field Type</t></is></c>')
      Endcase
    Endwith

*-- Writes the workbook XMLs for direct table/grid to workbook output
  Protected Procedure WriteDirectXMLs
    Lparameters tcPath, taSheetNames
    Local lcFileName, loException, lcSheetName, lnSheet, lnNumSheets
    External Array taSheetNames
*-*	Create the supporting XML files
    With This
      loException = Null
      Try
        lnNumSheets = Alen(taSheetNames, 1)
      Catch To loException
      Endtry
      If !Isnull(loException)
        Return False
      Endif
      Try
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        \\<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        \\<Default Extension="xml" ContentType="application/xml"/>
        \\<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>
        \\<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>
        For lnSheet=1 To lnNumSheets
          \\<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet' + Transform(lnSheet) + '.xml"/>
        Endfor
        \\<Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>
        \\<Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>
        \\</Types>
        Set Textmerge To
        Strtofile(lcOutStr, (tcPath + "[Content_Types].xml"))

        Set Textmerge On To Memvar lcOutStr Noshow
*-*	Save the .rels XML
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        \\<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        \\<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        \\</Relationships>
        Set Textmerge To
        Strtofile(lcOutStr, (tcPath + "_rels\.rels"))

*-*	Write the workbook relationships file
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        For lnSheet=1 To lnNumSheets
          \\<Relationship Id="rId' + Transform(lnSheet+1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + Transform(lnSheet) + '.xml"/>
        Endfor
        \\</Relationships>
        Set Textmerge To
        Strtofile(lcOutStr, (tcPath + "xl\_rels\workbook.xml.rels"))

*-*	Save the app XML
        Set Textmerge On To Memvar lcOutStr Noshow
        lcFileName = tcPath + "docProps\app.xml"
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        \\<Application>ExcelXMLX Release ' + Transform(.ExcelXlsxRelease) + '</Application>
        If !Empty(.CompanyName)
          \\<Company>' + .CompanyName + '</Company>
        Endif
        \\</Properties>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

*-*	Save the core XML
        Set Textmerge On To Memvar lcOutStr Noshow
        lcFileName = tcPath + "docProps\core.xml"
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<cp:coreProperties
        \\xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        \\xmlns:dc="http://purl.org/dc/elements/1.1/"
        \\xmlns:dcterms="http://purl.org/dc/terms/"
        \\xmlns:dcmitype="http://purl.org/dc/dcmitype/"
        \\xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        If !Empty(.Title)
          \\<dc:title>' + .GetXMLString(.Title) + '</dc:title>
        Endif
        If !Empty(.Subject)
          \\<dc:subject>' + .GetXMLString(.Subject) + '</dc:subject>
        Endif
        If !Empty(.CreatorName)
          \\<dc:creator>' + .GetXMLString(.CreatorName) + '</dc:creator>
        Endif
        If !Empty(.UserName)
          \\<cp:lastModifiedBy>' + .GetXMLString(.UserName) + '</cp:lastModifiedBy>
        Endif
        \\<dcterms:created xsi:type="dcterms:W3CDTF">' + Ttoc(Datetime(), 3) + 'Z</dcterms:created>
        \\<dcterms:modified xsi:type="dcterms:W3CDTF">' + Ttoc(Datetime(), 3) + 'Z</dcterms:modified>
        \\</cp:coreProperties>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

*-*	Save the styles XML
        lcFileName = tcPath + "xl\styles.xml"
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        \\mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
        \\<fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font></fonts>
        \\<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
        \\<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
        \\<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
        \\<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
        \\<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
        \\</styleSheet>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

*-*	Write the workbook XML
        Set Textmerge On To Memvar lcOutStr Noshow
        lcFileName = tcPath + "xl\workbook.xml"
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
        \\<sheets>
        For lnSheet=1 To lnNumSheets
          lcSheetName = Iif(Empty(taSheetNames[lnSheet]), "Sheet" + Transform(lnSheet), taSheetNames[lnSheet])
          \\<sheet name="' + lcSheetName + '" sheetId="' + Transform(lnSheet) + '" r:id="rId' + Transform(lnSheet+1) + '"/>
        Endfor
        \\</sheets>
        \\<calcPr calcId="0"/>
        \\</workbook>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

      Catch To loException
        Set Step On
        .ErrorLevelId = 18
        Raiseevent(This, "OnShowErrorMessage", 18, "Unable to create workbook supporting XMLs")
      Endtry
    Endwith
    Return Isnull(loException)

*-- Writes the relationships xml file
  Protected Procedure WriteRelationshipsXML
    Lparameters tnWB, tcPath
    Local lcBaseType, llSuccess
    With This
      Raiseevent(This, "OnShowStatusMessage", 2, 5)
      lcBaseType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
      Try
*-*	Write the workbook relationships file
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">

        Select xl_relationships
        Scan For workbook = tnWB
          \\<Relationship Id="rId' + Transform(xl_relationships.relid) + '" Target="' + Alltrim(xl_relationships.Target) + '" Type="' + lcBaseType + Alltrim(xl_relationships.reltype) + '"/>
        Endscan

        \\</Relationships>
        llSuccess = True
        Set Textmerge To
        .SaveAsUtf8(tcPath + "xl\_rels\workbook.xml.rels", lcOutStr)

      Catch To loException
        Set Step On
        llSuccess = False
        .ErrorLevelId = 16
        Raiseevent(This, "OnShowErrorMessage", 16, "Unable to create workbook.xml.rels")
      Endtry
    Endwith
    Return llSuccess

*-- Writes the Sheet XML files
  Protected Procedure WriteSheetXMLs
    Lparameters tnWB, tnSheetNbr, tcPath
    Local llSuccess, lcAlias, lhFile, lcRow, lnMinCol, lnMaxCol, lnWidth, lnLastRow, lcDyDescent
    Local lcRowDefHt, lcFileName
*-*	Write the sheet XML
    With This
      lcAlias = Alias()
*-*	Write the sheet
      Try
        Set Textmerge On To Memvar lcOutStr Noshow
        lcFileName = tcPath + "xl\worksheets\sheet" + Transform(tnSheetNbr) + ".xml"
*	lcDyDescent = ' x14ac:dyDescent="0.25"'                && Uncomment to add back...
        lcDyDescent = ''
        lcRowDefHt  = ' ht="30"'

*-*	Write sheet XML header information
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\ xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        \\ xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac"
        \\ xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

*-*	Write the tab color definition if set
        Do Case
        Case xl_sheets.tabcolorndx > 0
          \\<sheetPr codeName="Sheet<<Transform(tnSheetNbr)>>"><tabColor indexed="<<Transform(xl_sheets.tabcolorndx)>>"/></sheetPr>

        Case !Empty(xl_sheets.tabcolorrgb)
          \\<sheetPr codeName="Sheet<<Transform(tnSheetNbr)>>"><tabColor rgb="<<xl_sheets.tabcolorrgb>>"/></sheetPr>
        Endcase

*-*	Verify cells have been written
        Select xl_cells
        Locate For workbook = tnWB And Sheet = xl_sheets.Sheet And !celldeleted
        If Found()
*-*		Write the sheet view information which includes the freeze pane information
          Select cel.Sheet, Max(cel.cellrow) As maxrow, Max(cel.cellcol) As maxcol ;
            FROM xl_cells As cel ;
            WHERE cel.workbook = tnWB ;
            AND cel.Sheet = xl_sheets.Sheet ;
            AND cel.celldeleted = False ;
            GROUP By 1 ;
            INTO Cursor t_max Readwrite

          \\<dimension ref="A1:<<.ColumnIndexToAscii(t_max.maxcol) + Transform(t_max.maxrow)>>"/>
          Do Case
          Case xl_sheets.ysplit > 0 And xl_sheets.xsplit = 0      && Freeze top row(s) only
            \\<sheetViews><sheetView workbookViewId="0">
            lcRow = Transform(xl_sheets.ysplit+1)
            \\<pane state="frozenSplit" activePane="bottomLeft" topLeftCell="A<<lcRow>>" ySplit="<<Transform(xl_sheets.ysplit)>>"/>
            \\<selection activeCell="A<<lcRow>>" sqref="A<<lcRow>>" pane="bottomLeft"/>

          Case xl_sheets.ysplit = 0 And xl_sheets.xsplit > 0      && Freeze side column(s) only
            \\<sheetViews><sheetView workbookViewId="0">
            lcColumn = .ColumnIndexToAscii(xl_sheets.xsplit+1)
            \\<pane state="frozenSplit" activePane="bottomLeft" topLeftCell="<<lcColumn>>1" xSplit="<<Transform(xl_sheets.xsplit)>>"/>
            \\<selection activeCell="<<lcColumn>>1" sqref="<<lcColumn>>1" pane="topRight"/>

          Case xl_sheets.ysplit > 0 And xl_sheets.xsplit > 0      && Freeze both top row(s) and side column(s)
            \\<sheetViews><sheetView workbookViewId="0">
            lcRow = Transform(xl_sheets.ysplit+1)
            lcColumn = .ColumnIndexToAscii(xl_sheets.xsplit+1)
            \\<pane state="frozenSplit" activePane="bottomRight" topLeftCell="<<lcColumn + lcRow>>" ySplit="<<Transform(xl_sheets.ysplit)>>" xSplit="<<Transform(xl_sheets.xsplit)>>"/>
            \\<selection activeCell="<<.ColumnIndexToAscii(xl_sheets.xsplit)>>1" sqref="<<.ColumnIndexToAscii(xl_sheets.xsplit)>>1" pane="topRight"/>
            \\<selection activeCell="<<.ColumnIndexToAscii(xl_sheets.xsplit) + lcRow>>" sqref="<<.ColumnIndexToAscii(xl_sheets.xsplit) + lcRow>>" pane="bottomLeft"/>
            \\<selection activeCell="<<lcColumn + lcRow>>" sqref="<<lcColumn + lcRow>>" pane="bottomRight"/>

          Otherwise
            \\<sheetViews><sheetView workbookViewId="0"><selection activeCell="A1" sqref="A1"/>
*				FWRITE(lhFile, '<sheetViews><sheetView tabSelected="1" workbookViewId="0">')
          Endcase
          \\</sheetView></sheetViews>
          \\<sheetFormatPr defaultRowHeight="15"<<lcDyDescent>>/>

*-*		Write sheet Column widths
          Select clm.Column, clm.Width, clm.bestfit From xl_colwidths As clm ;
            WHERE clm.workbook = tnWB ;
            AND clm.Sheet = xl_sheets.Sheet ;
            ORDER By clm.Column ;
            INTO Cursor t_colwidths Readwrite

          If _Tally > 0
            \\<cols>
            Select t_colwidths
            Goto Top
            lnMinCol = t_colwidths.Column
            lnMaxCol = t_colwidths.Column
            lnWidth  = t_colwidths.Width
            Skip In t_colwidths
            Do While !Eof('t_colwidths')
              Do Case
*					CASE t_colwidths.bestfit    && Not yet working!!!
*						FWRITE(lhFile, '<col min="' + TRANSFORM(t_colwidths.column) + '" max="' + TRANSFORM(t_colwidths.column) + '" bestFit="1" customWidth="1"/>')

              Case t_colwidths.Width = lnWidth And t_colwidths.Column-1 = lnMinCol
                lnMaxCol = t_colwidths.Column

              Otherwise
                \\<col min="<<Transform(lnMinCol)>>max="<<Transform(lnMaxCol)>>" width="<<Transform(lnWidth+0.7109375)>>" customWidth="1"/>
                lnMinCol = t_colwidths.Column
                lnMaxCol = t_colwidths.Column
                lnWidth  = t_colwidths.Width
              Endcase
              Skip In t_colwidths
            Enddo
            \\<col min="<<Transform(lnMinCol)>>" max="<<Transform(lnMaxCol)>>" width="<<Transform(lnWidth+0.7109375)>>" customWidth="1"/>
            \\</cols>
          Endif
          Use In Select('t_colwidths')

*-*		Write sheet cell data values
          \\<sheetData>
*		lnLastRow = 0
          lnHeadRow = 0
          Select xl_cells
          Set Order To Tag cellindex
          Scan For workbook = tnWB And Sheet = xl_sheets.Sheet And !celldeleted
*-*			Write row header
            If lnHeadRow != xl_cells.cellrow
              If lnHeadRow != 0
                \\</row>
              Endif
              lnHeadRow = xl_cells.cellrow
              lnMaxCol  = .GetRowMaxColumn(tnWB, xl_sheets.Sheet, lnHeadRow)
              If Seek(.GetRowHeightKeyExpr(tnWB, xl_sheets.Sheet, xl_cells.cellrow), "xl_rowheights", "shrow")
                \\<row r="<<Transform(xl_cells.cellrow)>>" spans="1:<<Transform(lnMaxCol)>>" ht="<<Transform(xl_rowheights.Height)>>" customHeight="1">
              Else
                \\<row r="<<Transform(xl_cells.cellrow)>>" spans="1:<<Transform(lnMaxCol)>>"<<lcRowDefHt>>>
              Endif
            Endif
*-*			Write columns
            lcStyle = Iif(xl_cells.cellxfs > 0, ' s="' + Transform(xl_cells.cellxfs) + '"', '')
            Do Case
            Case xl_cells.DataType = DATA_TYPE_TIME
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \\<v><<Alltrim(xl_cells.cellvalue)>></v></c>')

            Case xl_cells.DataType = DATA_TYPE_DATE
              \<c r="<<ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \<v><<Alltrim(xl_cells.cellvalue)>></v></c>')

            Case xl_cells.DataType = DATA_TYPE_DATETIME
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \\<v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_LOGICAL
              \<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>> t="s">
              \\<v><<Transform(xl_cells.stringid)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_CHAR
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>> t="s">
              \\<v><<Transform(xl_cells.stringid)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_INT
              \<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \<v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_FLOAT
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \\<v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_CURRENCY
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>>
              \\<v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_PERCENT
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>"<<lcStyle >>>
              \\<v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_FORMULA
              \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>> t="str">
              \\<f><<Alltrim(xl_cells.cellformula)>></f><v><<Alltrim(xl_cells.cellvalue)>></v></c>

            Case xl_cells.DataType = DATA_TYPE_NONE
              If !Empty(lcStyle)
                \\<c r="<<.ColumnIndexToAscii(xl_cells.cellcol) + Transform(xl_cells.cellrow)>>"<<lcStyle>>></c>
              Endif
            Endcase
          Endscan
          \\</row>
          \\</sheetData>
          Set Order To
          Use In t_max

        Else
*-*		No cell data has been created/written
          \\<dimension ref="A1"/>
          \\<sheetViews>
          \\<sheetView workbookViewId="0"/>
          \\</sheetViews>
          \\<sheetFormatPr defaultRowHeight="14.4"/>
          \\<sheetData/>
        Endif
*-*	Write sheet merge data
        Select mrg.begrow, mrg.begcol, mrg.endrow, mrg.endcol From xl_mergecells As mrg ;
          WHERE mrg.workbook = tnWB ;
          AND mrg.Sheet = xl_sheets.Sheet ;
          ORDER By mrg.begrow ;
          INTO Cursor t_merge

        If _Tally > 0
          \\<mergeCells>
          Select t_merge
          Scan
            \\<mergeCell ref="<<.ColumnIndexToAscii(t_merge.begcol) + Transform(t_merge.begrow)>>:<<.ColumnIndexToAscii(t_merge.endcol) + Transform(t_merge.endrow)>>"/>
          Endscan
          \\</mergeCells>
        Endif
        Use In t_merge

*-*	Write the cell validation data
        Select xl_validation
        Count For xl_validation.workbook = tnWB And xl_validation.Sheet = xl_sheets.Sheet And !Deleted() To lnValidCnt
        If lnValidCnt > 0
          \\<dataValidations count="<<Transform(lnValidCnt)>>">
          Scan
            Select xl.cellrow, xl.cellcol From xl_cells As xl ;
              WHERE BinToC(xl.workbook)+BinToC(xl.Sheet)+BinToC(xl.validndx) = BinToC(tnWB)+BinToC(xl_sheets.Sheet)+BinToC(xl_validation.validndx) ;
              ORDER By xl.cellrow, xl.cellcol ;
              INTO Cursor t_celllist
            Select t_celllist
            lcSqRef = ""
            Scan
              lcSqRef = lcSqRef + .ColumnIndexToAscii(t_celllist.cellcol) + Transform(t_celllist.cellrow) + " "
            Endscan
            lcSqRef = Alltrim(lcSqRef)
            Do Case
            Case xl_validation.vtype = NONE_VALID_TYPE
              lcValidationText = 'type="none"'

            Case xl_validation.vtype = WHOLE_VALID_TYPE
              lcValidationText = 'type="whole"'

            Case xl_validation.vtype = DECIMAL_VALID_TYPE
              lcValidationText = 'type="decimal"'

            Case xl_validation.vtype = LIST_VALID_TYPE
              lcValidationText = 'type="list"'

            Case xl_validation.vtype = DATE_VALID_TYPE
              lcValidationText = 'type="date"'

            Case xl_validation.vtype = TIME_VALID_TYPE
              lcValidationText = 'type="time"'

            Case xl_validation.vtype = TXTLEN_VALID_TYPE
              lcValidationText = 'type="textLength"'

            Case xl_validation.vtype = CUSTOM_VALID_TYPE
              lcValidationText = 'type="custom"'

            Otherwise
              lcValidationText = 'type="none"'
            Endcase
            If xl_validation.vstyle > 0
              Do Case
              Case xl_validation.vstyle = STOP_VALID_STYLE
                lcValidationText = lcValidationText + ' errorStyle="stop"'

              Case xl_validation.vstyle = WARN_VALID_STYLE
                lcValidationText = lcValidationText + ' errorStyle="warning"'

              Case xl_validation.vstyle = INFO_VALID_STYLE
                lcValidationText = lcValidationText + ' errorStyle="information"'
              Endcase
            Endif
            If xl_validation.voperator > 0
              Do Case
              Case xl_validation.voperator = BETWEEN_VALID_OPER
                lcValidationText = lcValidationText + ' operator="between"'

              Case xl_validation.voperator = NOTBETW_VALID_OPER
                lcValidationText = lcValidationText + ' operator="notBetween"'

              Case xl_validation.voperator = EQUAL_VALID_OPER
                lcValidationText = lcValidationText + ' operator="equal"'

              Case xl_validation.voperator = NOTEQUAL_VALID_OPER
                lcValidationText = lcValidationText + ' operator="notEqual"'

              Case xl_validation.voperator = LESSTHAN_VALID_OPER
                lcValidationText = lcValidationText + ' operator="lessThan"'

              Case xl_validation.voperator = LESSOREQUAL_VALID_OPER
                lcValidationText = lcValidationText + ' operator="lessThanOrEqual"'

              Case xl_validation.voperator = GREATTHAN_VALID_OPER
                lcValidationText = lcValidationText + ' operator="greaterThan"'

              Case xl_validation.voperator = GREATOREQUAL_VALID_OPER
                lcValidationText = lcValidationText + ' operator="greaterThanOrEqual"'
              Endcase
            Endif
            lcValidationText = lcValidationText + ' allowBlank="' + Iif(xl_validation.allowblank, '1', '0') + '"'
            lcValidationText = lcValidationText + ' showInputMessage="' + Iif(xl_validation.showinpmsg, '1', '0') + '"'
            lcValidationText = lcValidationText + ' showErrorMessage="' + Iif(xl_validation.showerrmsg, '1', '0') + '"'
            If !Empty(xl_validation.errtitle)
              lcValidationText = lcValidationText + ' errorTitle="' + Alltrim(xl_validation.errtitle) + '"'
            Endif
            If !Empty(xl_validation.errmsg)
              lcValidationText = lcValidationText + ' error="' + Alltrim(xl_validation.errmsg) + '"'
            Endif
            If !Empty(xl_validation.vprompt)
              lcValidationText = lcValidationText + ' prompt="' + Alltrim(xl_validation.vprompt) + '"'
            Endif
            \\<dataValidation <<lcValidationText>> sqref="<<lcSqRef>>">
            If xl_validation.formula
              \\<formula1><<Alltrim(xl_validation.formula1)>></formula1>
              If !Empty(xl_validation.formula2)
                \\<formula2><<Alltrim(xl_validation.formula2)>></formula2>
              Endif
            Else
              \\<formula1>"<<Alltrim(xl_validation.formula1)>>"</formula1>
              If !Empty(xl_validation.formula2)
                \\<formula2>"<<Alltrim(xl_validation.formula2)>"</formula2>
              Endif
            Endif
            \\</dataValidation>
          Endscan
          Use In Select('t_celllist')
          \\</dataValidations>
        Endif

*-*	Write sheet margin data
        \\<pageMargins left="<<Rtrim(Transform(xl_sheets.mleft), 1, '0')>>"
        \\right="<<Rtrim(Transform(xl_sheets.mright), 1, '0')>>"
        \\top="<<Rtrim(Transform(xl_sheets.mtop), 1, '0')>>"
        \\bottom="<<Rtrim(Transform(xl_sheets.mbot), 1, '0')>>"
        \\header="<<Rtrim(Transform(xl_sheets.mheader), 1, '0')>>"
        \\footer="<<Rtrim(Transform(xl_sheets.mfooter), 1, '0')>>"/>

*-*	Write the sheet page print setup
        \\<pageSetup orientation="<<Iif(xl_sheets.prnorient=PORTRAIT_PRINT_ORIENTATION, 'portrait', 'landscape')>>"
        If xl_sheets.Paperwidth > 0 And xl_sheets.paperheight > 0
          \\ paperHeight="<<Transform(xl_sheets.paperheight) + xl_sheets.paperdimen>>"
          \\ paperWidth="<<Transform(xl_sheets.Paperwidth) + xl_sheets.paperdimen>>"
        Else
          If xl_sheets.Papersize > 0
            \\ paperSize="' + Transform(xl_sheets.Papersize) + '"
          Endif
        Endif
        Do Case
        Case xl_sheets.fittowidth > 0 And xl_sheets.fittoheight > 0
          \\ fitToWidth="<<Transform(xl_sheets.fittowidth)>>"
          \\ fitToHeight="<<Transform(xl_sheets.fittoheight)>>"

        Case xl_sheets.fittowidth = 0 And xl_sheets.fittoheight > 0
          \\ fitToHeight="<<Transform(xl_sheets.fittoheight)>>"

        Case xl_sheets.fittowidth > 0 And xl_sheets.fittoheight = 0
          \\ fitToWidth="<<Transform(xl_sheets.fittowidth)>>"

        Otherwise
          If xl_sheets.Scale != 100
            \\ scale="<Transform(xl_sheets.Scale)>>"
          Endif
        Endcase
        \\/>

*-*	Write the header and footer for the sheet if defined
        If Seek(BinToC(tnWB)+BinToC(xl_sheets.Sheet), "xl_hdrfooterdefn", "wbsheet")
          \\<headerFooter
          If xl_hdrfooterdefn.alignmargin
            \\ alignWithMargins="1"
          Endif
          If xl_hdrfooterdefn.difffirstpg
            \\ differentFirst="1"
          Endif
          If xl_hdrfooterdefn.diffoddeven
            \\ differentOddEven="1"
          Endif
          If xl_hdrfooterdefn.scalewdoc
            \\ scaleWithDoc="1"
          Endif
          \\>

*-*		Get the page header text
          lcOddText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_HDR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddHeader>', lcOddText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_HDR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddHeader>', lcOddText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_HDR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddHeader>', lcOddText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcOddText)
            lcOddText = lcOddText + '</oddHeader>'
          Endif

          lcEvenText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_HDR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenHeader>', lcEvenText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_HDR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenHeader>', lcEvenText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_HDR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenHeader>', lcEvenText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcEvenText)
            lcEvenText = lcEvenText + '</evenHeader>'
          Endif

          lcSameText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_HDR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddHeader>', lcSameText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_HDR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddHeader>', lcSameText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_HDR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddHeader>', lcSameText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcSameText)
            lcSameText = lcSameText + '</oddHeader>'
          Endif
          If Empty(lcSameText)
            lcHeaderText = lcOddText + lcEvenText
          Else
            lcHeaderText = lcSameText
          Endif

*-*		Get the page footer text
          lcOddText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_FTR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddFooter>', lcOddText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_FTR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddFooter>', lcOddText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_ODD_PAGE, HEADERFOOTER_SECT_FTR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcOddText = Iif(Empty(lcOddText), '<oddFooter>', lcOddText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcOddText)
            lcOddText = lcOddText + '</oddFooter>'
          Endif

          lcEvenText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_FTR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenFooter>', lcEvenText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_FTR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenFooter>', lcEvenText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_EVEN_PAGE, HEADERFOOTER_SECT_FTR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcEvenText = Iif(Empty(lcEvenText), '<evenFooter>', lcEvenText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcEvenText)
            lcEvenText = lcEvenText + '</evenFooter>'
          Endif

          lcSameText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_FTR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddFooter>', lcSameText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_FTR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddFooter>', lcSameText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_SAME_PAGE, HEADERFOOTER_SECT_FTR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcSameText = Iif(Empty(lcSameText), '<oddFooter>', lcSameText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcSameText)
            lcSameText = lcSameText + '</oddFooter>'
          Endif
          If Empty(lcSameText)
            lcFooterText = lcOddText + lcEvenText
          Else
            lcFooterText = lcSameText
          Endif

*-*		Get the first page text
          lcFirstHdrText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_HDR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcFirstHdrText = Iif(Empty(lcFirstHdrText), '<firstHeader>', lcFirstHdrText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_HDR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcFirstHdrText = Iif(Empty(lcFirstHdrText), '<firstHeader>', lcFirstHdrText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_HDR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcFirstHdrText = Iif(Empty(lcFirstHdrText), '<firstHeader>', lcFirstHdrText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcFirstHdrText)
            lcFirstHdrText = lcFirstHdrText + '</firstHeader>'
          Endif

          lcFirstFtrText = ""
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_FTR_LEFT), "xl_hdrfootertext", "hdrfootr")
            lcFirstFtrText = Iif(Empty(lcFirstFtrText), '<firstFooter>', lcFirstFtrText) + '&amp;L&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_FTR_CENTER), "xl_hdrfootertext", "hdrfootr")
            lcFirstFtrText = Iif(Empty(lcFirstFtrText), '<firstFooter>', lcFirstFtrText) + '&amp;C&amp;' + .GetXmlHeaderFooterText()
          Endif
          If Seek(.GetHeaderFooterKeyExpr(tnWB, xl_sheets.Sheet, HEADERFOOTER_FIRST_PAGE, HEADERFOOTER_SECT_FTR_RIGHT), "xl_hdrfootertext", "hdrfootr")
            lcFirstFtrText = Iif(Empty(lcFirstFtrText), '<firstFooter>', lcFirstFtrText) + '&amp;R&amp;' + .GetXmlHeaderFooterText()
          Endif
          If !Empty(lcFirstFtrText)
            lcFirstFtrText = lcFirstFtrText + '</firstFooter>'
          Endif
*-*		Write the header/footer text to XML
          \\<<lcHeaderText + lcFooterText + lcFirstHdrText + lcFirstFtrText>></headerFooter>
        Endif
*-*	Completed writing sheet
        \\</worksheet>
        llSuccess = True
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

      Catch To loException
        Set Step On
        llSuccess = False
        .ErrorLevelId = 14
        Raiseevent(This, "OnShowErrorMessage", 14, loException.Message)

      Endtry
      If !Empty(lcAlias)
        Select (lcAlias)
      Endif
    Endwith
    Return llSuccess

*-- Writes the sharedStrings XML
  Protected Procedure WriteStringsXML
    Lparameters tnWB, tcPath
    Local lcFileName, lnTotCount, lnUniqueCnt, loException, llSuccess, lnRelId
    With This
      Raiseevent(This, "OnShowStatusMessage", 2, 2)
      Try
        Select xl_strformat
        Set Order To Tag sortorder
        Count For workbook = tnWB And !Deleted() To lnTotCount

        Select xl_strings
        Set Order To Tag sortorder
        Count For workbook = tnWB And !Deleted() To lnUniqueCnt
        lnTotCount = lnTotCount + lnUniqueCnt

        If lnUniqueCnt > 0
*-*		If the shared strings xml has not been defined in the relationships, add it
          If !Seek(BinToC(tnWB)+"sharedStrings", "xl_relationships", "reltype")
            lnRelId = .GetNextId(tnWB, "xl_relationships")
            Insert Into xl_relationships (workbook, relid, reltype, Target) Values (tnWB, lnRelId, "sharedStrings", "sharedStrings.xml")
          Endif

*-*		Output the file
          Set Textmerge On To Memvar lcOutStr Noshow
          lcFileName = tcPath + "xl\sharedStrings.xml"

          \\'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          \<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="<<Transform(lnTotCount)>>" uniqueCount="<<Transform(lnUniqueCnt)>>">   && Change 19Jun2018: missplaced >

          Select xl_strings
          Scan For workbook = tnWB And !Deleted()
            If xl_strings.Formatted
              \\<si>
              Select xl_strformat
              Scan For workbook = tnWB And Id = xl_strings.Id And !Deleted()
                \\<r>
                If xl_strformat.fbold Or xl_strformat.fitalic Or xl_strformat.fcolor>0 Or !Empty(xl_strformat.fname) Or xl_strformat.Fsize>0 ;
                    OR xl_strformat.strkthr Or !Empty(xl_strformat.fvpos) Or xl_strformat.Theme>0 Or xl_strformat.indexed>0 Or xl_strformat.uline != UNDERLINE_NONE

                  \\<rPr>
                  If xl_strformat.fbold
                    \\<b/>
                  Endif
                  If xl_strformat.fitalic
                    \\<i/>
                  Endif
                  If xl_strformat.uline != UNDERLINE_NONE
                    \\<u/>
                  Endif
                  If xl_strformat.strkthr
                    \\<strike/>
                  Endif
                  If xl_strformat.Fsize > 0
                    \\<sz val="<<Transform(xl_strformat.Fsize)>>"/>
                  Endif
                  Do Case
                  Case xl_strformat.fcolor > 0
                    \\<color rgb="<<.ConvertColorToHex(xl_strformat.fcolor)>>"/>

                  Case xl_strformat.indexed > 0
                    \\<color indexed="<<Transform(xl_strformat.indexed)>>"/>

                  Case xl_strformat.Theme > 0
                    \\<color theme="<<Transform(xl_strformat.Theme)>>" tint="<<Transform(xl_strformat.tint, "99.999999999999999")>>"/>
                  Endcase
                  If !Empty(xl_strformat.fvpos)
                    \\<vertAlign val="<<Alltrim(xl_strformat.fvpos)>>"/>
                  Endif
                  If !Empty(xl_strformat.fname)
                    \\<rFont val="<<Alltrim(xl_strformat.fname)>>"/>
                  Endif
                  \\</rPr>
                Endif
*					IF xl_strformat.presvspace
                \\<t xml:space="preserve"><<Chrtran(xl_strformat.stringxml, Chr(160), Chr(32))>></t>
*					ELSE
*						FWRITE(lhFile, '<t>' + xl_strformat.stringxml + '</t>')
*					ENDIF
                \\</r>
              Endscan
              \\</si>
              Select xl_strings
            Else
              If xl_strings.presvspace
                \\<si><t xml:space="preserve"><<Chrtran(xl_strformat.stringxml, Chr(160), Chr(32))>></t></si>
              Else
                \\<si><t><<xl_strings.stringxml>></t></si>
              Endif
            Endif
          Endscan
          Set Order To
          \\</sst>
          Set Textmerge To
          .SaveAsUtf8(lcFileName, lcOutStr)
        Else
*-*		No strings are assigned to cells; remove from relationships if present
          Delete From xl_relationships Where workbook = tnWB And reltype = "sharedStrings"
        Endif
        llSuccess = True

      Catch To loException
        Set Step On
        llSuccess = False
        .ErrorLevelId = 15
        Raiseevent(This, "OnShowErrorMessage", 15, loException.Message)
      Endtry
    Endwith
    Return llSuccess

*-- Writes the styles.xml support file
  Protected Procedure WriteStylesXML
    Lparameters tnWB, tcPath
    Local lcAlias, lcFileName, lnNumFmtsId, lnRecCount, lcNumFmtId, lnRotation, llIndexedColors
    Local llMruColors
* TODO: заменить на textmerge
    With This
      Raiseevent(This, "OnShowStatusMessage", 2, 3)
      Try
        lcAlias = Alias()
        lcFileName = tcPath + "xl\styles.xml"
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        \\xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        \\mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">

*-*	Get last used custom number format Id
        lnNumFmtsId = 0
        Select xl_numFmts
        Set Order To Tag sortorder
        Scan For workbook = tnWB And xl_numFmts.Id < 900
          If xl_numFmts.Id > lnNumFmtsId
            lnNumFmtsId = xl_numFmts.Id
          Endif
        Endscan

*-*	Write the Number format section
        Count For workbook = tnWB To lnRecCount
        \\<numFmts count="<<Transform(lnRecCount)>>">
        Scan For workbook = tnWB
          If xl_numFmts.Id > 900
            lnNumFmtsId = lnNumFmtsId + 1
            \\<numFmt numFmtId="<<Transform(lnNumFmtsId)>>" formatCode="<<Alltrim(xl_numFmts.formatxml)>"/>
            Replace xl_numFmts.tempid With lnNumFmtsId In xl_numFmts
          Else
            \\<numFmt numFmtId="<<Transform(xl_numFmts.Id)>>" formatCode="<<Alltrim(xl_numFmts.formatxml)>"/>
            Replace xl_numFmts.tempid With xl_numFmts.Id In xl_numFmts
          Endif
        Endscan
        \\</numFmts>
        Set Order To

*-*	Write the fonts section
        Select xl_fonts
        Set Order To Tag sortorder
        Count For workbook = tnWB To lnRecCount
        \\<fonts count="<<Transform(lnRecCount)>" x14ac:knownFonts="1">
        Scan For workbook = tnWB
          \\<font>
          \\<name val="<<Alltrim(xl_fonts.fname)>>"/>
          \\<sz val="<<Transform(xl_fonts.Fsize)>>"/>
          If xl_fonts.fbold
            \\<b/>
          Endif
          If xl_fonts.fitalic
            \\<i/>
          Endif
          If !Empty(xl_fonts.uline)
            \\<u val="<<Alltrim(xl_fonts.uline)>>"/>
          Endif
          If xl_fonts.strkthr
            \\<strike/>
          Endif
          If !Empty(xl_fonts.fvpos)
            \\<vertAlign val="<<Alltrim(xl_fonts.fvpos)>>"/>
          Endif
          Do Case
          Case xl_fonts.fcolor > 0
            \\<color rgb="<<.ConvertColorToHex(xl_fonts.fcolor)>>"/>

          Case xl_fonts.indexed > 0
            \\<color indexed="<<Transform(xl_fonts.indexed)>>"/>

          Case xl_fonts.Theme > 0
            \\<color theme="<<Transform(xl_fonts.Theme)>>" tint="<<Transform(xl_fonts.tint, "99.999999999999999")>>"/>
          Endcase
          \\</font>
        Endscan
        \\</fonts>
        Set Order To

*-*	Write the fills section
        Select xl_fills
        Set Order To Tag sortorder
        Count For workbook = tnWB To lnRecCount
        \\<fills count="<<Transform(lnRecCount)>>">
        Scan For workbook = tnWB
          \\<fill><patternFill patternType="<<Alltrim(xl_fills.patttype)>>"
          Do Case
          Case xl_fills.fgindexed > 0 And xl_fills.bgindexed > 0
            \\>
            \\<fgColor indexed="<<Transform(xl_fills.fgindexed)>>"/>
            \\<bgColor indexed="<<Transform(xl_fills.bgindexed)>"/>
            \\</patternFill>

          Case xl_fills.fgcolor > 0 And xl_fills.bgindexed > 0
            \\>
            \\<fgColor rgb="<<.ConvertColorToHex(xl_fills.fgcolor)>>"/>
            \\<bgColor indexed="<<Transform(xl_fills.bgindexed)>>"/>
            \\</patternFill>

          Case xl_fills.bgcolor > 0 And xl_fills.fgindexed > 0
            \\>
            \\<fgColor indexed="<<Transform(xl_fills.fgindexed)>>"/>
            \\<bgColor rgb="<<.ConvertColorToHex(xl_fills.bgcolor)>>"/>
            \\</patternFill>')

          Case xl_fills.fgcolor > 0 And xl_fills.bgcolor > 0
            \\>
            \\<fgColor rgb="<<.ConvertColorToHex(xl_fills.fgcolor)>>"/>
            \\<bgColor rgb="<<.ConvertColorToHex(xl_fills.bgcolor)>>"/>
            \\</patternFill>

          Case xl_fills.Theme > 0
            \\>
            \\<fgColor theme="<<Transform(xl_fills.Theme)>>" tint="<<Transform(xl_fills.tint, "99.999999999999999")>>"/>
            \\<bgColor indexed="64"/>')
            \\</patternFill>')

          Otherwise
            \\/>
          Endcase
          \\</fill>
        Endscan
        \\</fills>

*-*	Write the borders section
        Select xl_borders
        Set Order To Tag sortorder
        Count For workbook = tnWB To lnRecCount
        \\<borders count="<<Transform(lnRecCount)>>">
        Scan For workbook = tnWB
          \\<border<<Iif(xl_borders.diagdn=1, ' diagonalDown="1"', '') + Iif(xl_borders.diagup=1, ' diagonalUp="1"', '')>>>
          If Empty(xl_borders.lstyle)
            \\<left/>
          Else
            If xl_borders.lstyle = BORDER_STYLE_NONE
              \\<left
            Else
              \\<left style="<<Alltrim(xl_borders.lstyle)>>"
            Endif
            Do Case
            Case xl_borders.lcolor > 0
              \\>
              \\<color rgb="<<.ConvertColorToHex(xl_borders.lcolor)>"/>
              \\</left>

            Case xl_borders.lindexed > 0
              \\>
              \\<color indexed="<<Transform(xl_borders.lindexed)>>"/>
              \\</left>

            Case xl_borders.ltheme > 0
              \\>
              \\<color theme="<<Transform(xl_borders.ltheme)>>" tint="<<Transform(xl_borders.ltint, "99.999999999999999")>>"/>
              \\</left>

            Otherwise
              \\/>
            Endcase
          Endif
          If Empty(xl_borders.rstyle)
            \\<right/>
          Else
            If xl_borders.rstyle = BORDER_STYLE_NONE
              \\<right
            Else
              \\<right style="<<Alltrim(xl_borders.rstyle)>>"
            Endif
            Do Case
            Case xl_borders.rcolor > 0
              \\>
              \\<color rgb="<<.ConvertColorToHex(xl_borders.rcolor)>>"/>
              \\</right>

            Case xl_borders.rindexed > 0
              \\>
              \\<color indexed="<<Transform(xl_borders.rindexed)>>"/>
              \\</right>


            Case xl_borders.rtheme > 0
              \\>
              \\<color theme="<<Transform(xl_borders.rtheme)>>" tint="<<Transform(xl_borders.rtint, "99.999999999999999")>>"/>
              \\</right>

            Otherwise
              \\/>
            Endcase
          Endif
          If Empty(xl_borders.tstyle)
            \\<top/>
          Else
            If xl_borders.tstyle = BORDER_STYLE_NONE
              \\<top
            Else
              \<top style="<<Alltrim(xl_borders.tstyle)>>"
            Endif
            Do Case
            Case xl_borders.tcolor > 0
              \\>
              \\<color rgb="<<.ConvertColorToHex(xl_borders.tcolor)>>"/>
              \\</top>')

            Case xl_borders.tindexed > 0
              \\>
              \\<color indexed="<<Transform(xl_borders.tindexed)>"/>
              \\</top>

            Case xl_borders.ttheme > 0
              \\>
              \\<color theme="<<Transform(xl_borders.ttheme)>>" tint="<<Transform(xl_borders.ttint, "99.999999999999999")>>"/>
              \\</top>')

            Otherwise
              \\/>
            Endcase
          Endif
          If Empty(xl_borders.bstyle)
            \\<bottom/>
          Else
            If xl_borders.bstyle = BORDER_STYLE_NONE
              \\<bottom
            Else
              \\<bottom style="<<Alltrim(xl_borders.bstyle)>>"
            Endif
            Do Case
            Case xl_borders.bcolor > 0
              \\>
              \\<color rgb="<<.ConvertColorToHex(xl_borders.bcolor)>>"/>
              \\</bottom>

            Case xl_borders.bindexed > 0
              \\>
              \\<color indexed="<<Transform(xl_borders.bindexed)>>"/>
              \\</bottom>

            Case xl_borders.btheme > 0
              \\>
              \\<color theme="<<Transform(xl_borders.btheme)>>" tint="<<Transform(xl_borders.btint, "99.999999999999999")>>"/>
              \\</bottom>

            Otherwise
              \\/>
            Endcase
          Endif
          If Empty(xl_borders.dstyle)
            \\<diagonal/>
          Else
            If xl_borders.dstyle = BORDER_STYLE_NONE
              \\<diagonal
            Else
              \\<diagonal style="<<Alltrim(xl_borders.dstyle)>>"
            Endif
            Do Case
            Case xl_borders.dcolor > 0
              \\>
              \\<color rgb="<<.ConvertColorToHex(xl_borders.dcolor)>>"/>
              \\</diagonal>

            Case xl_borders.dindexed > 0
              \\>
              \\<color indexed="<<Transform(xl_borders.dindexed)>"/>
              \\</diagonal>

            Case xl_borders.dtheme > 0
              \\>
              \\<color theme="<<Transform(xl_borders.dtheme)>>" tint="<<Transform(xl_borders.dtint, "99.999999999999999")>>"/>
              \\</diagonal>

            Otherwise
              \\/>
            Endcase
          Endif
          \\</border>
        Endscan
        \\</borders>
*-*	Write the cellStyleXfs section

*-*	Write the cellXfs section
        Select xl_cellxfs
        Set Order To Tag sortorder
        Count For workbook = tnWB To lnRecCount
        Fwrite(lhFile, '<cellXfs count="' + Transform(lnRecCount) + '">')
        Scan For workbook = tnWB
          If xl_cellxfs.numfmtid > 0 And Seek(BinToC(tnWB)+BinToC(xl_cellxfs.numfmtid), "xl_numfmts", "id")
            lcNumFmtId = Transform(xl_numFmts.tempid)
          Else
            lcNumFmtId = Transform(xl_cellxfs.numfmtid)
          Endif
          \\<xf numFmtId="<<lcNumFmtId>>" fontId="<<Transform(xl_cellxfs.fontid)>>"
          \\ fillId="<<Transform(xl_cellxfs.fillid)>>" borderId="<<Transform(xl_cellxfs.borderid)>>"
          \\<<IIF(lcNumFmtId = '0', '', ' applyNumberFormat="1"') + Iif(xl_cellxfs.fontid > 0, ' applyFont="1"', '') + IIF(xl_cellxfs.fillid > 0, ' applyFill="1"', '') + Iif(xl_cellxfs.borderid > 0, ' applyBorder="1"', '')>> applyAlignment="1"

          If !Empty(xl_cellxfs.halign) Or !Empty(xl_cellxfs.valign) Or xl_cellxfs.wraptext > 0 Or xl_cellxfs.indent > 0 Or xl_cellxfs.Rotation > 0
            \\><alignment
            If !Empty(xl_cellxfs.halign)
              \\ horizontal="<<Alltrim(xl_cellxfs.halign)>>"
            Endif
            If !Empty(xl_cellxfs.valign)
              \\ vertical="<<Alltrim(xl_cellxfs.valign)>>"
            Endif
            If xl_cellxfs.wraptext > 0
              \\ wrapText="1"
            Endif
            If xl_cellxfs.indent > 0
              \\ indent="<<Transform(xl_cellxfs.indent)>>"
            Endif
            If xl_cellxfs.Rotation != 0
              \\ textRotation="<<Transform(xl_cellxfs.Rotation)>>"
            Endif
            \\/></xf>
          Else
            \\/>
          Endif
        Endscan
        \\</cellXfs>

*-*	Write the colors section
        llIndexedColors = Seek(BinToC(tnWB)+BinToC(1), "xl_ndxcolors", "indexid")
        llMruColors     = Seek(BinToC(tnWB)+BinToC(1), "xl_mrucolors", "indexid")
        If llIndexedColors Or llMruColors
          \\<colors>
          If llIndexedColors
            \\<indexedColors>
            Select xl_ndxcolors
            Set Order To Tag indexid
            Scan For workbook = tnWB
              \\<rgbColor rgb="<<xl_ndxcolors.rgbcolor>>"/>
            Endscan
            \\</indexedColors>
          Endif

          If llMruColors
            \\<mruColors>
            Select xl_mrucolors
            Set Order To Tag indexid
            Scan For workbook = tnWB
              \\<color rgb="<<xl_ndxcolors.rgbcolor>>"/>
            Endscan
            \\</mruColors>
          Endif
          \\</colors>
        Endif

*-*	Close the style sheet
        \\</styleSheet>
        llSuccess = True

      Catch To loException
        Set Step On
        llSuccess = False
        .ErrorLevelId = 17
        Raiseevent(This, "OnShowErrorMessage", 17, "Unable to create styles.xml")

      Finally
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)
      Endtry
      If !Empty(lcAlias)
        Select (lcAlias)
      Endif
    Endwith
    Return llSuccess

*-- Writes the [Content_Types].xml, .rels, core.xml, workbook.xml.rels and app.xml files
  Protected Procedure WriteSupportXMLs
    Lparameters tnWB, tcPath
    Local llSuccess, loException, lhFile, lnSheet, lnUniqueCnt, lcCurr, lcColor, lnTotCount, lcFileName, lcBaseContentType
    With This
      Raiseevent(This, "OnShowStatusMessage", 2, 1)
      Try
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
        \\<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
        \\<Default Extension="xml" ContentType="application/xml"/>
        \\<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>

        lcBaseContentType = '<Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.'
        Select xl_relationships
        Scan For workbook = tnWB And !Deleted()
          Do Case
          Case xl_relationships.reltype = "calcChain"
*				FWRITE(lhFile, lcBaseContentType + 'calcChain+xml" PartName="/xl/calcChain.xml"/>')

          Case xl_relationships.reltype = "drawings"
*				FWRITE(lhFile, lcBaseContentType + 'drawing+xml" PartName="/xl/drawings/drawing1.xml"/>')

          Case xl_relationships.reltype = "externalLink"
*				FWRITE(lhFile, lcBaseContentType + 'externalLink+xml" PartName="/xl/externalLinks/externalLink1.xml"/>')

          Case xl_relationships.reltype = "sharedStrings"
            Fwrite(lhFile, lcBaseContentType + 'sharedStrings+xml" PartName="/xl/sharedStrings.xml"/>')

          Case xl_relationships.reltype = "styles"
            Fwrite(lhFile, lcBaseContentType + 'styles+xml" PartName="/xl/styles.xml"/>')

          Case xl_relationships.reltype = "theme"
*				FWRITE(lhFile, lcBaseContentType + 'theme+xml" PartName="/xl/theme/theme1.xml"/>')

          Case xl_relationships.reltype = "worksheet"
            Fwrite(lhFile, lcBaseContentType + 'worksheet+xml" PartName="/xl/worksheets/sheet' + Transform(xl_relationships.Sheet) + '.xml"/>')
          Endcase
        Endscan

        \\<Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>
        \\<Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>
        \\</Types>
        Set Textmerge To
        .SaveAsUtf8(tcPath + "[Content_Types].xml", lcOutStr)

*-*	Save the .rels XML
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        \\<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
        \\<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
        \\<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
        \\</Relationships>
        Set Textmerge To
        .SaveAsUtf8(tcPath + "_rels\.rels", lcOutStr)

*-*	Save the app XML
        Set Textmerge On To Memvar lcOutStr Noshow
        lcFileName = tcPath + "docProps\app.xml"
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
        \\<Application>ExcelXMLX Release ' + Transform(.ExcelXlsxRelease) + '</Application>
        If !Empty(.CompanyName)
          \\<Company>' + .CompanyName + '</Company>
        Endif
        \\</Properties>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

*-*	Save the core XML
        lcFileName = tcPath + "docProps\core.xml"
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \\<cp:coreProperties
        \\xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
        \\xmlns:dc="http://purl.org/dc/elements/1.1/"
        \\xmlns:dcterms="http://purl.org/dc/terms/"
        \\xmlns:dcmitype="http://purl.org/dc/dcmitype/"
        \\xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
        If !Empty(.Title)
          \\<dc:title>' + .GetXMLString(.Title) + '</dc:title>
        Endif
        If !Empty(.Subject)
          \\<dc:subject>' + .GetXMLString(.Subject) + '</dc:subject>
        Endif
        If !Empty(.CreatorName)
          \\<dc:creator>' + .GetXMLString(.CreatorName) + '</dc:creator>
        Endif
        If !Empty(.UserName)
          \\<cp:lastModifiedBy>' + .GetXMLString(.UserName) + '</cp:lastModifiedBy>
        Endif
        \\<dcterms:created xsi:type="dcterms:W3CDTF">' + Ttoc(Datetime(), 3) + 'Z</dcterms:created>
        \\<dcterms:modified xsi:type="dcterms:W3CDTF">' + Ttoc(Datetime(), 3) + 'Z</dcterms:modified>
        \\</cp:coreProperties>
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)

        llSuccess = True

      Catch To loException
        Set Step On
        llSuccess = False
        .ErrorLevelId = 18
        Raiseevent(This, "OnShowErrorMessage", 18, "Unable to create workbook supporting XMLs")
      Endtry
    Endwith
    Return llSuccess

*-- Writes the workbook XML
  Protected Procedure WriteWorkbookXML
    Lparameters tnWB, tcPath
    Local lcFileName, lhFile, loException, llSuccess, lcSheetName
    With This
      Raiseevent(This, "OnShowStatusMessage", 2, 4)
      Try
        lcFileName = tcPath + "xl\workbook.xml"
        Set Textmerge On To Memvar lcOutStr Noshow
        \\<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        \<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">

*-*	Write thw sheet information
        \\<sheets>
        Select xl_relationships
        Set Order To Tag relid
        Scan For workbook = tnWB And reltype = "worksheet"
          If Seek(BinToC(tnWB)+BinToC(xl_relationships.Sheet), "xl_sheets", "sheetndx")
            Do Case
            Case xl_sheets.state = VISIBLE_SHEET_STATE
              \\<sheet r:id="rId<<Transform(xl_relationships.relid)>>" sheetId="<<Transform(xl_relationships.Sheet)>>" name="<<Alltrim(xl_sheets.shname)>>"/>

            Case xl_sheets.state = HIDDEN_SHEET_STATE
              \\<sheet r:id="rId<<Transform(xl_relationships.relid)>>" sheetId="<<Transform(xl_relationships.Sheet)>>" state="hidden" name="<<Alltrim(xl_sheets.shname)>>"/>

            Case xl_sheets.state = VERYHIDDEN_SHEET_STATE
              \\<sheet r:id="rId<<Transform(xl_relationships.relid)>>" sheetId="<<Transform(xl_relationships.Sheet)>>" state="veryHidden" name="<<Alltrim(xl_sheets.shname)>>"/>
            Endcase
          Else
            Throw
          Endif
        Endscan
        Set Order To
        \\</sheets>

*-*	Write the named range information
        Select xl_namerange
        Locate For workbook = tnWB And !Deleted()
        If Found()
          \\<definedNames>
          Scan For workbook = tnWB And !Deleted()
            If (xl_namerange.begcol = 0 And xl_namerange.begrow = 0) Or (xl_namerange.begcol = 0 And xl_namerange.begrow = 0)
              Loop
            Endif
            \\<definedName name="<<Alltrim(xl_namerange.rname)>>"
            If !Empty(xl_namerange.Comment)
              \\ comment="<<Alltrim(xl_namerange.Comment)>>"
            Endif
            If xl_namerange.scope = SCOPE_SH_NAMED_RANGE
              \\ localSheetId="<<Transform(xl_namerange.Sheet)>>"
            Endif
            \\>

            lcSheetName = .GetSheetName(tnWB, xl_namerange.Sheet)
            Fwrite(lhFile, Iif(Occurs(" ", lcSheetName)>0, "'", "") + lcSheetName + Iif(Occurs(" ", lcSheetName)>0, "'", ""))
            Do Case
            Case xl_namerange.begcol = 0 And xl_namerange.begrow > 0
              \\!$<<Transform(xl_namerange.begrow)>>

            Case xl_namerange.begcol > 0 And xl_namerange.begrow = 0
              \\!$<<.ColumnIndexToAscii(xl_namerange.begcol)>>

            Otherwise
              \\!$<<.ColumnIndexToAscii(xl_namerange.begcol)>>$<<Transform(xl_namerange.begrow)>>
            Endcase
            Do Case
            Case xl_namerange.endcol = 0 And xl_namerange.endrow > 0
              \\:$<<Transform(xl_namerange.endrow)>>

            Case xl_namerange.endcol > 0 And xl_namerange.begrow = 0
              \\:$<<.ColumnIndexToAscii(xl_namerange.endcol)>>

            Otherwise
              \\:$<<.ColumnIndexToAscii(xl_namerange.endcol)>>$<<Transform(xl_namerange.endrow)>>
            Endcase
            \\</definedName>
          Endscan
          \\</definedNames>
        Endif

*-*	Close the workbook
        \\<calcPr calcId="0"/>
        \\</workbook>

        llSuccess = True

      Catch To loException
        Set Step On
        llSuccess = False

      Finally
        Set Textmerge To
        .SaveAsUtf8(lcFileName, lcOutStr)
      Endtry
    Endwith
    Return llSuccess

  Protected Procedure Destroy
    Set Debugout To
    With This
      .DestroyWorkingCursors()
      .OnDestroy()
    Endwith

  Protected Procedure Init
    Lparameters tnCodePage
    Local loException, lnSelect
    With This
      If Pcount() > 0 And Vartype(tnCodePage) = "N"
        .Codepage = tnCodePage
      Endif
      lnSelect = Select()
      .CreateWorkingCursors()
      If .DeclareWinApi
        .DeclareDLL()
      Endif
      Try
        .oXDom = Createobject('MSXML2.DOMDocument')
      Catch To loException
        .oXDom = Null
      Endtry
      .InitalizeIdValues()
      .OnInit()
      Select (lnSelect)
    Endwith

Enddefine
*
*-- EndDefine: vfpxworkbookxlsx
**************************************************
