Attribute VB_Name = "C_PublicFunctions"
' ************************************************************************************************
' © 2011 Dun & Bradstreet
'
' This program is free software; you can redistribute it and/or modify it under the terms of
' the GNU General Public License as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
' without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
' See the GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along with this program;
' if not, write to the Free Software Foundation, Inc., 51 Franklin St, Boston, MA 20110-1301, USA.
' ************************************************************************************************

Option Explicit


Public Function sWbCountry2IsoCode(sWbCtryCode As String) As String
    On Error GoTo ErrHandler
    
    sWbCountry2IsoCode = Application.WorksheetFunction.VLookup(sWbCtryCode, ThisWorkbook.Names("Ctry_Codes").RefersToRange, 3, False)
    Exit Function
    
ErrHandler:
    sWbCountry2IsoCode = ""
End Function


Public Function sWbCurr2IsoCode(sWbCurrCode As String) As String
    On Error GoTo ErrHandler
    
    sWbCurr2IsoCode = Application.WorksheetFunction.VLookup(sWbCurrCode, ThisWorkbook.Names("Curr_Codes").RefersToRange, 3, False)
    Exit Function
    
ErrHandler:
    sWbCurr2IsoCode = ""
End Function


Public Function sWbContDesc(sWbContCode As String) As String
    On Error GoTo ErrHandler
    
    If sWbContCode = "0" Or sWbContCode = "" Then GoTo ErrHandler
    
    sWbContDesc = Application.WorksheetFunction.VLookup(sWbContCode, ThisWorkbook.Names("Cont_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbContDesc = ""
End Function


Public Function sWbNatIDDesc(sWbNatIDCode As String) As String
    On Error GoTo ErrHandler
    
    sWbNatIDDesc = StrConv(Application.WorksheetFunction.VLookup(sWbNatIDCode, ThisWorkbook.Names("NatID_Codes").RefersToRange, 2, False), vbProperCase)
    Exit Function
    
ErrHandler:
    sWbNatIDDesc = ""
End Function


Public Function sWbCurrCodeDesc(sWbCurrCode As String) As String
    On Error GoTo ErrHandler
    
    If sCurrCode = "0000" Or sCurrCode = "" Then GoTo ErrHandler
    
    sWbCurrCodeDesc = Application.WorksheetFunction.VLookup(sCurrCode, ThisWorkbook.Names("Curr_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbCurrCodeDesc = ""
End Function


Public Function sWbAnnSalesIndDesc(sWbAnnSales As String, sWbAnnSalesInd As String) As String
    On Error GoTo ErrHandler
    
    If Right(sWbAnnSales, 15) = "000000000000000" Or sWbAnnSales = "" Then GoTo ErrHandler
    Select Case sWbAnnSalesInd
    Case "0"
        sWbAnnSalesIndDesc = "Actual"
    Case "1"
        sWbAnnSalesIndDesc = "Low end of range"
    Case "2"
        sWbAnnSalesIndDesc = "Estimated"
    Case "3"
        sWbAnnSalesIndDesc = "Modelled"
    Case Else
        sWbAnnSalesIndDesc = ""
    End Select
    Exit Function

ErrHandler:
    sWbAnnSalesIndDesc = ""
End Function


Public Function sWbImpExpCodeDesc(sWbImpExpCode As String) As String
    On Error GoTo ErrHandler
    
    If sWbImpExpCode = "G" Or sWbImpExpCode = "" Then GoTo ErrHandler
    
    sWbImpExpCodeDesc = Application.WorksheetFunction.VLookup(sWbImpExpCode, ThisWorkbook.Names("ImportExport_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbImpExpCodeDesc = ""
End Function


Public Function sWbLegalStatusCodeDesc(sWbLegalStatusCode As String) As String
    On Error GoTo ErrHandler
    
    If sWbLegalStatusCode = "000" Or sWbLegalStatusCode = "" Then GoTo ErrHandler
    
    sWbLegalStatusCodeDesc = Application.WorksheetFunction.VLookup(sWbLegalStatusCode, ThisWorkbook.Names("LegalStatus_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbLegalStatusCodeDesc = ""
End Function


Public Function sWbStatusCodeDesc(sWbStatusCode As String) As String
    On Error GoTo ErrHandler
    
    If sWbStatusCode = "" Then GoTo ErrHandler
    
    sWbStatusCodeDesc = Application.WorksheetFunction.VLookup(sWbStatusCode, ThisWorkbook.Names("Status_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbStatusCodeDesc = ""
End Function


Public Function sWbSubsidiaryCodeDesc(sWbSubsidiaryCode As String) As String
    On Error GoTo ErrHandler
    
    If sWbSubsidiaryCode = "" Then GoTo ErrHandler
    
    sWbSubsidiaryCodeDesc = Application.WorksheetFunction.VLookup(sWbSubsidiaryCode, ThisWorkbook.Names("Subsidiary_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbSubsidiaryCodeDesc = ""
End Function


Public Function sWbSIC87Desc(sWbSIC_4Code As String) As String
    On Error GoTo ErrHandler
    
    If sWbSIC_4Code = "9999" Or sWbSIC_4Code = "" Then GoTo ErrHandler
    
    sWbSIC87Desc = Application.WorksheetFunction.VLookup(sWbSIC_4Code, ThisWorkbook.Names("SIC87_Codes").RefersToRange, 2, False)
    Exit Function
    
ErrHandler:
    sWbSIC87Desc = ""
End Function


Public Function sWbSIC_2Code(sWbSIC_4Code As String) As String
    On Error GoTo ErrHandler
    
    If sWbSIC_4Code = "9999" Or sWbSIC_4Code = "" Then GoTo ErrHandler
    
    sWbSIC_2Code = Application.WorksheetFunction.VLookup(sWbSIC_4Code, ThisWorkbook.Names("SIC87_Codes").RefersToRange, 3, False)
    Exit Function
    
ErrHandler:
    sWbSIC_2Code = ""
End Function


Public Function sWbSIC87_2Desc(sWbSIC_4Code As String) As String
    On Error GoTo ErrHandler
    
    If sWbSIC_4Code = "9999" Or sWbSIC_4Code = "" Then GoTo ErrHandler
    
    sWbSIC87_2Desc = Application.WorksheetFunction.VLookup(sWbSIC_4Code, ThisWorkbook.Names("SIC87_Codes").RefersToRange, 4, False)
    Exit Function
    
ErrHandler:
    sWbSIC87_2Desc = ""
End Function


Public Function sWbSIC_1Code(sWbSIC_4Code As String) As String
    On Error GoTo ErrHandler
    
    If sWbSIC_4Code = "9999" Or sWbSIC_4Code = "" Then GoTo ErrHandler
    
    sWbSIC_1Code = Application.WorksheetFunction.VLookup(sWbSIC_4Code, ThisWorkbook.Names("SIC87_Codes").RefersToRange, 5, False)
    Exit Function
    
ErrHandler:
    sWbSIC_1Code = ""
End Function


Public Function sWbSIC87_1Desc(sWbSIC_4Code As String) As String
    On Error GoTo ErrHandler
    
    If sWbSIC_4Code = "9999" Or sWbSIC_4Code = "" Then GoTo ErrHandler
    
    sWbSIC87_1Desc = Application.WorksheetFunction.VLookup(sWbSIC_4Code, ThisWorkbook.Names("SIC87_Codes").RefersToRange, 6, False)
    Exit Function
    
ErrHandler:
    sWbSIC87_1Desc = ""
End Function


Public Function sWbLocalActCodeDesc(sWbLocalActCode As String) As String
    On Error GoTo ErrHandler

    If sWbLocalActCode = "000" Or sWbLocalActCode = "" Then GoTo ErrHandler

    sWbLocalActCodeDesc = Application.WorksheetFunction.VLookup(sWbLocalActCode, ThisWorkbook.Names("LocalAct_Codes").RefersToRange, 2, False)
    Exit Function

ErrHandler:
    sWbLocalActCodeDesc = ""
End Function


Public Function sWbNumEmplIndDesc(sWbNumEmpl As String, sWbNumEmplInd As String) As String
    On Error GoTo ErrHandler
    
    If sWbNumEmpl = "0000000" Or sWbNumEmpl = "" Then GoTo ErrHandler
    Select Case sWbNumEmplInd
    Case "0"
        sWbNumEmplIndDesc = "Actual"
    Case "1"
        sWbNumEmplIndDesc = "Low end of range"
    Case "2"
        sWbNumEmplIndDesc = "Estimated"
    Case "3"
        sWbNumEmplIndDesc = "Modelled"
    Case Else
        sWbNumEmplIndDesc = ""
    End Select
    Exit Function

ErrHandler:
    sWbNumEmplIndDesc = ""
End Function


Public Function iWbStartYear(sWbStartYear As String) As Variant
    On Error GoTo ErrHandler
    
    If sWbStartYear = "0000" Or sWbStartYear = "" Then GoTo ErrHandler
    iWbStartYear = CInt(sWbStartYear)
    Exit Function
    
ErrHandler:
    iWbStartYear = ""
End Function


Public Function lWbAnnSales(sWbAnnSales As String) As Variant
    On Error GoTo ErrHandler
    
    If Right(sWbAnnSales, 15) = "000000000000000" Or sWbAnnSales = "" Then GoTo ErrHandler
    lWbAnnSales = CLng(sWbAnnSales)
    Exit Function
    
ErrHandler:
    lWbAnnSales = ""
End Function


Public Function lWbNumEmpl(sWbNumEmpl As String) As Variant
    On Error GoTo ErrHandler
    
    If sWbNumEmpl = "0000000" Or sWbNumEmpl = "" Then GoTo ErrHandler
    lWbNumEmpl = CLng(sWbNumEmpl)
    Exit Function
    
ErrHandler:
    lWbNumEmpl = ""
End Function


Public Function lWbNumFamMembers(sWbNumFamMembers As String) As Variant
    On Error GoTo ErrHandler
    
    If sWbNumFamMembers = "00000" Or sWbNumFamMembers = "" Then GoTo ErrHandler
    lWbNumFamMembers = CLng(sWbNumFamMembers)
    Exit Function
    
ErrHandler:
    lWbNumFamMembers = ""
End Function


Public Function sWbStartYearDistr(sStartYear As Variant) As String
    On Error GoTo ErrHandler
    
    Dim lStartYear As Long
    
    If sStartYear = "" Then GoTo ErrHandler
    lStartYear = CLng(sStartYear)
    
    Select Case lStartYear
    Case 0 To 1899
        sWbStartYearDistr = "1. < 1900"
    Case 1900 To 1949
        sWbStartYearDistr = "2. >= 1900 and < 1950"
    Case 1950 To 1974
        sWbStartYearDistr = "3. >= 1950 and < 1975"
    Case 1975 To 1989
        sWbStartYearDistr = "4. >= 1975 and < 1990"
    Case 1990 To 1999
        sWbStartYearDistr = "5. >= 1990 and < 2000"
    Case 2000 To 2009
        sWbStartYearDistr = "6. >= 2000 and < 2010"
    Case Else
        sWbStartYearDistr = "7. >= 2010"
    End Select
    Exit Function

ErrHandler:
    sWbStartYearDistr = ""
End Function


Public Function sWbAnnSalesDistr(sAnnSales As Variant) As String
    On Error GoTo ErrHandler
    
    Dim lAnnSales As Long
    
    If sAnnSales = "" Then GoTo ErrHandler
    lAnnSales = CLng(sAnnSales)
    
    Select Case lAnnSales
    Case 0 To 249999
        sWbAnnSalesDistr = "01. < $250K"
    Case 250000 To 499999
        sWbAnnSalesDistr = "02. >= $250K and < $500K"
    Case 500000 To 999999
        sWbAnnSalesDistr = "03. >= $500K and < $1M"
    Case 1000000 To 2499999
        sWbAnnSalesDistr = "04. >= $1M and < $2.5M"
    Case 2500000 To 4999999
        sWbAnnSalesDistr = "05. >= $2.5M and < $5M"
    Case 5000000 To 9999999
        sWbAnnSalesDistr = "06. >= $5M and < $10M"
    Case 10000000 To 24999999
        sWbAnnSalesDistr = "07. >= $10M and < $25M"
    Case 25000000 To 49999999
        sWbAnnSalesDistr = "08. >= $25M and < $50M"
    Case 50000000 To 99999999
        sWbAnnSalesDistr = "09. >= $50M and < $100M"
    Case 100000000 To 249999999
        sWbAnnSalesDistr = "10. >= $100M and < $250M"
    Case Else
        sWbAnnSalesDistr = "11. >= $250M"
    End Select
    Exit Function

ErrHandler:
    sWbAnnSalesDistr = ""
End Function


