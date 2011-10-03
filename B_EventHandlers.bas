Attribute VB_Name = "B_EventHandlers"
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

Option Private Module
Option Explicit


'Format header row
Private Sub FormatHdr(HdrRange As Range)
    With HdrRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Interior.ColorIndex = 15
        .Font.Italic = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub


'Get the Excel currency code format string associated with the WB currency code
Private Function sWbCurrCodeFormatStr(sCurrCode As String) As String
    On Error GoTo ErrHandler
    
    If sCurrCode = "" Or sCurrCode = "0000" Then GoTo ErrHandler
    
    sWbCurrCodeFormatStr = Application.WorksheetFunction.VLookup(sCurrCode, ThisWorkbook.Names("Curr_Codes").RefersToRange, 7, False)
    Exit Function
    
ErrHandler:
    sWbCurrCodeFormatStr = ""
End Function


'Apply the correct format to a cell containing an annual sales figure
Private Sub FormatAnnualSalesFigure(rCell As Range, sCurrCode As String)
    Dim sFormatStr As String
        
    If Not IsEmpty(rCell) Then
        sFormatStr = sWbCurrCodeFormatStr(sCurrCode)
        If Len(sFormatStr) > 0 Then rCell.NumberFormat = sFormatStr
    End If
End Sub


'Format and process the Lhv1784 layout for optimal use in Excel
Private Sub ProcessLhv1784Xl(wsLhv1784Raw As Worksheet, wsLhv1784Xl As Worksheet)
    Dim lRow As Long
    Dim lNumRows As Long
    Dim strDate As String

    With wsLhv1784Xl
        lRow = 1
        Do While Not IsEmpty(wsLhv1784Raw.Cells(lRow, Col_Raw_Duns))
            .Cells(lRow, Col_Xl_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Duns)
            lRow = lRow + 1
        Loop
                    
        lNumRows = lRow - 1
        
        For lRow = 1 To lNumRows
            .Cells(lRow, Col_Xl_System_ID) = wsLhv1784Raw.Cells(lRow, Col_Raw_System_ID)
            .Cells(lRow, Col_Xl_Record_ID) = wsLhv1784Raw.Cells(lRow, Col_Raw_Record_ID)
            .Cells(lRow, Col_Xl_Filler1) = wsLhv1784Raw.Cells(lRow, Col_Raw_Filler1)
            'DUNS already filled out
            .Cells(lRow, Col_Xl_Bus_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Bus_Nme)
            .Cells(lRow, Col_Xl_Trdg_style_1) = wsLhv1784Raw.Cells(lRow, Col_Raw_Trdg_style_1)
            .Cells(lRow, Col_Xl_Reg_Adr_ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Reg_Adr_ind)
            .Cells(lRow, Col_Xl_Adr1) = wsLhv1784Raw.Cells(lRow, Col_Raw_Adr1)
            .Cells(lRow, Col_Xl_Adr2) = wsLhv1784Raw.Cells(lRow, Col_Raw_Adr2)
            .Cells(lRow, Col_Xl_Post_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Post_Cd)
            .Cells(lRow, Col_Xl_City_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_City_Nme)
            .Cells(lRow, Col_Xl_State_Prov) = wsLhv1784Raw.Cells(lRow, Col_Raw_State_Prov)
            .Cells(lRow, Col_Xl_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ctry_Nme)
            .Cells(lRow, Col_Xl_City_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_City_Cd)
            .Cells(lRow, Col_Xl_Cnty_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Cnty_Cd)
            .Cells(lRow, Col_Xl_State_Prov_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_State_Prov_Cd)
            .Cells(lRow, Col_Xl_State_Prov_Abbr) = wsLhv1784Raw.Cells(lRow, Col_Raw_State_Prov_Abbr)
            .Cells(lRow, Col_Xl_Ctry_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ctry_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_ISO_Ctry_Cd) = "ISO Country Code"
            Else
                .Cells(lRow, Col_Xl_ISO_Ctry_Cd).FormulaR1C1 = "=sWbCountry2IsoCode(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Cont_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Cont_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Cont_Desc) = "Continent Code Description"
            Else
                .Cells(lRow, Col_Xl_Cont_Desc).FormulaR1C1 = "=sWbContDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_MAdr) = wsLhv1784Raw.Cells(lRow, Col_Raw_MAdr)
            .Cells(lRow, Col_Xl_MPost_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MPost_Cd)
            .Cells(lRow, Col_Xl_MCity_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCity_Nme)
            .Cells(lRow, Col_Xl_MCnty_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCnty_Nme)
            .Cells(lRow, Col_Xl_MState_Prov_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_MState_Prov_Nme)
            .Cells(lRow, Col_Xl_MCtry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCtry_Nme)
            .Cells(lRow, Col_Xl_MCity_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCity_Cd)
            .Cells(lRow, Col_Xl_MCnty_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCnty_Cd)
            .Cells(lRow, Col_Xl_MState_Prov_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MState_Prov_Cd)
            .Cells(lRow, Col_Xl_MState_Prov_Abbr) = wsLhv1784Raw.Cells(lRow, Col_Raw_MState_Prov_Abbr)
            .Cells(lRow, Col_Xl_MCtry_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCtry_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_ISO_MCtry_Cd) = "ISO Mailing Country Code"
            Else
                .Cells(lRow, Col_Xl_ISO_MCtry_Cd).FormulaR1C1 = "=sWbCountry2IsoCode(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_MCont_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_MCont_Cd)
            .Cells(lRow, Col_Xl_Ctry_Access_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ctry_Access_Cd)
            .Cells(lRow, Col_Xl_Tel) = wsLhv1784Raw.Cells(lRow, Col_Raw_Tel)
            .Cells(lRow, Col_Xl_Cable_Telx) = wsLhv1784Raw.Cells(lRow, Col_Raw_Cable_Telx)
            .Cells(lRow, Col_Xl_Fax) = wsLhv1784Raw.Cells(lRow, Col_Raw_Fax)
            .Cells(lRow, Col_Xl_NatID) = wsLhv1784Raw.Cells(lRow, Col_Raw_NatID)
            .Cells(lRow, Col_Xl_NatID_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_NatID_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_NatID_Desc) = "National Identification System Description"
            Else
                .Cells(lRow, Col_Xl_NatID_Desc).FormulaR1C1 = "=sWbNatIDDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_OOB_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_OOB_Ind)
            .Cells(lRow, Col_Xl_Marketable_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Marketable_Ind)
            .Cells(lRow, Col_Xl_Yr_Strt).Value = wsLhv1784Raw.Cells(lRow, Col_Raw_Yr_Strt)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Yr_Strt_Int) = "Year Started (Numeric)"
            Else
                .Cells(lRow, Col_Xl_Yr_Strt_Int).FormulaR1C1 = "=iWbStartYear(RC[-1])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Yr_Strt_Distr) = "Year Started (distribution)"
            Else
                .Cells(lRow, Col_Xl_Yr_Strt_Distr).FormulaR1C1 = "=sWbStartYearDistr(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Ann_Sales) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ann_Sales)
            .Cells(lRow, Col_Xl_Crcy_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Crcy_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_ISO_Crcy_Cd) = "Currency Code (ISO)"
            Else
                .Cells(lRow, Col_Xl_ISO_Crcy_Cd).FormulaR1C1 = "=sWbCurr2IsoCode(RC[-1])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Ann_Sales_Formatted) = "Annual Sales Local (formatted)"
            Else
                .Cells(lRow, Col_Xl_Ann_Sales_Formatted).FormulaR1C1 = "=lWbAnnSales(RC[-3])"
                FormatAnnualSalesFigure .Cells(lRow, Col_Xl_Ann_Sales_Formatted), .Cells(lRow, Col_Xl_Crcy_Cd).Value
            End If
            .Cells(lRow, Col_Xl_Ann_Sales_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ann_Sales_Ind)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Ann_Sales_Desc) = "Annual Sales Indicator Description"
            Else
                .Cells(lRow, Col_Xl_Ann_Sales_Desc).FormulaR1C1 = "=sWbAnnSalesIndDesc(RC[-5], RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Ann_Sales_US) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ann_Sales_US)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Ann_Sales_US_Formatted) = "Annual Sales USD (formatted)"
            Else
                .Cells(lRow, Col_Xl_Ann_Sales_US_Formatted).FormulaR1C1 = "=lWbAnnSales(RC[-1])"
                FormatAnnualSalesFigure .Cells(lRow, Col_Xl_Ann_Sales_US_Formatted), "0020"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Ann_Sales_US_Distr) = "Annual Sales USD (distribution)"
            Else
                .Cells(lRow, Col_Xl_Ann_Sales_US_Distr).FormulaR1C1 = "=sWbAnnSalesDistr(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Emp_Here) = wsLhv1784Raw.Cells(lRow, Col_Raw_Emp_Here)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Emp_Here_Int) = "Employees Here (Numeric)"
            Else
                .Cells(lRow, Col_Xl_Emp_Here_Int).FormulaR1C1 = "=lWbNumEmpl(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Emp_Here_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Emp_Here_Ind)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Emp_Here_Desc) = "Employees Here Indicator Description"
            Else
                .Cells(lRow, Col_Xl_Emp_Here_Desc).FormulaR1C1 = "=sWbNumEmplIndDesc(RC[-3], RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Emp_Tot) = wsLhv1784Raw.Cells(lRow, Col_Raw_Emp_Tot)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Emp_Tot_Int) = "Employees Total (Numeric)"
            Else
                .Cells(lRow, Col_Xl_Emp_Tot_Int).FormulaR1C1 = "=lWbNumEmpl(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Emp_Tot_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Emp_Tot_Ind)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Emp_Tot_Desc) = "Employees Total Indicator Description"
            Else
                .Cells(lRow, Col_Xl_Emp_Tot_Desc).FormulaR1C1 = "=sWbNumEmplIndDesc(RC[-3], RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Prin_Incl_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Prin_Incl_Ind)
            .Cells(lRow, Col_Xl_Imp_Exp_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Imp_Exp_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Imp_Exp_Desc) = "Import/Export/Agent Description"
            Else
                .Cells(lRow, Col_Xl_Imp_Exp_Desc).FormulaR1C1 = "=sWbImpExpCodeDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Lgl_Stat) = wsLhv1784Raw.Cells(lRow, Col_Raw_Lgl_Stat)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Lgl_Stat_Desc) = "Legal Status Description"
            Else
                .Cells(lRow, Col_Xl_Lgl_Stat_Desc).FormulaR1C1 = "=sWbLegalStatusCodeDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Control_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Control_Ind)
            .Cells(lRow, Col_Xl_Stat_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Stat_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Stat_Desc) = "Status Code Description"
            Else
                .Cells(lRow, Col_Xl_Stat_Desc).FormulaR1C1 = "=sWbStatusCodeDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Subs_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Subs_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Subs_Desc) = "Subsidiary Code Description"
            Else
                .Cells(lRow, Col_Xl_Subs_Desc).FormulaR1C1 = "=sWbSubsidiaryCodeDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_CEO_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_CEO_Nme)
            .Cells(lRow, Col_Xl_CEO_Title) = wsLhv1784Raw.Cells(lRow, Col_Raw_CEO_Title)
            .Cells(lRow, Col_Xl_Rep_Date) = wsLhv1784Raw.Cells(lRow, Col_Raw_Rep_Date)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Rep_Date_Date) = "Report Date (Date)"
            Else
                strDate = wsLhv1784Raw.Cells(lRow, Col_Raw_Rep_Date)
                If Len(strDate) > 0 Then
                    .Cells(lRow, Col_Xl_Rep_Date_Date) = Left(strDate, 4) & "/" & Mid(strDate, 5, 2) & "/" & Right(strDate, 2)
                    .Cells(lRow, Col_Xl_Rep_Date_Date).NumberFormat = "[$-409]mmm-yy;@"
                End If
            End If
            .Cells(lRow, Col_Xl_Filler2) = wsLhv1784Raw.Cells(lRow, Col_Raw_Filler2)
            .Cells(lRow, Col_Xl_Prev_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Prev_Duns)
            .Cells(lRow, Col_Xl_US1987_Sic1) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic1)
            .Cells(lRow, Col_Xl_Line_Bus_Descr) = wsLhv1784Raw.Cells(lRow, Col_Raw_Line_Bus_Descr)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_US1987_Sic1_Desc) = "Line of Business Description"
            Else
                .Cells(lRow, Col_Xl_US1987_Sic1_Desc) = "=sWbSIC87Desc(RC[-2])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_US1987_Sic1_2) = "U.S. 1987 SIC 1 (2 Digit)"
            Else
                .Cells(lRow, Col_Xl_US1987_Sic1_2) = "=sWbSIC_2Code(RC[-3])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_US1987_Sic1_2Desc) = "U.S. 1987 SIC 1 (2 Digit) Description"
            Else
                .Cells(lRow, Col_Xl_US1987_Sic1_2Desc) = "=sWbSIC87_2Desc(RC[-4])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_US1987_Sic1_1) = "U.S. 1987 SIC 1 (1 Digit)"
            Else
                .Cells(lRow, Col_Xl_US1987_Sic1_1) = "=sWbSIC_1Code(RC[-5])"
            End If
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_US1987_Sic1_1Desc) = "U.S. 1987 SIC 1 (1 Digit) Description"
            Else
                .Cells(lRow, Col_Xl_US1987_Sic1_1Desc) = "=sWbSIC87_1Desc(RC[-6])"
            End If
            .Cells(lRow, Col_Xl_US1987_Sic2) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic2)
            .Cells(lRow, Col_Xl_US1987_Sic3) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic3)
            .Cells(lRow, Col_Xl_US1987_Sic4) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic4)
            .Cells(lRow, Col_Xl_US1987_Sic5) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic5)
            .Cells(lRow, Col_Xl_US1987_Sic6) = wsLhv1784Raw.Cells(lRow, Col_Raw_US1987_Sic6)
            .Cells(lRow, Col_Xl_Prmy_Loc_Act_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Prmy_Loc_Act_Cd)
            .Cells(lRow, Col_Xl_Act_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Act_Ind)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Act_Desc) = "Local Activity Code Description"
            Else
                .Cells(lRow, Col_Xl_Act_Desc) = "=sWbLocalActCodeDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_HQ_Prnt_Fil) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Fil)
            .Cells(lRow, Col_Xl_HQ_Prnt_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Duns)
            .Cells(lRow, Col_Xl_HQ_Prnt_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Nme)
            .Cells(lRow, Col_Xl_HQ_Prnt_Adr) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Adr)
            .Cells(lRow, Col_Xl_HQ_Prnt_Post_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Post_Cd)
            .Cells(lRow, Col_Xl_HQ_Prnt_City_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_City_Nme)
            .Cells(lRow, Col_Xl_HQ_Prnt_State_Prov_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_State_Prov_Nme)
            .Cells(lRow, Col_Xl_HQ_Prnt_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Ctry_Nme)
            .Cells(lRow, Col_Xl_HQ_Prnt_City_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_City_Cd)
            .Cells(lRow, Col_Xl_HQ_Prnt_Cnty_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Cnty_Cd)
            .Cells(lRow, Col_Xl_HQ_Prnt_State_Prov_Abbr) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_State_Prov_Abbr)
            .Cells(lRow, Col_Xl_HQ_Prnt_Ctry_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Ctry_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_HQ_Prnt_ISO_Ctry_Cd) = "Headquarter/Parent ISO Country Code"
            Else
                .Cells(lRow, Col_Xl_HQ_Prnt_ISO_Ctry_Cd).FormulaR1C1 = "=sWbCountry2IsoCode(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_HQ_Prnt_Cont_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_HQ_Prnt_Cont_Cd)
            .Cells(lRow, Col_Xl_Dom_Ult_Fil) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Fil)
            .Cells(lRow, Col_Xl_Dom_Ult_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Duns)
            .Cells(lRow, Col_Xl_Dom_Ult_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Nme)
            .Cells(lRow, Col_Xl_Dom_Ult_Adr) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Adr)
            .Cells(lRow, Col_Xl_Dom_Ult_Post_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Post_Cd)
            .Cells(lRow, Col_Xl_Dom_Ult_City_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_City_Nme)
            .Cells(lRow, Col_Xl_Dom_Ult_State_Prov_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_State_Prov_Nme)
            .Cells(lRow, Col_Xl_Dom_Ult_City_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_City_Cd)
            .Cells(lRow, Col_Xl_Dom_Ult_State_Prov_Abbr) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_State_Prov_Abbr)
            .Cells(lRow, Col_Xl_Dom_Ult_Ctry_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Dom_Ult_Ctry_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Dom_Ult_ISO_Ctry_Cd) = "Domestic Ultimate ISO Country Code"
            Else
                .Cells(lRow, Col_Xl_Dom_Ult_ISO_Ctry_Cd).FormulaR1C1 = "=sWbCountry2IsoCode(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Gbl_Ult_Ind) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Ind)
            .Cells(lRow, Col_Xl_Gbl_Ult_Fil) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Fil)
            .Cells(lRow, Col_Xl_Gbl_Ult_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Duns)
            .Cells(lRow, Col_Xl_Gbl_Ult_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Nme)
            .Cells(lRow, Col_Xl_Gbl_Ult_Adr) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Adr)
            .Cells(lRow, Col_Xl_Gbl_Ult_Post_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Post_Cd)
            .Cells(lRow, Col_Xl_Gbl_Ult_City_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_City_Nme)
            .Cells(lRow, Col_Xl_Gbl_Ult_State_Prov_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_State_Prov_Nme)
            .Cells(lRow, Col_Xl_Gbl_Ult_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Ctry_Nme)
            .Cells(lRow, Col_Xl_Gbl_Ult_City_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_City_Cd)
            .Cells(lRow, Col_Xl_Gbl_Ult_Cnty_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Cnty_Cd)
            .Cells(lRow, Col_Xl_Gbl_Ult_State_Prov_Abbr) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_State_Prov_Abbr)
            .Cells(lRow, Col_Xl_Gbl_Ult_Ctry_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Ctry_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Gbl_Ult_ISO_Ctry_Cd) = "Global Ultimate ISO Country Code"
            Else
                .Cells(lRow, Col_Xl_Gbl_Ult_ISO_Ctry_Cd).FormulaR1C1 = "=sWbCountry2IsoCode(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Gbl_Ult_Cont_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Cont_Cd)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Gbl_Ult_Cont_Desc) = "Global Ultimate Continent Code Description"
            Else
                .Cells(lRow, Col_Xl_Gbl_Ult_Cont_Desc).FormulaR1C1 = "=sWbContDesc(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Fam_Memb_Gbl) = wsLhv1784Raw.Cells(lRow, Col_Raw_Fam_Memb_Gbl)
            If lRow = 1 Then
                .Cells(lRow, Col_Xl_Fam_Memb_Gbl_Int) = "Number of Family Members (Numeric)"
            Else
                .Cells(lRow, Col_Xl_Fam_Memb_Gbl_Int).FormulaR1C1 = "=lWbNumFamMembers(RC[-1])"
            End If
            .Cells(lRow, Col_Xl_Gbl_Dias_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Dias_Cd)
            .Cells(lRow, Col_Xl_Gbl_Hier_Cd) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Hier_Cd)
            .Cells(lRow, Col_Xl_Fam_Upd_Date) = wsLhv1784Raw.Cells(lRow, Col_Raw_Fam_Upd_Date)
        Next

        FormatHdr .Cells(1, Col_Xl_System_ID)
        FormatHdr .Cells(1, Col_Xl_Record_ID)
        FormatHdr .Cells(1, Col_Xl_Filler1)
        FormatHdr .Cells(1, Col_Xl_Duns)
        FormatHdr .Cells(1, Col_Xl_Bus_Nme)
        FormatHdr .Cells(1, Col_Xl_Trdg_style_1)
        FormatHdr .Cells(1, Col_Xl_Reg_Adr_ind)
        FormatHdr .Cells(1, Col_Xl_Adr1)
        FormatHdr .Cells(1, Col_Xl_Adr2)
        FormatHdr .Cells(1, Col_Xl_Post_Cd)
        FormatHdr .Cells(1, Col_Xl_City_Nme)
        FormatHdr .Cells(1, Col_Xl_State_Prov)
        FormatHdr .Cells(1, Col_Xl_Ctry_Nme)
        FormatHdr .Cells(1, Col_Xl_City_Cd)
        FormatHdr .Cells(1, Col_Xl_Cnty_Cd)
        FormatHdr .Cells(1, Col_Xl_State_Prov_Cd)
        FormatHdr .Cells(1, Col_Xl_State_Prov_Abbr)
        FormatHdr .Cells(1, Col_Xl_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_ISO_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Cont_Cd)
        FormatHdr .Cells(1, Col_Xl_Cont_Desc)
        FormatHdr .Cells(1, Col_Xl_MAdr)
        FormatHdr .Cells(1, Col_Xl_MPost_Cd)
        FormatHdr .Cells(1, Col_Xl_MCity_Nme)
        FormatHdr .Cells(1, Col_Xl_MCnty_Nme)
        FormatHdr .Cells(1, Col_Xl_MState_Prov_Nme)
        FormatHdr .Cells(1, Col_Xl_MCtry_Nme)
        FormatHdr .Cells(1, Col_Xl_MCity_Cd)
        FormatHdr .Cells(1, Col_Xl_MCnty_Cd)
        FormatHdr .Cells(1, Col_Xl_MState_Prov_Cd)
        FormatHdr .Cells(1, Col_Xl_MState_Prov_Abbr)
        FormatHdr .Cells(1, Col_Xl_MCtry_Cd)
        FormatHdr .Cells(1, Col_Xl_MCtry_Cd)
        FormatHdr .Cells(1, Col_Xl_ISO_MCtry_Cd)
        FormatHdr .Cells(1, Col_Xl_MCont_Cd)
        FormatHdr .Cells(1, Col_Xl_Ctry_Access_Cd)
        FormatHdr .Cells(1, Col_Xl_Tel)
        FormatHdr .Cells(1, Col_Xl_Cable_Telx)
        FormatHdr .Cells(1, Col_Xl_Fax)
        FormatHdr .Cells(1, Col_Xl_NatID)
        FormatHdr .Cells(1, Col_Xl_NatID_Cd)
        FormatHdr .Cells(1, Col_Xl_NatID_Desc)
        FormatHdr .Cells(1, Col_Xl_OOB_Ind)
        FormatHdr .Cells(1, Col_Xl_Marketable_Ind)
        FormatHdr .Cells(1, Col_Xl_Yr_Strt)
        FormatHdr .Cells(1, Col_Xl_Yr_Strt_Int)
        FormatHdr .Cells(1, Col_Xl_Yr_Strt_Distr)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales)
        FormatHdr .Cells(1, Col_Xl_Crcy_Cd)
        FormatHdr .Cells(1, Col_Xl_ISO_Crcy_Cd)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_Formatted)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_Ind)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_Desc)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_US)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_US_Formatted)
        FormatHdr .Cells(1, Col_Xl_Ann_Sales_US_Distr)
        FormatHdr .Cells(1, Col_Xl_Emp_Here)
        FormatHdr .Cells(1, Col_Xl_Emp_Here_Int)
        FormatHdr .Cells(1, Col_Xl_Emp_Here_Ind)
        FormatHdr .Cells(1, Col_Xl_Emp_Here_Desc)
        FormatHdr .Cells(1, Col_Xl_Emp_Tot)
        FormatHdr .Cells(1, Col_Xl_Emp_Tot_Int)
        FormatHdr .Cells(1, Col_Xl_Emp_Tot_Ind)
        FormatHdr .Cells(1, Col_Xl_Emp_Tot_Desc)
        FormatHdr .Cells(1, Col_Xl_Prin_Incl_Ind)
        FormatHdr .Cells(1, Col_Xl_Imp_Exp_Cd)
        FormatHdr .Cells(1, Col_Xl_Imp_Exp_Desc)
        FormatHdr .Cells(1, Col_Xl_Lgl_Stat)
        FormatHdr .Cells(1, Col_Xl_Lgl_Stat_Desc)
        FormatHdr .Cells(1, Col_Xl_Control_Ind)
        FormatHdr .Cells(1, Col_Xl_Stat_Cd)
        FormatHdr .Cells(1, Col_Xl_Stat_Desc)
        FormatHdr .Cells(1, Col_Xl_Subs_Cd)
        FormatHdr .Cells(1, Col_Xl_Subs_Desc)
        FormatHdr .Cells(1, Col_Xl_CEO_Nme)
        FormatHdr .Cells(1, Col_Xl_CEO_Title)
        FormatHdr .Cells(1, Col_Xl_Rep_Date)
        FormatHdr .Cells(1, Col_Xl_Rep_Date_Date)
        FormatHdr .Cells(1, Col_Xl_Filler2)
        FormatHdr .Cells(1, Col_Xl_Prev_Duns)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1)
        FormatHdr .Cells(1, Col_Xl_Line_Bus_Descr)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1_Desc)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1_2)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1_2Desc)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1_1)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic1_1Desc)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic2)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic3)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic4)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic5)
        FormatHdr .Cells(1, Col_Xl_US1987_Sic6)
        FormatHdr .Cells(1, Col_Xl_Prmy_Loc_Act_Cd)
        FormatHdr .Cells(1, Col_Xl_Act_Ind)
        FormatHdr .Cells(1, Col_Xl_Act_Desc)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Fil)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Duns)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Nme)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Adr)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Post_Cd)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_City_Nme)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_State_Prov_Nme)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Ctry_Nme)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_City_Cd)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Cnty_Cd)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_State_Prov_Abbr)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_ISO_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_HQ_Prnt_Cont_Cd)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Fil)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Duns)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Nme)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Adr)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Post_Cd)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_City_Nme)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_State_Prov_Nme)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_City_Cd)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_State_Prov_Abbr)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Dom_Ult_ISO_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Ind)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Fil)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Duns)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Nme)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Adr)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Post_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_City_Nme)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_State_Prov_Nme)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Ctry_Nme)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_City_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Cnty_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_State_Prov_Abbr)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_ISO_Ctry_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Cont_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Ult_Cont_Desc)
        FormatHdr .Cells(1, Col_Xl_Fam_Memb_Gbl)
        FormatHdr .Cells(1, Col_Xl_Fam_Memb_Gbl_Int)
        FormatHdr .Cells(1, Col_Xl_Gbl_Dias_Cd)
        FormatHdr .Cells(1, Col_Xl_Gbl_Hier_Cd)
        FormatHdr .Cells(1, Col_Xl_Fam_Upd_Date)
        
        .Columns(Col_Xl_System_ID).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Record_ID).ColumnWidth = Col_Width_12
        .Columns(Col_Xl_Filler1).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Xl_Bus_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_Xl_Trdg_style_1).ColumnWidth = Col_Width_Name2
        .Columns(Col_Xl_Reg_Adr_ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Adr1).ColumnWidth = Col_Width_Adr1
        .Columns(Col_Xl_Adr2).ColumnWidth = Col_Width_Adr2
        .Columns(Col_Xl_Post_Cd).ColumnWidth = Col_Width_Post_Cd
        .Columns(Col_Xl_City_Nme).ColumnWidth = Col_Width_City_Nme
        .Columns(Col_Xl_State_Prov).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
        .Columns(Col_Xl_City_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Cnty_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_State_Prov_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_State_Prov_Abbr).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_ISO_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Cont_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Cont_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_MAdr).ColumnWidth = Col_Width_MAdr
        .Columns(Col_Xl_MPost_Cd).ColumnWidth = Col_Width_Post_Cd
        .Columns(Col_Xl_MCity_Nme).ColumnWidth = Col_Width_City_Nme
        .Columns(Col_Xl_MCnty_Nme).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_MState_Prov_Nme).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_MCtry_Nme).ColumnWidth = Col_Width_Ctry_Nme
        .Columns(Col_Xl_MCity_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_MCnty_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_MState_Prov_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_MState_Prov_Abbr).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_MCtry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_ISO_MCtry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_MCont_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Ctry_Access_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Tel).ColumnWidth = Col_Width_Tel
        .Columns(Col_Xl_Cable_Telx).ColumnWidth = Col_Width_Tel
        .Columns(Col_Xl_Fax).ColumnWidth = Col_Width_Tel
        .Columns(Col_Xl_NatID).ColumnWidth = Col_Width_NatID
        .Columns(Col_Xl_NatID_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_NatID_Desc).ColumnWidth = Col_Width_Desc
        .Columns(Col_Xl_OOB_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Marketable_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Yr_Strt).ColumnWidth = Col_Width_Year
        .Columns(Col_Xl_Yr_Strt_Int).ColumnWidth = Col_Width_Year
        .Columns(Col_Xl_Yr_Strt_Distr).ColumnWidth = Col_Width_Desc_Med
        .Columns(Col_Xl_Ann_Sales).ColumnWidth = Col_Width_Large_Str_Num
        .Columns(Col_Xl_Crcy_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_ISO_Crcy_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Ann_Sales_Formatted).ColumnWidth = Col_Width_Large_Curr_Num
        .Columns(Col_Xl_Ann_Sales_Ind).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Ann_Sales_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Ann_Sales_US).ColumnWidth = Col_Width_Large_Str_Num
        .Columns(Col_Xl_Ann_Sales_US_Formatted).ColumnWidth = Col_Width_Large_Curr_Num
        .Columns(Col_Xl_Ann_Sales_US_Distr).ColumnWidth = Col_Width_Desc_Med
        .Columns(Col_Xl_Emp_Here).ColumnWidth = Col_Width_Small_Str_Num
        .Columns(Col_Xl_Emp_Here_Int).ColumnWidth = Col_Width_Small_Int_Num
        .Columns(Col_Xl_Emp_Here_Ind).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Emp_Here_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Emp_Tot).ColumnWidth = Col_Width_Small_Str_Num
        .Columns(Col_Xl_Emp_Tot_Int).ColumnWidth = Col_Width_Small_Int_Num
        .Columns(Col_Xl_Emp_Tot_Ind).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Emp_Tot_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Prin_Incl_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Imp_Exp_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Imp_Exp_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Lgl_Stat).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Lgl_Stat_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Control_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Stat_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Stat_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Subs_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Subs_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_CEO_Nme).ColumnWidth = Col_Width_CEO_Name
        .Columns(Col_Xl_CEO_Title).ColumnWidth = Col_Width_CEO_Title
        .Columns(Col_Xl_Rep_Date).ColumnWidth = Col_Width_Str_Date
        .Columns(Col_Xl_Rep_Date_Date).ColumnWidth = Col_Width_Date_Date
        .Columns(Col_Xl_Filler2).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Prev_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Xl_US1987_Sic1).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_Line_Bus_Descr).ColumnWidth = Col_Width_Desc
        .Columns(Col_Xl_US1987_Sic1_Desc).ColumnWidth = Col_Width_Desc
        .Columns(Col_Xl_US1987_Sic1_2).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic1_2Desc).ColumnWidth = Col_Width_Desc
        .Columns(Col_Xl_US1987_Sic1_1).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic1_1Desc).ColumnWidth = Col_Width_Desc
        .Columns(Col_Xl_US1987_Sic2).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic3).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic4).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic5).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_US1987_Sic6).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_Prmy_Loc_Act_Cd).ColumnWidth = Col_Width_SIC
        .Columns(Col_Xl_Act_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Act_Desc).ColumnWidth = Col_Width_Desc_Med
        .Columns(Col_Xl_HQ_Prnt_Fil).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_HQ_Prnt_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Xl_HQ_Prnt_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_Xl_HQ_Prnt_Adr).ColumnWidth = Col_Width_Adr1
        .Columns(Col_Xl_HQ_Prnt_Post_Cd).ColumnWidth = Col_Width_Post_Cd
        .Columns(Col_Xl_HQ_Prnt_City_Nme).ColumnWidth = Col_Width_City_Nme
        .Columns(Col_Xl_HQ_Prnt_State_Prov_Nme).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_HQ_Prnt_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
        .Columns(Col_Xl_HQ_Prnt_City_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_HQ_Prnt_Cnty_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_HQ_Prnt_State_Prov_Abbr).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_HQ_Prnt_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_HQ_Prnt_ISO_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_HQ_Prnt_Cont_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Dom_Ult_Fil).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Dom_Ult_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Xl_Dom_Ult_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_Xl_Dom_Ult_Adr).ColumnWidth = Col_Width_Adr1
        .Columns(Col_Xl_Dom_Ult_Post_Cd).ColumnWidth = Col_Width_Post_Cd
        .Columns(Col_Xl_Dom_Ult_City_Nme).ColumnWidth = Col_Width_City_Nme
        .Columns(Col_Xl_Dom_Ult_State_Prov_Nme).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_Dom_Ult_City_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Dom_Ult_State_Prov_Abbr).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Dom_Ult_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Dom_Ult_ISO_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Gbl_Ult_Ind).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Gbl_Ult_Fil).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Gbl_Ult_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Xl_Gbl_Ult_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_Xl_Gbl_Ult_Adr).ColumnWidth = Col_Width_Adr1
        .Columns(Col_Xl_Gbl_Ult_Post_Cd).ColumnWidth = Col_Width_Post_Cd
        .Columns(Col_Xl_Gbl_Ult_City_Nme).ColumnWidth = Col_Width_City_Nme
        .Columns(Col_Xl_Gbl_Ult_State_Prov_Nme).ColumnWidth = Col_Width_18
        .Columns(Col_Xl_Gbl_Ult_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
        .Columns(Col_Xl_Gbl_Ult_City_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Gbl_Ult_Cnty_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Gbl_Ult_State_Prov_Abbr).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Gbl_Ult_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Gbl_Ult_ISO_Ctry_Cd).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Gbl_Ult_Cont_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Gbl_Ult_Cont_Desc).ColumnWidth = Col_Width_Desc_Short
        .Columns(Col_Xl_Fam_Memb_Gbl).ColumnWidth = Col_Width_8
        .Columns(Col_Xl_Fam_Memb_Gbl_Int).ColumnWidth = Col_Width_Small_Int_Num
        .Columns(Col_Xl_Gbl_Dias_Cd).ColumnWidth = Col_Width_12
        .Columns(Col_Xl_Gbl_Hier_Cd).ColumnWidth = Col_Width_6
        .Columns(Col_Xl_Fam_Upd_Date).ColumnWidth = Col_Width_Str_Date

        'Hide a couple of columns
        .Columns(Col_Xl_Filler1).Hidden = True
        .Columns(Col_Xl_Reg_Adr_ind).Hidden = True
        .Columns(Col_Xl_State_Prov).Hidden = True
        .Columns(Col_Xl_Ctry_Nme).Hidden = True
        .Columns(Col_Xl_City_Cd).Hidden = True
        .Columns(Col_Xl_Cnty_Cd).Hidden = True
        .Columns(Col_Xl_State_Prov_Cd).Hidden = True
        .Columns(Col_Xl_State_Prov_Abbr).Hidden = True
        .Columns(Col_Xl_Ctry_Cd).Hidden = True
        .Columns(Col_Xl_Cont_Cd).Hidden = True
        .Columns(Col_Xl_Cont_Desc).Hidden = True
        .Columns(Col_Xl_MCnty_Nme).Hidden = True
        .Columns(Col_Xl_MState_Prov_Nme).Hidden = True
        .Columns(Col_Xl_MCtry_Nme).Hidden = True
        .Columns(Col_Xl_MCity_Cd).Hidden = True
        .Columns(Col_Xl_MCnty_Cd).Hidden = True
        .Columns(Col_Xl_MState_Prov_Cd).Hidden = True
        .Columns(Col_Xl_MState_Prov_Abbr).Hidden = True
        .Columns(Col_Xl_MCtry_Cd).Hidden = True
        .Columns(Col_Xl_MCont_Cd).Hidden = True
        .Columns(Col_Xl_Cable_Telx).Hidden = True
        .Columns(Col_Xl_NatID_Cd).Hidden = True
        .Columns(Col_Xl_NatID_Desc).Hidden = True
        .Columns(Col_Xl_Marketable_Ind).Hidden = True
        .Columns(Col_Xl_Yr_Strt).Hidden = True
        .Columns(Col_Xl_Yr_Strt_Distr).Hidden = True
        .Columns(Col_Xl_Ann_Sales).Hidden = True
        .Columns(Col_Xl_Crcy_Cd).Hidden = True
        .Columns(Col_Xl_ISO_Crcy_Cd).Hidden = True
        .Columns(Col_Xl_Ann_Sales_Ind).Hidden = True
        .Columns(Col_Xl_Ann_Sales_US).Hidden = True
        .Columns(Col_Xl_Ann_Sales_US_Distr).Hidden = True
        .Columns(Col_Xl_Emp_Here).Hidden = True
        .Columns(Col_Xl_Emp_Here_Ind).Hidden = True
        .Columns(Col_Xl_Emp_Tot).Hidden = True
        .Columns(Col_Xl_Emp_Tot_Ind).Hidden = True
        .Columns(Col_Xl_Imp_Exp_Cd).Hidden = True
        .Columns(Col_Xl_Lgl_Stat).Hidden = True
        .Columns(Col_Xl_Control_Ind).Hidden = True
        .Columns(Col_Xl_Stat_Cd).Hidden = True
        .Columns(Col_Xl_Stat_Desc).Hidden = True
        .Columns(Col_Xl_Subs_Cd).Hidden = True
        .Columns(Col_Xl_Subs_Desc).Hidden = True
        .Columns(Col_Xl_Rep_Date).Hidden = True
        .Columns(Col_Xl_Rep_Date_Date).Hidden = True
        .Columns(Col_Xl_Filler2).Hidden = True
        .Columns(Col_Xl_Prev_Duns).Hidden = True
        .Columns(Col_Xl_Line_Bus_Descr).Hidden = True
        .Columns(Col_Xl_Act_Ind).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_Fil).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_State_Prov_Nme).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_Ctry_Nme).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_City_Cd).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_Cnty_Cd).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_State_Prov_Abbr).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_Ctry_Cd).Hidden = True
        .Columns(Col_Xl_HQ_Prnt_Cont_Cd).Hidden = True
        .Columns(Col_Xl_Dom_Ult_Fil).Hidden = True
        .Columns(Col_Xl_Dom_Ult_State_Prov_Nme).Hidden = True
        .Columns(Col_Xl_Dom_Ult_City_Cd).Hidden = True
        .Columns(Col_Xl_Dom_Ult_State_Prov_Abbr).Hidden = True
        .Columns(Col_Xl_Dom_Ult_Ctry_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Ind).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Fil).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_State_Prov_Nme).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Ctry_Nme).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_City_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Cnty_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_State_Prov_Abbr).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Ctry_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Cont_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Ult_Cont_Desc).Hidden = True
        .Columns(Col_Xl_Fam_Memb_Gbl).Hidden = True
        .Columns(Col_Xl_Fam_Memb_Gbl_Int).Hidden = True
        .Columns(Col_Xl_Gbl_Dias_Cd).Hidden = True
        .Columns(Col_Xl_Gbl_Hier_Cd).Hidden = True
        .Columns(Col_Xl_Fam_Upd_Date).Hidden = True 'Use the last column for defining the name below!
    End With

    'Name the range containing the D&B data enrichment
    If XlNameExists(sXlNameDnbData) Then ThisWorkbook.Names(sXlNameDnbData).Delete
    ThisWorkbook.Names.Add Name:=sXlNameDnbData, RefersToR1C1:="=" & sLhv1784Xl & "!R1C1:R" & lNumRows & "C" & Col_Xl_Fam_Upd_Date
End Sub


'Create reports listing the DUNS & GU duplicates
Private Sub CreateDuplicateReports(wsLhv1784Raw As Worksheet)
    Dim wsDunsDupsReport As Worksheet
    Dim lRow As Long, lNumRows As Long
    Dim rDups As Range
    Dim vColorIndex As Variant
    
    Const Col_Dups_System_ID = 1
    Const Col_Dups_Sap_Cust_Num = 2
    Const Col_Dups_Duns = 3
    Const Col_Dups_Bus_Nme = 4
    Const Col_Dups_Ctry_Nme = 5
    
    Const Col_GU_Dups_System_ID = 1
    Const Col_GU_Dups_Sap_Cust_Num = 2
    Const Col_GU_Dups_Duns = 3
    Const Col_GU_Dups_GU_Duns = 4
    Const Col_GU_Dups_Bus_Nme = 5
    Const Col_GU_Dups_Ctry_Nme = 6
    Const Col_GU_Dups_GU_Nme = 7
    Const Col_GU_Dups_GU_Ctry_Nme = 8
    
    'Create the duplicate DUNS report
    Set wsDunsDupsReport = ThisWorkbook.Worksheets.Add(After:=ActiveSheet, Type:=xlWorksheet)
    
    With wsDunsDupsReport
        .Name = sDunsDupsReport
        .Cells(2, 4).Select
        ActiveWindow.FreezePanes = True
        ActiveWindow.Zoom = 80
        .Rows(1).RowHeight = 25.5

        lRow = 1

        Do While Not IsEmpty(wsLhv1784Raw.Cells(lRow, Col_Raw_Duns))
            .Cells(lRow, Col_Dups_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Duns)
            lRow = lRow + 1
        Loop
        
        lNumRows = lRow - 1
    
        For lRow = 1 To lNumRows
            .Cells(lRow, Col_Dups_System_ID) = wsLhv1784Raw.Cells(lRow, Col_Raw_System_ID)
            .Cells(lRow, Col_Dups_Sap_Cust_Num) = wsLhv1784Raw.Cells(lRow, Col_Raw_Record_ID)
            'DUNS already filled out
            .Cells(lRow, Col_Dups_Bus_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Bus_Nme)
            .Cells(lRow, Col_Dups_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ctry_Nme)
        Next

        FormatHdr .Cells(1, Col_Dups_System_ID)
        FormatHdr .Cells(1, Col_Dups_Sap_Cust_Num)
        FormatHdr .Cells(1, Col_Dups_Duns)
        FormatHdr .Cells(1, Col_Dups_Bus_Nme)
        FormatHdr .Cells(1, Col_Dups_Ctry_Nme)

        .Columns(Col_Dups_System_ID).ColumnWidth = Col_Width_6
        .Columns(Col_Dups_Sap_Cust_Num).ColumnWidth = Col_Width_12
        .Columns(Col_Dups_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_Dups_Bus_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_Dups_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
    End With
    
    'Name the range containing the DUNS duplicates
    If XlNameExists(sXlNameDnbDunsDups) Then ThisWorkbook.Names(sXlNameDnbDunsDups).Delete
    ThisWorkbook.Names.Add Name:=sXlNameDnbDunsDups, RefersToR1C1:="=" & sDunsDupsReport & "!R1C1:R" & lNumRows & "C" & Col_Dups_Ctry_Nme

    'Perform a sort on DUNS
    Set rDups = ThisWorkbook.Names(sXlNameDnbDunsDups).RefersToRange
    With rDups
        .Sort Key1:=.Cells(1, Col_Dups_Duns), Header:=xlYes, DataOption1:=xlSortNormal

        'Highlight the DUNS duplicates
        For lRow = 3 To .Rows.Count
            If .Cells(lRow, Col_Dups_Duns).Offset(-1, 0).Value = .Cells(lRow, Col_Dups_Duns).Value Then
                'Mark the duplicates
                .Cells(lRow, Col_Dups_Duns).Offset(-1, 0).Interior.ColorIndex = 36
                .Cells(lRow, Col_Dups_Duns).Interior.ColorIndex = 36
            End If
        Next
    
        'Remove all singular DUNS
        For lRow = 2 To .Rows.Count
            If Not .Cells(lRow, Col_Dups_Duns).Interior.ColorIndex = 36 Then
                .Cells(lRow, Col_Dups_Duns).EntireRow.Delete
                If lRow > .Rows.Count Then
                    Exit For
                Else
                    lRow = lRow - 1
                End If
            End If
        Next
    
        'Apply formating to duplicates
        vColorIndex = 34
        For lRow = 2 To .Rows.Count
            .Cells(lRow, Col_Dups_Duns).Interior.ColorIndex = xlColorIndexNone
            Range(.Cells(lRow, Col_Dups_Bus_Nme), .Cells(lRow, Col_Dups_Ctry_Nme)).Interior.ColorIndex = vColorIndex
            If Not (.Cells(lRow, Col_Dups_Duns).Offset(1, 0).Value = .Cells(lRow, Col_Dups_Duns).Value) Then
                If vColorIndex = 34 Then
                    vColorIndex = 37
                Else
                    vColorIndex = 34
                End If
            End If
        Next
    End With

    'Create the duplicate Global Ultimate DUNS report
    Set wsDunsDupsReport = ThisWorkbook.Worksheets.Add(After:=ActiveSheet, Type:=xlWorksheet)
    
    With wsDunsDupsReport
        .Name = sGblUltDupsReport
        .Cells(2, 5).Select
        ActiveWindow.FreezePanes = True
        ActiveWindow.Zoom = 80
        .Rows(1).RowHeight = 25.5

        lRow = 1

        Do While Not IsEmpty(wsLhv1784Raw.Cells(lRow, Col_Raw_Duns))
            .Cells(lRow, Col_GU_Dups_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Duns)
            lRow = lRow + 1
        Loop
        
        lNumRows = lRow - 1
    
        For lRow = 1 To lNumRows
            .Cells(lRow, Col_GU_Dups_System_ID) = wsLhv1784Raw.Cells(lRow, Col_Raw_System_ID)
            .Cells(lRow, Col_GU_Dups_Sap_Cust_Num) = wsLhv1784Raw.Cells(lRow, Col_Raw_Record_ID)
            'DUNS already filled out
            .Cells(lRow, Col_GU_Dups_GU_Duns) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Duns)
            .Cells(lRow, Col_GU_Dups_Bus_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Bus_Nme)
            .Cells(lRow, Col_GU_Dups_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Ctry_Nme)
            .Cells(lRow, Col_GU_Dups_GU_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Nme)
            .Cells(lRow, Col_GU_Dups_GU_Ctry_Nme) = wsLhv1784Raw.Cells(lRow, Col_Raw_Gbl_Ult_Ctry_Nme)
        Next

        FormatHdr .Cells(1, Col_GU_Dups_System_ID)
        FormatHdr .Cells(1, Col_GU_Dups_Sap_Cust_Num)
        FormatHdr .Cells(1, Col_GU_Dups_Duns)
        FormatHdr .Cells(1, Col_GU_Dups_GU_Duns)
        FormatHdr .Cells(1, Col_GU_Dups_Bus_Nme)
        FormatHdr .Cells(1, Col_GU_Dups_Ctry_Nme)
        FormatHdr .Cells(1, Col_GU_Dups_GU_Nme)
        FormatHdr .Cells(1, Col_GU_Dups_GU_Ctry_Nme)

        .Columns(Col_GU_Dups_System_ID).ColumnWidth = Col_Width_6
        .Columns(Col_GU_Dups_Sap_Cust_Num).ColumnWidth = Col_Width_12
        .Columns(Col_GU_Dups_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_GU_Dups_GU_Duns).ColumnWidth = Col_Width_DUNS
        .Columns(Col_GU_Dups_Bus_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_GU_Dups_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
        .Columns(Col_GU_Dups_GU_Nme).ColumnWidth = Col_Width_Name1
        .Columns(Col_GU_Dups_GU_Ctry_Nme).ColumnWidth = Col_Width_Ctry_Nme
    End With
    
    'Name the range containing the DUNS duplicates
    If XlNameExists(sXlNameDnbGblUltDups) Then ThisWorkbook.Names(sXlNameDnbGblUltDups).Delete
    ThisWorkbook.Names.Add Name:=sXlNameDnbGblUltDups, RefersToR1C1:="=" & sGblUltDupsReport & "!R1C1:R" & lNumRows & "C" & Col_GU_Dups_GU_Ctry_Nme

    'Perform a sort on Global Ultimate DUNS
    Set rDups = ThisWorkbook.Names(sXlNameDnbGblUltDups).RefersToRange
    With rDups
        .Sort Key1:=.Cells(1, Col_GU_Dups_GU_Duns), Key2:=.Cells(1, Col_GU_Dups_Duns), Header:=xlYes, DataOption1:=xlSortNormal

        'Highlight the DUNS duplicates
        For lRow = 3 To .Rows.Count
            If Not IsEmpty(.Cells(lRow, Col_GU_Dups_GU_Duns)) And (.Cells(lRow, Col_GU_Dups_GU_Duns).Offset(-1, 0).Value = .Cells(lRow, Col_GU_Dups_GU_Duns).Value) Then
                'Mark the duplicates
                .Cells(lRow, Col_GU_Dups_GU_Duns).Offset(-1, 0).Interior.ColorIndex = 36
                .Cells(lRow, Col_GU_Dups_GU_Duns).Interior.ColorIndex = 36
            End If
        Next
    
        'Remove all singular Global Ultimate DUNS
        For lRow = 2 To .Rows.Count
            If Not .Cells(lRow, Col_GU_Dups_GU_Duns).Interior.ColorIndex = 36 Then
                .Cells(lRow, Col_GU_Dups_GU_Duns).EntireRow.Delete
                If lRow > .Rows.Count Then
                    Exit For
                Else
                    lRow = lRow - 1
                End If
            End If
        Next
    
        'Apply formating to Global Ultimate duplicates
        vColorIndex = 34
        For lRow = 2 To .Rows.Count
            .Cells(lRow, Col_GU_Dups_GU_Duns).Interior.ColorIndex = xlColorIndexNone
            Range(.Cells(lRow, Col_GU_Dups_Bus_Nme), .Cells(lRow, Col_GU_Dups_GU_Ctry_Nme)).Interior.ColorIndex = vColorIndex
            If Not (.Cells(lRow, Col_GU_Dups_GU_Duns).Offset(1, 0).Value = .Cells(lRow, Col_GU_Dups_GU_Duns).Value) Then
                If vColorIndex = 34 Then
                    vColorIndex = 37
                Else
                    vColorIndex = 34
                End If
            End If
        Next
    End With
End Sub


'Hide empty pivot items
Private Sub HideEmptyPivotItems(pFld As PivotField)
    On Error GoTo ErrHandler
    
    'Maybe there are no empty items
    pFld.PivotItems("").Visible = False

ErrHandler:
End Sub


'Create standard reports based on the D&B data-append
Private Sub CreateStandardReports()
    Dim wsReport As Worksheet
    Dim pcDnbData As PivotCache
    Dim ptReport As PivotTable
    Dim sFLD1 As String, sFLD2 As String, sFLD3 As String
    Dim lCol As Long
    
    'Create standard reports
    With ThisWorkbook
        'Create the data source for the pivot tables
        Set pcDnbData = .PivotCaches.Add(SourceType:=xlDatabase, SourceData:=sXlNameDnbData)
            
        'Create a worksheet for the country distribution report
        If WsExists(sCtryDistr) Then .Worksheets(sCtryDistr).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sCtryDistr
        
        'Create the country distribution pivot
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_Cont_Desc).Value
        sFLD2 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_Ctry_Nme).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sCtryDistr & "'" & "!R3C1", TableName:=sPivotCtryDistr, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD2), "Country count", xlCount
        ptReport.AddFields RowFields:=Array(sFLD1, sFLD2)
        wsReport.UsedRange.Columns.AutoFit
        
        'Create a worksheet for the DUNS v. global ultimate country distribution pivot
        If WsExists(sDunsGblUltCtryDistr) Then .Worksheets(sDunsGblUltCtryDistr).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sDunsGblUltCtryDistr
        
        'Create the DUNS v. global ultimate country distribution pivot
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_ISO_Ctry_Cd).Value
        sFLD2 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_Gbl_Ult_ISO_Ctry_Cd).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sDunsGblUltCtryDistr & "'" & "!R3C1", TableName:=sPivotDunsGblUltCtryDistr, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD2), "Count global ult. country ", xlCount
        ptReport.AddFields RowFields:=sFLD1, ColumnFields:=sFLD2
        HideEmptyPivotItems ptReport.PivotFields(sFLD2)
        wsReport.UsedRange.Columns.AutoFit
        For lCol = 2 To wsReport.UsedRange.Columns.Count - 1
            wsReport.UsedRange.Columns(lCol).ColumnWidth = Col_Width_6
        Next
        
        'Create a worksheet for the SIC distribution reports
        If WsExists(sActCodeDistr) Then .Worksheets(sActCodeDistr).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sActCodeDistr
        
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_US1987_Sic1_1Desc).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sActCodeDistr & "'" & "!R3C1", TableName:=sPivotActCodeDistr, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD1), "Count of Top Level SIC", xlCount
        ptReport.AddFields RowFields:=sFLD1
        ptReport.PivotFields(sFLD1).AutoSort xlDescending, "Count of Top Level SIC"
        HideEmptyPivotItems ptReport.PivotFields(sFLD1)
        wsReport.UsedRange.Columns.AutoFit
    
        If WsExists(sActCodeTop10) Then .Worksheets(sActCodeTop10).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sActCodeTop10
        
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_US1987_Sic1_Desc).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sActCodeTop10 & "'" & "!R3C1", TableName:=sPivotActCodeTop10, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD1), "SIC Top 10", xlCount
        ptReport.AddFields RowFields:=sFLD1
        ptReport.PivotFields(sFLD1).AutoSort xlDescending, "SIC Top 10"
        ptReport.PivotFields(sFLD1).AutoShow xlAutomatic, xlTop, 10, "SIC Top 10"
        HideEmptyPivotItems ptReport.PivotFields(sFLD1)
        ptReport.ColumnGrand = False
        wsReport.UsedRange.Columns.AutoFit
    
        'Create a worksheet for the annual sales distribution reports
        If WsExists(sAnnSalesDistr) Then .Worksheets(sAnnSalesDistr).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sAnnSalesDistr
        
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_Ann_Sales_US_Distr).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sAnnSalesDistr & "'" & "!R3C1", TableName:=sPivotAnnSalesDistr, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD1), "Count Annual Sales Distribution", xlCount
        ptReport.AddFields RowFields:=sFLD1
        HideEmptyPivotItems ptReport.PivotFields(sFLD1)
        wsReport.UsedRange.Columns.AutoFit
        
        'Create a worksheet for the start year distribution reports
        If WsExists(sStartYearDistr) Then .Worksheets(sStartYearDistr).Delete
        Set wsReport = .Worksheets.Add(After:=ActiveSheet)
        wsReport.Name = sStartYearDistr
        
        sFLD1 = .Names(sXlNameDnbData).RefersToRange.Cells(1, Col_Xl_Yr_Strt_Distr).Value
        Set ptReport = pcDnbData.CreatePivotTable(TableDestination:="'" & sStartYearDistr & "'" & "!R3C1", TableName:=sPivotStartYearDistr, DefaultVersion:=xlPivotTableVersion10)
        ptReport.AddDataField ptReport.PivotFields(sFLD1), "Count Start Year Distribution", xlCount
        ptReport.AddFields RowFields:=sFLD1
        HideEmptyPivotItems ptReport.PivotFields(sFLD1)
        wsReport.UsedRange.Columns.AutoFit
    End With
End Sub


'Initialization upon opening this workbook
Public Sub Auto_Open()
    On Error GoTo ErrHandler
    
    'Dimension the variables to be used while reading the files
    Dim bProcessLhv1784Xl As Boolean: bProcessLhv1784Xl = True
    Dim bRemoveRawData As Boolean: bRemoveRawData = False
    Dim bCreateDuplicateReports As Boolean: bCreateDuplicateReports = True
    Dim bCreateStandardReports As Boolean: bCreateStandardReports = True
    Dim sFileOpenPath As String
    Dim arrFldInfo As Variant
    Dim wsLhv1784Raw As Worksheet
    Dim wsLhv1784Xl As Worksheet
    Dim wsLayout1784 As Worksheet
    Dim lRow As Long: lRow = 2
    Dim lCol As Long: lCol = 1
    
    'Don't execute auto_open if data is already available in the workbook
    If WsExists(sLhv1784Raw) Then Err.Raise Number:=errWsRawExists
    
    'Let the user choose the 1784 file
    With Application.FileDialog(msoFileDialogFilePicker)
        .AllowMultiSelect = False
        .Title = "Select 1784 text file"
        .Show

        If .SelectedItems.Count = 1 Then
            sFileOpenPath = .SelectedItems(1)
        Else
            Err.Raise Number:=errNoFileSpecified
        End If
    End With
        
    'Specify the 1784 file format and open the fixed width text file
    arrFldInfo = Array( _
        Array(0, 2), Array(2, 2), Array(11, 2), Array(101, 2), Array(191, 2), Array(192, 2), Array(256, 2), Array(320, 2), Array(350, 2), Array(380, 2), Array(400, 2), Array(406, 2), Array(409, 2), Array(412, 2), Array(416, 2), Array(419, 2), Array(428, 2), Array(429, 2), Array(461, 2), Array(491, 2), _
        Array(521, 2), Array(551, 2), Array(571, 2), Array(577, 2), Array(580, 2), Array(583, 2), Array(587, 2), Array(590, 2), Array(599, 2), Array(600, 2), Array(616, 2), Array(621, 2), Array(625, 2), Array(641, 2), Array(657, 2), Array(673, 2), Array(733, 2), Array(793, 2), Array(834, 2), Array(838, 2), _
        Array(842, 2), Array(846, 2), Array(850, 2), Array(854, 2), Array(858, 2), Array(866, 2), Array(869, 2), Array(873, 2), Array(891, 2), Array(892, 2), Array(907, 2), Array(911, 2), Array(918, 2), Array(919, 2), Array(926, 2), Array(927, 2), Array(928, 2), Array(929, 2), Array(932, 2), Array(933, 2), _
        Array(934, 2), Array(935, 2), Array(937, 2), Array(946, 2), Array(954, 2), Array(956, 2), Array(965, 2), Array(1055, 2), Array(1119, 2), Array(1149, 2), Array(1179, 2), Array(1199, 2), Array(1205, 2), Array(1208, 2), Array(1212, 2), Array(1215, 2), Array(1224, 2), Array(1225, 2), Array(1227, 2), _
        Array(1236, 2), Array(1326, 2), Array(1390, 2), Array(1420, 2), Array(1450, 2), Array(1456, 2), Array(1459, 2), Array(1463, 2), Array(1472, 2), Array(1473, 2), Array(1475, 2), Array(1484, 2), Array(1574, 2), Array(1638, 2), Array(1668, 2), Array(1698, 2), Array(1718, 2), Array(1724, 2), _
        Array(1727, 2), Array(1731, 2), Array(1734, 2), Array(1743, 2), Array(1744, 2), Array(1749, 2), Array(1758, 2), Array(1760, 2), Array(1768, 2), Array(1769, 2), Array(1770, 2), Array(1780, 2))
        
    Application.ScreenUpdating = False 'No screen updates at this stage
    Application.Calculation = xlManual 'This for speeding things up
        
    Workbooks.OpenText Filename:=sFileOpenPath, Origin:=1252, StartRow:=1, DataType:=xlFixedWidth, FieldInfo:=arrFldInfo
        
    'Move the raw data into this workbook
    With ActiveSheet
        .Name = sLhv1784Raw
        .Move before:=ThisWorkbook.Sheets(1)
    End With
        
    Set wsLhv1784Raw = ThisWorkbook.Sheets(sLhv1784Raw)
    Set wsLayout1784 = ThisWorkbook.Sheets(sLayout1784)

    'Add a header row
    With wsLhv1784Raw
        .Rows(1).Insert Shift:=xlDown
        .Rows(1).RowHeight = 25.5
        Do While Not IsEmpty(wsLayout1784.Cells(lRow, Col_Layout_DESC))
            .Cells(1, lCol).Value = wsLayout1784.Cells(lRow, Col_Layout_DESC)
            FormatHdr .Cells(1, lCol)
            lRow = lRow + 1
            lCol = lCol + 1
        Loop
    End With
        
    'Format the view
    wsLhv1784Raw.Activate
    wsLhv1784Raw.Cells(2, 1).Select
    ActiveWindow.FreezePanes = True
    ActiveWindow.Zoom = 80
    
    'Process the raw data for optimal use in Excel
    If bProcessLhv1784Xl Then
        'Add a worksheet with 1784 data formatted for use in Excel
        Set wsLhv1784Xl = ThisWorkbook.Worksheets.Add(After:=wsLhv1784Raw)
        With wsLhv1784Xl
            .Name = sLhv1784Xl
            .Cells(2, 5).Select
            ActiveWindow.FreezePanes = True
            ActiveWindow.Zoom = 80
            .Rows(1).RowHeight = 25.5
        End With
        
        'Most of the actual processing is done in sub ProcessLhv1784Xl
        ProcessLhv1784Xl wsLhv1784Raw, wsLhv1784Xl
        
        wsLhv1784Xl.Calculate 'just this once
    End If

    'Create the duplicate (DUNS & GU) reports
    If bCreateDuplicateReports Then CreateDuplicateReports wsLhv1784Raw
    
    'Create a few example reports based on the D&B data-append
    If bCreateStandardReports Then CreateStandardReports
    
    'If applicable remove the raw data without user prompt
    If bRemoveRawData Then
        Application.DisplayAlerts = False
        wsLhv1784Raw.Delete
        Application.DisplayAlerts = True
    End If

ErrHandler:
    If Err.Number > 0 Then Debug.Print sErrMsg(Err.Number)
   
    Application.ScreenUpdating = True
End Sub


