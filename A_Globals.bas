Attribute VB_Name = "A_Globals"
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

'Worksheet names
Public Const sLhv1784Raw = "Lhv1784Raw" 'Raw D&B 1784 data
Public Const sLhv1784Xl = "Lhv1784Xl" 'D&B 1784 data formatted for Excel
Public Const sDunsDupsReport = "DunsDuplReport" 'D&B DUNS duplicate report
Public Const sGblUltDupsReport = "GlobalUltDuplReport" 'D&B global ultimate duplicate report
Public Const sCtryDistr = "D&B Country Distr" 'Country distribution report
Public Const sDunsGblUltCtryDistr = "D&B Duns v. Gbl Ult Ctry Distr" 'Country distribution report
Public Const sActCodeDistr = "D&B Act. Code Distr" 'Top level SIC distribution report
Public Const sActCodeTop10 = "D&B Act. Code Top 10" 'SIC top 10 report
Public Const sAnnSalesDistr = "D&B Annual Sales Distr" 'Annual sales distribution report
Public Const sStartYearDistr = "D&B Start Year Distr" 'Start year distribution report
Public Const sLayout1784 = "Layout1784" 'Lhv1784 specification

'Name of the append area
Public Const sXlNameDnbData = "DnbData" 'Name associated with the range containing the D&B data append
Public Const sXlNameDnbDunsDups = "DnbDunsDups" 'Name associated with the range containing the D&B DUNS duplicates
Public Const sXlNameDnbGblUltDups = "DnbGblUltDups" 'Name associated with the range containing the D&B global ultimate duplicates

'Names of pivot tables
Public Const sPivotCtryDistr = "CountryDistr" 'Country distribution report
Public Const sPivotDunsGblUltCtryDistr = "DunsGblUltCtryDistr" 'Country distribution report
Public Const sPivotActCodeDistr = "ActCodeDistr" 'Top level SIC distribution report
Public Const sPivotActCodeTop10 = "ActCodeTop10" 'SIC top 10 report
Public Const sPivotAnnSalesDistr = "AnnSalesDistr" 'Annual sales distribution report
Public Const sPivotStartYearDistr = "StartYearDistr" 'Start year distribution report

'Application error codes
Public Const errWsRawExists = 32767
Public Const errNoFileSpecified = 32768

'Columns Lhv1784Raw
Public Const Col_Raw_System_ID = 109
Public Const Col_Raw_Sap_Cust_Num = 108
Public Const Col_Raw_Filler1 = 1
Public Const Col_Raw_Duns = 2
Public Const Col_Raw_Bus_Nme = 3
Public Const Col_Raw_Trdg_style_1 = 4
Public Const Col_Raw_Reg_Adr_ind = 5
Public Const Col_Raw_Adr1 = 6
Public Const Col_Raw_Adr2 = 7
Public Const Col_Raw_Post_Cd = 16
Public Const Col_Raw_City_Nme = 8
Public Const Col_Raw_State_Prov = 9
Public Const Col_Raw_Ctry_Nme = 10
Public Const Col_Raw_City_Cd = 11
Public Const Col_Raw_Cnty_Cd = 12
Public Const Col_Raw_State_Prov_Cd = 13
Public Const Col_Raw_State_Prov_Abbr = 14
Public Const Col_Raw_Ctry_Cd = 15
Public Const Col_Raw_Cont_Cd = 17
Public Const Col_Raw_MAdr = 18
Public Const Col_Raw_MPost_Cd = 28
Public Const Col_Raw_MCity_Nme = 19
Public Const Col_Raw_MCnty_Nme = 20
Public Const Col_Raw_MState_Prov_Nme = 21
Public Const Col_Raw_MCtry_Nme = 22
Public Const Col_Raw_MCity_Cd = 23
Public Const Col_Raw_MCnty_Cd = 24
Public Const Col_Raw_MState_Prov_Cd = 25
Public Const Col_Raw_MState_Prov_Abbr = 26
Public Const Col_Raw_MCtry_Cd = 27
Public Const Col_Raw_MCont_Cd = 29
Public Const Col_Raw_Ctry_Access_Cd = 32
Public Const Col_Raw_Tel = 33
Public Const Col_Raw_Cable_Telx = 34
Public Const Col_Raw_Fax = 35
Public Const Col_Raw_NatID = 30
Public Const Col_Raw_NatID_Cd = 31
Public Const Col_Raw_OOB_Ind = 106
Public Const Col_Raw_Marketable_Ind = 107
Public Const Col_Raw_Yr_Strt = 47
Public Const Col_Raw_Ann_Sales = 48
Public Const Col_Raw_Crcy_Cd = 51
Public Const Col_Raw_Ann_Sales_Ind = 49
Public Const Col_Raw_Ann_Sales_US = 50
Public Const Col_Raw_Emp_Here = 52
Public Const Col_Raw_Emp_Here_Ind = 53
Public Const Col_Raw_Emp_Tot = 54
Public Const Col_Raw_Emp_Tot_Ind = 55
Public Const Col_Raw_Prin_Incl_Ind = 56
Public Const Col_Raw_Imp_Exp_Cd = 57
Public Const Col_Raw_Lgl_Stat = 58
Public Const Col_Raw_Control_Ind = 59
Public Const Col_Raw_Stat_Cd = 60
Public Const Col_Raw_Subs_Cd = 61
Public Const Col_Raw_CEO_Nme = 36
Public Const Col_Raw_CEO_Title = 37
Public Const Col_Raw_Rep_Date = 64
Public Const Col_Raw_Filler2 = 62
Public Const Col_Raw_Prev_Duns = 63
Public Const Col_Raw_US1987_Sic1 = 39
Public Const Col_Raw_Line_Bus_Descr = 38
Public Const Col_Raw_US1987_Sic2 = 40
Public Const Col_Raw_US1987_Sic3 = 41
Public Const Col_Raw_US1987_Sic4 = 42
Public Const Col_Raw_US1987_Sic5 = 43
Public Const Col_Raw_US1987_Sic6 = 44
Public Const Col_Raw_Prmy_Loc_Act_Cd = 45
Public Const Col_Raw_Act_Ind = 46
Public Const Col_Raw_HQ_Prnt_Fil = 65
Public Const Col_Raw_HQ_Prnt_Duns = 66
Public Const Col_Raw_HQ_Prnt_Nme = 67
Public Const Col_Raw_HQ_Prnt_Adr = 68
Public Const Col_Raw_HQ_Prnt_Post_Cd = 76
Public Const Col_Raw_HQ_Prnt_City_Nme = 69
Public Const Col_Raw_HQ_Prnt_State_Prov_Nme = 70
Public Const Col_Raw_HQ_Prnt_Ctry_Nme = 71
Public Const Col_Raw_HQ_Prnt_City_Cd = 72
Public Const Col_Raw_HQ_Prnt_Cnty_Cd = 73
Public Const Col_Raw_HQ_Prnt_State_Prov_Abbr = 74
Public Const Col_Raw_HQ_Prnt_Ctry_Cd = 75
Public Const Col_Raw_HQ_Prnt_Cont_Cd = 77
Public Const Col_Raw_Dom_Ult_Fil = 78
Public Const Col_Raw_Dom_Ult_Duns = 79
Public Const Col_Raw_Dom_Ult_Nme = 80
Public Const Col_Raw_Dom_Ult_Adr = 81
Public Const Col_Raw_Dom_Ult_Post_Cd = 87
Public Const Col_Raw_Dom_Ult_City_Nme = 82
Public Const Col_Raw_Dom_Ult_State_Prov_Nme = 83
Public Const Col_Raw_Dom_Ult_City_Cd = 84
Public Const Col_Raw_Dom_Ult_State_Prov_Abbr = 86
Public Const Col_Raw_Dom_Ult_Ctry_Cd = 85
Public Const Col_Raw_Gbl_Ult_Ind = 88
Public Const Col_Raw_Gbl_Ult_Fil = 89
Public Const Col_Raw_Gbl_Ult_Duns = 90
Public Const Col_Raw_Gbl_Ult_Nme = 91
Public Const Col_Raw_Gbl_Ult_Adr = 92
Public Const Col_Raw_Gbl_Ult_Post_Cd = 100
Public Const Col_Raw_Gbl_Ult_City_Nme = 93
Public Const Col_Raw_Gbl_Ult_State_Prov_Nme = 94
Public Const Col_Raw_Gbl_Ult_Ctry_Nme = 95
Public Const Col_Raw_Gbl_Ult_City_Cd = 96
Public Const Col_Raw_Gbl_Ult_Cnty_Cd = 97
Public Const Col_Raw_Gbl_Ult_State_Prov_Abbr = 98
Public Const Col_Raw_Gbl_Ult_Ctry_Cd = 99
Public Const Col_Raw_Gbl_Ult_Cont_Cd = 101
Public Const Col_Raw_Fam_Memb_Gbl = 102
Public Const Col_Raw_Gbl_Dias_Cd = 103
Public Const Col_Raw_Gbl_Hier_Cd = 104
Public Const Col_Raw_Fam_Upd_Date = 105

'Columns Lhv1784Xl
Public Const Col_Xl_System_ID = 1
Public Const Col_Xl_Sap_Cust_Num = 2
Public Const Col_Xl_Filler1 = 3
Public Const Col_Xl_Duns = 4
Public Const Col_Xl_Bus_Nme = 5
Public Const Col_Xl_Trdg_style_1 = 6
Public Const Col_Xl_Reg_Adr_ind = 7
Public Const Col_Xl_Adr1 = 8
Public Const Col_Xl_Adr2 = 9
Public Const Col_Xl_Post_Cd = 10
Public Const Col_Xl_City_Nme = 11
Public Const Col_Xl_State_Prov = 12
Public Const Col_Xl_Ctry_Nme = 13
Public Const Col_Xl_City_Cd = 14
Public Const Col_Xl_Cnty_Cd = 15
Public Const Col_Xl_State_Prov_Cd = 16
Public Const Col_Xl_State_Prov_Abbr = 17
Public Const Col_Xl_Ctry_Cd = 18
Public Const Col_Xl_ISO_Ctry_Cd = 19
Public Const Col_Xl_Cont_Cd = 20
Public Const Col_Xl_Cont_Desc = 21
Public Const Col_Xl_MAdr = 22
Public Const Col_Xl_MPost_Cd = 23
Public Const Col_Xl_MCity_Nme = 24
Public Const Col_Xl_MCnty_Nme = 25
Public Const Col_Xl_MState_Prov_Nme = 26
Public Const Col_Xl_MCtry_Nme = 27
Public Const Col_Xl_MCity_Cd = 28
Public Const Col_Xl_MCnty_Cd = 29
Public Const Col_Xl_MState_Prov_Cd = 30
Public Const Col_Xl_MState_Prov_Abbr = 31
Public Const Col_Xl_MCtry_Cd = 32
Public Const Col_Xl_ISO_MCtry_Cd = 33
Public Const Col_Xl_MCont_Cd = 34
Public Const Col_Xl_Ctry_Access_Cd = 35
Public Const Col_Xl_Tel = 36
Public Const Col_Xl_Cable_Telx = 37
Public Const Col_Xl_Fax = 38
Public Const Col_Xl_NatID = 39
Public Const Col_Xl_NatID_Cd = 40
Public Const Col_Xl_NatID_Desc = 41
Public Const Col_Xl_OOB_Ind = 42
Public Const Col_Xl_Marketable_Ind = 43
Public Const Col_Xl_Yr_Strt = 44
Public Const Col_Xl_Yr_Strt_Int = 45
Public Const Col_Xl_Yr_Strt_Distr = 46
Public Const Col_Xl_Ann_Sales = 47
Public Const Col_Xl_Crcy_Cd = 48
Public Const Col_Xl_ISO_Crcy_Cd = 49
Public Const Col_Xl_Ann_Sales_Formatted = 50
Public Const Col_Xl_Ann_Sales_Ind = 51
Public Const Col_Xl_Ann_Sales_Desc = 52
Public Const Col_Xl_Ann_Sales_US = 53
Public Const Col_Xl_Ann_Sales_US_Formatted = 54
Public Const Col_Xl_Ann_Sales_US_Distr = 55
Public Const Col_Xl_Emp_Here = 56
Public Const Col_Xl_Emp_Here_Int = 57
Public Const Col_Xl_Emp_Here_Ind = 58
Public Const Col_Xl_Emp_Here_Desc = 59
Public Const Col_Xl_Emp_Tot = 60
Public Const Col_Xl_Emp_Tot_Int = 61
Public Const Col_Xl_Emp_Tot_Ind = 62
Public Const Col_Xl_Emp_Tot_Desc = 63
Public Const Col_Xl_Prin_Incl_Ind = 64
Public Const Col_Xl_Imp_Exp_Cd = 65
Public Const Col_Xl_Imp_Exp_Desc = 66
Public Const Col_Xl_Lgl_Stat = 67
Public Const Col_Xl_Lgl_Stat_Desc = 68
Public Const Col_Xl_Control_Ind = 69
Public Const Col_Xl_Stat_Cd = 70
Public Const Col_Xl_Stat_Desc = 71
Public Const Col_Xl_Subs_Cd = 72
Public Const Col_Xl_Subs_Desc = 73
Public Const Col_Xl_CEO_Nme = 74
Public Const Col_Xl_CEO_Title = 75
Public Const Col_Xl_Rep_Date = 76
Public Const Col_Xl_Rep_Date_Date = 77
Public Const Col_Xl_Filler2 = 78
Public Const Col_Xl_Prev_Duns = 79
Public Const Col_Xl_US1987_Sic1 = 80
Public Const Col_Xl_Line_Bus_Descr = 81
Public Const Col_Xl_US1987_Sic1_Desc = 82
Public Const Col_Xl_US1987_Sic1_2 = 83
Public Const Col_Xl_US1987_Sic1_2Desc = 84
Public Const Col_Xl_US1987_Sic1_1 = 85
Public Const Col_Xl_US1987_Sic1_1Desc = 86
Public Const Col_Xl_US1987_Sic2 = 87
Public Const Col_Xl_US1987_Sic3 = 88
Public Const Col_Xl_US1987_Sic4 = 89
Public Const Col_Xl_US1987_Sic5 = 90
Public Const Col_Xl_US1987_Sic6 = 91
Public Const Col_Xl_Prmy_Loc_Act_Cd = 92
Public Const Col_Xl_Act_Ind = 93
Public Const Col_Xl_Act_Desc = 94
Public Const Col_Xl_HQ_Prnt_Fil = 95
Public Const Col_Xl_HQ_Prnt_Duns = 96
Public Const Col_Xl_HQ_Prnt_Nme = 97
Public Const Col_Xl_HQ_Prnt_Adr = 98
Public Const Col_Xl_HQ_Prnt_Post_Cd = 99
Public Const Col_Xl_HQ_Prnt_City_Nme = 100
Public Const Col_Xl_HQ_Prnt_State_Prov_Nme = 101
Public Const Col_Xl_HQ_Prnt_Ctry_Nme = 102
Public Const Col_Xl_HQ_Prnt_City_Cd = 103
Public Const Col_Xl_HQ_Prnt_Cnty_Cd = 104
Public Const Col_Xl_HQ_Prnt_State_Prov_Abbr = 105
Public Const Col_Xl_HQ_Prnt_Ctry_Cd = 106
Public Const Col_Xl_HQ_Prnt_ISO_Ctry_Cd = 107
Public Const Col_Xl_HQ_Prnt_Cont_Cd = 108
Public Const Col_Xl_Dom_Ult_Fil = 109
Public Const Col_Xl_Dom_Ult_Duns = 110
Public Const Col_Xl_Dom_Ult_Nme = 111
Public Const Col_Xl_Dom_Ult_Adr = 112
Public Const Col_Xl_Dom_Ult_Post_Cd = 113
Public Const Col_Xl_Dom_Ult_City_Nme = 114
Public Const Col_Xl_Dom_Ult_State_Prov_Nme = 115
Public Const Col_Xl_Dom_Ult_City_Cd = 116
Public Const Col_Xl_Dom_Ult_State_Prov_Abbr = 117
Public Const Col_Xl_Dom_Ult_Ctry_Cd = 118
Public Const Col_Xl_Dom_Ult_ISO_Ctry_Cd = 119
Public Const Col_Xl_Gbl_Ult_Ind = 120
Public Const Col_Xl_Gbl_Ult_Fil = 121
Public Const Col_Xl_Gbl_Ult_Duns = 122
Public Const Col_Xl_Gbl_Ult_Nme = 123
Public Const Col_Xl_Gbl_Ult_Adr = 124
Public Const Col_Xl_Gbl_Ult_Post_Cd = 125
Public Const Col_Xl_Gbl_Ult_City_Nme = 126
Public Const Col_Xl_Gbl_Ult_State_Prov_Nme = 127
Public Const Col_Xl_Gbl_Ult_Ctry_Nme = 128
Public Const Col_Xl_Gbl_Ult_City_Cd = 129
Public Const Col_Xl_Gbl_Ult_Cnty_Cd = 130
Public Const Col_Xl_Gbl_Ult_State_Prov_Abbr = 131
Public Const Col_Xl_Gbl_Ult_Ctry_Cd = 132
Public Const Col_Xl_Gbl_Ult_ISO_Ctry_Cd = 133
Public Const Col_Xl_Gbl_Ult_Cont_Cd = 134
Public Const Col_Xl_Gbl_Ult_Cont_Desc = 135
Public Const Col_Xl_Fam_Memb_Gbl = 136
Public Const Col_Xl_Fam_Memb_Gbl_Int = 137
Public Const Col_Xl_Gbl_Dias_Cd = 138
Public Const Col_Xl_Gbl_Hier_Cd = 139
Public Const Col_Xl_Fam_Upd_Date = 140

'Columns Layout1784
Public Const Col_Layout_ID = 1
Public Const Col_Layout_POS = 2
Public Const Col_Layout_NAME = 3
Public Const Col_Layout_DESC = 4
Public Const Col_Layout_START_POS = 5
Public Const Col_Layout_WIDTH = 6

'Column widths
Public Const Col_Width_6 = 6
Public Const Col_Width_8 = 8
Public Const Col_Width_12 = 12
Public Const Col_Width_18 = 18
Public Const Col_Width_DUNS = 11.14
Public Const Col_Width_Name1 = 40
Public Const Col_Width_Name2 = 25
Public Const Col_Width_Adr1 = 30
Public Const Col_Width_Adr2 = 25
Public Const Col_Width_MAdr = 15
Public Const Col_Width_Post_Cd = 9
Public Const Col_Width_City_Nme = 19
Public Const Col_Width_Ctry_Nme = 20
Public Const Col_Width_NatID = 16
Public Const Col_Width_Desc_Short = 15
Public Const Col_Width_Desc_Med = 25
Public Const Col_Width_Desc = 40
Public Const Col_Width_Tel = 12.5
Public Const Col_Width_Year = 6
Public Const Col_Width_SIC = 7
Public Const Col_Width_Str_Date = 9
Public Const Col_Width_Date_Date = 12
Public Const Col_Width_Small_Str_Num = 9
Public Const Col_Width_Small_Int_Num = 8
Public Const Col_Width_Large_Str_Num = 20
Public Const Col_Width_Large_Curr_Num = 19
Public Const Col_Width_CEO_Name = 25
Public Const Col_Width_CEO_Title = 30


'Check if a worksheet exists
Public Function WsExists(sWsName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sDummy As String
    sDummy = ThisWorkbook.Worksheets(sWsName).Name
    
    WsExists = True
    Exit Function
    
ErrHandler:
    WsExists = False
End Function


'Check if an Excel name exists
Public Function XlNameExists(sXlName As String) As Boolean
    On Error GoTo ErrHandler
    
    Dim sDummy As String
    sDummy = ThisWorkbook.Names(sXlName).Name
    
    XlNameExists = True
    Exit Function
    
ErrHandler:
    XlNameExists = False
End Function


'Return an English error message
Public Function sErrMsg(lErrNum As Long) As String
    Select Case lErrNum
    Case errNoFileSpecified
        sErrMsg = "No file specified"
    Case Else
        sErrMsg = Err.Description
    End Select
End Function


'Remove the generated sheets & names before save
Private Sub CleanUpWb()
    On Error Resume Next
    
    Application.DisplayAlerts = False
    
    ThisWorkbook.Worksheets(sLhv1784Raw).Delete
    ThisWorkbook.Worksheets(sLhv1784Xl).Delete
    ThisWorkbook.Worksheets(sDunsDupsReport).Delete
    ThisWorkbook.Worksheets(sGblUltDupsReport).Delete
    ThisWorkbook.Worksheets(sCtryDistr).Delete
    ThisWorkbook.Worksheets(sDunsGblUltCtryDistr).Delete
    ThisWorkbook.Worksheets(sActCodeDistr).Delete
    ThisWorkbook.Worksheets(sActCodeTop10).Delete
    ThisWorkbook.Worksheets(sAnnSalesDistr).Delete
    ThisWorkbook.Worksheets(sStartYearDistr).Delete
    
    ThisWorkbook.Names(sXlNameDnbData).Delete
    ThisWorkbook.Names(sXlNameDnbDunsDups).Delete
    ThisWorkbook.Names(sXlNameDnbGblUltDups).Delete

    Application.DisplayAlerts = True
End Sub

