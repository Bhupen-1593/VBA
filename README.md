# VBA
'My VBA Codes

Option Explicit
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open()
On Error Resume Next
Call Data_Verification                                                                                                                               '_here data is first verified by counting the rows for every WTG to be equal                                                                                                                                                                 if they are not equal then then excel wont allow us to progress further_
End Sub
Private Sub Data_Verification()                                                                                                                       '_this is the sub-routine called at line number 8_
Dim ans1 As Variant                                                                                                                                   '_this is variable number 1_
Dim ans2 As Variant '_this is variable number 2_
Sheets("Generation").Select                                                                                                                           '_this line of code selects the excel sheet named 'Generation'_
If Range("AB1").Value = Range("AC1").Value And Range("AB1").Value = Range("AD1").Value _
And Range("AB1").Value = Range("AE1").Value And Range("AB1").Value = Range("AC1").Value Then
ans1 = MsgBox("Data is verified you can proceed", vbOKOnly)                                                                                           '_code lines 14 to 16 checks the number of rows for every individual WTG                                                                                                                                                                      for program to progress further rows must be equal in number for the code to proceed_
On Error Resume Next                                                                                                                                  '_if rows are not equal then this line reurns error and line 21 is executed skipping                                                                                                                                                         line 18 and 19 else Workbook_Open_Copy1 and Workbook_Open_Copy2 sub-routines are run
Call Workbook_Open_Copy1
Call Workbook_Open_Copy2
Else
ans2 = MsgBox("No. of rows of different WTGs are not equal excel will not proceed", vbOKOnly)
End If
End Sub
---------------------------------------------------------FIRST PART ENDS-----------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open_Copy1()                                                                                                                                       '_this sub-routine will open downloaded file and copy the required                                                                                                                                                                             contents to the analysis file_
Dim wsdaily1 As Workbook                                                                                                                                                 '_Variable 1_
Dim answer1 As Variant                                                                                                                                                    '_Variable 2_

answer1 = MsgBox("Open and copy GeeCee Generation as on" & VBA.Format(VBA.Date - 1, "dd-mm-yy"), vbYesNo)                                                            '_messeage box with yes and no option is created and stored in variable                                                                                                                                                                       called answer1_
If answer1 = vbYes Then

Set wsdaily1 = Workbooks.Open("D:/BHUPEN PCC/client data/CRMS Daily/Gee Cee CRMS/CRMS Daily Download/gc" _                                                               '_the workbook data type has to be set, here the path of downloaded                                                                                                                                                                 file having WTG data is stored and opened inside the variable named wsdaily1_
& VBA.Format(VBA.Date - 1, "dd-mm-yy"))

If Range("A2").Value <> "" And Range("A3").Value <> "" Then                                                                                                    _'code lines from 35 to 43 checks the downloaded files content and if there                                                                                                                                                                      is required content available then only lines 36 to 38 are excuted where                                                                                                                                                                     contents are copied to analysis workbook else lines 40 to 42 are excuted                                                                                                                                                                     where the same things copy to analysis workbook happens_
    wsdaily1.Sheets(1).Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy
    ThisWorkbook.Sheets(1).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
Else
    wsdaily1.Sheets(1).Range("A2", Range("A2").End(xlToRight)).Copy
    ThisWorkbook.Sheets(1).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
End If

wsdaily1.Sheets(2).Activate                                                                                                                                     _'code lines 45 to 54 copies the contents of sheets 2nd of the downloaded                                                                                                                                                                       file and repeats the same execution of proces as process described above_
If Range("A2").Value <> "" And Range("A3").Value <> "" Then
    wsdaily1.Sheets(2).Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy
    ThisWorkbook.Sheets(2).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
Else
    wsdaily1.Sheets(2).Range("A2", Range("A2").End(xlToRight)).Copy
    ThisWorkbook.Sheets(2).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
End If

End If
End Sub
-------------------------------------------------------------------------------SECOND PART ENDS------------------------------------------------------------------------------------------------------------------
Private Sub Workbook_Open_Copy2()                                                                                                    ' _this part is similar to second part only difference is new WTG file is considered to execute the VBA                                                                                                                                         code_
Dim wsdaily2 As Workbook
Dim answer2 As Variant

answer2 = MsgBox("Open and copy Savla Generation as on" & VBA.Format(VBA.Date - 1, "dd-mm-yy"), vbYesNo)
If answer2 = vbYes Then

Set wsdaily2 = Workbooks.Open("D:/BHUPEN PCC/client data/CRMS Daily/savla/CRMS download daily/savt" _
& VBA.Format(VBA.Date - 1, "dd-mm-yy"))

If Range("A2").Value <> "" And Range("A3").Value <> "" Then
    wsdaily2.Sheets(1).Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy
    ThisWorkbook.Sheets(1).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
Else
    wsdaily2.Sheets(1).Range("A2", Range("A2").End(xlToRight)).Copy
    ThisWorkbook.Sheets(1).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
End If

wsdaily2.Sheets(2).Activate
If Range("A2").Value <> "" And Range("A3").Value <> "" Then
    wsdaily2.Sheets(2).Range("A2", Range("A2").End(xlToRight).End(xlDown)).Copy
    ThisWorkbook.Sheets(2).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
Else
    wsdaily2.Sheets(2).Range("A2", Range("A2").End(xlToRight)).Copy
    ThisWorkbook.Sheets(2).Activate
    Range("A2").End(xlDown).Offset(1, 0).PasteSpecial
End If

End If
End Sub




