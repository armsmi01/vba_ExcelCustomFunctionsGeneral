'title: vba_ExcelCustomFunctionsGeneral
'note: insert into stand-alone module, do not insert into existing modules or objects
'
'author: Michael J. Armstrong
'contact: marmstrong310@gmail.com
'© 2017, 2018 Michael J. Armstrong, Toronto, Ontario, Canada
'Distributed under the terms of the GNU General Public License v3.0
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.

Function OptimizeOn()
'Disable all process items to optimize VBA code
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

End Function

Function OptimizeOff()
'Disable all process items to optimize VBA code
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True

End Function

Function CounterResults(count As Double, time As Double) As String
'Add in MsgBox() as argument, returns results of count
    If count > 1 Then
        CounterResults = "This code processed " & count & " items in " & time & " seconds"
    ElseIf count = 1 Then
        CounterResults = "This code processed " & count & " item in " & time & " seconds"
    Else
        CounterResults = "This code processed nothing but cycled through all items in " & time & " seconds"
    End If

End Function

Function IfError(formula As Variant, error_result As String)
'VBA IFERROR function
'Return error_result if formula result is an error, or the value of the formula itself otherwise.
    On Error GoTo ErrorHandler

    If IsError(formula) Then
        IfError = error_result
    Else
        IfError = formula
    End If
    Exit Function
ErrorHandler:
    Resume Next

End Function

Function isLeapYear(yr As Integer) As Boolean
'A Leap Year occurs in years evenly divisible by four,
'except for centenary years not divisible by 400.
    'if the year divided by 400 has a remainder of zero
    If yr Mod 400 = 0 Then
        'then it is a leap year
        isLeapYear = True

    'or else if the year divided by 100 has a remainder of zero
    ElseIf yr Mod 100 = 0 Then
        'then it is NOT a leap year
        isLeapYear = False

    'or else if the year divided by 4 has a remainder of zero
    ElseIf yr Mod 4 = 0 Then
        'then it is a leap year
        isLeapYear = True
    'or else (default value)
    Else
        'it is NOT a leap year
        isLeapYear = False
    'end divisor test
    End If

End Function

Function NetAmount(debit As Double, credit As Double) As Double
    NetAmount = debit - credit
End Function

Function ItemAge(ItemDate As Date, AgeDate As Date) As Double
    ItemAge = ItemDate - AgeDate
End Function

Function GetColumnNumber(find_value As String, sheet_name As String) As Integer
    'Get column number of matching text (column header)
	Dim res As Object
    'First try for 'xlWhole' match (match the entire cell contents)
    Set res = Sheets(sheet_name).Cells(1, 1).EntireRow.Find(What:=find_value _
                                                           , LookIn:=xlValues _
                                                           , LookAt:=xlWhole _
                                                           , SearchOrder:=xlByColumns _
                                                           , SearchDirection:=xlPrevious _
                                                           , MatchCase:=False)
    'If can't match on whole, the try matching part of the cell (will return the last find)
    If res Is Nothing Then
            Set res = Sheets(sheet_name).Cells(1, 1).EntireRow.Find(What:=find_value _
                                                           , LookIn:=xlValues _
                                                           , LookAt:=xlPart _
                                                           , SearchOrder:=xlByColumns _
                                                           , SearchDirection:=xlPrevious _
                                                           , MatchCase:=False)
            If res Is Nothing Then
                GetColumnNumber = 0
            Else
                GetColumnNumber = res.Column
            End If
    Else
        GetColumnNumber = res.Column
    End If
End Function

Function InList(value As Variant, list As Range) As Boolean
	'checks if the item is already present in the list
    If IsError(Application.VLookup(value, list, 1, 0)) Then 
        InList = False
    Else
        InList = True
    End If
            
End Function

Function michaelize(amount As Variant) As Double
    christina = Len(amount)
    michaelize = Round((amount ^ 2) / christina, 2)
End Function

Function TransCcyID(Lgr As String, DocCcy As String, ComCcy As String) As Boolean
    'Transaction Currency for JDE GL
    If DocCcy = ComCcy And Lgr = "AA" Then
        TransCcyID = True
    ElseIf DocCcy <> ComCcy And Lgr = "CA" Then
        TransCcyID = True
    Else
        TransCcyID = False
    End If
End Function
