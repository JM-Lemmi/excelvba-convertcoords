Function ConvertToDecimal(dms As String) As Double
 ' Declare the variables to be double precision floating-point.
 Dim degrees As Double
 Dim minutes As Double
 Dim seconds As Double
 ' Declare the special characters for splitting
 Dim vbDblQuote As String
 vbDblQuote = Chr(34)
 
 ' Set degree to value before ° of Argument Passed.
 degrees = Val(Left(dms, InStr(1, dms, "°") - 1))
 ' Set minutes to the value between the ° and the ' and then divides by 60.
 ' The Val function converts the text string to a number.
 minutes = Val(Mid(dms, InStr(1, dms, "°") + 1, InStr(1, dms, "'") - InStr(1, dms, "°") - 1)) / 60
 ' Set seconds to the value between the ' and the " and then divides by 3600.
 seconds = Val(Mid(dms, InStr(1, dms, "'") + 1, InStr(1, dms, vbDblQuote) - InStr(1, dms, "'") - 1)) / 3600
 Convert_Decimal = degrees + minutes + seconds
End Function

' Courtesy of https://glenbambrick.com/2015/08/16/dms-to-dd-excel/
' Double Quotes from https://stackoverflow.com/a/28507279
