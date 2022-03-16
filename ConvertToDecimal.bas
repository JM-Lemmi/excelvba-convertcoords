Attribute VB_Name = "ConvertToDecimal"
Function Convert_Decimal(Degree_Deg As String) As Double
 ' Declare the variables to be double precision floating-point.
 Dim degrees As Double
 Dim minutes As Double
 Dim seconds As Double
 ' Set degree to value before "°" of Argument Passed.
 degrees = Val(Left(Degree_Deg, InStr(1, Degree_Deg, "°") - 1))
 ' Set minutes to the value between the "°" and the "'"
 ' of the text string for the variable Degree_Deg divided by
 ' 60. The Val function converts the text string to a number.
 minutes = Val(Mid(Degree_Deg, InStr(1, Degree_Deg, "°") + 2, _
 InStr(1, Degree_Deg, "'") - InStr(1, Degree_Deg, _
 "°") - 2)) / 60
 ' Set seconds to the number to the right of "'" that is
 ' converted to a value and then divided by 3600.
 seconds = Val(Mid(Degree_Deg, InStr(1, Degree_Deg, "'") + _
 2, Len(Degree_Deg) - InStr(1, Degree_Deg, "'") - 2)) _
 / 3600
 Convert_Decimal = degrees + minutes + seconds
End Function

' Courtesy of https://glenbambrick.com/2015/08/16/dms-to-dd-excel/
