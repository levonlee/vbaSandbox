Option Compare Database
Option Explicit

Public Function RegexReplace( _
        originalText As Variant, _
        regexPattern As String, _
        replaceText As String, _
        Optional GlobalReplace As Boolean = True) As Variant
    Dim rtn As Variant
    Dim objRegExp As Object  ' RegExp

    rtn = originalText
    If Not IsNull(rtn) Then
        Set objRegExp = CreateObject("VBScript.RegExp")
        objRegExp.Pattern = regexPattern
        objRegExp.Global = GlobalReplace
        rtn = objRegExp.Replace(originalText, replaceText)
        Set objRegExp = Nothing
    End If
    RegexReplace = rtn
End Function
