<div align="center">

## Filling a form in the WebBrowser Control \(automatically\)


</div>

### Description

Ever want to be able to automate a web page process through a Visual Basic front end or form you've created? This code helps you control a web page through the WebBrowser control to emulate actually entering data and submitting it through the web page itself. Check it out. This code is from http://vbpoint.cjb.net/ and there are more useful code there also.
 
### More Info
 
In sample I will fill the Altavista search box, with the WebBrowser control. Below I will list some subs and functions which are used in this sample. Open a new project (standard exe) and place a WebBrowser control, a textbox, and a command button on form1.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Chris Kesler](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/chris-kesler.md)
**Level**          |Beginner
**User Rating**    |4.9 (78 globes from 16 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Internet/ HTML](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/internet-html__1-34.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/chris-kesler-filling-a-form-in-the-webbrowser-control-automatically__1-13439/archive/master.zip)





### Source Code

```
Add the following code to form1:
Private Sub Command1_Click()
 Dim doc As HTMLDocument 'Reference MSHTML.TLB - may end up being IHTMLDocument3
 'go to the altavista (text) search page
 WebBrowser1.Navigate "http://www.altavista.com/cgi-bin/query?text"
 'Wait until page is loaded
 Do
 DoEvents
 Loop Until Not WebBrowser1.Busy
 'Make doc reference to the document inside the webbrowser control
 Set doc = WebBrowser1.Document
 'Set field q with the value of Text1
 SetInputField doc, 0, "q", Text1
 'Submit the form (same result as click the search button)
 doc.Forms(0).submit
 'Wait until result are loaded
 Do
 DoEvents
 Loop Until Not WebBrowser1.Busy
 MsgBox "Altavista search result loaded"
End Sub
'Add the following code to a module:
Public Sub SetInputField(doc As HTMLDocument, Form As Integer, Name As String, Value As String)
'doc = HTMLDocument, can be retrieved
' from webbrowser --> webbrowser.document
'Form = number of the form
' (if only one form in the doc --> Form = 0)
'Name = Name of the field you would like to fill
'Value = The new value for the input field called name
'PRE: Legal parameters entered
'POST: Input field with name Name on form Form in document doc will be filled with Value
 For q = 0 To doc.Forms(Form).length - 1
 If doc.Forms(Form)(q).Name = Name Then
 doc.Forms(Form)(q).Value = Value
 Exit For
 End If
 Next q
End Sub
'Additional useful subs:
'Sub to get the contents from a textbox:
Public Function GetInputField(doc As HTMLDocument, Form As Integer, Name As String) As String
 For q = 0 To doc.Forms(Form).Length - 1
 If doc.Forms(Form)(q).Name = Name Then
 GetInputField = doc.Forms(From)(q).Value
 Exit For
 End If
 Next q
End Function
'Sub to set a Checkbox:
Public Sub SetCheckBox(doc As HTMLDocument, Form As Integer, Name As String, Value As Boolean)
 For q = 0 To doc.Forms(Form).Length - 1
 If doc.Forms(Form)(q).Name = Name Then
 doc.Forms(From)(q).Checked = Value
 Exit For
 End If
 Next q
End Sub
'Sub set a radio button:
Public Sub SetRadioButton(doc As HTMLDocument, Form As Integer, Name As String, Name2 As String)
 For q = 0 To doc.Forms(Form).Length - 1
 If (doc.Forms(Form)(q).Name = Name) And (doc.Forms(Form)(q).Value = Name2) Then
 doc.Forms(From)(q).Checked = True
 Exit For
 End If
 Next q
End Sub
'Sub to make a selection in a ComboBox with Option Values:
Public Function SetComboBoxValue(ByVal doc As IHTMLDocument3, Form As Integer, Name As String, Name2 As String)
Dim q, i
For q = 0 To doc.Forms(Form).length - 1
  If (doc.Forms(Form)(q).Name = Name) Then
    For i = 0 To doc.Forms(Form)(q).length - 1
      If doc.Forms(Form)(q).Options(i).Value = Name2 Then
        doc.Forms(Form)(q).Options(i).Selected = True
        Exit For
      End If
    Next i
  End If
Next q
End Function
'Sub to make a selection in a ComboBox without Option Values:
Public Function SetComboTextValue(ByVal doc As IHTMLDocument3, Form As Integer, Name As String, Name2 As String)
Dim q, i
For q = 0 To doc.Forms(Form).length - 1
  If (doc.Forms(Form)(q).Name = Name) Then
    For i = 0 To doc.Forms(Form)(q).length - 1
      If doc.Forms(Form)(q).Options(i).Text = Name2 Then
        doc.Forms(Form)(q).Options(i).Selected = True
        Exit For
      End If
    Next
  End If
Next q
End Function
```

