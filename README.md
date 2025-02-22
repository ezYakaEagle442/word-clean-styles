# word-clean-styles
VB Macro to clean unused styles 

```bash
Public Sub CleanStyles()
For Each oStyle In ActiveDocument.Styles
If Not oStyle.BuiltIn Then
With ActiveDocument.Content
.Find.ClearFormatting
.Find.Style = ActiveDocument.Styles(oStyle)
If Not .Find.Execute() Then
oStyle.Delete:   n = n + 1
End If
End With
End If
Next oStyle MsgBox Str(n) & "Unusedd styles have been removed."
End Sub
```

