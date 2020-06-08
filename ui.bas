Dim msgs As Object

Public Sub init_messages()
    Set msgs = CreateObject("Scripting.Dictionary")
    Select Case Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Case 1049: ' Russian
        msgs("Embed data") = "Внедрить данные"
        msgs("Break links") = "Разорвать связи"
        msgs("Clean designs") = "Очистить темы"
        msgs("Send") = "Отправить"
        msgs("Paste && replace") = "Вставить с заменой"
        msgs("Send only selected slides (%s)?\nSelected slides numbers: %s") = _
             "Отправить только выделенные слайды (%s)?\nНомера выбранных слайдов: %s"
        msgs("Removed unused designs (with templates): %s\nRemoved unused templates: %s") = _
             "Удалено неиспользуемых тем (с образцами слайдов): %s\nУдалено неиспользуемых образцов слайдов: %s"
        msgs("Failed to identify selection: objects embedded in charts are not supported") = _
             "Не удалось определить выделение: объекты внутри диаграмм не поддерживаются"
        msgs("Select one object on slide") = "Выберите один объект на слайде"
        msgs("More than one object is found in clipboard") = "В буфере обмена должен находиться один объект"
    End Select
End Sub

Public Function tr(message) As String
    'Translates message to UI language
    If msgs.exists(message) Then tr = msgs(message) _
    Else tr = message
End Function

Public Function fmt(s, ParamArray values()) As String
    'Replaces %s in a string with values passed as arguments
    'Replaces \n with vbCrLf
    fmt = s
    fmt = Replace$(fmt, "\n", vbCrLf)
    For Each v In values
        fmt = Replace$(fmt, "%s", v, Count:=1)
    Next v
End Function
