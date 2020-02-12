Function project_id() As Long
    ' Generate project ID (1-1000) to find it in Application.AddIns collection
    Static unique As Long
    If unique = 0 Then
        Randomize
        unique = Int((1000 - 1 + 1) * Rnd + 1)
    End If
    project_id = unique
End Function

Function project_name() As String
    Dim i
    For Each i In Application.AddIns
        On Error GoTo project_name__next:
        If i.Loaded Then ' Otherwise add-in will be loaded on Application.Run
            If project_id = Application.Run(i.Name & "!project_id") Then
                project_name = i.Name
                Exit Function
            End If
        End If
project_name__next:
    Next i
End Function

Sub Auto_Open()
    Dim bar As CommandBar, b As CommandBarButton
    Call Auto_Close
    Set bar = CommandBars.Add(barId, Temporary:=True)

    Set b = bar.Controls.Add(msoControlButton)
    b.Caption = "Внедрить данные"
    b.OnAction = project_name & "!chartDataRecover"
    b.Style = msoButtonIconAndCaption
    b.FaceId = 17

    Set b = bar.Controls.Add(msoControlButton)
    b.Caption = "Разорвать связи"
    b.OnAction = project_name & "!chartBreakLinks"
    b.Style = msoButtonIconAndCaption
    b.FaceId = 1088

    Set b = bar.Controls.Add(msoControlButton)
    b.Caption = "Очистить темы"
    b.OnAction = project_name & "!remove_unused_designs"
    b.Style = msoButtonIconAndCaption
    b.FaceId = 47

    Set b = bar.Controls.Add(msoControlButton)
    b.Caption = "Отправить"
    b.OnAction = project_name & "!send_via_outlook"
    b.Style = msoButtonIconAndCaption
    b.FaceId = 24

    bar.Visible = True
End Sub

Sub Auto_Close()
    Dim i As CommandBar
    For Each i In CommandBars
        If i.Name = barId Then i.Delete
    Next i
End Sub
