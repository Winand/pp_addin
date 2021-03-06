Option Explicit
Private Const barId As String = "Winand's Tools"

Public Sub save_selected(filepath)
    ' Export selected slides to `filepath`
    Dim sl, sel_ids, pr, cur_idx, del_idc() As Long
    Set sel_ids = CreateObject("Scripting.Dictionary")
    For Each sl In selectedSlides
        Set sel_ids(sl.SlideID) = sl
    Next sl

    ActivePresentation.SaveCopyAs filepath
    Set pr = Presentations.Open(filepath, WithWindow:=False)
    If ActivePresentation.Slides.Count - sel_ids.Count > 0 Then
        ReDim del_idc(1 To ActivePresentation.Slides.Count - sel_ids.Count)
        
        For Each sl In pr.Slides
            If Not sel_ids.exists(sl.SlideID) Then
                cur_idx = cur_idx + 1
                del_idc(cur_idx) = sl.SlideIndex
            End If
        Next sl
        pr.Slides.Range(del_idc).Delete
    End If
    Call remove_unused_designs__internal(pr)
    pr.Save
    pr.Close
End Sub

Private Function generate_temp_path() As String
    ' Generate path to save active presentation in temp folder
    Dim file_name As String
    file_name = ActivePresentation.Name & IIf(ActivePresentation.Path = "", ".pptx", "")
    generate_temp_path = Environ("Temp") & "\" & file_name
End Function

Private Sub new_outlook_msg(subject, attachment_path)
    ' Create and display new Outlook message
    ' with `subject` and attach file `attachment_path`
    Dim objMsg, app
    Const olMailItem = 0
    Set app = CreateObject("Outlook.Application")
    Set objMsg = app.CreateItem(olMailItem)
    objMsg.subject = subject
    objMsg.Attachments.Add attachment_path
    objMsg.Display
    AppActivate app.ActiveInspector.Caption ' Bring message to front
End Sub

Public Sub send_selected_via_outlook()
    ' Creates new Outlook message and attaches selected slides from active presentation
    Dim tmp_file_path
    If is_protected_view Then Exit Sub
    tmp_file_path = generate_temp_path
    save_selected tmp_file_path
    
    On Error GoTo send__outlook_error:
    new_outlook_msg subject:=ActivePresentation.Name, attachment_path:=tmp_file_path

send__outlook_error:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    Kill tmp_file_path
End Sub

Public Sub send_via_outlook()
    ' Creates new Outlook message and attaches active presentation
    ' If slide thumbnails are selected (not all) calls `send_selected_via_outlook`
    If ActiveWindow.Selection.Type = ppSelectionSlides Then
        Dim slide_rng As SlideRange
        Set slide_rng = ActiveWindow.Selection.SlideRange
        If slide_rng.Count < ActiveWindow.Presentation.Slides.Count Then
            Select Case _
                MsgBox(fmt(tr("Send only selected slides (%s)?\nSelected slides numbers: %s"), _
                           slide_rng.Count, to_text_range(slide_rng)), _
                       vbInformation + vbYesNoCancel)
            Case vbYes:
                Call send_selected_via_outlook
                Exit Sub
            Case vbNo: 'Continue
            Case vbCancel:
                Exit Sub
            End Select
        End If
    End If
    Dim tmp_file_path
    If is_protected_view Then Exit Sub
    If is_saved_to_disk(ActivePresentation) Then 'actual state is already saved
        tmp_file_path = ActivePresentation.FullName
    Else
        tmp_file_path = generate_temp_path
        ActivePresentation.SaveCopyAs tmp_file_path
    End If
    On Error GoTo send__outlook_error:
    new_outlook_msg subject:=ActivePresentation.Name, attachment_path:=tmp_file_path

send__outlook_error:
    If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
    If tmp_file_path <> ActivePresentation.FullName Then
        ' Delete file if it's not main presentation file
        Kill tmp_file_path
    End If
End Sub

Private Function to_text_range(slide_rng As SlideRange) As String
    'Converts slide range to text: "1-3,5,8-13"
    Dim arr, i, prev As Long, rng_started As Boolean
    Set arr = CreateObject("System.Collections.ArrayList")
    For Each i In slide_rng
        arr.Add i.SlideNumber
    Next i
    arr.Sort
    For i = 0 To arr.Count - 1
        If i = 0 Then
            to_text_range = arr(i)
        ElseIf arr(i) - prev = 1 Then
            If Not rng_started Then to_text_range = to_text_range & "-"
            rng_started = True
        Else
            If rng_started Then to_text_range = to_text_range & prev
            to_text_range = to_text_range & ", " & arr(i)
            rng_started = False
        End If
        prev = arr(i)
    Next i
    If rng_started Then to_text_range = to_text_range & prev
End Function

Function is_protected_view() As Boolean
    ' Check if currently opened window is protected view window
    Dim tmp
    On Error GoTo err__is_protected_view:
    Set tmp = ActiveWindow
err__is_protected_view:
    ' In protected view `ActiveProtectedViewWindow` is used instead of `ActiveWindow`
    is_protected_view = IIf(Err.Number = &H80048240, True, False)
End Function

Public Function is_saved_to_disk(pr) As Boolean
    ' Check if actual state of presentation `pr` is saved to disk because
    ' `Saved` property returns True for new unchanged presentations
    If pr.Path <> "" And pr.Saved Then is_saved_to_disk = True
End Function

Function get_used_layouts(pr)
    ' Returns used designs and layouts in a presentation `pr` as a Dictionary:
    ' Design1->number_of_users, Design1{null_char}Layout1->number_of_users, etc.
    Dim used_layouts, layout, sl, l_name, i
    Set used_layouts = CreateObject("Scripting.Dictionary")
    For Each sl In pr.Slides
        Set layout = sl.CustomLayout
        l_name = layout.Design.Name
        used_layouts(l_name) = used_layouts(l_name) + 1
        l_name = layout.Design.Name & vbNullChar & layout.Name
        used_layouts(l_name) = used_layouts(l_name) + 1
    Next sl
    Set get_used_layouts = used_layouts
End Function

Private Function remove_unused_designs__internal(pr)
    ' Remove unused designs and layouts in `pr` presentation
    ' Return number of unused designs and unused layouts within used designs
    ' Returns Array(removed_designs, removed_layouts)
    Dim used_layouts, d, l
    Dim removed As Long, removed_d As Long
    Dim col As New Collection

    Set used_layouts = get_used_layouts(pr)
    For Each d In pr.Designs
        If Not used_layouts.exists(d.Name) Then
            col.Add d
            removed_d = removed_d + 1
        End If
    Next d
    For Each d In col
        d.Delete
    Next d
    Set col = New Collection
    For Each d In pr.Designs
        For Each l In d.SlideMaster.CustomLayouts
            If Not used_layouts.exists(d.Name & vbNullChar & l.Name) Then
                col.Add l
                removed = removed + 1
            End If
        Next l
    Next d
    For Each l In col
        l.Delete
    Next l
    remove_unused_designs__internal = Array(removed_d, removed)
End Function

Sub remove_unused_designs()
    Dim result
    If is_protected_view Then Exit Sub
    result = remove_unused_designs__internal(ActivePresentation)
    MsgBox fmt(tr("Removed unused designs (with templates): %s\nRemoved unused templates: %s"), _
               result(0), result(1)), vbInformation
End Sub

Function chartTemplatesFolder() As String
    On Error GoTo er:
    Dim templatesFolder As String
    templatesFolder = CreateObject("WScript.Shell").RegRead( _
        "HKCU\Software\Microsoft\Office\" & Application.Version & "\Common\General\Templates")
    chartTemplatesFolder = Environ("AppData") & "\Microsoft\" & templatesFolder & "\Charts"
Exit Function
er:
End Function

Private Function zorder_reached(new_z, target_z, direction As MsoZOrderCmd) As Long
    'Check if target ZOrder has been reached or passed already
    'Returns:
    '* True or -1 (reached), False or 0 (not yet reached),
    '* msoSendBackward, msoBringForward (passed already, change direction and step back)
    If new_z = target_z Then
        zorder_reached = True
    ElseIf direction = msoBringForward And new_z > target_z Then
        zorder_reached = msoSendBackward
    ElseIf direction = msoSendBackward And new_z < target_z Then
        zorder_reached = msoBringForward
    End If
End Function

Sub setZOrder(obj, pos)
    'Move object to specified ZOrder /pos/
    Dim direction As Long
    direction = IIf(obj.ZOrderPosition < pos, msoBringForward, msoSendBackward)
    While zorder_reached(obj.ZOrderPosition, pos, direction) = False
        obj.ZOrder direction
    Wend
    direction = zorder_reached(obj.ZOrderPosition, pos, direction)
    if direction <> True then obj.ZOrder direction 'one step back
End Sub

Function hasTitle(obj) As Boolean
    'Some charts' HasTitle=False while ChartTitle is present
    On Error Resume Next
    If obj.hasTitle Then hasTitle = True _
    Else hasTitle = Not obj.ChartTitle Is Nothing
End Function

Sub kill_or_not(ByRef fp As String)
On Error GoTo er
    Kill fp
er:
End Sub

Function getLongestSeries(sc) As Long
    Dim i As Long
    If sc.Count Then
        getLongestSeries = 1
        For i = 1 To sc.Count
            If UBound(sc(i).Values) > UBound(sc(getLongestSeries).Values) Then getLongestSeries = i
        Next i
    End If
End Function

Function rng(wb, er, ec, Optional sr = 1, Optional sc = 1) As Object
    Set rng = wb.Range(wb.Cells(sr, sc), wb.Cells(er, ec))
End Function

Sub copyPos(o1, o2, Optional resize As Boolean = True)
    'Set global (in slide coordinates) position of /o2/ to /o1/
    Dim top As Long, left As Long, o_tmp
    If resize Then
        o2.Width = o1.Width
        o2.Height = o1.Height
    End If
    Set o_tmp = o1
    While TypeName(o_tmp) <> "Slide"
        left = left + o_tmp.left
        top = top + o_tmp.top
        Set o_tmp = o_tmp.Parent
    Wend
    o2.left = left
    o2.top = top
End Sub

Function getParentSlide(sh) As Object
    'iterates through nesting shapes until "root" slide is found
    Set getParentSlide = sh.Parent
'    Debug.Print getParentSlide.Name, TypeName(getParentSlide)
    While TypeName(getParentSlide) <> "Slide"
        Set getParentSlide = getParentSlide.Parent
    Wend
End Function

'Function getXValues(ser) As Variant
''Get Values if XValues fail or empty
'On Error GoTo er:
'    Dim i
'    getXValues = ser.XValues
'    If Not IsEmpty(getXValues(0)) Then Exit Function
''    For Each i In getXValues
''        If Not IsEmpty(i) Then Exit Function
''    Next i
'er: getXValues = ser.Values 'FIXME: 1,2,3,4...
'End Function

Function getXValues(ser) As Variant
'Get Values if XValues fail or empty
On Error GoTo er:
    getXValues = ser.XValues
Exit Function
er: getXValues = ser.Values 'FIXME: 1,2,3,4...
End Function

Sub chartDataRecover()
    Dim isl As slide, j As shape, ch As Chart, ch2 As Chart, k As Long, charts As New Collection, ws As Object
    Dim ls As Long, i, o1, o2, left_shift As Long
    
    If is_protected_view Then Exit Sub
    For Each isl In selectedSlides
        For Each j In isl.Shapes
            If j.Type = msoChart And _
            j.left < ActivePresentation.PageSetup.SlideWidth Then
                If j.left < left_shift Then left_shift = j.left
                charts.Add j.Chart
                unbox j.Chart, charts
            End If
        Next j
    Next isl
    Debug.Print charts.Count
    For Each ch In charts
        ls = getLongestSeries(ch.SeriesCollection)
        If ls Then
            Set ch2 = getParentSlide(ch).Shapes.AddChart.Chart
            copyPos ch.Parent, ch2.Parent
            ch2.hasTitle = ch.hasTitle
            If ch2.hasTitle Then _
                ch2.ChartTitle.Caption = ch.ChartTitle.Caption
            Call setZOrder(ch2.Parent, ch.Parent.ZOrderPosition)

            Set ws = ch2.ChartData.Workbook.WorkSheets(1)
            Call ws.Range("A2:D5").ClearContents
            Set o1 = rng(ws, UBound(ch.SeriesCollection(ls).Values) + 1, ch.SeriesCollection.Count + 1)
            ws.ListObjects(1).resize o1
            rng(ws, UBound(ch.SeriesCollection(ls).Values) + 1, 1, 2, 1) = _
                ws.Application.Transpose(getXValues(ch.SeriesCollection(ls)))
            o1.wraptext = False

            ch.SaveChartTemplate "winand_temp"
            ch2.ApplyChartTemplate "winand_temp" 'apply AFTER data source resize
            
            For Each i In ch.Axes 'FIXME: Если нет подписей оси, то TickLabels выдаёт ошибку (?)
                ch2.Axes(i.Type, i.AxisGroup).TickLabels.NumberFormat = i.TickLabels.NumberFormat 'Fix percent labels format
            Next i
            For k = 1 To ch.SeriesCollection.Count
                Set o1 = ch.SeriesCollection(k)
                Set o2 = ch2.SeriesCollection(k)
                ws.Cells(1, k + 1) = o1.Name
                rng(ws, UBound(o1.Values) + 1, k + 1, 2, k + 1) = _
                    ws.Application.Transpose(o1.Values)
                If o1.HasDataLabels Then o2.DataLabels.NumberFormat = o1.DataLabels.NumberFormat 'Fix percent labels format
                ch2.Refresh 'Otherwise series is invisible
            Next k
            ws.Parent.Close
            If Not hasTitle(ch) And hasTitle(ch2) Then ch2.ChartTitle.Delete 'Title (of a series) may be added to a new chart (though HasTitle=False) even if the old one has no title
            ch.Parent.left = -left_shift + ch.Parent.left + ActivePresentation.PageSetup.SlideWidth 'go out!
        End If
    Next ch
    kill_or_not chartTemplatesFolder() & "\winand_temp.crtx"
End Sub

Sub unbox(ch, toCol) 'FIXME: toCol is not used
    Dim i, j, l As Single, t As Single, w As Single, h As Single, slide As Object, Name As String
    Set slide = getParentSlide(ch)
    For Each j In ch.Shapes
        If j.HasChart Then
            ActiveWindow.View.GotoSlide slide.SlideIndex
            l = j.left: t = j.top: w = j.Width: h = j.Height
            Name = Int(Rnd * 100) & "_" & j.Name
            j.Name = Name
            j.Select
            ActiveWindow.Selection.Cut
            slide.Shapes.Paste
            Set i = slide.Shapes(Name)
            i.left = l + ch.Parent.left: i.top = t + ch.Parent.top
            i.Width = w: i.Height = h
            setZOrder i, ch.Parent.ZOrderPosition + 1
        End If
    Next j
End Sub

Function selectedSlides() As Collection
    Dim sel As Selection, sl As New Collection, i As slide, View As PpViewType
    If Presentations.Count Then
        ActiveWindow.ViewType = ppViewNormal
        Set sel = ActiveWindow.Selection
        If sel.Type <> ppSelectionSlides Then
            ensureSlideSelected
            sl.Add ActiveWindow.View.slide
        Else
            For Each i In sel.SlideRange
                sl.Add i
            Next i
        End If
    End If
    Set selectedSlides = sl
End Function

Sub ensureSlideSelected()
    On Error GoTo 1:
    Dim cnt As Long
    cnt = ActiveWindow.Selection.SlideRange.Count
Exit Sub
1:  ActiveWindow.ViewType = ppViewSlide
    ActiveWindow.ViewType = ppViewNormal
End Sub

Sub chartBreakLinks()
On Error GoTo er:
    Dim i As shape
    If is_protected_view Then Exit Sub
    For Each i In ActiveWindow.Selection.ShapeRange
        i.LinkFormat.BreakLink
    Next i
er:
If Err.Number Then Debug.Print Err.Description
End Sub

Function selected_shapes(Optional ByRef embedded_sel) As ShapeRange
'Get selected shapes or empty `ShapeRange`
'`embedded_sel` flag can be set if shapes inside a chart are selected.
'               Zero-len ShapeRange is returned then.
On Error GoTo err__selected_shapes:
    Dim sel As Selection
    Set sel = ActiveWindow.Selection
    'Do not rely on sel.Type, 'cause ppSelectionText can be set when
    'text in a selected shape is being edited and in slide notes too
    If sel.HasChildShapeRange Then
        Set selected_shapes = sel.ChildShapeRange 'Shapes inside a group
    Else
        Set selected_shapes = sel.ShapeRange
    End If
    If selected_shapes.Count = 0 Then
        embedded_sel = True 'Warning: an object inside a chart is selected
        Err.Raise -1
    Else
        embedded_sel = False
    End If
Exit Function
err__selected_shapes: 'Return zero length range
    Set selected_shapes = ActiveWindow.View.slide.Shapes.Range(0)
End Function

Function paste_source_formatting() As ShapeRange
'Pastes data with source formatting. Works with charts and tables
On Error GoTo err__paste_source_formatting:
    Dim old_sel As ShapeRange, old_shape_count As Long, new_shape_count As Long
    Dim slide_shapes As Shapes, arr, i As Long
    Set slide_shapes = ActiveWindow.View.slide.Shapes
    old_shape_count = slide_shapes.Count
    Set old_sel = selected_shapes()
    'Multiple charts fail to be pasted if a chart is selected on a slide
    If old_sel.Count Then ActiveWindow.Selection.Unselect
    'PasteExcelTableSourceFormatting, PasteExcelChartSourceFormatting, PasteSourceFormatting
    CommandBars.ExecuteMso "PasteSourceFormatting"
    DoEvents 'Wait for `ExecuteMso` result
    new_shape_count = slide_shapes.Count
    ReDim arr(1 To new_shape_count - old_shape_count) As Long
    For i = old_shape_count + 1 To new_shape_count
        arr(i - old_shape_count) = i
    Next i
    Set paste_source_formatting = slide_shapes.Range(arr)
    old_sel.Select 'Restore selection
Exit Function
err__paste_source_formatting: 'Return zero length range
    Debug.Print "paste_source_formatting err:", Err.Description
    Set paste_source_formatting = slide_shapes.Range(0)
End Function

Sub paste_and_replace_shape()
' Заменяет выделенный объект объектом из буфера обмена,
' сохраняя положение и ZOrder
' FIXME: не поддерживаются (не тестировалось) вложенные объекты
On Error GoTo err__paste_and_replace_shape:
    Dim rng As ShapeRange, old_obj As Shape, new_obj As Shape, is_emb_selection As Boolean
    Set rng = selected_shapes(is_emb_selection)
    If is_emb_selection Then _
        Err.Raise -1, , tr("Failed to identify selection: objects embedded in charts are not supported")
    If rng.Count <> 1 Then _
        Err.Raise -1, , tr("Select one object on slide")
    Set old_obj = rng(1)
    Set rng = paste_source_formatting()
    If rng.Count = 0 Then _
        Set rng = ActiveWindow.View.slide.Shapes.Paste 'fallback
    Set new_obj = rng(1)
    If rng.Count > 1 Then
        rng.Delete
        old_obj.Select 'If text is pasted focus is set on it
        Err.Raise -1, , tr("More than one object is found in clipboard")
    End If
    copyPos old_obj, new_obj
    setZOrder new_obj, old_obj.ZOrderPosition
    old_obj.Delete
    new_obj.Select
Exit Sub
err__paste_and_replace_shape:
    MsgBox Err.Description, vbExclamation
End Sub
