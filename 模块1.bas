Sub a()
    '初始化
    Dim ExcelFileType(3), SubName(11), TemplateSubject
    Dim FileLack$, TemplateName$, TeacherList(), ReviewList(), ReviewName()
    ReDim TeacherList(0), ReviewList(0)
    FileLack = "缺少"
    ExcelFileType(1) = "模板"
    ExcelFileType(2) = "教师名单"
    ExcelFileType(3) = "阅卷名单"
    SubName(1) = "语文"
    SubName(2) = "数学"
    SubName(3) = "英语"
    SubName(4) = "政治"
    SubName(5) = "历史"
    SubName(6) = "地理"
    SubName(7) = "物理"
    SubName(8) = "化学"
    SubName(9) = "生物"
    SubName(10) = "文综"
    SubName(11) = "理综"
    '打开所有excel文件
    ExcelFile = Dir(ThisWorkbook.Path & "\" & "*.xls*")
    If ExcelFile <> "" Then Workbooks.Open (ThisWorkbook.Path & "\" & ExcelFile)
    Do Until 0
        ExcelFile = Dir
        If ExcelFile = "" Then Exit Do
        Workbooks.Open (ThisWorkbook.Path & "\" & ExcelFile)
    Loop
    '确定文件类型，记录文件名
    For FileCount = 1 To Workbooks.Count
        Select Case Workbooks(FileCount).Sheets(1).Name
            Case "教师名单"
                ExcelFileType(2) = ""
                ReDim Preserve TeacherList(UBound(TeacherList()) + 1)
                TeacherList(UBound(TeacherList())) = Workbooks(FileCount).Name
            Case "评卷员模板"
                ExcelFileType(1) = ""
                TemplateName = Workbooks(FileCount).Name
                For cou = 1 To 11
                    If InStr(TemplateName, SubName(cou)) > 0 Then TemplateSubject = SubName(cou)
                Next
            '工作表是否包含学科，包含则为阅卷名单
            Case Else
                For i = 1 To 11
                    If InStr(Workbooks(FileCount).Sheets(1).Name, SubName(i)) > 0 Then Exit For
                Next
                If i < 12 Then
                    ExcelFileType(3) = ""
                    ReDim Preserve ReviewList(UBound(ReviewList()) + 1)
                    ReviewList(UBound(ReviewList())) = Workbooks(FileCount).Name
                End If
        End Select
    Next
    '检测是否缺失文件
    For i = 1 To 3
        If ExcelFileType(i) <> "" Then FileLack = FileLack & vbCrLf & ExcelFileType(i)
    Next
    If FileLack <> "缺少" Then
        FileLack = FileLack & vbCrLf & "文件"
        MsgBox FileLack
        Exit Sub
    End If
    MsgBox "教师名单有" & UBound(TeacherList()) & "个" & ",阅卷名单有" & UBound(ReviewList()) & "个"
    '开始填充数据
    Open Workbooks(1).Path & "\错误日志.txt" For Output As #1 '准备错误日志文件
    Print #1, "以下教师在教师名单中未找到，请检查名字是否正确，手动添加"
    Dim FindErr As Boolean, AlFind As Boolean
    AlFind = False
    FindErr = False
    ExamCount = 1
    Do
        ExamIndex = Mid(Workbooks(TemplateName).Sheets(1).Cells(2, ExamCount).Text, 2, Len(Workbooks(TemplateName).Sheets(1).Cells(2, ExamCount).Text) - 2)
        '模板题号对应阅卷题号
        For ReviewListCount = 1 To UBound(ReviewList())
            ReDim ReviewName(0) '每次初始化存放当前题目阅卷教师的数组
            ReviewListExamCount = 3 '题目序号行位置
            Do Until Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 3) = ""
                If InStr(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 3), "第" & ExamIndex & "题") <> 0 Then
                    '拆分阅卷名单单题姓名
                    For ReviewListNameResolve = 1 To Len(RTrim(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 4)))
                        Compare = Mid(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 4), ReviewListNameResolve, 1)
                        If Asc(Compare) <> 32 And Asc(Compare) <> -23636 And Asc(Compare) <> -24158 Then TemReviewName = TemReviewName & Compare
                        If Asc(Compare) = 32 Or Asc(Compare) = -23636 Or Asc(Compare) = -24158 Then
                            ReDim Preserve ReviewName(UBound(ReviewName()) + 1)
                            ReviewName(UBound(ReviewName())) = TemReviewName
                            TemReviewName = ""
                            For ReviewListNameResolveBlank = ReviewListNameResolve To Len(RTrim(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 4)))
                                Compare = Mid(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 4), ReviewListNameResolveBlank, 1)
                                If Asc(Compare) <> 32 And Asc(Compare) <> -23636 And Asc(Compare) <> -24158 Then
                                    ReviewListNameResolve = ReviewListNameResolveBlank - 1
                                    Exit For
                                End If
                            Next
                        End If
                    Next
                    ReDim Preserve ReviewName(UBound(ReviewName()) + 1)
                    ReviewName(UBound(ReviewName())) = TemReviewName
                    TemReviewName = ""
                        '教师名单里寻找拆分出的名字并填充
                    TemplateRow = 4 '从模板第4行开始填充
					For TeacherListNameFind = 1 To UBound(ReviewName())
                        AlFind = False
                        For TeacherListName = 1 To UBound(TeacherList()) '教师名单循环变量
                            For TeacherListNameFind2 = 3 To Workbooks(TeacherList(TeacherListName)).Sheets(1).UsedRange.Rows.Count '教师名单从第一行到最后一行
                                If Workbooks(TeacherList(TeacherListName)).Sheets(1).Cells(TeacherListNameFind2, 2) = ReviewName(TeacherListNameFind) Then
                                    Workbooks(TemplateName).Sheets(1).Cells(TemplateRow, ExamCount + 2) = Workbooks(TeacherList(TeacherListName)).Sheets(1).Cells(TeacherListNameFind2, 1)
                                    TemplateRow = TemplateRow + 1
                                    AlFind = True
                                    Exit For
                                End If
                            Next
							If AlFind = True Then Exit For
                        Next
						If AlFind = False Then '没有在教师名单中找到该阅卷教师信息
                            Print #1, Workbooks(ReviewList(ReviewListCount)).Name, "第" & ExamIndex & "题", ReviewName(TeacherListNameFind) '写入日志
                            FindErr = True
                        End If
                    Next
                End If
                ReviewListExamCount = ReviewListExamCount + 1
            Loop
        Next
        ExamCount = ExamCount + 3
    Loop Until Workbooks(TemplateName).Sheets(1).Cells(2, ExamCount) = ""
    Close #1
    If FindErr = False Then
        Kill Workbooks(1).Path & "\错误日志.txt"
        MsgBox "模板填充成功'"
    Else
        MsgBox "模板填充未成功，请检查错误日志"
        Shell "notepad.exe " & Workbooks(1).Path & "\错误日志.txt", vbNormalFocus
    End If
End Sub
'缺失功能：
'特殊题目阅卷教师的设置（没有例子）