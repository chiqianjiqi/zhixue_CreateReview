Sub a()
    '��ʼ��
    Dim ExcelFileType(3), SubName(11), TemplateSubject
    Dim FileLack$, TemplateName$, TeacherList(), ReviewList(), ReviewName()
    ReDim TeacherList(0), ReviewList(0)
    FileLack = "ȱ��"
    ExcelFileType(1) = "ģ��"
    ExcelFileType(2) = "��ʦ����"
    ExcelFileType(3) = "�ľ�����"
    SubName(1) = "����"
    SubName(2) = "��ѧ"
    SubName(3) = "Ӣ��"
    SubName(4) = "����"
    SubName(5) = "��ʷ"
    SubName(6) = "����"
    SubName(7) = "����"
    SubName(8) = "��ѧ"
    SubName(9) = "����"
    SubName(10) = "����"
    SubName(11) = "����"
    '������excel�ļ�
    ExcelFile = Dir(ThisWorkbook.Path & "\" & "*.xls*")
    If ExcelFile <> "" Then Workbooks.Open (ThisWorkbook.Path & "\" & ExcelFile)
    Do Until 0
        ExcelFile = Dir
        If ExcelFile = "" Then Exit Do
        Workbooks.Open (ThisWorkbook.Path & "\" & ExcelFile)
    Loop
    'ȷ���ļ����ͣ���¼�ļ���
    For FileCount = 1 To Workbooks.Count
        Select Case Workbooks(FileCount).Sheets(1).Name
            Case "��ʦ����"
                ExcelFileType(2) = ""
                ReDim Preserve TeacherList(UBound(TeacherList()) + 1)
                TeacherList(UBound(TeacherList())) = Workbooks(FileCount).Name
            Case "����Աģ��"
                ExcelFileType(1) = ""
                TemplateName = Workbooks(FileCount).Name
                For cou = 1 To 11
                    If InStr(TemplateName, SubName(cou)) > 0 Then TemplateSubject = SubName(cou)
                Next
            '�������Ƿ����ѧ�ƣ�������Ϊ�ľ�����
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
    '����Ƿ�ȱʧ�ļ�
    For i = 1 To 3
        If ExcelFileType(i) <> "" Then FileLack = FileLack & vbCrLf & ExcelFileType(i)
    Next
    If FileLack <> "ȱ��" Then
        FileLack = FileLack & vbCrLf & "�ļ�"
        MsgBox FileLack
        Exit Sub
    End If
    MsgBox "��ʦ������" & UBound(TeacherList()) & "��" & ",�ľ�������" & UBound(ReviewList()) & "��"
    '��ʼ�������
    Open Workbooks(1).Path & "\������־.txt" For Output As #1 '׼��������־�ļ�
    Print #1, "���½�ʦ�ڽ�ʦ������δ�ҵ������������Ƿ���ȷ���ֶ����"
    Dim FindErr As Boolean, AlFind As Boolean
    AlFind = False
    FindErr = False
    ExamCount = 1
    Do
        ExamIndex = Mid(Workbooks(TemplateName).Sheets(1).Cells(2, ExamCount).Text, 2, Len(Workbooks(TemplateName).Sheets(1).Cells(2, ExamCount).Text) - 2)
        'ģ����Ŷ�Ӧ�ľ����
        For ReviewListCount = 1 To UBound(ReviewList())
            ReDim ReviewName(0) 'ÿ�γ�ʼ����ŵ�ǰ��Ŀ�ľ��ʦ������
            ReviewListExamCount = 3 '��Ŀ�����λ��
            Do Until Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 3) = ""
                If InStr(Workbooks(ReviewList(ReviewListCount)).Sheets(TemplateSubject).Cells(ReviewListExamCount, 3), "��" & ExamIndex & "��") <> 0 Then
                    '����ľ�������������
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
                        '��ʦ������Ѱ�Ҳ�ֳ������ֲ����
                    TemplateRow = 4 '��ģ���4�п�ʼ���
					For TeacherListNameFind = 1 To UBound(ReviewName())
                        AlFind = False
                        For TeacherListName = 1 To UBound(TeacherList()) '��ʦ����ѭ������
                            For TeacherListNameFind2 = 3 To Workbooks(TeacherList(TeacherListName)).Sheets(1).UsedRange.Rows.Count '��ʦ�����ӵ�һ�е����һ��
                                If Workbooks(TeacherList(TeacherListName)).Sheets(1).Cells(TeacherListNameFind2, 2) = ReviewName(TeacherListNameFind) Then
                                    Workbooks(TemplateName).Sheets(1).Cells(TemplateRow, ExamCount + 2) = Workbooks(TeacherList(TeacherListName)).Sheets(1).Cells(TeacherListNameFind2, 1)
                                    TemplateRow = TemplateRow + 1
                                    AlFind = True
                                    Exit For
                                End If
                            Next
							If AlFind = True Then Exit For
                        Next
						If AlFind = False Then 'û���ڽ�ʦ�������ҵ����ľ��ʦ��Ϣ
                            Print #1, Workbooks(ReviewList(ReviewListCount)).Name, "��" & ExamIndex & "��", ReviewName(TeacherListNameFind) 'д����־
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
        Kill Workbooks(1).Path & "\������־.txt"
        MsgBox "ģ�����ɹ�'"
    Else
        MsgBox "ģ�����δ�ɹ������������־"
        Shell "notepad.exe " & Workbooks(1).Path & "\������־.txt", vbNormalFocus
    End If
End Sub
'ȱʧ���ܣ�
'������Ŀ�ľ��ʦ�����ã�û�����ӣ�