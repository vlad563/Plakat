Attribute VB_Name = "NewMacros"
Sub ������1()
'������ �� ������� Ctrl + D
'������� ������
    Selection.WholeStory
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    
    
     ' ��̲�� ������������ ������� �� ���������

'������� ������
    With Selection.Find
        .Text = "^s" '��������� �������
        .Replacement.Text = " "     '�������
        .Forward = True             '����*������ true
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True      '����*������ true
    End With
'���������� ��� ������ ������
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
    
    
    ' ��������� �����˲� ����� ������� �������ֲ�
    
'������� ������
    With Selection.Find
        .Text = " {1;}([.,:;\!\?])" '��������� �������
        .Replacement.Text = "\1"    '�������
        .Forward = True             '����*������ true
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True      '����*������ true
    End With
'���������� ��� ������ ������
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1
   
    
    ' ������� ������ ϲ��� ����� �������ֲ�
    
    '������� ������
    With Selection.Find
        .Text = "([.,:;\!\?])" '��������� �������
        .Replacement.Text = "\1 "    '�������
        .Forward = True             '����*������ true
        .Wrap = wdStore
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True      '����*������ true
    End With
'���������� ��� ������ ������
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveLeft Unit:=wdCharacter, Count:=1


    ' ��������� ������ �����˲�, �� ������������

'������� ������
    With Selection.Find
        .Text = " {2;}"              '��������� 2 � ����� ������
        .Replacement.Text = " "      '������� �� ���� �����
        .Forward = True              '����*������ true
        .Wrap = wdQuestion
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchAllWordForms = False
        .MatchSoundsLike = False
        .MatchWildcards = True       '����*������ true
    End With
'���������� ��� ������ ������
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.MoveRight Unit:=wdCharacter, Count:=1
    
End Sub