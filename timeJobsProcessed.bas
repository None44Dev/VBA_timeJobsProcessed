Attribute VB_Name = "timeJobsProcessed"
Option Explicit
Private Const STARTJOBDAY  As Double = 0.375, ENDJOBDAY As Double = 0.75
Private dictJobDay As New Dictionary
Private dictDateLog As New Dictionary
Private Sub CalculateWorkDays()
Dim i&
Dim Infinity As Double
Dim arrTableJobDay As Variant
Dim arrTableDate As Variant
Dim TimeAndStatusTmp As Variant
Dim arrTmp As Variant

Call DisabledApps(False, False)
Call FncTime(Infinity, False)

' ��������� �������� �� ������
arrTableJobDay = [tableJobDay].Columns(1).Value2
arrTableDate = [tableDate].Columns(1).Resize(, 2).Value2

' ���� ������� ���� � �������
For i = LBound(arrTableJobDay) To UBound(arrTableJobDay)
    dictJobDay(arrTableJobDay(i, 1)) = dictJobDay(arrTableJobDay(i, 1)) + dictJobDay.Count
Next i

ReDim arrResults(LBound(arrTableDate) To UBound(arrTableDate), 1 To 2)
'���������� �� ������ ����
For i = LBound(arrTableDate) To UBound(arrTableDate)
    TimeAndStatusTmp = TimeInWork(arrTableDate(i, 1), arrTableDate(i, 2))                            '����� ����������
    arrResults(i, 1) = TimeAndStatusTmp                                                              '���������� ����� � ������
Next i

dictJobDay.RemoveAll                                                                                 '������� �������
dictDateLog.RemoveAll                                                                                '������� �������
[tableDate].Columns(3).Resize(, 2) = arrResults                                                      '���������� ���������

Call DisabledApps(True, True)
Call FncTime(Infinity, True)
End Sub
Private Function TimeInWork(ByVal DateReg As Double, ByVal DateComp As Double) As Double: TimeInWork = 0
Dim Int_LDate&, Int_UDate&, tmpDays&
Dim StatusReg%, StatusComp%, StatusJobDay%
Dim tmpStopTmpDay As Boolean
Dim TimeReg As Double
Dim TimeComp As Double

Int_LDate = Int(DateReg)
Int_UDate = Int(DateComp)
TimeReg = DateReg - Int_LDate
TimeComp = DateComp - Int_UDate

' �������� ���� � ����
Select Case True
    Case Int_LDate = Int_UDate
        If Not dictJobDay.Exists(Int_LDate) Then TimeInWork = 0: Exit Function                                                                      '�������� ���� � ���� = ��������
        TimeInWork = (CheckStatusTime(TimeReg, TimeComp, 1) * 24)                                                                                   '������ ���� � ����, �����
'��������� ��������
    Case Else
        StatusJobDay = CheckJobDay(Int_LDate, Int_UDate)                                                                                            '�������� ������ ��� �� �������
        If dictDateLog.Exists(Int_LDate & "|" & Int_UDate) Then tmpDays = dictDateLog(Int_LDate & "|" & Int_UDate): tmpStopTmpDay = True            '��������� ���� � �������, ���� ���� - ��������� ����� ������� ����
        Select Case StatusJobDay
            Case Is = 1
                Select Case tmpStopTmpDay: Case False: tmpDays = FindJobDay(Int_LDate, Int_UDate) * 9: End Select                                   '���� �� � ������� ����
                TimeInWork = tmpDays                                                                                                                '���������������� � ��������� � ��������
            Case Is = 2
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate) - 2) * 9: End Select                             '���� �� � ������� ����
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      '���������������� � ��������� � ������
            Case Is = 3
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate, 3) - 1) * 9: End Select                          '���� �� � ������� ����
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      '���������������� � ������ � ��������� � ��������
            Case Is = 4
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate, 4) - 1) * 9: End Select                          '���� �� � ������� ����
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      '���������������� � ��������, �� ��������� � ������
        End Select
        If tmpStopTmpDay Then dictDateLog.Add Int_LDate & "|" & Int_UDate, tmpDays                                                                  '���������� ����� ���� ���|��� � �������, ��� ���������� ������ ������� ����
End Select

End Function
Private Function CheckJobDay(ByRef Int_LDate As Long, ByRef Int_UDate As Long) As Integer
Select Case True
    Case Not dictJobDay.Exists(Int_LDate) And Not dictJobDay.Exists(Int_UDate): CheckJobDay = 1         '���������������� � ��������� � ��������
    Case dictJobDay.Exists(Int_LDate) And dictJobDay.Exists(Int_UDate): CheckJobDay = 2                 '���������������� � ��������� � ������
    Case dictJobDay.Exists(Int_LDate) And Not dictJobDay.Exists(Int_UDate): CheckJobDay = 3             '���������������� � ������ � ��������� � ��������
    Case Not dictJobDay.Exists(Int_LDate) And dictJobDay.Exists(Int_UDate): CheckJobDay = 4             '���������������� � ��������, �� ��������� � ������
End Select
End Function
Private Function FindJobDay(Int_LDate As Long, Int_UDate As Long, Optional ByRef JobDay As Integer) As Double

'�������� �� ���� ����������, ���� ������� ����
While Not dictJobDay.Exists(Int_UDate) And Int_UDate > Int_LDate
    Int_UDate = Int_UDate - 1
Wend

'�������� �� ���� �����������, ���� ������� ����
While Not dictJobDay.Exists(Int_LDate) And Int_LDate < Int_UDate
    Int_LDate = Int_LDate + 1
Wend

'��������
Select Case True
    Case Int_LDate = Int_UDate:
        Select Case JobDay
            Case 3, 4: FindJobDay = dictJobDay(Int_UDate) - dictJobDay(Int_LDate) + 1                   '���� ������ ��� 3, 4
            Case 0: FindJobDay = 0                                                                      '���� �� ������� ������ ���
        End Select
    Case Else: FindJobDay = dictJobDay(Int_UDate) - (dictJobDay(Int_LDate) - 1)
End Select

End Function
Private Function CheckStatusTime(ByRef TimeReg As Double, ByRef TimeComp As Double, ByRef sameDayProcessed As Integer) As Double: Dim StatusReg As Integer, StatusComp As Integer
Select Case sameDayProcessed
    Case Is = 1, 2: StatusReg = CheckReg(TimeReg): StatusComp = CheckComp(TimeComp)                                      '����� ��� �������
    Case Is = 4: StatusReg = 0: StatusComp = CheckComp(TimeComp)                                                         '����� ������� ������ ��� ����������
    Case Is = 3: StatusReg = CheckReg(TimeReg): StatusComp = 0                                                           '����� ������� ������ ��� �����������
End Select

Select Case True
    Case StatusReg = 1 And StatusComp = 1 And sameDayProcessed = 1: CheckStatusTime = TimeComp - TimeReg                 '������ � �������� � 9:00 �� 18:00, ���� � ����
    Case StatusReg = 1 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = ENDJOBDAY - TimeReg                '������ � �������� � 9:00 �� 00:00, ���� � ����
    Case StatusReg = 2 And StatusComp = 1 And sameDayProcessed = 1: CheckStatusTime = TimeComp - STARTJOBDAY             '������ � �������� � 00:00 �� 18:00, ���� � ����
    Case StatusReg = 2 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = STARTJOBDAY                        '������ � �������� � 00:00 �� 00:00, ���� � ����
    Case StatusReg = 2 And StatusComp = 2 And sameDayProcessed = 1: CheckStatusTime = 0                                  '���������������� � ��������� �� 09:00, ���� � ����
    Case StatusReg = 3 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = 0                                  '���������������� � ��������� ����� 18:00, ���� � ����
    Case sameDayProcessed = 2, 3, 4:                                                                                     '���� ������ ���� �� "��������"
        Select Case StatusReg
            Case Is = 0: CheckStatusTime = 0
            Case Is = 1: CheckStatusTime = ENDJOBDAY - TimeReg
            Case Is = 2: CheckStatusTime = STARTJOBDAY
            Case Is = 3: CheckStatusTime = 0
        End Select
        Select Case StatusComp
            Case Is = 0: CheckStatusTime = CheckStatusTime + 0
            Case Is = 1: CheckStatusTime = CheckStatusTime + (TimeComp - STARTJOBDAY)
            Case Is = 2: CheckStatusTime = CheckStatusTime + 0
            Case Is = 3: CheckStatusTime = CheckStatusTime + STARTJOBDAY
        End Select
End Select

End Function
Private Function CheckReg(ByRef TimeReg As Double) As Integer
Select Case TimeReg
    Case STARTJOBDAY To ENDJOBDAY: CheckReg = 1             '� ��������� � 9:00 �� 18:00
    Case Is < STARTJOBDAY: CheckReg = 2                     '������ 9:00
    Case Is > ENDJOBDAY: CheckReg = 3                       '������ 18:00
End Select
End Function
Private Function CheckComp(ByRef TimeComp As Double) As Integer
Select Case TimeComp
    Case STARTJOBDAY To ENDJOBDAY: CheckComp = 1            '� ��������� � 9:00 �� 18:00
    Case Is < STARTJOBDAY: CheckComp = 2                    '������ 9:00
    Case Is > ENDJOBDAY: CheckComp = 3                      '������ 18:00
End Select
End Function
Private Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
'��������� application
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function
Private Function FncTime(ByRef Infinity As Double, ByVal Output As Boolean)
'����������� ����� ���������� �������
If Not Output Then Infinity = Timer
If Output Then MsgBox "������! " & Format(Timer - Infinity, "0.00 ���"), vbInformation
End Function
