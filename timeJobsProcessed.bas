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

' Получение массивов из таблиц
arrTableJobDay = [tableJobDay].Columns(1).Value2
arrTableDate = [tableDate].Columns(1).Resize(, 2).Value2

' сбор рабочих дней в словарь
For i = LBound(arrTableJobDay) To UBound(arrTableJobDay)
    dictJobDay(arrTableJobDay(i, 1)) = dictJobDay(arrTableJobDay(i, 1)) + dictJobDay.Count
Next i

ReDim arrResults(LBound(arrTableDate) To UBound(arrTableDate), 1 To 2)
'Проходимся по каждой дате
For i = LBound(arrTableDate) To UBound(arrTableDate)
    TimeAndStatusTmp = TimeInWork(arrTableDate(i, 1), arrTableDate(i, 2))                            'Вызов обработчик
    arrResults(i, 1) = TimeAndStatusTmp                                                              'Записываем Время в работе
Next i

dictJobDay.RemoveAll                                                                                 'Очищаем словарь
dictDateLog.RemoveAll                                                                                'Очищаем словарь
[tableDate].Columns(3).Resize(, 2) = arrResults                                                      'Выписываем результат

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

' Проверка день в день
Select Case True
    Case Int_LDate = Int_UDate
        If Not dictJobDay.Exists(Int_LDate) Then TimeInWork = 0: Exit Function                                                                      'Проверка день в день = выходной
        TimeInWork = (CheckStatusTime(TimeReg, TimeComp, 1) * 24)                                                                                   'Расчёт день в день, время
'Остальные паттерны
    Case Else
        StatusJobDay = CheckJobDay(Int_LDate, Int_UDate)                                                                                            'Получаем статус дат из функции
        If dictDateLog.Exists(Int_LDate & "|" & Int_UDate) Then tmpDays = dictDateLog(Int_LDate & "|" & Int_UDate): tmpStopTmpDay = True            'Проверяем даты в словаре, если есть - исключаем поиск рабочих дней
        Select Case StatusJobDay
            Case Is = 1
                Select Case tmpStopTmpDay: Case False: tmpDays = FindJobDay(Int_LDate, Int_UDate) * 9: End Select                                   'Есть ли в словаре даты
                TimeInWork = tmpDays                                                                                                                'Зарегистрировано и выполнено в выходные
            Case Is = 2
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate) - 2) * 9: End Select                             'Есть ли в словаре даты
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      'Зарегистрировано и выполнено в будние
            Case Is = 3
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate, 3) - 1) * 9: End Select                          'Есть ли в словаре даты
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      'Зарегистрировано в будние и выполнено в выходной
            Case Is = 4
                Select Case tmpStopTmpDay: Case False: tmpDays = (FindJobDay(Int_LDate, Int_UDate, 4) - 1) * 9: End Select                          'Есть ли в словаре даты
                TimeInWork = tmpDays + (CheckStatusTime(TimeReg, TimeComp, StatusJobDay) * 24)                                                      'Зарегистрировано в выходные, но выполнено в будние
        End Select
        If tmpStopTmpDay Then dictDateLog.Add Int_LDate & "|" & Int_UDate, tmpDays                                                                  'Добавление ключа даты рег|вып в словарь, для исключения поиска рабочих дней
End Select

End Function
Private Function CheckJobDay(ByRef Int_LDate As Long, ByRef Int_UDate As Long) As Integer
Select Case True
    Case Not dictJobDay.Exists(Int_LDate) And Not dictJobDay.Exists(Int_UDate): CheckJobDay = 1         'Зарегистрировано и выполнено в выходные
    Case dictJobDay.Exists(Int_LDate) And dictJobDay.Exists(Int_UDate): CheckJobDay = 2                 'Зарегистрировано и выполнено в будние
    Case dictJobDay.Exists(Int_LDate) And Not dictJobDay.Exists(Int_UDate): CheckJobDay = 3             'Зарегистрировано в будние и выполнено в выходной
    Case Not dictJobDay.Exists(Int_LDate) And dictJobDay.Exists(Int_UDate): CheckJobDay = 4             'Зарегистрировано в выходные, но выполнено в будние
End Select
End Function
Private Function FindJobDay(Int_LDate As Long, Int_UDate As Long, Optional ByRef JobDay As Integer) As Double

'Вычитаем из даты выполнения, ищем рабочий день
While Not dictJobDay.Exists(Int_UDate) And Int_UDate > Int_LDate
    Int_UDate = Int_UDate - 1
Wend

'Вычитаем из даты регистрации, ищем рабочий день
While Not dictJobDay.Exists(Int_LDate) And Int_LDate < Int_UDate
    Int_LDate = Int_LDate + 1
Wend

'Проверки
Select Case True
    Case Int_LDate = Int_UDate:
        Select Case JobDay
            Case 3, 4: FindJobDay = dictJobDay(Int_UDate) - dictJobDay(Int_LDate) + 1                   'Если статус дат 3, 4
            Case 0: FindJobDay = 0                                                                      'Если не передал статус дат
        End Select
    Case Else: FindJobDay = dictJobDay(Int_UDate) - (dictJobDay(Int_LDate) - 1)
End Select

End Function
Private Function CheckStatusTime(ByRef TimeReg As Double, ByRef TimeComp As Double, ByRef sameDayProcessed As Integer) As Double: Dim StatusReg As Integer, StatusComp As Integer
Select Case sameDayProcessed
    Case Is = 1, 2: StatusReg = CheckReg(TimeReg): StatusComp = CheckComp(TimeComp)                                      'Вызов обе функции
    Case Is = 4: StatusReg = 0: StatusComp = CheckComp(TimeComp)                                                         'Вызов функции только для выполнения
    Case Is = 3: StatusReg = CheckReg(TimeReg): StatusComp = 0                                                           'Вызов функции только для регистрации
End Select

Select Case True
    Case StatusReg = 1 And StatusComp = 1 And sameDayProcessed = 1: CheckStatusTime = TimeComp - TimeReg                 'Входит в диапазон с 9:00 до 18:00, день в день
    Case StatusReg = 1 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = ENDJOBDAY - TimeReg                'Входит в диапазон с 9:00 до 00:00, день в день
    Case StatusReg = 2 And StatusComp = 1 And sameDayProcessed = 1: CheckStatusTime = TimeComp - STARTJOBDAY             'Входит в диапазон с 00:00 до 18:00, день в день
    Case StatusReg = 2 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = STARTJOBDAY                        'Входит в диапазон с 00:00 до 00:00, день в день
    Case StatusReg = 2 And StatusComp = 2 And sameDayProcessed = 1: CheckStatusTime = 0                                  'Зарегистрировано и выполнено до 09:00, день в день
    Case StatusReg = 3 And StatusComp = 3 And sameDayProcessed = 1: CheckStatusTime = 0                                  'Зарегистрировано и выполнено после 18:00, день в день
    Case sameDayProcessed = 2, 3, 4:                                                                                     'Если статус даты не "Выходной"
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
    Case STARTJOBDAY To ENDJOBDAY: CheckReg = 1             'в диапазоне с 9:00 до 18:00
    Case Is < STARTJOBDAY: CheckReg = 2                     'меньше 9:00
    Case Is > ENDJOBDAY: CheckReg = 3                       'больше 18:00
End Select
End Function
Private Function CheckComp(ByRef TimeComp As Double) As Integer
Select Case TimeComp
    Case STARTJOBDAY To ENDJOBDAY: CheckComp = 1            'в диапазоне с 9:00 до 18:00
    Case Is < STARTJOBDAY: CheckComp = 2                    'меньше 9:00
    Case Is > ENDJOBDAY: CheckComp = 3                      'больше 18:00
End Select
End Function
Private Function DisabledApps(OnOff As Boolean, CalculateApp As Boolean): If CalculateApp = False Then CalculateApp = xlManual Else CalculateApp = xlAutomatic
'отключаем application
Application.ScreenUpdating = OnOff
Application.Calculation = CalculateApp
Application.AskToUpdateLinks = OnOff
Application.DisplayAlerts = OnOff
End Function
Private Function FncTime(ByRef Infinity As Double, ByVal Output As Boolean)
'расчитываем время выполнения макроса
If Not Output Then Infinity = Timer
If Output Then MsgBox "Готово! " & Format(Timer - Infinity, "0.00 сек"), vbInformation
End Function
