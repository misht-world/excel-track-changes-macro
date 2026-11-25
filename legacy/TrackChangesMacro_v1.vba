	Dim previousValues As Object
	Dim initialValues As Object
	Dim initialFormulas As Object
	Dim disableEvents As Boolean
	
	Private Sub Workbook_Open()
	    ' Инициализируем словари для хранения предыдущих и начальных значений и формул
	    Set previousValues = CreateObject("Scripting.Dictionary")
	    Set initialValues = CreateObject("Scripting.Dictionary")
	    Set initialFormulas = CreateObject("Scripting.Dictionary")
	    ' Сохраняем начальные значения и формулы всех ячеек
	    SaveInitialData
	    SavePreviousValues
	    disableEvents = False
	End Sub
	
	Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
	    If disableEvents Then Exit Sub
	    ' Отключаем события, чтобы избежать рекурсивных вызовов
	    disableEvents = True
	
	    Dim cell As Range
	    For Each cell In Target
	        Dim cellAddress As String
	        cellAddress = Sh.Name & "!" & cell.Address
	
	        Dim initialValue As Variant
	        Dim initialFormula As Variant
	        Dim currentValue As Variant
	        Dim currentFormula As Variant
	
	        currentValue = cell.Value
	        If cell.HasFormula Then
	            currentFormula = cell.Formula
	        Else
	            currentFormula = ""
	        End If
	
	        ' Проверяем, существует ли ячейка в начальных значениях
	        If initialValues.Exists(cellAddress) Then
	            initialValue = initialValues(cellAddress)
	            initialFormula = initialFormulas(cellAddress)
	        Else
	            ' Если ячейка не была сохранена в начальных значениях, считаем, что она была пустой
	            initialValue = ""
	            initialFormula = ""
	            initialValues.Add cellAddress, initialValue
	            initialFormulas.Add cellAddress, initialFormula
	        End If
	
	        ' Преобразуем значения в строки для корректного сравнения
	        Dim currValStr As String
	        Dim initValStr As String
	
	        currValStr = CStr(currentValue)
	        initValStr = CStr(initialValue)
	
	        ' Проверяем, изменилось ли значение по сравнению с начальными данными
	        If currValStr <> initValStr Then
	            ' Значение отличается от начального, меняем цвет текста на синий
	            cell.Font.Color = RGB(0, 0, 255)
	        Else
	            ' Значение совпадает с начальными данными, меняем цвет текста на чёрный
	            cell.Font.Color = RGB(0, 0, 0)
	        End If
	
	        ' Обновляем предыдущие значения и формулы
	        previousValues(cellAddress) = currentValue
	        previousValues("Formula_" & cellAddress) = currentFormula
	
	    Next cell
	
	    ' Включаем события обратно
	    disableEvents = False
	End Sub
	
	Private Sub Workbook_SheetCalculate(ByVal Sh As Object)
	    If disableEvents Then Exit Sub
	    ' Отключаем события, чтобы избежать рекурсивных вызовов
	    disableEvents = True
	    Dim ws As Worksheet
	    Dim cell As Range
	
	    ' Перебираем все листы и ячейки для сравнения значений и формул
	    For Each ws In ThisWorkbook.Sheets
	        For Each cell In ws.UsedRange
	            Dim cellAddress As String
	            cellAddress = ws.Name & "!" & cell.Address
	
	            Dim initialValue As Variant
	            Dim initialFormula As Variant
	            Dim previousValue As Variant
	            Dim previousFormula As Variant
	            Dim currentValue As Variant
	            Dim currentFormula As Variant
	
	            currentValue = cell.Value
	            If cell.HasFormula Then
	                currentFormula = cell.Formula
	            Else
	                currentFormula = ""
	            End If
	
	            ' Проверяем, существует ли ячейка в начальных значениях
	            If initialValues.Exists(cellAddress) Then
	                initialValue = initialValues(cellAddress)
	                initialFormula = initialFormulas(cellAddress)
	            Else
	                ' Если ячейка не была сохранена в начальных значениях, считаем, что она была пустой
	                initialValue = ""
	                initialFormula = ""
	                initialValues.Add cellAddress, initialValue
	                initialFormulas.Add cellAddress, initialFormula
	            End If
	
	            ' Получаем предыдущие значения и формулы
	            If previousValues.Exists(cellAddress) Then
	                previousValue = previousValues(cellAddress)
	            Else
	                previousValue = initialValue
	            End If
	
	            If previousValues.Exists("Formula_" & cellAddress) Then
	                previousFormula = previousValues("Formula_" & cellAddress)
	            Else
	                previousFormula = initialFormula
	            End If
	
	            ' Преобразуем значения и формулы в строки для корректного сравнения
	            Dim prevValStr As String
	            Dim currValStr As String
	            Dim initValStr As String
	            Dim prevFormulaStr As String
	            Dim currFormulaStr As String
	
	            prevValStr = CStr(previousValue)
	            currValStr = CStr(currentValue)
	            initValStr = CStr(initialValue)
	            prevFormulaStr = CStr(previousFormula)
	            currFormulaStr = CStr(currentFormula)
	
	            ' Проверяем, изменилось ли значение или формула по сравнению с предыдущими
	            If prevValStr <> currValStr Or prevFormulaStr <> currFormulaStr Then
	                ' Проверяем, равно ли текущее значение начальному значению
	                If currValStr = initValStr Then
	                    ' Значение вернулось к начальному, меняем цвет текста на чёрный
	                    cell.Font.Color = RGB(0, 0, 0)
	                Else
	                    ' Значение отличается от начального, меняем цвет текста на синий
	                    cell.Font.Color = RGB(0, 0, 255)
	                End If
	                ' Обновляем предыдущие значения и формулы
	                previousValues(cellAddress) = currentValue
	                previousValues("Formula_" & cellAddress) = currentFormula
	            End If
	        Next cell
	    Next ws
	    ' Включаем события обратно
	    disableEvents = False
	End Sub
	
	Private Sub SaveInitialData()
	    ' Сохраняем начальные значения и формулы всех ячеек
	    Dim cell As Range
	    Dim ws As Worksheet
	    For Each ws In ThisWorkbook.Sheets
	        For Each cell In ws.UsedRange
	            Dim cellAddress As String
	            cellAddress = ws.Name & "!" & cell.Address
	            initialValues(cellAddress) = cell.Value
	            If cell.HasFormula Then
	                initialFormulas(cellAddress) = cell.Formula
	            Else
	                initialFormulas(cellAddress) = ""
	            End If
	        Next cell
	    Next ws
	End Sub
	
	Private Sub SavePreviousValues()
	    ' Обновляем текущие значения и формулы всех ячеек
	    Dim cell As Range
	    Dim ws As Worksheet
	    For Each ws In ThisWorkbook.Sheets
	        For Each cell In ws.UsedRange
	            Dim cellAddress As String
	            cellAddress = ws.Name & "!" & cell.Address
	            previousValues(cellAddress) = cell.Value
	            If cell.HasFormula Then
	                previousValues("Formula_" & cellAddress) = cell.Formula
	            Else
	                previousValues("Formula_" & cellAddress) = ""
	            End If
	        Next cell
	    Next ws
End Sub