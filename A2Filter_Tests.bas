Option Explicit

Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
  'this method runs once per module.
  Set Assert = CreateObject("Rubberduck.AssertClass")
  Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
  'this method runs once per module.
  Set Assert = Nothing
  Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
  'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
  'this method runs after every test in the module.
End Sub


'@TestMethod
Public Sub A2_Filter_AND_TestMethod()
  '   On Error GoTo TestFail
  Dim varReturn() As Variant
  Dim a2_Data() As Variant
  Dim a2_Crit() As Variant
  
  ' Пока работает только с одиночными критерями
  ' ToDo: Переделать тесты для многих
  ' в массиве данных 1 строка, в критериях 1 условие
  ReDim a2_Data(1 To 1, 1 To 1)
  a2_Data(1, 1) = "11"

  ReDim a2_Crit(1 To 2, 1 To 3)
  a2_Crit(1, 1) = 1 ' столбец
  a2_Crit(1, 2) = a2_Data(1, 1) ' критерий
  a2_Crit(1, 3) = "EQUAL_TEXT" ' метод фильтрации
  
  varReturn = A2_Filter_AND(a2_Data(), a2_Crit())
  If varReturn(1, 1) <> a2_Data(1, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"

  a2_Crit(1, 2) = "0" ' критерий
  varReturn = A2_Filter_AND(a2_Data(), a2_Crit())
  If isArray_Bound(varReturn) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"

  ' в массиве данных 2 строки, в критериях 2 условия
  ReDim a2_Data(1 To 2, 1 To 1)
  a2_Data(1, 1) = "11"
  a2_Data(2, 1) = "21"

  ReDim a2_Crit(1 To 2, 1 To 3)
  a2_Crit(1, 1) = 1 ' столбец
  a2_Crit(1, 2) = a2_Data(1, 1) ' критерий
  a2_Crit(1, 3) = "EQUAL_TEXT" ' метод фильтрации
  a2_Crit(2, 1) = 1 ' столбец
  a2_Crit(2, 2) = a2_Data(2, 1) ' критерий
  a2_Crit(2, 3) = "EQUAL_TEXT" ' метод фильтрации
  
  varReturn = A2_Filter_AND(a2_Data(), a2_Crit())
  If varReturn(1, 1) <> a2_Data(1, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"
  If varReturn(2, 1) <> a2_Data(2, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"

  ' 3 строки данных
  ReDim a2_Data(1 To 3, 1 To 3)
  a2_Data(1, 1) = "11": a2_Data(1, 2) = "12": a2_Data(1, 3) = "13"
  a2_Data(2, 1) = "21": a2_Data(2, 2) = "22": a2_Data(2, 3) = "23"
  a2_Data(3, 1) = "31": a2_Data(3, 2) = "32": a2_Data(3, 3) = "33"
   
  ReDim a2_Crit(1 To 2, 1 To 3)
  a2_Crit(1, 1) = 1 ' столбец
  a2_Crit(1, 2) = a2_Data(3, 1) ' критерий
  a2_Crit(1, 3) = "EQUAL_TEXT" ' метод фильтрации
   
  varReturn = A2_Filter_AND(a2_Data(), a2_Crit())
  If varReturn(1, 1) <> a2_Data(3, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"
   
  a2_Crit(2, 1) = 1 ' столбец
  a2_Crit(2, 2) = a2_Data(1, 1) ' критерий
  a2_Crit(2, 3) = "EQUAL_TEXT" ' метод фильтрации
   
  varReturn = A2_Filter_AND(a2_Data(), a2_Crit())
  If varReturn(1, 1) <> a2_Data(1, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"
  If varReturn(2, 1) <> a2_Data(3, 1) Then Err.Raise 567, "A2_Filter_AND(a2_Data(),a2_Crit())"
   
TestExit:
  Exit Sub
TestFail:
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub Collection_A2_Rows_Numbers_All_TestMethod()
  On Error GoTo TestFail
  Dim a2() As Variant
  a2 = Mock.G_a2
  Dim varReturn As Collection
  Set varReturn = Collection_A2_Rows_Numbers_All(a2())
  If varReturn.count <> (UBound(a2) - LBound(a2) + 1) Then Err.Raise 567, "Collection_A2_Rows_Numbers_All(a2())"
TestExit:
  Mock.wb.Close False
  Exit Sub
TestFail:
  Mock.wb.Close False
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub Collection_Rows_Copy_TestMethod()
  On Error GoTo TestFail
  Dim a2_Data() As Variant
  Dim a2_Crit() As Variant

  a2_Data = Mock.G_a2
  a2_Crit = Mock.G_a2
  Dim varReturn As Collection
  Set varReturn = Collection_Rows_Copy(a2_Data(), a2_Crit())
  
  'if varReturn <> 0 Then Err.Raise 567, "Collection_Rows_Copy(a2_Data(),a2_Crit())"
TestExit:
  Mock.wb.Close False
  Exit Sub
TestFail:
  Mock.wb.Close False
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub Criteria_Check_TestMethod()
  On Error GoTo TestFail
  Dim varReturn As Boolean
  Dim var As Variant
  Dim a2_Crit() As Variant
  Dim row_Crit_ As Long
  'Массив критериев:
  'Столбец номер, Критерий, Метод фильтации с указанием учета регистра
  ReDim a2_Crit(1 To 2, 1 To 3)
  row_Crit_ = 2
  '  a2_Crit(row_Crit_, 1) = 9 ' в этой процедуре не нужен
  
  var = "z"
  a2_Crit(row_Crit_, 2) = var
  a2_Crit(row_Crit_, 3) = "EQUAL_TEXT" ' сравнить с учётом регистра
  
  varReturn = Criteria_Check(var, a2_Crit(), row_Crit_)
  If varReturn = False Then Err.Raise 567, "Criteria_Check(var,a2_Crit(),row_Crit_)"
  
  var = "z"
  a2_Crit(row_Crit_, 2) = "w"
  a2_Crit(row_Crit_, 3) = "EQUAL_TEXT" ' сравнить с учётом регистра
  
  varReturn = Criteria_Check(var, a2_Crit(), row_Crit_)
  If varReturn Then Err.Raise 567, "Criteria_Check(var,a2_Crit(),row_Crit_)"
  
TestExit:
  Exit Sub
TestFail:
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub A2_Copy_Rows_Collection_TestMethod()
  On Error GoTo TestFail
  Dim a2_Sour() As Variant
  Dim coll_Rows As New Collection
  Dim varReturn() As Variant

  ReDim a2_Sour(1 To 3, 1 To 2)
  a2_Sour(1, 1) = 11: a2_Sour(1, 2) = 12
  a2_Sour(3, 1) = 31: a2_Sour(3, 2) = 32

  coll_Rows.Add 1, CStr(1)
  coll_Rows.Add 3, CStr(2)

  varReturn = A2_Copy_Rows_Collection(a2_Sour(), coll_Rows)
  If UBound(varReturn) <> coll_Rows.count Then Err.Raise 567, "A2_Copy_Rows_Collection(a2_Sour(),coll_Rows)"
  If varReturn(2, 2) <> 32 Then Err.Raise 567, "A2_Copy_Rows_Collection(a2_Sour(),coll_Rows)"

TestExit:
  Exit Sub
TestFail:
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub Row_Meets_CriteriaS_AND_TestMethod()
'  On Error GoTo TestFail
  Dim a2_Data() As Variant
  Dim row_Data As Long
  Dim a2_Crit() As Variant
  ReDim a2_Data(1 To 3, 1 To 3)
  a2_Data(1, 1) = "11"
  a2_Data(1, 2) = "12"
  a2_Data(1, 3) = "13"
  a2_Data(2, 1) = "21"
  a2_Data(2, 2) = "22"
  a2_Data(2, 3) = "23"
  a2_Data(3, 1) = "31"
  a2_Data(3, 2) = "32"
  a2_Data(3, 3) = "33"

  row_Data = 3

  ReDim a2_Crit(1 To 1, 1 To 3)
  a2_Crit(1, 1) = 1 ' столбец
  a2_Crit(1, 2) = "31" ' критерий
  a2_Crit(1, 3) = "EQUAL_TEXT" ' метод фильтрации
  Dim varReturn As Boolean
  varReturn = Row_Meets_CriteriaS_AND(a2_Data(), row_Data, a2_Crit())
  If varReturn = False Then Err.Raise 567, "Row_Meets_CriteriaS_AND(a2_Data(),row_Data,a2_Crit())"

TestExit:
  Exit Sub
TestFail:
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub isArray_Bound_TestMethod()
 
  On Error GoTo TestFail
 
  Dim a2() As Variant
 
  Dim varReturn As Boolean
 
  varReturn = isArray_Bound(a2)
 
  If varReturn Then Err.Raise 567, "isArray_Bound(a2)"
 
  ReDim a2(1 To 1)
 
  varReturn = isArray_Bound(a2)
 
  If varReturn = False Then Err.Raise 567, "isArray_Bound(a2)"
 
TestExit:
 
  Exit Sub
 
TestFail:
 
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
 
End Sub


'@TestMethod
Public Sub String_Add_Symbols_TestMethod()
   On Error GoTo TestFail
   Dim sVal As String
   Dim length As Long
   Dim symb As String
   sVal = Mock.G_String
   length = Len(sVal) + 9
   symb = Mock.G_String
   Dim varReturn As String
   varReturn = String_Add_Symbols(sVal, length, symb)
   If Len(varReturn) <> Len(sVal) + 9 Then Err.Raise 567, "String_Add_Symbols(sVal,length,symb)"
   
   sVal = "ZZ"
   length = Len(sVal) + 3
   symb = Mock.G_Symb_Rand_ASCII_Range(1, 255)
   varReturn = String_Add_Symbols(sVal, length, symb)
   If Mid(varReturn, 3, 1) <> symb Then Err.Raise 567, "String_Add_Symbols(sVal,length,symb)"
TestExit:
   Mock.wb.Close False
   Exit Sub
TestFail:
   Mock.wb.Close False
   Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub


'@TestMethod
Public Sub Element_Meet_CriteriaS_TestMethod()
  On Error GoTo TestFail
  Dim varReturn As Boolean
  Dim element As Variant
  Dim column As Long
  Dim a2_Crit() As Variant
  element = Mock.G_Variant
  a2_Crit = A2_Crit_Booking
  column = a2_Crit(1, 1)
  varReturn = Element_Meet_CriteriaS(element, column, a2_Crit())
  If varReturn Then Err.Raise 567, "Element_Meet_CriteriaS(element,column,a2_Crit())"

  element = Mock.G_Variant
  a2_Crit(1, 2) = element
  varReturn = Element_Meet_CriteriaS(element, column, a2_Crit())
  If varReturn = False Then Err.Raise 567, "Element_Meet_CriteriaS(element,column,a2_Crit())"

TestExit:
  Mock.wb.Close False
  Exit Sub
TestFail:
  Mock.wb.Close False
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub



'@TestMethod
Public Sub Column_in_Criterias_TestMethod()
  On Error GoTo TestFail
  Dim varReturn As Boolean
  Dim column As Long
  Dim a2_Crit() As Variant
  ReDim a2_Crit(1 To 2, 1 To 1)
  
  column = 1
  varReturn = Column_in_Criterias(column, a2_Crit())
  If varReturn Then Err.Raise 567, "Column_in_Criterias(column,a2_Crit())"

  a2_Crit(1, 1) = 1
  varReturn = Column_in_Criterias(column, a2_Crit())
  If varReturn = False Then Err.Raise 567, "Column_in_Criterias(column,a2_Crit())"

TestExit:
  Exit Sub
TestFail:
  Assert.Fail "Test error: #" & Err.Number & " - " & Err.Description
End Sub

