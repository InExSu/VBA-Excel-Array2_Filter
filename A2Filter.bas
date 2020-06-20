Attribute VB_Name = "A2Filter"
Option Explicit

' ������ ���������� � 3 ������� ���������
Public Const EQUAL_TEXT As String = "EQUAL_TEXT" ' ������ ����� ��� �����  ��������
Public Const EQUAL_BINA As String = "EQUAL_BINA" ' ������ ����� �   ������ ��������
Public Const CONTA_TEXT As String = "CONTA_TEXT" ' ������ �������� ��� �����  ��������
Public Const CONTA_BINA As String = "CONTA_BINA" ' ������ �������� �   ������ ��������
' ToDo: ������� ��� ��� � �����


Sub A2_Filter_AND_FW_Live()
   Dim _
      a2_Data() As Variant, _
      a2_Crit() As Variant
   
   a2_Data = ActiveSheet.UsedRange.Value
   
   ReDim a2_Crit(1 To 2, 1 To 3)
   '   a2_Crit(1, 1) = 1
   '   a2_Crit(1, 2) = "Code"
   '   a2_Crit(1, 3) = "EQUAL_TEXT"
   '   a2_Crit(2, 1) = 1
   '   a2_Crit(2, 2) = "Date"
   '   a2_Crit(2, 3) = "EQUAL_TEXT"
   
   a2_Crit(1, 1) = 3
   a2_Crit(1, 2) = "Work"
   a2_Crit(1, 3) = "CONTA_TEXT"
   
   a2_Data = A2_Filter_AND(a2_Data, a2_Crit)
   
 MsgBox A2_2_String(a2_Data)
   
End Sub


Function A2_Filter_AND( _
   a2_Data() As Variant, _
   a2_Crit() As Variant) _
   As Variant()
   ' test yes
   ' ���������� ������� ���������� �� ������ ��������, �� ���������� ���������
   ' �������� � - ������ �������� �������� �� ���� ��������� ��������
  
   ' ToDo: ���������� ������� A2_Filter_OR - ����������� �� ��������� ��� -
   ' ��� ������� ������� ��������� �������� �������� (����� �����)
  
   '������ ���������:
   '������� �����, ��������, ����� ��������� � ��������� ����� ��������

   ' ������ ���������� ������ � Criteria_Check

   A2_Filter_AND = A2_Copy_Rows_Collection( _
      a2_Data, _
      Collection_Rows_Copy( _
      a2_Data, a2_Crit))

End Function


Function Collection_Rows_Copy( _
   a2_Data() As Variant, _
   a2_Crit() As Variant) _
   As Collection
   ' ��������� ������� �����, ������� ����� ����������
   '������ ���������:
   '������� �����, ��������, ����� ��������� � ��������� ����� ��������

   Dim _
      row_Data As Long, _
      coll_Copy As New Collection
  
   ' ������ �� ������� ������� � �������
   For row_Data = LBound(a2_Data) To UBound(a2_Data)
      
      If Row_Meets_Criteria_AND( _
         a2_Data, row_Data, _
         a2_Crit) Then
          
         coll_Copy.Add row_Data
          
      End If
   Next row_Data
    
   Set Collection_Rows_Copy = coll_Copy
    
End Function


Function Row_Meets_Criteria_AND( _
   a2_Data() As Variant, _
   row_Data As Long, _
   a2_Crit() As Variant) _
   As Boolean
   ' ������������� �� ������ ������ � ������� ���������
  
   '������ ���������:
   '������� �����, ��������, ����� ��������� � ��������� ����� ��������

   Dim _
      row_Crit As Long, _
      col_Data As Long, _
      bingo  As Boolean
  
   For row_Crit = LBound(a2_Crit) To UBound(a2_Crit)
      For col_Data = LBound(a2_Data, 2) To UBound(a2_Data, 2)
         
         '  ���� ������� ��������� � ������ ��������
         If col_Data = a2_Crit(row_Crit, 1) Then
            bingo = True
            
            ' ���� ������� ������� �� �������� ��������
            If Criteria_Check( _
               a2_Data(row_Data, col_Data), _
               a2_Crit, row_Crit) = False Then
               
               bingo = False
               Exit For

            End If
         End If
      Next col_Data
      
      '���� ������ �������� ������� - ������ ������ ������ �������
      If bingo Then Exit For
   
   Next row_Crit
  
   Row_Meets_Criteria_AND = bingo
  
End Function


Function Criteria_Check( _
   var_Desti As Variant, _
   a2_Crit() As Variant, _
   row_Crite As Long) _
   As Boolean
   ' test yes
   ' ��������� �� ������� � ����������

   '������ ���������:
   '������� �����, ��������, ����� ��������� � ��������� ����� ��������
  
   Dim _
      bingo As Boolean, _
      vCrit As Variant
  
   vCrit = a2_Crit(row_Crite, 2)
  
   '   Debug.Assert vCrit <> "11"
   '   Debug.Assert vCrit <> "31"
  
   Select Case LCase$(a2_Crit(row_Crite, 3))
      
      Case LCase$(EQUAL_TEXT)
         If LCase$(var_Desti) = LCase$(vCrit) Then
            bingo = True
         End If
     
      Case LCase$(EQUAL_BINA)
         If var_Desti = vCrit Then
            bingo = True
         End If
      
      Case LCase$(CONTA_TEXT)
         If InStr(1, var_Desti, vCrit, vbTextCompare) > 0 Then
            bingo = True
         End If
      
      Case LCase$(CONTA_BINA)
         If InStr(var_Desti, vCrit) > 0 Then
            bingo = True
         End If
   
   End Select
  
   Criteria_Check = bingo
  
End Function


Function Collection_A2_Rows_Numbers_All( _
   a2() As Variant) _
   As Collection
  
   ' test yes
   ' ������� ��������� ������� ����� �������
  
   Dim _
      colle As New Collection, _
      count As Long

   For count = LBound(a2) To UBound(a2)
      colle.Add count
   Next count

   Set Collection_A2_Rows_Numbers_All = colle

End Function


Function Option_Compare() _
   As String
   ' ������� �������� ������ ��������� �����
   If "z" = "Z" Then
      Option_Compare = "Text"
   Else
      Option_Compare = "Binary" ' default text comparison method.
   End If
End Function


Function A2_Copy_Rows_Collection( _
   a2_Sour() As Variant, _
   coll_Rows As Collection) _
   As Variant()
  
   ' test yes
   ' ���������� ������ ������� � ����� ������, �� ������� ���������
   
   Dim a2_Dest() As Variant
  
   If coll_Rows.count > 0 Then
      
      ReDim a2_Dest( _
         LBound(a2_Sour) To LBound(a2_Sour) + coll_Rows.count - 1, _
         LBound(a2_Sour, 2) To UBound(a2_Sour, 2))

      Dim row As Long

      For row = LBound(a2_Dest) To UBound(a2_Dest)
  
         A2_Row_Copy _
            a2_Sour, _
            coll_Rows.Item(row), _
            a2_Dest, _
            row

      Next
  
   End If
   
   A2_Copy_Rows_Collection = a2_Dest

End Function


Sub A2_Row_Copy( _
   a2_Sour() As Variant, _
   lRow_Sour As Long, _
   a2_Dest() As Variant, _
   lRow_Dest As Long)
   ' test yes
   ' ���������� ������ �� ������� � ������

   Dim lCol As Long

   For lCol = LBound(a2_Sour, 2) To UBound(a2_Sour, 2)

      a2_Dest(lRow_Dest, lCol) = _
         a2_Sour(lRow_Sour, lCol)

   Next
End Sub


Function isArray_Bound( _
   a2() As Variant) _
   As Boolean
   ' code test coverage
   ' �������� ������������� �������, isAx
 
   Dim var As Variant
 
   var = Not a2
 
   If var <> -1 Then isArray_Bound = True
 
End Function


Sub A2_2_String_test()
   ReDim a2(1 To 2, 1 To 2)
   a2(1, 1) = "123456789"
   a2(1, 2) = "123456789"
   a2(2, 1) = "12345678"
   a2(2, 2) = "1234567890"

   MsgBox A2_2_String(a2)

End Sub


Function A2_2_String( _
   a2() As Variant, _
   Optional el_Width As Long = 9, _
   Optional separato As String = " | ", _
   Optional s_Add As String = "_") _
   As String
   ' ������� ������ � ���� ������-�������

   Dim sVal As String, _
      sAdd As String

   Dim row As Long, col As Long
   For row = LBound(a2) To UBound(a2)
      For col = LBound(a2, 2) To UBound(a2, 2)
                                                
         sAdd = Left$(String_Add_Symbols(CStr(a2(row, col)), el_Width, s_Add), _
            el_Width)
         
         Debug.Assert Len(sAdd) = 9
         sVal = sVal & sAdd & separato
                                                                                                                  
      Next col
      
      sVal = sVal & vbNewLine
      
   Next row

   A2_2_String = sVal

End Function


Function String_Add_Symbols( _
   sVal As String, _
   length As Long, _
   Optional symb As String = " ") _
   As String
   ' ��������� ������ ��������� �� ������ �����
   
   Dim count As Long
   count = length - Len(sVal)

   If count > 0 Then
      String_Add_Symbols = sVal & String(count, symb)
   Else
      String_Add_Symbols = sVal
   End If

End Function
