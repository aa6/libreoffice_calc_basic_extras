Function DuplicateRow_ObjectType(TestObject as Variant, Optional TestCase As String)
  If TypeName(TestObject) <> "Object" Then
    DuplicateRow_ObjectType = TypeName(TestObject)
    Exit Function
  End If
  If IsMissing(TestCase) Then
    TestCase = "ImpossibleObject"
  End If
  Select Case TestCase
    Case "ImpossibleObject"
      On Local Error Goto Label_DuplicateRow_ObjectTest_StrangeShit
      error_trigger = TestObject.Row1
      DuplicateRow_ObjectType = TestCase
      Exit Function
      Label_DuplicateRow_ObjectTest_StrangeShit:
      DuplicateRow_ObjectType =  DuplicateRow_ObjectType(TestObject,"Cell")
    Case "Cell"
      On Local Error Goto Label_DuplicateRow_ObjectTest_1
      error_trigger = TestObject.CellAddress.Row
      error_trigger = TestObject.CellAddress.Column
      DuplicateRow_ObjectType = TestCase
      Exit Function
      Label_DuplicateRow_ObjectTest_1:
      DuplicateRow_ObjectType =  DuplicateRow_ObjectType(TestObject,"Unknown")
    Case Else
      DuplicateRow_ObjectType = TestCase
  End Select

End Function