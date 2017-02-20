Function DuplicateRow_ObjectType(TestObject as Variant, Optional TestCase As String)
    test_object = TestObject
    Select Case TRUE
      Case IsMissing(TestCase)
          If TypeName(test_object) <> "Object" Then
              DuplicateRow_ObjectType = TypeName(test_object)
          Else
              DuplicateRow_ObjectType = DuplicateRow_ObjectType(TestObject,"ImpossibleObject")
          End If      
      Case TestCase = "ImpossibleObject"
          On Local Error Goto LABELDUPLICATEROWOBJECTTESTFAIL0
          error_trigger = test_object.Row1
          Goto LABELDUPLICATEROWOBJECTTESTSUCCESS
          LABELDUPLICATEROWOBJECTTESTFAIL0:
          DuplicateRow_ObjectType =  DuplicateRow_ObjectType(TestObject,"Cell")
      Case TestCase = "Cell"
          On Local Error Goto LABELDUPLICATEROWOBJECTTESTFAIL1
          error_trigger = test_object.CellAddress.Row
          error_trigger = test_object.CellAddress.Column
          Goto LABELDUPLICATEROWOBJECTTESTSUCCESS
          LABELDUPLICATEROWOBJECTTESTFAIL1:
          DuplicateRow_ObjectType =  DuplicateRow_ObjectType(TestObject,"Unknown")
      Case Else
          LABELDUPLICATEROWOBJECTTESTSUCCESS:
          DuplicateRow_ObjectType = TestCase
    End Select
End Function