Attribute VB_Name = "Table_Operation"
Option Explicit

Sub ImportData(firstLine As Integer, firstColumn As Integer)

  Dim Sql As String, sqlFields As String, tableName As String
  Dim line As Integer, column As Integer

  Application.ScreenUpdating = False

  'Get Table Name
  tableName = Cells(FindLine("Table Name", 1), 2)

  'Get Import fields
  sqlFields = SqlImportFields(ArrayLine(FindLine("Import Data", 1) + 1))

  Sql = SqlSelectQuery(tableName, sqlFields)

  RunImportSql Sql, firstLine, firstColumn

  Application.ScreenUpdating = True
End Sub

Sub RunImportSql(sqlQuery As String, firstLine As Integer, firstColumn As Integer)

  Dim rs As ADODB.Recordset
  Dim line As Integer, column As Integer

  Set rs = New ADODB.Recordset

  ConnectProductionServer
  rs.Open sqlQuery, oConn

    line = 0
      Do Until rs.EOF

          For column = 0 To rs.Fields.Count -1
              Cells(line + firstLine, column + firstColumn) = rs.Fields(column).Value

          Next
          line = line + 1
          rs.MoveNext
      Loop

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

End Sub

Function SqlSelectQuery(tableName As String, sqlFields As String) As String 'add filter

  SqlSelectQuery = "SELECT " & sqlFields & " FROM " & tableName

End Function

Function FindLine(lookupValue As String, column As Integer) As Integer
    Dim lookupLine As Integer, cellValue As String

    lookupLine = 1
    cellValue = Cells(lookupLine, column).Value

    While cellValue <> lookupValue
      lookupLine = lookupLine + column
      cellValue = Cells(lookupLine, column).Value
    Wend

    FindLine = lookupLine
End Function

Sub testcode()
 Dim testArray() As String
  testArray = ArrayLine(3)
  msgbox SqlImportFields(testArray)

End Sub

Function ArrayLine(line As Integer) As String()
  Dim iColumn As Integer, sArray() As String

  'Dimentionate the array
  iColumn = 1
  While Cells(line, iColumn) <> ""
    ReDim sArray(iColumn)
    iColumn = iColumn + 1
  Wend

  'Populate the array
  iColumn = 1
  While Cells(line, iColumn) <> ""
    sArray(iColumn) = Cells(line, iColumn)
    iColumn = iColumn + 1
  Wend

  ArrayLine = sArray
End Function

function SizeArray (vArray () as variant) as Integer
  dim arrayItem as variant, iArray as Integer

  iArray =0
  for each arrayItem in vArray
    iArray = iarray + 1
  Next arrayItem

  SizeArray = iArray
end function

Function SqlImportFields(sArray() As String) As String
  '''Reformat an array into fields statements for SQL queries
  Dim sqlField As Variant, sFields As String
  sFields = ""

  For Each sqlField In sArray
    sFields = sFields & sqlField & ","
  Next sqlField

  sFields = Mid(sFields, 2)
  sFields = Left(sFields, Len(sFields) - 1)

  SqlImportFields = sFields
End Function

function sqlImportFilters(sArrayField as string, sArrayValue as string) as String
  '''Reformat two arrays into filters statements for SQL queries (WHERE)
  Dim sqlField As Variant, sFields As String, sValue as string
  Dim dimArray as Integer

  sFields = ""

  dimArray = 0
  For Each sqlField In sArrayField
    sFields = sFields & sqlField & ","
  Next sqlField

sFields = Mid(sFields, 2)
sFields = Left(sFields, Len(sFields) - 1)

SqlImportFields = sFields

end function

Sub UpdateData(firstLine As Integer, firstColumn As Integer)

  Dim Sql As String, sqlFields As String, tableName As String, countSql As Integer
  Dim maxField As Integer
  Dim tableLine As Integer

  '_____________________________________

  'Creation de la requete SQL UPDATE
  '_____________________________________


  ConnectProductionServer

      tableName = Cells(1, 2)
      maxField = 1
      tableLine = firstLine

      While Cells(tableLine, firstColumn) <> ""

          Cells(5, 1) = Cells(tableLine, firstColumn)

          maxField = 2
          countSql = 1
          sqlFields = ""
          While Cells(3, maxField) <> ""

                      If countSql > 1 Then

                          sqlFields = sqlFields & " , "
                      End If

                      If Cells(5, maxField) <> "" Then

                          sqlFields = sqlFields & Cells(3, maxField) & " = '" & Cells(5, maxField) & "'"
                          Else
                          sqlFields = sqlFields & Cells(3, maxField) & " = NULL"
                      End If

              maxField = maxField + 1
              countSql = countSql + 1
          Wend
          'Timestamp ????

          Sql = "UPDATE " & tableName & " SET " & sqlFields & " WHERE " & Cells(3, 1) & " = " & Cells(5, 1)
          'MsgBox Sql
          oConn.Execute Sql
          tableLine = tableLine + 1

      Wend
  'MsgBox Sql
      oConn.Close
      Set oConn = Nothing

End Sub
