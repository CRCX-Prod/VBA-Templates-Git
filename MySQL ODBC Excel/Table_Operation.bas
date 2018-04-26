Attribute VB_Name = "Table_Operation"
Option Explicit

Sub ImportData(firstLine As Integer, firstColumn As Integer)

  Dim rs As ADODB.Recordset
  Dim Sql As String, SqlFields As String, tableName As String
  Dim maxField As Integer, line As Integer, column As Integer
  Dim stringRange As Range

  Application.ScreenUpdating = False
  Set rs = New ADODB.Recordset

  '_____________________________________

  'Creation de la requete SQL
  '_____________________________________

  tableName = LookupTableName
  maxField = 1

  While Cells(9, maxField) <> ""
      If maxField = 1 Then   'field 1
          SqlFields = Cells(9, maxField)
      Else            'other fields
          SqlFields = SqlFields & "," & Cells(9, maxField)
      End If
  maxField = maxField + 1
  Wend

  Sql = "SELECT " & SqlFields & " FROM " & tableName
  '    MsgBox Sql
  '    Range("A12").Value = Sql

      ConnectProductionServer
      rs.Open Sql, oConn ', adOpenDynamic, adLockOptimistic

  '_____________________________________

  'Effacer les donnees existantes
  '_____________________________________





  '_____________________________________

  'Affichage des donnees
  '_____________________________________

  line = 0
      Do Until rs.EOF

          For column = 0 To maxField - 2
              Cells(line + firstLine, column + firstColumn) = rs.Fields(column).Value

          Next
          line = line + 1
          rs.MoveNext
      Loop

  Application.ScreenUpdating = True

  oConn.Close
  Set oConn = Nothing
  Set rs = Nothing

End Sub

sub testcode ()

  msgbox LookupTableName

end sub

Function LookupTableName() As String
  'look in A column : "Table Name",
  'Revert Column B
  LookupTableName = Cells(FindLine ("Table Name"), 2)
End Function

function FindLine (lookupValue as String) as String
    Dim lookupLine As Integer, cellValue As String

    lookupLine = 1
    cellValue = Cells(lookupLine, 1).Value

    While cellValue <> lookupValue
      lookupLine = lookupLine + 1
      cellValue = Cells(lookupLine, 1).Value
    Wend

    FindLine = lookupLine
end function

Function LookupImportFields() As String

  'look in A column : "Table Name",
  'Revert Column B
  Dim lookupLine As Integer, cellValue As String

  lookupLine = 1
  cellValue = Cells(lookupLine, 1).Value

  While cellValue <> "Table Name"
    lookupLine = lookupLine + 1
    cellValue = Cells(lookupLine, 1).Value
  Wend

  SqlTableName = Cells(lookupLine, 2)

End Function


Sub UpdateData(firstLine As Integer, firstColumn As Integer)

  Dim Sql As String, SqlFields As String, tableName As String, countSql As Integer
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
          SqlFields = ""
          While Cells(3, maxField) <> ""

                      If countSql > 1 Then

                          SqlFields = SqlFields & " , "
                      End If

                      If Cells(5, maxField) <> "" Then

                          SqlFields = SqlFields & Cells(3, maxField) & " = '" & Cells(5, maxField) & "'"
                          Else
                          SqlFields = SqlFields & Cells(3, maxField) & " = NULL"
                      End If

              maxField = maxField + 1
              countSql = countSql + 1
          Wend
          'Timestamp ????

          Sql = "UPDATE " & tableName & " SET " & SqlFields & " WHERE " & Cells(3, 1) & " = " & Cells(5, 1)
          'MsgBox Sql
          oConn.Execute Sql
          tableLine = tableLine + 1

      Wend
  'MsgBox Sql
      oConn.Close
      Set oConn = Nothing

End Sub
