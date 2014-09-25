
Imports System.Data.SQLite

''' <summary>
''' Test code for creating and updating sqlite databases
''' </summary>
''' <remarks></remarks>
Public Class Form1

    Private dbConn As SQLite.SQLiteConnection
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        dbConn.Close()
    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' Setup database connection
        dbConn = New SQLite.SQLiteConnection("Datasource=test.db;Version=3;New=True")
        Try
            dbConn.Open()
        Catch ex As Exception
            MsgBox(ex.Message + vbCrLf + vbCrLf + ex.StackTrace, vbCritical + vbOK, "Unable to open database")
            End
        End Try

        CreateTables()
        UpdateTables()
    End Sub

    Private Sub CreateTable(tableName As String, tableFields As String)

        Dim strSQL As String
        Dim command As SQLiteCommand

        command = dbConn.CreateCommand()

        Try
            command.CommandText = "DROP TABLE [" + tableName + "]"
            command.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("Unable to drop [" + tableName + "]")
        End Try

        strSQL = "CREATE TABLE [" + tableName + "] (" + _
            "[ID] INTEGER NOT NULL PRIMARY KEY AUTOINCREMENT, " + _
            "[CreationTimestamp] TIMESTAMP DEFAULT CURRENT_TIMESTAMP NOT NULL," + _
            "[UpdatedTimestamp] TIMESTAMP DEFAULT CURRENT_TIMESTAMP NULL," + tableFields + ");"
        
        Try
            command.CommandText = strSQL
            command.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox(ex.Message + vbCrLf + vbCrLf + ex.StackTrace, vbCritical + vbOK, _
                   "Unable to create table '" + tableName + "'")
            End
        End Try
    End Sub

    Private Sub CreateTables()
        ' Create tables if they don't already exist
        CreateTable("Customer", "[Forename] TEXT, " + _
                    "[Surname] TEXT, " + _
                    "[Address1] TEXT, " + _
                    "[Address2] TEXT")
        InsertData("Customer", "[Forename],[Surname],[Address1],[Address2]", _
                   """Donald"",""Duck"",""1 The Lake"",""Under the Rainbow""")

        CreateTable("Facility", "[CustomerID] INT NOT NULL," + _
                    "[FacilityType] TEXT, " + _
                    "[Volume] TEXT," + _
                    "[Airspace] TEXT," + _
                    "FOREIGN KEY (CustomerID) REFERENCES Customer(ID)")
    End Sub

    Private Sub InsertData(tableName As String, fieldList As String, dataList As String)
        Dim strSQL As String = "INSERT INTO " + tableName + "(" + fieldList + ") VALUES (" + dataList + ");"
        Dim command As SQLiteCommand
        Dim rv As Int16

        Try
            command = dbConn.CreateCommand()
            command.CommandText = strSQL
            rv = command.ExecuteNonQuery()
            If rv <> 1 Then MsgBox("Failed to insert " + vbCrLf + strSQL)
        Catch ex As Exception
            MsgBox(ex.Message + vbCrLf + vbCrLf + "Attempting: " + strSQL + vbCrLf + vbCrLf + _
                   ex.StackTrace, vbCritical + vbOK, "Unable to insert data '" + tableName + "'")
            End
        End Try
    End Sub

    Private Sub UpdateTables()
        ' Update table structure
    End Sub


    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Dim strQuery As String
        If txtForename.Text <> vbNullString Then
            strQuery = "[Forename]=""" & txtForename.Text & """"
        End If
        If txtSurname.Text <> vbNullString Then
            If strQuery <> vbNullString Then strQuery += " AND "
            strQuery += "[Surname]=""" & txtSurname.Text & """"
        End If
        If strQuery = vbNullString Then
            MsgBox("Nothing to search for")
        Else
            strQuery = "SELECT COUNT(*) FROM Customer WHERE " + strQuery + ";"
            Dim command As SQLiteCommand
            Try
                command = dbConn.CreateCommand()
                command.CommandText = strQuery

                Dim rv As Int16
                rv = command.ExecuteScalar()
                MsgBox(rv)
            Catch ex As Exception
                MsgBox(ex.Message + vbCrLf + vbCrLf + "Attempting: " + strQuery + vbCrLf + vbCrLf + _
                       ex.StackTrace, vbCritical + vbOK, "Attempting to search")
            End Try

        End If
    End Sub
End Class
