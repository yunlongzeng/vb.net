Imports System.Windows.Forms
Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Public Function finddata(sqltxt)
        DataGridView1.DataSource = Nothing
        Dim connstr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\database\Ins.accdb"

        Dim conn As New OleDb.OleDbConnection(connstr)
        conn.Open() '打开连接
        Dim da As New OleDb.OleDbDataAdapter()
        da.SelectCommand = New OleDbCommand(sqltxt, conn)
        Dim dt As New DataTable
        Try
            da.Fill(dt)
            DataGridView1.DataSource = dt
            DataGridView1.Columns(0).Width = 60
            DataGridView1.Columns(1).Width = 70
            DataGridView1.Columns(2).Width = 80
            DataGridView1.Columns(3).Width = 150
            DataGridView1.Columns(4).Width = 70
            DataGridView1.Columns(5).Width = 80
            DataGridView1.Columns(6).Width = 90
        Catch ex As Exception
            'MsgBox("异常")
        End Try
        conn.Close()

    End Function

    Public Function updatedata(sqltxt)

        Dim connstr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\database\Ins.accdb"

        Dim conn As New OleDb.OleDbConnection(connstr)
        conn.Open() '打开连接
        Dim da As New OleDb.OleDbDataAdapter()
        da.SelectCommand = New OleDbCommand(sqltxt, conn)
        da.SelectCommand.ExecuteNonQuery()

        conn.Close()
    End Function

    Public Function codeexist(sqltxt)
        Dim Con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=\\TSHAIPS96-31.mdir.co\SHA_Engineering\Instrumentation\CPQExcelView\database\Ins.accdb")
        Con.Open()
        Dim objCommand As New OleDbCommand(sqltxt, Con)
        Dim objReader As OleDbDataReader = objCommand.ExecuteReader()
        Dim strData As String
        If objReader.HasRows Then
            While objReader.Read()
                strData = String.Empty
                For intIndex As Integer = 0 To objReader.FieldCount - 1
                    strData &= objReader.Item(intIndex).ToString
                Next
            End While

        End If
        Return strData
    End Function
    Private Sub DataGridView1_CellContentClick(sender As Object, e As Windows.Forms.DataGridViewCellEventArgs)

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.KeyPreview = True '注册窗体的键盘事件
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        If e.KeyCode = Keys.Enter Then
            Dim datasheet As String
            If TextBox1.Text = "8" Then
                datasheet = "SOV"
            ElseIf TextBox1.Text = "14" Then
                datasheet = "AOV"
            ElseIf TextBox1.Text = "18" Then
                datasheet = "PPS"
            ElseIf TextBox1.Text = "19" Then
                datasheet = "AFR"
            ElseIf TextBox1.Text = "23" Then
                datasheet = "VB"
            Else
                MsgBox("SmartCode is wrong!")
                datasheet = "空"
            End If
            Dim sqltxt As String
            If TextBox3.Text <> "" Then
                sqltxt = "Select * from " & datasheet & " where IdCode = '" & TextBox2.Text _
                & "' OR TypeCode like '%" & TextBox3.Text & "%'"
            Else
                sqltxt = "Select * from " & datasheet & " where IdCode = '" & TextBox2.Text & "'"
            End If
            finddata(sqltxt)
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Dim datasheet As String
        If TextBox1.Text = "8" Then
            datasheet = "SOV"
        ElseIf TextBox1.Text = "14" Then
            datasheet = "AOV"
        ElseIf TextBox1.Text = "18" Then
            datasheet = "PPS"
        ElseIf TextBox1.Text = "19" Then
            datasheet = "AFR"
        ElseIf TextBox1.Text = "23" Then
            datasheet = "VB"
        Else
            MsgBox("SmartCode is wrong!")
            datasheet = "空"
        End If
        Dim sqltxt As String
        If TextBox3.Text <> "" Then
            sqltxt = "Select * from " & datasheet & " where IdCode = '" & TextBox2.Text _
                & "' OR TypeCode like '%" & TextBox3.Text & "%'"
        Else
            sqltxt = "Select * from " & datasheet & " where IdCode = '" & TextBox2.Text & "'"
        End If
        finddata(sqltxt)

    End Sub

    Private Sub ListView1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        DataGridView1.EndEdit()
        Dim datasheet As String
        If TextBox1.Text = "8" Then
            datasheet = "SOV"
        ElseIf TextBox1.Text = "14" Then
            datasheet = "AOV"
        ElseIf TextBox1.Text = "18" Then
            datasheet = "PPS"
        ElseIf TextBox1.Text = "19" Then
            datasheet = "AFR"
        ElseIf TextBox1.Text = "23" Then
            datasheet = "VB"
        Else
            MsgBox("SmartCode is wrong!")
            datasheet = "空"
        End If
        Dim a = DataGridView1.RowCount
        If a = 0 Then  '添加空行
            Dim connstr As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\CPQExcelView\Ins.accdb"
            Dim sqltxt As String

            sqltxt = "Select * from " & datasheet & " where IdCode = '增加空行'"

            Dim conn As New OleDb.OleDbConnection(connstr)
            conn.Open() '打开连接
            Dim da As New OleDb.OleDbDataAdapter()
            da.SelectCommand = New OleDbCommand(sqltxt, conn)
            Dim dt As New DataTable
            Try
                da.Fill(dt)
                DataGridView1.DataSource = dt
                DataGridView1.Columns(0).Width = 60
                DataGridView1.Columns(1).Width = 70
                DataGridView1.Columns(2).Width = 80
                DataGridView1.Columns(3).Width = 150
                DataGridView1.Columns(4).Width = 70
                DataGridView1.Columns(5).Width = 80
                DataGridView1.Columns(6).Width = 90
            Catch ex As Exception
                'MsgBox("异常")
            End Try

        Else
            For i = 0 To a - 2
                If TextBox1.Text = 8 Then
                    Dim d1 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(2).Value), "", DataGridView1.Rows(i).Cells(2).Value))
                    Dim d2 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), "", DataGridView1.Rows(i).Cells(3).Value))
                    Dim d3 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(4).Value), "", DataGridView1.Rows(i).Cells(4).Value))
                    Dim d4 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(5).Value), "", DataGridView1.Rows(i).Cells(5).Value))
                    Dim d5 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(6).Value), "", DataGridView1.Rows(i).Cells(6).Value))
                    Dim d6 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(7).Value), "", DataGridView1.Rows(i).Cells(7).Value))
                    If d1 = "" Or d2 = "" Or d3 = "" Or d4 = "" Or d5 = "" Or d6 = "" Then
                        MsgBox("Information isn't complete!")
                    Else
                        Dim sqltx As String
                        sqltx = "Select count(*) from " & datasheet & " where IdCode = '" & DataGridView1.Rows(i).Cells(2).Value & "'"
                        Dim existline = codeexist(sqltx)
                        If existline = "0" Then
                            Dim sqltxt As String
                            sqltxt = "INSERT INTO SOV(SmartCode,IdCode,TypeCode,Bracket,InsSize,ExhaustPort,InsFunction,InsConnection) VALUES('8','" _
                        & DataGridView1.Rows(i).Cells(2).Value & "','" & DataGridView1.Rows(i).Cells(3).Value & "','" & DataGridView1.Rows(i).Cells(4).Value & "','" _
                         & DataGridView1.Rows(i).Cells(5).Value & "','" & DataGridView1.Rows(i).Cells(6).Value & "','" & DataGridView1.Rows(i).Cells(7).Value & "','" _
                          & DataGridView1.Rows(i).Cells(8).Value & "')"

                            updatedata(sqltxt)
                            MsgBox("Insert successful.")
                        Else
                            Dim intResult As Integer
                            intResult = MessageBox.Show("Do you want to replace?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

                            If intResult = DialogResult.OK Then
                                Dim sqltxt As String
                                sqltxt = "Update " & datasheet & " SET TypeCode = '" & DataGridView1.Rows(i).Cells(3).Value & "',Bracket= '" & DataGridView1.Rows(i).Cells(4).Value & "',InsSize= '" & DataGridView1.Rows(i).Cells(5).Value & "',ExhaustPort= '" & DataGridView1.Rows(i).Cells(6).Value & "',InsFunction = '" & DataGridView1.Rows(i).Cells(7).Value & "',InsConnection= '" & DataGridView1.Rows(i).Cells(8).Value & "' WHERE IdCode = '" _
                                & DataGridView1.Rows(i).Cells(2).Value & "'"
                                updatedata(sqltxt)
                                MsgBox("Replace successful.")
                            End If
                        End If
                    End If
                ElseIf TextBox1.Text = 14 Then
                    Dim d1 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(2).Value), "", DataGridView1.Rows(i).Cells(2).Value))
                    Dim d2 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), "", DataGridView1.Rows(i).Cells(3).Value))
                    Dim d3 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(4).Value), "", DataGridView1.Rows(i).Cells(4).Value))
                    Dim d4 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(5).Value), "", DataGridView1.Rows(i).Cells(5).Value))
                    Dim d5 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(6).Value), "", DataGridView1.Rows(i).Cells(6).Value))
                    If d1 = "" Or d2 = "" Or d3 = "" Or d4 = "" Or d5 = "" Then
                        MsgBox("Information isn't complete!")
                    Else
                        Dim sqltx As String
                        sqltx = "Select count(*) from " & datasheet & " where IdCode = '" & DataGridView1.Rows(i).Cells(2).Value & "'"
                        Dim existline = codeexist(sqltx)
                        If existline = "0" Then
                            Dim sqltxt As String
                            sqltxt = "INSERT INTO AOV(SmartCode,IdCode,TypeCode,Bracket,InsSize,ExhaustPort) VALUES('14','" _
                        & DataGridView1.Rows(i).Cells(2).Value & "','" & DataGridView1.Rows(i).Cells(3).Value & "','" & DataGridView1.Rows(i).Cells(4).Value & "','" _
                         & DataGridView1.Rows(i).Cells(5).Value & "','" & DataGridView1.Rows(i).Cells(6).Value & "')"

                            updatedata(sqltxt)
                            MsgBox("Insert successful.")
                        Else
                            Dim intResult As Integer
                            intResult = MessageBox.Show("Do you want to replace?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

                            If intResult = DialogResult.OK Then
                                Dim sqltxt As String
                                sqltxt = "Update " & datasheet & " SET TypeCode = '" & DataGridView1.Rows(i).Cells(3).Value & "',Bracket= '" & DataGridView1.Rows(i).Cells(4).Value & "',InsSize= '" & DataGridView1.Rows(i).Cells(5).Value & "',ExhaustPort= '" & DataGridView1.Rows(i).Cells(6).Value & "' WHERE IdCode = '" _
                                & DataGridView1.Rows(i).Cells(2).Value & "'"
                                updatedata(sqltxt)
                                MsgBox("Replace successful.")
                            End If
                        End If
                    End If
                ElseIf TextBox1.Text = 19 Then
                    Dim d1 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(2).Value), "", DataGridView1.Rows(i).Cells(2).Value))
                    Dim d2 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), "", DataGridView1.Rows(i).Cells(3).Value))
                    Dim d3 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(4).Value), "", DataGridView1.Rows(i).Cells(4).Value))
                    Dim d4 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(5).Value), "", DataGridView1.Rows(i).Cells(5).Value))
                    Dim d5 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(6).Value), "", DataGridView1.Rows(i).Cells(6).Value))
                    If d1 = "" Or d2 = "" Or d3 = "" Or d4 = "" Or d5 = "" Then
                        MsgBox("Information isn't complete!")
                    Else
                        Dim sqltx As String
                        sqltx = "Select count(*) from " & datasheet & " where IdCode = '" & DataGridView1.Rows(i).Cells(2).Value & "'"
                        Dim existline = codeexist(sqltx)
                        If existline = "0" Then
                            Dim sqltxt As String
                            sqltxt = "INSERT INTO AFR(SmartCode,IdCode,TypeCode,Bracket,InsSize,InsMaterial) VALUES('19','" _
                        & DataGridView1.Rows(i).Cells(2).Value & "','" & DataGridView1.Rows(i).Cells(3).Value & "','" & DataGridView1.Rows(i).Cells(4).Value & "','" _
                         & DataGridView1.Rows(i).Cells(5).Value & "','" & DataGridView1.Rows(i).Cells(6).Value & "')"

                            updatedata(sqltxt)
                            MsgBox("Insert successful.")
                        Else
                            Dim intResult As Integer
                            intResult = MessageBox.Show("Do you want to replace?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

                            If intResult = DialogResult.OK Then
                                Dim sqltxt As String
                                sqltxt = "Update " & datasheet & " SET TypeCode = '" & DataGridView1.Rows(i).Cells(3).Value & "',Bracket= '" & DataGridView1.Rows(i).Cells(4).Value & "',InsSize= '" & DataGridView1.Rows(i).Cells(5).Value & "',InsMaterial= '" & DataGridView1.Rows(i).Cells(6).Value & "' WHERE IdCode = '" _
                                & DataGridView1.Rows(i).Cells(2).Value & "'"
                                updatedata(sqltxt)
                                MsgBox("Replace successful.")
                            End If
                        End If
                    End If
                ElseIf TextBox1.Text = 18 Then
                    Dim d1 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(2).Value), "", DataGridView1.Rows(i).Cells(2).Value))
                    Dim d2 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), "", DataGridView1.Rows(i).Cells(3).Value))
                    Dim d3 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(4).Value), "", DataGridView1.Rows(i).Cells(4).Value))
                    Dim d4 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(5).Value), "", DataGridView1.Rows(i).Cells(5).Value))
                    Dim d5 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(6).Value), "", DataGridView1.Rows(i).Cells(6).Value))
                    If d1 = "" Or d2 = "" Or d3 = "" Or d4 = "" Or d5 = "" Then
                        MsgBox("Information isn't complete!")
                    Else
                        Dim sqltx As String
                        sqltx = "Select count(*) from " & datasheet & " where IdCode = '" & DataGridView1.Rows(i).Cells(2).Value & "'"
                        Dim existline = codeexist(sqltx)
                        If existline = "0" Then
                            Dim sqltxt As String
                            sqltxt = "INSERT INTO PPS(SmartCode,IdCode,TypeCode,Bracket,InsSize,ExhaustPort) VALUES('14','" _
                        & DataGridView1.Rows(i).Cells(2).Value & "','" & DataGridView1.Rows(i).Cells(3).Value & "','" & DataGridView1.Rows(i).Cells(4).Value & "','" _
                         & DataGridView1.Rows(i).Cells(5).Value & "','" & DataGridView1.Rows(i).Cells(6).Value & "')"

                            updatedata(sqltxt)
                            MsgBox("Insert successful.")
                        Else
                            Dim intResult As Integer
                            intResult = MessageBox.Show("Do you want to replace?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

                            If intResult = DialogResult.OK Then
                                Dim sqltxt As String
                                sqltxt = "Update " & datasheet & " SET TypeCode = '" & DataGridView1.Rows(i).Cells(3).Value & "',Bracket= '" & DataGridView1.Rows(i).Cells(4).Value & "',InsSize= '" & DataGridView1.Rows(i).Cells(5).Value & "',ExhaustPort= '" & DataGridView1.Rows(i).Cells(6).Value & "' WHERE IdCode = '" _
                                & DataGridView1.Rows(i).Cells(2).Value & "'"
                                updatedata(sqltxt)
                                MsgBox("Replace successful.")
                            End If
                        End If
                    End If
                ElseIf TextBox1.Text = 23 Then
                    Dim d1 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(2).Value), "", DataGridView1.Rows(i).Cells(2).Value))
                    Dim d2 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(3).Value), "", DataGridView1.Rows(i).Cells(3).Value))
                    Dim d3 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(4).Value), "", DataGridView1.Rows(i).Cells(4).Value))
                    Dim d4 = CStr(IIf(IsDBNull(DataGridView1.Rows(i).Cells(5).Value), "", DataGridView1.Rows(i).Cells(5).Value))
                    If d1 = "" Or d2 = "" Or d3 = "" Or d4 = "" Then
                        MsgBox("Information isn't complete!")
                    Else
                        Dim sqltx As String
                        sqltx = "Select count(*) from " & datasheet & " where IdCode = '" & DataGridView1.Rows(i).Cells(2).Value & "'"
                        Dim existline = codeexist(sqltx)
                        If existline = "0" Then
                            Dim sqltxt As String
                            sqltxt = "INSERT INTO VB(SmartCode,IdCode,TypeCode,Bracket,InsSize) VALUES('14','" _
                        & DataGridView1.Rows(i).Cells(2).Value & "','" & DataGridView1.Rows(i).Cells(3).Value & "','" & DataGridView1.Rows(i).Cells(4).Value & "','" _
                         & DataGridView1.Rows(i).Cells(5).Value & "')"

                            updatedata(sqltxt)
                            MsgBox("Insert successful.")
                        Else
                            Dim intResult As Integer
                            intResult = MessageBox.Show("Do you want to replace?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

                            If intResult = DialogResult.OK Then
                                Dim sqltxt As String
                                sqltxt = "Update " & datasheet & " SET TypeCode = '" & DataGridView1.Rows(i).Cells(3).Value & "',Bracket= '" & DataGridView1.Rows(i).Cells(4).Value & "',InsSize= '" & DataGridView1.Rows(i).Cells(5).Value & "' WHERE IdCode = '" _
                                & DataGridView1.Rows(i).Cells(2).Value & "'"
                                updatedata(sqltxt)
                                MsgBox("Replace successful.")
                            End If
                        End If
                    End If
                Else
                    MsgBox("SmartCode is wrong!")
                End If
            Next
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        DataGridView1.EndEdit()
        Dim datasheet As String
        If TextBox1.Text = "8" Then
            datasheet = "SOV"
        ElseIf TextBox1.Text = "14" Then
            datasheet = "AOV"
        ElseIf TextBox1.Text = "18" Then
            datasheet = "PPS"
        ElseIf TextBox1.Text = "19" Then
            datasheet = "AFR"
        ElseIf TextBox1.Text = "23" Then
            datasheet = "VB"
        Else
            MsgBox("SmartCode is wrong!")
            datasheet = "空"
        End If
        Dim a = DataGridView1.RowCount
        Dim d1 = CStr(IIf(IsDBNull(DataGridView1.CurrentRow.Cells(2).Value), "", DataGridView1.CurrentRow.Cells(2).Value))

        If a = 0 Then  '判断是否有数据
            MsgBox("You choosed the blank line!")
        ElseIf d1 = "" Then
            MsgBox("You choosed the blank line!")
        Else
            Dim sqltxt As String
            sqltxt = "Delete from " & datasheet & " WHERE IdCode = '" _
            & DataGridView1.CurrentRow.Cells(2).Value & "'"
            Dim intResult As Integer
            intResult = MessageBox.Show("Do you want to delete?", "Warning!", MessageBoxButtons.OKCancel, MessageBoxIcon.Asterisk, MessageBoxDefaultButton.Button1)

            If intResult = DialogResult.OK Then
                updatedata(sqltxt)
                sqltxt = "Select * from " & datasheet & " where IdCode = '" & TextBox2.Text & "'"
                finddata(sqltxt)
            End If
        End If
    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

End Class