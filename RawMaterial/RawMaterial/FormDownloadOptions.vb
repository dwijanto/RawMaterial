Imports System.Data
Imports System.Data.Common
Imports Npgsql
Public Class FormDownloadOptions

    'Private DataAdapter1 As DbDataAdapter
    Private DataAdapter1 As NpgsqlDataAdapter
    Private Param As DbParameter
    Private Dataset1 As DataSet
    Private DataTable1 As DataTable
    Private BindingSource1 As New BindingSource
    Private DataCollection As Collection
    Private Check1Changed As Boolean = False


    Private Sub FormDownloadOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call ConnectData()
        Call FillData()
        Call BindingData()
        CheckBox1_CheckedChanged(Me, e)
    End Sub
    Private Sub ConnectData()
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)

        'Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        'Dim conn As DbConnection = factory.CreateConnection

        Dim conn = New NpgsqlConnection(DbTools1.getConnectionString)
        'conn.ConnectionString = DbTools1.getConnectionString
        Try
            'Dim selectcmd As DbCommand = factory.CreateCommand
            Dim selectcmd = New NpgsqlCommand()
            selectcmd.Connection = conn
            selectcmd.CommandText = "select paramdt.paramdtid,paramdt.paramname,paramdt.cvalue,paramdt.bvalue from rmparamdt paramdt" & _
                                    " left join rmparamhd ph on ph.paramhdid = paramdt.paramhdid" & _
                                    " where ph.paramname = 'DownloadOptions'" & _
                                    " order by paramdt.ivalue "
            selectcmd.CommandType = CommandType.Text
            DataAdapter1 = New NpgsqlDataAdapter
            DataAdapter1.SelectCommand = selectcmd
            Dim updatecmd = New NpgsqlCommand()
            'Dim updatecmd As DbCommand = factory.CreateCommand
            updatecmd.Connection = conn
            updatecmd.CommandText = "Update rmparamdt set paramdtid=@paramdtid,cvalue=@cvalue,bvalue=@bvalue where paramdtid=@paramdtid"
            DataAdapter1.UpdateCommand = updatecmd
            Param = updatecmd.CreateParameter
            Param.SourceColumn = "paramdtid"
            Param.ParameterName = "@paramdtid"
            Param.DbType = DbType.Int64
            Param.Value = DataRowVersion.Original
            updatecmd.Parameters.Add(Param)

            Param = updatecmd.CreateParameter
            Param.SourceColumn = "cvalue"
            Param.ParameterName = "@cvalue"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updatecmd.Parameters.Add(Param)

            Param = updatecmd.CreateParameter
            Param.SourceColumn = "bvalue"
            Param.ParameterName = "@bvalue"
            Param.DbType = DbType.Boolean
            Param.Value = DataRowVersion.Current
            updatecmd.Parameters.Add(Param)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub

    Private Sub ConnectDataFactory()
        Dim DbTools1 As New DJLib.Dbtools(myUserid, myPassword)
        Dim factory As DbProviderFactory = DbProviderFactories.GetFactory(My.Settings.dbProviderFactory)
        Dim conn As DbConnection = factory.CreateConnection
        conn.ConnectionString = DbTools1.getConnectionString
        Try
            Dim selectcmd As DbCommand = factory.CreateCommand
            selectcmd.Connection = conn
            selectcmd.CommandText = "select paramdt.paramdtid,paramdt.paramname,paramdt.cvalue,paramdt.bvalue from rmparamdt paramdt" & _
                                    " left join rmparamhd ph on ph.paramhdid = paramdt.paramhdid" & _
                                    " where ph.paramname = 'DownloadOptions'" & _
                                    " order by paramdt.ivalue "
            selectcmd.CommandType = CommandType.Text
            DataAdapter1 = factory.CreateDataAdapter
            DataAdapter1.SelectCommand = selectcmd

            Dim updatecmd As DbCommand = factory.CreateCommand
            updatecmd.Connection = conn
            updatecmd.CommandText = "Update rmparamdt set paramdtid=@paramdtid,cvalue=@cvalue,bvalue=@bvalue where paramdtid=@paramdtid"
            DataAdapter1.UpdateCommand = updatecmd
            Param = updatecmd.CreateParameter
            Param.SourceColumn = "paramdtid"
            Param.ParameterName = "@paramdtid"
            Param.DbType = DbType.Int64
            Param.Value = DataRowVersion.Original
            updatecmd.Parameters.Add(Param)

            Param = updatecmd.CreateParameter
            Param.SourceColumn = "cvalue"
            Param.ParameterName = "@cvalue"
            Param.DbType = DbType.String
            Param.Value = DataRowVersion.Current
            updatecmd.Parameters.Add(Param)

            Param = updatecmd.CreateParameter
            Param.SourceColumn = "bvalue"
            Param.ParameterName = "@bvalue"
            Param.DbType = DbType.Boolean
            Param.Value = DataRowVersion.Current
            updatecmd.Parameters.Add(Param)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub FillData()
        Dataset1 = New DataSet("paramdt")
        DataAdapter1.FillSchema(Dataset1, SchemaType.Source)
        DataAdapter1.MissingSchemaAction = MissingSchemaAction.AddWithKey 'Will add Primary & Unique key info
        DataTable1 = New DataTable("paramdt")
        DataAdapter1.Fill(DataTable1)
        'DataView1 = New DataView(Dataset1.Tables("product")) 
    End Sub
    Private Sub BindingData()
        dataCollection = New Collection
        For Each DataRow In DataTable1.Rows
            If DataRow(1).ToString = "useproxy" Then
                dataCollection.Add(DataRow(0), DataRow(1).ToString)
                CheckBox1.Checked = DataRow(3)
            Else
                Select Case DataRow(1).ToString
                    Case "proxyserver"
                        TextBox1.Text = DataRow(2).ToString
                    Case "server"
                        TextBox3.Text = DataRow(2).ToString
                    Case "port"
                        TextBox2.Text = DataRow(2).ToString
                    Case "username"
                        TextBox4.Text = DataRow(2).ToString
                    Case "password"
                        TextBox5.Text = DataRow(2).ToString
                    Case "saveto"
                        TextBox6.Text = DataRow(2).ToString
                    Case "url"
                        TextBox7.Text = DataRow(2).ToString
                End Select
                dataCollection.Add(DataRow(0), DataRow(1).ToString)
            End If
        Next


    End Sub

    Private Sub bOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bOk.Click
        SaveData()
        DataAdapter1.Update(DataTable1)
        Call bCancel_Click(Me, e)
    End Sub
    '
    Private Sub bCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bCancel.Click
        Me.Close()
    End Sub

    Private Sub SaveData()
        Dim DataRow1 As DataRow
        If Check1Changed Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("useproxy"))
            DataRow1.BeginEdit()
            DataRow1("bvalue") = CheckBox1.Checked
            DataRow1.EndEdit()
        End If
        If TextBox1.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("proxyserver"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox1.Text
            DataRow1.EndEdit()
        End If
        If TextBox2.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("port"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox2.Text
            DataRow1.EndEdit()
        End If
        If TextBox3.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("server"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox3.Text
            DataRow1.EndEdit()
        End If
        If TextBox4.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("username"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox4.Text
            DataRow1.EndEdit()
        End If
        If TextBox5.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("password"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox5.Text
            DataRow1.EndEdit()
        End If
        If TextBox6.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("saveto"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox6.Text
            DataRow1.EndEdit()
        End If
        If TextBox7.Modified Then
            DataRow1 = DataTable1.Rows.Find(dataCollection("url"))
            DataRow1.BeginEdit()
            DataRow1("cvalue") = TextBox7.Text
            DataRow1.EndEdit()
        End If

    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Check1Changed = True
        TextBox1.Enabled = CheckBox1.Checked
        TextBox2.Enabled = CheckBox1.Checked
        TextBox3.Enabled = CheckBox1.Checked
        TextBox4.Enabled = CheckBox1.Checked
        TextBox5.Enabled = CheckBox1.Checked
    End Sub
End Class