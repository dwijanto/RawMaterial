Public Class seriesSelection
    Private DataTable1 As DataTable
    Private DataTable2 As DataTable
    Private DataTable3 As DataTable
    Dim myCheck As Boolean = False
    Dim myCheck2 As Boolean = False
    Dim unit As String = vbEmpty


    Public Event CurrencyChanged(ByVal CrCyCode As String)
    Private _Sqlstr As String
    Public Property Sqlstr() As String
        Get
            Return _Sqlstr
        End Get
        Set(ByVal value As String)
            _Sqlstr = value
        End Set
    End Property



    Private _DataSet As DataSet
    Public Property [DataSet]() As DataSet
        Get
            Return _DataSet
        End Get
        Set(ByVal value As DataSet)
            _DataSet = value
        End Set
    End Property

    Private _DataSet1 As DataSet
    Public Property [DataSet1]() As DataSet
        Get
            Return _DataSet1
        End Get
        Set(ByVal value As DataSet)
            _DataSet1 = value
        End Set
    End Property
    Private _DataSet2 As DataSet
    Public Property [DataSet2]() As DataSet
        Get
            Return _DataSet2
        End Get
        Set(ByVal value As DataSet)
            _DataSet2 = value
        End Set
    End Property


    Private _SeriesLegend As String
    Public Property SeriesLegend() As String
        Get
            Return _SeriesLegend
        End Get
        Set(ByVal value As String)
            _SeriesLegend = value
        End Set
    End Property
    Private Sub seriesSelection_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not DesignMode Then
            Try
                Call FillData()
                Call BindingData()
                myCheck = True
                Call roComboBox1_SelectedIndexChanged(Me, e)
                myCheck2 = True
            Catch ex As Exception

            End Try
        End If
    End Sub

    Private Sub FillData()
        If Not _DataSet Is Nothing Then
            Try
                DataTable1 = New DataTable("Master")
                DataTable2 = New DataTable("Detail")
                DataTable3 = New DataTable("Currency")

                DataTable1 = _DataSet.Tables(0)
                DataTable2 = _DataSet.Tables(1)
                DataTable3 = _DataSet.Tables(2)
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub BindingData()
        If Not _DataSet Is Nothing Then
            Try
                RoComboBox1.DataSource = DataTable1.DefaultView
                RoComboBox1.DisplayMember = "master"
                RoComboBox1.ValueMember = "master"
                RoComboBox2.DataSource = Nothing
                RoComboBox3.DataSource = DataTable3.DefaultView
                RoComboBox3.DisplayMember = "crcycode"
                RoComboBox3.ValueMember = "crcycode"
            Catch ex As Exception

            End Try
        End If

    End Sub
    Private Sub roComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoComboBox1.SelectedIndexChanged
        If myCheck Then
            If RoComboBox1.Text = "" Then
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
            End If
            Try
                Dim myFilter As String = "sheetname='" & RoComboBox1.Text.ToString & "'"
                DataTable2.DefaultView.RowFilter = myFilter
                RoComboBox2.DataSource = DataTable2.DefaultView
                RoComboBox2.DisplayMember = "namesource" '"rawmaterialname"
                RoComboBox2.ValueMember = "keyid"
                myCheck2 = True
                RoComboBox2_SelectedIndexChanged(Me, e)
            Catch ex As Exception
            End Try
        End If
    End Sub

    Private Sub RoComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoComboBox2.SelectedIndexChanged
        If myCheck2 Then
            If Not _DataSet Is Nothing Then
                Try
                    unit = ""
                    Dim myrow() As DataRow = Nothing
                    myrow = _DataSet.Tables(1).Select("keyid='" + RoComboBox2.SelectedValue.ToString & "'")
                    If myrow(0).ItemArray(7).ToString <> "" Then
                        unit = " (" & myrow(0).ItemArray(7).ToString & " )"
                    End If
                Catch ex As Exception
                End Try
            End If
            Dim toolTip1 As ToolTip = New ToolTip()
            toolTip1.AutoPopDelay = 0
            toolTip1.InitialDelay = 0
            toolTip1.ReshowDelay = 0
            toolTip1.ShowAlways = True
            toolTip1.SetToolTip(Me.RoComboBox2, RoComboBox2.Text.ToString())

        End If
    End Sub

    Private Sub RoComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoComboBox3.SelectedIndexChanged
        RaiseEvent CurrencyChanged(RoComboBox3.Text.ToString)
    End Sub
End Class
