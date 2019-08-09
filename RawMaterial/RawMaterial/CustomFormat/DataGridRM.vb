Imports DJLib.Dbtools
Public Class DataGridRM
    Dim mDataSet As DataSet
    Dim BindingSource1 As New BindingSource
    Dim CurrencyManager1 As CurrencyManager
    Public Event RecordCount(ByVal RecordCount As Long)
    Dim mRecordCount As Long
    Dim tooltip1 As New ToolTip
    Public Property DataSet1 As DataSet
        Get
            Return mDataSet
        End Get
        Set(ByVal value As DataSet)
            mDataSet = value
        End Set
    End Property
    Public ReadOnly Property GetRecordCount As Long
        Get
            Return mRecordCount
        End Get
    End Property

    Private Sub DataGridRM_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
    Public Sub LoadData()

        If Not DesignMode Then
            Try
                Call FillData()
                Call BindingData()
            Catch ex As Exception
            End Try
        End If
    End Sub
    Private Sub FillData()
        ComboBox1.DataSource = DataSet1.Tables(2).DefaultView
        ComboBox1.DisplayMember = "crcycode"
        ComboBox1.ValueMember = "crcycode"

    End Sub
    Private Sub BindingData()
        Try
            If Not (DataSet1 Is Nothing) Then
                BindingSource1.DataSource = DataSet1.Tables(1).DefaultView 'DataTable1 'dataview1
                DataGridView1.DataSource = BindingSource1
                With DataGridView1
                    .RowsDefaultCellStyle.BackColor = Color.Bisque
                    .AlternatingRowsDefaultCellStyle.BackColor = Color.White
                End With
                Me.DataGridView1.Columns("Date").DefaultCellStyle.Format = "dd-MM-yyyy"
                CurrencyManager1 = CType(Me.BindingContext(BindingSource1), CurrencyManager)
                mRecordCount = CurrencyManager1.Count
                ShowRecordCount(mRecordCount)


                CheckedListBox1.DataSource = DataSet1.Tables(3).DefaultView
                CheckedListBox1.DisplayMember = "namesource" '"rawmaterialname"
                CheckedListBox1.ValueMember = "keyid"
                tooltip1.SetToolTip(GroupBox2, "If no date selected, All Price history is displayed")
                VisibleExport(CheckBox1.Checked)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub ShowRecordCount(ByVal myRecordCount As Long)
        RaiseEvent RecordCount(myRecordCount)
    End Sub

    Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged, DateTimePicker1.ValueChanged
        Button1.Enabled = DateTimePicker1.Checked Or DateTimePicker2.Checked
    End Sub

    Private Sub ApplyFilter()
        Dim myArray(0) As String
        Dim i As Integer = 0
        If DateTimePicker1.Checked Then
            ReDim Preserve myArray(0)
            DateTimePicker1.Format = DateTimePickerFormat.Short
            myArray(0) = "Date > =#" & DateTimePicker1.Value.Date & "#"
            i = i + 1

        End If

        If DateTimePicker2.Checked Then
            ReDim Preserve myArray(i)
            DateTimePicker1.Format = DateTimePickerFormat.Short
            myArray(i) = "Date <=#" & DateTimePicker2.Value.Date & "#"
            i = i + 1
        End If
        Dim myfilter As String = Join(myArray, " and ")
        DataSet1.Tables(1).DefaultView.RowFilter = myfilter
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        mRecordCount = CurrencyManager1.Count
        ShowRecordCount(mRecordCount)
    End Sub


    'Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
    '    If ComboBox1.Text = "" Then
    '        CheckedListBox1.DataSource = Nothing
    '    Else
    '        CheckedListBox1.DataSource = DataSet1.Tables(3).DefaultView
    '        CheckedListBox1.DisplayMember = "rawmaterialname"
    '        CheckedListBox1.ValueMember = "keyid"
    '    End If
    'End Sub


    Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged
        'If ComboBox1.Text <> "" And ComboBox1.Text <> "System.Data.DataRowView" Then
        'If ComboBox1.Text <> "System.Data.DataRowView" Then
        Select Case sender.selectedindex
            Case 0
                Dim chkstate As CheckState
                chkstate = CheckedListBox1.GetItemCheckState(0)
                For i = 0 To sender.items.count - 1
                    CheckedListBox1.SetItemChecked(i, chkstate)
                    DataGridView1.Columns.Item(i).visible = True
                Next
            Case Else
                CheckedListBox1.SetItemChecked(0, 0)
                Dim mycountlist As Integer = Countlist(CheckedListBox1)
                If CheckedListBox1.Items.Count = mycountlist + 1 Then
                    CheckedListBox1.SetItemChecked(0, True)
                End If

                Dim mycheck As CheckState
                If mycountlist = 0 Then
                    For i = 1 To sender.items.count - 1
                        DataGridView1.Columns.Item(i).visible = True
                    Next
                Else
                    For i = 1 To sender.items.count - 1

                        mycheck = CheckedListBox1.GetItemCheckState(i)
                        DataGridView1.Columns.Item(i).visible = mycheck
                    Next
                End If
                
             
        End Select
        'End If
    End Sub

    Public Function Countlist(ByVal myCheckedListbox As CheckedListBox) As Integer
        Dim count As Integer = 0
        For i = 0 To myCheckedListbox.Items.Count - 1
            If CheckedListBox1.GetItemCheckState(i) Then
                count += 1
            End If
        Next
        Return count
    End Function
    Public Function getStringBuilderSingle(ByRef sheetName As String) As String
        Dim myreturn As String = String.Empty
        Dim myCriteria(0) As String
        Dim myIndex As Integer = 0
        Dim myDate1 As Date = Nothing
        Dim myDate2 As Date = Nothing
        Dim myList As String = getselecteditems(CheckedListBox1)

        If myList <> "" Then
            ReDim Preserve myCriteria(myIndex)
            myCriteria(myIndex) = myList
            myIndex = myIndex + 1
        Else
            ReDim Preserve myCriteria(myIndex)
            myCriteria(myIndex) = "hd.sheetname = '" & sheetName & "'"
            myIndex = myIndex + 1
        End If
        If DateTimePicker1.Checked Then
            DateTimePicker1.Format = DateTimePickerFormat.Short
            myDate1 = DateTimePicker1.Value
            ReDim Preserve myCriteria(myIndex)
            myCriteria(myIndex) = "tx.txdate >= " & DateFormatyyyyMMdd(myDate1)
            myIndex = myIndex + 1
        End If
        If DateTimePicker2.Checked Then
            DateTimePicker2.Format = DateTimePickerFormat.Short
            myDate2 = DateTimePicker2.Value
            ReDim Preserve myCriteria(myIndex)
            myCriteria(myIndex) = "tx.txdate <= " & DateFormatyyyyMMdd(myDate2)
            myIndex = myIndex + 1
        End If
        myreturn = Join(myCriteria, " and ")
        If myreturn <> "" Then
            myreturn = " where " & myreturn
        End If
        Return myreturn
    End Function

    Private Function getselecteditems(ByVal CLB As CheckedListBox) As String
        Dim mylist(0) As String
        Dim myIndex As Integer = 0
        Dim myReturn As String = String.Empty
        If Not (CLB.GetItemCheckState(0) = CheckState.Checked) Then
            For i = 1 To CLB.Items.Count - 1
                If CLB.GetItemCheckState(i) = CheckState.Checked Then
                    ReDim Preserve mylist(myIndex)
                    CLB.SetSelected(i, True)
                    mylist(myIndex) = "tx.keyid='" & CLB.SelectedValue & "'"
                    myIndex = myIndex + 1
                End If
            Next
            myReturn = Join(mylist, " or ")
        End If
        If myReturn <> "" Then
            myReturn = "(" & myReturn & ")"
        End If
        Return myReturn
    End Function

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        VisibleExport(CheckBox1.Checked)
        If Not CheckBox1.Checked Then
            ComboBox1.Text = ""
        End If
    End Sub
    Private Sub VisibleExport(ByVal myStatus As Boolean)
        ComboBox1.Visible = myStatus
        Label3.Visible = myStatus
        Label7.Visible = myStatus
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ApplyFilter()
    End Sub
End Class
