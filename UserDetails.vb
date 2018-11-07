Imports C1.Win.C1FlexGrid
Imports System.Data.SqlClient

Public Class UserDetails
    Dim docNum As Integer
    Const rowCount = 15
    Const vDoc = 0
    Const vName = 1
    Const vDOB = 2
    Const vMotherName = 3
    Const vGender = 4
    Const vAddress1 = 5
    Const vAddress2 = 6
    Const vAddress3 = 7
    Const vState = 8
    Const vCity = 9
    Const vContact1 = 10
    Const vContact2 = 11
    Const vMobileNumber = 12
    Const vTelephoneNumber = 13
    Const vHobbies = 14

    Private Sub Registeration_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        docNum = 1
        'docNum += 1
        vsfUserDetails.Rows(vMobileNumber).Visible = False
        vsfUserDetails.AllowMerging = AllowMergingEnum.FixedOnly
        vsfUserDetails.Cols(0).AllowMerging = True
        'Dim rng As C1.Win.C1FlexGrid.CellRange = vsfUserDetails.GetCellRange(5, 0, 7, 0)
        'rng.Data = "Address"
        With vsfUserDetails
            .Rows.Count = rowCount
            .SetData(vDoc, 0, "Document Number")
            .SetData(vDoc, 1, docNum)
            .SetData(vName, 0, "Name")
            .SetData(vDOB, 0, "Date Of Birth")
            .SetData(vMotherName, 0, "Mother Name")
            .SetData(vGender, 0, "Gender")
            .SetData(vAddress1, 0, "Address")
            .SetData(vAddress2, 0, "Address")
            .SetData(vAddress3, 0, "Address")
            .SetData(vState, 0, "State")
            .SetData(vCity, 0, "City")
            .SetData(vContact1, 0, "Contact")
            .SetData(vContact1, 1, "Mobile Number")
            .SetData(vContact2, 0, "Contact")
            .SetData(vContact2, 1, "Telephone Number")
            .Rows(vTelephoneNumber).Visible = False
            '.SetData(vMobileNumber, 0, "Mobile Number")
            '.SetData(vTelephoneNumber, 0, "Telephone Number")
            .SetData(vHobbies, 0, "Hobbies")
        End With
    End Sub

    Private Sub vsfUserDetails_BeforeEdit(sender As Object, e As RowColEventArgs) Handles vsfUserDetails.BeforeEdit
        If e.Col = 1 Then
            Select Case e.Row
                Case vGender
                    vsfUserDetails.Rows(e.Row).ComboList = "Male|Female"
                Case vDOB
                    vsfUserDetails.Rows(e.Row).DataType = GetType(DateTime)
                Case vCity
                    vsfUserDetails.Rows(e.Row).ComboList = "|"
                Case vState
                    vsfUserDetails.Rows(e.Row).DataMap = LoadState()
                    'vsfUserDetails.Rows(e.Row).ComboList = LoadState()
                Case vContact1
                    vsfUserDetails.Rows(e.Row).DataType = GetType(Boolean)
                Case vContact2
                    vsfUserDetails.Rows(e.Row).DataType = GetType(Boolean)
                Case vHobbies
                    vsfUserDetails.Rows(e.Row).ComboList = "..."
            End Select


        End If
    End Sub

    Private Function LoadState()
        Dim sqlConn As SqlConnection
        Dim reader As SqlDataReader
        Dim query As String
        Dim cmd As SqlCommand

        sqlConn = New SqlConnection()
        sqlConn.ConnectionString = "server=(local);database=Registration;Trusted_Connection=True"
        Try
            sqlConn.Open()
            query = "Select * From State"
            cmd = New SqlCommand(query, sqlConn)
            reader = cmd.ExecuteReader

            Dim state As New Specialized.ListDictionary
            While reader.Read
                state.Add(reader.Item(0), reader.Item(1))
            End While
            sqlConn.Close()
            Return state
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlConn.Dispose()
        End Try
    End Function

    Private Function LoadCity()
        Dim sqlConn As SqlConnection
        Dim reader As SqlDataReader
        Dim query As String
        Dim cmd As SqlCommand
        Dim city As String = "|"
        sqlConn = New SqlConnection()
        sqlConn.ConnectionString = "server=(local);database=Registration;Trusted_Connection=True"
        Try
            sqlConn.Open()
            query = "Select * From City where StateId = '" & vsfUserDetails.GetData(vState, 1) & "'"
            cmd = New SqlCommand(query, sqlConn)
            reader = cmd.ExecuteReader

            'Dim city As New Specialized.ListDictionary
            'While reader.Read
            '    city.Add(reader.Item(0), reader.Item(2))
            'End While
            sqlConn.Close()
            Return city
        Catch ex As Exception
            MsgBox(ex.Message)
        Finally
            sqlConn.Dispose()
        End Try
    End Function

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim sqlConn As SqlConnection
        Dim query As String
        Dim cmd As SqlCommand
        sqlConn = New SqlConnection()
        sqlConn.ConnectionString = "server=(local);database=Registration;Trusted_Connection=True"

        Try
            Dim dialog As DialogResult
            dialog = MessageBox.Show("Are you sure you want to save", "Save", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If dialog = DialogResult.Yes Then
                sqlConn.Open()
                query = "Insert into [User](Name,DOB,MotherName,Gender,Address,CityId,StateId,MobileNumber,TelephoneNumber,Hobbies) 
Values ('" & vsfUserDetails.GetData(vName, 1) & "','" & vsfUserDetails.GetData(vDOB, 1) & "','" & vsfUserDetails.GetData(vMotherName, 1) & "','" & vsfUserDetails.GetData(vGender, 1) & "','" & vsfUserDetails.GetData(vAddress1, 1) & " " & vsfUserDetails.GetData(vAddress2, 1) & " " & vsfUserDetails.GetData(vAddress3, 1) & "','" & vsfUserDetails.GetData(vCity, 1) & "','" & vsfUserDetails.GetData(vState, 1) & "','" & vsfUserDetails.GetData(vMobileNumber, 1) & "','" & vsfUserDetails.GetData(vTelephoneNumber, 1) & "','" & vsfUserDetails.GetData(vHobbies, 1) & "')"
                cmd = New SqlCommand(query, sqlConn)

                If cmd.ExecuteNonQuery > 0 Then
                    MessageBox.Show("Data Saved Successfully", "Successfull !")
                End If
                sqlConn.Close()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        For i As Integer = 0 To vsfUserDetails.Rows.Count - 1
            vsfUserDetails.Rows(i).Item(1) = Nothing
        Next

        docNum = docNum + 1
        vsfUserDetails.Rows(0).Item(1) = docNum

        'Me.Controls.Clear()
        'InitializeComponent()
        'Registeration_Load(e, e)
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Are you sure you want to Exit", "Exit", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
        If dialog = DialogResult.Yes Then
            End
        End If
    End Sub

    Private Sub vsfUserDetails_AfterEdit(sender As Object, e As RowColEventArgs) Handles vsfUserDetails.AfterEdit
        Dim query As String
        Dim reader As SqlDataReader
        Dim sqlConn As SqlConnection
        Dim cmd As SqlCommand
        Dim adapter As New SqlDataAdapter
        Dim bindingSource As New BindingSource
        Dim table As New DataTable
        Dim city As New Specialized.ListDictionary()
        'Dim city As String = "|"

        Try
            If e.Col = 1 And e.Row = 8 Then
                sqlConn = New SqlConnection()
                sqlConn.ConnectionString = "server=(local);database=Registration;Trusted_Connection=True"
                sqlConn.Open()
                query = "Select * From City where StateId = '" & vsfUserDetails.GetData(vState, 1) & "'"
                'query = "SELECT * FROM City where StateId IN (SELECT StateId FROM State where StateName = '" & vsfUserDetails.GetData(vState, 1) & "')"
                cmd = New SqlCommand(query, sqlConn)
                reader = cmd.ExecuteReader

                'While reader.Read
                '    If city = "" Then
                '        city = reader.Item(2)
                '    Else
                '        city = city & "|" & reader.Item(2)
                '    End If
                'End While

                While reader.Read
                    city.Add(reader.Item(0), reader.Item(2))
                End While

                vsfUserDetails.Rows(vCity).DataMap = city
                sqlConn.Close()
            End If

            If e.Col = 1 And e.Row = 9 Then
                sqlConn = New SqlConnection()
                sqlConn.ConnectionString = "server=(local);database=Registration;Trusted_Connection=True"
                sqlConn.Open()

                Dim query1 As String
                query1 = "Select * from City where StateId = '" & vsfUserDetails.GetData(vState, 1) & "' And CityName = '" & vsfUserDetails.GetData(vCity, 1) & "'"
                cmd = New SqlCommand(query1, sqlConn)
                reader = cmd.ExecuteReader
                'adapter.SelectCommand = cmd
                'adapter.Fill(table)
                'bindingSource.DataSource = table

                If vsfUserDetails.GetData(vCity, 1) <> Nothing And reader.HasRows = False Then
                    Dim query2 As String
                    query2 = "Insert Into City(StateId,CityName) Values ('" & vsfUserDetails.GetData(vState, 1) & "','" & vsfUserDetails.GetData(vCity, 1) & "') "
                    cmd = New SqlCommand(query2, sqlConn)
                    reader.Close()
                    reader = cmd.ExecuteReader
                End If
                sqlConn.Close()
            End If

            If e.Col = 1 Then
                With vsfUserDetails
                    Select Case e.Row
                '    Case vContact
                '        If Not IsNumeric(.Editor.Text) Then
                '            MsgBox("Enter numeric value")
                '            e.Cancel = True
                '        End If
                        Case vContact1
                            .SetData(vMobileNumber, 0, "Mobile Number")
                            If (vsfUserDetails.GetData(e.Row, 1)) = "True" Then
                                vsfUserDetails.Rows(vMobileNumber).Visible = True
                            Else
                                vsfUserDetails.Rows(vMobileNumber).Visible = False
                            End If

                        Case vContact2
                            .SetData(vTelephoneNumber, 0, "Telephone Number")
                            If (vsfUserDetails.GetData(e.Row, 1)) = "True" Then
                                vsfUserDetails.Rows(vTelephoneNumber).Visible = True
                            Else
                                vsfUserDetails.Rows(vTelephoneNumber).Visible = False
                            End If
                    End Select
                End With
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub vsfUserDetails_CellButtonClick(sender As Object, e As RowColEventArgs) Handles vsfUserDetails.CellButtonClick
        If e.Col = 1 Then
            Dim hobbies As String = InputBox("Enter your hobbies")
            If hobbies = "" Then
                MsgBox("Please enter hobbies")
            Else
                'vsfNewEmployee.SetData(vHobbies, 0, hobbies)
                vsfUserDetails(vHobbies, e.Col) = hobbies
            End If

        End If
    End Sub

    Private Sub vsfUserDetails_ValidateEdit(sender As Object, e As ValidateEditEventArgs) Handles vsfUserDetails.ValidateEdit
        If e.Col = 1 Then
            With vsfUserDetails
                Select Case e.Row
                    Case vMobileNumber
                        If Not IsNumeric(.Editor.Text) And .Editor.Text <> Nothing Then
                            MessageBox.Show("Enter numeric value", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            e.Cancel = True
                        End If
                    Case vTelephoneNumber
                        If Not IsNumeric(.Editor.Text) And .Editor.Text <> Nothing Then
                            MessageBox.Show("Enter numeric value", "Error !", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                            e.Cancel = True
                        End If
                End Select
            End With
        End If
    End Sub

End Class
