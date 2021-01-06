Imports System.Data
Imports System.Data.SqlClient

Public Class SkillMatrixClass
    Public gridMainDbConnection As SqlConnection
    Public gridMainDbCommand As SqlCommand
    Public gridMainDbConnectionState As Boolean = False

    Public Function OpenMainDbConnection() As Boolean

        Me.gridMainDbConnection = New SqlConnection
        Me.gridMainDbCommand = New SqlCommand

        Me.gridMainDbConnection.ConnectionString = "Data Source=WPPHL039SQL01;" &
                            "Initial Catalog=RPA_GRID" & ";" &
                            "Persist Security Info=True;" &
                            "Integrated Security=SSPI;" &
                            "Connect Timeout=30;"


        If Me.gridMainDbConnection.State = ConnectionState.Open Then
            Me.CloseMainDbConnection()
        End If

        Try

            Me.gridMainDbConnection.Open()

            Me.gridMainDbCommand.Connection = Me.gridMainDbConnection

            If Me.gridMainDbConnection.State = ConnectionState.Open Then
                Me.gridMainDbConnectionState = True
                OpenMainDbConnection = True
            Else
                Me.gridMainDbConnectionState = False
                OpenMainDbConnection = False
            End If

        Catch ex As SqlException
            OpenMainDbConnection = False

        End Try

    End Function

    Public Sub CloseMainDbConnection()

        Me.gridMainDbConnection.Close()
        Me.gridMainDbConnectionState = False

    End Sub


    'GET SEARCH ITEMS'
    Public Function GetTowerList() As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()

            Me.gridMainDbCommand.CommandText = "  SELECT DISTINCT [TOWER] FROM [dbo].[vQuery_Team] WHERE [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Tower];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("Tower"))

                End While

                dr.Close()
                If temp.Count > 1 Then
                    temp.Add("-ALL-")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetSumTowerList() As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [Tower] FROM [dbo].[vQuery_Team] WHERE [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Tower];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("Tower"))

                End While

                dr.Close()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetDepartmentList(ByVal _tower As String) As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _tower)

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [Department] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Department];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("Department"))

                End While

                dr.Close()
                'If temp.Count > 1 Then
                '    temp.Add("-ALL-")
                'End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetSumDepartmentList(ByVal _tower As String) As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _tower)

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [Department] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Department];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("Department"))

                End While

                dr.Close()
                If temp.Count > 1 Then
                    temp.Add("-ALL-")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetSegmentList(ByVal _Tower As String, ByVal _Dept As String) As DataTable

        Dim temp As New DataTable

        temp.Columns.Add("Id")
        temp.Columns.Add("Segment")

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _Tower)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Department", _Dept)
            If _Dept = "-ALL-" Then
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Segment];"
            Else
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [Department]=@Department AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Segment];"

            End If

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read
                    temp.Rows.Add(dr.Item("Id"), dr.Item("Segment"))


                End While

                dr.Close()

                If temp.Rows.Count > 1 Then
                    temp.Rows.Add(0, "-ALL-")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetSumSegmentList(ByVal _Tower As String, ByVal _Dept As String) As DataTable

        Dim temp As New DataTable

        temp.Columns.Add("Id")
        temp.Columns.Add("Segment")

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _Tower)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Department", _Dept)
            If _Dept = "-ALL-" Then
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull)  ORDER BY [Segment];"
            Else
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [Department]=@Department AND [ID] IN (SELECT DISTINCT TeamId from vQuery_TeamPull) ORDER BY [Segment];"

            End If

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader
                If _Dept = "-ALL-" Then

                    temp.Rows.Add(0, "-ALL-")
                Else

                    While dr.Read

                        temp.Rows.Add(dr.Item("Id"), dr.Item("Segment"))


                    End While

                End If

                dr.Close()
                If temp.Rows.Count > 1 Then
                    temp.Rows.Add(0, "-ALL-")
                End If



            Catch ex As Exception
            End Try

            Me.CloseMainDbConnection()


        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function
    Public Function GetAgentInfo(ByVal _tower As String, ByVal _dept As String, ByVal _seg As Integer, ByVal _segStr As String) As DataTable
        Dim temp As New DataTable

        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            If _tower = "-ALL-" Then
                da = New SqlDataAdapter("SELECT DISTINCT * FROM [dbo].[vQuery_SkillViewInfo];", Me.gridMainDbConnection)
            ElseIf _dept = "-ALL-" Then
                da = New SqlDataAdapter("SELECT DISTINCT * FROM [dbo].[vQuery_SkillViewInfo] WHERE [Tower]='" & _tower & "' ORDER BY [Tower];", Me.gridMainDbConnection)
            ElseIf _segStr = "-ALL-" Then
                da = New SqlDataAdapter("SELECT DISTINCT * FROM [dbo].[vQuery_SkillViewInfo] WHERE [Department]='" & _dept & "' ORDER BY [Tower];", Me.gridMainDbConnection)
            Else
                da = New SqlDataAdapter("SELECT DISTINCT * FROM [dbo].[vQuery_SkillViewInfo] WHERE [TeamId]=" & _seg & " ORDER BY [Name];", Me.gridMainDbConnection)
            End If

            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If


        If Not IsNothing(temp) Then
            Dim dt = Me.GetSkillInfo()
            Dim dt2 = Me.GetSPInfo(_seg)
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    temp.Columns.Add(" " & dt.Rows(i).Item("SkillName") & " ")
                Next


            End If

            For i = 0 To temp.Rows.Count - 1
                If Not IsNothing(dt2) Then
                    If dt2.Rows.Count > 0 Then
                        For j = 0 To dt2.Rows.Count - 1

                            If temp.Rows(i).Item("EmpNo") = dt2.Rows(j).Item("EmpNo") Then
                                temp.Rows(i).Item(" " & dt2.Rows(j).Item("SkillName") & " ") = "X"
                            End If

                        Next

                    End If

                End If

            Next

        End If


        'If temp.Columns(0).ColumnName = "TeamId" Then
        '    temp.Columns.Remove("TeamId")
        'End If




        Return temp
    End Function


    'Public Function GetAgentLogInfo(ByVal _seg As Integer) As List(Of AgentLogInfo)

    '    Dim temp As New List(Of AgentLogInfo)
    '    If OpenMainDbConnection() = True Then
    '        Me.gridMainDbCommand.Parameters.Clear()
    '        Me.gridMainDbCommand.Parameters.AddWithValue("@TeamId", _seg)

    '        Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[vQuery_AgentInfoLog] WHERE [TeamId]=@TeamId ORDER BY [EmpName];"


    '        Try

    '            Dim dr As SqlDataReader
    '            dr = Me.gridMainDbCommand.ExecuteReader

    '            While dr.Read

    '                Dim temp2 As New AgentLogInfo
    '                temp2.TeamId = dr.Item("TeamId")
    '                temp2.EmpNo = dr.Item("EmpNo")
    '                temp2.EmpName = dr.Item("EmpName")
    '                temp2.Domestic = dr.Item("Domestic")
    '                temp2.Supervisor = If(IsDBNull(dr.Item("Supervisor")), "", dr.Item("Supervisor"))
    '                temp2.Manager = If(IsDBNull(dr.Item("Manager")), "", dr.Item("Manager"))
    '                temp2.OrderProcess = If(IsDBNull(dr.Item("OrderProcess")), "", dr.Item("OrderProcess"))
    '                temp2.Credit = If(IsDBNull(dr.Item("Credit")), "", dr.Item("Credit"))
    '                temp2.Rebill = If(IsDBNull(dr.Item("Rebill")), "", dr.Item("Rebill"))
    '                temp2.Complaints = If(IsDBNull(dr.Item("Complaints")), "", dr.Item("Complaints"))
    '                temp2.ShippingDiscrepancy = If(IsDBNull(dr.Item("ShippingDiscrepancy")), "", dr.Item("ShippingDiscrepancy"))
    '                temp2.DocumentRequests = If(IsDBNull(dr.Item("DocumentRequest")), "", dr.Item("DocumentRequest"))
    '                temp2.PricingProduct = If(IsDBNull(dr.Item("PricingProduct")), "", dr.Item("PricingProduct"))
    '                temp2.Implementation = If(IsDBNull(dr.Item("Implementation")), "", dr.Item("Implementation"))
    '                temp2.ValueLinkTrained = If(IsDBNull(dr.Item("ValueLinkTrained")), "", dr.Item("ValueLinkTrained"))
    '                temp2.Router = If(IsDBNull(dr.Item("Router")), "", dr.Item("Router"))
    '                temp2.Phone = If(IsDBNull(dr.Item("Phone")), "", dr.Item("Phone"))
    '                temp2.Email = If(IsDBNull(dr.Item("Email")), "", dr.Item("Email"))
    '                temp2.Cases = If(IsDBNull(dr.Item("Cases")), "", dr.Item("Cases"))
    '                temp2.Commercial = If(IsDBNull(dr.Item("Commercial")), 0, dr.Item("Commercial"))
    '                temp2.DCS = If(IsDBNull(dr.Item("DCS")), 0, dr.Item("DCS"))
    '                temp2.GoV = If(IsDBNull(dr.Item("GoV")), 0, dr.Item("GoV"))
    '                temp2.Speciality = If(IsDBNull(dr.Item("Speciality")), 0, dr.Item("Speciality"))
    '                temp2.HSRouter = If(IsDBNull(dr.Item("HSRouter")), 0, dr.Item("HSRouter"))

    '                temp2.Action = If(IsDBNull(dr.Item("Action")), "", dr.Item("Action"))
    '                temp2.ModifiedBy = If(IsDBNull(dr.Item("ModifiedBy")), "", dr.Item("ModifiedBy"))
    '                temp2.ModifiedDate = If(IsDBNull(dr.Item("ModifiedDate")), Nothing, dr.Item("ModifiedDate"))
    '                temp2.ReviewBy = If(IsDBNull(dr.Item("ReviewBy")), "", dr.Item("ReviewBy"))
    '                temp2.ReviewDate = If(IsDBNull(dr.Item("ReviewDate")), Nothing, dr.Item("ReviewDate"))

    '                temp.Add(temp2)

    '            End While


    '            dr.Close()


    '        Catch ex As Exception
    '            MsgBox(ex.Message)
    '        End Try


    '        Me.CloseMainDbConnection()



    '    Else
    '        MsgBox("Failed to open main database. Please check you connection.")
    '    End If

    '    Return temp
    'End Function
    Public Function GetSkillList() As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [SkillHome] FROM [dbo].[tblAMSkill] ORDER BY [SkillHome];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("SkillHome"))

                End While

                dr.Close()
                If temp.Count > 1 Then
                    temp.Add("-ALL-")
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If


        Return temp
    End Function

    Public Function GetSkillViewList(ByRef _temp As List(Of String), ByVal _seg As Integer, ByVal ChannelParam As String, ByVal CoreBizParam As String, ByVal HomeSkillParam As String, ByVal ChannelPrioParam As String) As DataTable
        Dim tempEmpNo As New DataTable
        Dim temp As New DataTable
        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            If _seg = 0 Then
                da = New SqlDataAdapter("  SELECT DISTINCT [EmpNo] FROM [dbo].[vQuery_SkillViewInfo] WHERE [TeamId]  IN (SELECT DISTINCT TeamId from vQuery_TeamPull)  AND ( " & ChannelParam & " OR " & ChannelPrioParam & " ) AND  (" & CoreBizParam & ") AND  (" & HomeSkillParam & ");", Me.gridMainDbConnection)
            Else
                da = New SqlDataAdapter("  SELECT DISTINCT [EmpNo] FROM [dbo].[vQuery_SkillViewInfo] WHERE [TeamId]= " & _seg & "  AND ( " & ChannelParam & " OR " & ChannelPrioParam & " ) AND  (" & CoreBizParam & ") AND  (" & HomeSkillParam & ");", Me.gridMainDbConnection)
            End If


            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(tempEmpNo)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If



        If Not IsNothing(tempEmpNo) Then
            If tempEmpNo.Rows.Count > 0 Then

                Dim EmpNoStr As String = ""
                Dim SkillHomeStr As String = ""

                For I = 0 To tempEmpNo.Rows.Count - 1
                    If EmpNoStr = "" Then
                        EmpNoStr = EmpNoStr & "'" & tempEmpNo(I).Item("EmpNo") & "'"
                    Else
                        EmpNoStr = EmpNoStr & ",'" & tempEmpNo(I).Item("EmpNo") & "'"
                    End If
                Next


                For Each itm In _temp
                    If SkillHomeStr = "" Then
                        SkillHomeStr = SkillHomeStr & "'" & itm & "'"
                    Else
                        SkillHomeStr = SkillHomeStr & ",'" & itm & "'"
                    End If
                Next

                If EmpNoStr = "" Or SkillHomeStr = "" Then Return temp

                If OpenMainDbConnection() = True Then


                    Dim da As New SqlDataAdapter()
                    'da = New SqlDataAdapter("SELECT * FROM [dbo].[tblAMSkill] WHERE [SkillHome]='" & itm & "' ORDER BY [Id]; ", Me.gridMainDbConnection)
                    da = New SqlDataAdapter("SELECT        TOP (100) PERCENT dbo.tblAMSkill.SkillHome As [Home Skill], dbo.tblAMSkill.SkillName AS  [Skill Preference], COUNT(*) AS [HC]
                    FROM            dbo.tblAMSkill INNER JOIN
                                             dbo.tblAMSkillPreference ON dbo.tblAMSkill.Id = dbo.tblAMSkillPreference.IdSkill
                    WHERE        (dbo.tblAMSkillPreference.EmpNo IN (" & EmpNoStr & ") AND  [SkillHome] IN (" & SkillHomeStr & "))
                    GROUP BY dbo.tblAMSkill.SkillHome, dbo.tblAMSkill.SkillName
                    ORDER BY dbo.tblAMSkill.SkillHome, dbo.tblAMSkill.SkillName ", Me.gridMainDbConnection)
                    da.SelectCommand.CommandTimeout = 1000
                    Try
                        da.Fill(temp)
                    Catch ex As Exception
                    End Try

                    Me.CloseMainDbConnection()


                End If


            End If
        End If


        Return temp
    End Function
    Public Function GetSkillView(ByVal _dept As String, ByVal _seg As Integer, ByVal ChannelParam As String, ByVal CoreBizParam As String, ByVal HomeSkillParam As String, ByVal ChannelPrioParam As String) As DataTable
        Dim temp As New DataTable

        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            If _seg = 0 Then
                da = New SqlDataAdapter("  SELECT DISTINCT * FROM [dbo].[vQuery_SMSummaryView] WHERE [TeamId]  IN (SELECT DISTINCT TeamId from vQuery_TeamPull WHERE [Segment] = '" & _dept & "')  AND ( " & ChannelParam & " OR " & ChannelPrioParam & " ) AND  (" & CoreBizParam & ") AND  (" & HomeSkillParam & ") ORDER BY [Name];", Me.gridMainDbConnection)
            Else
                Dim sql As String = "  SELECT DISTINCT * FROM [dbo].[vQuery_SMSummaryView] WHERE [TeamId]= " & _seg & "  AND ( " & ChannelParam & " OR " & ChannelPrioParam & " ) AND  (" & CoreBizParam & ") AND  (" & HomeSkillParam & ") ORDER BY [Name];"
                da = New SqlDataAdapter(sql, Me.gridMainDbConnection)
            End If

            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If


        If temp.Rows.Count > 1 Then
            Dim dt = Me.GetSkillInfo()
            Dim dt2 = Me.GetSPInfo(_seg)
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    temp.Columns.Add(" " & dt.Rows(i).Item("SkillName") & " ")
                Next


            End If

            For i = 0 To temp.Rows.Count - 1
                If Not IsNothing(dt2) Then
                    If dt2.Rows.Count > 0 Then
                        For j = 0 To dt2.Rows.Count - 1

                            If temp.Rows(i).Item("EmpNo") = dt2.Rows(j).Item("EmpNo") Then
                                temp.Rows(i).Item(" " & dt2.Rows(j).Item("SkillName") & " ") = "X"
                            End If

                        Next

                    End If

                End If

            Next

        End If


        If temp.Columns(0).ColumnName = "TeamId" Then
            temp.Columns.Remove("TeamId")
            temp.Columns.Remove("EmpNo")
        End If



        Return temp
    End Function



    'GET LOGS'
    Public Function GetCFLog(ByVal _temp As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_CoreFunctionLog] WHERE [TeamID]=" & _temp & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetHSLog(ByVal _temp As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_HomeSkillLog] WHERE [TeamID]=" & _temp & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetCHLog(ByVal _temp As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_ChannelLog] WHERE [TeamID]=" & _temp & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetSPLog(ByVal _temp As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT DISTINCT * FROM [dbo].[vQuery_SkillPreferenceLog] WHERE [TeamID]=" & _temp & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        If Not IsNothing(temp) Then
            Dim dt = Me.GetSkillInfo() 'skill names
            Dim dt2 = Me.GetSPTeamLogInfo(_temp)
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    temp.Columns.Add(" " & dt.Rows(i).Item("SkillName") & " ")
                Next


            End If

            For i = 0 To temp.Rows.Count - 1
                If Not IsNothing(dt2) Then
                    If dt2.Rows.Count > 0 Then
                        For j = 0 To dt2.Rows.Count - 1

                            If temp.Rows(i).Item("EmpNo") = dt2.Rows(j).Item("EmpNo") Then
                                temp.Rows(i).Item(" " & dt2.Rows(j).Item("SkillName") & " ") = "X"
                            End If

                        Next

                    End If

                End If

            Next

        End If

        Return temp
    End Function


    'GET PER TAB INFO'
    Public Function GetCFInfo(ByVal _EmpNo As String) As CoreFunctionInfo
        Dim temp As New CoreFunctionInfo

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)

            Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMCoreFunction] WHERE [EmpNo]=@EmpNo;"


            Try

                Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

                While dr.Read
                    temp.OrderProcess = If(IsDBNull(dr.Item("OrderProcess")), 0, dr.Item("OrderProcess"))
                    temp.Credit = If(IsDBNull(dr.Item("Credit")), 0, dr.Item("Credit"))
                    temp.Rebill = If(IsDBNull(dr.Item("Rebill")), 0, dr.Item("Rebill"))
                    temp.Returns = If(IsDBNull(dr.Item("Returns")), 0, dr.Item("Returns"))
                    temp.Complaints = If(IsDBNull(dr.Item("Complaints")), 0, dr.Item("Complaints"))
                    temp.ShippingDiscrepancy = If(IsDBNull(dr.Item("ShippingDiscrepancy")), 0, dr.Item("ShippingDiscrepancy"))
                    temp.DocumentRequests = If(IsDBNull(dr.Item("DocumentRequests")), 0, dr.Item("DocumentRequests"))
                    temp.PricingProduct = If(IsDBNull(dr.Item("PricingProduct")), 0, dr.Item("PricingProduct"))
                    temp.Implementation = If(IsDBNull(dr.Item("Implementation")), 0, dr.Item("Implementation"))

                    temp.CarrierDisposition = If(IsDBNull(dr.Item("CarrierDisposition")), 0, dr.Item("CarrierDisposition"))
                    temp.AccountMaintenance = If(IsDBNull(dr.Item("AccountMaintenance")), 0, dr.Item("AccountMaintenance"))
                    temp.SpecialProcessOrders = If(IsDBNull(dr.Item("SpecialProcessOrders")), 0, dr.Item("SpecialProcessOrders"))
                    temp.Research = If(IsDBNull(dr.Item("Research")), 0, dr.Item("Research"))

                    temp.ValueLinkTrained = If(IsDBNull(dr.Item("ValueLinkTrained")), 0, dr.Item("ValueLinkTrained"))
                    temp.Comment = If(IsDBNull(dr.Item("Comment")), "", dr.Item("Comment"))
                End While


                dr.Close()


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If

        Return temp
    End Function
    Public Function GetHSInfo(ByVal _EmpNo As String) As HomeSkillInfo
        Dim temp As New HomeSkillInfo

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)

            Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMHomeSkill] WHERE [EmpNo]=@EmpNo;"


            Try

                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read


                    temp.BackOffice = If(IsDBNull(dr.Item("BackOffice")), 0, dr.Item("BackOffice"))
                    temp.CHI = If(IsDBNull(dr.Item("CHI")), 0, dr.Item("CHI"))
                    temp.Concierge = If(IsDBNull(dr.Item("Concierge")), 0, dr.Item("Concierge"))
                    temp.Commercial = If(IsDBNull(dr.Item("Commercial")), 0, dr.Item("Commercial"))
                    temp.GOV = If(IsDBNull(dr.Item("GOV")), 0, dr.Item("GOV"))
                    temp.IDNC = If(IsDBNull(dr.Item("IDNC")), 0, dr.Item("IDNC"))
                    temp.Kaiser = If(IsDBNull(dr.Item("Kaiser")), 0, dr.Item("Kaiser"))
                    temp.Router = If(IsDBNull(dr.Item("Router")), 0, dr.Item("Router"))
                    temp.SalesSupport = If(IsDBNull(dr.Item("SalesSupport")), 0, dr.Item("SalesSupport"))
                    temp.Specialty = If(IsDBNull(dr.Item("Specialty")), 0, dr.Item("Specialty"))
                    temp.Tradex = If(IsDBNull(dr.Item("Tradex")), 0, dr.Item("Tradex"))

                End While


                dr.Close()


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If

        Return temp
    End Function
    Public Function GetCHInfo(ByVal _EmpNo As String) As ChannelInfo
        Dim temp As New ChannelInfo

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)

            Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMChannel] WHERE [EmpNo]=@EmpNo;"


            Try

                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read
                    temp.Router = If(IsDBNull(dr.Item("Router")), 0, dr.Item("Router"))
                    temp.Phone = If(IsDBNull(dr.Item("Phone")), 0, dr.Item("Phone"))
                    temp.Email = If(IsDBNull(dr.Item("Email")), 0, dr.Item("Email"))
                    temp.Cases = If(IsDBNull(dr.Item("Cases")), 0, dr.Item("Cases"))

                    temp.PhonePrio = If(IsDBNull(dr.Item("PhonePrio")), "", dr.Item("PhonePrio"))
                    temp.CasesEmailPrio = If(IsDBNull(dr.Item("CasesEmailPrio")), "", dr.Item("CasesEmailPrio"))
                    temp.Reason = If(IsDBNull(dr.Item("Reason")), "", dr.Item("Reason"))
                End While


                dr.Close()


            Catch ex As Exception
                MsgBox(ex.Message)
            End Try


            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If

        Return temp
    End Function

    Public Function GetSkillPrefInfo(ByVal _EmpNo As String) As List(Of Integer)
        Dim temp As New List(Of Integer)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)

            Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMSkillPreference] WHERE [EmpNo]=@EmpNo ORDER BY [IdSkill];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("IdSkill"))

                End While

                dr.Close()

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If

        Return temp
    End Function

    Public Function GetSkillInfo() As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[tblAMSkill] ORDER BY [SkillHome];", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function

    Public Function GetPrimarySkillInfo(_str As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[tblAMSkill] WHERE [ID] IN " & _str & " ORDER BY [ID];", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function

    Public Function GetSPInfo(ByVal _seg As Integer) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_SkillPrefView] WHERE [TeamId]=" & _seg & " ORDER BY [EmpNo];", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetSPLogInfo(ByVal _emp As Integer) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_SkillPrefLogView] WHERE [EmpNo]=" & _emp & ";", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function

    Public Function GetSPTeamLogInfo(ByVal _temp As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_SkillPrefLogView] WHERE [TeamId]=" & _temp & ";", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function

    'SAVE PER TAB'
    Public Function SaveCFInfo(ByVal _temp As CoreFunctionInfo) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then




            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@OrderProcess", _temp.OrderProcess)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Credit", _temp.Credit)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Rebill", _temp.Rebill)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Returns", _temp.Returns)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Complaints", _temp.Complaints)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ShippingDiscrepancy", _temp.ShippingDiscrepancy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@DocumentRequests", _temp.DocumentRequests)
            Me.gridMainDbCommand.Parameters.AddWithValue("@PricingProduct", _temp.PricingProduct)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Implementation", _temp.Implementation)

            Me.gridMainDbCommand.Parameters.AddWithValue("@CarrierDisposition", _temp.CarrierDisposition)
            Me.gridMainDbCommand.Parameters.AddWithValue("@AccountMaintenance", _temp.AccountMaintenance)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SpecialProcessOrders", _temp.SpecialProcessOrders)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Research", _temp.Research)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ValueLinkTrained", _temp.ValueLinkTrained)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Comment", _temp.Comment)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Pending")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")


            gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMCoreFunctionLog] ([EmpNo],[OrderProcess],[Credit],[Rebill],[Returns],[Complaints],[ShippingDiscrepancy],[DocumentRequests],[PricingProduct],[Implementation],[CarrierDisposition],[AccountMaintenance],[SpecialProcessOrders],[Research],[ValueLinkTrained],[Comment],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@OrderProcess,@Credit,@Rebill,@Returns,@Complaints,@ShippingDiscrepancy,@DocumentRequests,@PricingProduct,@Implementation,@CarrierDisposition,@AccountMaintenance,@SpecialProcessOrders,@Research,@ValueLinkTrained,@Comment,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"

            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True


                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function
    Public Function SaveHSInfo(ByVal _temp As HomeSkillInfo) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then


            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@BackOffice", _temp.BackOffice)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CHI", _temp.CHI)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Concierge", _temp.Concierge)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Commercial", _temp.Commercial)
            Me.gridMainDbCommand.Parameters.AddWithValue("@DCS", _temp.DCS)
            Me.gridMainDbCommand.Parameters.AddWithValue("@GOV", _temp.GOV)
            Me.gridMainDbCommand.Parameters.AddWithValue("@IDNC", _temp.IDNC)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Kaiser", _temp.Kaiser)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Router", _temp.Router)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SalesSupport", _temp.SalesSupport)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Specialty", _temp.Specialty)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tradex", _temp.Tradex)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SupplyAssurance", _temp.SupplyAssurance)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CET", _temp.CET)

            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Pending")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")


            gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMHomeSkillLog] ([EmpNo],[Commercial],[DCS],[GOV],[Specialty],[Router],[BackOffice],[CHI],[Concierge],[IDNC],[Kaiser],[SalesSupport],[Tradex],[SupplyAssurance],[CET],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@Commercial,@DCS,@GOV,@Specialty,@Router,@BackOffice,@CHI,@Concierge,@IDNC,@Kaiser,@SalesSupport,@Tradex,@SupplyAssurance,@CET,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"
            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function
    Public Function SaveCHInfo(ByVal _temp As ChannelInfo) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then


            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Router", _temp.Router)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Phone", _temp.Phone)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Email", _temp.Email)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Cases", _temp.Cases)

            Me.gridMainDbCommand.Parameters.AddWithValue("@PhonePrio", _temp.PhonePrio)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CasesEmailPrio", _temp.CasesEmailPrio)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Reason", _temp.Reason)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Pending")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")


            gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMChannelLog] ([EmpNo],[Router],[Phone],[Email],[Cases],[Action],[PhonePrio],[CasesEmailPrio],[Reason],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@Router,@Phone,@Email,@Cases,@Action,@PhonePrio,@CasesEmailPrio,@Reason,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"

            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function
    Public Function SaveSPInfo(ByRef _temp As List(Of Integer), ByVal _EmpNo As String) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then



            For Each itm In _temp
                Me.gridMainDbCommand.Parameters.Clear()
                Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)
                Me.gridMainDbCommand.Parameters.AddWithValue("@IdSkill", itm)

                Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
                Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

                Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Pending")

                Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
                Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")



                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMSkillPreferenceLog] ([EmpNo],[IdSkill],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@IdSkill,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"

                Try
                    If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                        temp = True

                    End If
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try







            Next
            Me.CloseMainDbConnection()
        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If

        Return temp

    End Function


    'SAVE PER LOG'
    Public Function SaveCFLog(ByVal _temp As CoreFunctionInfo, ByVal _Id As Integer) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then




            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@OrderProcess", _temp.OrderProcess)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Credit", _temp.Credit)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Rebill", _temp.Rebill)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Returns", _temp.Returns)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Complaints", _temp.Complaints)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ShippingDiscrepancy", _temp.ShippingDiscrepancy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@DocumentRequests", _temp.DocumentRequests)
            Me.gridMainDbCommand.Parameters.AddWithValue("@PricingProduct", _temp.PricingProduct)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Implementation", _temp.Implementation)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CarrierDisposition", _temp.CarrierDisposition)
            Me.gridMainDbCommand.Parameters.AddWithValue("@AccountMaintenance", _temp.AccountMaintenance)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SpecialProcessOrders", _temp.SpecialProcessOrders)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Research", _temp.Research)

            Me.gridMainDbCommand.Parameters.AddWithValue("@ValueLinkTrained", _temp.ValueLinkTrained)


            Me.gridMainDbCommand.Parameters.AddWithValue("@Comment", _temp.Comment)

            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", _temp.ModifiedBy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", _temp.ModifiedDate)



            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Approved")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Date.Now)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)
            Me.gridMainDbCommand.Parameters.AddWithValue("@LastModifiedDate", Date.Now)

            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMCoreFunction] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunction] SET OrderProcess=@OrderProcess,Credit=@Credit,Rebill=@Rebill,Returns=@Returns,Complaints=@Complaints,ShippingDiscrepancy=@ShippingDiscrepancy,DocumentRequests=@DocumentRequests,PricingProduct=@PricingProduct,Implementation=@Implementation,CarrierDisposition=@CarrierDisposition,AccountMaintenance=@AccountMaintenance,SpecialProcessOrders=@SpecialProcessOrders,Research=@Research,ValueLinkTrained=@ValueLinkTrained,ReviewDate=@ReviewDate,ReviewBy=@ReviewBy,Comment=@Comment WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMCoreFunction] ([EmpNo],[OrderProcess],[Credit],[Rebill],[Returns],[Complaints],[ShippingDiscrepancy],[DocumentRequests],[PricingProduct],[Implementation],[CarrierDisposition],[AccountMaintenance],[SpecialProcessOrders],[Research],[ValueLinkTrained],[Comment],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate],[LastModifiedDate]) VALUES (@EmpNo,@OrderProcess,@Credit,@Rebill,@Returns,@Complaints,@ShippingDiscrepancy,@DocumentRequests,@PricingProduct,@Implementation,@CarrierDisposition,@AccountMaintenance,@SpecialProcessOrders,@Research,@ValueLinkTrained,@Comment,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate,@LastModifiedDate)"
            End If

            dr.Close()



            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunctionLog] SET OrderProcess=@OrderProcess,Credit=@Credit,Rebill=@Rebill,Returns=@Returns,Complaints=@Complaints,ShippingDiscrepancy=@ShippingDiscrepancy,DocumentRequests=@DocumentRequests,PricingProduct=@PricingProduct,Implementation=@Implementation,CarrierDisposition=@CarrierDisposition,AccountMaintenance=@AccountMaintenance,SpecialProcessOrders=@SpecialProcessOrders,Research=@Research,ValueLinkTrained=@ValueLinkTrained,ReviewDate=@ReviewDate,ReviewBy=@ReviewBy,Action=@Action WHERE Id=" & _Id & ";"
                    Me.gridMainDbCommand.ExecuteNonQuery()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function

    Public Function SaveHSLog(ByVal _temp As HomeSkillInfo, ByVal _Id As Integer) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then


            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@BackOffice", _temp.BackOffice)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CHI", _temp.CHI)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Concierge", _temp.Concierge)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Commercial", _temp.Commercial)
            Me.gridMainDbCommand.Parameters.AddWithValue("@DCS", _temp.DCS)
            Me.gridMainDbCommand.Parameters.AddWithValue("@GOV", _temp.GOV)
            Me.gridMainDbCommand.Parameters.AddWithValue("@IDNC", _temp.IDNC)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Kaiser", _temp.Kaiser)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Router", _temp.Router)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SalesSupport", _temp.SalesSupport)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Specialty", _temp.Specialty)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tradex", _temp.Tradex)
            Me.gridMainDbCommand.Parameters.AddWithValue("@SupplyAssurance", _temp.SupplyAssurance)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CET", _temp.CET)

            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", _temp.ModifiedBy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", _temp.ModifiedDate)


            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Approved")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Date.Now())
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)


            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMHomeSkill] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMHomeSkill] SET BackOffice=@BackOffice,CHI=@CHI,Concierge=@Concierge,Commercial=@Commercial,DCS=@DCS,GOV=@GOV,IDNC=@IDNC,Kaiser=@Kaiser,Router=@Router,SalesSupport=@SalesSupport,Specialty=@Specialty,Tradex=@Tradex,SupplyAssurance=@SupplyAssurance,CET=@CET WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMHomeSkill] ([EmpNo],[Commercial],[DCS],[GOV],[Specialty],[Router],[BackOffice],[CHI],[Concierge],[IDNC],[Kaiser],[SalesSupport],[Tradex],[SupplyAssurance],[CET],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@Commercial,@DCS,@GOV,@Specialty,@Router,@BackOffice,@CHI,@Concierge,@IDNC,@Kaiser,@SalesSupport,@Tradex,@SupplyAssurance,@CET,@ModifiedBy,@ModifiedDate)"
            End If

            dr.Close()


            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMHomeSkillLog] SET BackOffice=@BackOffice,CHI=@CHI,Concierge=@Concierge,Commercial=@Commercial,DCS=@DCS,GOV=@GOV,IDNC=@IDNC,Kaiser=@Kaiser,Router=@Router,SalesSupport=@SalesSupport,Specialty=@Specialty,Tradex=@Tradex,SupplyAssurance=@SupplyAssurance,CET=@CET,Action=@Action,ReviewDate=@ReviewDate,ReviewBy=@ReviewBy WHERE Id=" & _Id & ";"
                    Me.gridMainDbCommand.ExecuteNonQuery()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function

    Public Function SaveCHLog(ByVal _temp As ChannelInfo, ByVal _Id As Integer) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then


            Me.gridMainDbCommand.Parameters.Clear()

            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Router", _temp.Router)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Phone", _temp.Phone)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Email", _temp.Email)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Cases", _temp.Cases)

            Me.gridMainDbCommand.Parameters.AddWithValue("@PhonePrio", _temp.PhonePrio)
            Me.gridMainDbCommand.Parameters.AddWithValue("@CasesEmailPrio", _temp.CasesEmailPrio)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Reason", _temp.Reason)

            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", _temp.ModifiedBy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", _temp.ModifiedDate)


            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Approved")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Date.Now())
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)



            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMChannel] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMChannel] SET Router=@Router,Phone=@Phone,Email=@Email,Cases=@Cases,PhonePrio=@PhonePrio,CasesEmailPrio=@CasesEmailPrio,Reason=@Reason WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMChannel] ([EmpNo],[Router],[Phone],[Email],[Cases],[PhonePrio],[CasesEmailPrio],[Reason],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@Router,@Phone,@Email,@Cases,@PhonePrio,@CasesEmailPrio,@Reason,@ModifiedBy,@ModifiedDate)"
            End If

            dr.Close()

            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMChannelLog] SET Router=@Router,Phone=@Phone,Email=@Email,Cases=@Cases,PhonePrio=@PhonePrio,CasesEmailPrio=@CasesEmailPrio,Reason=@Reason,Action=@Action,ReviewDate=@ReviewDate,ReviewBy=@ReviewBy WHERE Id=" & _Id & ";"
                    Me.gridMainDbCommand.ExecuteNonQuery()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function

    Public Function SaveSPLog(ByVal _temp As SkillPreference) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then
            gridMainDbCommand.CommandText = "DELETE FROM [dbo].[tblAMSkillPreference] WHERE [EmpNo]= " & _temp.EmpNo & ";"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            dr.Close()

            Me.gridMainDbCommand.Parameters.Clear()

            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", _temp.ModifiedBy)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", _temp.ModifiedDate)


            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Approved")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Date.Now())
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)

            gridMainDbCommand.CommandText = "  INSERT INTO [dbo].[tblAMSkillPreference] ([EmpNo],[IdSkill],[ModifiedBy],[ModifiedDate])
                                            SELECT [EmpNo],[IdSkill],[ModifiedBy],[ModifiedDate]
                                            FROM [dbo].[tblAMSkillPreferenceLog]
                                            WHERE [ModifiedBy]=@ModifiedBy AND [ModifiedDate]=@ModifiedDate AND [EmpNo]=@EmpNo"





            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMSkillPreferenceLog] SET ReviewBy=@ReviewBy,ReviewDate=@ReviewDate,Action=@Action WHERE EmpNo=@EmpNo AND [ModifiedBy]=@ModifiedBy AND [ModifiedDate]=@ModifiedDate ;"
                    Me.gridMainDbCommand.ExecuteNonQuery()
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function


    Public Function SaveReviewedLog(ByVal _temp As CoreFunctionInfo) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then




            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Reviewed")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)

            gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunction] SET ReviewBy=@ReviewBy,ReviewDate=@ReviewDate WHERE EmpNo=@EmpNo;"


            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True

                End If
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Me.CloseMainDbConnection()



        Else
            MsgBox("Failed to open main database. Please check you connection.")
        End If



        Return temp
    End Function


End Class

Public Class AgentInfo
    Implements IDisposable

    Public Property TeamId As Integer
    Public Property EmpNo As String
    Public Property Domestic As Boolean
    Public Property EmpName As String
    Public Property Supervisor As String
    Public Property Manager As String


    Public Property CoreFunction As CoreFunctionInfo
    Public Property ChanelInfo As ChannelInfo
    Public Property HomeSkillInfo As HomeSkillInfo
    Public Property SkillPreference As List(Of SkillPreference)



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

Public Class AgentViewInfo
    Implements IDisposable

    Public Property TeamId As Integer
    Public Property EmpNo As String
    Public Property Domestic As Boolean
    Public Property EmpName As String
    Public Property Supervisor As String
    Public Property Manager As String
    Public Property Status As Boolean


    Public Property OrderProcess As String
    Public Property Credit As String
    Public Property Rebill As String
    Public Property Returns As String
    Public Property Complaints As String
    Public Property ShippingDiscrepancy As String
    Public Property DocumentRequests As String
    Public Property PricingProduct As String
    Public Property Implementation As String
    Public Property ValueLinkTrained As String


    Public Property Router As String
    Public Property Phone As String
    Public Property Email As String
    Public Property Cases As String


    Public Property Commercial As Boolean
    Public Property DCS As Boolean
    Public Property GoV As Boolean
    Public Property Speciality As Boolean
    Public Property HSRouter As Boolean

    Public Property SkillPreference As List(Of SkillPreference)



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

Public Class AgentLogInfo
    Implements IDisposable

    Public Property TeamId As Integer
    Public Property EmpNo As String
    Public Property Domestic As Boolean
    Public Property EmpName As String
    Public Property Supervisor As String
    Public Property Manager As String


    Public Property OrderProcess As String
    Public Property Credit As String
    Public Property Rebill As String
    Public Property Returns As String
    Public Property Complaints As String
    Public Property ShippingDiscrepancy As String
    Public Property DocumentRequests As String
    Public Property PricingProduct As String
    Public Property Implementation As String
    Public Property ValueLinkTrained As String


    Public Property Router As String
    Public Property Phone As String
    Public Property Email As String
    Public Property Cases As String


    Public Property Commercial As Boolean
    Public Property DCS As Boolean
    Public Property GoV As Boolean
    Public Property Speciality As Boolean
    Public Property HSRouter As Boolean

    Public Property Action As String
    Public Property ModifiedBy As String
    Public Property ModifiedDate As DateTime?
    Public Property ReviewBy As String
    Public Property ReviewDate As DateTime?


    Public Property SkillPreference As List(Of SkillPreference)



#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

Public Class CoreFunctionInfo
    Implements IDisposable


    Public Property EmpNo As Integer
    Public Property OrderProcess As String
    Public Property Credit As String
    Public Property Rebill As String
    Public Property Returns As String
    Public Property Complaints As String
    Public Property ShippingDiscrepancy As String
    Public Property DocumentRequests As String
    Public Property PricingProduct As String
    Public Property Implementation As String
    Public Property CarrierDisposition As String
    Public Property AccountMaintenance As String
    Public Property SpecialProcessOrders As String
    Public Property Research As String


    Public Property ValueLinkTrained As String
    Public Property Comment As String
    Public Property ModifiedBy As String
    Public Property ModifiedDate As DateTime

#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
Public Class HomeSkillInfo
    Implements IDisposable


    Public Property EmpNo As Integer
    Public Property BackOffice As Boolean
    Public Property CHI As Boolean
    Public Property Concierge As Boolean
    Public Property Commercial As Boolean
    Public Property DCS As Boolean
    Public Property GOV As Boolean
    Public Property IDNC As Boolean
    Public Property Kaiser As Boolean
    Public Property Router As Boolean
    Public Property SalesSupport As Boolean
    Public Property Specialty As Boolean
    Public Property Tradex As Boolean
    Public Property SupplyAssurance As Boolean
    Public Property CET As Boolean


    Public Property ModifiedBy As String
    Public Property ModifiedDate As DateTime


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
Public Class ChannelInfo
    Implements IDisposable


    Public Property EmpNo As Integer
    Public Property Router As String
    Public Property Phone As String
    Public Property Email As String
    Public Property Cases As String

    Public Property PhonePrio As String
    Public Property CasesEmailPrio As String
    Public Property Reason As String
    Public Property ModifiedBy As String
    Public Property ModifiedDate As DateTime


#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class
Public Class SkillPreference
    Implements IDisposable


    Public Property EmpNo As String
    Public Property Skill As Integer
    Public Property ModifiedBy As String
    Public Property ModifiedDate As DateTime




#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls



    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If



            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        disposedValue = True
    End Sub



    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub



    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        ' TODO: uncomment the following line if Finalize() is overridden above.
        ' GC.SuppressFinalize(Me)
    End Sub
#End Region
End Class

