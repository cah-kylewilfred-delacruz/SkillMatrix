Imports System.Data
Imports System.Data.SqlClient

Public Class SkillMatrixClass
    Public gridMainDbConnection As SqlConnection
    Public gridMainDbCommand As SqlCommand
    Public gridMainDbConnectionState As Boolean = False

    Public Function OpenMainDbConnection() As Boolean

        Me.gridMainDbConnection = New SqlConnection
        Me.gridMainDbCommand = New SqlCommand

        Me.gridMainDbConnection.ConnectionString = "Data Source=LWPF0RR5J1;" &
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

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [Tower] FROM [dbo].[vQuery_Team] ORDER BY [Tower];"

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    temp.Add(dr.Item("Tower"))

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
    Public Function GetDepartmentList(ByVal _tower As String) As List(Of String)

        Dim temp As New List(Of String)

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _tower)

            Me.gridMainDbCommand.CommandText = "SELECT DISTINCT [Department] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower ORDER BY [Department];"

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
    Public Function GetSegmentList(ByVal _Tower As String, ByVal _Dept As String) As DataTable

        Dim temp As New DataTable

        temp.Columns.Add("Id")
        temp.Columns.Add("Segment")

        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@Tower", _Tower)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Department", _Dept)
            If _Dept = "ALL" Then
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower  ORDER BY [Segment];"
            Else
                Me.gridMainDbCommand.CommandText = "SELECT [Id],[Segment] FROM [dbo].[vQuery_Team] WHERE [Tower]=@Tower AND [Department]=@Department ORDER BY [Segment];"

            End If

            Try
                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read
                    temp.Rows.Add(dr.Item("Id"), dr.Item("Segment"))


                End While

                dr.Close()

                'If temp.Rows.Count > 1 Then
                '    temp.Rows.Add(0, "-ALL-")
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
    Public Function GetAgentInfo(ByVal _seg As Integer) As DataTable
        Dim temp As New DataTable

        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_SkillViewInfo] WHERE [TeamId]=" & _seg & " ORDER BY [Name];", Me.gridMainDbConnection)
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
        If temp.Columns(0).ColumnName = "TeamId" Then
            temp.Columns.Remove("TeamId")
        End If




        Return temp
    End Function
    Public Function GetAgentLogInfo(ByVal _seg As Integer) As List(Of AgentLogInfo)

        Dim temp As New List(Of AgentLogInfo)
        If OpenMainDbConnection() = True Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@TeamId", _seg)

            Me.gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[vQuery_AgentInfoLog] WHERE [TeamId]=@TeamId ORDER BY [EmpName];"


            Try

                Dim dr As SqlDataReader
                dr = Me.gridMainDbCommand.ExecuteReader

                While dr.Read

                    Dim temp2 As New AgentLogInfo
                    temp2.TeamId = dr.Item("TeamId")
                    temp2.EmpNo = dr.Item("EmpNo")
                    temp2.EmpName = dr.Item("EmpName")
                    temp2.Domestic = dr.Item("Domestic")
                    temp2.Supervisor = If(IsDBNull(dr.Item("Supervisor")), "", dr.Item("Supervisor"))
                    temp2.Manager = If(IsDBNull(dr.Item("Manager")), "", dr.Item("Manager"))
                    temp2.OrderProcess = If(IsDBNull(dr.Item("OrderProcess")), "", dr.Item("OrderProcess"))
                    temp2.Credit = If(IsDBNull(dr.Item("Credit")), "", dr.Item("Credit"))
                    temp2.Rebill = If(IsDBNull(dr.Item("Rebill")), "", dr.Item("Rebill"))
                    temp2.Complaints = If(IsDBNull(dr.Item("Complaints")), "", dr.Item("Complaints"))
                    temp2.ShippingDiscrepancy = If(IsDBNull(dr.Item("ShippingDiscrepancy")), "", dr.Item("ShippingDiscrepancy"))
                    temp2.DocumentRequests = If(IsDBNull(dr.Item("DocumentRequest")), "", dr.Item("DocumentRequest"))
                    temp2.PricingProduct = If(IsDBNull(dr.Item("PricingProduct")), "", dr.Item("PricingProduct"))
                    temp2.Implementation = If(IsDBNull(dr.Item("Implementation")), "", dr.Item("Implementation"))
                    temp2.ValueLinkTrained = If(IsDBNull(dr.Item("ValueLinkTrained")), "", dr.Item("ValueLinkTrained"))
                    temp2.Router = If(IsDBNull(dr.Item("Router")), "", dr.Item("Router"))
                    temp2.Phone = If(IsDBNull(dr.Item("Phone")), "", dr.Item("Phone"))
                    temp2.Email = If(IsDBNull(dr.Item("Email")), "", dr.Item("Email"))
                    temp2.Cases = If(IsDBNull(dr.Item("Cases")), "", dr.Item("Cases"))
                    temp2.Commercial = If(IsDBNull(dr.Item("Commercial")), 0, dr.Item("Commercial"))
                    temp2.DCS = If(IsDBNull(dr.Item("DCS")), 0, dr.Item("DCS"))
                    temp2.GoV = If(IsDBNull(dr.Item("GoV")), 0, dr.Item("GoV"))
                    temp2.Speciality = If(IsDBNull(dr.Item("Speciality")), 0, dr.Item("Speciality"))
                    temp2.HSRouter = If(IsDBNull(dr.Item("HSRouter")), 0, dr.Item("HSRouter"))

                    temp2.Action = If(IsDBNull(dr.Item("Action")), "", dr.Item("Action"))
                    temp2.ModifiedBy = If(IsDBNull(dr.Item("ModifiedBy")), "", dr.Item("ModifiedBy"))
                    temp2.ModifiedDate = If(IsDBNull(dr.Item("ModifiedDate")), Nothing, dr.Item("ModifiedDate"))
                    temp2.ReviewBy = If(IsDBNull(dr.Item("ReviewBy")), "", dr.Item("ReviewBy"))
                    temp2.ReviewDate = If(IsDBNull(dr.Item("ReviewDate")), Nothing, dr.Item("ReviewDate"))

                    temp.Add(temp2)

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

    'GET LOGS'
    Public Function GetCFLog(ByVal _EmpNo As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_CoreFunctionLog] WHERE [EmpNo]=" & _EmpNo & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetHSLog(ByVal _EmpNo As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_HomeSkillLog] WHERE [EmpNo]=" & _EmpNo & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetCHLog(ByVal _EmpNo As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_ChannelLog] WHERE [EmpNo]=" & _EmpNo & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
        End If

        Return temp
    End Function
    Public Function GetSPLog(ByVal _EmpNo As String) As DataTable
        Dim temp As New DataTable


        If OpenMainDbConnection() = True Then
            Dim da As New SqlDataAdapter()
            da = New SqlDataAdapter("SELECT * FROM [dbo].[vQuery_SkillPreferenceLog] WHERE [EmpNo]=" & _EmpNo & " ORDER BY [ModifiedDate] DESC;", Me.gridMainDbConnection)
            da.SelectCommand.CommandTimeout = 1000
            Try
                da.Fill(temp)
            Catch ex As Exception
            End Try
            Me.CloseMainDbConnection()
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
                    temp.OrderProcess = If(IsNothing(dr.Item("OrderProcess")), 0, dr.Item("OrderProcess"))
                    temp.Credit = If(IsNothing(dr.Item("Credit")), 0, dr.Item("Credit"))
                    temp.Rebill = If(IsNothing(dr.Item("Rebill")), 0, dr.Item("Rebill"))
                    temp.Returns = If(IsNothing(dr.Item("Returns")), 0, dr.Item("Returns"))
                    temp.Complaints = If(IsNothing(dr.Item("Complaints")), 0, dr.Item("Complaints"))
                    temp.ShippingDiscrepancy = If(IsNothing(dr.Item("ShippingDiscrepancy")), 0, dr.Item("ShippingDiscrepancy"))
                    temp.DocumentRequests = If(IsNothing(dr.Item("DocumentRequests")), 0, dr.Item("DocumentRequests"))
                    temp.PricingProduct = If(IsNothing(dr.Item("PricingProduct")), 0, dr.Item("PricingProduct"))
                    temp.Implementation = If(IsNothing(dr.Item("Implementation")), 0, dr.Item("Implementation"))
                    temp.ValueLinkTrained = If(IsNothing(dr.Item("ValueLinkTrained")), 0, dr.Item("ValueLinkTrained"))
                    temp.Comment = If(IsNothing(dr.Item("Comment")), "", dr.Item("Comment"))
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


                    temp.BackOffice = If(IsNothing(dr.Item("BackOffice")), 0, dr.Item("BackOffice"))
                    temp.CHI = If(IsNothing(dr.Item("CHI")), 0, dr.Item("CHI"))
                    temp.Concierge = If(IsNothing(dr.Item("Concierge")), 0, dr.Item("Concierge"))
                    temp.Commercial = If(IsNothing(dr.Item("Commercial")), 0, dr.Item("Commercial"))
                    temp.GOV = If(IsNothing(dr.Item("GOV")), 0, dr.Item("GOV"))
                    temp.IDNC = If(IsNothing(dr.Item("IDNC")), 0, dr.Item("IDNC"))
                    temp.Kaiser = If(IsNothing(dr.Item("Kaiser")), 0, dr.Item("Kaiser"))
                    temp.Router = If(IsNothing(dr.Item("Router")), 0, dr.Item("Router"))
                    temp.SalesSupport = If(IsNothing(dr.Item("SalesSupport")), 0, dr.Item("SalesSupport"))
                    temp.Specialty = If(IsNothing(dr.Item("Specialty")), 0, dr.Item("Specialty"))
                    temp.Tradex = If(IsNothing(dr.Item("Tradex")), 0, dr.Item("Tradex"))

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
                    temp.Router = If(IsNothing(dr.Item("Router")), "", dr.Item("Router"))
                    temp.Phone = If(IsNothing(dr.Item("Phone")), "", dr.Item("Phone"))
                    temp.Email = If(IsNothing(dr.Item("Email")), "", dr.Item("Email"))
                    temp.Cases = If(IsNothing(dr.Item("Cases")), "", dr.Item("Cases"))
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
            da = New SqlDataAdapter("SELECT * FROM [dbo].[tblAMSkill] ORDER BY [Id];", Me.gridMainDbConnection)
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
            Me.gridMainDbCommand.Parameters.AddWithValue("@ValueLinkTrained", _temp.ValueLinkTrained)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Comment", _temp.Comment)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")

            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMCoreFunction] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunction] SET OrderProcess=@OrderProcess,Credit=@Credit,Rebill=@Rebill,Returns=@Returns,Complaints=@Complaints,ShippingDiscrepancy=@ShippingDiscrepancy,DocumentRequests=@DocumentRequests,PricingProduct=@PricingProduct,Implementation=@Implementation,ValueLinkTrained=@ValueLinkTrained,ModifiedDate=@ModifiedDate,Comment=@Comment WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMCoreFunction] ([EmpNo],[OrderProcess],[Credit],[Rebill],[Returns],[Complaints],[ShippingDiscrepancy],[DocumentRequests],[PricingProduct],[Implementation],[ValueLinkTrained],[Comment],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@OrderProcess,@Credit,@Rebill,@Returns,@Complaints,@ShippingDiscrepancy,@DocumentRequests,@PricingProduct,@Implementation,@ValueLinkTrained,@Comment,@ModifiedBy,@ModifiedDate)"
            End If

            dr.Close()
            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMCoreFunctionLog] ([EmpNo],[OrderProcess],[Credit],[Rebill],[Returns],[Complaints],[ShippingDiscrepancy],[DocumentRequests],[PricingProduct],[Implementation],[ValueLinkTrained],[Comment],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@OrderProcess,@Credit,@Rebill,@Returns,@Complaints,@ShippingDiscrepancy,@DocumentRequests,@PricingProduct,@Implementation,@ValueLinkTrained,@Comment,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"
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



            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")

            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMHomeSkill] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMHomeSkill] SET BackOffice=@BackOffice,CHI=@CHI,Concierge=@Concierge,Commercial=@Commercial,DCS=@DCS,GOV=@GOV,IDNC=@IDNC,Kaiser=@Kaiser,Router=@Router,SalesSupport=@SalesSupport,Specialty=@Specialty,Tradex=@Tradex,        ModifiedDate=@ModifiedDate,ModifiedBy=@ModifiedBy WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMHomeSkill] ([EmpNo],[Commercial],[DCS],[GOV],[Specialty],[Router],[BackOffice],[CHI],[Concierge],[IDNC],[Kaiser],[SalesSupport],[Tradex],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@Commercial,@DCS,@GOV,@Specialty,@Router,@BackOffice,@CHI,@Concierge,@IDNC,@Kaiser,@SalesSupport,@Tradex,@ModifiedBy,@ModifiedDate)"


            End If

            dr.Close()
            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMHomeSkillLog] ([EmpNo],[Commercial],[DCS],[GOV],[Specialty],[Router],[BackOffice],[CHI],[Concierge],[IDNC],[Kaiser],[SalesSupport],[Tradex],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@Commercial,@DCS,@GOV,@Specialty,@Router,@BackOffice,@CHI,@Concierge,@IDNC,@Kaiser,@SalesSupport,@Tradex,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"
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
    Public Function SaveCHInfo(ByVal _temp As ChannelInfo) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then


            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _temp.EmpNo)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Router", _temp.Router)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Phone", _temp.Phone)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Email", _temp.Email)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Cases", _temp.Cases)


            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")

            gridMainDbCommand.CommandText = "SELECT * FROM [dbo].[tblAMChannel] WHERE [EmpNo]=@EmpNo;"

            Dim dr As SqlDataReader = gridMainDbCommand.ExecuteReader

            If dr.Read Then
                gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMChannel] SET Router=@Router,Phone=@Phone,Email=@Email,Cases=@Cases,ModifiedDate=@ModifiedDate,ModifiedBy=@ModifiedBy WHERE EmpNo=@EmpNo;"
            Else
                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMChannel] ([EmpNo],[Router],[Phone],[Email],[Cases],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@Router,@Phone,@Email,@Cases,@ModifiedBy,@ModifiedDate)"


            End If
            dr.Close()

            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMChannelLog] ([EmpNo],[Router],[Phone],[Email],[Cases],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@Router,@Phone,@Email,@Cases,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"
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
    Public Function SaveSPInfo(ByRef _temp As List(Of Integer), ByVal _EmpNo As String) As Boolean
        Dim temp As Boolean
        If OpenMainDbConnection() Then
            Me.gridMainDbCommand.Parameters.Clear()
            Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)

            gridMainDbCommand.CommandText = "DELETE FROM [dbo].[tblAMSkillPreference] WHERE EmpNo=@EmpNo;"
            Me.gridMainDbCommand.ExecuteNonQuery()

            For Each itm In _temp
                Me.gridMainDbCommand.Parameters.Clear()
                Me.gridMainDbCommand.Parameters.AddWithValue("@EmpNo", _EmpNo)
                Me.gridMainDbCommand.Parameters.AddWithValue("@IdSkill", itm)

                Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedDate", Now.ToString)
                Me.gridMainDbCommand.Parameters.AddWithValue("@ModifiedBy", Environment.UserName)

                Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "")

                Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", DBNull.Value)
                Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", "")


                gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMSkillPreference] ([EmpNo],[IdSkill],[ModifiedBy],[ModifiedDate]) VALUES (@EmpNo,@IdSkill,@ModifiedBy,@ModifiedDate)"

                Try
                    If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                        temp = True
                        gridMainDbCommand.CommandText = "INSERT INTO [dbo].[tblAMSkillPreferenceLog] ([EmpNo],[IdSkill],[Action],[ModifiedBy],[ModifiedDate],[ReviewBy],[ReviewDate]) VALUES (@EmpNo,@IdSkill,@Action,@ModifiedBy,@ModifiedDate,@ReviewBy,@ReviewDate)"
                        Me.gridMainDbCommand.ExecuteNonQuery()
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
            Me.gridMainDbCommand.Parameters.AddWithValue("@ValueLinkTrained", _temp.ValueLinkTrained)
            Me.gridMainDbCommand.Parameters.AddWithValue("@Comment", _temp.Comment)



            Me.gridMainDbCommand.Parameters.AddWithValue("@Action", "Approved")

            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewDate", Date.Now)
            Me.gridMainDbCommand.Parameters.AddWithValue("@ReviewBy", Environment.UserName)


            gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunction] SET OrderProcess=@OrderProcess,Credit=@Credit,Rebill=@Rebill,Returns=@Returns,Complaints=@Complaints,ShippingDiscrepancy=@ShippingDiscrepancy,DocumentRequests=@DocumentRequests,PricingProduct=@PricingProduct,Implementation=@Implementation,ValueLinkTrained=@ValueLinkTrained,ModifiedDate=@ModifiedDate,Comment=@Comment WHERE EmpNo=@EmpNo;"

            Try
                If Me.gridMainDbCommand.ExecuteNonQuery() > 0 Then
                    temp = True
                    gridMainDbCommand.CommandText = "UPDATE [dbo].[tblAMCoreFunctionLog] SET OrderProcess=@OrderProcess,Credit=@Credit,Rebill=@Rebill,Returns=@Returns,Complaints=@Complaints,ShippingDiscrepancy=@ShippingDiscrepancy,DocumentRequests=@DocumentRequests,PricingProduct=@PricingProduct,Implementation=@Implementation,ValueLinkTrained=@ValueLinkTrained,ReviewDate=@ReviewDate,ReviewBy=@ReviewBy,Action=@Action WHERE Id=" & _Id & ";"
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

