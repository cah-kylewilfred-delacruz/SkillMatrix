Imports System.Data

Class MainWindow
    Dim a, b, c, d, x, y As Boolean
    Dim sm As SkillMatrixClass
    Dim agentLst
    Dim dt, dt2 As New DataTable
    Dim SkillMatrixVersion As FileVersionInfo

    Public Sub New()
        sm = New SkillMatrixClass
        Dim dt, dt2 As New DataTable
        Dim v
        ' This call is required by the designer.
        InitializeComponent()

        SkillMatrixVersion = FileVersionInfo.GetVersionInfo(Environment.CurrentDirectory & "\Skill Matrix.exe")

        lblVersion.Content = " Skill Matrix v" & SkillMatrixVersion.FileVersion.ToString

        sm.SaveversionLog(SkillMatrixVersion.FileVersion.ToString)

        ' Add any initialization after the InitializeComponent() call.
        a = 1
        b = 0
        c = 0
        d = 0

        x = 1
        y = 0

        For Each itm In sm.GetTowerList
            cmbTower.Items.Add(itm)
        Next
        lstBO.Items.Clear()
        lstCHI.Items.Clear()
        lstComm.Items.Clear()
        lstConcierge.Items.Clear()
        lstDCS.Items.Clear()
        lstGov.Items.Clear()
        lstIDNC.Items.Clear()
        lstKaiser.Items.Clear()
        lstRouter.Items.Clear()
        lstSalesSupport.Items.Clear()
        lstSpecialty.Items.Clear()
        lstTradex.Items.Clear()
        lstSupplyAssurance.Items.Clear()
        lstCET.Items.Clear()



        lstPrimarySkill.Items.Clear()


        dt = sm.GetSkillInfo()

        If dt.Rows.Count > 0 Then
            For i = 0 To dt.Rows.Count - 1
                Dim x As New System.Windows.Controls.CheckBox

                x.Content = dt.Rows(i).Item("SkillName")
                x.Name = "ID" & dt.Rows(i).Item("Id")
                x.Width = 170
                x.Height = 30
                x.FontSize = 12
                x.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
                x.VerticalContentAlignment = Windows.VerticalAlignment.Center
                x.VerticalAlignment = Windows.VerticalAlignment.Center
                x.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
                x.Foreground = Windows.Media.Brushes.White
                x.Margin = New Thickness(5, 0, 5, 0)
                Me.RegisterName(x.Name, x)

                Select Case dt.Rows(i).Item("SkillHome")
                    Case "Back Office"
                        lstBO.Items.Add(x)
                    Case "CHI"
                        lstCHI.Items.Add(x)
                    Case "Commercial"
                        lstComm.Items.Add(x)
                    Case "Concierge"
                        lstConcierge.Items.Add(x)
                    Case "DCS"
                        lstDCS.Items.Add(x)
                    Case "GOV"
                        lstGov.Items.Add(x)
                    Case "IDNC"
                        lstIDNC.Items.Add(x)
                    Case "Kaiser"
                        lstKaiser.Items.Add(x)

                    Case "Router"
                        lstRouter.Items.Add(x)
                    Case "Sales Support"
                        lstSalesSupport.Items.Add(x)
                    Case "Specialty"
                        lstSpecialty.Items.Add(x)
                    Case "Tradex"
                        lstTradex.Items.Add(x)
                    Case "SupplyAssurance"
                        lstSupplyAssurance.Items.Add(x)
                    Case "CET"
                        lstCET.Items.Add(x)
                End Select


            Next
        End If



    End Sub


    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click

        'If cmbSegment.Text = "" Or cmbDept.Text = "" Or cmbTower.Text = "" Then
        '    MsgBox("Please Select Tower/Department/Segment")

        'Else
        grdSearch.Visibility = Visibility.Visible
        grdView.Visibility = Visibility.Visible
        grdEdit.Visibility = Visibility.Collapsed
        grdNameInfo.Visibility = Visibility.Collapsed


        'gvAgents.Columns("LastModifiedDate").IsVisible = False



        load()

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As RoutedEventArgs) Handles btnCancel.Click
        grdSearch.Visibility = Visibility.Visible
        grdEdit.Visibility = Visibility.Collapsed
        grdNameInfo.Visibility = Visibility.Collapsed
        btnSave.Visibility = Visibility.Visible
        btnSearch.Visibility = Visibility.Visible

        If y Then
            load()
        End If


        'gvAgents.Columns("EmpNo").IsVisible = False
    End Sub

    Private Sub btnEdit_Click(sender As Object, e As RoutedEventArgs) Handles btnEdit.Click
        gvEdit()
    End Sub

    Private Sub load()
        gvAgents.ItemsSource = sm.GetAgentInfo(cmbTower.Text, cmbDept.Text, cmbSegment.SelectedValue, cmbSegment.Text).DefaultView
        gvAgents.Columns("EmpNo").IsVisible = False
        gvAgents.Columns("Domestic").IsVisible = False
        gvAgents.Columns("TeamId").IsVisible = False
    End Sub


    Private Sub tabCBF_MouseEnter(sender As Object, e As MouseEventArgs) Handles tabCBF.MouseEnter
        tabCBF.Background = New SolidColorBrush(Color.FromRgb(43, 54, 60))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(43, 54, 60))
    End Sub

    Private Sub tabCBF_MouseLeave(sender As Object, e As MouseEventArgs) Handles tabCBF.MouseLeave
        If a Then
            tabCBF.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        Else
            tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        End If
    End Sub

    Private Sub tabCHL_MouseEnter(sender As Object, e As MouseEventArgs) Handles tabCHL.MouseEnter
        tabCHL.Background = New SolidColorBrush(Color.FromRgb(43, 54, 60))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(43, 54, 60))
    End Sub

    Private Sub tabCHL_MouseLeave(sender As Object, e As MouseEventArgs) Handles tabCHL.MouseLeave
        If b Then
            tabCHL.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        Else
            tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        End If
    End Sub

    Private Sub tabHSkill_MouseEnter(sender As Object, e As MouseEventArgs) Handles tabHSkill.MouseEnter
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(43, 54, 60))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(43, 54, 60))
    End Sub

    Private Sub tabHSkill_MouseLeave(sender As Object, e As MouseEventArgs) Handles tabHSkill.MouseLeave
        If c Then
            tabHSkill.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        Else
            tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        End If
    End Sub


    Private Sub tabSS_MouseEnter(sender As Object, e As MouseEventArgs) Handles tabSS.MouseEnter
        tabSS.Background = New SolidColorBrush(Color.FromRgb(43, 54, 60))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(43, 54, 60))
    End Sub

    Private Sub tabSS_MouseLeave(sender As Object, e As MouseEventArgs) Handles tabSS.MouseLeave
        If d Then
            tabSS.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        Else
            tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        End If
    End Sub
    Private Sub tabCBF_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles tabCBF.MouseDown

        a = 1
        b = 0
        c = 0
        d = 0

        tabCBF.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))


        If x Then
            stkCBFunction.Visibility = Visibility.Visible
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed

        Else

            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed

            grdCFLogs.Visibility = Visibility.Visible
            grdCHLogs.Visibility = Visibility.Hidden
            grdHSLogs.Visibility = Visibility.Hidden
            grdSPLogs.Visibility = Visibility.Hidden

            If gvCFLogs.Items.Count > 0 Then
                gvCFLogs.Columns("Id").IsVisible = False
                gvCFLogs.Columns("EmpNo").IsVisible = False
                gvCFLogs.Columns("TeamId").IsVisible = False
            End If
        End If
    End Sub
    Private Sub tabHSkill_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles tabHSkill.MouseDown
        a = 0
        b = 0
        c = 1
        d = 0
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        If x Then
            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Visible
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed
        Else
            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed


            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Visible
            grdSPLogs.Visibility = Visibility.Collapsed

            If gvHSLogs.Items.Count > 0 Then
                gvHSLogs.Columns("Id").IsVisible = False
                gvHSLogs.Columns("EmpNo").IsVisible = False
                gvHSLogs.Columns("TeamId").IsVisible = False
            End If
        End If


    End Sub

    Private Sub tabCHL_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles tabCHL.MouseDown
        a = 0
        b = 1
        c = 0
        d = 0
        tabCHL.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

        If x Then
            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Visible
            stkSpecSkill.Visibility = Visibility.Collapsed

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed

        Else
            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Visible
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed

            If gvChLogs.Items.Count > 0 Then
                gvChLogs.Columns("Id").IsVisible = False
                gvChLogs.Columns("EmpNo").IsVisible = False
                gvChLogs.Columns("TeamId").IsVisible = False
            End If
        End If



    End Sub

    Private Sub tabSS_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles tabSS.MouseDown
        a = 0
        b = 0
        c = 0
        d = 1
        tabSS.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))


        If x Then
            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Visible

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed
        Else

            stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Collapsed


            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Visible

            If gvSPLogs.Items.Count > 0 Then
                'gvSPLogs.Columns("Id").IsVisible = False
                gvSPLogs.Columns("EmpNo").IsVisible = False
                gvSPLogs.Columns("TeamId").IsVisible = False
            End If
        End If

    End Sub


    Private Sub cmbTower_DropDownClosed(sender As Object, e As EventArgs) Handles cmbTower.DropDownClosed
        If cmbTower.Text = "" Then

        Else
            Try
                cmbDept.Items.Clear()
                cmbSegment.ItemsSource = Nothing
            Catch ex As Exception

            End Try
            For Each itm In sm.GetDepartmentList(cmbTower.SelectedItem)
                cmbDept.Items.Add(itm)
            Next
        End If


    End Sub

    Private Sub cmbDept_DropDownClosed(sender As Object, e As EventArgs) Handles cmbDept.DropDownClosed
        If cmbDept.Text = "" Then

        Else
            cmbSegment.ItemsSource = sm.GetSegmentList(cmbTower.SelectedItem, cmbDept.SelectedItem).DefaultView
        End If
    End Sub

    Private Sub gvAgents_MouseDoubleClick(sender As Object, e As MouseButtonEventArgs) Handles gvAgents.MouseDoubleClick
        gvEdit()
    End Sub


    Private Sub gvEdit()

        Dim agentLst
        Dim cf As New CoreFunctionInfo
        Dim hs As New HomeSkillInfo
        Dim ch As New ChannelInfo
        Dim sp As New List(Of Integer)
        Dim sm As New SkillMatrixClass
        btnApprove.Visibility = Visibility.Collapsed
        btnSave.Visibility = Visibility.Visible
        agentLst = gvAgents.SelectedItem
        lstPrimarySkill.Items.Clear()



        a = 1
        b = 0
        c = 0
        d = 0

        x = 1
        y = 0
        tabCBF.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

        stkCBFunction.Visibility = Visibility.Visible
        stkHomeSkill.Visibility = Visibility.Collapsed
        stkChannel.Visibility = Visibility.Collapsed
        stkSpecSkill.Visibility = Visibility.Collapsed

        grdCFLogs.Visibility = Visibility.Collapsed
        grdCHLogs.Visibility = Visibility.Collapsed
        grdHSLogs.Visibility = Visibility.Collapsed
        grdSPLogs.Visibility = Visibility.Collapsed


        If Not IsNothing(agentLst) Then


            lblEmp.Content = agentLst.Item("Name")
            lblSup.Content = agentLst.Item("Supervisor")
            lblManager.Content = agentLst.Item("Manager")
            lblStatus.Content = If(agentLst.Item("Status") = True, "Active", "Inactive")
            lblDomestic.Content = If(agentLst.Item("Domestic") = True, "Yes", "No")

            sp = sm.GetSkillPrefInfo(agentLst.item("EmpNo"))

            Dim str As String = ""
            For Each itm In sp
                If str = "" Then
                    str = "(" & itm
                Else
                    str = str & "," & itm
                End If
            Next

            str = str & ")"

            Dim pschk As New DataTable

            pschk = sm.GetPrimarySkillInfo(str)

            If pschk.Rows.Count > 0 Then
                For i = 0 To pschk.Rows.Count - 1
                    Dim x As New System.Windows.Controls.CheckBox

                    x.Content = pschk.Rows(i).Item("SkillName")
                    x.Name = "IDPS" & pschk.Rows(i).Item("Id")
                    x.Width = 170
                    x.Height = 30
                    x.FontSize = 12
                    x.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
                    x.VerticalContentAlignment = Windows.VerticalAlignment.Center
                    x.VerticalAlignment = Windows.VerticalAlignment.Center
                    x.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
                    x.Foreground = Windows.Media.Brushes.White
                    x.Margin = New Thickness(5, 0, 5, 0)

                    Try
                        Me.RegisterName(x.Name, x)
                    Catch ex As Exception
                        Me.UnregisterName(x.Name)
                        Me.RegisterName(x.Name, x)
                    End Try


                    lstPrimarySkill.Items.Add(x)
                Next
            End If

            cf = sm.GetCFInfo(agentLst.Item("EmpNo"))

            tab1CmbOrderProcess.SelectedIndex = cf.OrderProcess
            tab1CmbCredit.SelectedIndex = cf.Credit
            tab1CmbRebill.SelectedIndex = cf.Rebill
            tab1CmbReturn.SelectedIndex = cf.Returns
            tab1CmbComplaints.SelectedIndex = cf.Complaints
            tab1CmbShipDisc.SelectedIndex = cf.ShippingDiscrepancy
            tab1CmbPrcProd.SelectedIndex = cf.PricingProduct
            tab1CmbDocumentRequest.SelectedIndex = cf.DocumentRequests
            tab1CmbImplementation.SelectedIndex = cf.Implementation
            tab1CmbCarrierDisposition.SelectedIndex = cf.CarrierDisposition
            tab1CmbAccountMaintenance.SelectedIndex = cf.AccountMaintenance
            tab1CmbSPO.SelectedIndex = cf.SpecialProcessOrders
            tab1CmbResearch.SelectedIndex = cf.Research
            cmbVLTrained.SelectedIndex = cf.ValueLinkTrained
            tab1TxtComment.Text = cf.Comment
            lblPrimarySkill.Content = cf.PrimarySkill

            hs = sm.GetHSInfo(agentLst.Item("EmpNo"))



            tab2ChkBackOffice.IsChecked = hs.BackOffice
            tab2ChkCHI.IsChecked = hs.CHI
            tab2ChkConcierge.IsChecked = hs.Concierge
            tab2ChkCommercial.IsChecked = hs.Commercial
            tab2ChkDCS.IsChecked = hs.DCS
            tab2ChkGOV.IsChecked = hs.GOV
            tab2ChkIDNC.IsChecked = hs.IDNC
            tab2ChkKaiser.IsChecked = hs.Kaiser
            tab2ChkRouter.IsChecked = hs.Router
            tab2ChkSalesSupport.IsChecked = hs.SalesSupport
            tab2ChkSpecialty.IsChecked = hs.Specialty
            tab2ChkTradex.IsChecked = hs.Tradex


            If tab2ChkBackOffice.IsChecked Then
                stkBackOffice.Visibility = Visibility.Visible
            End If
            If tab2ChkCHI.IsChecked Then
                stkCHI.Visibility = Visibility.Visible
            End If
            If tab2ChkConcierge.IsChecked Then
                stkConcierge.Visibility = Visibility.Visible
            End If
            If tab2ChkCommercial.IsChecked Then
                stkCommercial.Visibility = Visibility.Visible
            End If
            If tab2ChkDCS.IsChecked Then
                stkDCS.Visibility = Visibility.Visible
            End If
            If tab2ChkGOV.IsChecked Then
                stkGovernment.Visibility = Visibility.Visible
            End If
            If tab2ChkIDNC.IsChecked Then
                stkIDNC.Visibility = Visibility.Visible
            End If
            If tab2ChkKaiser.IsChecked Then
                stkKaiser.Visibility = Visibility.Visible
            End If
            If tab2ChkRouter.IsChecked Then
                stkRouter.Visibility = Visibility.Visible
            End If
            If tab2ChkSalesSupport.IsChecked Then
                stkSalesSupport.Visibility = Visibility.Visible
            End If
            If tab2ChkSpecialty.IsChecked Then
                stkSpecialty.Visibility = Visibility.Visible
            End If
            If tab2ChkTradex.IsChecked Then
                stkTradex.Visibility = Visibility.Visible
            End If


            ch = sm.GetCHInfo(agentLst.Item("EmpNo"))

            tab3CmbRouter.SelectedIndex = ch.Router
            tab3CmbPhone.SelectedIndex = ch.Phone
            tab3CmbEmail.SelectedIndex = ch.Email
            tab3CmbCase.SelectedIndex = ch.Cases
            tab3CmbBackOffice.SelectedIndex = ch.BackOfficeProf

            tab3CmbPhonePRIO.Text = ch.PhonePrio
            tab3CmbEmailCasePrio.Text = ch.CasesEmailPrio
            tab3CmbBackOfficePrio.Text = ch.BackOfficePrio
            tab3TxtReason.Text = ch.Reason


            'sp = sm.GetSkillPrefInfo(agentLst.Item("EmpNo"))

            If Not IsNothing(sp) Then
                If sp.Count > 0 Then
                    For Each item In sp
                        Dim cb As System.Windows.Controls.CheckBox = Me.FindName("ID" & item)
                        If Not IsNothing(cb) Then
                            cb.IsChecked = True
                        End If

                        Dim cb1 As System.Windows.Controls.CheckBox = Me.FindName("IDPS" & item)
                        If Not IsNothing(cb) Then
                            cb1.IsChecked = True
                        End If
                    Next
                End If
            End If



            grdSearch.Visibility = Visibility.Collapsed
            grdEdit.Visibility = Visibility.Visible
            grdNameInfo.Visibility = Visibility.Visible
        End If


    End Sub
    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        Dim agentLst

        agentLst = gvAgents.SelectedItem
        Dim tempCF = New CoreFunctionInfo
        Dim tempHS = New HomeSkillInfo
        Dim tempCH = New ChannelInfo
        Dim tempSP = New List(Of SkillPreference)
        Dim _EmpNo = agentLst.Item("EmpNo")

        'Try

        'Catch ex As Exception
        '    Dim agentLst As DataRow
        'End Try
        'agentLst = gvAgents.SelectedItem

        Dim CF, HS, CH, SP As Boolean

        If a Then
            tempCF.EmpNo = _EmpNo
            tempCF.OrderProcess = tab1CmbOrderProcess.SelectedIndex
            tempCF.Credit = tab1CmbCredit.SelectedIndex
            tempCF.Rebill = tab1CmbRebill.SelectedIndex
            tempCF.Returns = tab1CmbReturn.SelectedIndex
            tempCF.Complaints = tab1CmbComplaints.SelectedIndex
            tempCF.PricingProduct = tab1CmbPrcProd.SelectedIndex
            tempCF.ShippingDiscrepancy = tab1CmbShipDisc.SelectedIndex
            tempCF.DocumentRequests = tab1CmbDocumentRequest.SelectedIndex
            tempCF.Implementation = tab1CmbImplementation.SelectedIndex
            tempCF.CarrierDisposition = tab1CmbCarrierDisposition.SelectedIndex
            tempCF.AccountMaintenance = tab1CmbAccountMaintenance.SelectedIndex
            tempCF.SpecialProcessOrders = tab1CmbSPO.SelectedIndex
            tempCF.Research = tab1CmbResearch.SelectedIndex
            tempCF.PrimarySkill = lblPrimarySkill.Content
            tempCF.ValueLinkTrained = cmbVLTrained.SelectedIndex
            tempCF.Xipay = cmbXiPay.SelectedIndex

            tempCF.Comment = tab1TxtComment.Text


            CF = sm.SaveCFInfo(tempCF)
            If CF Then
                MsgBox("Sucess Saving Core Business Function")
                a = 0
                b = 1
                c = 0
                d = 0
                tabCHL.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                stkCBFunction.Visibility = Visibility.Hidden
                stkHomeSkill.Visibility = Visibility.Hidden
                stkChannel.Visibility = Visibility.Visible
                stkSpecSkill.Visibility = Visibility.Hidden

            Else
                MsgBox("FailedSaving Core Business Function")
            End If

        ElseIf b Then
            If (tab3CmbPhonePRIO.Text <> "P1" Or tab3CmbEmailCasePrio.Text <> "P1" Or tab3CmbBackOfficePrio.Text <> "P1") And tab3TxtReason.Text = "" Then
                MsgBox("Reason is required!", MsgBoxStyle.Critical, "ERROR")
                Exit Sub
            End If

            tempCH.EmpNo = _EmpNo
            tempCH.Router = tab3CmbRouter.SelectedIndex
            tempCH.Phone = tab3CmbPhone.SelectedIndex
            tempCH.Email = tab3CmbEmail.SelectedIndex
            tempCH.Cases = tab3CmbCase.SelectedIndex
            tempCH.BackOfficeProf = tab3CmbBackOffice.SelectedIndex

            tempCH.PhonePrio = tab3CmbPhonePRIO.Text
            tempCH.CasesEmailPrio = tab3CmbEmailCasePrio.Text
            tempCH.BackOfficePrio = tab3CmbBackOfficePrio.Text
            tempCH.Reason = tab3TxtReason.Text

            CH = sm.SaveCHInfo(tempCH)
            If CH Then
                MsgBox("Sucess Saving Channel")
                a = 0
                b = 0
                c = 1
                d = 0
                tabHSkill.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                stkCBFunction.Visibility = Visibility.Hidden
                stkHomeSkill.Visibility = Visibility.Visible
                stkChannel.Visibility = Visibility.Hidden
                stkSpecSkill.Visibility = Visibility.Hidden

            Else
                MsgBox("FailedSaving Saving Channel")
            End If


        ElseIf c Then
            tempHS.EmpNo = _EmpNo
            tempHS.BackOffice = tab2ChkBackOffice.IsChecked
            tempHS.CHI = tab2ChkCHI.IsChecked
            tempHS.Commercial = tab2ChkCommercial.IsChecked
            tempHS.Concierge = tab2ChkConcierge.IsChecked
            tempHS.Commercial = tab2ChkCommercial.IsChecked
            tempHS.DCS = tab2ChkDCS.IsChecked
            tempHS.GOV = tab2ChkGOV.IsChecked
            tempHS.IDNC = tab2ChkIDNC.IsChecked
            tempHS.Kaiser = tab2ChkKaiser.IsChecked
            tempHS.Router = tab2ChkRouter.IsChecked
            tempHS.SalesSupport = tab2ChkSalesSupport.IsChecked
            tempHS.Specialty = tab2ChkSpecialty.IsChecked
            tempHS.Tradex = tab2ChkTradex.IsChecked
            tempHS.SupplyAssurance = tab2ChkSupplyAssurance.IsChecked
            tempHS.CET = tab2ChkCET.IsChecked


            HS = sm.SaveHSInfo(tempHS)
            If HS Then
                MsgBox("Sucess Saving Home Skill")
                a = 0
                b = 0
                c = 0
                d = 1
                tabSS.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                stkCBFunction.Visibility = Visibility.Hidden
                stkHomeSkill.Visibility = Visibility.Hidden
                stkChannel.Visibility = Visibility.Hidden
                stkSpecSkill.Visibility = Visibility.Visible

            Else
                MsgBox("Failed Saving Home Skill")
            End If

        ElseIf d Then
            Dim lstSkillId As New List(Of Integer)


            If tab2ChkBackOffice.IsChecked Then
                For Each itm In lstBO.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If



            If tab2ChkCHI.IsChecked Then
                For Each itm In lstCHI.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            If tab2ChkCommercial.IsChecked Then
                For Each itm In lstComm.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If
            If tab2ChkConcierge.IsChecked Then
                For Each itm In lstConcierge.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If
            If tab2ChkDCS.IsChecked Then
                For Each itm In lstDCS.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If
            If tab2ChkGOV.IsChecked Then
                For Each itm In lstGov.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If
            If tab2ChkIDNC.IsChecked Then
                For Each itm In lstIDNC.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If



            If tab2ChkKaiser.IsChecked Then
                For Each itm In lstKaiser.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            If tab2ChkRouter.IsChecked Then
                For Each itm In lstRouter.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If



            If tab2ChkSalesSupport.IsChecked Then
                For Each itm In lstSalesSupport.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            If tab2ChkSpecialty.IsChecked Then
                For Each itm In lstSpecialty.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If
            If tab2ChkTradex.IsChecked Then
                For Each itm In lstTradex.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            If tab2ChkSupplyAssurance.IsChecked Then
                For Each itm In lstSupplyAssurance.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            If tab2ChkCET.IsChecked Then
                For Each itm In lstCET.Items

                    Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
                    If Not IsNothing(cb) Then
                        If cb.IsChecked Then
                            lstSkillId.Add(CType(itm.Name.ToString.Substring(2), Integer))
                        End If
                    End If

                Next
            End If

            SP = sm.SaveSPInfo(lstSkillId, agentLst.Item("EmpNo"))
            If SP Then
                MsgBox("Sucess Saving Skill Preference")

                grdSearch.Visibility = Visibility.Visible
                grdEdit.Visibility = Visibility.Collapsed
                grdNameInfo.Visibility = Visibility.Collapsed


                gvAgents.ItemsSource = sm.GetAgentInfo(cmbTower.Text, cmbDept.Text, cmbSegment.SelectedValue, cmbSegment.Text).DefaultView
                'gvAgents.Columns("EmpNo").IsVisible = False
            End If
        End If
    End Sub


    Private Sub btnLogs_Click(sender As Object, e As RoutedEventArgs) Handles btnLogs.Click
        'Dim agentlst
        x = 0
        y = 1

        a = 1
        b = 0
        c = 0
        d = 0
        stkLogs.Visibility = Visibility.Visible

        grdCFLogs.Visibility = Visibility.Visible
        grdSearch.Visibility = Visibility.Visible
        grdEdit.Visibility = Visibility.Visible
        grdNameInfo.Visibility = Visibility.Collapsed

        stkCBFunction.Visibility = Visibility.Collapsed
        stkHomeSkill.Visibility = Visibility.Collapsed
        stkChannel.Visibility = Visibility.Collapsed
        stkSpecSkill.Visibility = Visibility.Collapsed

        grdCFLogs.Visibility = Visibility.Visible
        grdCHLogs.Visibility = Visibility.Collapsed
        grdHSLogs.Visibility = Visibility.Collapsed
        grdSPLogs.Visibility = Visibility.Collapsed
        btnApprove.Visibility = Visibility.Visible
        btnSearch.Visibility = Visibility.Collapsed
        'agentlst = gvAgents.SelectedItem
        If Not IsNothing(cmbSegment.SelectedValue) Then
            gvCFLogs.ItemsSource = sm.GetCFLog(cmbSegment.SelectedValue).DefaultView
            gvCHLogs.ItemsSource = sm.GetCHLog(cmbSegment.SelectedValue).DefaultView
            gvHSLogs.ItemsSource = sm.GetHSLog(cmbSegment.SelectedValue).DefaultView
            gvSPLogs.ItemsSource = sm.GetSPLog(cmbSegment.SelectedValue).DefaultView


            If gvCFLogs.Items.Count > 0 Then
                gvCFLogs.Columns("Id").IsVisible = False
                gvCFLogs.Columns("EmpNo").IsVisible = False
                gvCFLogs.Columns("TeamId").IsVisible = False
            End If
            'If gvCHLogs.Items.Count > 0 Then
            '    gvCHLogs.Columns("Id").IsVisible = False
            '    gvCHLogs.Columns("EmpNo").IsVisible = False
            '    gvCHLogs.Columns("TeamId").IsVisible = False
            'End If
            'If gvHSLogs.Items.Count > 0 Then
            '    gvHSLogs.Columns("Id").IsVisible = False
            '    gvHSLogs.Columns("EmpNo").IsVisible = False
            '    gvHSLogs.Columns("TeamId").IsVisible = False
            'End If
            'If gvSPLogs.Items.Count > 0 Then
            '    'gvSPLogs.Columns("Id").IsVisible = False
            '    gvSPLogs.Columns("EmpNo").IsVisible = False
            '    gvSPLogs.Columns("TeamId").IsVisible = False
            'End If


            'lblEmp.Content = agentlst.Item("Name")
            'lblSup.Content = agentlst.Item("Supervisor")
            'lblManager.Content = agentlst.Item("Manager")
            'lblStatus.Content = If(agentlst.Item("Status") = True, "Active", "Inactive")
            'lblDomestic.Content = If(agentlst.Item("Domestic") = True, "Yes", "No")
            btnSave.Visibility = Visibility.Collapsed
        End If

        tabCBF.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
        tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

        tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
        'Else
        '    MsgBox("Please select an employee to view logs")
        'End If
    End Sub


    'Private Sub btnReviewed_Click(sender As Object, e As RoutedEventArgs) Handles btnReviewed.Click
    '    Dim agentLst As DataRow
    '    agentLst = gvAgents.SelectedItem
    '    Dim tempCF = New CoreFunctionInfo
    '    Dim _EmpNo = agentLst.Item("EmpNo")
    '    Dim CF As Boolean

    '    tempCF.EmpNo = _EmpNo

    '    CF = sm.SaveReviewedLog(tempCF)
    '    If CF Then
    '        MsgBox("Selected Item is Reviewed")
    '        gvAgents.ItemsSource = sm.GetAgentInfo(cmbSegment.SelectedValue)
    '        gvAgents.Columns("EmpNo").IsVisible = False
    '    End If
    'End Sub
    Private Sub tab2ChkBackOffice_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkBackOffice.Checked
        stkBackOffice.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkBackOffice_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkBackOffice.Unchecked
        stkBackOffice.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkCHI_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCHI.Checked
        stkCHI.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkCHI_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCHI.Unchecked
        stkCHI.Visibility = Visibility.Collapsed
    End Sub
    Private Sub tab2ChkCommercial_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCommercial.Checked
        stkCommercial.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkCommercial_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCommercial.Unchecked
        stkCommercial.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkConcierge_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkConcierge.Checked
        stkConcierge.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkConcierge_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkConcierge.Unchecked
        stkConcierge.Visibility = Visibility.Collapsed
    End Sub
    Private Sub tab2ChkDCS_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkDCS.Checked
        stkDCS.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkDCS_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkDCS.Unchecked
        stkDCS.Visibility = Visibility.Collapsed
    End Sub
    Private Sub tab2ChkGoV_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkGOV.Checked
        stkGovernment.Visibility = Visibility.Visible
    End Sub
    Private Sub tab2ChkGoV_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkGOV.Unchecked
        stkGovernment.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkIDNC_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkIDNC.Checked
        stkIDNC.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkIDNC_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkIDNC.Unchecked
        stkIDNC.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkKaiser_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkKaiser.Checked
        stkKaiser.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkKaiser_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkKaiser.Unchecked
        stkKaiser.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkRouter_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkRouter.Checked
        stkRouter.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkRouter_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkRouter.Unchecked
        stkRouter.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkSalesSupport_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSalesSupport.Checked
        stkSalesSupport.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkSalesSupport_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSalesSupport.Unchecked
        stkSalesSupport.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkSpecialty_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSpecialty.Checked
        stkSpecialty.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkSpecialty_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSpecialty.Unchecked
        stkSpecialty.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkTradex_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkTradex.Checked
        stkTradex.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkTradex_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkTradex.Unchecked
        stkTradex.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkSupplyAssurance_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSupplyAssurance.Checked
        stkSuplyAssurance.Visibility = Visibility.Visible
    End Sub
    Private Sub tab2ChkSupplyAssurance_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkSupplyAssurance.Unchecked
        stkSuplyAssurance.Visibility = Visibility.Collapsed
    End Sub

    Private Sub tab2ChkCET_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCET.Checked
        stkCET.Visibility = Visibility.Visible
    End Sub

    Private Sub tab2ChkCET_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkCET.Unchecked
        stkCET.Visibility = Visibility.Collapsed
    End Sub

    Private Sub btnApprove_Click(sender As Object, e As RoutedEventArgs) Handles btnApprove.Click
        'Dim agentLst
        'Dim drCF, drHS, drCH, drSP
        'agentLst = gvAgents.SelectedItem

        Dim n As Integer

        Dim tempHS = New HomeSkillInfo
        Dim tempCH = New ChannelInfo
        Dim tempSP = New List(Of SkillPreference)
        Dim tempSP2 As New SkillPreference
        'Dim _EmpNo = agentLst.Item("EmpNo")

        Dim CF, HS, CH, SP As Boolean

        If a Then
            'drCF = gvCFLogs.SelectedItems
            n = 0
            'For Each itm In drCF
            '    _lstId.Add(drCF.Item("ID"))
            'Next


            'For i = 0 To gvCFLogs.SelectedItems.Count - 1
            '    Dim obj As DataRow = gvCFLogs.SelectedItems(i)
            '    _lstId.Add(

            '    obj.Item("ResearchType") = cmbResearchType.Text
            '    obj.Item("Status") = cmbStatus.Text
            '    obj.Item("Notes") = txtNotes.Text
            '    obj.Item("PurchaseOrder") = txtPONo.Text
            '    obj.Item("GLorCostCenter") = txtGLorCostCenter.Text
            '    gvAudits.SelectedItems(i) = obj
            'Next




            For i = 0 To gvCFLogs.SelectedItems.Count - 1
                Dim obj As DataRowView = gvCFLogs.SelectedItems(i)
                Dim tempCF = New CoreFunctionInfo

                tempCF.EmpNo = obj.Item("EmpNo")
                If obj.Item("Order Process") = "" Then
                    tempCF.OrderProcess = 0
                ElseIf obj.Item("Order Process") = "New" Then
                    tempCF.OrderProcess = 1
                ElseIf obj.Item("Order Process") = "Emerging" Then
                    tempCF.OrderProcess = 2
                ElseIf obj.Item("Order Process") = "Stable" Then
                    tempCF.OrderProcess = 3
                Else
                    tempCF.OrderProcess = 4
                End If

                If obj.Item("Credit") = "" Then
                    tempCF.Credit = 0
                ElseIf obj.Item("Credit") = "New" Then
                    tempCF.Credit = 1
                ElseIf obj.Item("Credit") = "Emerging" Then
                    tempCF.Credit = 2
                ElseIf obj.Item("Credit") = "Stable" Then
                    tempCF.Credit = 3
                Else
                    tempCF.Credit = 4
                End If

                If obj.Item("Rebill") = "" Then
                    tempCF.Rebill = 0
                ElseIf obj.Item("Rebill") = "New" Then
                    tempCF.Rebill = 1
                ElseIf obj.Item("Rebill") = "Emerging" Then
                    tempCF.Rebill = 2
                ElseIf obj.Item("Rebill") = "Stable" Then
                    tempCF.Rebill = 3
                Else
                    tempCF.Rebill = 4
                End If

                If obj.Item("Returns") = "" Then
                    tempCF.Returns = 0
                ElseIf obj.Item("Returns") = "New" Then
                    tempCF.Returns = 1
                ElseIf obj.Item("Returns") = "Emerging" Then
                    tempCF.Returns = 2
                ElseIf obj.Item("Returns") = "Stable" Then
                    tempCF.Returns = 3
                Else
                    tempCF.Returns = 4
                End If

                If obj.Item("Complaints") = "" Then
                    tempCF.Complaints = 0
                ElseIf obj.Item("Complaints") = "New" Then
                    tempCF.Complaints = 1
                ElseIf obj.Item("Complaints") = "Emerging" Then
                    tempCF.Complaints = 2
                ElseIf obj.Item("Complaints") = "Stable" Then
                    tempCF.Complaints = 3
                Else
                    tempCF.Complaints = 4
                End If

                If obj.Item("Pricing Product") = "" Then
                    tempCF.PricingProduct = 0
                ElseIf obj.Item("Pricing Product") = "New" Then
                    tempCF.PricingProduct = 1
                ElseIf obj.Item("Pricing Product") = "Emerging" Then
                    tempCF.PricingProduct = 2
                ElseIf obj.Item("Pricing Product") = "Stable" Then
                    tempCF.PricingProduct = 3
                Else
                    tempCF.PricingProduct = 4
                End If

                If obj.Item("Shipping Discrepancy") = "" Then
                    tempCF.ShippingDiscrepancy = 0
                ElseIf obj.Item("Shipping Discrepancy") = "New" Then
                    tempCF.ShippingDiscrepancy = 1
                ElseIf obj.Item("Shipping Discrepancy") = "Emerging" Then
                    tempCF.ShippingDiscrepancy = 2
                ElseIf obj.Item("Shipping Discrepancy") = "Stable" Then
                    tempCF.ShippingDiscrepancy = 3
                Else
                    tempCF.ShippingDiscrepancy = 4
                End If

                If obj.Item("Document Request") = "" Then
                    tempCF.DocumentRequests = 0
                ElseIf obj.Item("Document Request") = "New" Then
                    tempCF.DocumentRequests = 1
                ElseIf obj.Item("Document Request") = "Emerging" Then
                    tempCF.DocumentRequests = 2
                ElseIf obj.Item("Document Request") = "Stable" Then
                    tempCF.DocumentRequests = 3
                Else
                    tempCF.DocumentRequests = 4
                End If

                If obj.Item("Implementation") = "" Then
                    tempCF.Implementation = 0
                ElseIf obj.Item("Implementation") = "New" Then
                    tempCF.Implementation = 1
                ElseIf obj.Item("Implementation") = "Emerging" Then
                    tempCF.Implementation = 2
                ElseIf obj.Item("Implementation") = "Stable" Then
                    tempCF.Implementation = 3
                Else
                    tempCF.Implementation = 4
                End If

                If obj.Item("CarrierDisposition") = "" Then
                    tempCF.CarrierDisposition = 0
                ElseIf obj.Item("CarrierDisposition") = "New" Then
                    tempCF.CarrierDisposition = 1
                ElseIf obj.Item("CarrierDisposition") = "Emerging" Then
                    tempCF.CarrierDisposition = 2
                ElseIf obj.Item("CarrierDisposition") = "Stable" Then
                    tempCF.CarrierDisposition = 3
                Else
                    tempCF.CarrierDisposition = 4
                End If


                If obj.Item("AccountMaintenance") = "" Then
                    tempCF.AccountMaintenance = 0
                ElseIf obj.Item("AccountMaintenance") = "New" Then
                    tempCF.AccountMaintenance = 1
                ElseIf obj.Item("AccountMaintenance") = "Emerging" Then
                    tempCF.AccountMaintenance = 2
                ElseIf obj.Item("AccountMaintenance") = "Stable" Then
                    tempCF.AccountMaintenance = 3
                Else
                    tempCF.AccountMaintenance = 4
                End If

                If obj.Item("SpecialProcessOrders") = "" Then
                    tempCF.SpecialProcessOrders = 0
                ElseIf obj.Item("SpecialProcessOrders") = "New" Then
                    tempCF.SpecialProcessOrders = 1
                ElseIf obj.Item("SpecialProcessOrders") = "Emerging" Then
                    tempCF.SpecialProcessOrders = 2
                ElseIf obj.Item("SpecialProcessOrders") = "Stable" Then
                    tempCF.SpecialProcessOrders = 3
                Else
                    tempCF.SpecialProcessOrders = 4
                End If

                If obj.Item("Research") = "" Then
                    tempCF.Research = 0
                ElseIf obj.Item("Research") = "New" Then
                    tempCF.Research = 1
                ElseIf obj.Item("Research") = "Emerging" Then
                    tempCF.Research = 2
                ElseIf obj.Item("Research") = "Stable" Then
                    tempCF.Research = 3
                Else
                    tempCF.Research = 4
                End If

                If obj.Item("ValueLinkTrained") = "No" Then
                    tempCF.ValueLinkTrained = 0
                Else
                    tempCF.ValueLinkTrained = 1
                End If

                If obj.Item("Xipay") = "No" Then
                    tempCF.Xipay = 0
                Else
                    tempCF.Xipay = 1
                End If


                tempCF.PrimarySkill = obj.Item("PrimarySkill")
                tempCF.Comment = obj.Item("Comment")
                tempCF.ModifiedBy = obj.Item("ModifiedBy")
                tempCF.ModifiedDate = obj.Item("ModifiedDate")

                CF = sm.SaveCFLog(tempCF, obj.Item("Id"))
                If CF Then
                    n += 1
                End If

            Next

            Dim str As String
            If n > 1 Then
                str = " Item"
            Else
                str = " Items"
            End If

            'tempCF.EmpNo = _EmpNo

            'If CF Then
            MsgBox(n & str & " Approved", MsgBoxStyle.Information, vbInformation)
            gvCFLogs.ItemsSource = sm.GetCFLog(cmbSegment.SelectedValue).DefaultView


            a = 0
            b = 1
            c = 0
            d = 0
            tabCHL.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

            tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

            grdCFLogs.Visibility = Visibility.Visible
            grdHSLogs.Visibility = Visibility.Hidden
            grdCHLogs.Visibility = Visibility.Hidden
            grdSPLogs.Visibility = Visibility.Hidden

            'Else
            '    MsgBox("Failed Saving Core Business Function")
            'End If

        ElseIf b Then
            n = 0
            For i = 0 To gvChLogs.SelectedItems.Count - 1
                Dim obj As DataRowView = gvChLogs.SelectedItems(i)
                tempCH.EmpNo = obj.Item("EmpNo")

                If IsDBNull(obj.Item("Router")) Then
                    tempCH.Router = 0
                ElseIf obj.Item("Router") = "New" Then
                    tempCH.Router = 1
                ElseIf obj.Item("Router") = "Emerging" Then
                    tempCH.Router = 2
                ElseIf obj.Item("Router") = "Stable" Then
                    tempCH.Router = 3
                Else
                    tempCH.Router = 4
                End If

                If IsDBNull(obj.Item("Phone")) Then
                    tempCH.Phone = 0
                ElseIf obj.Item("Phone") = "New" Then
                    tempCH.Phone = 1
                ElseIf obj.Item("Phone") = "Emerging" Then
                    tempCH.Phone = 2
                ElseIf obj.Item("Phone") = "Stable" Then
                    tempCH.Phone = 3
                Else
                    tempCH.Phone = 4
                End If

                If IsDBNull(obj.Item("Email")) Then
                    tempCH.Email = 0
                ElseIf obj.Item("Email") = "New" Then
                    tempCH.Email = 1
                ElseIf obj.Item("Email") = "Emerging" Then
                    tempCH.Email = 2
                ElseIf obj.Item("Email") = "Stable" Then
                    tempCH.Email = 3
                Else
                    tempCH.Email = 4
                End If

                If IsDBNull(obj.Item("Cases")) Then
                    tempCH.Cases = 0
                ElseIf obj.Item("Cases") = "New" Then
                    tempCH.Cases = 1
                ElseIf obj.Item("Cases") = "Emerging" Then
                    tempCH.Cases = 2
                ElseIf obj.Item("Cases") = "Stable" Then
                    tempCH.Cases = 3
                Else
                    tempCH.Cases = 4
                End If

                If IsDBNull(obj.Item("BackOfficeProf")) Then
                    tempCH.BackOfficeProf = 0
                ElseIf obj.Item("BackOffice") = "New" Then
                    tempCH.BackOfficeProf = 1
                ElseIf obj.Item("BackOffice") = "Emerging" Then
                    tempCH.BackOfficeProf = 2
                ElseIf obj.Item("BackOffice") = "Stable" Then
                    tempCH.BackOfficeProf = 3
                Else
                    tempCH.BackOfficeProf = 4
                End If

                tempCH.PhonePrio = If(IsDBNull(obj.Item("PhonePrio")), "", obj.Item("PhonePrio"))
                tempCH.CasesEmailPrio = If(IsDBNull(obj.Item("CasesEmailPrio")), "", obj.Item("CasesEmailPrio")) ' obj.Item("CasesEmailPrio")
                tempCH.BackOfficePrio = If(IsDBNull(obj.Item("BackOfficePrio")), "", obj.Item("BackOfficePrio")) ' obj.Item("BackOfficePrio")
                tempCH.Reason = If(IsDBNull(obj.Item("Reason")), "", obj.Item("Reason")) ' obj.Item("Reason")
                tempCH.ModifiedBy = obj.Item("ModifiedBy")
                tempCH.ModifiedDate = obj.Item("ModifiedDate")
                CH = sm.SaveCHLog(tempCH, obj.Item("Id"))
                If CH Then
                    n += 1
                End If
            Next
            Dim str As String
            If n > 1 Then
                str = " Items"
            Else
                str = " Item"
            End If
            'drCH = gvChLogs.SelectedItem
            'tempCH.EmpNo = _EmpNo

            'If CH Then
            MsgBox(n & str & " Approved", MsgBoxStyle.Information, vbInformation)
            gvChLogs.ItemsSource = sm.GetCHLog(cmbSegment.SelectedValue).DefaultView


            a = 0
            b = 0
            c = 1
            d = 0

            tabHSkill.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

            tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))



            grdCFLogs.Visibility = Visibility.Hidden
            grdHSLogs.Visibility = Visibility.Visible
            grdCHLogs.Visibility = Visibility.Hidden
            grdSPLogs.Visibility = Visibility.Hidden

            'Else
            '    MsgBox("Failed Saving Saving Channel")
            'End If


        ElseIf c Then
            n = 0
            For i = 0 To gvChLogs.SelectedItems.Count - 1
                Dim obj As DataRowView = gvChLogs.SelectedItems(i)

                tempHS.EmpNo = obj.Item("EmpNo")
                tempHS.BackOffice = obj.Item("BackOffice")
                tempHS.CHI = obj.Item("CHI")
                tempHS.Commercial = obj.Item("Commercial")
                tempHS.Concierge = obj.Item("Concierge")
                tempHS.Commercial = obj.Item("Commercial")
                tempHS.DCS = obj.Item("DCS")
                tempHS.GOV = obj.Item("GOV")
                tempHS.IDNC = obj.Item("IDNC")
                tempHS.Kaiser = obj.Item("Kaiser")
                tempHS.Router = obj.Item("Router")
                tempHS.SalesSupport = obj.Item("SalesSupport")
                tempHS.Specialty = obj.Item("Specialty")
                tempHS.Tradex = obj.Item("Tradex")
                tempHS.SupplyAssurance = obj.Item("SupplyAssurance")
                tempHS.CET = obj.Item("CET")
                tempHS.ModifiedBy = obj.Item("ModifiedBy")
                tempHS.ModifiedDate = obj.Item("ModifiedDate")

                HS = sm.SaveHSLog(tempHS, obj.Item("Id"))

                If HS Then
                    n += 1
                End If
            Next


            Dim str As String
            If n > 1 Then
                str = " Item"
            Else
                str = " Items"
            End If

            'drHS = gvHSLogs.SelectedItem
            ''tempHS.EmpNo = _EmpNo

            'If HS Then
            MsgBox(n & str & " Approved", MsgBoxStyle.Information, vbInformation)
            gvHSLogs.ItemsSource = sm.GetCHLog(cmbSegment.SelectedValue).DefaultView

            a = 0
            b = 0
            c = 0
            d = 1
            tabSS.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
            tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

            tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
            tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))



            grdCFLogs.Visibility = Visibility.Hidden
            grdHSLogs.Visibility = Visibility.Hidden
            grdCHLogs.Visibility = Visibility.Hidden
            grdSPLogs.Visibility = Visibility.Visible

            'Else
            'MsgBox("Failed Saving Home Skill")
            'End If

        ElseIf d Then
            n = 0
            For i = 0 To gvChLogs.SelectedItems.Count - 1
                Dim obj As DataRowView = gvChLogs.SelectedItems(i)

                tempSP2.EmpNo = obj.Item("EmpNo")

                tempSP2.ModifiedBy = obj.Item("ModifiedBy")
                tempSP2.ModifiedDate = obj.Item("ModifiedDate")
                SP = sm.SaveSPLog(tempSP2)
                If SP Then
                    n += 1
                End If
            Next


            Dim str As String
            If n > 1 Then
                str = " Items"
            Else
                str = " Item"
            End If



            MsgBox(n & str & " Approved", MsgBoxStyle.Information, vbInformation)

            grdSearch.Visibility = Visibility.Visible
            grdEdit.Visibility = Visibility.Collapsed
            grdNameInfo.Visibility = Visibility.Collapsed


            load()

            'gvAgents.ItemsSource = sm.GetAgentInfo(cmbTower.Text, cmbDept.Text, cmbSegment.SelectedValue, cmbSegment.Text).DefaultView
            'gvAgents.Columns("EmpNo").IsVisible = False


        End If







    End Sub

    Private Sub btnSummary_Click(sender As Object, e As RoutedEventArgs) Handles btnSummary.Click
        Dim frm As New frmSummary() 'cmbSegment.SelectedValue)
        frm.ShowDialog()
    End Sub

    Private Sub btnChange_Click(sender As Object, e As RoutedEventArgs) Handles btnChange.Click
        stkChange.Visibility = Visibility.Visible
        stkPrimarySkill.Visibility = Visibility.Visible
    End Sub

    Private Sub btnCncl_Click(sender As Object, e As RoutedEventArgs) Handles btnCncl.Click
        stkChange.Visibility = Visibility.Collapsed
        stkPrimarySkill.Visibility = Visibility.Collapsed
    End Sub

    Private Sub btnSet_Click(sender As Object, e As RoutedEventArgs) Handles btnSet.Click
        Dim str As String = ""
        For Each itm In lstPrimarySkill.Items
            Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
            If Not IsNothing(cb) Then
                If cb.IsChecked = True Then
                    If str = "" Then
                        str = itm.Content
                    Else
                        str = str & "/" & itm.Content
                    End If
                End If
            End If

        Next

        lblPrimarySkill.Content = str

        stkChange.Visibility = Visibility.Collapsed
        stkPrimarySkill.Visibility = Visibility.Collapsed

        MsgBox("Primary skill has been set. Please click save to apply changes", vbInformation, "Agent Primary Skills")
    End Sub

    Private Sub topBar_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles topBar.MouseDown
        DragMove()
    End Sub

    Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
        Close()
    End Sub

    Private Sub btnMax_Click(sender As Object, e As RoutedEventArgs) Handles btnMax.Click
        If Me.WindowState = WindowState.Maximized Then
            Me.WindowState = WindowState.Normal
        Else
            Me.WindowState = WindowState.Maximized
        End If
    End Sub
End Class
