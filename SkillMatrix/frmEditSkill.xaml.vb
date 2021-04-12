Imports System.Data
Imports System.IO
Imports Telerik.Windows.Controls

Public Class frmEditSkill
    Dim a, b, c, d As New Boolean
    'Dim agentLst
    Dim sm As SkillMatrixClass

    Dim lstEmpNo As List(Of String)
    Public Sub New(ByVal _lstEmpNo As List(Of String))
        Dim sm As New SkillMatrixClass
        Dim dt, dt2 As New DataTable
        Dim str1 As String
        a = 1
        b = 0
        c = 0
        d = 0

        lstEmpNo = _lstEmpNo
        ' This call is required by the designer.
        InitializeComponent()
        str1 = "("
        For Each itm In lstEmpNo
            If str1 = "(" Then
                str1 = str1 & "'" & itm & "'"
            Else
                str1 = str1 & "," & "'" & itm & "'"
            End If

        Next
        str1 = str1 & ")"


        dt = sm.GetEmpNameList(str1)

        lstEmployee.Items.Clear()


        '<Label Content="Test" Foreground="White" BorderBrush="White"  HorizontalContentAlignment="Left" VerticalContentAlignment="Center" Height="30" Width="170" Margin="5,0" FontSize="11" />
        For i = 0 To dt.Rows.Count - 1
            Dim x As New System.Windows.Controls.Label
            x.Content = dt.Rows(i).Item("EmpName")
            x.Name = "E" & dt.Rows(i).Item("EmpNo")
            x.Width = 170
            x.Height = 30
            x.FontSize = 11
            x.HorizontalContentAlignment = Windows.HorizontalAlignment.Left
            x.VerticalContentAlignment = Windows.VerticalAlignment.Center
            x.VerticalAlignment = Windows.VerticalAlignment.Center
            x.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
            x.Foreground = Windows.Media.Brushes.White
            x.Margin = New Thickness(5, 0, 5, 0)
            Me.RegisterName(x.Name, x)

            lstEmployee.Items.Add(x)
        Next

        ' Add any initialization after the InitializeComponent() call.
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



        'lstPrimarySkill.Items.Clear()


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
        cbf()
    End Sub

    Private Sub cbf()

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


        stkCBFunction.Visibility = Visibility.Visible
        stkHomeSkill.Visibility = Visibility.Collapsed
        stkChannel.Visibility = Visibility.Collapsed
        stkSpecSkill.Visibility = Visibility.Collapsed

        grdCFLogs.Visibility = Visibility.Collapsed
        grdCHLogs.Visibility = Visibility.Collapsed
        grdHSLogs.Visibility = Visibility.Collapsed
        grdSPLogs.Visibility = Visibility.Collapsed

        If gvCFLogs.Items.Count > 0 Then
            gvCFLogs.Columns("Id").IsVisible = False
            gvCFLogs.Columns("EmpNo").IsVisible = False
            gvCFLogs.Columns("TeamId").IsVisible = False
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

        stkCBFunction.Visibility = Visibility.Collapsed
        stkHomeSkill.Visibility = Visibility.Visible
        stkChannel.Visibility = Visibility.Collapsed
        stkSpecSkill.Visibility = Visibility.Collapsed

        grdCFLogs.Visibility = Visibility.Collapsed
        grdCHLogs.Visibility = Visibility.Collapsed
        grdHSLogs.Visibility = Visibility.Collapsed
        grdSPLogs.Visibility = Visibility.Collapsed

        If gvHSLogs.Items.Count > 0 Then
            gvHSLogs.Columns("Id").IsVisible = False
            gvHSLogs.Columns("EmpNo").IsVisible = False
            gvHSLogs.Columns("TeamId").IsVisible = False
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


        stkCBFunction.Visibility = Visibility.Collapsed
        stkHomeSkill.Visibility = Visibility.Collapsed
        stkChannel.Visibility = Visibility.Visible
        stkSpecSkill.Visibility = Visibility.Collapsed

        grdCFLogs.Visibility = Visibility.Collapsed
        grdCHLogs.Visibility = Visibility.Collapsed
        grdHSLogs.Visibility = Visibility.Collapsed
        grdSPLogs.Visibility = Visibility.Collapsed

        If gvChLogs.Items.Count > 0 Then
            gvChLogs.Columns("Id").IsVisible = False
            gvChLogs.Columns("EmpNo").IsVisible = False
            gvChLogs.Columns("TeamId").IsVisible = False
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


        stkCBFunction.Visibility = Visibility.Collapsed
            stkHomeSkill.Visibility = Visibility.Collapsed
            stkChannel.Visibility = Visibility.Collapsed
            stkSpecSkill.Visibility = Visibility.Visible

            grdCFLogs.Visibility = Visibility.Collapsed
            grdCHLogs.Visibility = Visibility.Collapsed
            grdHSLogs.Visibility = Visibility.Collapsed
            grdSPLogs.Visibility = Visibility.Collapsed


        If gvSPLogs.Items.Count > 0 Then
            'gvSPLogs.Columns("Id").IsVisible = False
            gvSPLogs.Columns("EmpNo").IsVisible = False
            gvSPLogs.Columns("TeamId").IsVisible = False
        End If


    End Sub

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

    Private Function GetSkillId() As List(Of Integer)
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

        Return lstSkillId
    End Function

    Private Sub btnChange_Click(sender As Object, e As RoutedEventArgs) Handles btnChange.Click
        Dim lstPrimarySkills As New List(Of Integer)
        Dim _str As String
        Dim dt As DataTable
        lstPrimarySkills = GetSkillId()

        _str = "("
        For Each itm In lstPrimarySkills
            If _str = "(" Then
                _str = _str & "'" & itm & "'"
            Else
                _str = _str & "," & "'" & itm & "'"
            End If

        Next
        _str = _str & ")"

        sm = New SkillMatrixClass

        dt = sm.GetPrimarySkillInfo(_str)
        Dim frm As New frmChangePrimary(dt)
        frm.ShowDialog()
        lblPrimarySkill.Content = frm.lblPrimarySkill.Content

        'stkChange.Visibility = Visibility.Visible
        'stkPrimarySkill.Visibility = Visibility.Visible
    End Sub



    Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click
        'Dim agentLst
        Dim date1 As New DateTime
        date1 = Format(Date.Now(), "MM/dd/yyyy hh:mm")
        'agentLst = gvAgents.SelectedItem
        Dim _cfCount, _hsCount, _chCount, _spCount As New Integer
        _cfCount = 0
        _hsCount = 0
        _chCount = 0
        _spCount = 0

        For Each emp In lstEmpNo
            Dim tempCF = New CoreFunctionInfo
            Dim tempHS = New HomeSkillInfo
            Dim tempCH = New ChannelInfo
            Dim tempSP = New List(Of SkillPreference)
            Dim _EmpNo = emp

            'Try

            'Catch ex As Exception
            '    Dim agentLst As DataRow
            'End Try
            'agentLst = gvAgents.SelectedItem

            Dim CF, HS, CH, SP As Boolean

            If a Then
                tempCF.EmpNo = _EmpNo
                tempCF.OrderProcess = If(tab1CmbOrderProcess.Text = "", 0, tab1CmbOrderProcess.SelectedIndex)
                tempCF.Credit = If(tab1CmbCredit.Text = "", 0, tab1CmbCredit.SelectedIndex)
                tempCF.Rebill = If(tab1CmbRebill.Text = "", 0, tab1CmbRebill.SelectedIndex)
                tempCF.Returns = If(tab1CmbReturn.Text = "", 0, tab1CmbReturn.SelectedIndex)
                tempCF.Complaints = If(tab1CmbComplaints.Text = "", 0, tab1CmbComplaints.SelectedIndex)
                tempCF.PricingProduct = If(tab1CmbPrcProd.Text = "", 0, tab1CmbPrcProd.SelectedIndex)
                tempCF.ShippingDiscrepancy = If(tab1CmbPrcProd.Text = "", 0, tab1CmbPrcProd.SelectedIndex)
                tempCF.DocumentRequests = If(tab1CmbDocumentRequest.Text = "", 0, tab1CmbDocumentRequest.SelectedIndex)
                tempCF.Implementation = If(tab1CmbImplementation.Text = "", 0, tab1CmbImplementation.SelectedIndex)
                tempCF.CarrierDisposition = If(tab1CmbCarrierDisposition.Text = "", 0, tab1CmbCarrierDisposition.SelectedIndex)
                tempCF.AccountMaintenance = If(tab1CmbAccountMaintenance.Text = "", 0, tab1CmbAccountMaintenance.SelectedIndex)
                tempCF.SpecialProcessOrders = If(tab1CmbSPO.Text = "", 0, tab1CmbSPO.SelectedIndex)
                tempCF.Research = If(tab1CmbResearch.Text = "", 0, tab1CmbResearch.SelectedIndex)
                tempCF.PrimarySkill = lblPrimarySkill.Content
                tempCF.ValueLinkTrained = If(cmbVLTrained.SelectedIndex = 1, 1, 0)
                tempCF.Xipay = If(cmbXiPay.SelectedIndex = 1, 1, 0)

                tempCF.Comment = tab1TxtComment.Text


                CF = sm.SaveCFInfo(tempCF)
                If CF Then
                    _cfCount = _cfCount + 1
                    'MsgBox("Sucess Saving Core Business Function")
                    'a = 0
                    'b = 1
                    'c = 0
                    'd = 0
                    'tabCHL.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                    'tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                    'tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                    'stkCBFunction.Visibility = Visibility.Hidden
                    'stkHomeSkill.Visibility = Visibility.Hidden
                    'stkChannel.Visibility = Visibility.Visible
                    'stkSpecSkill.Visibility = Visibility.Hidden

                    'Else
                    '    MsgBox("Failed Saving Core Business Function")
                End If

            ElseIf b Then
                If (tab3CmbPhonePRIO.Text <> "P1" Or tab3CmbEmailCasePrio.Text <> "P1" Or tab3CmbBackOfficePrio.Text <> "P1") And tab3TxtReason.Text = "" Then
                    MsgBox("Reason is required!", MsgBoxStyle.Critical, "ERROR")
                    Exit Sub
                End If

                tempCH.EmpNo = _EmpNo
                tempCH.Router = If(tab3CmbRouter.Text = "", 0, tab3CmbRouter.SelectedIndex)
                tempCH.Phone = If(tab3CmbRouter.Text = "", 0, tab3CmbRouter.SelectedIndex)
                tempCH.Email = If(tab3CmbEmail.Text = "", 0, tab3CmbEmail.SelectedIndex)
                tempCH.Cases = If(tab3CmbCase.Text = "", 0, tab3CmbCase.SelectedIndex)
                tempCH.BackOfficeProf = If(tab3CmbBackOffice.Text = "", 0, tab3CmbBackOffice.SelectedIndex)

                tempCH.PhonePrio = tab3CmbPhonePRIO.Text
                tempCH.CasesEmailPrio = tab3CmbEmailCasePrio.Text
                tempCH.BackOfficePrio = tab3CmbBackOfficePrio.Text
                tempCH.Reason = tab3TxtReason.Text

                CH = sm.SaveCHInfo(tempCH)
                If CH Then
                    _chCount = _chCount + 1
                    'MsgBox("Sucess Saving Channel")
                    'a = 0
                    'b = 0
                    'c = 1
                    'd = 0
                    'tabHSkill.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                    'tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                    'tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabSS.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                    'stkCBFunction.Visibility = Visibility.Hidden
                    'stkHomeSkill.Visibility = Visibility.Visible
                    'stkChannel.Visibility = Visibility.Hidden
                    'stkSpecSkill.Visibility = Visibility.Hidden

                    'Else
                    'MsgBox("Failed Saving Channel")
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
                    _hsCount = _hsCount + 1
                    'MsgBox("Sucess Saving Home Skill")
                    'a = 0
                    'b = 0
                    'c = 0
                    'd = 1
                    'tabSS.Background = New SolidColorBrush(Color.FromRgb(234, 57, 57))
                    'tabSS.BorderBrush = New SolidColorBrush(Color.FromRgb(234, 57, 57))

                    'tabCBF.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCBF.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabHSkill.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabHSkill.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCHL.Background = New SolidColorBrush(Color.FromRgb(35, 42, 46))
                    'tabCHL.BorderBrush = New SolidColorBrush(Color.FromRgb(35, 42, 46))

                    'stkCBFunction.Visibility = Visibility.Hidden
                    'stkHomeSkill.Visibility = Visibility.Hidden
                    'stkChannel.Visibility = Visibility.Hidden
                    'stkSpecSkill.Visibility = Visibility.Visible

                    'Else
                    '    MsgBox("Failed Saving Home Skill")
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

                SP = sm.SaveSPInfo(lstSkillId, _EmpNo, date1)
                If SP Then
                    _spCount = _spCount + 1
                    'MsgBox("Sucess Saving Skill Preference")

                    'grdSearch.Visibility = Visibility.Visible
                    'grdEdit.Visibility = Visibility.Collapsed
                    'grdNameInfo.Visibility = Visibility.Collapsed


                    'gvAgents.ItemsSource = sm.GetAgentInfo(cmbTower.Text, cmbDept.Text, cmbSegment.SelectedValue, cmbSegment.Text).DefaultView
                    'gvAgents.Columns("EmpNo").IsVisible = False
                End If
            End If
        Next
        If a Then
            MsgBox(_cfCount & " Changes has been made in Core Business")
        ElseIf b Then
            MsgBox(_chCount & " Changes has been made in Channel")
        ElseIf c Then
            MsgBox(_hsCount & " Changes has been made in HomeSkill")
        ElseIf d Then
            MsgBox(_spCount & " Changes has been made in Skil Preference")
        End If

    End Sub
End Class
