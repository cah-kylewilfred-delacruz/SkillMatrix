Imports System.Data
Imports System.IO
Imports System.Windows.Forms
Imports Microsoft.Win32
Imports SkillMatrix.SkillMatrixClass
Imports Telerik.Windows.Controls
Imports SaveFileDialog = Microsoft.Win32.SaveFileDialog

Public Class frmSummary

    Dim sm = New SkillMatrix.SkillMatrixClass
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

        'For Each itm In sm.GetSkillList()
        '    cmbHomeSkill.Items.Add(itm)
        'Next
        Dim dt As New DataTable
        dt = sm.GetSkillInfo()
        gvSummary.ItemsSource = dt

        For Each itm In sm.GetSumTowerList
            cmbTower.Items.Add(itm)
        Next
    End Sub


    Private Sub cmbTower_DropDownClosed(sender As Object, e As EventArgs) Handles cmbTower.DropDownClosed
        If cmbTower.Text = "" Then

        Else
            Try
                cmbDept.Items.Clear()
                cmbSegment.ItemsSource = Nothing
            Catch ex As Exception

            End Try
            For Each itm In sm.GetSumDepartmentList(cmbTower.SelectedItem)
                cmbDept.Items.Add(itm)
            Next
        End If



    End Sub

    Private Sub cmbDept_DropDownClosed(sender As Object, e As EventArgs) Handles cmbDept.DropDownClosed
        If cmbDept.Text = "" Then

        Else
            cmbSegment.ItemsSource = sm.GetSumSegmentList(cmbTower.SelectedItem, cmbDept.SelectedItem).DefaultView
        End If
    End Sub





    'Private Sub chkBackOffice_Checked(sender As Object, e As RoutedEventArgs) Handles chkBackOffice.Checked
    '    Dim dt As New DataView
    '    dt = sm.GetSkillInfo().DefaultView
    '    dt.RowFilter = "SkillHome = " & "[Back Office]"
    '    gvSummary.ItemsSource = dt
    'End Sub
    'Private Sub topBar_MouseDown(sender As Object, e As MouseButtonEventArgs) Handles topBar.MouseDown
    '    DragMove()
    'End Sub

    'Private Sub btnClose_Click(sender As Object, e As RoutedEventArgs) Handles btnClose.Click
    '    Close()
    'End Sub

    'Private Sub btnMax_Click(sender As Object, e As RoutedEventArgs) Handles btnMax.Click
    '    If Me.WindowState = WindowState.Maximized Then
    '        Me.WindowState = WindowState.Normal
    '    Else
    '        Me.WindowState = WindowState.Maximized
    '    End If
    'End Sub

    Private Sub btnSearch_Click(sender As Object, e As RoutedEventArgs) Handles btnSearch.Click
        Dim tempCh As New List(Of String)


        Dim ChannelPrio As String = ""
        Dim CoreBiz As String = ""
        Dim ChannelProf As String = ""


        Dim temp As New List(Of String)
        gvSummary.ItemsSource = Nothing
        If chkBackOffice.IsChecked Then
            temp.Add("Back Office")
        End If
        If chkCHI.IsChecked Then
            temp.Add("CHI")
        End If
        If chkConcierge.IsChecked Then
            temp.Add("Concierge")
        End If
        If chkCommercial.IsChecked Then
            temp.Add("Commercial")
        End If
        If chkDCS.IsChecked Then
            temp.Add("DCS")
        End If
        If chkGov.IsChecked Then
            temp.Add("GOV")
        End If
        If chkIDNC.IsChecked Then
            temp.Add("IDNC")
        End If
        If chkKaiser.IsChecked Then
            temp.Add("Kaiser")
        End If
        If chkRouter.IsChecked Then
            temp.Add("Router")
        End If
        If chkSalesSupport.IsChecked Then
            temp.Add("Sales Support")
        End If
        If chkSpecialty.IsChecked Then
            temp.Add("Specialty")
        End If
        If chkCarrierDispo.IsChecked Then
            temp.Add("CarrierDisposition")
        End If
        If chkAccountMaintenance.IsChecked Then
            temp.Add("AccountMaintenance")
        End If
        If chkSPO.IsChecked Then
            temp.Add("SpecialProcessOrders")
        End If
        If chkResearch.IsChecked Then
            temp.Add("Research")
        End If



        'If chkPhone.IsChecked = False And chkCases.IsChecked = False And chkEmail.IsChecked = False And chkCHRouter.IsChecked = False Then
        '    MsgBox("Please Select a Channel", vbInformation)
        '    Exit Sub
        'End If

        'If chkP1.IsChecked = False And chkP2.IsChecked = False And chkP3.IsChecked = False And chkException.IsChecked = False And chkSystem.IsChecked = False And chkProject.IsChecked = False And chkLOA.IsChecked = False Then
        '    MsgBox("Please Select a Channel Profeciency", vbInformation)
        '    Exit Sub
        'End If



        If cmbTower.SelectedValue = "" And cmbSegment.SelectedValue = "" Then
            MsgBox("Please Select a select segment", vbInformation)
            Exit Sub
        End If

        If chkP1.IsChecked Then
            ChannelPrio = "('P1'"
        End If
        If chkP2.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('P2'"
            Else
                ChannelPrio = ChannelPrio & ",'P2'"
            End If
        End If
        If chkP3.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('P3'"
            Else
                ChannelPrio = ChannelPrio & ",'P3'"
            End If
        End If
        If chkException.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('Exception'"
            Else
                ChannelPrio = ChannelPrio & ",'Exception'"
            End If
        End If
        If chkSystem.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('System'"
            Else
                ChannelPrio = ChannelPrio & ",'System'"
            End If
        End If
        If chkProject.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('Project'"
            Else
                ChannelPrio = ChannelPrio & ",'Project'"
            End If
        End If
        If chkLOA.IsChecked Then
            If ChannelPrio = "" Then
                ChannelPrio = "('LOA'"
            Else
                ChannelPrio = ChannelPrio & ",'LOA'"
            End If
        End If



        If chkCHNew.IsChecked Then
            ChannelProf = "('New'"
        End If
        If chkCHEmerging.IsChecked Then
            If ChannelProf = "" Then
                ChannelProf = "('Emerging'"
            Else
                ChannelProf = ChannelProf & ",'Emerging'"
            End If
        End If
        If chkCHStable.IsChecked Then
            If ChannelProf = "" Then
                ChannelProf = "('Stable'"
            Else
                ChannelProf = ChannelProf & ",'Stable'"
            End If
        End If
        If chkCHMaster.IsChecked Then
            If ChannelProf = "" Then
                ChannelProf = "('Master'"
            Else
                ChannelProf = ChannelProf & ",'Master'"
            End If
        End If



        ChannelPrio = ChannelPrio & ")"
        ChannelProf = ChannelProf & ")"

        If chkNew.IsChecked Then
            CoreBiz = "('New'"
        End If
        If chkEmerging.IsChecked Then
            If CoreBiz = "" Then
                CoreBiz = "('Emerging'"
            Else
                CoreBiz = CoreBiz & ",'Emerging'"
            End If
        End If
        If chkStable.IsChecked Then
            If CoreBiz = "" Then
                CoreBiz = "('Stable'"
            Else
                CoreBiz = CoreBiz & ",'Stable'"
            End If
        End If
        If chkMaster.IsChecked Then
            If CoreBiz = "" Then
                CoreBiz = "('Master'"
            Else
                CoreBiz = CoreBiz & ",'Master'"
            End If
        End If

        CoreBiz = CoreBiz & ")"
        Dim str As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""



        If chkPhonePrio.IsChecked Then
            str4 = "[PhonePrio] IN " & ChannelPrio
        End If
        If chkCasesEmailPrio.IsChecked Then
            If str4 = "" Then
                str4 = "[CasesEmailPrio] IN " & ChannelPrio
            Else str4 = str4 & " OR " & "[CasesEmailPrio] IN " & ChannelPrio
            End If
        End If
        If chkBackOfficePrio.IsChecked Then
            If str4 = "" Then
                str4 = "[BackOfficePrio] IN " & ChannelPrio
            Else str4 = str4 & " OR " & "[BackOfficePrio] IN " & ChannelPrio
            End If
        End If
        If chkCHRouterPrio.IsChecked Then
            If str4 = "" Then
                str4 = "[RouterPrio] IN " & ChannelPrio
            Else str4 = str4 & " OR " & "[RouterPrio] IN " & ChannelPrio
            End If
        End If


        If chkPhone.IsChecked Then
            str = "[Phone] IN " & ChannelProf
        End If
        If chkEmail.IsChecked Then
            If str = "" Then
                str = "[Email] IN " & ChannelProf
            Else str = str & " OR " & "[Email] IN " & ChannelProf
            End If
        End If
        If chkCases.IsChecked Then
            If str = "" Then
                str = "[Cases] IN " & ChannelProf
            Else str = str & " OR " & "[Cases] IN " & ChannelProf
            End If
        End If
        If chkCHRouter.IsChecked Then
            If str = "" Then
                str = "[Router] IN " & ChannelProf
            Else str = str & " OR " & "[Router] IN " & ChannelProf
            End If
        End If

        If chkCHBackOffice.IsChecked Then
            If str = "" Then
                str = "[BackOfficeProf] IN " & ChannelProf
            Else str = str & " OR " & "[BackOfficeProf] IN " & ChannelProf
            End If
        End If

        '------------------------------------------------------------------



        If chkOrderProcess.IsChecked Then
            str2 = "[OrderProcess] IN " & CoreBiz
        End If
        If chkCredit.IsChecked Then
            If str2 = "" Then
                str2 = "[Credit] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[Credit] IN " & CoreBiz
            End If
        End If
        If chkRebill.IsChecked Then
            If str2 = "" Then
                str2 = "[Rebill] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[Rebill] IN " & CoreBiz
            End If
        End If
        If chkReturns.IsChecked Then
            If str2 = "" Then
                str2 = "[Returns] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[Returns] IN " & CoreBiz
            End If
        End If
        If chkComplaints.IsChecked Then
            If str2 = "" Then
                str2 = "[Complaints] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[Complaints] IN " & CoreBiz
            End If
        End If
        If chkShipDec.IsChecked Then
            If str2 = "" Then
                str2 = "[ShippingDiscrepancy] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[ShippingDiscrepancy] IN " & CoreBiz
            End If
        End If
        If chkDocReq.IsChecked Then
            If str2 = "" Then
                str2 = "[DocumentRequest] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[DocumentRequest] IN " & CoreBiz
            End If
        End If
        If chkPricProd.IsChecked Then
            If str2 = "" Then
                str2 = "[PricingProduct] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[PricingProduct] IN " & CoreBiz
            End If
        End If
        If chkImp.IsChecked Then
            If str2 = "" Then
                str2 = "[Implementation] IN " & CoreBiz
            Else str2 = str2 & " OR " & "[Implementation] IN " & CoreBiz
            End If
        End If

        If chkBackOffice.IsChecked Then
            str3 = "[BackOffice]=1"
        End If
        If chkCHI.IsChecked Then
            If str3 = "" Then
                str3 = "[CHI]=1"
            Else str3 = str3 & " OR " & "[CHI]=1"
            End If
        End If
        If chkConcierge.IsChecked Then
            If str3 = "" Then
                str3 = "[Concierge]=1"
            Else str3 = str3 & " OR " & "[Concierge]=1"
            End If
        End If
        If chkCommercial.IsChecked Then
            If str3 = "" Then
                str3 = "[Commercial]=1"
            Else str3 = str3 & " OR " & "[Commercial]=1"
            End If
        End If
        If chkDCS.IsChecked Then
            If str3 = "" Then
                str3 = "[DCS]=1"
            Else str3 = str3 & " OR " & "[DCS]=1"
            End If
        End If
        If chkGov.IsChecked Then
            If str3 = "" Then
                str3 = "[GOV]=1"
            Else str3 = str3 & " OR " & "[GOV]=1"
            End If
        End If
        If chkIDNC.IsChecked Then
            If str3 = "" Then
                str3 = "[IDNC]=1"
            Else str3 = str3 & " OR " & "[IDNC]=1"
            End If
        End If
        If chkKaiser.IsChecked Then
            If str3 = "" Then
                str3 = "[Kaiser]=1"
            Else str3 = str3 & " OR " & "[Kaiser]=1"
            End If
        End If
        If chkRouter.IsChecked Then
            If str3 = "" Then
                str3 = "[HSRouter]=1"
            Else str3 = str3 & " OR " & "[HSRouter]=1"
            End If
        End If
        If chkSalesSupport.IsChecked Then
            If str3 = "" Then
                str3 = "[SalesSupport]=1"
            Else str3 = str3 & " OR " & "[SalesSupport]=1"
            End If
        End If
        If chkSpecialty.IsChecked Then
            If str3 = "" Then
                str3 = "[Specialty]=1"
            Else str3 = str3 & " OR " & "[Specialty]=1"
            End If
        End If
        If chkCarrierDispo.IsChecked Then
            If str3 = "" Then
                str3 = "[CarrierDisposition]=1"
            Else str3 = str3 & " OR " & "[CarrierDisposition]=1"
            End If
        End If
        If chkAccountMaintenance.IsChecked Then
            If str3 = "" Then
                str3 = "[AccountMaintenance]=1"
            Else str3 = str3 & " OR " & "[AccountMaintenance]=1"
            End If
        End If
        If chkSPO.IsChecked Then
            If str3 = "" Then
                str3 = "[SpecialProcessOrders]=1"
            Else str3 = str3 & " OR " & "[SpecialProcessOrders]=1"
            End If
        End If
        If chkResearch.IsChecked Then
            If str3 = "" Then
                str3 = "[Research]=1"
            Else str3 = str3 & " OR " & "[Research]=1"
            End If
        End If

        If str = "" Then
            MsgBox("Please Select a Channel", vbInformation)
            Exit Sub
        End If

        If str2 = "" Then
            MsgBox("Please Select a Core Business", vbInformation)
            Exit Sub
        End If
        If str3 = "" Then
            MsgBox("Please Select a Home Skill", vbInformation)
            Exit Sub
        End If

        If str4 = "" Then
            MsgBox("Please Select a Channel Priority", vbInformation)
            Exit Sub
        End If


        If Not IsNothing(temp) Then
            gvSummary.ItemsSource = sm.GetSkillViewList(temp, cmbSegment.SelectedValue, str, str2, str3, str4)
        End If

        gvFiltered.ItemsSource = sm.GetSkillView(cmbDept.Text, cmbSegment.SelectedValue, str, str2, str3, str4).DefaultView

        If gvFiltered.Items.Count = 0 Then
            MsgBox("No Items Matched")
        Else
            MsgBox(gvFiltered.Items.Count & " Items Matched")
        End If

    End Sub

    Private Sub btnExport_Click(sender As Object, e As RoutedEventArgs) Handles btnExport.Click
        Dim extension As String = "xls"
        Dim dialog As New SaveFileDialog() With {
     .DefaultExt = extension,
     .Filter = String.Format("{1} files (.{0})|.{0}|All files (.)|.", extension, "Excel"),
     .FilterIndex = 1
    }

        Try
            If dialog.ShowDialog() <> False Then
                Using stream As Stream = dialog.OpenFile()
                    gvFiltered.Export(stream, New GridViewExportOptions() With {
             .Format = ExportFormat.Html,
             .ShowColumnHeaders = True,
             .ShowColumnFooters = False,
             .ShowGroupFooters = False
            })
                End Using
                MessageBox.Show("Export Success", "Grid", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
                Dim fullPath As String = dialog.FileName
                Dim psi As ProcessStartInfo = New ProcessStartInfo()
                psi.FileName = Path.GetFileName(fullPath)
                psi.WorkingDirectory = Path.GetDirectoryName(fullPath)
                psi.Arguments = "p1=hardCodedv1 p2=v2"
                Process.Start(psi)
            End If
        Catch ex As Exception

        End Try
    End Sub
End Class
