Imports System.Data

Public Class frmChangePrimary
    Dim pschk As New DataTable
    Public Sub New(ByVal _temp As DataTable)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        pschk = _temp
        lstPrimarySkill.Items.Clear()

        If pschk.Rows.Count > 0 Then
            For i = 0 To pschk.Rows.Count - 1
                Dim x As New System.Windows.Controls.CheckBox

                x.Content = pschk.Rows(i).Item("SkillName")
                x.Name = "IDPS" & pschk.Rows(i).Item("Id")
                x.Width = 220
                x.Height = 30
                x.FontSize = 11
                x.HorizontalContentAlignment = Windows.HorizontalAlignment.Center
                x.VerticalContentAlignment = Windows.VerticalAlignment.Center
                x.VerticalAlignment = Windows.VerticalAlignment.Center
                x.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
                x.Foreground = Windows.Media.Brushes.White
                x.IsChecked = False
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

        MsgBox("Primary skill has been set.", vbInformation, "Agent Primary Skills")
    End Sub

    Private Sub btnCncl_Click(sender As Object, e As RoutedEventArgs) Handles btnCncl.Click
        Close()
    End Sub
End Class
