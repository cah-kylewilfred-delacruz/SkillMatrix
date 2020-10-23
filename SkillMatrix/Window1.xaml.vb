Imports System.Data

Public Class Window1
    'Public Sub New()

    '    ' This call is required by the designer.
    '    InitializeComponent()

    '    ' Add any initialization after the InitializeComponent() call.






    '    Dim dt As New DataTable



    '    lstComm.Items.Clear()
    '    lstGov.Items.Clear()
    '    lstDCS.Items.Clear()
    '    lstSpecial.Items.Clear()
    '    lstRouter.Items.Clear()

    '    If dt.Rows.Count > 0 Then
    '        For i = 0 To dt.Rows.Count - 1
    '            Dim x As New System.Windows.Controls.CheckBox
    '            x.Content = dt.Rows(i).Item("SkillName")
    '            x.Name = "ID" & dt.Rows(i).Item("Id")
    '            x.Width = 100
    '            x.Height = 25
    '            x.FontSize = 11
    '            x.HorizontalContentAlignment = Windows.HorizontalAlignment.Left
    '            x.VerticalContentAlignment = Windows.VerticalAlignment.Center
    '            x.VerticalAlignment = Windows.VerticalAlignment.Top
    '            x.HorizontalAlignment = Windows.HorizontalAlignment.Stretch
    '            x.Foreground = Windows.Media.Brushes.Black
    '            x.Margin = New Thickness(0, 0, 0, 0)
    '            Me.RegisterName(x.Name, x)

    '            Select Case dt.Rows(i).Item("SkillHome")
    '                Case "Comercial"
    '                    lstComm.Items.Add(x)
    '                Case "DCS"
    '                    lstDCS.Items.Add(x)
    '                Case "Gov"
    '                    lstGov.Items.Add(x)
    '                Case "Speciality"
    '                    lstSpecial.Items.Add(x)
    '                Case "Router"
    '                    lstRouter.Items.Add(x)
    '            End Select


    '        Next
    '    End If



    '    Dim TempUserInfo As New AgentInfo(



    '    If Not IsNothing(TempUserInfo.SkillPreference) Then
    '        If TempUserInfo.SkillPreference.Count > 0 Then
    '            For Each item In TempUserInfo.SkillPreference
    '                Dim cb As System.Windows.Controls.CheckBox = Me.FindName("ID" & item.Skill)
    '                If Not IsNothing(cb) Then
    '                    cb.IsChecked = True
    '                End If
    '            Next
    '        End If
    '    End If


    'End Sub

    'Private Sub btnSave_Click(sender As Object, e As RoutedEventArgs) Handles btnSave.Click

    '    Dim lstSkillId As New List(Of Integer)


    '    If tab2ChkComercial.IsChecked Then
    '        For Each itm In lstComm.Items

    '            Dim cb As System.Windows.Controls.CheckBox = Me.FindName(itm.Name)
    '            If Not IsNothing(cb) Then
    '                If cb.IsChecked Then
    '                    lstSkillId.Add(Val(itm.Name.ToString.Substring(3)))
    '                End If
    '            End If

    '        Next
    '    End If

    '    'Save(lstSkillId,TempUserInfo.EmpNo)

    'End Sub

    'Private Sub tab2ChkComercial_Checked(sender As Object, e As RoutedEventArgs) Handles tab2ChkComercial.Checked
    '    stpComm.Width = 130
    'End Sub

    'Private Sub tab2ChkComercial_Unchecked(sender As Object, e As RoutedEventArgs) Handles tab2ChkComercial.Unchecked
    '    stpComm.Width = 0
    'End Sub
End Class
'Public Class AgentInfo
'    Implements IDisposable



'    Public Property TeamId As Integer
'    Public Property EmpNo As String
'    Public Property Domestic As Boolean
'    Public Property EmpName As String
'    Public Property Supervisor As String
'    Public Property Manager As String

'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' To detect redundant calls



'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                ' TODO: dispose managed state (managed objects).
'            End If



'            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'            ' TODO: set large fields to null.
'        End If
'        disposedValue = True
'    End Sub



'    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
'    'Protected Overrides Sub Finalize()
'    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub



'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'        Dispose(True)
'        ' TODO: uncomment the following line if Finalize() is overridden above.
'        ' GC.SuppressFinalize(Me)
'    End Sub
'#End Region
'End Class
'Public Class CoreFunctionInfo
'    Implements IDisposable




'    Public Property EmpNo As Integer
'    Public Property OrderProcess As Integer
'    Public Property Credit As Integer
'    Public Property Rebill As Integer
'    Public Property Returns As Integer
'    Public Property Complaints As Integer
'    Public Property ShippingDiscrepancy As Integer
'    Public Property DocumentRequests As Integer
'    Public Property PricingProduct As Integer
'    Public Property Implementation As Integer



'    Public Property ValueLinkTrained As Boolean
'    Public Property Comment As String
'    Public Property ModifiedBy As String
'    Public Property ModifiedDate As DateTime



'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' To detect redundant calls



'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                ' TODO: dispose managed state (managed objects).
'            End If



'            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'            ' TODO: set large fields to null.
'        End If
'        disposedValue = True
'    End Sub



'    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
'    'Protected Overrides Sub Finalize()
'    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub



'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'        Dispose(True)
'        ' TODO: uncomment the following line if Finalize() is overridden above.
'        ' GC.SuppressFinalize(Me)
'    End Sub
'#End Region
'End Class
'Public Class HomeSkillInfo
'    Implements IDisposable




'    Public Property EmpNo As Integer
'    Public Property Commercial As Boolean
'    Public Property DCS As Boolean
'    Public Property GoV As Boolean
'    Public Property Speciality As Boolean
'    Public Property Router As Boolean
'    Public Property ModifiedBy As String
'    Public Property ModifiedDate As DateTime




'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' To detect redundant calls



'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                ' TODO: dispose managed state (managed objects).
'            End If



'            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'            ' TODO: set large fields to null.
'        End If
'        disposedValue = True
'    End Sub



'    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
'    'Protected Overrides Sub Finalize()
'    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub



'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'        Dispose(True)
'        ' TODO: uncomment the following line if Finalize() is overridden above.
'        ' GC.SuppressFinalize(Me)
'    End Sub
'#End Region
'End Class
'Public Class ChannelInfo
'    Implements IDisposable




'    Public Property EmpNo As Integer
'    Public Property Router As String
'    Public Property Phone As String
'    Public Property Email As String
'    Public Property Cases As String
'    Public Property ModifiedBy As String
'    Public Property ModifiedDate As DateTime




'#Region "IDisposable Support"
'    Private disposedValue As Boolean ' To detect redundant calls



'    ' IDisposable
'    Protected Overridable Sub Dispose(disposing As Boolean)
'        If Not disposedValue Then
'            If disposing Then
'                ' TODO: dispose managed state (managed objects).
'            End If



'            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
'            ' TODO: set large fields to null.
'        End If
'        disposedValue = True
'    End Sub



'    ' TODO: override Finalize() only if Dispose(disposing As Boolean) above has code to free unmanaged resources.
'    'Protected Overrides Sub Finalize()
'    '    ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'    '    Dispose(False)
'    '    MyBase.Finalize()
'    'End Sub



'    ' This code added by Visual Basic to correctly implement the disposable pattern.
'    Public Sub Dispose() Implements IDisposable.Dispose
'        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
'        Dispose(True)
'        ' TODO: uncomment the following line if Finalize() is overridden above.
'        ' GC.SuppressFinalize(Me)
'    End Sub
'#End Region
'End Class

