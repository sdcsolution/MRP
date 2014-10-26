Imports System.Data.OleDb

Public Class frmConfiguration
    Public DayChoose As String
    Public MonthChoose As Integer
    Private WorkID(100) As Integer

    Private Sub frmConfiguration_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mdlConnectDB.ShowDepartment(lbDepartment)
        mdlConnectDB.ShowGroup(lbGroup)
    End Sub

    Private Sub btnSearchMasterCalendar_Click(sender As Object, e As EventArgs) Handles btnSearchMasterCalendar.Click
        'Generate Day in Year
        'mdlConnectDB.MASTERSCHEDULE_GENERATE(cbCalendarYear.Text)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcJan, 1, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcFeb, 2, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcMarch, 3, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcApril, 4, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcMay, 5, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcJune, 6, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcJuly, 7, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcAugust, 8, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcSeptember, 9, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcOctober, 10, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcNovember, 11, 2014)
        mdlConnectDB.MASTERSCHEDULE_COLOR(mcDecember, 12, 2014)
    End Sub

    Private Sub ClearSelectionCalendar(i As Integer)
        If i <> 1 Then
            mcJan.ClearSelection()
        End If
        If i <> 2 Then
            mcFeb.ClearSelection()
        End If
        If i <> 3 Then
            mcMarch.ClearSelection()
        End If
        If i <> 4 Then
            mcApril.ClearSelection()
        End If
        If i <> 5 Then
            mcMay.ClearSelection()
        End If
        If i <> 6 Then
            mcJune.ClearSelection()
        End If
        If i <> 7 Then
            mcJuly.ClearSelection()
        End If
        If i <> 8 Then
            mcAugust.ClearSelection()
        End If
        If i <> 9 Then
            mcSeptember.ClearSelection()
        End If
        If i <> 10 Then
            mcOctober.ClearSelection()
        End If
        If i <> 11 Then
            mcNovember.ClearSelection()
        End If
        If i <> 12 Then
            mcDecember.ClearSelection()
        End If
    End Sub

    Private Sub mcJan_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcJan.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcJan.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(1)
    End Sub

    Private Sub mcFeb_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcFeb.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcFeb.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(2)
    End Sub

    Private Sub mcMarch_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcMarch.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcMarch.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(3)
    End Sub

    Private Sub mcApril_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcApril.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcApril.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(4)
    End Sub

    Private Sub mcMay_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcMay.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcMay.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(5)
    End Sub

    Private Sub mcJune_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcJune.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcJune.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(6)
    End Sub

    Private Sub mcJuly_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcJuly.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcJuly.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(7)
    End Sub

    Private Sub mcAugust_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcAugust.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcAugust.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(8)
    End Sub

    Private Sub mcSeptember_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcSeptember.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcSeptember.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(9)
    End Sub

    Private Sub mcOctober_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcOctober.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcOctober.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(10)
    End Sub

    Private Sub mcNovember_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcNovember.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcNovember.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(11)
    End Sub

    Private Sub mcDecember_DaySelected(sender As Object, e As DaySelectedEventArgs) Handles mcDecember.DaySelected
        Dim m_daysSelected() As String = e.Days
        Dim m_dates As SelectedDatesCollection = mcDecember.SelectedDates
        lbDaySelected.Text = m_dates(0).ToShortDateString()
        ClearSelectionCalendar(12)
    End Sub

    Private Sub lbDepartment_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbDepartment.SelectedIndexChanged
        lbJobDesc.Items.Clear()
        mdlConnectDB.ShowPosition(lbPosition, lbDepartment.Text)
    End Sub

    Private Sub Add1ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Add1ToolStripMenuItem.Click
        mdlConnectDB.EDIT_MASTERSCHEDULE(lbDaySelected.Text, 0, "22121084")
        btnSearchMasterCalendar.PerformClick()
    End Sub

    Private Sub Add2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Add2ToolStripMenuItem.Click
        mdlConnectDB.EDIT_MASTERSCHEDULE(lbDaySelected.Text, 1, "22121084")
        btnSearchMasterCalendar.PerformClick()
    End Sub

    Private Sub lbPosition_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lbPosition.SelectedIndexChanged
        mdlConnectDB.ShowJobDesc(lbJobDesc, CheckPosID(lbPosition.Text, lbDepartment.Text))
    End Sub

    Private Sub btnAddAC_Click(sender As Object, e As EventArgs) Handles btnAddAC.Click
        AddAccountCode(txtAccountCode.Text, txtDescripAC.Text, WorkID(cbWorkAC.SelectedIndex + 1), "22121084")
        DGVList(dgvAccountCode, "SELECT * FROM V_ACCOUNTCODE")
        MessageBox.Show("Update Completed !!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        btnClearAC.PerformClick()
    End Sub

    Private Sub XtraTabControl1_Click(sender As Object, e As EventArgs) Handles XtraTabControl1.Click
        If XtraTabControl1.SelectedTabPageIndex = 4 Then
            BOMComboListDataArray(cbWorkAC, "MRP_Workcenter_LIST", "WorkCenterName,WorkID", "Show", "True")
            Array.Copy(DataID, WorkID, DataID.Length)
            DGVList(dgvAccountCode, "SELECT * FROM V_ACCOUNTCODE")
        End If
    End Sub

    Private Sub btnClearAC_Click(sender As Object, e As EventArgs) Handles btnClearAC.Click
        txtAccountCode.Text = ""
        cbWorkAC.Text = ""
        txtDescripAC.Text = ""
    End Sub
End Class