

'If user pressed any buttons, this will tell which one
Public Enum enmTaskProgressUserResponse
    None = 0
    Cancel = 1
    Pause = 2
    SkipNext = 4
End Enum

Public Class TaskProgressDialog
    Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer
    Private Const HWND_TOP As Short = 0
    Private Const HWND_TOPMOST As Short = -1
    Private Const HWND_NOTOPMOST As Short = -2
    Private Const SWP_NOMOVE As Short = &H2S
    Private Const SWP_NOSIZE As Short = &H1S

    'Keep track of time elapsed
    Private m_dtStartTime As Date
    Private m_ProgressForm As frmTaskProgress
    Private m_DelayedShowTimer As Threading.Timer

    'Progress bar control doesn't allow Max = Min so we do Max = Min + near_zero
    Private Const mdNEAR_ZERO_VALUE As Single = 1.0E-31
    Private Const sUNKNOWN_TIME_REMAINING As String = "(calculating...)" 'Show if we can't decide time remaining
    Friend IsButtonAreaShown As Boolean = True
    Friend IsProgressAreaShown As Boolean = True

    'This is the first function to be called. Pass various parameters to initialize.
    'TaskMessage - A main heading message (Example: Exporting...)
    'TaskDetailMessage - Use it for a step performed (Example Reading source files...)
    'IsProcessingFinished - Indicates whether task is in progress (or about to start) or it is finished. For the first time call set to False. When task is finished and you want to show 100% progress, call this function with this param = True
    Public Sub Show(ByVal dialogTitle As String, ByVal taskMessage As String, ByVal taskDetailMessage As String, ByVal isShowPauseButton As Boolean, ByVal isShowSkipNextButton As Boolean, ByVal isShowCancelButton As Boolean, ByVal isShowProgressBar As Boolean, ByVal minProgressValue As Double, ByVal maxProgressValue As Double, ByVal showAfterMillisecondsDelay As Integer)
        m_ProgressForm = New frmTaskProgress

        SetDialogText(dialogTitle, taskMessage, taskDetailMessage)
        m_ProgressForm.cmdPause.Visible = isShowPauseButton
        m_ProgressForm.cmdSkipNext.Visible = isShowSkipNextButton
        m_ProgressForm.cmdCancel.Visible = CBool(isShowCancelButton)
        m_ProgressForm.prbTaskProgress.Visible = isShowProgressBar
        m_ProgressForm.lblProgressPercentage.Visible = isShowProgressBar
        m_ProgressForm.IsProcessingFinished = False

        SetTimeRemainingText(sUNKNOWN_TIME_REMAINING)

        m_ProgressForm.IsCancelClicked = False
        m_ProgressForm.IsPauseClicked = False
        m_ProgressForm.IsSkipNextClicked = False

        SetDialogLayout(isShowPauseButton, isShowSkipNextButton, isShowCancelButton, isShowProgressBar)

        m_ProgressForm.prbTaskProgress.Minimum = CType(minProgressValue, Integer)
        MaxValue = maxProgressValue
        SetCurrentProgressValue(0)
        Call UpdateMinMaxLabels()
        m_dtStartTime = Now

        If showAfterMillisecondsDelay > 0 Then
            'Do not show now
            m_DelayedShowTimer = New Threading.Timer(AddressOf DelayedShowTimerCallBack, m_DelayedShowTimer, showAfterMillisecondsDelay, Threading.Timeout.Infinite)
        Else
            ShowForm()
        End If
    End Sub

    Private Sub SetTimeRemainingText(ByVal timeRemaingText As String)
        m_ProgressForm.lblTimeRemaining.Text = timeRemaingText
        m_ProgressForm.lblTimeRemaining.Visible = (m_ProgressForm.prbTaskProgress.Visible And (timeRemaingText <> sUNKNOWN_TIME_REMAINING))
        m_ProgressForm.lblTimeRemainingCaption.Visible = (m_ProgressForm.prbTaskProgress.Visible And (timeRemaingText <> sUNKNOWN_TIME_REMAINING))
    End Sub

    Friend Sub SetDialogLayout(ByVal isShowPauseButton As Boolean, ByVal isShowSkipNextButton As Boolean, ByVal isShowCancelButton As Boolean, ByVal isShowProgressBar As Boolean)
        If (IsButtonAreaShown = True) And (isShowCancelButton = False) And (isShowPauseButton = False) And (isShowSkipNextButton = False) Then
            IsButtonAreaShown = False
            m_ProgressForm.ClientSize = New Drawing.Size(m_ProgressForm.ClientSize.Width, m_ProgressForm.ButtonSeperator.Top)
            If (isShowProgressBar = False) And (IsProgressAreaShown = True) Then
                IsProgressAreaShown = False
                m_ProgressForm.ClientSize = New Drawing.Size(m_ProgressForm.ClientSize.Width, m_ProgressForm.prbTaskProgress.Top)
            End If
        End If
    End Sub

    Private Sub DelayedShowTimerCallBack(ByVal state As Object)
        DirectCast(state, Threading.Timer).Dispose()
        If Not (m_DelayedShowTimer Is Nothing) Then
            If DirectCast(state, Threading.Timer) Is m_DelayedShowTimer Then
                m_DelayedShowTimer = Nothing
            End If
        End If

        ShowForm()
    End Sub

    Private Sub ShowForm()
        m_ProgressForm.Show()
        SetWindowPos(m_ProgressForm.Handle.ToInt32, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
        m_ProgressForm.Refresh()
    End Sub

    'Check if any button was pressed. We do not use DoEvents because of re-enterancy problems. This function
    'will call API to check the Windows message queue to see if a button was pressed.
    Public Function GetUserResponse() As enmTaskProgressUserResponse
        Dim eUserResponse As enmTaskProgressUserResponse = enmTaskProgressUserResponse.None

        If m_ProgressForm.IsCancelClicked = True Then
            eUserResponse = eUserResponse Or enmTaskProgressUserResponse.Cancel
        End If
        If m_ProgressForm.IsPauseClicked = True Then
            eUserResponse = eUserResponse Or enmTaskProgressUserResponse.Pause
        End If
        If m_ProgressForm.IsSkipNextClicked = True Then
            eUserResponse = eUserResponse Or enmTaskProgressUserResponse.SkipNext
        End If

        'If user did pressed any button, call DoEvents. This will give user a feedback that
        'we have acklowledged what she was doing or otherwise UI will appear as frozen to her.
        If eUserResponse <> enmTaskProgressUserResponse.None Then
            System.Windows.Forms.Application.DoEvents()
        End If

        Return eUserResponse
    End Function


    'Wrappers for Form.Visible. This is not needed but it's gives bit of control for extensibility
    'Users might need to hide this form for a while
    Public Property IsVisible() As Boolean
        Get
            If m_ProgressForm Is Nothing Then
                Return False
            Else
                Return m_ProgressForm.Visible
            End If
        End Get
        Set(ByVal Value As Boolean)
            If m_ProgressForm Is Nothing Then
                'Ignore
            Else
                m_ProgressForm.Visible = Value
                If Value = True Then
                    ShowForm()
                End If
            End If
        End Set
    End Property

    Public Sub UpdateAsFinished(ByVal showDialogInFinishedState As Boolean)
        If Not (m_DelayedShowTimer Is Nothing) Then
            m_DelayedShowTimer.Dispose()
            m_DelayedShowTimer = Nothing
        End If

        'Hide form because we will re-show it as modal
        If m_ProgressForm.Visible Then
            m_ProgressForm.Hide()
        End If
        If showDialogInFinishedState = True Then
            m_ProgressForm.IsProcessingFinished = True
            m_ProgressForm.cmdCancel.Text = "&Close"
            m_ProgressForm.cmdPause.Visible = False
            m_ProgressForm.cmdSkipNext.Visible = False
            SetTimeRemainingText("0 Seconds")
            m_ProgressForm.ShowDialog()
            m_ProgressForm.Close()
        End If
    End Sub

    'Call this function to update the progress value and/or step performed in task. While doing this
    'the ByRef param UserResponse will return you the value if user pressed any buttons on the form like Cancel or Pause
    Public Sub UpdateProgress(ByVal taskDetailMessage As String, ByVal progressValue As Double, ByRef userResponse As enmTaskProgressUserResponse)
        UpdateProgress(taskDetailMessage, progressValue, userResponse, Double.NaN, Double.NaN)
    End Sub

    Public Sub UpdateProgress(ByVal taskDetailMessage As String, ByVal progressValue As Double, ByRef userResponse As enmTaskProgressUserResponse, ByVal progressMinValue As Double, ByVal progressMaxValue As Double)
        'Update min/max values if passed
        If Double.IsNaN(progressMinValue) = False Then
            m_ProgressForm.prbTaskProgress.Minimum = CType(progressMinValue, Integer)
        End If
        If Double.IsNaN(progressMaxValue) = False Then
            MaxValue = progressMaxValue
        End If

        If taskDetailMessage Is Nothing Then
            'Ignore
        Else
            m_ProgressForm.lblTaskDetailMessage.Text = taskDetailMessage
        End If

        'Update progress value. If it's outside the range then limit to allowed extrems
        If (progressValue <= MaxValue) And (progressValue >= m_ProgressForm.prbTaskProgress.Minimum) Then
            SetCurrentProgressValue(progressValue)
        Else
            If (progressValue > MaxValue) Then
                SetCurrentProgressValue(MaxValue)
            Else
                SetCurrentProgressValue(m_ProgressForm.prbTaskProgress.Minimum)
            End If
        End If

        'Calculate percentage of task completed
        Dim dPercentageProgress As Double
        If MaxValue - m_ProgressForm.prbTaskProgress.Minimum <> 0 Then
            dPercentageProgress = 100 * (progressValue - m_ProgressForm.prbTaskProgress.Minimum) / (MaxValue - m_ProgressForm.prbTaskProgress.Minimum)
            If dPercentageProgress > 100 Then
                dPercentageProgress = 100
            End If
            m_ProgressForm.lblProgressPercentage.Text = "(" & dPercentageProgress.ToString("00.0") & "% complete)"
        Else
            m_ProgressForm.lblProgressPercentage.Text = sUNKNOWN_TIME_REMAINING
        End If

        'Update UI
        Call UpdateMinMaxLabels()

        'Calculate how much time might be required for 100% progress
        Call UpdateRemainingTime()

        m_ProgressForm.ToolTip1.ToolTipTitle = Format(GetCurrentProgressValue)

        'This is needed for some reason to show updated values in controls
        m_ProgressForm.Refresh()

        'See if any button was pressed and let the caller know about it
        userResponse = GetUserResponse()
    End Sub

    'Unload the progress dialog
    Public Sub UnloadProgressDialog()
        If Not (m_DelayedShowTimer Is Nothing) Then
            m_DelayedShowTimer.Dispose()
            m_DelayedShowTimer = Nothing
        End If

        m_ProgressForm.Close()
        m_ProgressForm.Dispose()
        m_ProgressForm = Nothing
    End Sub

    Private m_totalOfCurrentSampleTimeRemainingValues As Double = 0
    Private m_timeRemainingValueCurrentSampleCount As Integer = 0
    'Get the remaining time and decide what to show
    Private Sub UpdateRemainingTime()

        'If max = min then we can't decide
        Static dLastProgressValue As Double
        Static dtLastUpdateTime As Date
        Static dTimeRemainingFromLastUpdate As Double
        Dim dMinTimeRemaining As Double
        Dim dMaxTimeRemaining As Double
        Dim sMinTimeRemaining As String
        Dim sMaxTimeRemaining As String
        Dim sTimeRemaining As String
        If MaxValue - m_ProgressForm.prbTaskProgress.Minimum <> 0 Then
            If Now.Subtract(dtLastUpdateTime).TotalSeconds >= 1.5 Then
                'Canculate time remaining using max, min, current progress and time elapsed so far
                m_totalOfCurrentSampleTimeRemainingValues += GetTimeRemaining(m_dtStartTime, Now, m_ProgressForm.prbTaskProgress.Minimum, GetCurrentProgressValue)
                m_timeRemainingValueCurrentSampleCount += 1

                Dim timeRemainingFromTotalAverage As Double = m_totalOfCurrentSampleTimeRemainingValues / m_timeRemainingValueCurrentSampleCount

                'Save last values to calculate time remaining based on most recent updates

                'Keep track of what was last value we had calulated

                'We sample last values at the interval of 3 seconds to avoide fluctuating recalculated values
                If Now.Subtract(dtLastUpdateTime).TotalSeconds >= 4 Then
                    dTimeRemainingFromLastUpdate = GetTimeRemaining(dtLastUpdateTime, Now, dLastProgressValue, GetCurrentProgressValue)
                    dLastProgressValue = GetCurrentProgressValue()
                    dtLastUpdateTime = Now
                End If

                'Find out max and min value of our estimates
                If timeRemainingFromTotalAverage > dTimeRemainingFromLastUpdate Then
                    dMinTimeRemaining = dTimeRemainingFromLastUpdate
                    dMaxTimeRemaining = timeRemainingFromTotalAverage
                Else
                    dMinTimeRemaining = timeRemainingFromTotalAverage
                    dMaxTimeRemaining = dTimeRemainingFromLastUpdate
                End If

                'Format min and max values (ex. 30 seconds, 4 hours, 2 days etc)
                sMinTimeRemaining = FormateTimeRemaining(dMinTimeRemaining, sUNKNOWN_TIME_REMAINING)
                sMaxTimeRemaining = FormateTimeRemaining(dMaxTimeRemaining, sUNKNOWN_TIME_REMAINING)

                'Build the string which we will show as time remaining
                'If min-max values doesn't differ too much we show "2 minutes to 5 minues" else
                'we just show max value because something like "2 minutes to 5 hours" doesn't make sense
                If (sMinTimeRemaining <> sUNKNOWN_TIME_REMAINING) And (sMaxTimeRemaining <> sUNKNOWN_TIME_REMAINING) Then
                    If (dMaxTimeRemaining / dMinTimeRemaining >= 2) And ((dMaxTimeRemaining / dMinTimeRemaining <= 6)) Then
                        sTimeRemaining = "Approximately between " + sMinTimeRemaining & " to " & sMaxTimeRemaining
                    Else
                        sTimeRemaining = sMaxTimeRemaining
                    End If
                    'If we can't estimate min or max
                ElseIf (sMinTimeRemaining = sUNKNOWN_TIME_REMAINING) And (sMaxTimeRemaining = sUNKNOWN_TIME_REMAINING) Then
                    sTimeRemaining = sUNKNOWN_TIME_REMAINING
                    'See if we have max value
                ElseIf (sMaxTimeRemaining <> sUNKNOWN_TIME_REMAINING) Then
                    sTimeRemaining = "Approximately " + sMaxTimeRemaining
                    'See if we have min value
                ElseIf (sMinTimeRemaining <> sUNKNOWN_TIME_REMAINING) Then
                    sTimeRemaining = "Approximately " + sMinTimeRemaining
                    'If code comes here, we forgot to take care of some possibilities"
                Else
                    sTimeRemaining = "(error)"
                End If
                SetTimeRemainingText(sTimeRemaining)
            Else
                'Do not update if progress reports are less then 1.5 sec apart
            End If
        Else
            SetTimeRemainingText(sUNKNOWN_TIME_REMAINING)
        End If
    End Sub

    'Format min and max values (ex. 30 seconds, 4 hours, 2 days etc)
    Private Function FormateTimeRemaining(ByVal TimeRemaining As Double, ByVal StringForUnknownTimeRemaining As String) As String
        'Format the numbers of seconds in terms of mins, hrs, days etc
        Dim sFormattedTimeRemaining As String
        If CInt(TimeRemaining) <= 0 Then
            sFormattedTimeRemaining = StringForUnknownTimeRemaining
        ElseIf TimeRemaining > (CDbl(25) * 3600) Then
            sFormattedTimeRemaining = CInt(TimeRemaining / (24 * 3600)) & " Days"
        ElseIf TimeRemaining > 3601 Then
            sFormattedTimeRemaining = CInt(TimeRemaining / 3600) & " Hours"
        ElseIf TimeRemaining > 61 Then
            sFormattedTimeRemaining = CInt(TimeRemaining / 60) & " Minutes"
        Else
            sFormattedTimeRemaining = CInt(TimeRemaining) & " Seconds"
        End If
        FormateTimeRemaining = sFormattedTimeRemaining
    End Function

    'Canculate time remaining using max, min, current progress and time elapsed so far
    Private Function GetTimeRemaining(ByVal StartTime As Date, ByVal EndTime As Date, ByVal StartValue As Double, ByVal EndValue As Double) As Double
        'Calculate time spent
        Dim dTimeSpentInSeconds As Double
        If StartTime.ToOADate() <> 0 Then
            dTimeSpentInSeconds = (EndTime.Subtract(StartTime).TotalMilliseconds + 1) / 1000.0 'Add 0.1 millisecond because if it is too fast DateDiff will return 0
        Else
            dTimeSpentInSeconds = 0
        End If

        'Normalize values Min to Max as 0 to 1
        Dim dProgressPerUnit As Double
        dProgressPerUnit = (EndValue - StartValue) / (MaxValue - m_ProgressForm.prbTaskProgress.Minimum)

        'Find out how much time we might spend using simple propotion
        Dim dTimeSpentPerUnit As Double
        If dProgressPerUnit <> 0 Then
            dTimeSpentPerUnit = dTimeSpentInSeconds / dProgressPerUnit
        Else
            dTimeSpentPerUnit = 0
        End If

        'Find the remaining time
        GetTimeRemaining = dTimeSpentPerUnit - dTimeSpentInSeconds
    End Function

    'If user code needs to show any message, it can't use MsgBox because this form will be on the top
    Public Function ShowMessageBox(ByVal prompt As String, ByVal buttons As MsgBoxStyle, ByVal title As String) As MsgBoxResult
        'Make this form on top so if this form is hidden behind other windows, MessageBox still shows up
        ShowMessageBox = MsgBox(prompt, buttons, title)
    End Function

    Public Sub SetDialogText(ByVal dialogTitle As String, ByVal taskMessage As String, ByVal taskDetailMessage As String)
        If Not (dialogTitle Is Nothing) Then
            m_ProgressForm.Text = dialogTitle
        End If
        If Not (taskMessage Is Nothing) Then
            m_ProgressForm.lblTaskMessage.Text = taskMessage
        End If
        If Not (taskDetailMessage Is Nothing) Then
            m_ProgressForm.lblTaskDetailMessage.Text = taskDetailMessage
        End If
    End Sub


    'This property wraps m_ProgressForm.prbTaskProgress.Maximum because m_ProgressForm.prbTaskProgress has limitation that Max can not be same as Min. But this form does support Max=Min. To do this we do Max=Min+near_zero_value while writing and subtract it while reading
    Private Property MaxValue() As Double
        Get
            'If Max is up to twice our offset, it means it's value with out offset so subtract it
            If (m_ProgressForm.prbTaskProgress.Maximum - m_ProgressForm.prbTaskProgress.Minimum) <= (mdNEAR_ZERO_VALUE * 2) Then
                Return m_ProgressForm.prbTaskProgress.Maximum - mdNEAR_ZERO_VALUE
            Else
                Return m_ProgressForm.prbTaskProgress.Maximum
            End If
        End Get
        Set(ByVal Value As Double)
            'If value is same as min value, error would be raised, so add some offset
            If (Value - m_ProgressForm.prbTaskProgress.Minimum) <= mdNEAR_ZERO_VALUE Then
                m_ProgressForm.prbTaskProgress.Maximum = CType(Value + mdNEAR_ZERO_VALUE, Integer)
            Else
                m_ProgressForm.prbTaskProgress.Maximum = CType(Value, Integer)
            End If
        End Set
    End Property

    'Update the labels at the progress bar showing min/max values
    Private Sub UpdateMinMaxLabels()
        m_ProgressForm.lblMaxValue.Text = CStr(MaxValue)
        m_ProgressForm.lblMinValue.Text = CStr(m_ProgressForm.prbTaskProgress.Minimum)
        m_ProgressForm.lblCurrentValue.Text = GetCurrentProgressValue() & "/" & MaxValue
    End Sub

    Private m_progressValue As Double
    Public Sub SetCurrentProgressValue(ByVal progressValue As Double)
        m_progressValue = progressValue
        m_ProgressForm.prbTaskProgress.Value = CType(m_progressValue, Integer)
    End Sub
    Public Function GetCurrentProgressValue() As Double
        GetCurrentProgressValue = m_progressValue
    End Function
End Class

