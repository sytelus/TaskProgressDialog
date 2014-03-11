Imports System.Collections.Specialized

Public Class SharedTaskProgress
    Private Shared m_SharedTaskProgressDialog As TaskProgressDialog = Nothing
    Private Shared m_NestedLevelCount As Integer = 0
    Private Shared m_ProgressBarPropertiesCallStack As New Stack
    Private Shared m_ProgressRangeAvailableForChild As Double = 0.5

    Public Shared Property ProgressRangeAvailableForChild() As Double
        Get
            Return m_ProgressRangeAvailableForChild
        End Get
        Set(ByVal Value As Double)
            m_ProgressRangeAvailableForChild = Value
            DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).ProgressRangeAvailableForChild = Value
        End Set
    End Property

    Public Shared Function GetUserResponse() As enmTaskProgressUserResponse
        Return m_SharedTaskProgressDialog.GetUserResponse
    End Function



    Public Shared Property IsVisible() As Boolean
        Get
            If m_SharedTaskProgressDialog Is Nothing Then
                Return False
            Else
                Return m_SharedTaskProgressDialog.IsVisible
            End If
        End Get
        Set(ByVal Value As Boolean)
            If m_SharedTaskProgressDialog Is Nothing Then
                'Ignore
            Else
                m_SharedTaskProgressDialog.IsVisible = Value
            End If
        End Set
    End Property

    Public Shared Sub UpdateAsFinished()
        UpdateAsFinished(False)
    End Sub
    Public Shared Sub UpdateAsFinished(ByVal showDialogInFinishedState As Boolean)
        UpdateProgress(CurrentVirtualMaxValue)
        'Check if this function doesn't cause stack underflow
        If m_NestedLevelCount > 0 Then
            m_NestedLevelCount -= 1
            Dim currentProgressBarProperties As ProgressBarProperties = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties)
            currentProgressBarProperties.TotalLengthUsedByChilds += ConvertPhysicalToVirtual((CurrentPhysicalProgressValue - currentProgressBarProperties.LastChildPhysicalStart))
            m_ProgressRangeAvailableForChild = currentProgressBarProperties.ProgressRangeAvailableForChild
            SetDialogText(currentProgressBarProperties.DialogTitle, currentProgressBarProperties.TaskMessage, currentProgressBarProperties.TaskDetailMessage)
            m_ProgressBarPropertiesCallStack.Pop()  'Discard values of child whoes execusion just ended
            If m_NestedLevelCount = 0 Then
                m_SharedTaskProgressDialog.UpdateAsFinished(showDialogInFinishedState)
            Else
                'Do not show dialog as finished
            End If
        Else
            'Ignore this call because there are no more pending calls
        End If
    End Sub
    Public Shared ReadOnly Property CurrentVirtualMaxValue() As Double
        Get
            If m_ProgressBarPropertiesCallStack.Count > 0 Then
                Return (DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).VirtualMaximum)
            Else
                Return Double.MaxValue
            End If
        End Get
    End Property
    Public Shared ReadOnly Property CurrentVirtualMinValue() As Double
        Get
            If m_ProgressBarPropertiesCallStack.Count > 0 Then
                Return (DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).VirtualMinimum)
            Else
                Return Double.MinValue
            End If
        End Get
    End Property
    Public Shared Sub UpdateProgress(ByVal taskDetailMessage As String, ByVal progressValue As Double)
        Dim userResponse As enmTaskProgressUserResponse
        UpdateProgress(taskDetailMessage, progressValue, userResponse)
    End Sub
    Public Shared Sub UpdateProgress(ByVal progressValue As Double)
        Dim userResponse As enmTaskProgressUserResponse
        UpdateProgress(Nothing, progressValue, userResponse)
    End Sub
    Public Shared Sub UpdateProgress(ByVal taskDetailMessage As String)
        UpdateProgress(taskDetailMessage, CurrentVirtualProgressValue)
    End Sub
    Public Shared Sub UpdateProgress(ByVal taskDetailMessage As String, ByVal progressValue As Double, ByRef userResponse As enmTaskProgressUserResponse)
        If m_ProgressBarPropertiesCallStack.Count > 0 Then
            Dim currentProgressBarProperties As ProgressBarProperties = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties)

            Dim progressValueAfterChildUseConsideration As Double = ((CurrentVirtualMinValue + currentProgressBarProperties.TotalLengthUsedByChilds)) + (((CurrentVirtualMaxValue - (CurrentVirtualMinValue + currentProgressBarProperties.TotalLengthUsedByChilds)) / (CurrentVirtualMaxValue - CurrentVirtualMinValue)) * (progressValue - CurrentVirtualMinValue))

            'If (progressValueAfterChildUseConsideration < CurrentVirtualProgressValue) Then Stop

            Dim scaledProgressValue As Double
            If currentProgressBarProperties.VirtualMaximum - currentProgressBarProperties.VirtualMinimum = 0 Then
                scaledProgressValue = currentProgressBarProperties.PhysicalMinimum
            Else
                scaledProgressValue = currentProgressBarProperties.PhysicalMinimum + ((currentProgressBarProperties.PhysicalMaximum - currentProgressBarProperties.PhysicalMinimum) * ((progressValueAfterChildUseConsideration - currentProgressBarProperties.VirtualMinimum) / (currentProgressBarProperties.VirtualMaximum - currentProgressBarProperties.VirtualMinimum)))
            End If

            If Not (m_SharedTaskProgressDialog Is Nothing) Then
                m_SharedTaskProgressDialog.UpdateProgress(taskDetailMessage, scaledProgressValue, userResponse, Double.NaN, Double.NaN)
            End If
        Else
            'Ignore the call because .Show wasn't called  first
        End If
    End Sub
    Public Shared Sub UpdateAsFinishedAndUnload()
        UpdateAsFinished()
        UnloadProgressDialog()
    End Sub
    Public Shared Sub UnloadProgressDialog()
        UnloadProgressDialog(False)
    End Sub
    Public Shared Sub UnloadProgressDialog(ByVal forceUnload As Boolean)
        If (m_NestedLevelCount = 0) Or (forceUnload = True) Then
            m_SharedTaskProgressDialog.UnloadProgressDialog()
            m_SharedTaskProgressDialog = Nothing
        Else
            'Nested level is non zero
            'Restore the visibility
            IsVisible = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).IsDialogVisible
        End If
    End Sub
    Public Shared Function ShowMessageBox(ByVal prompt As String, ByVal buttons As MsgBoxStyle, ByVal title As String) As MsgBoxResult
        Return m_SharedTaskProgressDialog.ShowMessageBox(prompt, buttons, title)
    End Function
    Public Shared ReadOnly Property CurrentPhysicalProgressValue() As Double
        Get
            Return m_SharedTaskProgressDialog.GetCurrentProgressValue
        End Get
    End Property
    Public Shared Function ConvertPhysicalToVirtual(ByVal valueToConvert As Double) As Double
        If m_ProgressBarPropertiesCallStack.Count > 0 Then
            Dim currentProgressBarProperties As ProgressBarProperties = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties)
            If (currentProgressBarProperties.PhysicalMaximum - currentProgressBarProperties.PhysicalMinimum) <> 0 Then
                Return (((valueToConvert - currentProgressBarProperties.PhysicalMinimum) / (currentProgressBarProperties.PhysicalMaximum - currentProgressBarProperties.PhysicalMinimum)) * (currentProgressBarProperties.VirtualMaximum - currentProgressBarProperties.VirtualMinimum)) + currentProgressBarProperties.VirtualMinimum
            Else
                Return 0
            End If
        Else
            'TODO: should return something else
            Return CurrentVirtualMinValue
        End If
    End Function
    Public Shared ReadOnly Property CurrentVirtualProgressValue() As Double
        Get
            Return ConvertPhysicalToVirtual(CurrentPhysicalProgressValue)
        End Get
    End Property
    Public Shared Sub SetDialogText(ByVal dialogTitle As String, ByVal taskMessage As String, ByVal taskDetailMessage As String)
        m_SharedTaskProgressDialog.SetDialogText(dialogTitle, taskMessage, taskDetailMessage)
    End Sub
    Public Shared Sub Show(ByVal taskMessage As String, ByVal showAfterMillisecondsDelay As Integer)
        Show(Nothing, taskMessage, Nothing, False, False, False, False, 0, 1, False, showAfterMillisecondsDelay)
    End Sub
    Public Shared Sub Show(ByVal dialogTitle As String, ByVal taskMessage As String)
        Show(dialogTitle, taskMessage, String.Empty, False, False, 0, 1)
    End Sub
    Public Shared Sub Show(ByVal taskMessage As String, ByVal taskDetailMessage As String, ByVal minProgressValue As Double, ByVal maxProgressValue As Double, ByVal showAfterMillisecondsDelay As Integer)
        Show(Nothing, taskMessage, taskDetailMessage, False, False, False, True, minProgressValue, maxProgressValue, False, showAfterMillisecondsDelay)
    End Sub
    Public Shared Sub Show(ByVal taskMessage As String, ByVal taskDetailMessage As String, ByVal minProgressValue As Double, ByVal maxProgressValue As Double)
        Show(Nothing, taskMessage, taskDetailMessage, False, True, minProgressValue, maxProgressValue)
    End Sub
    Public Shared Sub Show(ByVal dialogTitle As String, ByVal taskMessage As String, ByVal taskDetailMessage As String, ByVal isShowCancelButton As Boolean, ByVal isShowProgressBar As Boolean, ByVal minProgressValue As Double, ByVal maxProgressValue As Double)
        Show(dialogTitle, taskMessage, taskDetailMessage, False, False, isShowCancelButton, isShowProgressBar, minProgressValue, maxProgressValue, False, 0)
    End Sub
    Private Shared m_WindowTitle As String = "Progress"
    Public Shared Property WindowTitle() As String
        Get
            Return m_WindowTitle
        End Get
        Set(ByVal value As String)
            m_WindowTitle = value
        End Set
    End Property

    Public Shared Sub Show(ByVal dialogTitle As String, ByVal taskMessage As String, ByVal taskDetailMessage As String, ByVal isShowPauseButton As Boolean, ByVal isShowSkipNextButton As Boolean, ByVal isShowCancelButton As Boolean, ByVal isShowProgressBar As Boolean, ByVal minProgressValue As Double, ByVal maxProgressValue As Double, ByVal resetPreviousCalls As Boolean, ByVal showAfterMillisecondsDelay As Integer)
        m_NestedLevelCount = m_NestedLevelCount + 1
        Dim newTaskMessage As String = taskMessage
        If newTaskMessage Is Nothing Then
            newTaskMessage = "Working..."
        End If
        Dim newDialogTitle As String = dialogTitle
        If newDialogTitle Is Nothing Then
            newDialogTitle = WindowTitle
        End If

        If (m_SharedTaskProgressDialog Is Nothing) Or (resetPreviousCalls = True) Then
            If Not (m_SharedTaskProgressDialog Is Nothing) Then
                m_SharedTaskProgressDialog.UnloadProgressDialog()
            End If
            m_NestedLevelCount = 1
            m_ProgressBarPropertiesCallStack.Clear()
            m_ProgressBarPropertiesCallStack.Push(New ProgressBarProperties(minProgressValue, maxProgressValue, minProgressValue, maxProgressValue, dialogTitle, taskMessage, taskDetailMessage, True, m_ProgressRangeAvailableForChild, 0, 0))

            m_SharedTaskProgressDialog = New TaskProgressDialog
            m_SharedTaskProgressDialog.Show(newDialogTitle, newTaskMessage, taskDetailMessage, isShowPauseButton, isShowSkipNextButton, isShowCancelButton, isShowProgressBar, minProgressValue, maxProgressValue, 0)
        Else
            'Do not reset dialog, set new limits for progress bar

            DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).LastChildPhysicalStart = CurrentPhysicalProgressValue

            'Set min at where we left off
            Dim newPhysicalMin As Double = CurrentPhysicalProgressValue

            'Max = min + ((last Parent's max-last Parent's min)/range)
            Dim parentPhysicalMin As Double = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).PhysicalMinimum
            Dim parentPhysicalMax As Double = DirectCast(m_ProgressBarPropertiesCallStack.Peek, ProgressBarProperties).PhysicalMaximum
            Dim newPhysicalMax As Double = newPhysicalMin + ((parentPhysicalMax - newPhysicalMin) * ProgressRangeAvailableForChild)
            m_ProgressBarPropertiesCallStack.Push(New ProgressBarProperties(newPhysicalMin, newPhysicalMax, minProgressValue, maxProgressValue, dialogTitle, taskMessage, taskDetailMessage, IsVisible, m_ProgressRangeAvailableForChild, 0, 0))
            IsVisible = True
            SetDialogText(newDialogTitle, taskMessage, taskDetailMessage)
            m_SharedTaskProgressDialog.SetDialogLayout(isShowPauseButton, isShowSkipNextButton, isShowCancelButton, isShowProgressBar)
        End If
    End Sub

    Private Class ProgressBarProperties
        Public PhysicalMinimum As Double
        Public PhysicalMaximum As Double
        Public VirtualMinimum As Double
        Public VirtualMaximum As Double
        Public LastChildPhysicalStart As Double
        Public TotalLengthUsedByChilds As Double
        Public DialogTitle As String
        Public TaskMessage As String
        Public TaskDetailMessage As String
        Public IsDialogVisible As Boolean
        Public ProgressRangeAvailableForChild As Double
        Public Sub New(ByVal physicalMin As Double, ByVal physicalMax As Double, ByVal virtualMin As Double, ByVal virtualMax As Double, ByVal thisDialogTitle As String, ByVal thisTaskMessage As String, ByVal thisTaskDetail As String, ByVal thisIsDialogVisible As Boolean, ByVal thisProgressRangeAvailableForChild As Double, ByVal thisLastChildPhysicalStart As Double, ByVal thisTotalLengthUsedByChilds As Double)
            PhysicalMinimum = physicalMin
            PhysicalMaximum = physicalMax
            VirtualMinimum = virtualMin
            VirtualMaximum = virtualMax
            DialogTitle = thisDialogTitle
            TaskMessage = thisTaskMessage
            TaskDetailMessage = thisTaskDetail
            IsDialogVisible = thisIsDialogVisible
            ProgressRangeAvailableForChild = thisProgressRangeAvailableForChild
            LastChildPhysicalStart = thisLastChildPhysicalStart
            TotalLengthUsedByChilds = thisTotalLengthUsedByChilds
        End Sub
    End Class
End Class
