Option Strict Off
Option Explicit On 

Friend Class frmTaskProgress
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
	Public Sub New()
		MyBase.New()
        'This call is required by the Windows Form Designer.
		InitializeComponent()
	End Sub
	'Form overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Required by the Windows Form Designer
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
    Public WithEvents lblProgressPercentage As System.Windows.Forms.Label
    Public WithEvents lblTimeRemaining As System.Windows.Forms.Label
    Public WithEvents lblTaskDetailMessage As System.Windows.Forms.Label
    Public WithEvents cmdSkipNext As System.Windows.Forms.Button
    Public WithEvents cmdPause As System.Windows.Forms.Button
    Public WithEvents cmdCancel As System.Windows.Forms.Button
    Public WithEvents prbTaskProgress As System.Windows.Forms.ProgressBar
    Public WithEvents lblCurrentValue As System.Windows.Forms.Label
    Public WithEvents lblMaxValue As System.Windows.Forms.Label
    Public WithEvents lblMinValue As System.Windows.Forms.Label
    Public WithEvents lblTimeRemainingCaption As System.Windows.Forms.Label
    Public WithEvents lblTaskMessage As System.Windows.Forms.Label
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Public WithEvents ButtonSeperator As System.Windows.Forms.GroupBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmTaskProgress))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.lblProgressPercentage = New System.Windows.Forms.Label
        Me.lblTimeRemaining = New System.Windows.Forms.Label
        Me.lblTaskDetailMessage = New System.Windows.Forms.Label
        Me.cmdSkipNext = New System.Windows.Forms.Button
        Me.cmdPause = New System.Windows.Forms.Button
        Me.cmdCancel = New System.Windows.Forms.Button
        Me.prbTaskProgress = New System.Windows.Forms.ProgressBar
        Me.ButtonSeperator = New System.Windows.Forms.GroupBox
        Me.lblCurrentValue = New System.Windows.Forms.Label
        Me.lblMaxValue = New System.Windows.Forms.Label
        Me.lblMinValue = New System.Windows.Forms.Label
        Me.lblTimeRemainingCaption = New System.Windows.Forms.Label
        Me.lblTaskMessage = New System.Windows.Forms.Label
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblProgressPercentage
        '
        Me.lblProgressPercentage.AutoSize = True
        Me.lblProgressPercentage.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblProgressPercentage.Location = New System.Drawing.Point(192, 120)
        Me.lblProgressPercentage.Name = "lblProgressPercentage"
        Me.lblProgressPercentage.Size = New System.Drawing.Size(17, 12)
        Me.lblProgressPercentage.TabIndex = 16
        Me.lblProgressPercentage.Text = "0%"
        '
        'lblTimeRemaining
        '
        Me.lblTimeRemaining.Location = New System.Drawing.Point(88, 160)
        Me.lblTimeRemaining.Margin = New System.Windows.Forms.Padding(0, 1, 3, 3)
        Me.lblTimeRemaining.Name = "lblTimeRemaining"
        Me.lblTimeRemaining.Size = New System.Drawing.Size(240, 16)
        Me.lblTimeRemaining.TabIndex = 11
        Me.lblTimeRemaining.Text = "(calculating...)"
        Me.lblTimeRemaining.Visible = False
        '
        'lblTaskDetailMessage
        '
        Me.lblTaskDetailMessage.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblTaskDetailMessage.Font = New System.Drawing.Font("Arial", 7.0!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaskDetailMessage.Location = New System.Drawing.Point(0, 138)
        Me.lblTaskDetailMessage.Margin = New System.Windows.Forms.Padding(3, 3, 3, 1)
        Me.lblTaskDetailMessage.Name = "lblTaskDetailMessage"
        Me.lblTaskDetailMessage.Size = New System.Drawing.Size(336, 20)
        Me.lblTaskDetailMessage.TabIndex = 7
        Me.lblTaskDetailMessage.Text = "Task details..."
        '
        'cmdSkipNext
        '
        Me.cmdSkipNext.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSkipNext.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSkipNext.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSkipNext.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSkipNext.Location = New System.Drawing.Point(79, 184)
        Me.cmdSkipNext.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.cmdSkipNext.Name = "cmdSkipNext"
        Me.cmdSkipNext.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSkipNext.Size = New System.Drawing.Size(82, 25)
        Me.cmdSkipNext.TabIndex = 6
        Me.cmdSkipNext.Text = "S&kip Next"
        '
        'cmdPause
        '
        Me.cmdPause.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPause.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPause.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPause.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPause.Location = New System.Drawing.Point(167, 184)
        Me.cmdPause.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.cmdPause.Name = "cmdPause"
        Me.cmdPause.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPause.Size = New System.Drawing.Size(81, 25)
        Me.cmdPause.TabIndex = 4
        Me.cmdPause.Text = "&Pause"
        '
        'cmdCancel
        '
        Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCancel.Location = New System.Drawing.Point(255, 184)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCancel.Size = New System.Drawing.Size(81, 25)
        Me.cmdCancel.TabIndex = 3
        Me.cmdCancel.Text = "Cancel"
        '
        'prbTaskProgress
        '
        Me.prbTaskProgress.Location = New System.Drawing.Point(0, 92)
        Me.prbTaskProgress.Name = "prbTaskProgress"
        Me.prbTaskProgress.Size = New System.Drawing.Size(334, 22)
        Me.prbTaskProgress.Step = 1
        Me.prbTaskProgress.Style = System.Windows.Forms.ProgressBarStyle.Continuous
        Me.prbTaskProgress.TabIndex = 0
        Me.prbTaskProgress.Text = "[Formatted]"
        Me.prbTaskProgress.Value = 75
        '
        'ButtonSeperator
        '
        Me.ButtonSeperator.BackColor = System.Drawing.SystemColors.Control
        Me.ButtonSeperator.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSeperator.ForeColor = System.Drawing.SystemColors.ControlText
        Me.ButtonSeperator.Location = New System.Drawing.Point(2, 175)
        Me.ButtonSeperator.Margin = New System.Windows.Forms.Padding(3, 1, 3, 3)
        Me.ButtonSeperator.Name = "ButtonSeperator"
        Me.ButtonSeperator.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ButtonSeperator.Size = New System.Drawing.Size(462, 4)
        Me.ButtonSeperator.TabIndex = 2
        Me.ButtonSeperator.TabStop = False
        '
        'lblCurrentValue
        '
        Me.lblCurrentValue.AutoSize = True
        Me.lblCurrentValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblCurrentValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCurrentValue.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCurrentValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblCurrentValue.Location = New System.Drawing.Point(131, 119)
        Me.lblCurrentValue.Name = "lblCurrentValue"
        Me.lblCurrentValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCurrentValue.Size = New System.Drawing.Size(16, 12)
        Me.lblCurrentValue.TabIndex = 10
        Me.lblCurrentValue.Text = "0/0"
        '
        'lblMaxValue
        '
        Me.lblMaxValue.AutoSize = True
        Me.lblMaxValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblMaxValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMaxValue.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMaxValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMaxValue.Location = New System.Drawing.Point(308, 120)
        Me.lblMaxValue.Name = "lblMaxValue"
        Me.lblMaxValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMaxValue.Size = New System.Drawing.Size(31, 12)
        Me.lblMaxValue.TabIndex = 9
        Me.lblMaxValue.Text = "Label1"
        Me.lblMaxValue.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'lblMinValue
        '
        Me.lblMinValue.AutoSize = True
        Me.lblMinValue.BackColor = System.Drawing.SystemColors.Control
        Me.lblMinValue.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblMinValue.Font = New System.Drawing.Font("Arial", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMinValue.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblMinValue.Location = New System.Drawing.Point(0, 119)
        Me.lblMinValue.Name = "lblMinValue"
        Me.lblMinValue.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblMinValue.Size = New System.Drawing.Size(31, 12)
        Me.lblMinValue.TabIndex = 8
        Me.lblMinValue.Text = "Label1"
        '
        'lblTimeRemainingCaption
        '
        Me.lblTimeRemainingCaption.AutoSize = True
        Me.lblTimeRemainingCaption.BackColor = System.Drawing.SystemColors.Control
        Me.lblTimeRemainingCaption.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTimeRemainingCaption.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTimeRemainingCaption.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTimeRemainingCaption.Location = New System.Drawing.Point(0, 160)
        Me.lblTimeRemainingCaption.Margin = New System.Windows.Forms.Padding(3, 3, 1, 0)
        Me.lblTimeRemainingCaption.Name = "lblTimeRemainingCaption"
        Me.lblTimeRemainingCaption.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTimeRemainingCaption.Size = New System.Drawing.Size(87, 14)
        Me.lblTimeRemainingCaption.TabIndex = 5
        Me.lblTimeRemainingCaption.Text = "Time Remaining:"
        Me.lblTimeRemainingCaption.Visible = False
        '
        'lblTaskMessage
        '
        Me.lblTaskMessage.AutoSize = True
        Me.lblTaskMessage.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblTaskMessage.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTaskMessage.Location = New System.Drawing.Point(0, 0)
        Me.lblTaskMessage.Name = "lblTaskMessage"
        Me.lblTaskMessage.Size = New System.Drawing.Size(215, 15)
        Me.lblTaskMessage.TabIndex = 0
        Me.lblTaskMessage.Text = "Your task is being processed now..."
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, 22)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(344, 62)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 17
        Me.PictureBox1.TabStop = False
        '
        'frmTaskProgress
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(338, 215)
        Me.ControlBox = False
        Me.Controls.Add(Me.lblTimeRemaining)
        Me.Controls.Add(Me.lblTimeRemainingCaption)
        Me.Controls.Add(Me.lblProgressPercentage)
        Me.Controls.Add(Me.lblCurrentValue)
        Me.Controls.Add(Me.lblMaxValue)
        Me.Controls.Add(Me.lblMinValue)
        Me.Controls.Add(Me.lblTaskMessage)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.lblTaskDetailMessage)
        Me.Controls.Add(Me.cmdSkipNext)
        Me.Controls.Add(Me.cmdPause)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.ButtonSeperator)
        Me.Controls.Add(Me.prbTaskProgress)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmTaskProgress"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Task Progress"
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region


    '********************************************************************************
    'Author: Shital Shah
    'Date: VB6 Original code : June, 2002
    'Purpose:
    '       A reusable, extensible dialog box to show the task progress with following features:
    '        -Automatically calculates time remaining using complex algorithm
    '        -Allows to show which steps is being performed in current task
    '        -Enables you to provide options for Cancel, Skip and Pause
    '        -DoEvents safe
    '        -Automatically retains it's position on top even if it's non modal
    '        -Shows progress in various formats including %.
    'Usage:
    '       * Call DisplayProgressDialog to initialize and show dialog for the first time
    '       * Call UpdateProgress to update the progress value/message any time, check user response, update time remaining
    '       * Call UnloadProgressDialog to unload the dialog
    '       * Use GetUserResponse to check user response any time
    '       * Use IsVisible to show/hide this dialog
    '       * Use ShowMessageBox to show a message to user.
    'Dependency: None
    '********************************************************************************

    Public IsCancelClicked As Boolean = False
    Public IsPauseClicked As Boolean = False
    Public IsSkipNextClicked As Boolean = False

    Public IsProcessingFinished As Boolean = False
    Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
        IsSkipNextClicked = True
        'Just hide the form. Caller should use UnloadDialog to actuallu unload the form
        If IsProcessingFinished = True Then
            Me.Hide()
        End If
    End Sub
    Private Sub cmdSkipNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSkipNext.Click
        IsSkipNextClicked = True
    End Sub
    Private Sub cmdPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPause.Click
        IsPauseClicked = False
    End Sub
End Class