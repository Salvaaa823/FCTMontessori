Imports System.Diagnostics

Public Class Form1
    Inherits Form

    ' Declaración de los controles de la interfaz de usuario
    Private WithEvents appListBox As ListBox
    Private WithEvents timeTextBox As TextBox
    Private WithEvents startButton As Button
    Private WithEvents countdownLabel As Label
    Private WithEvents unlockButton As Button
    Private WithEvents passwordTextBox As TextBox

    Public Sub New()
        ' Inicializar los controles de la interfaz de usuario
        appListBox = New ListBox()
        timeTextBox = New TextBox()
        startButton = New Button()
        countdownLabel = New Label()
        unlockButton = New Button()
        passwordTextBox = New TextBox()

        ' Configuración de los controles
        appListBox.Location = New Point(20, 20)
        appListBox.Size = New Size(150, 120)

        timeTextBox.Location = New Point(190, 20)
        timeTextBox.Size = New Size(100, 20)

        startButton.Location = New Point(190, 50)
        startButton.Size = New Size(100, 30)
        startButton.Text = "Bloquear"

        countdownLabel.Location = New Point(190, 90)
        countdownLabel.Size = New Size(100, 20)
        countdownLabel.TextAlign = ContentAlignment.MiddleCenter

        unlockButton.Location = New Point(190, 120)
        unlockButton.Size = New Size(100, 30)
        unlockButton.Text = "Desbloquear"
        unlockButton.Enabled = False

        passwordTextBox.Location = New Point(190, 160)
        passwordTextBox.Size = New Size(100, 20)
        passwordTextBox.PasswordChar = "*"c

        ' Configuración del formulario principal
        Me.Text = "App Locker"
        Me.Size = New Size(320, 220)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.Controls.Add(appListBox)
        Me.Controls.Add(timeTextBox)
        Me.Controls.Add(startButton)
        Me.Controls.Add(countdownLabel)
        Me.Controls.Add(unlockButton)
        Me.Controls.Add(passwordTextBox)
    End Sub

    ' Resto del código de la clase MainForm ...
    Private lockedProcess As Process
    Private lockTime As Integer
    Private Const adminPassword As String = "1234"
    Private isLocked As Boolean = False

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        PopulateAppList()
    End Sub

    Private Sub PopulateAppList()
        appListBox.Items.Add("Word")
        appListBox.Items.Add("Excel")
        appListBox.Items.Add("PowerPoint")
    End Sub

    Private Sub startButton_Click(sender As Object, e As EventArgs) Handles startButton.Click
        If appListBox.SelectedIndex >= 0 Then
            If Integer.TryParse(timeTextBox.Text, lockTime) Then
                LockApp()
            Else
                MessageBox.Show("Ingrese un valor válido para el tiempo de bloqueo.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Else
            MessageBox.Show("Seleccione una aplicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub LockApp()
        Dim selectedApp As String = appListBox.SelectedItem.ToString()
        Dim processName As String = ""

        Select Case selectedApp
            Case "Word"
                processName = "WINWORD"
            Case "Excel"
                processName = "EXCEL"
            Case "PowerPoint"
                processName = "POWERPNT"
        End Select

        Dim processes As Process() = Process.GetProcessesByName(processName)
        If processes.Length > 0 Then
            lockedProcess = processes(0)
            lockedProcess.WaitForInputIdle()

            Dim hwnd As IntPtr = lockedProcess.MainWindowHandle
            WindowStateUtils.MaximizeWindow(hwnd)

            isLocked = True
            countdownTimer.Start()
            UnlockUI()
        Else
            MessageBox.Show("No se encontró la aplicación seleccionada.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub countdownTimer_Tick(sender As Object, e As EventArgs) Handles countdownTimer.Tick
        lockTime -= 1
        countdownLabel.Text = lockTime.ToString()

        If lockTime <= 0 Then
            countdownTimer.Stop()
            UnlockApp()
        End If
    End Sub

    Private Sub UnlockApp()
        If isLocked Then
            isLocked = False
            countdownLabel.Text = ""

            lockedProcess.CloseMainWindow()
            lockedProcess.WaitForExit()

            lockedProcess.Dispose()
            lockedProcess = Nothing

            LockUI()
        End If
    End Sub

    Private Sub LockUI()
        appListBox.Enabled = True
        timeTextBox.Enabled = True
        startButton.Enabled = True
        unlockButton.Enabled = False
    End Sub

    Private Sub UnlockUI()
        appListBox.Enabled = False
        timeTextBox.Enabled = False
        startButton.Enabled = False
        unlockButton.Enabled = True
    End Sub

    Private Sub unlockButton_Click(sender As Object, e As EventArgs) Handles unlockButton.Click
        Dim password As String = passwordTextBox.Text
        If password = adminPassword Then
            UnlockApp()
        Else
            MessageBox.Show("Contraseña incorrecta.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            passwordTextBox.Text = ""
        End If
    End Sub
End Class

Public Class WindowStateUtils
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function ShowWindowAsync(hWnd As IntPtr, nCmdShow As Integer) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function IsIconic(hWnd As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
    End Function

    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetShellWindow() As IntPtr
    End Function

    Public Shared Sub MaximizeWindow(hWnd As IntPtr)
        ShowWindowAsync(hWnd, 3) ' SW_MAXIMIZE
        SetForegroundWindow(hWnd)
    End Sub

    Public Shared Function IsAppInForeground(hWnd As IntPtr) As Boolean
        Dim foregroundWindow As IntPtr = GetForegroundWindow()
        If hWnd = foregroundWindow Then
            Return True
        End If

        Dim shellWindow As IntPtr = GetShellWindow()
        If foregroundWindow = shellWindow Then
            Dim nextWindow As IntPtr = GetWindow(foregroundWindow, 2) ' GW_HWNDNEXT
            While nextWindow <> IntPtr.Zero
                If IsWindowVisible(nextWindow) Then
                    If nextWindow = hWnd Then
                        Return True
                    Else
                        Return False
                    End If
                End If

                nextWindow = GetWindow(nextWindow, 2) ' GW_HWNDNEXT
            End While
        End If

        Return False
    End Function
End Class
