Imports System.IO
Imports System

Public Class Form1

    Private Property vmmserverinput As String
    Private Property homePage As String = "http://www.google.com"
    Private Property canBusResults As String = "http://www.yahoo.com"
    Private Property HvacPage As String = "http://www.ask.com"
    Private Property driveModesPage As String = "http://www.myharman.com"
    Private Property powerModing As String = "http://www.facebook.com"
    Private Property reportbug As String = "https://elvis.harman.com"
    Private Property customscript As String = "http://ltu.edu"

    Private Sub TBDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TBDToolStripMenuItem.Click
        MsgBox("Coming Soon!" & Chr(13) & "Thank you for your patience.", MsgBoxStyle.Information, "Under Development")
    End Sub


    Private Sub HowToUseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToUseToolStripMenuItem.Click
        WebBrowser1.Navigate(homePage)

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Dim response As MsgBoxResult
        response = MsgBox("Are you sure you want to exit?", MsgBoxStyle.Question + MsgBoxStyle.YesNo, "Confirm")
        If response = MsgBoxResult.Yes Then
            Me.Dispose()
        ElseIf response = MsgBoxResult.No Then
            Exit Sub
        End If
    End Sub

    Private Sub TBDToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles TBDToolStripMenuItem1.Click
        MsgBox("Coming Soon!" & Chr(13) & "Thank you for your patience.", MsgBoxStyle.Information, "Under Development")
    End Sub


    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click

        Dim vmmserver As String = ""
        If RadioButton1.Checked = True Then
            vmmserverinput = "CIPI.exe"
        End If
        If RadioButton2.Checked = True Then
            vmmserverinput = "CUSW.exe"
        End If
        If RadioButton3.Checked = True Then
            vmmserverinput = "520.exe"
        End If
        If RadioButton4.Checked = True Then
            vmmserverinput = "940.exe"
        End If
        If RadioButton5.Checked = True Then
            vmmserverinput = "MP.exe"
        End If
        If vmmserver <> "" Then
            MsgBox("Please choose a VMM server to launch. ")
        End If
        Try
            System.Diagnostics.Process.Start(vmmserverinput)

        Catch a As Exception
            MsgBox("Error: Unable to find " & vmmserverinput, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try

    End Sub


    Private Sub InfoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles InfoToolStripMenuItem.Click
        MsgBox("Kevin Delong - Software Test Engineer (Automation)" & Chr(13) & "Kevin.Delong@harman.com" & Chr(13) & Chr(13) & "Amit Kulkarni - Software Tester (Intern)" & Chr(13) & "Amit.Kulkarni@harman.com", MsgBoxStyle.Information, "Contact Us")
    End Sub

    Private Sub PowerModingToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PowerModingToolStripMenuItem.Click
        MsgBox("Coming Soon!" & Chr(13) & "Thank you for your patience.", MsgBoxStyle.Information, "Under Development")
    End Sub

    Private Sub DriveModesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DriveModesToolStripMenuItem.Click
        MsgBox("Coming Soon!" & Chr(13) & "Thank you for your patience.", MsgBoxStyle.Information, "Under Development")
    End Sub

    Private Sub RUToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RUToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "RU_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub MPToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MPToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "mp_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub PNETgenericToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PNETgenericToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "pnet_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub CUSWgenericToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CUSWgenericToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "cusw_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub MaseratiToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaseratiToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "maserati_j6_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub MaseratiCMCToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaseratiCMCToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "maserati_cmc_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub Fiat520ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Fiat520ToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "fiat520_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & RUscript & " in folder")
        End Try
    End Sub

    Private Sub Fiat940ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles Fiat940ToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "fiat940_smoketest.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try
    End Sub

    Private Sub PNETgenericToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles PNETgenericToolStripMenuItem1.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "pnet_HVAC.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try
    End Sub

    Private Sub PNETRUToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PNETRUToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "pnet_ru_HVAC.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try
    End Sub

    Private Sub CUSWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CUSWToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "cusw_HVAC.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try
    End Sub

    Private Sub MaseratiJ6ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MaseratiJ6ToolStripMenuItem.Click
        Dim RUscript As String = ""
        RUscript = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "maserati_j6_HVAC.py")

        Try
            If RUscript = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                System.Diagnostics.Process.Start(RUscript)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & RUscript, MsgBoxStyle.Critical, "Error: missing " & vmmserverinput & " in folder")
        End Try
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        WebBrowser1.Navigate(homePage)
    End Sub

    Private Sub RequirementsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RequirementsToolStripMenuItem.Click
        WebBrowser1.Navigate(canBusResults)
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim resultpool As String = ""
        resultpool = InputBox("Clear your result pool", "Result File Path Needed", "C:\")

        Try
            If resultpool = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                For Each file As String In My.Computer.FileSystem.GetFiles(resultpool)
                    ListBox1.Items.Remove(file)
                Next
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & resultpool, MsgBoxStyle.Critical, "Error: missing " & resultpool & " in folder")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim resultpool As String = ""
        resultpool = InputBox("Enter path for results", "Result File Path Needed", "C:\")
        Try
            If resultpool = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Information, "Incorrect file name")
            Else
                For Each file As String In My.Computer.FileSystem.GetFiles(resultpool)
                    ListBox1.Items.Add(file)
                Next
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & resultpool, MsgBoxStyle.Critical, "Error: missing " & resultpool & " in folder")
        End Try
    End Sub

    Private Sub ImLostHelpToolStripMenuItem_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim listselected As String = ListBox1.SelectedItem
        Try
            If listselected = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Critical, "File cannot be launched!")
            Else
                System.Diagnostics.Process.Start(listselected)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & listselected, MsgBoxStyle.Critical, "Error: missing " & listselected & " in folder")
        End Try
    End Sub

    Private Sub EnableLoggingpyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles EnableLoggingpyToolStripMenuItem.Click
        Dim script As String = ""
        script = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "enableLogging.py")
        Try
            If script = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Critical, "File cannot be launched!")
            Else
                System.Diagnostics.Process.Start(script)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & script, MsgBoxStyle.Critical, "Error: missing " & script & " in folder")
        End Try
    End Sub

    Private Sub TeraTermCommandsdocxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TeraTermCommandsdocxToolStripMenuItem.Click
        Dim script As String = ""
        script = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "teratermcommands.docx")
        Try
            If script = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Critical, "File cannot be launched!")
            Else
                System.Diagnostics.Process.Start(script)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & script, MsgBoxStyle.Critical, "Error: missing " & script & " in folder")
        End Try
    End Sub

    Private Sub ReportABugToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ReportABugToolStripMenuItem1.Click
        WebBrowser1.Navigate(reportbug)
    End Sub

    Private Sub SoftwareBugToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SoftwareBugToolStripMenuItem.Click
        WebBrowser1.Navigate(reportbug)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        ListBox1.Refresh()
    End Sub

    Private Sub ConfluencePageToolStripMenuItem2_Click(sender As Object, e As EventArgs) Handles ConfluencePageToolStripMenuItem2.Click
        WebBrowser1.Navigate(customscript)
    End Sub

    Private Sub TemplatepyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TemplatepyToolStripMenuItem.Click
        Dim script As String = ""
        script = InputBox("Enter python script filename and press OK to execute script", "Filename needed", "template.py")
        Try
            If script = "" Then
                MsgBox("Unable to find file specfied!", MsgBoxStyle.Critical, "File cannot be launched!")
            Else
                System.Diagnostics.Process.Start(script)
            End If

        Catch a As Exception
            MsgBox("Error: Unable to find " & script, MsgBoxStyle.Critical, "Error: missing " & script & " in folder")
        End Try
    End Sub

    Private Sub WhatIsSmokeTestsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WhatIsSmokeTestsToolStripMenuItem.Click
        MsgBox("Smoke tests will determine the basic functionality of each feature. To determine if something is considered broken, please run a script. After a script finished executing, a excel file will be created with a report of the current broken/limited functionality of features. Use the result pool to add a file path to your result folder. ", MsgBoxStyle.Information, "What are smoke tests?")
    End Sub

    Private Sub WhatIsHVACControlsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WhatIsHVACControlsToolStripMenuItem.Click
        MsgBox("HVAC/Controls scripts will focus on Heating, Ventilation, and Air Conditioning as well as Driver/Passenger, Heated/Vented Seats and steering wheel features. To ensure proper can-request values are observed and head unit updates accordingly.", MsgBoxStyle.Information, "What is HVAC/Controls?")
    End Sub

    Private Sub WhatDoYouNeedToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WhatDoYouNeedToolStripMenuItem.Click
        MsgBox("Contact us by visiting our confluence page below. You can request to submit your own custom script ideas and we will create them for you.  You can also learn how to create your own scripts by using the template.py file.", MsgBoxStyle.Information, "Automation - Increasing our scope")
    End Sub

    Private Sub ImLostHelpToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ImLostHelpToolStripMenuItem1.Click
        MsgBox("1) Add all your scripts/vmm servers to a folder with this application." & Chr(13) &
               "2) You can left-click on any script in the menu bar to run it. " & Chr(13) & "3) Follow the instructions in command prompt during script execution" & Chr(13) &
               "4) Be sure you are connected to the Harman network. The confluence page will appear below for scope of scripts.", MsgBoxStyle.Information, "Automation Help for Vehicle Interface")

    End Sub
End Class
