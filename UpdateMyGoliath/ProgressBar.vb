Public Class ProgressBar

    Private m_PercentComplete
    Private m_CurrentStep
        Private m_ProgressBar
        Private m_Title
        Private m_Text
        Private m_StatusBarText
        Private n
        Private strComputer
        Private intScreenHeight
        Private intScreenWidth
        Private objWMIService
        Private ColItems
        Private objItem

        '---Initialize defaults 
        Private Sub ProgessBar_Initialize()
            m_PercentComplete = 0
            m_CurrentStep = 0
            m_Title = "Progress"
            m_Text = ""
        End Sub


        '----Setting the Title of the Progress Bar
        Public Function SetTitle(pTitle)
            m_Title = pTitle
        End Function

        '----Setting the Text of the Progress Bar  
        Public Function SetText(pText)
            m_Text = pText
        End Function

        '----Setting Percentage Complete of the Progress Bar
        Public Function Update(percentComplete)
            m_PercentComplete = percentComplete
            UpdateProgressBar()
        End Function

        '----Refresh the Progress Bar
        Public Function BarRefresh()
            UpdateProgressBar()
        End Function

        '----Display the Progress Bar
        Public Function Show()
            strComputer = "."
            intScreenHeight = 0
            intScreenWidth = 0

        objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        ColItems = objWMIService.ExecQuery("Select * from Win32_DesktopMonitor",, 48)

        For Each objItem In ColItems
                intScreenHeight = objItem.ScreenHeight
                intScreenWidth = objItem.ScreenWidth
            Next

        m_ProgressBar = CreateObject("InternetExplorer.Application")

        '---in code, the colon acts as a line feed 
        m_ProgressBar.navigate2(("about:blank") : m_ProgressBar.width = 540 : m_ProgressBar.height = 125)
            m_ProgressBar.toolbar = False : m_ProgressBar.menubar = False : m_ProgressBar.statusbar = False
            m_ProgressBar.visible = True
        'm_ProgressBar.top = (intScreenHeight - m_ProgressBar.height) / 2
        'm_ProgressBar.left = (intScreenWidth - m_ProgressBar.width) / 2
        m_ProgressBar.document.write("<body Scroll=no style='margin:0px;padding:0px;'><div style='text-align:center;'><span name='pc' id='pc'>0</span></div>")
        m_ProgressBar.document.write("<div id='statusbar' name='statusbar' style='border:1px solid blue;line-height:10px;height:10px;color:blue;'></div>")
        m_ProgressBar.document.write("<div style='text-align:center'><span id='text' name='text'></span></div>")
    End Function

        '----Close the Progress Bar
        Public Function Close()
            m_ProgressBar.quit
        m_ProgressBar = Nothing
    End Function

        '----Updating the Progress Bar
        Private Function UpdateProgressBar()
            If m_PercentComplete = 0 Then
                m_StatusBarText = ""
            End If

            For n = m_CurrentStep To m_PercentComplete - 1
                m_StatusBarText = m_StatusBarText & "|"
                m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
                m_ProgressBar.Document.title = n & "% Complete : " & m_Title
                m_ProgressBar.Document.GetElementById("pc").InnerHtml = n & "% Complete : " & m_Title
            Threading.Thread.Sleep(10)
        Next
            m_ProgressBar.Document.GetElementById("statusbar").InnerHtml = m_StatusBarText
            m_ProgressBar.Document.title = m_PercentComplete & "% Complete : " & m_Title
            m_ProgressBar.Document.GetElementById("pc").InnerHtml = m_PercentComplete & "% Complete : " & m_Title
            m_ProgressBar.Document.GetElementById("text").InnerHtml = m_Text
            m_CurrentStep = m_PercentComplete
        End Function


    End Class
