Imports System.Data.OleDb
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing
Imports PdfSharp.Fonts
Imports System.IO

Public Class Crystal_Reporter
    Dim BCA(2) As Integer
    Dim BCA2 As Integer
    Dim MCA1 As Integer
    Dim MCA(2) As Integer
    Dim BBA1 As Integer
    Dim BBA(2) As Integer
    Dim CSE1 As Integer
    Dim CSE(2) As Integer
    Dim IT1 As Integer
    Dim IT(2) As Integer
    Dim CE1 As Integer
    Dim CE(2) As Integer
    Dim HMCT1 As Integer
    Dim HMCT(2) As Integer
    Dim MBA1 As Integer
    Dim MBA(2) As Integer

    Dim EE(2) As Integer
    Dim noOfAnswers As Integer
    Dim noOfBCAStudents(2) As Integer
    Dim noOfMCAStudents(2) As Integer
    Dim noOfCSEStudents(2) As Integer
    Dim noOfITStudents(2) As Integer
    Dim noOfMBAStudents(2) As Integer
    Dim noOfBHMCTStudents(2) As Integer
    Dim noOfBBAStudents(2) As Integer
    Dim noOfEEStudents(2) As Integer
    Dim noOfECEStudents(2) As Integer
    Dim noOfCEStudents(2) As Integer
    Dim absentStudent(1000) As String
    Dim noOfAbsentStudents As Integer
    Dim noOfPresentStudents As Integer
    ' Dim noOfAbsentStudents As Integer
    Dim totalNoOfStudents As Integer
    Dim filenameForRoutine As String
    Dim filenameForBacklog As String
    Dim filenameForBacklogRoutine As String
    Dim fileNameForAbsentee As String
    Dim fileNameForTopSheet As String
    Dim fileNameForReceipt As String
    Dim routineStatus As Boolean = False
    Dim absenteeStatus As Boolean = False
    Dim backlogStatus As Boolean = False
    Dim backlogRoutineStatus As Boolean = False
    Dim topSheetStatus As Boolean = False
    Dim databaseFileStatus As Boolean = False
    Dim receiptStatus As Boolean = False
    Dim noOfDays As Integer
    Dim mainMenu As String
    Dim databasePath As String
    Dim reportStatus As String
    Dim semester As Integer = 1
    Dim ConnString As String
    Dim isDatabaseConnected As Boolean = False
    Dim filename As String
    Dim k As Integer = 0
    Dim noOfBacklogStudents As Integer
    Public MyConnection As OleDbConnection = New OleDbConnection
    Public dr As OleDbDataReader
    Public dr2 As OleDbDataReader
    Public dr3 As OleDbDataReader
    Dim Examstudent(5000) As Student
    Dim AbStudent(5000) As AbsentStudent
    Dim ExaminationDate(60) As DateInfo
    Dim Student(5000) As BacklogStudent
    Dim halfInfo(4) As halfInfo
    Dim Paper(5000) As Paper
    Dim screen As String = "main"
    Private Sub Create_Routine()
        Dim dbQuery As String
        Dim dbquery2 As String
        Dim dbquery3 As String
        Dim datelist(60) As String ' a temporary storege for dates
        Dim datecount As Integer = 0
        Dim temp As Integer = 0
        Dim repeat As Boolean
        Dim k As Integer
        Dim half As Integer = 1
        Dim j As Integer
        'Code to fetch dates from database tables for entire examination and store in an array
        For i = semester To 8 Step 2
            dbQuery = "SELECT *  FROM Sem" & i
            Dim cmd As OleDbCommand = New OleDbCommand(dbQuery, MyConnection)
            dr = cmd.ExecuteReader
            While dr.Read
                repeat = False
                For temp = 0 To datecount - 1
                    If dr("Exam_Date").ToString.Equals(datelist(temp)) Then repeat = True : Exit For
                Next
                If repeat = False Then
                    datelist(datecount) = dr("Exam_Date").ToString
                    ComboBox1.Items.Add(dr("Exam_Date").ToString)
                    datecount += 1
                End If
            End While
        Next
        noOfDays = datecount - 1 ' total no of days of the examination

        'code to select each date, then add information of all the papers in the selected date
        For i = 0 To noOfDays
            k = 0
            ExaminationDate(i) = New DateInfo
            ExaminationDate(i).DateValue = datelist(i).ToString
            For j = semester To 8 Step 2 ' variable semester represents the selected semester(1 for odd, 2 for even)
                dbQuery = "SELECT * FROM Sem" & j & " WHERE(Exam_Date=" & "'" & datelist(i) & "'" & ")" ' choose a date 
                Dim cmd As OleDbCommand = New OleDbCommand(dbQuery, MyConnection)
                dr = cmd.ExecuteReader
                While dr.Read
                    'the code populates the examination date with all the information of the papers which fall in the specified date
                    ExaminationDate(i).paperCode(k) = dr("Paper_Code").ToString
                    ExaminationDate(i).stream(k) = dr("Stream").ToString
                    ExaminationDate(i).semester(k) = j
                    dbquery2 = "SELECT * FROM HalfRecord WHERE(Semester=" & "'" & j & "'" & ")"
                    Dim cmd2 As OleDbCommand = New OleDbCommand(dbquery2, MyConnection)
                    dr2 = cmd2.ExecuteReader
                    While dr2.Read
                        'code to obtain the information that in which half does the semester fall.
                        ExaminationDate(i).half(k) = dr2("Half").ToString
                    End While
                    dbquery3 = "SELECT * FROM PaperInfo WHERE(Paper_Code=" & "'" & ExaminationDate(i).paperCode(k) & "'" & ")"
                    Dim cmd3 As OleDbCommand = New OleDbCommand(dbquery3, MyConnection)
                    dr3 = cmd3.ExecuteReader
                    'code to obtain the paper name of the current selected paper from table"PaperInfo"
                    While dr3.Read
                        ExaminationDate(i).paperName(k) = dr3("Paper_Name").ToString
                    End While
                    k += 1
                    ExaminationDate(i).paperCode(k) = ""
                End While
            Next
        Next
        rpgt.Text = "Most recently Created Report- Date wise- Half wise routine"
        rpgt.ForeColor = Color.DarkCyan
        If reportStatus = "TopSheet" Or reportStatus = "backlog_routine" Then Exit Sub
        Generate_PDF_Report_for_Routine()
    End Sub
    Public Sub Generate_PDF_Report_for_Routine()
        Dim substring As String
        Dim j As Integer
        Dim document As PdfDocument = New PdfDocument
        Dim top As Integer = 40
        document.Info.Title = "Examination Routine"
        document.Info.Author = "Siliguri Institite of Technology"
        Dim page(30) As PdfPage
        For half = 1 To 2
            If half = 1 Then
                substring = "st"
            Else
                substring = "nd"
            End If
            For i = 0 To noOfDays
                j = 0
                k = 0
                page(i) = document.AddPage
                page(i).Size = PageSize.A4
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page(i))
                Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
                Dim font2 As XFont = New XFont("Segoe UI", 8, XFontStyle.Regular)
                Dim font3 As XFont = New XFont("Segoe UI", 7, XFontStyle.Regular)
                page(i).Orientation = PageOrientation.Portrait
                Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
                Do
                    gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(15, 28 + top))
                    gfx.DrawLine(pen, New XPoint(575, 10 + top), New XPoint(575, 28 + top))
                    gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(575, 10 + top))
                    gfx.DrawLine(pen, New XPoint(15, 28 + top), New XPoint(575, 28 + top))
                    gfx.DrawString("Siliguri Institute of Technology ", font, XBrushes.Black, New XRect(0, -20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopCenter)
                    gfx.DrawString("Examination Routine ", font3, XBrushes.Black, New XRect(0, -5 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopCenter)
                    gfx.DrawString("Date: " & ExaminationDate(i).DateValue, font, XBrushes.Black, New XRect(20, 10 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Half: " & half & substring, font, XBrushes.Black, New XRect(500, 10 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                    If ExaminationDate(i).half(k) = half Then ' if the half of the paper matches with the one being computed, print the paper
                        gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(575, 40 + top))
                        gfx.DrawString("Semester", font, XBrushes.Black, New XRect(20, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString("Stream", font, XBrushes.Black, New XRect(100, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString("Paper Code", font, XBrushes.Black, New XRect(200, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString("Paper Name", font, XBrushes.Black, New XRect(330, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawLine(pen, New XPoint(15, 60 + top), New XPoint(575, 60 + top))
                        gfx.DrawString(ExaminationDate(i).semester(k), font2, XBrushes.Black, New XRect(20, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(ExaminationDate(i).stream(k), font2, XBrushes.Black, New XRect(100, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(ExaminationDate(i).paperCode(k), font2, XBrushes.Black, New XRect(200, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawString(ExaminationDate(i).paperName(k) & "", font2, XBrushes.Black, New XRect(330, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                        gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))
                        j += 1
                        gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(15, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(95, 40 + top), New XPoint(95, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(195, 40 + top), New XPoint(195, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(325, 40 + top), New XPoint(325, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(575, 40 + top), New XPoint(575, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))
                    End If
                    k += 1
                Loop Until ExaminationDate(i).paperCode(k) = ""
            Next
        Next
        ' routineStatus = False
        ' code to check whether the Examination Routine request came directly from the user or indirectly via the backlog routine generation stub.
        '  MessageBox.Show(reportStatus)
        ' If Not reportStatus = "backlog_routine" Or Not reportStatus = "TopSheet" Then ' if the request is direct, then prompt the user to save the file and change the state of routine to true (complete)
        Save_File()
        document.Save(filename)
        filenameForRoutine = filename
        routineStatus = True
        ' End If
    End Sub
    Private Sub Save_File()
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "PDF File |"
        saveFileDialog1.Title = "Save a PDF File"
        saveFileDialog1.ShowDialog()
        filename = saveFileDialog1.FileName() & ".pdf"
    End Sub

    Private Sub dbm_MouseEnter(sender As Object, e As EventArgs) Handles dbm.MouseEnter
        dbm.Image = dbg.Image
    End Sub

    Private Sub dbm_MouseLeave(sender As Object, e As EventArgs) Handles dbm.MouseLeave
        dbm.Image = dbb.Image
    End Sub

    Private Sub prfm_Click(sender As Object, e As EventArgs) Handles prfm.Click
        screen = "preferences"
        Open_Preference_Screen()
        back.Visible = True
        back.Enabled = True
    End Sub
    Private Sub Close_Preference_Screen()
        LabelSemPref.Visible = False
        LabelEven.Visible = False
        LabelOdd.Visible = False
        semEven.Visible = False
        semEven.Enabled = False
        semOdd.Visible = False
        semOdd.Enabled = False
        labelpre.Visible = False
        LabelOutput.Visible = False
        OutputDoc.Visible = False
        OutputDoc.Enabled = False
        OutputPDF.Visible = False
        OutputPDF.Visible = False
        LabelPDF.Visible = False
        LabelDoc.Visible = False
        LabelPageSIze.Visible = False
        ComboBoxSize.Visible = False
        ComboBoxSize.Enabled = False
        OutputBorder.Visible = False
        OutputBorder.Enabled = False
        OutputNoBorder.Visible = False
        OutputNoBorder.Enabled = False
        LabelBordered.Visible = False
        LabelNoBorder.Visible = False
        LabelOutputStyle.Visible = False
        Show_Main_Menu()
    End Sub
    Private Sub Open_Preference_Screen()
        Hide_Main_Menu()
        LabelSemPref.Visible = True
        LabelOutput.Visible = True
        LabelEven.Visible = True
        LabelOdd.Visible = True
        semEven.Visible = True
        semEven.Enabled = True
        semOdd.Visible = True
        semOdd.Enabled = True
        labelpre.Visible = True
        OutputDoc.Visible = True
        OutputDoc.Enabled = True
        OutputPDF.Visible = True
        OutputPDF.Enabled = True
        LabelPDF.Visible = True
        LabelDoc.Visible = True
        LabelPageSIze.Visible = True
        ComboBoxSize.Visible = True
        ComboBoxSize.Enabled = True
        LabelOutputStyle.Visible = True
        OutputBorder.Visible = True
        OutputBorder.Enabled = True
        OutputNoBorder.Visible = True
        OutputNoBorder.Enabled = True
        LabelBordered.Visible = True
        LabelNoBorder.Visible = True
    End Sub

    Private Sub prfm_MouseEnter(sender As Object, e As EventArgs) Handles prfm.MouseEnter
        prfm.Image = prfg.Image

    End Sub

    Private Sub prfm_MouseLeave(sender As Object, e As EventArgs) Handles prfm.MouseLeave
        prfm.Image = prfb.Image
    End Sub

    Private Sub rptm_MouseEnter(sender As Object, e As EventArgs) Handles rptm.MouseEnter
        rptm.Image = rptg.Image
    End Sub

    Private Sub rptm_MouseLeave(sender As Object, e As EventArgs) Handles rptm.MouseLeave
        rptm.Image = rptb.Image
    End Sub



    Private Sub hlpm_MouseEnter(sender As Object, e As EventArgs) Handles hlpm.MouseEnter
        hlpm.Image = hlpg.Image
    End Sub


    Private Sub hlpm_MouseLeave(sender As Object, e As EventArgs) Handles hlpm.MouseLeave
        hlpm.Image = hlpb.Image
    End Sub

    Private Sub Form1_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        End
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case semester
            Case 1
                semOdd.Image = semYes.Image
                semEven.Image = semNo.Image
                semInd.Image = semOddInd.Image
            Case 2
                semOdd.Image = semNo.Image
                semEven.Image = semYes.Image
                semInd.Image = semEvenInd.Image
        End Select
        OutputDoc.Image = semYes.Image
        OutputPDF.Image = semNo.Image

    End Sub

    Private Sub Open_Database_Connection_Screen()
        Hide_Main_Menu()
        screen = "database"
        btnm.Visible = True
        btnm.Enabled = True
        btncm.Enabled = True
        btncm.Visible = True
        back.Visible = True
        back.Enabled = True
        LabelChoose.Visible = True
        LabelConnect.Visible = True
        LabelDbStatus.Visible = True
        LabelPath.Visible = True
        dbmain.Visible = True
    End Sub
    Private Sub Close_Database_Connection_Screen()
        Hide_Main_Menu()
        btnm.Visible = False
        btnm.Enabled = False
        btncm.Enabled = False
        btncm.Visible = False
        back.Visible = False
        LabelChoose.Visible = False
        LabelConnect.Visible = False
        LabelDbStatus.Visible = False
        LabelPath.Visible = False
        dbmain.Visible = False
    End Sub
    Private Sub Hide_Main_Menu()
        dbm.Visible = False
        dbm.Enabled = False
        rptm.Visible = False
        rptm.Enabled = False
        hlpm.Visible = False
        hlpm.Enabled = False
        prfm.Visible = False
        prfm.Enabled = False
        Labeldb.Visible = False
        Labelhlp.Visible = False
        Labelrp.Visible = False
        Labelpr.Visible = False
    End Sub
    Private Sub Show_Main_Menu()
        dbm.Visible = True
        dbm.Enabled = True
        rptm.Visible = True
        rptm.Enabled = True
        hlpm.Visible = True
        hlpm.Enabled = True
        prfm.Visible = True
        prfm.Enabled = True
        Labeldb.Visible = True
        Labelhlp.Visible = True
        Labelrp.Visible = True
        Labelpr.Visible = True
    End Sub

    Private Sub dbm_Click(sender As Object, e As EventArgs) Handles dbm.Click
        Open_Database_Connection_Screen()

    End Sub
    Private Sub Open_Report_Generation_Screen()
        Hide_Main_Menu()
        If isDatabaseConnected = False Then
            screen = "report"
            noDb.Visible = True
            LabelNoDb.Visible = True
            Exit Sub
        End If
        screen = "report"
        Routinem.Visible = True
        Routinem.Enabled = True
        blm.Visible = True
        blm.Enabled = True
        blrtm.Visible = True
        blrtm.Enabled = True
        dwr.Visible = True
        bsl.Visible = True
        dblr.Visible = True
        tpshm.Visible = True
        tpshm.Enabled = True
        LabelTop.Visible = True
        abm.Visible = True
        LabelAbm.Visible = True
        sit.Visible = True
        sit.Enabled = True
        receipt.Visible = True
    End Sub
    Private Sub btnm_MouseEnter(sender As Object, e As EventArgs) Handles btnm.MouseEnter
        btnm.Image = btnbg.Image
    End Sub

    Private Sub btnm_MouseLeave(sender As Object, e As EventArgs) Handles btnm.MouseLeave
        btnm.Image = btnb.Image
    End Sub
    Private Sub btncm_MouseEnter(sender As Object, e As EventArgs) Handles btncm.MouseEnter
        btncm.Image = btncg.Image
    End Sub

    Private Sub btncm_MouseLeave(sender As Object, e As EventArgs) Handles btncm.MouseLeave
        btncm.Image = btncb.Image
    End Sub

    Private Sub back_Click(sender As Object, e As EventArgs) Handles back.Click
        Select Case screen
            Case "report2"
                Close_Report_Preferences()
                Open_Report_Generation_Screen()
            Case "report"
                Close_Report_Generation_Screen()
                Show_Main_Menu()
                back.Visible = False
                back.Enabled = False
            Case "database"
                Close_Database_Connection_Screen()
                Show_Main_Menu()
                back.Visible = False
                back.Enabled = False
            Case "preferences"
                Close_Preference_Screen()
                Show_Main_Menu()
                back.Visible = False
                back.Enabled = False
        End Select
        LabelTopSheet.Visible = False
        noDb.Visible = False
        LabelNoDb.Visible = False


    End Sub
    Private Sub Close_Report_Generation_Screen()
        Routinem.Visible = False
        Routinem.Enabled = False
        blm.Visible = False
        blm.Enabled = False
        blrtm.Visible = False
        blrtm.Enabled = False
        dwr.Visible = False
        bsl.Visible = False
        dblr.Visible = False
        tpshm.Visible = False
        tpshm.Enabled = False
        LabelTop.Visible = False
        abm.Visible = False
        LabelAbm.Visible = False
        sit.Visible = False
        sit.Enabled = False
        receipt.Visible = False
    End Sub


    Private Sub btnm_Click(sender As Object, e As EventArgs) Handles btnm.Click
        OpenFileDialog1.Title = "Choose a MS Access Database File for Connection"
        OpenFileDialog1.ShowDialog()
    End Sub
    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

        Dim strm As System.IO.Stream
        Try
            strm = OpenFileDialog1.OpenFile()
            databasePath = OpenFileDialog1.FileName.ToString()
            LabelPath.Text = OpenFileDialog1.FileName.ToString()
            If databasePath.Substring(Len(databasePath) - 5).Equals("accdb") Then
                LabelPath.ForeColor = Color.DarkCyan
                databaseFileStatus = True

            Else
                LabelPath.ForeColor = Color.Coral
                LabelPath.Text = LabelPath.Text & vbNewLine & "Error. The file is not a MS Access database file. Please Choose a correct database file."
            End If
        Catch ex As Exception
            LabelPath.ForeColor = Color.Coral
            LabelPath.Text = "Could not load Database file. The file seems to be already open in another application. Please close it first"
        End Try

    End Sub
    Private Sub Create_Database_Connection()
        Dim connString As String
        If isDatabaseConnected = False Then
            Try
                connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source =" & databasePath
                MyConnection.ConnectionString = connString
                MyConnection.Open()
                ' DatabaseIcon.Image = DatabaseConnectedIcon.Image
                isDatabaseConnected = True
                LabelDbStatus.Text = "Database Connected"
                LabelDbStatus.ForeColor = Color.DarkCyan
                dbState.Image = dbStateOn.Image
                Timer1.Enabled = True
            Catch ex As Exception

            End Try

        End If
    End Sub

    Private Sub btncm_Click(sender As Object, e As EventArgs) Handles btncm.Click
        If databaseFileStatus = True Then
            Create_Database_Connection()
        End If

    End Sub
    Private Sub Routinem_MouseEnter(sender As Object, e As EventArgs) Handles Routinem.MouseEnter
        Routinem.Image = Routineg.Image

    End Sub

    Private Sub Reportm_MouseLeave(sender As Object, e As EventArgs) Handles Routinem.MouseLeave
        Routinem.Image = Routineb.Image
    End Sub

    Private Sub rptm_Click(sender As Object, e As EventArgs) Handles rptm.Click

        Open_Report_Generation_Screen()

        back.Visible = True
        back.Enabled = True
    End Sub
    Private Sub blm_MouseEnter(sender As Object, e As EventArgs) Handles blm.MouseEnter
        blm.Image = blg.Image
    End Sub

    Private Sub blm_MouseLeave(sender As Object, e As EventArgs) Handles blm.MouseLeave
        blm.Image = blb.Image
    End Sub
    Private Sub blrtm_MouseEnter(sender As Object, e As EventArgs) Handles blrtm.MouseEnter
        blrtm.Image = blrtg.Image
    End Sub

    Private Sub blrtm_MouseLeave(sender As Object, e As EventArgs) Handles blrtm.MouseLeave
        blrtm.Image = blrtb.Image
    End Sub
    Private Sub btnsm_MouseEnter(sender As Object, e As EventArgs)
        btnsm.Image = btnsg.Image
    End Sub


    Private Sub btnrsm_MouseEnter(sender As Object, e As EventArgs) Handles btnrsm.MouseEnter
        btnrsm.Image = btnrsg.Image
    End Sub

    Private Sub btnrsm_MouseLeave(sender As Object, e As EventArgs) Handles btnrsm.MouseLeave
        btnrsm.Image = btnrsb.Image
    End Sub

    Private Sub Open_Report_Preferences()
        screen = "report2"
        Close_Report_Generation_Screen()
        Select Case reportStatus
            Case "routine"
                reportMain.Text = "Generate Date wise - Half wise Routine"
            Case "backlog"
                reportMain.Text = "Generate Backlog Student list"
            Case "backlog_routine"
                reportMain.Text = "Generate Backlog Routine"
            Case "Absentee"
                reportMain.Text = "Generate Absentee List"
            Case "TopSheet"
                reportMain.Text = "Generate Top Sheet"
        End Select

        btnsm.Enabled = True
        btnsm.Visible = True
        btnrsm.Enabled = True
        btnrsm.Visible = True
        LabelGenerate.Visible = True
        LabelReport.Visible = True
        reportMain.Visible = True
        rpgt.Visible = True
    End Sub

    Private Sub Close_Report_Preferences()
        Open_Report_Generation_Screen()
        btnsm.Enabled = False
        btnsm.Visible = False
        btnrsm.Enabled = False
        btnrsm.Visible = False
        LabelGenerate.Visible = False
        LabelReport.Visible = False
        reportMain.Visible = False
        rpgt.Visible = False
        LabelDateTop.Visible = False
        LabelMax.Visible = False
        ComboBox1.Visible = False
        ComboBox1.Enabled = False
        TextBoxNoOfAnswerScriptsInAPacket.Visible = False
        TextBoxNoOfAnswerScriptsInAPacket.Enabled = False
    End Sub

    Private Sub Routinem_Click(sender As Object, e As EventArgs) Handles Routinem.Click
        reportStatus = "routine"

        Open_Report_Preferences()
    End Sub


    Private Sub btnrsm_Click(sender As Object, e As EventArgs) Handles btnrsm.Click
        Select Case reportStatus
            Case "routine"
                If Not routineStatus = True Then
                    Create_Routine()
                End If


            Case "backlog"
                If Not backlogStatus = True Then
                    Create_BacklogReport()
                End If

            Case "backlog_routine"
                If Not backlogRoutineStatus = True Then
                    Create_Backlog_Routine()
                End If
            Case "TopSheet"
                '  If Not TextBoxNoOfAnswerScriptsInAPacket.Equals("") Then
                'If routineStatus = False Then

                '                End If

                Count_No_Of_AbsentStudents()
                '              End If
            Case "Absentee"
                Count_No_Of_Students()
            Case "Receipt"
                Write_Receipt()
        End Select

    End Sub
    Private Sub Create_BacklogReport()
        Dim i As Integer
        Dim k(1000) As Integer
        Dim noOfPapers As Integer
        Dim noOfEntries As Integer
        Dim noOfPages As Integer
        Dim dbquery As String
        'Code to assign data of all backlog students from database table to Student 
        dbquery = "SELECT * FROM BacklogList"
        Dim cmd As OleDbCommand = New OleDbCommand(dbquery, MyConnection)
        dr = cmd.ExecuteReader
        i = 0
        While dr.Read
            Student(i) = New BacklogStudent
            Student(i).StudentName = dr("Student_Name").ToString
            Student(i).StudentRollNo = dr("Roll_No").ToString
            Student(i).BacklogPaperOne = dr("Paper1").ToString
            Student(i).BacklogPaperTwo = dr("Paper2").ToString
            Student(i).BacklogPaperThree = dr("Paper3").ToString
            i += 1
        End While
        noOfBacklogStudents = i - 1 ' total no of backlog students
        'Code to assign data of all the papers from database table to BacklogPaper object
        i = 0
        dbquery = "SELECT * FROM PaperInfo"
        Dim cmd2 As OleDbCommand = New OleDbCommand(dbquery, MyConnection)
        dr2 = cmd2.ExecuteReader
        i = 0
        'code to populate the Paper with all the paper codes and paper names
        While dr2.Read
            Paper(i) = New Paper
            Paper(i).PaperCode = dr2("Paper_Code").ToString()
            Paper(i).PaperName = dr2("Paper_Name").ToString()
            i += 1
        End While
        noOfPapers = i - 1 ' total no of papers in the database
        Dim b As Integer = 0
        'Code to check each paper with the backlog papers. If a match occurs, the student data will be appended to the list of students of the specific paper
        For i = 0 To noOfPapers
            For j = 0 To noOfBacklogStudents
                If Student(j).BacklogPaperOne.Equals(Paper(i).PaperCode) Or _
                 Student(j).BacklogPaperTwo.Equals(Paper(i).PaperCode) Or _
                 Student(j).BacklogPaperThree.Equals(Paper(i).PaperCode) Then
                    Paper(i).StudentName(k(i)) = Student(j).StudentName.ToString
                    Paper(i).StudentRollNo(k(i)) = Student(j).StudentRollNo
                    k(i) += 1
                    noOfEntries += 1
                End If
            Next
        Next
        If reportStatus = "backlog_routine" Then Exit Sub
        ' code to generate a pdf file for the report
        noOfPages = (noOfEntries / 37) + 1
        Top = 0
        Dim counter As Integer
        Dim pageNo As Integer
        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Created with PDFsharp"
        document.Info.Author = "Anish Sharma"
        Dim page(30) As PdfPage
        Dim rec As Integer
        Dim flag As Integer
        Dim ctr As Integer = 0
        For i = 0 To noOfPages - 1
            page(pageNo) = document.AddPage
            page(pageNo).Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page(pageNo))
            Dim font2 As XFont = New XFont("Segoe UI", 8, XFontStyle.Regular)
            Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
            Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
            While rec < noOfPapers
                gfx.DrawString("Backlog Report", font, XBrushes.Black, New XRect(0, 10 + Top, page(pageNo).Width.Point, page(i).Height.Point), XStringFormats.TopCenter)
                gfx.DrawString("Paper Code", font, XBrushes.Black, New XRect(20, 40 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                gfx.DrawString("Paper Name", font, XBrushes.Black, New XRect(100, 40 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                gfx.DrawString("Roll No", font, XBrushes.Black, New XRect(380, 40 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                gfx.DrawString("Student's Name", font, XBrushes.Black, New XRect(440, 40 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                gfx.DrawLine(pen, New XPoint(15, 40 + Top), New XPoint(575, 40 + Top))
                gfx.DrawLine(pen, New XPoint(15, 60 + Top), New XPoint(575, 60 + Top))
                While counter < k(rec)
                    gfx.DrawString(Paper(rec).PaperCode, font2, XBrushes.Black, New XRect(20, (ctr + 3) * 20 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(Paper(rec).PaperName, font2, XBrushes.Black, New XRect(100, (ctr + 3) * 20 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(Paper(rec).StudentRollNo(counter), font2, XBrushes.Black, New XRect(380, (ctr + 3) * 20 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(Paper(rec).StudentName(counter), font2, XBrushes.Black, New XRect(440, (ctr + 3) * 20 + Top, page(pageNo).Width.Point, page(pageNo).Height.Point), XStringFormats.TopLeft)
                    ctr += 1
                    gfx.DrawLine(pen, New XPoint(15, 40 + Top), New XPoint(15, (ctr + 3) * 20 + Top))
                    gfx.DrawLine(pen, New XPoint(95, 40 + Top), New XPoint(95, (ctr + 3) * 20 + Top))
                    gfx.DrawLine(pen, New XPoint(375, 40 + Top), New XPoint(375, (ctr + 3) * 20 + Top))
                    gfx.DrawLine(pen, New XPoint(430, 40 + Top), New XPoint(430, (ctr + 3) * 20 + Top))
                    gfx.DrawLine(pen, New XPoint(575, 40 + Top), New XPoint(575, (ctr + 3) * 20 + Top))
                    gfx.DrawLine(pen, New XPoint(15, (ctr + 3) * 20 + Top), New XPoint(575, (ctr + 3) * 20 + Top))
                    If ctr > (37) Then ' if the number of entries in a page exceed 37, then a new page must be created and the state of the information must be preserved.
                        flag = 1
                        pageNo += 1
                        ctr = 0
                        Exit While
                    End If
                    counter += 1
                End While
                If flag = 1 Then
                    flag = 0
                    counter += 1
                    Exit While
                Else
                    counter = 0
                End If
                rec += 1
            End While
        Next
        rpgt.Text = "Most recently Created Report- Backlog Student list"
        rpgt.ForeColor = Color.DarkCyan
        Save_File()
        document.Save(filename)
        filenameForBacklog = filename
        backlogStatus = True
    End Sub
    Private Sub Count_No_Of_AbsentStudents()

        Dim k As Integer = 0
        Dim dbquery As String
        dbquery = "SELECT * FROM AbsentList"
        Dim cmd As OleDbCommand = New OleDbCommand(dbquery, MyConnection)
        dr = cmd.ExecuteReader
        While dr.Read
            absentStudent(k) = dr("Roll_No").ToString.Trim
            k += 1
        End While
        noOfAbsentStudents = k
        ' MessageBox.Show("No of Absent Students= " & noOfAbsentStudents)
        Count_No_Of_Present_Students()
    End Sub
    Private Sub Count_No_Of_Present_Students()
        Dim k As Integer = 0
        Dim j As Integer
        Dim l As Integer = 0
        '  Dim l As Integer
        Dim absent As Boolean
        For i = semester To 8 Step 2
            Dim dbquery As String
            dbquery = "SELECT * FROM StudentListSem" & i
            Dim cmd As OleDbCommand = New OleDbCommand(dbquery, MyConnection)
            dr = cmd.ExecuteReader
            While dr.Read
                absent = False
                For j = 0 To noOfAbsentStudents - 1
                    If dr("Roll_No").ToString.Trim.Equals(absentStudent(j)) Then
                        'MessageBox.Show(absentStudent(j))
                        absent = True
                        Exit For
                    End If
                Next
                If absent = False Then
                    Examstudent(k) = New Student
                    Examstudent(k).StudentRollNo = dr("Roll_No").ToString.Trim
                    Examstudent(k).StudentStream = dr("Stream").ToString.Trim
                    Examstudent(k).semester = i
                    Label4.Text = Label4.Text & Examstudent(k).StudentRollNo & " " & Examstudent(k).StudentStream & " " & Examstudent(k).semester & vbNewLine
                    k += 1
                Else
                    AbStudent(l) = New AbsentStudent
                    AbStudent(l).StudentRollNo = dr("Roll_No").ToString.Trim
                    AbStudent(l).StudentStream = dr("Stream").ToString.Trim
                    ' AbStudent(l).semester = dr("Semester").ToString.Trim
                    AbStudent(l).semester = i
                    'MessageBox.Show(AbStudent(l).StudentRollNo & AbStudent(l).StudentStream)
                    l += 1

                End If

            End While
        Next
        noOfAbsentStudents = l
        noOfPresentStudents = k

        Dim dbquery2 As String
        dbquery2 = "SELECT * FROM HalfRecord"
        Dim cmd2 As OleDbCommand = New OleDbCommand(dbquery2, MyConnection)
        dr = cmd2.ExecuteReader
        Dim h As Integer
        While dr.Read
            halfInfo(h) = New halfInfo
            halfInfo(h).half = dr("Half").ToString
            halfInfo(h).semester = dr("Semester").ToString
            h += 1
        End While
        For h = 0 To 3
            For j = 0 To noOfAbsentStudents - 1
                If AbStudent(j).semester = halfInfo(h).semester Then
                    AbStudent(j).half = halfInfo(h).half
                End If
            Next
        Next
        Count_Total_No_Of_Students_In_Each_Semester()
    End Sub
    Private Sub Count_No_Of_Students()
        Dim j As Integer
        Dim SelectedDate As String
        Dim index As Integer
        SelectedDate = ComboBox1.SelectedItem.ToString
        ' MessageBox.Show(SelectedDate)
        For i = 0 To noOfDays
            If ExaminationDate(i).DateValue = SelectedDate Then
                index = i
                Exit For
            End If
        Next
        Label4.Text = ""
        '   MessageBox.Show("Total= " & totalNoOfStudents)
        j = 0
        While Not ExaminationDate(index).paperCode(j) = ""
            ' Label4.Text = Label4.Text & ExaminationDate(i).DateValue & " " & ExaminationDate(i).half(j) & vbNewLine
            For a = 0 To noOfPresentStudents - 1
                ' Label4.Text = Label4.Text & "Sem student " & Examstudent(a).semester & "Sem Exam " & ExaminationDate(0).semester(j) & vbNewLine & "Stream Student  " & Examstudent(a).StudentStream.Trim & "Stream Exam  " & ExaminationDate(0).stream(j).Trim

                If Examstudent(a).semester.Equals(ExaminationDate(index).semester(j)) And _
               Examstudent(a).StudentStream.Trim.Equals(ExaminationDate(index).stream(j).Trim) Then
                    ExaminationDate(index).noOfStudents(j) += 1
                End If

            Next
            For b = 0 To noOfAbsentStudents - 1

                If AbStudent(b).semester.Equals(ExaminationDate(index).semester(j)) And _
                AbStudent(b).StudentStream.Trim.Equals(ExaminationDate(index).stream(j).Trim) Then
                    ExaminationDate(index).absentStudent(j) += 1
                End If
            Next
            ' If ExaminationDate(index).half(j) = 1 Then
            ' Label4.Text = Label4.Text & "Stream= " & ExaminationDate(0).stream(j) & "Sem= " & ExaminationDate(0).semester(j) & " Half " & ExaminationDate(0).half(j) & "Total No of Students= " & ExaminationDate(0).noOfStudents(j) & vbNewLine
            ' End If

            j += 1
        End While

        Write_PDF(index)

        'Label4.Text = ""
        'MessageBox.Show("-----" & BCA(0) & " - " & BCA(1))
    End Sub
    Private Sub Write_PDF(ByVal index As String)
        Dim substring As String
        Dim absentflag As Boolean = False
        Dim j As Integer
        Dim document As PdfDocument = New PdfDocument

        Dim top As Integer = 40
        document.Info.Title = "Examination Routine"
        document.Info.Author = "Siliguri Institite of Technology"
        Dim page As PdfPage
        For half = 1 To 2
            If half = 1 Then
                substring = "st"
            Else
                substring = "nd"
            End If

            j = 0
            k = 0
            page = document.AddPage
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
            Dim font2 As XFont = New XFont("Segoe UI", 8, XFontStyle.Regular)
            Dim font3 As XFont = New XFont("Segoe UI", 7, XFontStyle.Regular)
            page.Orientation = PageOrientation.Portrait
            Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
            Do
                absentflag = False
                gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(15, 28 + top))
                gfx.DrawLine(pen, New XPoint(575, 10 + top), New XPoint(575, 28 + top))
                gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(575, 10 + top))
                gfx.DrawLine(pen, New XPoint(15, 28 + top), New XPoint(575, 28 + top))
                gfx.DrawString("Siliguri Institute of Technology ", font, XBrushes.Black, New XRect(0, -20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
                gfx.DrawString("Absent Student List ", font3, XBrushes.Black, New XRect(0, -5 + top, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
                gfx.DrawString("Date: " & ExaminationDate(index).DateValue, font, XBrushes.Black, New XRect(20, 10 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawString("Half: " & half & substring, font, XBrushes.Black, New XRect(500, 10 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                If ExaminationDate(index).half(k) = half Then ' if the half of the paper matches with the one being computed, print the paper
                    gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(575, 40 + top))
                    gfx.DrawString("Semester", font, XBrushes.Black, New XRect(20, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Stream", font, XBrushes.Black, New XRect(80, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Total Students", font, XBrushes.Black, New XRect(125, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Present Students", font, XBrushes.Black, New XRect(210, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Absent Students", font, XBrushes.Black, New XRect(310, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Absent Student Roll No", font, XBrushes.Black, New XRect(410, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawLine(pen, New XPoint(15, 60 + top), New XPoint(575, 60 + top))
                    gfx.DrawString(ExaminationDate(index).semester(k), font2, XBrushes.Black, New XRect(20, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).stream(k), font2, XBrushes.Black, New XRect(80, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).noOfStudents(k) + ExaminationDate(index).absentStudent(k) & "", font2, XBrushes.Black, New XRect(125, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).noOfStudents(k) & "", font2, XBrushes.Black, New XRect(210, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).absentStudent(k) & "", font2, XBrushes.Black, New XRect(310, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    If ExaminationDate(index).absentStudent(k) > 0 Then
                        For b = 0 To noOfAbsentStudents - 1
                            'Label4.Text = Label4.Text & "Absent Student - Stream " & AbStudent(b).StudentStream & " Absent Student Half= " & AbStudent(b).half & " Absent Student Semester= " & AbStudent(b).semester & "Exam Stream= " & ExaminationDate(index).stream(k) & " Exam Half= " & ExaminationDate(index).half(k) & " Exam Semester= " & ExaminationDate(index).semester(k) & vbNewLine
                            If AbStudent(b).StudentStream.Trim.Equals(ExaminationDate(index).stream(k).Trim) And AbStudent(b).semester = ExaminationDate(index).semester(k) Then
                                ' AbStudent(b).half = Equals(ExaminationDate(index).half(k) And 

                                'MessageBox.Show("Absent Student - Stream" & AbStudent(b).StudentStream & "Absent Student Half= " & AbStudent(b).half & "Absent Student Semester= " & AbStudent(b).semester & "Exam Half=" & ExaminationDate(index).half(k) & "Exam Stream=" & ExaminationDate(index).stream(k) & "Exam Semester= " & ExaminationDate(index).semester(k))

                                Label4.Text = Label4.Text & "Absent Student - Stream " & AbStudent(b).StudentStream & " Absent Student Half= " & AbStudent(b).half & " Absent Student Semester= " & AbStudent(b).semester & "Exam Stream= " & ExaminationDate(index).stream(k) & " Exam Half= " & ExaminationDate(index).half(k) & " Exam Semester= " & ExaminationDate(index).semester(k) & vbNewLine
                                gfx.DrawString(AbStudent(b).StudentRollNo & "", font2, XBrushes.Black, New XRect(410, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                                j += 1
                                absentflag = True
                            End If
                        Next
                    End If

                    '  gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))

                    If absentflag = False Then
                        j += 1
                        gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(15, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(72, 40 + top), New XPoint(72, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(120, 40 + top), New XPoint(120, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(205, 40 + top), New XPoint(205, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(305, 40 + top), New XPoint(305, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(405, 40 + top), New XPoint(405, (j + 3) * 20 + top))
                        gfx.DrawLine(pen, New XPoint(575, 40 + top), New XPoint(575, (j + 3) * 20 + top))
                    End If
                    gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))

                End If
                k += 1
            Loop Until ExaminationDate(index).paperCode(k) = ""



        Next
        ' routineStatus = False
        ' code to check whether the Examination Routine request came directly from the user or indirectly via the backlog routine generation stub.
        '  MessageBox.Show(reportStatus)
        ' If Not reportStatus = "backlog_routine" Or Not reportStatus = "TopSheet" Then ' if the request is direct, then prompt the user to save the file and change the state of routine to true (complete)
        Save_File()
        document.Save(filename)
        ' Process.Start(filename)
        fileNameForAbsentee = filename
        absenteeStatus = True
        rpgt.Text = "Most recently Created Report- Absent Student List"
        rpgt.ForeColor = Color.DarkCyan
        ' End If

    End Sub
    Private Sub Write_Receipt()
        Dim index As Integer
        Dim SelectedDate As String
        SelectedDate = ComboBox1.SelectedItem.ToString
        ' MessageBox.Show(SelectedDate)
        For i = 0 To noOfDays
            If ExaminationDate(i).DateValue = SelectedDate Then
                index = i
                Exit For
            End If
        Next
        Dim totalReceived As Integer
        Dim substring As String
        Dim absentflag As Boolean = False
        Dim j As Integer
        Dim document As PdfDocument = New PdfDocument
        Dim base As Integer = 40
        Dim top As Integer = 50
        document.Info.Title = "Examination Routine"
        document.Info.Author = "Siliguri Institite of Technology"
        Dim page As PdfPage
        For half = 1 To 2
            totalReceived = 0
            If half = 1 Then
                substring = "st"
            Else
                substring = "nd"
            End If

            j = 0
            k = 0
            page = document.AddPage
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
            Dim font2 As XFont = New XFont("Segoe UI", 8, XFontStyle.Regular)
            Dim font3 As XFont = New XFont("Segoe UI", 7, XFontStyle.Regular)
            page.Orientation = PageOrientation.Portrait
            Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
            Do
                absentflag = False
                gfx.DrawLine(pen, New XPoint(15, 10 + base), New XPoint(15, 28 + base))
                gfx.DrawLine(pen, New XPoint(575, 10 + base), New XPoint(575, 28 + base))
                gfx.DrawLine(pen, New XPoint(15, 10 + base), New XPoint(575, 10 + base))
                gfx.DrawLine(pen, New XPoint(15, 28 + base), New XPoint(575, 28 + base))
                gfx.DrawString("Siliguri Institute of Technology ", font, XBrushes.Black, New XRect(0, -20 + base, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
                ' gfx.DrawString("Absent Student List ", font3, XBrushes.Black, New XRect(0, -5 + top, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
                gfx.DrawString("Date: " & ExaminationDate(index).DateValue, font, XBrushes.Black, New XRect(20, 10 + base, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                gfx.DrawString("Half: " & half & substring, font, XBrushes.Black, New XRect(500, 10 + base, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                If ExaminationDate(index).half(k) = half Then ' if the half of the paper matches with the one being computed, print the paper
                    gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(575, 40 + top))
                    gfx.DrawString("Semester", font, XBrushes.Black, New XRect(20, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Stream", font, XBrushes.Black, New XRect(80, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Paper Code", font, XBrushes.Black, New XRect(125, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Paper Name", font, XBrushes.Black, New XRect(190, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("No of Copies", font, XBrushes.Black, New XRect(400, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("No of Packets", font, XBrushes.Black, New XRect(480, 40 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawLine(pen, New XPoint(15, 60 + top), New XPoint(575, 60 + top))
                    gfx.DrawString(ExaminationDate(index).semester(k), font2, XBrushes.Black, New XRect(20, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).stream(k), font2, XBrushes.Black, New XRect(80, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).paperCode(k) & "", font2, XBrushes.Black, New XRect(125, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).paperName(k) & "", font2, XBrushes.Black, New XRect(190, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString(ExaminationDate(index).noOfStudents(k) & "", font2, XBrushes.Black, New XRect(400, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    ' gfx.DrawString("" & "", font2, XBrushes.Black, New XRect(410, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    If ExaminationDate(index).noOfStudents(k) < noOfAnswers Then
                        totalReceived += 1
                        gfx.DrawString("1" & "", font2, XBrushes.Black, New XRect(490, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    ElseIf ExaminationDate(index).noOfStudents(k) Mod noOfAnswers = 0 Then
                        totalReceived += ExaminationDate(index).noOfStudents(k) \ noOfAnswers
                        gfx.DrawString(ExaminationDate(index).noOfStudents(k) \ noOfAnswers & "", font2, XBrushes.Black, New XRect(490, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    Else
                        totalReceived += (ExaminationDate(index).noOfStudents(k) \ noOfAnswers) + 1
                        gfx.DrawString((ExaminationDate(index).noOfStudents(k) \ noOfAnswers) + 1 & "", font2, XBrushes.Black, New XRect(490, (j + 3) * 20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    End If

                    ' 
                    '  gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))


                    j += 1
                    gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(15, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(72, 40 + top), New XPoint(72, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(120, 40 + top), New XPoint(120, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(188, 40 + top), New XPoint(188, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(395, 40 + top), New XPoint(395, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(475, 40 + top), New XPoint(475, (j + 3) * 20 + top))
                    gfx.DrawLine(pen, New XPoint(575, 40 + top), New XPoint(575, (j + 3) * 20 + top))

                    gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))

                End If
                k += 1
            Loop Until ExaminationDate(index).paperCode(k) = ""


            gfx.DrawString("Total No of Received Packets  " & totalReceived, font2, XBrushes.Black, New XRect(25, 30 + base, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
        Next

        ' routineStatus = False
        ' code to check whether the Examination Routine request came directly from the user or indirectly via the backlog routine generation stub.
        '  MessageBox.Show(reportStatus)
        ' If Not reportStatus = "backlog_routine" Or Not reportStatus = "TopSheet" Then ' if the request is direct, then prompt the user to save the file and change the state of routine to true (complete)
        'Save_File()
        rpgt.Text = "Most recently Created Report- Answer Paper Receipt"
        rpgt.ForeColor = Color.DarkCyan
        Save_File()
        document.Save(filename)
        fileNameForReceipt = filename
        receiptStatus = True
        ' document.Save("E:\New2.pdf")
        ' Process.Start("E:\New2.pdf")
        ' fileNameForAbsentee = filename
        ' absenteeStatus = True
        ' End If
    End Sub

    Private Sub Count_Total_No_Of_Students_In_Each_Semester()
        Dim dbquery As String
        dbquery = "SELECT * FROM HalfRecord"
        Dim cmd As OleDbCommand = New OleDbCommand(dbquery, MyConnection)
        dr = cmd.ExecuteReader
        Dim h As Integer
        While dr.Read
            halfInfo(h) = New halfInfo
            halfInfo(h).half = dr("Half").ToString
            halfInfo(h).semester = dr("Semester").ToString
            h += 1
        End While
        For h = 0 To 3
            For j = 0 To noOfPresentStudents - 1
                If Examstudent(j).semester = halfInfo(h).semester Then
                    Examstudent(j).half = halfInfo(h).half
                End If
            Next
        Next
        Dim i As Integer

        'code to determine half of the student an assign total no of students in each stream for both halves

        For i = 0 To noOfPresentStudents - 1
            '  If Examstudent(i).present = True Then
            Select Case Examstudent(i).StudentStream

                Case "BCA"
                    noOfBCAStudents(Examstudent(i).half - 1) += 1
                    Label2.Text = Label2.Text & "BCA half " & Examstudent(i).half & " No of Students " & noOfBCAStudents(Examstudent(i).half - 1) & vbNewLine
                Case "MCA"

                    noOfMCAStudents(Examstudent(i).half - 1) += 1

                Case "MBA"

                    noOfMBAStudents(Examstudent(i).half - 1) += 1

                Case "CSE"

                    noOfCSEStudents(Examstudent(i).half - 1) += 1
                    Label3.Text = Label3.Text & "CSE half " & Examstudent(i).half & " No of Students " & noOfCSEStudents(Examstudent(i).half - 1) & vbNewLine

                Case "IT"

                    noOfITStudents(Examstudent(i).half - 1) += 1
                    Label3.Text = Label3.Text & "IT half " & Examstudent(i).half & " No of Students " & noOfITStudents(Examstudent(i).half - 1) & vbNewLine

                Case "BBA"

                    noOfBBAStudents(Examstudent(i).half - 1) += 1

                Case "HMCT"

                    noOfBHMCTStudents(Examstudent(i).half - 1) += 1

                Case "CIVIL"

                    noOfCEStudents(Examstudent(i).half - 1) += 1

                Case "ECE"
                    noOfECEStudents(Examstudent(i).half - 1) += 1
            End Select
            ' End If
        Next
        '  MessageBox.Show(noOfMBAStudents(0) & " - " & noOfMBAStudents(1))
        Print_Top_Sheet()
    End Sub
    Private Sub Print_Top_Sheet()
        Dim document As PdfDocument = New PdfDocument
        Dim i As Integer = 0
        Dim half As Integer = 1
        Dim k As Integer
        Dim counter(12) As Integer
        Dim SelectedDate As String
        Dim index As Integer
        SelectedDate = ComboBox1.SelectedItem.ToString
        ' MessageBox.Show(SelectedDate)
        For i = 0 To noOfDays
            If ExaminationDate(i).DateValue = SelectedDate Then
                index = i
                Exit For
            End If
        Next
        For half = 1 To 2
            i = 0
            For k = 0 To 11
                counter(k) = 0
            Next
            While Not ExaminationDate(index).paperCode(i) = ""
                If ExaminationDate(index).half(i) = half Then
                    Select Case ExaminationDate(index).stream(i)
                        Case "BCA"
                            If counter(0) = 0 Then

                                Create_PDF_for_Top_Sheet("BCA", half, ExaminationDate(index).DateValue, document, noOfBCAStudents(half - 1))
                                counter(0) = 1
                            End If
                        Case "MCA"
                            If counter(1) = 0 Then

                                Create_PDF_for_Top_Sheet("MCA", half, ExaminationDate(index).DateValue, document, noOfMCAStudents(half - 1))
                                counter(1) = 1
                            End If

                        Case "CSE"
                            If counter(2) = 0 Then
                                '   MessageBox.Show("CSE " & noOfCSEStudents(half - 1))
                                '  Label1.Text = Label1.Text & "CSE= " & noOfCSEStudents(half - 1) & vbNewLine
                                Create_PDF_for_Top_Sheet("CSE", half, ExaminationDate(index).DateValue, document, noOfCSEStudents(half - 1))
                                counter(2) = 1
                            End If
                        Case "IT"
                            If counter(3) = 0 Then
                                Label1.Text = Label1.Text & "IT= " & noOfITStudents(half - 1) & vbNewLine
                                Create_PDF_for_Top_Sheet("IT", half, ExaminationDate(index).DateValue, document, noOfITStudents(half - 1))
                                counter(3) = 1
                            End If
                        Case "MBA"
                            If counter(4) = 0 Then
                                Label1.Text = Label1.Text & "MBA= " & noOfMBAStudents(half - 1) & vbNewLine
                                '  Label1.Text = Label1.Text & "BCA= " & noOfBCAStudents(half - 1) & vbNewLine
                                Create_PDF_for_Top_Sheet("MBA", half, ExaminationDate(index).DateValue, document, noOfMBAStudents(half - 1))
                                counter(4) = 1
                            End If
                        Case "BBA"
                            If counter(5) = 0 Then
                                Label1.Text = Label1.Text & "BBA= " & noOfBBAStudents(half - 1) & vbNewLine
                                Create_PDF_for_Top_Sheet("BBA", half, ExaminationDate(index).DateValue, document, noOfBBAStudents(half - 1))
                                counter(5) = 1
                            End If
                        Case "ECE"
                            If counter(6) = 0 Then
                                Label1.Text = Label1.Text & "ECE= " & noOfECEStudents(half - 1) & vbNewLine
                                Create_PDF_for_Top_Sheet("ECE", half, ExaminationDate(index).DateValue, document, noOfECEStudents(half - 1))
                                counter(6) = 1
                            End If
                        Case "EE"
                            If counter(7) = 0 Then
                                Create_PDF_for_Top_Sheet("EE", half, ExaminationDate(index).DateValue, document, noOfEEStudents(half - 1))
                                counter(7) = 1
                            End If
                        Case "CIVIL"
                            If counter(8) = 0 Then
                                Create_PDF_for_Top_Sheet("CIVIL", half, ExaminationDate(index).DateValue, document, noOfCEStudents(half - 1))
                                counter(8) = 1
                            End If
                        Case "HMCT"
                            If counter(8) = 0 Then
                                Create_PDF_for_Top_Sheet("HMCT", half, ExaminationDate(index).DateValue, document, noOfBHMCTStudents(half - 1))
                                counter(8) = 1
                            End If

                    End Select
                End If
                i += 1
            End While
        Next
        Save_File()
        document.Save(filename)
        ' Process.Start(filename)
        fileNameForTopSheet = filename
        topSheetStatus = True
        rpgt.Text = "Most recently Created Report- Top Sheet"
        rpgt.ForeColor = Color.DarkCyan
        ' document.Save("E:\New.pdf")
        ' Process.Start("E:\New.pdf")
    End Sub

    Private Sub Create_PDF_for_Top_Sheet(ByVal stream As String, ByVal half As Integer, ByVal ExamDate As String, ByRef document As PdfDocument, ByRef total As Integer)
        Dim exam As Integer
        noOfAnswers = CType(TextBoxNoOfAnswerScriptsInAPacket.Text, Integer)
        Dim totalNo As Integer
        Dim noOfEntries As Integer
        Dim noOfPages As Integer
        '   MessageBox.Show("Selected from" & stream & ". Total No of Students =" & total)
        Dim substring As String = "st"

        If half = 2 Then
            substring = "nd"
        End If
        Dim top As Integer = 40
        Dim k As Integer
        Dim left As Integer
        document.Info.Title = "Created with PDFsharp"
        document.Info.Author = "Siliguri Institute of Technology"

        noOfPages = (total \ noOfAnswers) + 1
        ' If Not total Mod noOfAnswers = 0 And total > noOfAnswers Then noOfPages += 1
        ' MessageBox.Show("Selected from " & stream & ". Total No of Students =" & total & "No of Pages= " & noOfPages)
        ' MessageBox.Show(noOfPages)
        For temp = 0 To noOfPages - 1
            top = 40
            left = 0
            k = 0
            totalNo = 0
            Dim page As PdfPage = New PdfPage
            page = document.AddPage
            page.Size = PageSize.A4
            Dim gfx As XGraphics = XGraphics.FromPdfPage(page)
            Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
            Dim font2 As XFont = New XFont("Segoe UI", 10, XFontStyle.Regular)
            Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
            gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(15, 28 + top))
            gfx.DrawLine(pen, New XPoint(575, 10 + top), New XPoint(575, 28 + top))
            gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(575, 10 + top))
            gfx.DrawLine(pen, New XPoint(15, 28 + top), New XPoint(575, 28 + top))
            gfx.DrawString("Maulana Abdul Kalam Azad ", font, XBrushes.Black, New XRect(0, -20 + top, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
            gfx.DrawString("University of Technology  ", font2, XBrushes.Black, New XRect(0, -5 + top, page.Width.Point, page.Height.Point), XStringFormats.TopCenter)
            gfx.DrawString("Date: " & ExamDate, font, XBrushes.Black, New XRect(20, 10 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            gfx.DrawString("Half: " & half, font, XBrushes.Black, New XRect(500, 10 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            gfx.DrawString("Stream: " & stream, font, XBrushes.Black, New XRect(20, 30 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            gfx.DrawString("Roll no of students : ", font2, XBrushes.Black, New XRect(20, 50 + top, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            top += 10
            gfx.DrawLine(pen, New XPoint(18, 60 + top), New XPoint(575, 60 + top))
            '  Label1.Text = Label1.Text & vbNewLine & " Date: " & ExamDate & " Stream " & stream & " Half" & half & vbNewLine
            For i = exam To noOfPresentStudents - 1
                If Examstudent(i).StudentStream = stream And Examstudent(i).half = half Then
                    gfx.DrawString(Examstudent(i).StudentRollNo, font2, XBrushes.Black, New XRect(20 + left, (40 + top) + 20, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    Label1.Text = Label1.Text & "   " & Examstudent(i).StudentRollNo
                    gfx.DrawLine(pen, New XPoint(18 + left, 60 + top), New XPoint(18 + left, 80 + top))
                    noOfEntries += 1
                    totalNo += 1
                    k += 1
                    left += 80
                    If noOfEntries = noOfAnswers Then
                        ' gfx.DrawString("noOfEntries= " & noOfEntries, font2, XBrushes.Black, New XRect(20, 400, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                        exam = i + 1
                        noOfEntries = 0
                        Exit For
                    End If
                    If k = 7 Then
                        gfx.DrawLine(pen, New XPoint(575, 60 + top), New XPoint(575, 80 + top))
                        top += 20
                        k = 0
                        left = 0
                        gfx.DrawLine(pen, New XPoint(18, 60 + top), New XPoint(575, 60 + top))
                        ' gfx.DrawString("k=7 ", font2, XBrushes.Black, New XRect(20, 600, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
                    End If
                End If
            Next
            gfx.DrawString("Total no of Answer scripts:  " & totalNo, font, XBrushes.Black, New XRect(400, 70, page.Width.Point, page.Height.Point), XStringFormats.TopLeft)
            gfx.DrawLine(pen, New XPoint(575, 60 + top), New XPoint(575, 80 + top))
            gfx.DrawLine(pen, New XPoint(18 + left, 60 + top), New XPoint(18 + left, 80 + top))
            gfx.DrawLine(pen, New XPoint(18, 80 + top), New XPoint(575, 80 + top))
        Next

    End Sub
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Static a As Integer
        If a Mod 2 = 0 Then
            dbState.Image = dbStateOff.Image
        Else
            dbState.Image = dbStateOn.Image
        End If
        If a = 10 Then
            Timer1.Enabled = False
            dbState.Image = dbStateOn.Image
        End If
        a += 1
    End Sub

    Private Sub blm_Click(sender As Object, e As EventArgs) Handles blm.Click
        reportStatus = "backlog"

        Open_Report_Preferences()
    End Sub

    Private Sub blrtm_Click(sender As Object, e As EventArgs) Handles blrtm.Click
        reportStatus = "backlog_routine"
        Open_Report_Preferences()
    End Sub
    Private Sub Create_Backlog_Routine()
        If routineStatus = False Then ' code to check if datewise routine is prepared. If not, then a request is sent to prepare it before preparing the backlog routine
            Create_Routine()
        End If
        If backlogStatus = False Then ' code to check if Backlog student list is prepared. If not, then a request is sent to prepare it before preparing the backlog routine
            Create_BacklogReport()
        End If
        Dim ct As Integer
        For i = 0 To noOfDays
            ct = 0
            ExaminationDate(i).noOfBacklogStudents(ct) = 0
            While Not ExaminationDate(i).paperCode(ct) = ""
                For ctr = 0 To noOfBacklogStudents
                    If ExaminationDate(i).paperCode(ct).Equals(Student(ctr).BacklogPaperOne) Or _
                    ExaminationDate(i).paperCode(ct).Equals(Student(ctr).BacklogPaperTwo) Or _
                    ExaminationDate(i).paperCode(ct).Equals(Student(ctr).BacklogPaperThree) Then
                        ExaminationDate(i).noOfBacklogStudents(ct) += 1
                    End If
                Next
                ct += 1
            End While
        Next
        rpgt.Text = "Most recently Created Report- Backlog Routine"
        rpgt.ForeColor = Color.DarkCyan
        Create_PDF_Backlog_Routine()
        backlogRoutineStatus = True
    End Sub
    Private Sub Create_PDF_Backlog_Routine()
        Dim top As Integer = 40
        Dim temp As Integer
        Dim tempPaper(500) As String
        Dim flag As Integer = 0
        Dim j As Integer
        Dim document As PdfDocument = New PdfDocument
        document.Info.Title = "Created with PDFsharp"
        document.Info.Author = "Anish Sharma"
        Dim page(30) As PdfPage
        Dim substring As String
        For half = 1 To 2
            If half = 1 Then
                substring = "st"
            Else
                substring = "nd"
            End If
            For i = 0 To noOfDays
                j = 0
                k = 0
                page(i) = document.AddPage
                page(i).Size = PageSize.A4
                Dim gfx As XGraphics = XGraphics.FromPdfPage(page(i))
                Dim font As XFont = New XFont("Segoe UI", 12, XFontStyle.Regular)
                Dim font2 As XFont = New XFont("Segoe UI", 8, XFontStyle.Regular)
                Dim font3 As XFont = New XFont("Segoe UI", 7, XFontStyle.Regular)
                page(i).Orientation = PageOrientation.Portrait
                Dim pen As XPen = New XPen(XColor.FromArgb(0, 0, 0))
                Do
                    gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(15, 28 + top))
                    gfx.DrawLine(pen, New XPoint(575, 10 + top), New XPoint(575, 28 + top))
                    gfx.DrawLine(pen, New XPoint(15, 10 + top), New XPoint(575, 10 + top))
                    gfx.DrawLine(pen, New XPoint(15, 28 + top), New XPoint(575, 28 + top))
                    gfx.DrawString("Backlog Routine ", font, XBrushes.Black, New XRect(0, -20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopCenter)
                    gfx.DrawString("Date: " & ExaminationDate(i).DateValue, font, XBrushes.Black, New XRect(20, 10 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                    gfx.DrawString("Half: " & half & substring, font, XBrushes.Black, New XRect(500, 10 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                    flag = 0
                    For temp = 0 To k - 1
                        If ExaminationDate(i).paperName(temp).Equals(ExaminationDate(i).paperName(k)) Then ' code to deny repetetion 
                            flag = 1
                            Exit For
                        End If
                    Next
                    If ExaminationDate(i).noOfBacklogStudents(k) >= 1 And flag = 0 Then
                        If ExaminationDate(i).half(k) = half Then
                            gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(575, 40 + top))
                            gfx.DrawString("Semester", font, XBrushes.Black, New XRect(20, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString("Stream", font, XBrushes.Black, New XRect(80, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString("Paper Code", font, XBrushes.Black, New XRect(125, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString("Paper Name", font, XBrushes.Black, New XRect(200, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString("Total", font, XBrushes.Black, New XRect(530, 40 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawLine(pen, New XPoint(15, 60 + top), New XPoint(575, 60 + top))
                            gfx.DrawString(ExaminationDate(i).semester(k), font2, XBrushes.Black, New XRect(20, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(ExaminationDate(i).stream(k), font2, XBrushes.Black, New XRect(80, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(ExaminationDate(i).paperCode(k), font2, XBrushes.Black, New XRect(125, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(ExaminationDate(i).paperName(k) & "", font2, XBrushes.Black, New XRect(200, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawString(ExaminationDate(i).noOfBacklogStudents(k) & "", font2, XBrushes.Black, New XRect(530, (j + 3) * 20 + top, page(i).Width.Point, page(i).Height.Point), XStringFormats.TopLeft)
                            gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))
                            j += 1
                            gfx.DrawLine(pen, New XPoint(15, 40 + top), New XPoint(15, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(75, 40 + top), New XPoint(75, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(120, 40 + top), New XPoint(120, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(195, 40 + top), New XPoint(195, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(575, 40 + top), New XPoint(575, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(525, 40 + top), New XPoint(525, (j + 3) * 20 + top))
                            gfx.DrawLine(pen, New XPoint(15, (j + 3) * 20 + top), New XPoint(575, (j + 3) * 20 + top))
                        End If
                    End If
                    k += 1
                Loop Until ExaminationDate(i).paperCode(k) = ""
            Next
        Next
        Save_File()
        document.Save(filename)
        filenameForBacklogRoutine = filename
    End Sub
    Private Sub btnsm_MouseEnter1(sender As Object, e As EventArgs) Handles btnsm.MouseEnter
        btnsm.Image = btnsg.Image
    End Sub
    Private Sub btnsm_MouseLeave1(sender As Object, e As EventArgs) Handles btnsm.MouseLeave
        btnsm.Image = btnsb.Image
    End Sub
    Private Sub btnsm_Click(sender As Object, e As EventArgs) Handles btnsm.Click
        Select Case reportStatus
            Case "routine"
                If routineStatus = True Then Process.Start(filenameForRoutine)
            Case "backlog"
                If backlogStatus = True Then Process.Start(filenameForBacklog)
            Case "backlog_routine"
                If backlogRoutineStatus = True Then Process.Start(filenameForBacklogRoutine)
            Case "TopSheet"
                If topSheetStatus = True Then Process.Start(fileNameForTopSheet)
            Case "Absentee"
                If absenteeStatus = True Then Process.Start(fileNameForAbsentee)
            Case "Receipt"
                If receiptStatus = True Then Process.Start(fileNameForReceipt)


        End Select

    End Sub
    Private Sub semOdd_Click(sender As Object, e As EventArgs) Handles semOdd.Click
        If semOdd.Image.Equals(semYes.Image) Then
            semOdd.Image = semNo.Image
            semEven.Image = semYes.Image
            semInd.Image = semEvenInd.Image
            semester = 2
        Else
            semOdd.Image = semYes.Image
            semEven.Image = semNo.Image
            semInd.Image = semOddInd.Image
            semester = 1
        End If
    End Sub

    Private Sub semEven_Click(sender As Object, e As EventArgs) Handles semEven.Click
        If semOdd.Image.Equals(semYes.Image) Then
            semOdd.Image = semNo.Image
            semEven.Image = semYes.Image
            semester = 2
            semInd.Image = semEvenInd.Image
        Else
            semOdd.Image = semYes.Image
            semEven.Image = semNo.Image
            semester = 1
            semInd.Image = semOddInd.Image

        End If
    End Sub

    Private Sub OutputPDF_Click(sender As Object, e As EventArgs) Handles OutputPDF.Click
        If OutputPDF.Image.Equals(semYes.Image) Then
            OutputPDF.Image = semNo.Image
            OutputDoc.Image = semYes.Image
            ' semInd.Image = semEvenInd.Image
        Else
            OutputPDF.Image = semYes.Image
            OutputDoc.Image = semNo.Image
            'semInd.Image = semOddInd.Image
        End If
    End Sub

    Private Sub OutputDoc_Click(sender As Object, e As EventArgs) Handles OutputDoc.Click
        If OutputDoc.Image.Equals(semYes.Image) Then
            OutputDoc.Image = semNo.Image
            OutputPDF.Image = semYes.Image

            ' semInd.Image = semEvenInd.Image
        Else
            OutputDoc.Image = semYes.Image
            OutputPDF.Image = semNo.Image

            ' semInd.Image = semOddInd.Image

        End If
    End Sub
    Private Sub OutputBorder_Click(sender As Object, e As EventArgs) Handles OutputBorder.Click
        If OutputBorder.Image.Equals(semYes.Image) Then
            OutputBorder.Image = semNo.Image
            OutputNoBorder.Image = semYes.Image
            ' semInd.Image = semEvenInd.Image
        Else
            OutputBorder.Image = semYes.Image
            OutputNoBorder.Image = semNo.Image
            'semInd.Image = semOddInd.Image
        End If
    End Sub

    Private Sub OutputNoBorder_Click(sender As Object, e As EventArgs) Handles OutputNoBorder.Click
        If OutputNoBorder.Image.Equals(semYes.Image) Then
            OutputNoBorder.Image = semNo.Image
            OutputBorder.Image = semYes.Image

            ' semInd.Image = semEvenInd.Image
        Else
            OutputNoBorder.Image = semYes.Image
            OutputBorder.Image = semNo.Image

            ' semInd.Image = semOddInd.Image

        End If
    End Sub
    Private Sub rst_Click(sender As Object, e As EventArgs) Handles rst.Click
        Application.Restart()
    End Sub

    Private Sub rst_MouseEnter(sender As Object, e As EventArgs) Handles rst.MouseEnter
        rst.Image = rst2.Image
    End Sub

    Private Sub rst_MouseLeave(sender As Object, e As EventArgs) Handles rst.MouseLeave
        rst.Image = rst1.Image
    End Sub

    Private Sub tpshm_Click(sender As Object, e As EventArgs) Handles tpshm.Click
        reportStatus = "TopSheet"
        If routineStatus = False Then
            ' Try
            'Create_Routine()
            ' Catch ex As Exception

            '  End Try

        End If
        Open_Report_Preferences()
        reportMain.Visible = False
        LabelTopSheet.Visible = True
        LabelMax.Visible = True
        LabelDateTop.Visible = True
        ComboBox1.Visible = True
        ComboBox1.Enabled = True
        TextBoxNoOfAnswerScriptsInAPacket.Visible = True
        TextBoxNoOfAnswerScriptsInAPacket.Enabled = True

    End Sub

    Private Sub tpshm_MouseEnter(sender As Object, e As EventArgs) Handles tpshm.MouseEnter
        tpshm.Image = tpshg.Image
    End Sub

    Private Sub tpshm_MouseLeave(sender As Object, e As EventArgs) Handles tpshm.MouseLeave
        tpshm.Image = tpshb.Image
    End Sub

    Private Sub dblr_Click(sender As Object, e As EventArgs) Handles dblr.Click

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        'Count_No_Of_AbsentStudents()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        'MessageBox.Show(37 Mod 2)
    End Sub

    Private Sub abm_Click(sender As Object, e As EventArgs) Handles abm.Click
        reportStatus = "Absentee"
        ComboBox1.Visible = True
        ComboBox1.Enabled = True
        Open_Report_Preferences()
    End Sub

    Private Sub abm_MouseEnter(sender As Object, e As EventArgs) Handles abm.MouseEnter
        abm.Image = abg.Image
    End Sub

    Private Sub abm_MouseLeave(sender As Object, e As EventArgs) Handles abm.MouseLeave
        abm.Image = abb.Image
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        Write_Receipt()
    End Sub

    Private Sub TitleBar_Click(sender As Object, e As EventArgs) Handles TitleBar.Click

    End Sub

    Private Sub Label4_Click(sender As Object, e As EventArgs) Handles Label4.Click

    End Sub

    Private Sub sit_Click(sender As Object, e As EventArgs) Handles sit.Click
        reportStatus = "Receipt"
        ComboBox1.Visible = True
        ComboBox1.Enabled = True
        Open_Report_Preferences()
    End Sub

    Private Sub sit_MouseEnter(sender As Object, e As EventArgs) Handles sit.MouseEnter
        sit.Image = sitg.Image
    End Sub

    Private Sub sit_MouseLeave(sender As Object, e As EventArgs) Handles sit.MouseLeave
        sit.Image = sitb.Image
    End Sub

    Private Sub LabelNoBorder_Click(sender As Object, e As EventArgs) Handles LabelNoBorder.Click

    End Sub
End Class
Public Class DateInfo
    Public DateValue As String
    Public stream(100) As String
    Public semester(100) As Integer
    Public paperCode(100) As String
    Public half(100) As String
    Public paperName(100) As String
    Public backlog(100) As Boolean
    Public noOfBacklogStudents(100) As Integer
    Public noOfStudents(10000) As Integer
    Public absentStudent(1000) As Integer
    Public noOfPackets(100) As Integer
    Dim noOfCopies(100) As Integer
End Class
Public Class BacklogStudent
    Public StudentName As String
    Public StudentRollNo As String
    Public BacklogPaperOne As String
    Public BacklogPaperTwo As String
    Public BacklogPaperThree As String
End Class
Public Class Paper
    Public StudentRollNo(500) As String
    Public StudentName(500) As String
    Public PaperCode As String
    Public PaperName As String
End Class
Public Class Student
    Public StudentRollNo As String
    Public StudentStream As String
    Public present As Boolean
    Public semester As Integer
    Public half As Integer
End Class
Public Class AbsentStudent
    Public StudentRollNo As String
    Public StudentStream As String
    Public present As Boolean
    Public semester As Integer
    Public half As Integer
End Class
Public Class halfInfo
    Public semester As Integer
    Public half As Integer
End Class