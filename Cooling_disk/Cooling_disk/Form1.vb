Imports System.IO
Imports System.Math
Imports Word = Microsoft.Office.Interop.Word


Public Class Form1

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Fan_sizing_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Fan_rapport_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

    Public Shared transfer() As String = {"McPhee and Johnson (2007) employed experimental and",
    "analytical methods for better understanding of convection through the fins of a brake rotor",
    "The experimental approach involved two aspects, assessment of both heat transfer And fluid motion",
    "A transient experiment was conducted to quantify the internal (fin) convection And external (rotor surface)",
    "convection terms for three nominal speeds.",
    "For the given experiment, conduction And radiation were determined to be negligible.",
    "Rotor rotational speeds of 342, 684 And 1025 rpm yielded fin convection heat transfer",
    "coefficients of 27.0, 52.7, 78.3 Wm-2 K-1, respectively, indicating a linear relationship.",
    "At the slowest speed, the internal convection represented 45.5% of the total heat transfer, increasing to 55.4% at 1025 rpm.",
    "The flow aspect of the experiment involved the determination of the velocity field through the internal passages formed by the radial fins.",
    "Utilizing PIV, the phase-averaged velocity field was determined.",
    "A number of detrimental flow patterns were observed, notably entrance effects",
    "and the presence of recirculation on the suction side of the fins"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For hh = 0 To (transfer.Length - 1)
            TextBox8.Text &= transfer(hh) & vbCrLf
        Next hh
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown14.ValueChanged
        Calc_shaft()
    End Sub

    Private Sub Calc_shaft()
        Dim f_od, f_id, F_length, F_coeff, F_temp, Shaft_area As Double
        Dim d_no, d_od, d_hub_od, d_Heat_trans, d_thick, fin_height As Double
        Dim d_area_actual, d_area_calc, area_factor1, area_factor2, area_eff As Double
        Dim power_conducted, power_transferred As Double
        Dim dT_conduct, dT_transfer As Double
        Dim temp_fan, temp_amb, temp_disk As Double
        Dim i As Integer

        '-------------- temps ----------------
        temp_fan = NumericUpDown5.Value
        temp_amb = NumericUpDown14.Value
        temp_disk = (temp_amb + temp_fan) / 2

        '-------------- shaft-----------------
        f_od = NumericUpDown1.Value / 1000      '[mm]->[m]
        f_id = NumericUpDown2.Value / 1000
        F_length = NumericUpDown3.Value / 1000
        F_coeff = NumericUpDown4.Value
        F_temp = NumericUpDown5.Value
        Shaft_area = Math.PI / 4 * (f_od ^ 2 - f_id ^ 2)

        '-------------- disk-----------------
        d_no = NumericUpDown11.Value    'Number of disks
        d_od = NumericUpDown9.Value / 1000
        d_hub_od = NumericUpDown7.Value / 1000
        d_thick = NumericUpDown10.Value / 1000
        d_Heat_trans = NumericUpDown6.Value
        d_area_actual = d_no * 2 * Math.PI / 4 * (d_od ^ 2 - d_hub_od ^ 2)  'Natural log !!


        fin_height = (d_od - d_hub_od) / 2
        area_factor1 = fin_height * (d_Heat_trans / (F_coeff * 0.5 * d_thick)) ^ 0.5
        area_factor2 = 1 + 0.35 * Math.Log(d_od / d_hub_od)

        area_eff = Math.Tanh(area_factor1 * area_factor2) / (area_factor1 * area_factor2)
        d_area_calc = d_area_actual * area_eff

        '-------------- heat ---------------
        If temp_disk > 0 Then        'Preventing VB start problems!!
            For i = 0 To 350
                dT_conduct = temp_fan - temp_disk
                dT_transfer = temp_disk - temp_amb

                power_conducted = Shaft_area * dT_conduct * F_coeff / F_length
                power_transferred = dT_transfer * d_area_calc * d_Heat_trans


                If (power_conducted > power_transferred) Then
                    temp_disk += 1
                Else
                    temp_disk -= 1
                End If
            Next

            TextBox7.Text = Math.Round(temp_disk, 1).ToString
        End If

        TextBox1.Text = Math.Round(Shaft_area, 2).ToString
        TextBox2.Text = Math.Round(d_area_actual, 2).ToString
        TextBox3.Text = Math.Round(area_eff, 3).ToString
        TextBox4.Text = Math.Round(d_area_calc, 2).ToString

        TextBox5.Text = Math.Round(power_conducted, 0).ToString
        TextBox6.Text = Math.Round(power_transferred, 0).ToString

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Write_to_word()
    End Sub

    'Write data to Word 
    'see https://msdn.microsoft.com/en-us/library/office/aa192495(v=office.11).aspx
    Private Sub Write_to_word()
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oTable As Word.Table
        Dim oPara1, oPara2 As Word.Paragraph
        Dim ufilename As String
        Dim row As Integer

        Try
            oWord = CType(CreateObject("Word.Application"), Word.Application)
            oWord.Visible = True
            oDoc = oWord.Documents.Add

            'Insert a paragraph at the beginning of the document. 
            oPara1 = oDoc.Content.Paragraphs.Add
            oPara1.Range.Text = "VTK Engineering"
            oPara1.Range.Font.Name = "Arial"
            oPara1.Range.Font.Size = 14
            oPara1.Range.Font.Bold = CInt(True)
            oPara1.Format.SpaceAfter = 1                '1 pt spacing after paragraph. 
            oPara1.Range.InsertParagraphAfter()

            oPara2 = oDoc.Content.Paragraphs.Add(oDoc.Bookmarks.Item("\endofdoc").Range)
            oPara2.Format.SpaceAfter = 1
            oPara2.Range.Font.Bold = CInt(False)
            oPara2.Range.Text = "Fan cooling disk sizing " & vbCrLf
            oPara2.Range.InsertParagraphAfter()

            '----------------------------------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 2)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 10
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)

            row = 1
            oTable.Cell(row, 1).Range.Text = "Project Name"
            oTable.Cell(row, 2).Range.Text = TextBox9.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Item number"
            oTable.Cell(row, 2).Range.Text = TextBox10.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Fan type "
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author "
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date "
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Fan data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Fan shaft dimensions"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft OD"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown1.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft ID"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown2.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Length casing-cooling disk"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown3.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Heat conductivity coeff"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown4.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[W/mK]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Max fan operating temp"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown5.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[c]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft cross section"
            oTable.Cell(row, 2).Range.Text = TextBox1.Text
            oTable.Cell(row, 3).Range.Text = "[m2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Cooling disk data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 7, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cooling disk data"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Number of disks"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown11.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Outside diameter disk"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown9.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Hub diameter"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown7.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uniform disk thcknes"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown10.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[W/mK]"
            row += 1

            oTable.Cell(row, 1).Range.Text = "Heat transfer (external)"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown6.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[W/m2K]"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Effective disk area"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[m2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Results data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 5, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Results"

            row += 1
            oTable.Cell(row, 1).Range.Text = "Ambient temperature"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown14.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[c]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Conducted power"
            oTable.Cell(row, 2).Range.Text = TextBox5.Text
            oTable.Cell(row, 3).Range.Text = "[W]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "To air transferred power"
            oTable.Cell(row, 2).Range.Text = TextBox6.Text
            oTable.Cell(row, 3).Range.Text = "[W]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Calculated shaft temperature"
            oTable.Cell(row, 2).Range.Text = TextBox7.Text
            oTable.Cell(row, 3).Range.Text = "[c]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(1.55)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            ufilename = "Fan_cooling_disk_report_" & TextBox2.Text & "_" & TextBox2.Text & DateTime.Now.ToString("_yyyy_MM_dd") & "(" & TextBox3.Text & ")" & ".docx"
            If Directory.Exists(dirpath_Rap) Then
                ufilename = dirpath_Rap & ufilename
            Else
                ufilename = dirpath_Home & ufilename
            End If
            oWord.ActiveDocument.SaveAs(ufilename.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " Problem storing file to " & dirpath_Rap)  ' Show the exception's message.
        End Try
    End Sub


End Class
