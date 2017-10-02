Imports System.Globalization
Imports System.IO
Imports System.Math
Imports System.Threading
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


    Public Shared howto() As String = {"How to ..",
   "Iterate the grey bearing house temperature",
   "until the purple values (generated and dissipated power) are identical."}

    'sleeve bearing housing; diam; cooling area; L/d ratio
    'Dodge always air cooling
    Public Shared b_house() As String = {
   "Renk ERWLQ 09 ;80-100; 0.37; 1.0",
   "Renk ERWLQ 11 ;100-125; 0.64; 1.0",
   "Renk ERWLQ 14 ;125-160; 0.74; 1.0",
   "Renk ERWLQ 18 ;160-200; 1.13; 1.0",
   "Renk ERWLQ 22 ;200-250; 1.48; 1.0",
   "Dodge 3-7/16"";87.31; 0.094; 1.8",
   "Dodge 3-15/16"";100.01; 0.093; 1.8",
   "Dodge 4-7/15"";112.71; 0.125; 1.8",
   "Dodge 4-15/16"";125.41; 0.188; 1.8",
   "Dodge 5-7/16"" ;138.11; 0.265; 1.8",
   "Dodge 6"";152.40; 0.412; 1.8",
   "Dodge 7"";177.80; 0.680; 1.8",
   "Dodge 8"";203.20; 1.059; 1.8",
   "Dodge 9"";228.60; 1.672; 1.8",
   "Dodge 10"";254.00; 2.574; 1.8",
   "Dodge 12"";304.80; 4.903; 1.8"}

    Public Shared sleeve_LD_ratio() As String = {
    "Sleeve Length/dia ratio",
    "Text book ~ 0.8",
    "Renk      ~ 1.0",
    "DVB       ~ 1.0",
    "Dodge     ~ 1.8"}

    'oil type; kin visco [mm2/s=cP]; Density [kg/m3]
    Public Shared oil() As String = {
    "ISO 3448 VG 32; 32; 857",
    "ISO 3448 VG 46; 46; 861",
    "ISO 3448 VG 68; 68; 865",
    "ISO 3448 VG 100; 100; 869"}

    'Seal material type; Young modulus [GPa]
    'From https://www.matbase.com/material-categories/natural-and-synthetic-polymers/
    Public Shared seal_mat() As String = {
    "PTFE (Teflon); 0.75",
    "Polypropylene; 1.40",
    "NBR (Nitrile rubber); 0.005",
    "Polyamide 11 (Nylon); 1.4",
    "Polyamide 12 (Nylon); 2.6",
    "Polyamide 46 (Nylon); 3.0",
    "Polyamide 66 (Nylon); 2.0",
    "Silicone Rubber; 0.005",
    "Fluor elastomere (Viton); 0.005"}

    'Explanation "Metal;Temp;[W/mK]"
    Public Shared mat_conductivity() As String = {
    "Admiralty Brass;20;111",
    "Aluminum-pure;93;215",
    "Aluminum-Bronze;20;76",
    "Antimony;20;19",
    "Beryllium;20;218",
    "Beryllium Copper;20;66",
    "Bismuth;20;8.5",
    "Cadmium;20;93",
    "Carbon Steel, 0.5% C (plate);20;54",
    "Carbon Steel, C45 (shaft steel);20;46",
    "Carbon Steel, 1.5% C (alloy);20;36",
    "Cartridge brass (UNS C26000);20;120",
    "Cast Iron, gray;21;60",
    "Chromium;20;90",
    "Cobalt;20;69",
    "Copper-pure;20;386",
    "Copper bronze (75% Cu, 25% Sn);20;26",
    "Copper brass (70% Cu, 30% Zi);20;111",
    "Cupronickel;20;29",
    "Gold;20;315",
    "Hastelloy B;20;10",
    "Hastelloy C;21;8.7",
    "Inconel or Alloy 600 @ 93c;93;15.7",
    "Inconel or Alloy 600 @ 200c;200;17.4",
    "Inconel or Alloy 600 @ 427c;427;20.9",
    "Iridium;100;147",
    "Iron-nodular pearlitic;100;31",
    "Iron-pure;20;73",
    "Iron-wrought;20;59",
    "Lead;20;35",
    "Manganese Bronze;20;106",
    "Magnesium;20;159",
    "Mercury;20;8.4",
   "Molybdenum;20;140",
   "Monel;100;26",
   "Nickel;20;90",
   "Nickel Wrought;100;70",
   "Niobium (Columbium);20;52",
   "Osmium;20;61",
   "Phosphor bronze (10% Sn, UNS C52400);20;50",
   "Platinum;20;73",
   "Plutonium;20;8.0",
   "Potassium;20;100",
   "Red Brass;20;159",
   "Rhodium;20;150",
   "Selenium;20;0.52",
   "Silicon;20;84",
   "Silver-pure;20;407",
   "Sodium;20;134",
   "Stainless 304(L) @ 100c;100;16.9",
   "Stainless 304(L) @ 315c;315;17.3",
   "Stainless 304(L) @ 538c;538;18.4",
   "Stainless 316(L) @ 100c;100;16.2",
   "Stainless 316(L) @ 500c;500;21.4",
   "Tantalum;20;54",
   "Thorium;20;42",
   "Tin;0;65",
   "Titanium;20;21",
   "Tungsten;20;168",
   "Uranium;20;24",
   "Vanadium;20;61",
   "Wrought Carbon Steel;0;59",
   "Yellow Brass;20;116",
   "Zinc;20;116",
   "Zirconium;0;23"}

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim words() As String
        Dim separators() As String = {";"}
        Dim hh As Integer

        '------------Geeft problemen binnen VTK---------------
        'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        'Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        For hh = 0 To (transfer.Length - 1)
            TextBox8.Text &= transfer(hh) & vbCrLf
        Next hh

        For hh = 0 To (Howto.Length - 1)
            TextBox37.Text &= howto(hh) & vbCrLf
        Next hh

        For hh = 0 To (sleeve_LD_ratio.Length - 1)
            TextBox40.Text &= sleeve_LD_ratio(hh) & vbCrLf
        Next hh

        '-------Fill combobox1,2 and 5 Steel selection------------------
        For hh = 0 To (mat_conductivity.Length - 2)  'Fill combobox3 with steel data
            words = mat_conductivity(hh).Split(separators, StringSplitOptions.None)
            ComboBox1.Items.Add(words(0))
            ComboBox2.Items.Add(words(0))

        Next hh

        For hh = 0 To (b_house.Length - 1)            'Fill combobox3 with steel data
            words = b_house(hh).Split(separators, StringSplitOptions.None)
            ComboBox3.Items.Add(words(0))
        Next hh

        For hh = 0 To (oil.Length - 1)            'Fill combobox4 with oil data
            words = oil(hh).Split(separators, StringSplitOptions.None)
            ComboBox4.Items.Add(words(0))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 9, -1))   'C45
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 1, -1))   'Aluminium-pure
        ComboBox3.SelectedIndex = CInt(IIf(ComboBox3.Items.Count > 0, 2, -1))   'Renk
        ComboBox4.SelectedIndex = CInt(IIf(ComboBox4.Items.Count > 0, 1, -1))   'Oil selection
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown14.ValueChanged
        Calc_shaft()
    End Sub

    Private Sub Calc_shaft()
        Dim shaft_OD, f_id, F_length, F_coeff, F_temp, Shaft_area As Double
        Dim d_no, d_od, d_hub_od, d_Heat_trans, d_thick, fin_height As Double
        Dim d_area_actual, d_area_calc, area_factor1, area_factor2, area_eff, d_coeff As Double
        Dim power_conducted, power_transferred As Double
        Dim dT_conduct, dT_transfer As Double
        Dim temp_fan, temp_amb, temp_disk As Double
        Dim i As Integer

        Calc_transfer()

        '-------------- temps ----------------
        temp_fan = NumericUpDown5.Value
        temp_amb = NumericUpDown14.Value
        temp_disk = (temp_amb + temp_fan) / 2

        '-------------- shaft-----------------
        shaft_OD = NumericUpDown1.Value / 1000      '[mm]->[m]
        f_id = NumericUpDown2.Value / 1000
        F_length = NumericUpDown3.Value / 1000
        F_coeff = NumericUpDown4.Value
        F_temp = NumericUpDown5.Value
        Shaft_area = Math.PI / 4 * (shaft_OD ^ 2 - f_id ^ 2)

        '-------------- disk-----------------
        d_no = NumericUpDown11.Value    'Number of disks
        d_od = NumericUpDown9.Value / 1000
        d_hub_od = NumericUpDown7.Value / 1000
        d_thick = NumericUpDown10.Value / 1000

        Double.TryParse(TextBox15.Text, d_Heat_trans)
        Double.TryParse(TextBox20.Text, d_coeff)
        d_area_actual = d_no * 2 * Math.PI / 4 * (d_od ^ 2 - d_hub_od ^ 2)  'Natural log !!

        fin_height = (d_od - d_hub_od) / 2
        area_factor1 = fin_height * (d_Heat_trans / (d_coeff * 0.5 * d_thick)) ^ 0.5
        area_factor2 = 1 + 0.35 * Math.Log(d_od / d_hub_od)

        area_eff = Math.Tanh(area_factor1 * area_factor2) / (area_factor1 * area_factor2)
        d_area_calc = d_area_actual * area_eff

        '-------------- heat ---------------
        If temp_disk > 0 Then        'Preventing VB start problems!!
            For i = 0 To 500
                dT_conduct = temp_fan - temp_disk
                dT_transfer = temp_disk - temp_amb

                power_conducted = Shaft_area * dT_conduct * F_coeff / F_length
                power_transferred = dT_transfer * d_area_calc * d_Heat_trans

                If Abs(power_conducted - power_transferred) < 2 Then
                    Exit For        'Speeding things up
                End If

                If (power_conducted > power_transferred) Then
                    temp_disk += 0.5
                Else
                    temp_disk -= 0.5
                End If
            Next

            TextBox7.Text = Math.Round(temp_disk, 1).ToString
        End If

        TextBox1.Text = Shaft_area.ToString("0.000")
        TextBox2.Text = d_area_actual.ToString("0.00")
        TextBox3.Text = Math.Round(area_eff, 3).ToString
        TextBox4.Text = Math.Round(d_area_calc, 2).ToString
        TextBox5.Text = Math.Round(power_conducted, 0).ToString
        TextBox6.Text = Math.Round(power_transferred, 0).ToString

        'Checks
        TextBox5.BackColor = CType(IIf(Abs(power_conducted - power_transferred) > 30, Color.Red, Color.White), Color)
        TextBox6.BackColor = TextBox5.BackColor
        NumericUpDown7.BackColor = CType(IIf(NumericUpDown7.Value <= NumericUpDown1.Value + 20, Color.Red, Color.Yellow), Color)
        NumericUpDown9.BackColor = CType(IIf(NumericUpDown9.Value <= NumericUpDown7.Value + 50, Color.Red, Color.Yellow), Color)
    End Sub

    Private Sub Calc_transfer()
        Dim d_od, d_id, dia, speed, Reynolds, ht, ro_air, ka_air, vel, mu, safety As Double
        Dim nusselt As Double

        d_od = NumericUpDown9.Value / 1000      '[mm]->[m]
        d_id = NumericUpDown7.Value / 1000      '[mm]->[m]
        dia = (d_od + d_id) / 2
        speed = NumericUpDown12.Value           '[rpm]
        vel = speed / 60 * PI * d_od            '[m/s]

        ka_air = 0.0257                         '[W/mK]conductivity air
        ro_air = 1.205                          '[ro] air

        mu = 1.983 / 10 ^ 5                     'dyn visco air [Pa.s]
        safety = 0.3

        Reynolds = ro_air * vel * dia / mu

        'See Ain Shams Engineering journal (2014) 5, 177-185
        If Reynolds >= 1000 And Reynolds < 1000000 Then
            nusselt = 0.022 * Reynolds ^ 0.821
            ht = nusselt * ka_air / dia     '[W/m2K]
        End If

        If Reynolds < 1000 Then
            nusselt = 10
            ht = nusselt * ka_air / dia     '[W/m2K]
        End If

        TextBox61.Text = ro_air.ToString("0.000")   '[kg/m3] air
        TextBox58.Text = nusselt.ToString("0")      '[W/mK]conductivity air
        TextBox57.Text = ka_air.ToString("0.000")   '[W/mK]conductivity air
        TextBox12.Text = d_od.ToString("0.00")
        TextBox13.Text = Reynolds.ToString("0")
        TextBox14.Text = mu.ToString
        TextBox15.Text = ht.ToString("0.0")
        TextBox19.Text = TextBox15.Text
        TextBox17.Text = Math.Round(vel, 1).ToString
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
            oPara2.Range.Text = "Fan cooling disk sizing" & vbCrLf
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
            oTable.Cell(row, 1).Range.Text = "Fan type"
            oTable.Cell(row, 2).Range.Text = TextBox11.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Author"
            oTable.Cell(row, 2).Range.Text = Environment.UserName
            row += 1
            oTable.Cell(row, 1).Range.Text = "Date"
            oTable.Cell(row, 2).Range.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            row += 1
            oTable.Columns(1).Width = oWord.InchesToPoints(2.5)   'Change width of columns 
            oTable.Columns(2).Width = oWord.InchesToPoints(4)

            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            '------------------ Fan data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 8, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Fan shaft dimensions"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft material"
            oTable.Cell(row, 2).Range.Text = ComboBox1.Text
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
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown4.Value, 1).ToString
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
            oTable.Columns(2).Width = oWord.InchesToPoints(2.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()

            '------------------ Cooling disk data----------------------
            'Insert a table, fill it with data and change the column widths.
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Cooling disk data"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Disk material"
            oTable.Cell(row, 2).Range.Text = ComboBox2.Text
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
            oTable.Cell(row, 1).Range.Text = "Disk conductivity coeff"
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[W/mK]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Disk heat transfer (external)"
            oTable.Cell(row, 2).Range.Text = TextBox15.Text
            oTable.Cell(row, 3).Range.Text = "[W/m2K]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Effective disk area"
            oTable.Cell(row, 2).Range.Text = TextBox4.Text
            oTable.Cell(row, 3).Range.Text = "[m2]"

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(2.1)
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
            oTable.Cell(row, 3).Range.Text = "[°c]"
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
            oTable.Cell(row, 3).Range.Text = "[°c]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(2.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            ufilename = "Fan_cooling_disk_report_" & TextBox9.Text & "_" & TextBox10.Text & DateTime.Now.ToString("_yyyy_MM_dd") & "(" & TextBox3.Text & ")" & ".docx"
            If Directory.Exists(dirpath_Rap) Then
                ufilename = dirpath_Rap & ufilename
            Else
                ufilename = dirpath_Home & ufilename
            End If
            oWord.ActiveDocument.SaveAs(ufilename.ToString)
        Catch ex As Exception
            MessageBox.Show(ex.Message & "Problem storing file to" & dirpath_Rap)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim separators() As String = {";"}

        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = mat_conductivity(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            NumericUpDown4.Value = CDec(words(2))       'Conductivity fan shaft
        End If
        Calc_shaft()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim separators() As String = {";"}

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = mat_conductivity(ComboBox2.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox20.Text = words(2)       'Conductivity cooling disk
        End If
        Calc_shaft()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, NumericUpDown12.ValueChanged, NumericUpDown12.Enter, TabControl1.Enter, TabPage3.Click
        Calc_shaft()
    End Sub

    'Sleeve bearing calculation
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click, TabPage4.Enter, NumericUpDown18.ValueChanged, NumericUpDown17.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown22.ValueChanged, NumericUpDown21.ValueChanged, NumericUpDown20.ValueChanged, NumericUpDown19.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown23.ValueChanged
        Calc_sleeve_bearing()
    End Sub
    Private Sub Calc_sleeve_bearing()
        Dim load_kg, load_N As Double
        Dim dia, b_length, speed, rps As Double
        Dim clearance, clear_ratio, renk_clear As Double
        Dim coeff, power_loss As Double
        Dim oil_pr, oil_dyn, oil_Kin, density As Double
        Dim heat_loss_house, area_house, coeff_house, dt As Double
        Dim friction_torque As Double
        Dim separators() As String = {";"}
        Dim temp As Decimal

        '-----------Renk cooling area-----------
        If (ComboBox3.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = b_house(ComboBox3.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox39.Text = words(1)               'diameter area
            TextBox29.Text = words(2)               'cooling area
            Decimal.TryParse(words(3), temp)
            IIf(temp > NumericUpDown19.Minimum And temp > NumericUpDown19.Maximum, temp, 1)
        End If

        If (ComboBox4.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = oil(ComboBox4.SelectedIndex).Split(separators, StringSplitOptions.None)
            Double.TryParse(words(1), oil_Kin)          'Kin viscosity
            Double.TryParse(words(2), density)          'Density
            oil_dyn = oil_Kin * density / 10 ^ 6        '[mm2/s]--> [N.s/m2]
            TextBox32.Text = oil_dyn.ToString("0.000")  '[N.s/m2]
        End If

        '----- load ----
        load_kg = NumericUpDown18.Value             '[kg]
        load_N = load_kg * 10                       '[N]
        TextBox26.Text = load_N.ToString("00")

        '------ speed in bearing----
        dia = NumericUpDown17.Value / 1000                  '[m]
        b_length = dia * NumericUpDown19.Value              '[m]
        TextBox30.Text = (b_length * 1000).ToString("00")   '[mm]

        rps = NumericUpDown16.Value / 60            '[rotation per second]
        speed = Math.PI * dia * rps                 '[m/s]
        TextBox22.Text = speed.ToString("0.0")

        '----- RENK clearance---
        If speed < 10 Then  'Speed < 10 [m/s]
            Select Case True
                Case dia < 0.1
                    renk_clear = 1.6 / 10 ^ 3
                Case dia >= 0.1 And dia < 0.25
                    renk_clear = 1.32 / 10 ^ 3
                Case dia >= 0.25
                    renk_clear = 1.12 / 10 ^ 3
            End Select
        Else            'Speed >= 10 [m/s]
            Select Case True
                Case dia < 0.1
                    renk_clear = 1.9 / 10 ^ 3
                Case dia >= 0.1 And dia < 0.25
                    renk_clear = 1.6 / 10 ^ 3
                Case dia >= 0.25
                    renk_clear = 1.32 / 10 ^ 3
            End Select
        End If
        TextBox35.Text = (renk_clear * 10 ^ 3).ToString("0.0")

        '----- clearance---
        clear_ratio = 1 / (2 * renk_clear)
        TextBox36.Text = clear_ratio.ToString("0.0")
        clearance = (dia * 0.5) / clear_ratio
        TextBox21.Text = (clearance * 1000).ToString("0")   '[mu]

        '--- oil film pressure----
        oil_pr = load_N / (dia * b_length)                  '[Pa]
        TextBox31.Text = (oil_pr / 10 ^ 6).ToString("0.00") '[MPa]
        If oil_pr > 2.5 * 10 ^ 6 Or oil_pr < 0.5 * 10 ^ 6 Then
            TextBox31.BackColor = Color.Red
        Else
            TextBox31.BackColor = Color.LightGreen
        End If

        '---- Petroff's equation--
        coeff = 2 * PI ^ 2 * oil_dyn * rps / oil_pr * clear_ratio
        TextBox33.Text = coeff.ToString("0.0000")    '[-]

        '--- power loss due to friction
        friction_torque = coeff * load_N * dia / 2    '[Nm]   
        power_loss = friction_torque * rps * 2 * PI    '[W]
        TextBox34.Text = friction_torque.ToString("0.0")
        TextBox23.Text = power_loss.ToString("00")

        '-----Heat loss house----
        coeff_house = NumericUpDown22.Value
        dt = NumericUpDown21.Value - NumericUpDown20.Value
        Double.TryParse(TextBox29.Text, area_house)         '[m2]
        heat_loss_house = area_house * dt * coeff_house     '[W]

        TextBox24.Text = area_house.ToString("0.00")
        TextBox25.Text = heat_loss_house.ToString("00")

        '============================================
        'Heat flowing trough shaft to the bearing
        Dim ht, s_area, power_con, c_length, dtt As Double

        ht = NumericUpDown4.Value       'Heat transfer coeff
        c_length = NumericUpDown8.Value / 1000  'cooling disk-bearing
        dtt = NumericUpDown23.Value - NumericUpDown6.Value
        s_area = PI / 4 * dia ^ 2       'shaft section area

        power_con = s_area * dtt * ht / c_length
        TextBox29.Text = ht.ToString("0.00")
        TextBox38.Text = power_con.ToString("0")
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        Calc_sleeve_bearing()
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        Calc_sleeve_bearing()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown31.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown27.ValueChanged, TabPage5.Enter
        Dim fric_coef, no_seals, pwr_seal As Double
        Dim force, torque, rpm, diam, omega As Double

        no_seals = NumericUpDown25.Value        '[-]
        fric_coef = NumericUpDown31.Value       '[-]
        force = NumericUpDown29.Value           '[N]
        rpm = NumericUpDown12.Value             '[rpm]
        diam = NumericUpDown27.Value / 1000     '[m]

        torque = force * fric_coef * (diam / 2) '[N.m]
        omega = rpm / 60 * 2 * PI               '[rad/s]
        pwr_seal = omega * torque * no_seals    '[W]

        TextBox45.Text = torque.ToString("0.00")
        TextBox44.Text = omega.ToString("0.0")
        TextBox43.Text = pwr_seal.ToString("0.0")

        '----------- shaft area -----------------
        Dim shaft_L, shaft_area, ht_coef, Pwr_air As Double
        Dim dt, dt_average As Double

        Double.TryParse(TextBox15.Text, ht_coef) '[W/m2K]
        shaft_L = NumericUpDown28.Value / 1000          '[m]
        shaft_area = 2 * shaft_L * diam * PI                '[m2]

        '-------------- heat ---------------
        dt = 0
        For i = 0 To 1400
            dt_average = dt / 2                             '[c]
            Pwr_air = dt_average * shaft_area * ht_coef     '[W]


            If Abs(Pwr_air - pwr_seal) < 0.1 Then
                Exit For        'Speeding things up
            End If

            If (Pwr_air < pwr_seal) Then
                dt += 0.1
            Else
                dt -= 0.5
            End If
        Next

        '------------------------------
        TextBox41.Text = shaft_area.ToString("0.000")
        TextBox51.Text = Pwr_air.ToString("0.0")
        TextBox48.Text = dt.ToString("0.0")
        TextBox56.Text = ht_coef.ToString("0.0")
        TextBox59.Text = rpm.ToString("00")
    End Sub

End Class
