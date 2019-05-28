Imports System.Globalization
Imports System.Text
Imports System.IO
Imports System.Math
Imports System.Threading
Imports System.Management
Imports Word = Microsoft.Office.Interop.Word


Public Class Form1

    '----------- directory's-----------
    Dim dirpath_Eng As String = "N:\Engineering\VBasic\Cool_disk_input\"
    Dim dirpath_Rap As String = "N:\Engineering\VBasic\Cool_disk_copy\"
    Dim dirpath_Home As String = "C:\Temp\"

    Dim transfer() As String = {"McPhee and Johnson (2007) employed experimental and",
   "analytical methods for better understanding of convection through the fins of a brake rotor",
   "The experimental approach involved two aspects, assessment of both heat transfer And fluid motion",
   " ",
   "A transient experiment was conducted to quantify the internal (fin) convection And external (rotor surface)",
   "convection terms for three nominal speeds.",
   "For the given experiment, conduction And radiation were determined to be negligible.",
   "Rotor rotational speeds of 342, 684 And 1025 rpm yielded fin convection heat transfer",
   "coefficients of 27.0, 52.7, 78.3 W/m2.K1, respectively, indicating a linear relationship.",
   "At the slowest speed, the internal convection represented 45.5% of the total heat transfer, increasing to 55.4% at 1025 rpm.",
   "The flow aspect of the experiment involved the determination of the velocity field through the internal passages formed by the radial fins.",
   "Utilizing PIV, the phase-averaged velocity field was determined.",
   "A number of detrimental flow patterns were observed, notably entrance effects",
   "and the presence of recirculation on the suction side of the fins"}

    Dim howto() As String = {"How to ..",
   "Iterate the grey bearing house temperature",
   "until the purple values (generated and dissipated power) are identical."}

    'sleeve bearing housing; diam; cooling area; L/d ratio
    'Dodge always air cooling
    Dim b_house() As String = {
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

    'Troesner cooling disk
    'Type, Dia, Thick, wide, dia hub, dia_shaft_max
    Dim Troester() As String = {
    "K150   ;150;  3; 30; 60 ; 50",
    "K200   ;200;  5; 30; 60 ; 50",
    "K250   ;250;  5; 34; 85 ; 70",
    "K315   ;315;  7; 54; 112; 95",
    "K400   ;400;  8; 68; 145;125",
    "K400 SO;400;  8; 74; 170;155",
    "K500   ;500; 10; 68; 180;165",
    "K630   ;630; 11; 78; 225;200"}

    Dim sleeve_LD_ratio() As String = {
    "Sleeve Length/dia ratio",
    "Text book ~ 0.8",
    "Renk      ~ 1.0",
    "DVB       ~ 1.0",
    "Dodge     ~ 1.8"}

    'oil type; kin visco [mm2/s=cP]; Density [kg/m3]
    Dim oil() As String = {
    "ISO 3448 VG 32; 32; 857",
    "ISO 3448 VG 46; 46; 861",
    "ISO 3448 VG 68; 68; 865",
    "ISO 3448 VG 100; 100; 869"}

    Dim oil_temp() As String = {
    "Oil temperatures",
    "Mineral oil operate @ 110-126 °C",
    "Full synthetic operate @ 110- 148 °C",
    "Keep the shaft temp below 100 °C",
    "  "}

    'Explanation "Metal;Temp;[W/mK]"
    Dim mat_conductivity() As String = {
    "Admiralty Brass;20;111",
    "Aluminum-235 (Troester) Kuhl Scheibe;20;145",
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
        Dim Pro_user As String

        '----Noodzakelijk ivm punt en komma binnen mat_conductivity()--------
        Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
        Thread.CurrentThread.CurrentUICulture = New CultureInfo("en-US")

        Pro_user = Environment.UserName     'User name on the screen
        Me.Text = Me.Text & " (" & Pro_user & ")"

        For hh = 0 To (transfer.Length - 1)
            TextBox8.Text &= transfer(hh) & vbCrLf
        Next hh

        For hh = 0 To (Howto.Length - 1)
            TextBox37.Text &= howto(hh) & vbCrLf
        Next hh

        For hh = 0 To (sleeve_LD_ratio.Length - 1)
            TextBox40.Text &= sleeve_LD_ratio(hh) & vbCrLf
        Next hh


        For hh = 0 To (oil_temp.Length - 1)
            TextBox70.Text &= oil_temp(hh) & vbCrLf
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

        For hh = 0 To (Troester.Length - 1)            'Fill combobox5 with Troester disk data
            words = Troester(hh).Split(separators, StringSplitOptions.None)
            ComboBox5.Items.Add(words(0))
        Next hh

        '----------------- prevent out of bounds------------------
        ComboBox1.SelectedIndex = CInt(IIf(ComboBox1.Items.Count > 0, 10, -1))  'C45
        ComboBox2.SelectedIndex = CInt(IIf(ComboBox2.Items.Count > 0, 1, -1))   'Aluminium-235 
        ComboBox3.SelectedIndex = CInt(IIf(ComboBox3.Items.Count > 0, 2, -1))   'Renk
        ComboBox4.SelectedIndex = CInt(IIf(ComboBox4.Items.Count > 0, 1, -1))   'Oil selection
        ComboBox5.SelectedIndex = CInt(IIf(ComboBox5.Items.Count > 0, 1, -1))   'Troester

        TextBox9.Text = "P" & Now.ToString("yy") & ".10"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown14.ValueChanged, TabPage1.Enter, NumericUpDown12.ValueChanged
        Calc_transfer()
        Calc_shaft()
    End Sub

    Private Sub Calc_shaft()
        Dim shaft_OD, f_id, F_length, F_coeff, Shaft_area As Double
        Dim d_no, d_od, d_hub_od, d_Heat_transf, d_thick, fin_height As Double
        Dim d_area_actual, d_area_calc, area_factor1, area_factor2, fin_eff, Conduct As Double
        Dim power_conducted, power_transferred As Double
        Dim dT_conduct, dT_transfer As Double
        Dim temp_fan, temp_amb, temp_disk As Double
        Dim i As Integer

        '-------------- temps ----------------
        temp_fan = NumericUpDown5.Value
        temp_amb = NumericUpDown14.Value
        temp_disk = (temp_amb + temp_fan) / 2

        '-------------- shaft-----------------
        shaft_OD = NumericUpDown1.Value / 1000      '[mm]->[m]
        f_id = NumericUpDown2.Value / 1000
        F_length = NumericUpDown3.Value / 1000
        F_coeff = NumericUpDown4.Value
        Shaft_area = Math.PI / 4 * (shaft_OD ^ 2 - f_id ^ 2)

        '-------------- disk-----------------
        d_no = NumericUpDown11.Value    'Number of disks
        d_od = NumericUpDown9.Value / 1000      '[m]
        d_hub_od = NumericUpDown7.Value / 1000  '[m]
        d_thick = NumericUpDown10.Value / 1000  '[m]

        Double.TryParse(TextBox15.Text, d_Heat_transf)   '[W/m2k]
        Double.TryParse(TextBox20.Text, Conduct)        '[W/mk]
        d_area_actual = d_no * 2 * Math.PI / 4 * (d_od ^ 2 - d_hub_od ^ 2)

        fin_height = (d_od - d_hub_od) / 2
        area_factor1 = fin_height * (d_Heat_transf / (Conduct * 0.5 * d_thick)) ^ 0.5
        area_factor2 = 1 + 0.35 * Math.Log(d_od / d_hub_od) 'Natural log !!
        'MessageBox.Show(Conduct.ToString)

        fin_eff = Math.Tanh(area_factor1 * area_factor2) / (area_factor1 * area_factor2)
        d_area_calc = d_area_actual * fin_eff

        '-------------- heat ---------------
        If temp_disk > 0 Then        'Preventing VB start problems!!
            For i = 0 To 500
                dT_conduct = temp_fan - temp_disk
                dT_transfer = temp_disk - temp_amb

                power_conducted = Shaft_area * dT_conduct * F_coeff / F_length
                power_transferred = dT_transfer * d_area_calc * d_Heat_transf

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

        TextBox72.Text = area_factor1.ToString("0.00")
        TextBox73.Text = area_factor2.ToString("0.00")

        TextBox3.Text = Math.Round(fin_eff, 3).ToString
        TextBox4.Text = Math.Round(d_area_calc, 2).ToString
        TextBox5.Text = Math.Round(power_conducted, 0).ToString
        TextBox6.Text = Math.Round(power_transferred, 0).ToString

        'Checks
        TextBox7.BackColor = CType(IIf(temp_disk > 100, Color.Red, Color.White), Color)
        TextBox5.BackColor = CType(IIf(Abs(power_conducted - power_transferred) > 15, Color.Red, Color.White), Color)
        TextBox6.BackColor = TextBox5.BackColor
        NumericUpDown7.BackColor = CType(IIf(NumericUpDown7.Value <= NumericUpDown1.Value + 20, Color.Red, Color.Yellow), Color)
        NumericUpDown9.BackColor = CType(IIf(NumericUpDown9.Value <= NumericUpDown7.Value + 50, Color.Red, Color.Yellow), Color)
    End Sub

    Private Sub Calc_transfer()
        Dim disk_od, d_hub, dia, d_shaft, speed As Double
        Dim ro_air, ka_air, vel, vel_shaft, mu As Double
        Dim nusselt, nusselt_shaft, reynolds_disk, reynolds_shaft As Double
        Dim ht, ht_shaft As Double

        NumericUpDown12.Increment = CDec(IIf(NumericUpDown12.Value >= 50, 10, 1))     'Speed [rpm]

        disk_od = NumericUpDown9.Value / 1000   '[mm]->[m]
        d_hub = NumericUpDown7.Value / 1000     '[mm]->[m]
        dia = (disk_od + d_hub) / 2
        speed = NumericUpDown12.Value           '[rpm]
        vel = speed / 60 * PI * disk_od         '[m/s]

        ka_air = 0.0257                         '[W/mK]conductivity air
        ro_air = 1.205                          '[ro] air
        'http://www.engineeringtoolbox.com/dry-air-properties-d_973.html
        mu = 1.846 / 10 ^ 5                     'dyn visco air [Pa.s] @ 300K

        reynolds_disk = ro_air * vel * dia / mu

        If reynolds_disk >= 1300000 Then reynolds_disk = 1300000
        'See Ain Shams Engineering journal (2014) 5, 177-185

        If reynolds_disk >= 1000 And reynolds_disk <= 1300000 Then
            nusselt = 0.022 * reynolds_disk ^ 0.821
            ht = nusselt * ka_air / dia     '[W/m2K]
        End If

        If reynolds_disk < 1000 Then
            nusselt = 10
            ht = nusselt * ka_air / dia     '[W/m2K]
        End If

        TextBox61.Text = ro_air.ToString("0.000")   '[kg/m3] air
        TextBox58.Text = nusselt.ToString("0")      '[W/mK]conductivity air
        TextBox57.Text = ka_air.ToString("0.000")   '[W/mK]conductivity air
        TextBox12.Text = disk_od.ToString("0.00")
        TextBox13.Text = reynolds_disk.ToString("0")
        TextBox14.Text = mu.ToString
        TextBox15.Text = ht.ToString("0.0")
        TextBox19.Text = TextBox15.Text
        TextBox17.Text = Math.Round(vel, 1).ToString
        TextBox65.Text = speed.ToString("0")

        '===================== shaft only===============
        d_shaft = NumericUpDown1.Value / 1000       '[mm]
        vel_shaft = speed / 60 * PI * d_shaft
        reynolds_shaft = ro_air * vel_shaft * d_shaft / mu

        'See Ain Shams Engineering journal (2014) 5, 177-185
        If reynolds_shaft >= 1000 And reynolds_shaft < 1000000 Then
            nusselt_shaft = 0.022 * reynolds_shaft ^ 0.821
            ht_shaft = nusselt_shaft * ka_air / d_shaft     '[W/m2K]
        End If

        If reynolds_shaft < 1000 Then
            nusselt_shaft = 10
            ht_shaft = nusselt_shaft * ka_air / d_shaft     '[W/m2K]
        End If

        TextBox62.Text = ht_shaft.ToString("0.0")
        TextBox64.Text = nusselt_shaft.ToString("0")
        TextBox63.Text = d_shaft.ToString("0.000")
        TextBox66.Text = vel_shaft.ToString("0.0")
        TextBox18.Text = reynolds_shaft.ToString("0")
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
            oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, 9, 3)
            oTable.Range.ParagraphFormat.SpaceAfter = 1
            oTable.Range.Font.Size = 9
            oTable.Range.Font.Bold = CInt(False)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            row = 1
            oTable.Cell(row, 1).Range.Text = "Fan shaft"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft material"
            oTable.Cell(row, 2).Range.Text = ComboBox1.Text
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft OD"
            oTable.Cell(row, 2).Range.Text = NumericUpDown1.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft ID"
            oTable.Cell(row, 2).Range.Text = NumericUpDown2.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Distance casing-cooling disk"
            oTable.Cell(row, 2).Range.Text = NumericUpDown3.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Heat conductivity coeff"
            oTable.Cell(row, 2).Range.Text = NumericUpDown4.Value.ToString("0.0")
            oTable.Cell(row, 3).Range.Text = "[W/m.K]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Max fan operating temp"
            oTable.Cell(row, 2).Range.Text = NumericUpDown5.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[°C]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Shaft speed"
            oTable.Cell(row, 2).Range.Text = NumericUpDown12.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[rpm]"
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
            oTable.Cell(row, 2).Range.Text = NumericUpDown11.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[-]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Outside diameter disk"
            oTable.Cell(row, 2).Range.Text = NumericUpDown9.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Hub diameter"
            oTable.Cell(row, 2).Range.Text = Round(NumericUpDown7.Value, 0).ToString
            oTable.Cell(row, 3).Range.Text = "[mm]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Uniform disk thickness"
            oTable.Cell(row, 2).Range.Text = NumericUpDown7.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[W/m.K]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Disk conductivity coeff."
            oTable.Cell(row, 2).Range.Text = TextBox20.Text
            oTable.Cell(row, 3).Range.Text = "[W/m.K]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Disk heat transfer (external)"
            oTable.Cell(row, 2).Range.Text = TextBox15.Text
            oTable.Cell(row, 3).Range.Text = "[W/m2.K]"
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
            oTable.Cell(row, 2).Range.Text = NumericUpDown14.Value.ToString("0")
            oTable.Cell(row, 3).Range.Text = "[°C]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Conducted power"
            oTable.Cell(row, 2).Range.Text = TextBox5.Text
            oTable.Cell(row, 3).Range.Text = "[W]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Power to air transferred"
            oTable.Cell(row, 2).Range.Text = TextBox6.Text
            oTable.Cell(row, 3).Range.Text = "[W]"
            row += 1
            oTable.Cell(row, 1).Range.Text = "Calculated shaft temperature"
            oTable.Cell(row, 2).Range.Text = TextBox7.Text
            oTable.Cell(row, 3).Range.Text = "[°C]"
            row += 1

            oTable.Columns(1).Width = oWord.InchesToPoints(2.0)   'Change width of columns
            oTable.Columns(2).Width = oWord.InchesToPoints(2.1)
            oTable.Columns(3).Width = oWord.InchesToPoints(0.8)
            oTable.Rows.Item(1).Range.Font.Bold = CInt(True)
            oDoc.Bookmarks.Item("\endofdoc").Range.InsertParagraphAfter()


            ufilename = "Cool_disk_report_" & TextBox9.Text & "_" & TextBox10.Text & DateTime.Now.ToString("_yyyy_MM_dd") & "(" & TextBox3.Text & ")" & ".docx"
            If Directory.Exists(dirpath_Rap) Then
                ufilename = dirpath_Rap & ufilename
            Else
                ufilename = dirpath_Home & ufilename
            End If
            oWord.ActiveDocument.SaveAs(ufilename.ToString)
        Catch ex As DirectoryNotFoundException
            MessageBox.Show(ex.Message & "Problem storing file to" & dirpath_Rap)  ' Show the exception's message.
        End Try
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim separators() As String = {";"}

        If (ComboBox1.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = mat_conductivity(ComboBox1.SelectedIndex).Split(separators, StringSplitOptions.None)
            NumericUpDown4.Value = CDec(words(2))       'Conductivity fan shaft
        End If
        Calc_transfer()
        Calc_shaft()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        Dim separators() As String = {";"}

        If (ComboBox2.SelectedIndex > -1) Then          'Prevent exceptions
            Dim words() As String = mat_conductivity(ComboBox2.SelectedIndex).Split(separators, StringSplitOptions.None)
            TextBox20.Text = words(2)       'Conductivity cooling disk
        End If
        Calc_transfer()
        Calc_shaft()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click, TabControl1.Enter, TabPage3.Enter
        Calc_transfer()
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
        TextBox22.Text = speed.ToString("0.00")

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
        TextBox21.Text = (clearance * 1000).ToString("0.00")   '[mm]

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

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click, NumericUpDown31.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown25.ValueChanged, NumericUpDown28.ValueChanged, TabPage5.Enter
        Calc_transfer()
        Calc_seal()
    End Sub
    Private Sub Calc_seal()
        Dim fric_coef, no_seals, pwr_seal As Double
        Dim force, torque, rpm, diam, omega As Double

        no_seals = NumericUpDown25.Value        '[-]
        fric_coef = NumericUpDown31.Value       '[-]
        force = NumericUpDown29.Value           '[N]
        rpm = NumericUpDown12.Value             '[rpm]
        diam = NumericUpDown1.Value / 1000      '[m]

        torque = force * fric_coef * (diam / 2) '[N.m]
        omega = rpm / 60 * 2 * PI               '[rad/s]
        pwr_seal = omega * torque * no_seals    '[W]

        TextBox45.Text = torque.ToString("0.00")
        TextBox44.Text = omega.ToString("0.0")
        TextBox43.Text = pwr_seal.ToString("0.0")

        '----------- shaft area -----------------
        Dim shaft_L, shaft_area, ht_coef, Pwr_air As Double
        Dim dt, dt_average As Double

        Double.TryParse(TextBox62.Text, ht_coef)        '[W/m2K]
        shaft_L = NumericUpDown28.Value / 1000          '[m]
        shaft_area = 2 * shaft_L * diam * PI            '[m2]

        '-------------- heat ---------------
        dt = 0
        For i = 0 To 4000
            dt_average = dt / 2                         '[c]
            Pwr_air = dt_average * shaft_area * ht_coef '[W]

            If Abs(Pwr_air - pwr_seal) < 0.2 Then
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
        TextBox67.Text = (diam * 1000).ToString("0")

        TextBox48.BackColor = CType(IIf(dt > 40, Color.Red, Color.LightGreen), Color)
        TextBox59.BackColor = CType(IIf(rpm > 60, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click, NumericUpDown36.ValueChanged, NumericUpDown34.ValueChanged, NumericUpDown33.ValueChanged, TabPage6.Enter
        Calc_transfer()
        Calc_stuff()
    End Sub
    Private Sub Calc_stuff()
        Dim i As Integer
        Dim gland_pressure, fric_coef, pwr_gland, gland_l As Double
        Dim force, torque, rpm, diam, omega As Double

        gland_pressure = NumericUpDown34.Value * 10 ^ 5        '[-]
        fric_coef = NumericUpDown33.Value       '[-]

        rpm = NumericUpDown12.Value             '[rpm]
        diam = NumericUpDown1.Value / 1000      '[m]
        gland_l = NumericUpDown36.Value / 1000  '[m]
        force = gland_pressure * diam * gland_l '[N]

        torque = force * fric_coef * (diam / 2) '[N.m]
        omega = rpm / 60 * 2 * PI               '[rad/s]
        pwr_gland = omega * torque              '[W]

        '----------- shaft area -----------------
        Dim shaft_L, shaft_area, ht_coef, Pwr_air As Double
        Dim dt, dt_average As Double

        Double.TryParse(TextBox62.Text, ht_coef)        '[W/m2K]
        shaft_L = NumericUpDown28.Value / 1000          '[m]
        shaft_area = 2 * shaft_L * diam * PI            '[m2]

        '-------------- heat ---------------
        dt = 0
        For i = 0 To 4000
            dt_average = dt / 2                         '[c]
            Pwr_air = dt_average * shaft_area * ht_coef '[W]

            If Abs(Pwr_air - pwr_gland) < 0.2 Then
                Exit For        'Speeding things up
            End If

            If (Pwr_air < pwr_gland) Then
                dt += 0.1
            Else
                dt -= 0.5
            End If
        Next

        TextBox42.Text = force.ToString("0")
        TextBox49.Text = torque.ToString("0.0")
        TextBox60.Text = rpm.ToString("0")
        TextBox50.Text = omega.ToString("0.0")
        TextBox52.Text = pwr_gland.ToString("0.0")
        TextBox53.Text = shaft_area.ToString("0.000")
        TextBox54.Text = dt.ToString("0")
        TextBox55.Text = Pwr_air.ToString("0.0")
        TextBox68.Text = (diam * 1000).ToString("0")
        TextBox69.Text = ht_coef.ToString("0.0")
        TextBox54.BackColor = CType(IIf(dt > 100, Color.Red, Color.LightGreen), Color)
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Save_tofile()
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Read_file()
    End Sub
    'Retrieve control settings and case_x_conditions from file
    'Split the file string into 5 separate strings
    'Each string represents a control type (combobox, checkbox,..)
    'Then split up the secton string into part to read into the parameters
    Private Sub Read_file()
        Dim control_words(), words() As String
        Dim i As Integer
        Dim ttt As Double
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim separators() As String = {";"}
        Dim separators1() As String = {"BREAK"}

        OpenFileDialog1.FileName = "Cool_disk*"
        If Directory.Exists(dirpath_Eng) Then
            OpenFileDialog1.InitialDirectory = dirpath_Eng  'used at VTK
        Else
            OpenFileDialog1.InitialDirectory = dirpath_Home  'used at home
        End If

        OpenFileDialog1.Title = "Open a Text File"
        OpenFileDialog1.Filter = "VTK Files|*.vtk"

        If OpenFileDialog1.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Dim readText As String = File.ReadAllText(OpenFileDialog1.FileName, Encoding.ASCII)
            control_words = readText.Split(separators1, StringSplitOptions.None) 'Split the read file content

            '----- retrieve case condition-----
            words = control_words(0).Split(separators, StringSplitOptions.None) 'Split first line the read file content
            TextBox9.Text = words(0)                  'Project number
            TextBox10.Text = words(1)                 'Item name
            TextBox11.Text = words(2)                 'Fan type

            '---------- terugzetten numeric controls -----------------
            FindControlRecursive(all_num, Me, GetType(NumericUpDown))
            all_num = all_num.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(1).Split(separators, StringSplitOptions.None)     'Split the read file content
            For i = 0 To all_num.Count - 1
                Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal numeric controls--
                If (i < words.Length - 1) Then
                    If Not (Double.TryParse(words(i + 1), ttt)) Then MessageBox.Show("Numeric controls conversion problem occured")
                    If ttt <= grbx.Maximum And ttt >= grbx.Minimum Then
                        grbx.Value = CDec(ttt)          'OK
                    Else
                        grbx.Value = grbx.Minimum       'NOK
                        MessageBox.Show("Numeric controls value out of ousode min-max range, Minimum value is used")
                    End If
                Else
                    MessageBox.Show("Warning last Numeric controls not found in file")  'NOK
                End If
            Next

            '---------- terugzetten combobox controls -----------------
            FindControlRecursive(all_combo, Me, GetType(ComboBox))
            all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(2).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_combo.Count - 1
                Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    grbx.SelectedItem = words(i + 1)
                Else
                    MessageBox.Show("Warning last combobox not found in file")
                End If
            Next

            '---------- terugzetten checkbox controls -----------------
            FindControlRecursive(all_check, Me, GetType(CheckBox))
            all_check = all_check.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(3).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_check.Count - 1
                Dim grbx As CheckBox = CType(all_check(i), CheckBox)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal checkboxes--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last checkbox not found in file")
                End If
            Next

            '---------- terugzetten radiobuttons controls -----------------
            FindControlRecursive(all_radio, Me, GetType(RadioButton))
            all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()                  'Alphabetical order
            words = control_words(4).Split(separators, StringSplitOptions.None) 'Split the read file content
            For i = 0 To all_radio.Count - 1
                Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
                '--- dit deel voorkomt problemen bij het uitbreiden van het aantal radiobuttons--
                If (i < words.Length - 1) Then
                    Boolean.TryParse(words(i + 1), grbx.Checked)
                Else
                    MessageBox.Show("Warning last radiobutton not found in file")
                End If
            Next

        End If
    End Sub
    '----------- Find all controls on form1------
    'Nota Bene, sequence of found control may be differen, List sort is required
    Public Shared Function FindControlRecursive(ByVal list As List(Of Control), ByVal parent As Control, ByVal ctrlType As Type) As List(Of Control)

        If parent Is Nothing Then Return list

        If parent.GetType Is ctrlType Then
            list.Add(parent)
        End If
        For Each child As Control In parent.Controls
            FindControlRecursive(list, child, ctrlType)
        Next
        Return list
    End Function

    'Save control settings and case_x_conditions to file
    Private Sub Save_tofile()
        Dim temp_string As String
        Dim filename As String = "Cool_disk_select_" & TextBox9.Text & "_" & TextBox10.Text & "_" & TextBox11.Text & DateTime.Now.ToString("_yyyy_MM_dd") & ".vtk"
        Dim all_num, all_combo, all_check, all_radio As New List(Of Control)
        Dim i As Integer

        If String.IsNullOrEmpty(TextBox10.Text) Then TextBox10.Text = "-"
        If String.IsNullOrEmpty(TextBox11.Text) Then TextBox11.Text = "-"

        temp_string = TextBox9.Text & ";" & TextBox10.Text & ";" & TextBox11.Text & ";"
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all numeric, combobox, checkbox and radiobutton controls -----------------
        FindControlRecursive(all_num, Me, GetType(NumericUpDown))   'Find the control
        all_num = all_num.OrderBy(Function(x) x.Name).ToList()      'Alphabetical order
        For i = 0 To all_num.Count - 1
            Dim grbx As NumericUpDown = CType(all_num(i), NumericUpDown)
            temp_string &= grbx.Value.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all combobox controls and save
        FindControlRecursive(all_combo, Me, GetType(ComboBox))      'Find the control
        all_combo = all_combo.OrderBy(Function(x) x.Name).ToList()   'Alphabetical order
        For i = 0 To all_combo.Count - 1
            Dim grbx As ComboBox = CType(all_combo(i), ComboBox)
            temp_string &= grbx.SelectedItem.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all checkbox controls and save
        FindControlRecursive(all_check, Me, GetType(CheckBox))      'Find the control
        all_check = all_check.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_check.Count - 1
            Dim grbx As CheckBox = CType(all_check(i), CheckBox)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '-------- find all radio controls and save
        FindControlRecursive(all_radio, Me, GetType(RadioButton))   'Find the control
        all_radio = all_radio.OrderBy(Function(x) x.Name).ToList()  'Alphabetical order
        For i = 0 To all_radio.Count - 1
            Dim grbx As RadioButton = CType(all_radio(i), RadioButton)
            temp_string &= grbx.Checked.ToString & ";"
        Next
        temp_string &= vbCrLf & "BREAK" & vbCrLf & ";"

        '---- if path not exist then create one----------
        Try
            If (Not System.IO.Directory.Exists(dirpath_Home)) Then System.IO.Directory.CreateDirectory(dirpath_Home)
            If (Not System.IO.Directory.Exists(dirpath_Eng)) Then System.IO.Directory.CreateDirectory(dirpath_Eng)
            If (Not System.IO.Directory.Exists(dirpath_Rap)) Then System.IO.Directory.CreateDirectory(dirpath_Rap)
        Catch ex As DirectoryNotFoundException
            MessageBox.Show("Line 1033, " & ex.Message)  ' Show the exception's message.
        End Try

        Try
            If CInt(temp_string.Length.ToString) > 100 Then      'String may be empty
                If Directory.Exists(dirpath_Eng) Then
                    File.WriteAllText(dirpath_Eng & filename, temp_string, Encoding.ASCII)      'used at VTK
                Else
                    File.WriteAllText(dirpath_Home & filename, temp_string, Encoding.ASCII)     'used at home
                End If
            End If
        Catch ex As DirectoryNotFoundException
            MessageBox.Show("Line 1045, " & ex.Message)  ' Show the exception's message.
        End Try
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click, ComboBox5.SelectedIndexChanged
        Troesner_disk()
    End Sub

    Private Sub Troesner_disk()
        'Write Troester Cooling data to input screen 
        Dim tmp As Decimal
        Dim separators() As String = {";"}
        If (ComboBox5.SelectedIndex > -1) Then      'Prevent exceptions
            Dim words() As String = Troester(ComboBox5.SelectedIndex).Split(separators, StringSplitOptions.None)

            'Diameter cooling disk
            Decimal.TryParse(words(1), tmp)
            NumericUpDown9.Value = tmp

            'Thickness cooling disk
            Decimal.TryParse(words(2), tmp)
            NumericUpDown10.Value = tmp

            'Diameter hub
            Decimal.TryParse(words(4), tmp)
            NumericUpDown7.Value = tmp

            ComboBox2.SelectedIndex = 1   'Aluminium-235 
        End If
    End Sub
End Class
