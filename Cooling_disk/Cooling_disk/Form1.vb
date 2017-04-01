Imports System
Imports System.IO
Imports System.Text
Imports System.Math
'Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Globalization
Imports System.Threading
'Imports Word = Microsoft.Office.Interop.Word
Imports System.Collections


Public Class Form1

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
            TextBox8.Text &= transfer(hh)
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


End Class
