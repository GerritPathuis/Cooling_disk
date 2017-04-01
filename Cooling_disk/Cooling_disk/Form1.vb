Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, MyBase.Load, NumericUpDown6.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged
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
        temp_disk = NumericUpDown13.Value

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
            For i = 0 To 5
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

            NumericUpDown13.Value = temp_disk
        End If
        If power_conducted < 0 Then power_conducted = 0
        If power_transferred < 0 Then power_transferred = 0

        TextBox1.Text = Math.Round(Shaft_area, 2).ToString
        TextBox2.Text = Math.Round(d_area_actual, 2).ToString
        TextBox3.Text = Math.Round(area_eff, 3).ToString
        TextBox4.Text = Math.Round(d_area_calc, 2).ToString

        TextBox5.Text = Math.Round(power_conducted, 0).ToString
        TextBox6.Text = Math.Round(power_transferred, 0).ToString

    End Sub


End Class
