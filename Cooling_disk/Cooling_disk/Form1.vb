Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, MyBase.Load
        Calc_shaft()
    End Sub

    Private Sub Calc_shaft()
        Dim f_od, f_id, F_length, F_coeff, F_temp, Shaft_area As Double
        Dim d_no, d_od, d_hub_od, d_coeff, d_thick, d_temp, d_area_actual, d_area_calc, area_factor As Double

        '-------------- shaft-----------------
        f_od = NumericUpDown1.Value
        f_id = NumericUpDown2.Value
        F_length = NumericUpDown3.Value
        F_coeff = NumericUpDown4.Value
        F_temp = NumericUpDown5.Value
        Shaft_area = Math.PI / 4 * (f_od ^ 2 - f_id ^ 2)

        TextBox1.Text = Math.Round(Shaft_area, 0).ToString


        '-------------- disk-----------------
        d_no = NumericUpDown11.Value
        d_od = NumericUpDown9.Value
        d_hub_od = NumericUpDown7.Value
        d_thick = NumericUpDown10.Value
        d_coeff = NumericUpDown6.Value
        d_area_actual = d_no * Math.PI / 4 * (d_od ^ 2 - d_hub_od ^ 2)

        area_factor = 0.5

        d_area_calc = d_area_actual * area_factor

        TextBox2.Text = Math.Round(d_area_actual, 0).ToString
        TextBox3.Text = Math.Round(area_factor, 2).ToString
        TextBox4.Text = Math.Round(d_area_calc, 0).ToString

    End Sub

End Class
