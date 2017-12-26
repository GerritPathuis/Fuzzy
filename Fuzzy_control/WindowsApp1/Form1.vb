Imports System.Net.NetworkInformation
Imports System.Net.Sockets
Imports System.Text
Imports System.IO
Imports System.IO.Ports
Imports System.ComponentModel
Imports System.Threading
Imports System.Drawing.Graphics
Imports System.Math
Imports System.Drawing
Imports System.Drawing.Image
Imports System.Windows.Forms
Imports System.Drawing.Printing

Public Class Form1
    'Based on Youtube, H462710 - Fuzzy Logic Control Example
    Public r_Bitmap As Bitmap
    Public q_Graphics As Graphics

    Public output, input, rc As Double
    Public NN(5) As Double  'Input set Negative Negative
    Public N(5) As Double   'Input set Negative
    Public Z(5) As Double   'Input set Zero
    Public P(5) As Double   'Input set Positive
    Public PP(5) As Double  'Input set Positive Positive
    Public Input_Range As Double
    Public Output_Range As Double

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, NumericUpDown9.ValueChanged, NumericUpDown8.ValueChanged, NumericUpDown7.ValueChanged, NumericUpDown6.ValueChanged, NumericUpDown5.ValueChanged, NumericUpDown4.ValueChanged, NumericUpDown3.ValueChanged, NumericUpDown2.ValueChanged, NumericUpDown16.ValueChanged, NumericUpDown15.ValueChanged, NumericUpDown14.ValueChanged, NumericUpDown13.ValueChanged, NumericUpDown12.ValueChanged, NumericUpDown11.ValueChanged, NumericUpDown10.ValueChanged, NumericUpDown1.ValueChanged, NumericUpDown31.ValueChanged, NumericUpDown30.ValueChanged, NumericUpDown29.ValueChanged, NumericUpDown28.ValueChanged, NumericUpDown20.ValueChanged, CheckBox5.CheckedChanged, CheckBox4.CheckedChanged, CheckBox3.CheckedChanged, CheckBox2.CheckedChanged, CheckBox1.CheckedChanged
        Input_Range = 20    'Volt (to do)
        Output_Range = 60   'Volt (to do)
        TextBox18.Text = Input_Range.ToString
        TextBox19.Text = Output_Range.ToString

        '===== Get Crisp input value =======
        input = NumericUpDown1.Value

        If CheckBox1.Checked Then   '===== NN (Negative-Negative) set =======
            NN(0) = NumericUpDown2.Value
            NN(1) = NumericUpDown8.Value
            NN(2) = NumericUpDown12.Value
            NN(4) = NumericUpDown30.Value.ToString
            If (input > NN(1) And input < NN(2)) Then
                NN(3) = Calc_DOM(input, NN(0), NN(1), NN(2))
            Else
                NN(3) = 0     'Degree of Memerbership
            End If
            'MessageBox.Show("Input=" & input.ToString & ", 0= " & NN(0).ToString & ", 1= " & NN(1).ToString & ", 2=" & NN(2).ToString & ",DOM=" & NN(3).ToString)
        End If
        TextBox2.Text = NN(3).ToString  'Input DOM

        If CheckBox2.Checked Then   '===== N set =======
            N(0) = NumericUpDown3.Value
            N(1) = NumericUpDown9.Value
            N(2) = NumericUpDown13.Value
            N(4) = NumericUpDown31.Value.ToString
            If (input > N(0) And input < N(2)) Then
                N(3) = Calc_DOM(input, N(0), N(1), N(2))
            Else
                N(3) = 0
            End If
        End If
        TextBox3.Text = N(3).ToString  'Input DOM

        If CheckBox3.Checked Then '===== Z set=======
            Z(0) = NumericUpDown4.Value
            Z(1) = NumericUpDown7.Value
            Z(2) = NumericUpDown11.Value
            Z(4) = NumericUpDown29.Value.ToString
            If (input > Z(0) And input < Z(2)) Then
                Z(3) = Calc_DOM(input, Z(0), Z(1), Z(2))
            Else
                Z(3) = 0
            End If
        End If
        TextBox4.Text = Z(3).ToString  'Input DOM

        If CheckBox4.Checked Then '===== P set=======
            P(0) = NumericUpDown5.Value
            P(1) = NumericUpDown6.Value
            P(2) = NumericUpDown10.Value
            P(4) = NumericUpDown28.Value.ToString
            If (input > P(0) And input < P(2)) Then
                P(3) = Calc_DOM(input, P(0), P(1), P(2))
            Else
                P(3) = 0
            End If
        End If
        TextBox5.Text = P(3).ToString  'Input DOM

        If CheckBox5.Checked Then '===== PP set=======
            PP(0) = NumericUpDown5.Value
            PP(1) = NumericUpDown6.Value
            PP(2) = NumericUpDown10.Value
            PP(4) = NumericUpDown20.Value.ToString
            If (input > PP(0) And input < PP(2)) Then
                PP(3) = Calc_DOM(input, PP(0), PP(1), PP(2))
            Else
                PP(3) = 0
            End If
        End If
        TextBox6.Text = PP(3).ToString  'Input DOM

        '----- Set output value -------
        TextBox7.Text = NN(4).ToString
        TextBox8.Text = N(4).ToString
        TextBox9.Text = Z(4).ToString
        TextBox10.Text = P(4).ToString
        TextBox11.Text = PP(4).ToString

        '----- Calc Crisp value ---
        NN(5) = NN(3) * NN(4)
        N(5) = N(3) * N(4)
        Z(5) = Z(3) * Z(4)
        P(5) = P(3) * P(4)
        PP(5) = PP(3) * PP(4)

        TextBox12.Text = NN(5).ToString
        TextBox13.Text = N(5).ToString
        TextBox14.Text = Z(5).ToString
        TextBox15.Text = P(5).ToString
        TextBox16.Text = PP(5).ToString
        output = NN(5) + N(5) + Z(5) + P(5) + PP(5)
        TextBox17.Text = output.ToString

        MakeNewBitmap2()    'Chear the graph
        drwg_input_set()
    End Sub

    Private Function Calc_DOM(inputt As Double, p0 As Double, p1 As Double, p2 As Double) As Double
        'Calculate the Degree Of Memebership
        Dim rc, dom As Double

        If p1 = p2 Then MessageBox.Show("Problem with RC calc")
        rc = 1 / (p1 - p2)
        dom = (inputt - p1) * rc
        Return (dom)
    End Function

    Private Sub MakeNewBitmap2() 'Balancer picture
        Dim wid As Integer = PictureBox2.Size.Width
        Dim hgt As Integer = PictureBox2.Size.Height
        r_Bitmap = New Bitmap(wid, hgt)
        q_Graphics = Graphics.FromImage(r_Bitmap)
        q_Graphics.Clear(Me.BackColor)
        PictureBox2.Image = r_Bitmap
    End Sub

    Private Sub drwg_input_set()
        Dim breed As Integer
        Dim hoog As Integer
        Dim x0, x1, x2, x3, x4, x_offset, left_margin As Integer
        Dim y0, y1, y2, y3, y4 As Integer
        Dim DOM0_height As Integer  'DOM 0 line
        Dim DOM1_height As Integer  'DOM 1 line
        Dim h_scale As Double       'horizonral scale
        Dim pen1 As New System.Drawing.Pen(Color.Yellow, 2) 'Used in the graph
        Dim pen2 As New System.Drawing.Pen(Color.White, 2) 'Used in the graph
        Dim pen3 As New System.Drawing.Pen(Color.Green, 2) 'Used in the graph
        Dim pen4 As New System.Drawing.Pen(Color.Blue, 2) 'Used in the graph
        Dim pen5 As New System.Drawing.Pen(Color.Red, 2) 'Used in the graph

        hoog = PictureBox2.Size.Height - 10
        breed = PictureBox2.Size.Width - 10

        DOM0_height = 20
        DOM1_height = PictureBox2.Size.Height - 20
        x_offset = 10
        left_margin = 50

        h_scale = (PictureBox2.Size.Width - left_margin - x_offset) / Input_Range

        Try
            q_Graphics.Clear(Color.Black)  'Clear graph

            'Draw DOM=0 and DOM=1 
            q_Graphics.DrawLine(pen2, left_margin, DOM0_height, breed, DOM0_height)
            q_Graphics.DrawLine(pen2, left_margin, DOM1_height, breed, DOM1_height)

            'Draw DOM lines
            Dim drawFont As New Font("Arial", 8)
            Dim drawBrush As New SolidBrush(Color.White)
            Dim drawPoint1 As New PointF(0, DOM0_height - 5)
            Dim drawPoint2 As New PointF(0, (DOM1_height - 5))
            q_Graphics.DrawString("DOM=1", drawFont, drawBrush, drawPoint1)
            q_Graphics.DrawString("DOM=0", drawFont, drawBrush, drawPoint2)

            If CheckBox1.Checked Then 'Draw NN_input_set 
                x0 = left_margin + (NN(0) + x_offset) * h_scale
                x1 = left_margin + (NN(1) + x_offset) * h_scale
                x2 = left_margin + (NN(2) + x_offset) * h_scale
                If (NN(0) <> NN(1)) Then q_Graphics.DrawLine(pen1, x0, DOM1_height, x1, DOM0_height)
                If (NN(1) <> NN(2)) Then q_Graphics.DrawLine(pen1, x1, DOM0_height, x2, DOM1_height)
            End If

                If CheckBox2.Checked Then 'Draw N_input_set 
                x0 = left_margin + (N(0) + x_offset) * h_scale
                x1 = left_margin + (N(1) + x_offset) * h_scale
                x2 = left_margin + (N(2) + x_offset) * h_scale
                If (N(0) <> N(1)) Then q_Graphics.DrawLine(pen2, x0, DOM1_height, x1, DOM0_height)
                If (N(1) <> N(2)) Then q_Graphics.DrawLine(pen2, x1, DOM0_height, x2, DOM1_height)
            End If

            If CheckBox3.Checked Then 'Draw Z_input_set 
                x0 = left_margin + (Z(0) + x_offset) * h_scale
                x1 = left_margin + (Z(1) + x_offset) * h_scale
                x2 = left_margin + (Z(2) + x_offset) * h_scale
                If (Z(0) <> Z(1)) Then q_Graphics.DrawLine(pen3, x0, DOM1_height, x1, DOM0_height)
                If (Z(1) <> Z(2)) Then q_Graphics.DrawLine(pen3, x1, DOM0_height, x2, DOM1_height)
            End If

            If CheckBox4.Checked Then 'Draw P_input_set 
                x0 = left_margin + (P(0) + x_offset) * h_scale
                x1 = left_margin + (P(1) + x_offset) * h_scale
                x2 = left_margin + (P(2) + x_offset) * h_scale
                If (P(0) <> P(1)) Then q_Graphics.DrawLine(pen4, x0, DOM1_height, x1, DOM0_height)
                If (P(1) <> P(2)) Then q_Graphics.DrawLine(pen4, x1, DOM0_height, x2, DOM1_height)
            End If

            If CheckBox5.Checked Then 'Draw P_input_set 
                x0 = left_margin + (PP(0) + x_offset) * h_scale
                x1 = left_margin + (PP(1) + x_offset) * h_scale
                x2 = left_margin + (PP(2) + x_offset) * h_scale
                If (PP(0) <> PP(1)) Then q_Graphics.DrawLine(pen5, x0, DOM1_height, x1, DOM0_height)
                If (PP(1) <> PP(2)) Then q_Graphics.DrawLine(pen5, x1, DOM0_height, x2, DOM1_height)
            End If

        Catch ex As Exception
            MsgBox("Error 105 writing problem" & ex.Message)
        End Try
    End Sub


End Class
