Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop

Public Class Form1
    Dim Senorg, Fecha, Dia, Mes, Year, Nombre, Cedula, FechaInicioDesc, Diai, Mesi, Yeari, DeudaDesc, MontoDesc, Periodo, Porcentaje, LeEntregueDesc, DeudaTotal, MontoTotal As String
    Dim Word1 As Word.Application
    Dim WordDoc As Word.Document
    Dim datosagregados As New DatosRegistro
    Dim consulta As String
    Private _adaptador As New MySqlDataAdapter
    Dim op As Boolean
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim thisDay As DateTime = DateTime.Today
        TextBox6.Text = thisDay.ToString("dd")
        TextBox13.Text = thisDay.ToString("MM")
        TextBox14.Text = thisDay.ToString("yy")
        op = False

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Word1 = New Word.Application
        Word1.Visible = True


        WordDoc = Word1.Documents.Open("C:\Users\JuanB\Desktop\activity6\profile2.dotx")
        Senorg = TextBox1.Text
        Dia = TextBox6.Text
        Mes = TextBox13.Text
        Year = TextBox14.Text
        Fecha = TextBox6.Text + "/" + TextBox13.Text + "/" + TextBox14.Text
        Nombre = TextBox7.Text
        Cedula = TextBox2.Text
        FechaInicioDesc = (TextBox3.Text + "/" + TextBox15.Text + "/" + TextBox16.Text)
        Diai = TextBox3.Text
        Mesi = TextBox15.Text
        Yeari = TextBox16.Text
        DeudaDesc = TextBox4.Text
        MontoDesc = TextBox5.Text
        Periodo = TextBox8.Text
        Porcentaje = TextBox9.Text
        LeEntregueDesc = TextBox10.Text
        DeudaTotal = TextBox11.Text
        MontoTotal = TextBox12.Text
        With WordDoc
            .FormFields("Srgerente").Result = Senorg
            .FormFields("Dia").Result = Dia
            .FormFields("Mes").Result = Mes
            .FormFields("Year").Result = Year
            .FormFields("Nombre").Result = Nombre
            .FormFields("Cedula").Result = Cedula
            .FormFields("Diai").Result = Diai
            .FormFields("Mesi").Result = Mesi
            .FormFields("Yeari").Result = Yeari
            .FormFields("Deudai").Result = DeudaDesc
            .FormFields("Dolaresi").Result = MontoDesc

            .FormFields("Periodo").Result = Periodo
            .FormFields("Porcentaje").Result = Porcentaje
            .FormFields("Leentregue").Result = LeEntregueDesc

            .FormFields("DeudaT").Result = DeudaTotal
            .FormFields("DolaresT").Result = MontoTotal

            .FormFields("Fecha").Result = Fecha

        End With

        Word1 = Nothing
        WordDoc = Nothing

        Conexion_Global()
        consulta = ("CALL Insertar('" & Cedula & "','" & Senorg & "','" & Nombre & "','" & Fecha & "','" & FechaInicioDesc & "','" & DeudaDesc & "','" & MontoDesc & "','" & Periodo & "','" & Porcentaje & "','" & DeudaTotal & "','" & MontoTotal & "','" & LeEntregueDesc & "');")
        _adaptador.InsertCommand = New MySqlCommand(consulta, _coneccion)
        _coneccion.Open()
        _adaptador.InsertCommand.Connection = _coneccion
        _adaptador.InsertCommand.ExecuteNonQuery()

        MsgBox("Recuerda siempre guardar y cerrar el archivo anterior en Word")


    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Senorg = TextBox1.Text
        Dia = TextBox6.Text
        Mes = TextBox13.Text
        Year = TextBox14.Text
        Fecha = TextBox6.Text + "/" + TextBox13.Text + "/" + TextBox14.Text
        Nombre = TextBox7.Text
        Cedula = TextBox2.Text
        FechaInicioDesc = (TextBox3.Text + "/" + TextBox15.Text + "/" + TextBox16.Text)
        Diai = TextBox3.Text
        Mesi = TextBox15.Text
        Yeari = TextBox16.Text
        DeudaDesc = TextBox4.Text
        MontoDesc = TextBox5.Text
        Periodo = TextBox8.Text
        Porcentaje = TextBox9.Text
        LeEntregueDesc = TextBox10.Text
        DeudaTotal = TextBox11.Text
        MontoTotal = TextBox12.Text

        MessageBox.Show("                   " + Cedula + vbNewLine + vbNewLine + Senorg + "    " + Fecha + vbNewLine + vbNewLine + Nombre + "          " + Cedula + vbNewLine + vbNewLine + FechaInicioDesc + "  " +
                        DeudaDesc + "   " + MontoDesc + vbNewLine + vbNewLine + Periodo + "      " + Porcentaje + "     " + LeEntregueDesc + vbNewLine + vbNewLine + DeudaTotal + " " + MontoTotal)



        MessageBox.Show("                                                      No. Cedula" + Cedula + vbNewLine + vbNewLine + " Senor Gerente de:" + Senorg + "    Panama," + Fecha + vbNewLine + "Presente,-" + vbNewLine + "Senor:" + vbNewLine + " Yo " + Nombre + "     con Cedula N# " + Cedula + " por medio de la presente" + vbNewLine + "autorizo a usted para que me descuente de mi salario total(Incluye sueldo base, comisiones, aumentos, etc.," + vbNewLine + "comenzando el " + FechaInicioDesc + "  la suma de " + DeudaDesc + ".  " + vbNewLine + "(B/." + MontoDesc + ") todas las" + Periodo + " que representan un" + Porcentaje + " por ciento de mi" + vbNewLine + "salario actual,y le entregue estos descuentos mensualmente a :" + LeEntregueDesc + vbNewLine + " para cubrir obligacion de Credito comercial que he contraido con este establecimiento." + vbNewLine + vbNewLine + "Esta autorizacion es de caracter irrevocable y esta vigente hasta cancelar la suma de" + DeudaTotal + vbNewLine + " (B/." + MontoTotal + "), y puede ser descontada en el nuevo empleo en caso de que ocurra cambio de trabajo, por" + vbNewLine + "Esta autorizacio de descuento inclute el descuento mensual de las vacaciones y de descuento correspondiente" + vbNewLine +
                        "...Para las formalidades del caso de agradezco a usted devuelva copia de esta autorizacion firmada al beneficiario."
                        )

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        TextBox1.Text = ""
        TextBox7.Text = ""
        TextBox2.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox10.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox3.Text = ""
        TextBox15.Text = ""
        TextBox16.Text = ""
    End Sub

    Private Sub Label11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label11.Click

    End Sub
End Class
