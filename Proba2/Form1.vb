Imports System.IO
Imports iTextSharp.text

Public Class Form1

    Dim Alumnos As List(Of Alumno) = New List(Of Alumno)

    Dim Alumnos_out As List(Of Alumno) = New List(Of Alumno)
    Dim extrI, extrM, extrD As Boolean


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Label1.Show()

        Dim codAlum As String = " "

        Dim extrOk, extrI, extrM, extrD, existD, existM As Boolean


        extrOk = True




        If File.Exists(OpenFileDialog1.FileName) Then

            ' ************************* LECTURA FICHERO SGA ***********************************

            extrI = ExtraerFichero(Alumnos, 1, OpenFileDialog1.FileName)


            ' ************************* LECTURA FICHERO MATRICULA ***********************************


            If (extrI = False) Then

                MessageBox.Show("Es imprescindible haber extraído previamente los expedientes de SGA", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Else
                If (File.Exists(OpenFileDialog2.FileName)) Then

                    asignarHZAlumno(Alumnos, OpenFileDialog2.FileName)
                    extrM = ExtraerFichero(Alumnos, 2, OpenFileDialog2.FileName)
                    If (extrM = False) Then
                        Exit Sub
                    End If
                Else
                    existM = False
                End If

                If (File.Exists(OpenFileDialog3.FileName)) Then

                    ' ************************* LECTURA FICHERO DAE ***********************************

                    extrD = ExtraerFichero(Alumnos, 3, OpenFileDialog3.FileName)

                    If (extrD = False) Then
                        Exit Sub
                    End If
                Else
                    existD = False
                End If


            End If

        Else

            MessageBox.Show("Es obligatorio extraer los datos del SGA", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            extrI = False


        End If


        If (extrI = True) And (extrM = True Or existM = False) And (extrD = True Or existD = False) Then
            Button3.Enabled = True
            MessageBox.Show("Extracción realizada correctamente", "OK", MessageBoxButtons.OK)

        End If


    End Sub

    Private Function ExtraerFichero(ByVal Alumnos As List(Of Alumno), ByVal tipoFich As Integer, ByVal ubicFich As String) As Boolean

        Dim objStreamReader As StreamReader
        Dim strLinea, strLineaCab As String
        Dim aDatos() As String

        Dim codAlum As String = " "
        Dim asignatura As Asignatura

        Dim extr As Boolean
        Dim comparacion1 As String
        Dim aape1, aape2, anom, aens, acurs As Integer
        Dim shz, sdesc, sdesk As Integer

        Dim i As Integer


        objStreamReader = New StreamReader(ubicFich)

        'Se establece la posición de los datos para cada tipo de fichero
        ' TipoFichero = SGA
        If (tipoFich = 1) Then

            comparacion1 = "hz_Kodea"
            i = 0
            aape1 = 1
            aape2 = 2
            anom = 3
            aens = 4
            acurs = 5
            shz = 6
            sdesc = 7
            sdesk = 8
            'sens = 9
            'scurs = 10

            'TipoFichero = Matrícula
        ElseIf (tipoFich = 2) Then
            comparacion1 = "COD_ALU"
            i = 0
            aape1 = 5
            aape2 = 4
            anom = 6
            aens = 22
            acurs = 23
            shz = 40
            sdesc = 41
            'sens
            'scurs

            'TipoFichero = DAE
        Else
            comparacion1 = "Alumno"
            i = 6
            shz = 7
            sdesc = 8
            sdesk = 9

        End If

        ' ************************* LECTURA FICHERO ***********************************
        Dim alumno = New Alumno()

        'Lectura cabecera fichero
        strLineaCab = objStreamReader.ReadLine

        aDatos = Split(strLineaCab, ";")

        'Controlar el tipo de fichero
        If (aDatos(i) = comparacion1) Then

            strLinea = objStreamReader.ReadLine


            'Continuar leyendo hasta el final del fichero
            Do While Not strLinea Is Nothing



                aDatos = Split(strLinea, ";")

                If (tipoFich = 2 And aDatos(aens) <> "Bachillerato") Then
                    RadioButton1.Enabled = False
                    RadioButton3.Enabled = False
                    RadioButton4.Enabled = False
                End If

                If (tipoFich = 1) Or (tipoFich = 2) Or (tipoFich = 3 And aDatos(0) <> "") Then


                    ' Si es un alumno todavía no tratado en el fichero

                    If (codAlum <> aDatos(i)) Then
                        If (tipoFich = 1) Then
                            alumno = New Alumno(aDatos(i), aDatos(aape1), aDatos(aape2), aDatos(anom), aDatos(aens), CInt(aDatos(acurs)))
                            Alumnos.Add(alumno)
                            alumno.SGA = New List(Of Asignatura)

                            'asignatura = New Asignatura(aDatos(shz), aDatos(sdesc), aDatos(sdesk))

                        Else

                            alumno = ObtenerAlumno(Alumnos, tipoFich, aDatos(i))

                            If (alumno Is Nothing) Then
                                alumno = New Alumno()
                                alumno.NombreCompleto = aDatos(i)
                                Alumnos.Add(alumno)
                            End If

                            If tipoFich = 2 Then
                                alumno.Matricula = New List(Of Asignatura)

                                'Añadir a la lista de matrícula la Lengua Extranjera
                                asignatura = New Asignatura(aDatos(31), aDatos(32), aDatos(32))
                                alumno.Matricula.Add(asignatura)

                                ' Añadir a la lista de matrícula Religión (en caso de que se imparta)
                                If (aDatos(33) = "SI") Then
                                    asignatura = New Asignatura(aDatos(34), aDatos(35), aDatos(35))
                                    alumno.Matricula.Add(asignatura)
                                End If

                            Else
                                alumno.DAE = New List(Of Asignatura)
                            End If

                        End If

                        codAlum = aDatos.GetValue(i)

                    End If


                    'Añadir asignatura de la línea correspondiente en la lista 
                    asignatura = New Asignatura(aDatos(shz), aDatos(sdesc), aDatos(sdesk))

                    If (tipoFich = 1) Then
                        alumno.SGA.Add(asignatura)
                    ElseIf tipoFich = 2 Then

                        alumno.Apellido1 = aDatos(aape1)
                        alumno.Apellido2 = aDatos(aape2)
                        alumno.Nombre = aDatos(anom)

                        alumno.Matricula.Add(asignatura)
                    Else
                        alumno.DAE.Add(asignatura)
                    End If

                End If
                strLinea = objStreamReader.ReadLine

            Loop

            'Cerrar fichero
            objStreamReader.Close()

            extr = True
        Else

            MessageBox.Show("Fichero incorrecto", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

            extr = False 'Fichero incorrecto
        End If

        Return extr

    End Function


    Public Function ObtenerAlumno(ByVal Alumnos As List(Of Alumno), ByVal TipoFich As Integer, ByVal Dato As String) As Alumno

        ' Comparación mediante el código HZ del alumno (1)

        If (TipoFich = 2) Then
            For Each alum As Alumno In Alumnos
                If (alum.CodigoHZ.Equals(Dato)) Then
                    Return alum
                    Exit For
                End If
            Next

            ' Comparación mediante el nombre completo del alumno (2)
        ElseIf (TipoFich = 3) Then
            For Each alum As Alumno In Alumnos
                If (alum.NombreCompleto.Equals(Dato)) Then
                    Return alum
                    Exit For
                End If
            Next
        End If
        Return Nothing

    End Function

    Private Sub asignarHZAlumno(ByVal Lista1 As List(Of Alumno), ByVal Fichero As String)
        For Each alum As Alumno In Lista1

            If (alum.CodigoHZ = "") Then
                alum.CodigoHZ = buscarHZAlumno(alum, Fichero)
            End If
        Next
    End Sub

    Private Function buscarHZAlumno(ByVal alumno As Alumno, ByVal Fichero As String) As String
        Dim objStreamReader As StreamReader
        Dim strlinea As String
        Dim codalum As String = ""
        Dim aDatos() As String
        objStreamReader = New StreamReader(Fichero)

        'Leer cabecera
        strlinea = objStreamReader.ReadLine

        strlinea = objStreamReader.ReadLine


        Do While Not strlinea Is Nothing
            aDatos = Split(strlinea, ";")

            'Controlar si es un alumno nuevo en el fichero (primera línea del alumno)
            If (codalum <> aDatos(0)) Then

                'Si el nombre del alumno pasado por parámetro coinciden con los datos del fichero
                If (aDatos(4) = alumno.Apellido2 And aDatos(5) = alumno.Apellido1 And aDatos(6) = alumno.Nombre) Then
                    Return aDatos(0)
                End If
            End If
            codalum = aDatos.GetValue(0)
        Loop
        Return ""



    End Function


    Private Sub Label1_Click(sender As Object, e As EventArgs)
        Me.Hide()

    End Sub

    

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim comp1, comp2, comp3, comp4, comp5, comp6 As Boolean

        'MessageBox.Show("Fichero procesado correctamente", "OK", MessageBoxButtons.OK)

        Alumnos_out.Clear()

        ' Comparación "SGA vs Matrícula"
        If (RadioButton1.Checked = True) Then

            For Each Alumno As Alumno In Alumnos
                inicializarEncontrados(Alumno.SGA, Alumno.Matricula, Alumno.DAE)
                comp1 = comparacion(1, Alumno, Alumno.SGA, Alumno.Matricula)
            Next

            'Comparación "DAE vs SGA"
        ElseIf (RadioButton2.Checked = True) Then
            For Each Alumno As Alumno In Alumnos
                inicializarEncontrados(Alumno.SGA, Alumno.DAE, Alumno.Matricula)
                comp2 = comparacion(2, Alumno, Alumno.SGA, Alumno.DAE)
            Next



            'Comparación "Matrícula vs DAE"
        ElseIf (RadioButton3.Checked = True) Then

            For Each Alumno As Alumno In Alumnos
                inicializarEncontrados(Alumno.Matricula, Alumno.DAE, Alumno.SGA)
                comp3 = comparacion(3, Alumno, Alumno.Matricula, Alumno.DAE)
            Next

            'TODOS
        ElseIf (RadioButton4.Checked = True) Then

            For Each Alumno As Alumno In Alumnos
                inicializarEncontrados(Alumno.Matricula, Alumno.DAE, Alumno.SGA)
                comp4 = comparacion(1, Alumno, Alumno.SGA, Alumno.Matricula)
                inicializarEncontrados(Alumno.Matricula, Alumno.DAE, Alumno.SGA)
                comp5 = comparacion(2, Alumno, Alumno.SGA, Alumno.DAE)
                inicializarEncontrados(Alumno.Matricula, Alumno.DAE, Alumno.SGA)
                comp6 = comparacion(3, Alumno, Alumno.Matricula, Alumno.DAE)
            Next

            'No se ha seleccionado ningun tipo de comparación
        ElseIf (RadioButton1.Checked = False) And (RadioButton2.Checked = False) And (RadioButton3.Checked = False) And (RadioButton4.Checked = False) Then
            MessageBox.Show("Es preciso seleccionar un tipo de comparación", "Select null", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            MessageBox.Show("Error", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If

        ImpresionPDF()
    End Sub

    Private Function comparacion(ByVal tipocomp As Integer, ByVal Alumno As Alumno, ByVal Lista1 As List(Of Asignatura), ByVal Lista2 As List(Of Asignatura)) As Boolean
        Dim al1 As Boolean
        Dim a, a1, a2, a3, a4 As Asignatura
        Dim alumno_out As Alumno

        ' Variable para controlar si el alumno existía en la lista, o se ha tenido que crear
        al1 = False

        'Buscar el alumno recibido como parámetro en la lista alumno_out según el códigoHZ
        alumno_out = ObtenerAlumno(Alumnos_out, 2, Alumno.CodigoHZ)

        'Si no se ha encontrado ningún alumno con dicho códigoHZ, realizar la búsqueda según nombre completo
        If (alumno_out Is Nothing) Then
            alumno_out = ObtenerAlumno(Alumnos_out, 1, Alumno.NombreCompleto)
        End If

        ' Si el alumno no existe en la lista alumno_out
        If (alumno_out Is Nothing) Then
            alumno_out = New Alumno(Alumno.CodigoHZ, Alumno.Apellido1, Alumno.Apellido2, Alumno.Nombre, Alumno.Enseñanza, Alumno.Curso)
            al1 = True
        End If

        If ((tipocomp = 1) And (Lista2 Is Nothing)) Or ((tipocomp = 3) And (Lista1 Is Nothing)) Then
            If (alumno_out.Matricula Is Nothing) Then
                alumno_out.Matricula = New List(Of Asignatura)
                Dim asignatura_out = New Asignatura("XXXX", "Alumno no matriculado", "Ikasle ez matrikulatua")
                alumno_out.Matricula.Add(asignatura_out)

            End If
        ElseIf (tipocomp = 2) And (Lista2 Is Nothing) Then
            alumno_out.DAE = New List(Of Asignatura)
            Dim asignatura_out = New Asignatura("XXXX", "Alumno no contemplado en el DAE", "DAEn esleitu gabeko ikaslea")
            alumno_out.DAE.Add(asignatura_out)

        ElseIf (((tipocomp = 2) Or (tipocomp = 1)) And (Lista1 Is Nothing)) Then
            alumno_out.SGA = New List(Of Asignatura)
            Dim asignatura_out = New Asignatura("XXXX", "Alumno no registrado en el SGA", "KASen erregistratu gabeko ikaslea")
            alumno_out.SGA.Add(asignatura_out)

        Else


            For Each asigI As Asignatura In Lista1

                'Buscar en la Lista2 las asignaturas de la Lista1
                a = obtenerAsignatura(Lista2, asigI.CodigoHZAsig)

                ' Si una asignatura de la Lista1 no existe en la Lista2, la añadiremos directamente a la Lista2 de alumnos_out
                If (a Is Nothing) Then

                    Dim asignatura_out = New Asignatura(asigI.CodigoHZAsig, asigI.Descripcion, asigI.Deskribapena)

                    If (tipocomp = 1) Then
                        If (alumno_out.Matricula Is Nothing) Then
                            alumno_out.Matricula = New List(Of Asignatura)
                            alumno_out.Matricula.Add(asignatura_out)
                        Else
                            a1 = obtenerAsignatura(alumno_out.Matricula, asigI.CodigoHZAsig)
                            If (a1 Is Nothing) Then
                                alumno_out.Matricula.Add(asignatura_out)
                            End If
                        End If



                    ElseIf (tipocomp = 2) Or (tipocomp = 3) Then
                        If (alumno_out.DAE Is Nothing) Then
                            alumno_out.DAE = New List(Of Asignatura)
                            alumno_out.DAE.Add(asignatura_out)
                        Else
                            a2 = obtenerAsignatura(alumno_out.DAE, asigI.CodigoHZAsig)
                            If (a2 Is Nothing) Then
                                alumno_out.DAE.Add(asignatura_out)
                            End If

                        End If
                    End If

                    ' Si una asignatura de la Lista1 existe en la Lista2, la añadiremos directamente a la lista alumnos_out
                Else
                    a.Enc = True
                    asigI.Enc = True

                End If



            Next


            ' Recorrer Lista2 para identificar aquellas asignaturas que no se han emparejado con ninguna de las de la Lista1 e introducirlas en la Lista1 de alumnos_out
            For Each asigM As Asignatura In Lista2
                If (asigM.Enc = False) Then
                    Dim asignatura_out = New Asignatura(asigM.CodigoHZAsig, asigM.Descripcion, asigM.Deskribapena)



                    If (tipocomp = 1) Or (tipocomp = 2) Then
                        If (alumno_out.SGA Is Nothing) Then
                            alumno_out.SGA = New List(Of Asignatura)
                            alumno_out.SGA.Add(asignatura_out)
                        Else
                            a3 = obtenerAsignatura(alumno_out.SGA, asigM.CodigoHZAsig)
                            If (a3 Is Nothing) Then
                                alumno_out.SGA.Add(asignatura_out)
                            End If
                        End If


                    Else
                        If (alumno_out.Matricula Is Nothing) Then
                            alumno_out.Matricula = New List(Of Asignatura)
                            alumno_out.Matricula.Add(asignatura_out)
                        Else
                            a4 = obtenerAsignatura(alumno_out.Matricula, asigM.CodigoHZAsig)
                            If (a4 Is Nothing) Then
                                alumno_out.Matricula.Add(asignatura_out)
                            End If

                        End If

                    End If


                End If
            Next


        End If

        If (al1 = True) Then
            Alumnos_out.Add(alumno_out)
        End If

        Return True

    End Function

    Private Function obtenerAsignatura(ByVal Asignaturas As List(Of Asignatura), ByVal HZcodAsig As String) As Asignatura

        For Each asig As Asignatura In Asignaturas
            If (asig.CodigoHZAsig.Equals(HZcodAsig)) Then
                Return asig
                Exit For
            End If
        Next

        Return Nothing
    End Function

    Private Sub inicializarEncontrados(ByVal Lista1 As List(Of Asignatura), ByVal Lista2 As List(Of Asignatura), ByVal Lista3 As List(Of Asignatura))
        If (Lista1 Is Nothing) Then
            'Do nothing
        Else
            For Each asig As Asignatura In Lista1
                asig.Enc = False
            Next
        End If

        If (Lista2 Is Nothing) Then
            'Do nothing
        Else
            For Each asig As Asignatura In Lista2
                asig.Enc = False
            Next
        End If

        If (Lista3 Is Nothing) Then
            'Do nothing
        Else
            For Each asig As Asignatura In Lista3
                asig.Enc = False
            Next
        End If

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        TextBox1.Text = OpenFileDialog1.FileName
    End Sub


    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        OpenFileDialog2.ShowDialog()
    End Sub

    Private Sub OpenFileDialog2_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog2.FileOk
        TextBox2.Text = OpenFileDialog2.FileName
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        OpenFileDialog3.ShowDialog()
    End Sub

    Private Sub OpenFileDialog3_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog3.FileOk
        TextBox3.Text = OpenFileDialog3.FileName
    End Sub



    Private Sub ImpresionPDF()

        ' Especificar destino fichero
        Dim SaveFileDialog As New SaveFileDialog
        Dim ruta As String

        With SaveFileDialog
            .Title = "Guardar"
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
            .Filter = "Archivos pdf (*.pdf)|*.pdf"
            .FileName = "Archivo"
            .OverwritePrompt = True
            .CheckPathExists = True

        End With


        If SaveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            ruta = SaveFileDialog.FileName
        Else
            ruta = String.Empty
            Exit Sub

        End If

        Try

            Dim document As New iTextSharp.text.Document(PageSize.A4)

            'Configurar documento
            document.PageSize.Rotate()
            document.AddAuthor("SGA")
            document.AddTitle("Resultado comparación")

            Dim writer As pdf.PdfWriter = pdf.PdfWriter.GetInstance(document, New System.IO.FileStream(ruta, System.IO.FileMode.Create))

            writer.ViewerPreferences = pdf.PdfWriter.PageLayoutSinglePage

            document.Open()


            Dim cb As pdf.PdfContentByte = writer.DirectContent
            Dim bf As pdf.BaseFont = pdf.BaseFont.CreateFont(pdf.BaseFont.HELVETICA, pdf.BaseFont.CP1252, pdf.BaseFont.NOT_EMBEDDED)


            cb.SetFontAndSize(bf, 10)
            cb.BeginText()

            cb.SetTextMatrix(50, 800)

            'Recorrer la lista alumnos_out (resultado de la comparación) para imprimirla
            For Each Alum_out As Alumno In Alumnos_out

                If ((Alum_out.SGA Is Nothing) And Alum_out.Matricula Is Nothing And Alum_out.DAE Is Nothing) Then
                    ' Si las tres listas están vacías no se ha de imprimir nada
                Else

                    document.Add(New Paragraph("ALUMNO:"))

                    document.Add(New Paragraph("(" & Alum_out.CodigoHZ & ")" & " " & Alum_out.Apellido1 & " " & Alum_out.Apellido2 & ", " & Alum_out.Nombre & " (" & Alum_out.Enseñanza & Alum_out.Curso & ")"))


                    If (Alum_out.SGA Is Nothing) Then
                        'SGA OK > No se ha de imprimir nada
                    Else
                        document.Add(New Paragraph("ASIGNATURAS QUE FALTAN / ERRÓNEAS EN EL SGA"))

                        For Each asigI As Asignatura In Alum_out.SGA
                            If asigI.CodigoHZAsig = "XXXX" Then
                                document.Add(New Paragraph("Alumno no registrado en el SGA"))
                                Exit For

                            End If
                            document.Add(New Paragraph(asigI.CodigoHZAsig & " " & asigI.Descripcion))

                        Next


                    End If




                    If (Alum_out.Matricula Is Nothing) Then
                        'Matricula OK > No se ha de imprimir nada
                    Else
                        document.Add(New Paragraph("ASIGNATURAS QUE FALTAN/ERRÓNEAS EN LA MATRÍCULA:"))

                        For Each asigM As Asignatura In Alum_out.Matricula


                            If asigM.CodigoHZAsig = "XXXX" Then
                                document.Add(New Paragraph("Alumno no matriculado"))
                                Exit For
                            End If
                            document.Add(New Paragraph(asigM.CodigoHZAsig & " " & asigM.Descripcion))

                        Next

                    End If




                    If (Alum_out.DAE Is Nothing) Then
                        'DAE OK > No se ha de imprimir nada
                    Else
                        document.Add(New Paragraph("ASIGNATURAS QUE FALTAN/ERRÓNEAS EN EL DAE:"))

                        For Each asigD As Asignatura In Alum_out.DAE

                            If asigD.CodigoHZAsig = "XXXX" Then
                                document.Add(New Paragraph("Alumno no contemplado en el DAE"))
                                Exit For
                            End If
                            document.Add(New Paragraph(asigD.CodigoHZAsig & " " & asigD.Descripcion))

                        Next

                    End If

                    document.Add(New Paragraph("*************************************************************"))

                End If

            Next

            cb.EndText()



            document.Close()

            MessageBox.Show("Resultado generado correctamente ", "OK", MessageBoxButtons.OK)

        Catch ex As Exception
            MessageBox.Show("Error en la generación", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try




    End Sub

    'Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load


    '#If DEBUG Then
    'OpenFileDialog1.FileName = "C:\Maite\FicheroSGA_PR3.csv"
    'OpenFileDialog2.FileName = "C:\Maite\Fichero_Matricula_PR3.csv"
    'OpenFileDialog3.FileName = "C:\Maite\Fichero_DAE_PR2.csv"

    'TextBox1.Text = OpenFileDialog1.FileName
    'TextBox2.Text = OpenFileDialog2.FileName
    'TextBox3.Text = OpenFileDialog3.FileName
    '#End If
    'End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class
