Imports Newtonsoft.Json.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class Form1
    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop

        ' получение полного имени файла из драгндропа.
        Dim raw_file() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        ' массив, получемый после отбасывания несоответствующих файлов(не json)
        Dim file() As String
        file = raw_file
        ' обработка исключений: Если файл имеет расширение не json, то необходимо отборсить его
        ' счетчик массива нужно будет пересчитать`
        For I As Integer = 0 To file.Length - 1
            Dim extension As String = FileIO.FileSystem.GetFileInfo(raw_file(I)).Extension

            If extension <> ".json" Then
                MsgBox("Файл " & FileIO.FileSystem.GetFileInfo(raw_file(I)).Name &
                       " не может быть обработан, так как имеет неверный формат")
            Else
                ReDim Preserve file(I)
                file(I) = raw_file(I)

            End If

        Next


        'если после первой обраотки массив пустой, то выйти из проедуры
        If IsNothing(file) Then
            Exit Sub
        End If


        For I As Integer = 0 To file.Length - 1

            Dim shortname As String = IO.Path.GetFileNameWithoutExtension(file(I))          ' имя файла без расширения
            Dim pathname As String = FileIO.FileSystem.GetFileInfo(file(I)).DirectoryName   ' путь

            Dim raw_json As New StreamReader(file(I))       ' сырой файл
            Dim contents As String = raw_json.ReadToEnd()   ' прочитанный файл
            Dim total_q As Int64 = 0                        ' количество закупленных шт. в чеке
            Dim json_obj                                    ' пропарсенный json

            Dim quantity As Int16
            Dim name As String
            Dim price As Double
            Dim sum_price As Double
            Dim user As String
            Dim bill_date
            Dim total_sum As Double
            Dim fOut As String = ""
            Dim items
            Dim dtime

            Try
                json_obj = JObject.Parse(contents)
            Catch ex As Exception
                MsgBox("Файл " & shortname & " не может быть обработан, так как не json-структурированным файлом")
                Continue For
            End Try


            If Not IsNothing(json_obj.SelectToken("['ticket.document.receipt']")) Then


                ' обработка исключений. Здесь проверяется корректность входящего json. Если присутствуют 
                ' метки fiscalDocumentNumber и fiscalDriveNumber, то, это, скорее всего, правильный файл
                Try

                    Dim c_block0 As String = json_obj.SelectToken("['ticket.document.receipt'].operator")
                    Dim c_block1 As String = json_obj.SelectToken("['ticket.document.receipt'].retailPlace")

                    If IsNothing(c_block0) Or IsNothing(c_block0) Then
                        Throw New Exception("JSONControlBlocksFail")
                    End If
                Catch ex As Exception
                    MsgBox("Файл " & shortname & " не может быть обработан, так как не является кассовым чеком")
                    Continue For
                End Try

                user = json_obj.SelectToken("['ticket.document.receipt'].retailPlace")               ' название магазина
                bill_date = json_obj.SelectToken("['ticket.document.receipt'].dateTime")      ' дата создания чека d unixtime


                total_sum = (json_obj.SelectToken("['ticket.document.receipt'].totalSum")) / 100 'итоговая сумма
                fOut = ""     ' основная стока, которая собирает всю информацию

                ' запись в основную строку
                fOut = user & ";Дата:;" & bill_date & ";" & Chr(10) & ";" & Chr(10)
                fOut = fOut & "Продукт;Количество;Цена за шт.;Цена;" & Chr(10)

                items = json_obj.SelectToken("['ticket.document.receipt'].items")   ' массив позиций покупок в чеке

                ' обработка всех позиций
                For Each j In items

                    quantity = j.SelectToken("quantity")           ' количество шт. позиции
                    name = j.SelectToken("name")                  ' название позиции
                    price = CDbl(j.SelectToken("price")) / 100    ' стоимость за штуку
                    sum_price = CDbl(j.SelectToken("sum")) / 100  ' итоговая стоимость

                    ' запись в основную строку
                    fOut = fOut & name & ";"
                    fOut = fOut & quantity.ToString & ";"
                    fOut = fOut & price.ToString & ";"
                    fOut = fOut & sum_price.ToString & ";" & Chr(10)

                    ' инкремент штук(они нигде не указаны в json)
                    total_q += quantity
                Next

            Else


                ' обработка исключений. Здесь проверяется корректность входящего json. Если присутствуют 
                ' метки fiscalDocumentNumber и fiscalDriveNumber, то, это, скорее всего, правильный файл
                Try

                    Dim c_block0 As String = json_obj.SelectToken("operator")
                    Dim c_block1 As String = json_obj.SelectToken("retailPlaceAddress")

                    If IsNothing(c_block0) Or IsNothing(c_block0) Then
                        Throw New Exception("JSONControlBlocksFail")
                    End If
                Catch ex As Exception
                    MsgBox("Файл " & shortname & " не может быть обработан, так как не является кассовым чеком")
                    Continue For
                End Try

                user = json_obj.SelectToken(".user")               ' название магазина
                bill_date = json_obj.SelectToken("dateTime")      ' дата создания чека d unixtime

                dtime = utime_to_date(bill_date) ' здесь конвертируется unixtime в datetime
                total_sum = CDbl(json_obj.SelectToken("totalSum")) / 100 'итоговая сумма
                fOut = ""     ' основная стока, которая собирает всю информацию

                ' запись в основную строку
                fOut = user & ";Дата:;" & dtime & ";" & Chr(10) & ";" & Chr(10)
                fOut = fOut & "Продукт;Количество;Цена за шт.;Цена;" & Chr(10)

                items = json_obj.SelectToken("items")   ' массив позиций покупок в чеке

                ' обработка всех позиций
                For Each j In items

                    quantity = j.SelectToken("quantity")           ' количество шт. позиции
                    name = j.SelectToken("name")                  ' название позиции
                    price = CDbl(j.SelectToken("price")) / 100    ' стоимость за штуку
                    sum_price = CDbl(j.SelectToken("sum")) / 100  ' итоговая стоимость

                    ' запись в основную строку
                    fOut = fOut & name & ";"
                    fOut = fOut & quantity.ToString & ";"
                    fOut = fOut & price.ToString & ";"
                    fOut = fOut & sum_price.ToString & ";" & Chr(10)

                    ' инкремент штук(они нигде не указаны в json)
                    total_q += quantity
                Next
            End If




            ' запись в основную строку
            fOut = fOut & Chr(10) & "ИТОГО:;" & total_q & ";;" & total_sum

            ' обработка исключений. Если csv файл существует и открыт, то 
            ' перейти на следующую итерацию
            If My.Computer.FileSystem.FileExists(pathname & "\" & shortname & ".csv") Then

                Dim flag As Integer = MsgBox("файл " & pathname & "\" & shortname &
                    ".csv уже существует. " & Chr(10) & "Хотите перезаписать его?", vbYesNo)

                ' если файл существует и перезаписать = да, проверить, не занят ли он и записать.
                ' если файл занят - вывести сообщение, что файл занят и перейти на следующую итерацию
                If flag = 6 Then

                    Try
                        My.Computer.FileSystem.WriteAllText(pathname & "\" & shortname & ".csv", fOut, False)
                    Catch ex As System.IO.IOException
                        MsgBox("файл " & pathname & "\" & shortname & ".csv Не может быть сохранен, так как занят другим приложением.")
                        Continue For
                    End Try
                End If
            Else
                My.Computer.FileSystem.WriteAllText(pathname & "\" & shortname & ".csv", fOut, False)

            End If

        Next
    End Sub
    Private Function utime_to_date(bill_date As Double)
        ' эта функция преобразует, unixtime в стандартный формат времени

        ' точка отсчета по unixtime и добавление секунд по unixtime
        Dim dtime = New DateTime(1970, 1, 1, 0, 0, 0, 0)
        utime_to_date = dtime.AddSeconds(bill_date)

    End Function
    Private Sub Form1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragEnter

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If

    End Sub
    Shared Function lic_compare() As Boolean
        ' функция для проверки файла лицензии 
        ' и фактического хэша FeatureSet

        ' файл лиицензии
        Dim lic As String = My.Computer.FileSystem.ReadAllText(Application.StartupPath & "\license.lic")


        Dim FSet = My.Computer.Registry.GetValue(
            "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardProduct", Nothing)

        Dim Sset = My.Computer.Registry.GetValue(
            "HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\CentralProcessor\0", "ProcessorNameString", Nothing)

        Dim Dset = FSet & " " & Sset

        ' хэширвание значения из реестра
        Dim req As String = GetHash(GetHash(Dset) & "FFFFFF")

        If lic = req Then
            lic_compare = True
        Else
            lic_compare = False
        End If

    End Function
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' в зависимости от состояния лицензии: заблокировать перетаскивание или нет
        If My.Computer.FileSystem.FileExists(Application.StartupPath & "\license.lic") Then

            If lic_compare() Then

                AllowDrop = True
                Label1.Text = "перетащите файлы " & Chr(10) & "в это окно"
                Label2.Text = "Лицензия активна"
                Label2.Visible = False
                Button1.Visible = False
            Else
                AllowDrop = False
                Label1.Text = "перетаскивание " & Chr(10) & "недоступно"
                MsgBox("Лицензия не действительна")
            End If
        Else

            AllowDrop = False
            Label1.Text = "перетаскивание " & Chr(10) & "недоступно"
            Label2.Text = "Лицензия неактивна"

        End If

    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Form3.Show()

    End Sub
    Shared Function GetHash(theInput As String) As String

        Using hasher As MD5 = MD5.Create()    ' create hash object

            ' Convert to byte array and get hash
            Dim dbytes As Byte() =
                 hasher.ComputeHash(Encoding.UTF8.GetBytes(theInput))

            ' sb to create string from bytes
            Dim sBuilder As New StringBuilder()

            ' convert byte data to hex string
            For n As Integer = 0 To dbytes.Length - 1
                sBuilder.Append(dbytes(n).ToString("X2"))
            Next n

            Return sBuilder.ToString()
        End Using

    End Function

End Class