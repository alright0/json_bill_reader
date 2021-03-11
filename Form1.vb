' версия 1.4

Imports Newtonsoft.Json.Linq
Imports System.Security.Cryptography
Imports System.Text
Imports System.IO

Public Class Form1
    Private Sub Form1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles Me.DragDrop

        ' получение полного имени файла из драгндропа.
        Dim raw_file() As String = CType(e.Data.GetData(DataFormats.FileDrop), String())

        ' обработка исключений: Если файл имеет расширение не json, 
        ' то необходимо отборсить его счетчик массива нужно будет пересчитать
        Dim inc% = 0            ' инкремент валидных файлов
        Dim file() As String    ' массив валидных файлов
        For Each fl In raw_file

            If FileIO.FileSystem.GetFileInfo(fl).Extension <> ".json" Then

                MsgBox(FileIO.FileSystem.GetFileInfo(fl).Name &
                       " не может быть обработан, так как имеет неверный формат")
                Continue For
            Else

                ReDim Preserve file(inc)
                inc += 1
                file(inc - 1) = fl
            End If
        Next

        'если после первой обраотки массив пустой, то выйти из проедуры
        If IsNothing(file) Then

            Exit Sub
        End If

        ' основной цикл обработки файлов
        For I As Integer = 0 To file.Length - 1

            Dim ShortName$ = IO.Path.GetFileNameWithoutExtension(file(I))          ' имя файла без расширения
            Dim PathName$ = FileIO.FileSystem.GetFileInfo(file(I)).DirectoryName   ' путь
            Dim FullName$ = FileIO.FileSystem.GetFileInfo(file(I)).Name            ' имя с расширением

            Dim raw_json As New StreamReader(file(I))   ' сырой файл
            Dim contents$ = raw_json.ReadToEnd()        ' прочитанный файл
            Dim total_q& = 0                            ' количество закупленных шт. в чеке

            Dim quantity&                         'long int
            Dim name$, user$, prefix$, fOut$      'string
            Dim price#, sum_price#, total_sum#    'double
            Dim items, dtime, bill_date, json_obj 'variant


            ' проверка, является ли файл json-файлом
            Try
                json_obj = JObject.Parse(contents)
            Catch ex As Exception
                MsgBox($"Файл {FullName} не может быть обработан, так как не является JSON-структурированным файлом")
                Continue For
            End Try


            ' переключение между схемами JSON для IOS и Android
            If Not IsNothing(json_obj.SelectToken("['ticket.document.receipt']")) Then
                prefix = "['ticket.document.receipt']."
            Else
                prefix = ""
            End If

            ' try здесь выступает в качестве валидатора json - если файл не является чеком, 
            ' то исключение будет поймано на списке items, 
            ' все остальные поля будут равны nothnig
            Try
                ' название магазина
                user = json_obj.SelectToken($"{prefix}user")

                'дата создания чека в unixtime
                bill_date = json_obj.SelectToken($"{prefix}dateTime")

                ' здесь конвертируется unixtime в datetime. 
                ' версия для android получает значение в unixtime, для IOS в datetime
                ' поэтому необходимо отловить правильную версию
                Try
                    dtime = CDate(utime_to_date(bill_date)).ToString("dd.MM.yyyy HH:mm:ss")
                Catch
                    dtime = CDate(bill_date).ToString("dd.MM.yyyy HH:mm:ss")
                End Try

                'итоговая сумма
                total_sum = CDbl(json_obj.SelectToken($"{prefix}totalSum")) / 100

                ' основная стока, которая собирает всю информацию
                fOut = ""

                ' запись в основную строку
                fOut = $"{user};Дата:;{dtime};{Chr(10)};{Chr(10)}"
                fOut = $"{fOut}Продукт;Количество;Цена за шт.;Цена;{Chr(10)}"

                ' массив позиций покупок в чеке
                items = json_obj.SelectToken($"{prefix}items")

                ' обработка всех позиций
                For Each j In items

                    quantity = j.SelectToken("quantity")          ' количество шт. позиции
                    name = j.SelectToken("name")                  ' название позиции
                    price = CDbl(j.SelectToken("price")) / 100    ' стоимость за штуку
                    sum_price = CDbl(j.SelectToken("sum")) / 100  ' итоговая стоимость

                    ' запись в основную строку
                    fOut = $"{fOut}{name};{quantity};{price};{sum_price};{Chr(10)}"

                    ' ирекурсивная запись всех строк
                    total_q += quantity
                Next

                ' запись в основную строку
                fOut = $"{fOut}{Chr(10)}ИТОГО:;{total_q};;{total_sum}"
                'fOut = fOut & Chr(10) & "ИТОГО:;" & total_q & ";;" & total_sum

            Catch ex As Exception
                MsgBox($"Файл {FullName} не может быть обработан, так как не является кассовым чеком")
                Continue For
            End Try

            ' обработка исключений. Если csv файл существует и открыт, то 
            ' перейти на следующую итерацию
            If My.Computer.FileSystem.FileExists($"{pathname}\{shortname}.csv") Then

                Dim flag As Integer = MsgBox($"файл {FullName} уже существует.{Chr(10)}Хотите перезаписать его?", vbYesNo)

                ' если файл существует, то проверить, не занят ли он и перезаписать.
                ' если файл занят - вывести сообщение, что файл занят и предложить 
                ' повторно перезаписать или пропустить
                If flag = 6 Then

                    Try
                        My.Computer.FileSystem.WriteAllText($"{pathname}\{shortname}.csv", fOut, False)
                    Catch ex As System.IO.IOException

                        ' в этом месте происходит обработка случая, когда файл, который надо перезаписать занят другим процессом.
                        Dim msgstatus% = MsgBox($"Файл {FullName} Не может быть сохранен, так как занят другим приложением!{Chr(10)}Закройте файл и попробуйте еще раз.", MsgBoxStyle.RetryCancel)
                        While msgstatus = 4
                            Try
                                My.Computer.FileSystem.WriteAllText($"{PathName}\{ShortName}.csv", fOut, False)
                                msgstatus = 0 ' необходимо сбросить файл ответа, иначе не получится выйти из цикла при успешной перезаписи
                            Catch
                                msgstatus% = MsgBox($"Файл {FullName} Не может быть сохранен, так как занят другим приложением!{Chr(10)}Закройте файл и попробуйте еще раз.", MsgBoxStyle.RetryCancel)
                            End Try
                        End While
                        If msgstatus = 2 Then Continue For
                    End Try
                End If
            Else
                My.Computer.FileSystem.WriteAllText($"{pathname}\{shortname}.csv", fOut, False)
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