Imports System.Data.SqlClient



Public Class Form1
    Dim POSPrint As New PrinterClass(Application.StartupPath)
    'Dim P As New PrinterClass(Application.StartupPath)
    Dim ConStr As String
    Dim StationNo As String
    Dim KodeCompany As String, KodeDivisi As String
    Dim oConn As New SqlConnection '("data source=DESKTOP-50PF94A\SQLEXPRESS; initial catalog=POSRT;user id=SA;password=indofood;Application Name=KOP_DSI;MultipleActiveResultSets=true;")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'With POSPrint
        '    'Printing Logo
        '    .SetFont(10, "Tahoma")
        '    .RTL = False
        '    .PrintLogo()

        '    'Printing Title
        '    '.FeedPaper(1)
        '    .AlignCenter()
        '    .BigFont()
        '    .Bold = True
        '    .WriteLine("Sales Receipt")

        '    'Printing Date
        '    .GotoSixth(1)
        '    .SmallFont()
        '    .AlignLeft()
        '    '.WriteChars("Date:")
        '    .WriteLine("Date : " + DateTime.Now.ToString)
        '    .WriteLine("")
        '    .DrawLine()
        '    .FeedPaper(1)

        '    'Printing Header
        '    .GotoCol(2)
        '    .WriteChars("#")
        '    .GotoSixth(2)
        '    .WriteChars("Description")
        '    .GotoSixth(5)
        '    .WriteChars("Count")
        '    .GotoSixth(6)
        '    .WriteChars("Total")
        '    .WriteLine("")
        '    .DrawLine()
        '    '.FeedPaper(1)

        '    'Printing Items
        '    .SmallFont()
        '    Dim i As Integer
        '    For i = 1 To 6
        '        .GotoSixth(1)
        '        .WriteChars(i)
        '        .GotoSixth(2)
        '        .WriteChars("Item# " & (Rnd() * 100) \ 1)
        '        .GotoSixth(5)
        '        .WriteChars(Rnd() * 10 \ 1)
        '        .GotoSixth(6)
        '        .WriteChars((Rnd() * 50 \ 1) & " JD(s)")
        '        .WriteLine("")
        '    Next

        '    'Printing Totals
        '    .NormalFont()
        '    .DrawLine()
        '    .GotoSixth(1)
        '    .UnderlineOn()
        '    .WriteChars("Total")
        '    .UnderlineOff()
        '    .GotoSixth(5)
        '    .WriteChars((Rnd() * 300 \ 1) & " JD(s)")
        '    .CutPaper() ' Can be used with real printer to cut the paper.

        '    'Ending the session
        '    .EndDoc()
        'End With
        'End

    End Sub

    Private Sub Form1_Deactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Deactivate
        oConn.Close()
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Label1.Text = POSPrint.PrinterWidth
        'POSPrint.PrinterWidth = 3600
        Try
            StationNo = My.Resources.StationNo
            KodeCompany = My.Resources.KodeCompany
            KodeDivisi = My.Resources.KodeDivisi

            ConStr = My.Resources.Conn
            oConn = New SqlConnection(ConStr)
            oConn.Open()
        Catch ex As Exception
            MessageBox.Show(ex.Message, "error")
            End
        End Try

    End Sub

    Public Function GetDataSet(ByVal sSQL As String) As DataSet



        Dim oComm As New SqlCommand
        Dim oSQLDataAdapter As SqlDataAdapter
        Dim dsRecordData As DataSet

        oComm.Connection = oConn
        oComm.CommandType = CommandType.Text
        oComm.CommandTimeout = 0
        oComm.CommandText = sSQL

        If oConn.State = 0 Then oConn.Open()

        oSQLDataAdapter = New SqlDataAdapter
        oSQLDataAdapter.SelectCommand = oComm

        'initialize dataset
        dsRecordData = New DataSet("RecordData")
        oSQLDataAdapter.Fill(dsRecordData)

        'oConn.Close()

        Return dsRecordData
    End Function

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'ReadDataAndPrint("01", "2", "", "Dela")
        ReadSpool()
    End Sub

    Private Sub ReadDataAndPrint(ByVal KodeCompany As String, ByVal KodeCabang As String, ByVal TX_No As String, ByVal AnggotaID As String, ByVal Kasir As String, ByVal StationNo As String)

        Dim ds As New DataSet
        Dim sAlign As String
        Dim sFormat As String
        Dim sSize As String
        Dim iPosCol As Integer
        Dim sNL As String
        Dim stxt As String

        Try

            With POSPrint
                .PrinterWidth = 3600
                .SetFont(8, "Tahoma")
                .RTL = False
                .PrintLogo()

                ds = GetDataSet("sp_GeneratePOSTbl '" + KodeCompany + "','" + KodeCabang + "','" + TX_No + "','" + AnggotaID + "','" + Kasir + "','" + StationNo + "'")

                Dim tableTMP As DataTable = ds.Tables(0)

                .AlignCenter()
                .BigFont()
                '.Bold = True
                .WriteLine("Kwitansi")
                '.WriteLine(CStr(POSPrint.PrinterWidth()))

                For Each row As DataRow In tableTMP.Rows
                    sAlign = Convert.ToString(row("Align")).ToUpper
                    sFormat = Convert.ToString(row("Format")).ToUpper
                    sSize = Convert.ToString(row("Size")).ToUpper
                    iPosCol = Convert.ToInt32(row("PosCol"))
                    sNL = Convert.ToString(row("NL")).ToUpper
                    stxt = Convert.ToString(row("txt"))

                    Select Case sAlign
                        Case "L" : .AlignLeft()
                        Case "C" : .AlignCenter()
                        Case "R" : .AlignRight()
                    End Select

                    .Bold = False
                    .UnderlineOff()
                    Select Case sFormat
                        Case "B" : .Bold = True
                        Case "U" : .UnderlineOn()
                    End Select

                    .NormalFont()
                    Select Case sFormat
                        Case "B" : .BigFont()
                        Case "S" : .SmallFont()
                        Case "R" : .NormalFont()
                    End Select

                    .GotoSixth(iPosCol)
                    If sNL = "Y" Then
                        .WriteLine(stxt)
                    Else
                        .WriteChars(stxt)
                    End If
                Next
                .CutPaper() ' Can be used with real printer to cut the paper.

                'Ending the session
                .EndDoc()

                txt1.AppendText("Print Success - Invoice:" + TX_No + " / " + DateTime.Now.ToString() + Environment.NewLine)
            End With

        Catch ex As Exception
            txt1.AppendText("************** Print Failed - Invoice:" + TX_No + " / " + DateTime.Now.ToString() + " Err:" + ex.Message.ToString() + Environment.NewLine)
        End Try
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ReadSpool()
    End Sub

    Private Sub ReadSpool()
        Dim command As SqlCommand = New SqlCommand("SELECT top 1 * FROM Util_SpoolPrint where CompanyID='" + KodeCompany + "' AND DivisiID='" + KodeDivisi + "' AND StationNo='" + StationNo + "'", oConn)

        If oConn.State = 0 Then oConn.Open()

        Dim reader As SqlDataReader = command.ExecuteReader()

        If reader.HasRows Then
            Timer1.Enabled = False
            While reader.Read()
                lblStat.Text = "Printing"
                ReadDataAndPrint(KodeCompany, reader.GetString(1), reader.GetString(2), reader.GetString(3), reader.GetString(4), reader.GetString(6))
                reader.Close()
                Application.DoEvents()
                lblStat.Text = "READY"
                Exit Sub
            End While

        End If
    End Sub
End Class
