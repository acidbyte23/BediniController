Option Explicit On

Imports System.IO.Ports
Imports System.Net
Imports System.Text.Encoding


Public Class Form1

    'set the inifile path
    Public ReadOnly inifile As String = Application.StartupPath + "\config.ini"


    'declare library
    Private Declare Function WritePrivateProfileStringA Lib "kernel32.dll" (ByVal lpSectionName As String, ByVal lpKeyName As String,
                                                                        ByVal lpString As String, ByVal lpFileName As String) As Long

    'declare GetPrivateProfileString from kernel32 library
    Private Declare Auto Function GetPrivateProfileString Lib "kernel32.dll" (ByVal lpAppName As String, ByVal lpKeyName As String,
                                                                          ByVal lpDefault As String, ByVal lpReturnedString As String,
                                                                          ByVal nSize As Integer, ByVal lpFileName As String) As Integer


    ' get strings from ini file function
    Public Shared Function sGetIni(ByVal sINIFile As String, ByVal sSection As String,
                                                       ByVal sKey As String, ByVal sDefault As String) As String
        'temporary space holders
        Dim sTemp As String = Space(255)
        Dim nLength As Integer

        'calling windows dll function
        nLength = GetPrivateProfileString(sSection, sKey, sDefault, sTemp, 255, sINIFile)

        'return result
        Return sTemp.Substring(0, nLength)

    End Function


    Dim msChart As MSChartWrapper.ChartForm

    Dim serialConnection As Boolean
    Dim serialDataRx(16) As Byte

    Dim graphPlot(1) As Double
    Dim yMinimumVoltage As Double
    Dim yMaximumVoltage As Double
    Dim yMinimumCurrent As Double
    Dim yMaximumCurrent As Double

    Dim goingDown As Boolean

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles connectUdp.Click

        If IpAdress.Text IsNot "" And Port.Text IsNot "" Then
            AxWinsock1.RemoteHost = IpAdress.Text
            AxWinsock1.RemotePort = Port.Text
            connectUdp.Visible = False
            disconnectUdp.Visible = True

            Try
                AxWinsock1.Connect()
            Catch ex As Exception
                connectUdp.Visible = True
                disconnectUdp.Visible = False
                'MsgBox(ex.Message)
            End Try
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles disconnectUdp.Click
        connectUdp.Visible = True
        disconnectUdp.Visible = False
        Try
            AxWinsock1.Close()
        Catch ex As Exception
            connectUdp.Visible = False
            disconnectUdp.Visible = True
            'MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub AxWinsock1_ConnectEvent(sender As Object, e As EventArgs) Handles AxWinsock1.ConnectEvent
        connectUdp.Visible = False
        disconnectUdp.Visible = True
    End Sub

    Private Sub Form1_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        'check if ini file exists
        If Not System.IO.File.Exists(inifile) Then
            'open new inifile for edit
            Dim TextFile As New System.IO.StreamWriter(inifile)
            'make default ini file
            TextFile.WriteLine("[FORM]")
            TextFile.WriteLine("name=Bedini Studio")
            TextFile.WriteLine("")
            TextFile.WriteLine("[BASIC]")
            TextFile.WriteLine("")
            TextFile.WriteLine("[PID]")
            TextFile.WriteLine("")
            TextFile.WriteLine("[PRIMAIRY]")
            TextFile.WriteLine("")
            TextFile.WriteLine("[SECUNDAIRY]")
            TextFile.WriteLine("")
            TextFile.WriteLine("[ADVANCED]")
            TextFile.WriteLine("")
            'save default ini file
            TextFile.Close()
        End If

        ' Get Comm Ports
        For Each sp As String In My.Computer.Ports.SerialPortNames
            commPorts.Items.Add(sp)
        Next

        ' Load Parameter Screens
        basisSetings.Visible = True
        pidSetings.Visible = False
        primairySettings.Visible = False
        secundairySettings.Visible = False
        advancedSettings.Visible = False



        yMinimumVoltage = 0
        yMaximumVoltage = 2

        graphPlot(0) = 0
        graphPlot(1) = 0

        commPorts.SelectedIndex = 0
        baudrateSerial.SelectedIndex = 0
    End Sub

    Private Sub connectSerial_Click(sender As Object, e As EventArgs) Handles connectSerial.Click
        SerialPort1.PortName = commPorts.SelectedItem.ToString
        SerialPort1.BaudRate = baudrateSerial.SelectedItem.ToString
        SerialPort1.ReadTimeout = 2000
        connectSerial.Visible = False
        disconnectSerial.Visible = True

        Try
            SerialPort1.Open()
        Catch ex As Exception
            connectSerial.Visible = True
            disconnectSerial.Visible = False
            'MsgBox(ex.Message)
        End Try

        If SerialPort1.IsOpen = True Then
            connectUdp.Visible = False
            serialConnection = True
        End If

    End Sub

    Private Sub disconnectSerial_Click(sender As Object, e As EventArgs) Handles disconnectSerial.Click
        connectSerial.Visible = True
        disconnectSerial.Visible = False

        Try
            SerialPort1.Close()
        Catch ex As Exception
            connectSerial.Visible = False
            disconnectSerial.Visible = True
            'MsgBox(ex.Message)
        End Try

        If SerialPort1.IsOpen = False Then
            connectUdp.Visible = True
            serialConnection = False
        End If

    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        If serialConnection = True Then

            For i As Integer = 0 To SerialPort1.BytesToRead
                serialDataRx(i) = SerialPort1.ReadByte()
            Next

            serialDataRecieved.Text = serialDataRecieved.Text + serialDataRx(0).ToString + serialDataRx(1).ToString + serialDataRx(2).ToString + serialDataRx(3).ToString + serialDataRx(4).ToString + serialDataRx(5).ToString + serialDataRx(6).ToString + serialDataRx(7).ToString + serialDataRx(8).ToString + serialDataRx(9).ToString + serialDataRx(10).ToString + serialDataRx(11).ToString + serialDataRx(12).ToString + serialDataRx(13).ToString + serialDataRx(14).ToString + serialDataRx(15).ToString + serialDataRx(16).ToString + vbNewLine

            If serialDataRx(0) = &H70 Then
                Dim bytes As Byte() = {serialDataRx(2), serialDataRx(3), serialDataRx(4), serialDataRx(5)}
                Select Case serialDataRx(1)
                    Case &H10
                        graphPlot(0) = BitConverter.ToDouble(bytes, 0)
                    Case &H20
                        graphPlot(1) = BitConverter.ToDouble(bytes, 0)
                    Case &H30

                    Case &H40

                    Case &H50

                End Select
            End If
        End If

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ' just a osc signal
        If graphPlot(0) > 999 Then
            goingDown = True
        End If
        If graphPlot(0) < -999 Then
            goingDown = False
        End If

        If goingDown = True Then
            graphPlot(0) = graphPlot(0) - 1
            graphPlot(1) = graphPlot(1) + 6
        Else
            graphPlot(0) = graphPlot(0) + 1
            graphPlot(1) = graphPlot(1) - 6
        End If

        ' from here

        yMinimumVoltage = graphPlot(0)
        yMinimumCurrent = graphPlot(1)

        Dim voltageAvg As Double = 0.0
        Dim currentAvg As Double = 0.0

        For i As Integer = 0 To Chart1.Series.Item("Series1").Points.Count
            If Chart1.Series.Item("Series1").Points.Count > i Then
                voltageAvg = voltageAvg + Chart1.Series.Item("Series1").Points.Item(i).YValues.GetValue(0)
            End If
        Next

        If Chart1.Series.Item("Series1").Points.Count > 0 Then
            yMinimumVoltage = Chart1.Series.Item("Series1").Points.FindMinByValue.YValues.GetValue(0)
            yMaximumVoltage = Chart1.Series.Item("Series1").Points.FindMaxByValue.YValues.GetValue(0)
        End If


        For i As Integer = 0 To Chart1.Series.Item("Series2").Points.Count
            If Chart1.Series.Item("Series2").Points.Count > i Then
                currentAvg = currentAvg + Chart1.Series.Item("Series2").Points.Item(i).YValues.GetValue(0)
            End If
        Next

        If Chart1.Series.Item("Series2").Points.Count > 0 Then
            yMinimumCurrent = Chart1.Series.Item("Series2").Points.FindMinByValue.YValues.GetValue(0)
            yMaximumCurrent = Chart1.Series.Item("Series2").Points.FindMaxByValue.YValues.GetValue(0)
        End If


        Dim voltageAvgTemp As Double
        Dim currentAvgTemp As Double

        voltageAvgTemp = (voltageAvg / Chart1.Series.Item("Series1").Points.Count)
        currentAvgTemp = (currentAvg / Chart1.Series.Item("Series2").Points.Count) * 1000

        Label7.Text = voltageAvgTemp.ToString
        Label8.Text = currentAvgTemp.ToString

        ' If Chart1.Series.Item("Series1").Points.Count >= indexMaximum And Chart1.Series.Item("Series1").Points.Count > indexMaximum Then
        ' yMaximumVoltage = Chart1.Series.Item("Series1").Points.Item(indexMaximum).YValues.GetValue(0)
        ' End If

        'primairy battery
        Chart1.Series.Item("Series1").Points.Add(graphPlot(0) / 100)


        Chart1.Series.Item("Series2").Points.Add((graphPlot(1) / 1000))


        If Chart1.Series.Item("Series1").Points.Count > 1000 Then
            Chart1.Series.Item("Series1").Points.RemoveAt(0)
        End If
        If Chart1.Series.Item("Series2").Points.Count > 1000 Then
            Chart1.Series.Item("Series2").Points.RemoveAt(0)
        End If

        'secundairy battery
        Chart4.Series.Item("Series1").Points.Add(graphPlot(0) / 100)


        Chart4.Series.Item("Series2").Points.Add((graphPlot(1) / 1000))


        'If Chart4.Series.Item("Series1").Points.Count > 1000 Then
        '    Chart4.Series.Item("Series1").Points.RemoveAt(0)
        'End If
        'If Chart4.Series.Item("Series2").Points.Count > 1000 Then
        '    Chart4.Series.Item("Series2").Points.RemoveAt(0)
        'End If

        Dim voltage As Double
        Dim voltage1 As Double
        Dim current As Double
        Dim current1 As Double

        If yMinimumVoltage >= 0 Then
            voltage = yMinimumVoltage
            voltage1 = yMinimumVoltage
        Else
            voltage = yMinimumVoltage
            voltage1 = yMinimumVoltage
        End If

        If yMinimumCurrent >= 0 Then
            current = yMinimumCurrent * 1000
            current1 = yMinimumCurrent
        Else
            current = yMinimumCurrent * 1000
            current1 = yMinimumCurrent
        End If


        If current >= 0 And voltage1 >= 0 Then
            If current1 < voltage1 Then
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = yMinimumCurrent.ToString
            Else
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = yMinimumVoltage.ToString
            End If
        ElseIf current >= 0 And voltage1 < 0 Then
            Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = voltage1.ToString
        ElseIf current < 0 And voltage1 >= 0 Then
            Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = current1.ToString
        ElseIf current1 < 0 And voltage1 < 0 Then
            If current1 < voltage1 Then
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = current1.ToString
            Else
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Minimum = voltage1.ToString
            End If
        End If


        Label3.Text = voltage.ToString
        Label17.Text = current.ToString


        If yMaximumVoltage >= 0 Then
            voltage = yMaximumVoltage
            voltage1 = yMaximumVoltage
        Else
            voltage = yMaximumVoltage
            voltage1 = yMaximumVoltage / 100
        End If

        If yMaximumCurrent >= 0 Then
            current = yMaximumCurrent * 1000
            current1 = yMaximumCurrent
        Else
            current = yMaximumCurrent * 1000
            current1 = yMaximumCurrent / 1000
        End If

        If current1 >= 0 And voltage1 >= 0 Then
            If current1 > voltage1 Then
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = yMaximumCurrent.ToString
            Else
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = yMaximumVoltage.ToString
            End If
        ElseIf current1 >= 0 And voltage1 < 0 Then
            Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = current1.ToString
        ElseIf current < 0 And voltage1 >= 0 Then
            Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = voltage1.ToString
        ElseIf current < 0 And voltage1 < 0 Then
            If current1 > voltage1 Then
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = current1.ToString
            Else
                Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = voltage1.ToString
            End If
        End If


        'If yMaximumVoltage > yMinimumVoltage Then
        'Chart1.ChartAreas.Item(" ChartArea1").AxisY.Maximum = voltage.ToString
        'End If

        Label5.Text = voltage.ToString
        Label18.Text = current.ToString

        voltage = graphPlot(0) / 100
        Label6.Text = voltage.ToString
        Label19.Text = graphPlot(1).ToString

        Dim currentWatts As Double
        Dim avarageWatts As Double

        currentWatts = (voltage) * (graphPlot(1) / 1000)
        Label22.Text = currentWatts.ToString

        avarageWatts = voltageAvgTemp * (currentAvgTemp / 1000)
        Label23.Text = avarageWatts.ToString

        ' watts in
        Chart2.Series.Item("Series1").Points.Add((currentWatts))

        If Chart2.Series.Item("Series1").Points.Count > 1000 Then
            Chart2.Series.Item("Series1").Points.RemoveAt(0)
        End If

        Chart2.Series.Item("Series2").Points.Add((avarageWatts))

        If Chart2.Series.Item("Series2").Points.Count > 1000 Then
            Chart2.Series.Item("Series2").Points.RemoveAt(0)
        End If


        'watts out
        Chart5.Series.Item("Series1").Points.Add((currentWatts))

        'If Chart5.Series.Item("Series1").Points.Count > 1000 Then
        '    Chart5.Series.Item("Series1").Points.RemoveAt(0)
        'End If

        Chart5.Series.Item("Series2").Points.Add((avarageWatts))

        'If Chart5.Series.Item("Series2").Points.Count > 1000 Then
        'Chart5.Series.Item("Series2").Points.RemoveAt(0)
        'End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        pidSetings.Visible = False
        primairySettings.Visible = False
        secundairySettings.Visible = False
        advancedSettings.Visible = False
        basisSetings.Visible = True
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        basisSetings.Visible = False
        primairySettings.Visible = False
        secundairySettings.Visible = False
        advancedSettings.Visible = False
        pidSetings.Visible = True
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        basisSetings.Visible = False
        pidSetings.Visible = False
        secundairySettings.Visible = False
        advancedSettings.Visible = False
        primairySettings.Visible = True
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        basisSetings.Visible = False
        pidSetings.Visible = False
        primairySettings.Visible = False
        advancedSettings.Visible = False
        secundairySettings.Visible = True
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        basisSetings.Visible = False
        pidSetings.Visible = False
        primairySettings.Visible = False
        secundairySettings.Visible = False
        advancedSettings.Visible = True
    End Sub

End Class
