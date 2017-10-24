
Imports System
Imports System.IO
Imports System.Collections

Public Class Form1
    Inherits System.Windows.Forms.Form

    Dim Counter As Integer
    Dim F As Boolean
    Dim ij As Integer
    Dim oXL As Microsoft.Office.Interop.Excel.Application
    Dim oWB As Microsoft.Office.Interop.Excel.Workbook
    Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
    Dim Manhol_no As Integer
    Dim Ground_Le As Integer
    Dim arrayst1(5) As String
    Dim ExelFileName As String
    Dim SaveFileName As String
    Dim OpenExcel As Boolean
    Dim No As Integer
    Dim No1 As Integer
    Dim a(20000) As Double
    Dim l(20000) As Single
    Dim k(20000) As Single
    Dim z(20000) As Single
    Dim k2(20000) As String
    Dim z2(20000) As String
    Dim gl1(20000) As Single
    Dim X1(20000) As Double
    Dim Y1(20000) As Double
    Dim X2(20000) As Double
    Dim Y2(20000) As Double
    Dim MaxSewerNo As Integer
    Dim si(20000) As Single
    Dim di(20000) As Single
    Dim s(200000) As Single
    Dim q5(20000) As Single
    Dim q6(20000) As Single
    Dim gl2(20000) As Single
    Dim il1(20000) As Single
    Dim il2(20000) As Single
    Dim dis(20000) As Single
    Dim d(20000) As Single
    Dim v(20000) As Single
    Dim my1(20000) As Single
    Dim nn(20000) As Integer
    Dim cl(20000) As Single
    Dim p2(20000) As Single
    Dim N As Integer
    Dim qz(20000) As Single
    Dim sk As Integer
    Dim vz As Single
    Dim ux As Integer
    Dim iin(20000) As Integer
    Dim m_po As Single
    Dim m_lc As Single
    Dim m_mv As Single
    Dim m_ks1 As Single
    Dim m_rf2 As Single
    Dim m_rf3 As Single
    Dim m_rf4 As Single
    Dim m_mi As Single
    Dim m_mx As Single
    Dim shpPoint As String
    Dim shpPoint2 As String
    Dim shpLine As String
    Dim hnd As Integer
    Dim count As Integer
    Dim tt1 As Integer
    Dim tt2 As Integer
    Dim xxx, yyy As Double
    Dim xxxStr As String
    Dim Man(20000) As Double
    Dim mh(20000) As Integer
    Dim Distance, xx2, yy2 As Double
    Dim FilewExt As String
    Dim Lptx(20000) As Double
    Dim Lpty(20000) As Double
    Dim ShpPointName As String
    Dim ShpLineName As String
    Dim xx(20000) As Double
    Dim yy(20000) As Double
    Dim ShpNme As String
    Dim Mah11(20000) As Single
    Dim Mah22(20000) As Single
    Dim MCoverDepth(20000) As Single
    Dim AddDis(20000) As Single
    Dim MinDiam(20000) As Single
    Dim CountNo As Integer
    Dim Sc As Integer
    
    Dim Length_sh As Integer
    Dim NetStr As String
    Dim Man_Str(20000) As String
    Dim Man2_Str(20000) As String
    Dim MaxGnd As Double
    Dim MinGnd As Double
    Dim Tol As Double
    Dim px1Coord(2) As Double
    Dim py1Coord(2) As Double
    Dim px2Coord(2) As Double
    Dim py2Coord(2) As Double
    Dim stst, vtst As Single





    Private Function GetsVal(ByVal vl As Double) As Double
        Dim k, k1 As Integer
        Dim st, st2, st3, st4 As String
        Dim Vn, Vt As Double
        If vl = 0 Then
            GetsVal = 0
            GoTo A100
        End If

        st = Str(vl)

        k = st.IndexOf(".")
        st4 = Mid(st, k + 2, 1)

        k1 = Len(Mid(st, k + 2, Len(st)))
        st3 = Mid(st, k + 1, Len(st))
        st2 = Str(Int(Val(vl)))

        Vn = Val(st3) * 10

        If k1 = 3 Or k1 = 2 And st4 = "0" Then
            Vt = Val(st2) + Vn
            GetsVal = Vt

        Else
            GetsVal = vl
        End If


A100:


    End Function

    Private Sub Form1_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub
     
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Tol = 5.0
        Manhol_no = 2
        Ground_Le = 3
        Length_sh = 1
        If IO.Directory.Exists("c:\Shapes") Then
            Exit Sub
        Else : IO.Directory.CreateDirectory("c:\Shapes")
        End If


    End Sub


   


    Private Sub ExitItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitItem.Click
        If OpenExcel = True Then
            oWB.Save()
            oWB.Close()
            oXL.Quit()
        End If
       
        End
        Form1.ActiveForm.Close()


    End Sub

    Private Sub ReadShp(ByVal shpPoint As String)
        Dim sf As New MapWinGIS.Shapefile
        Dim f As Boolean
        sf.Open(shpPoint)
        Dim st, st2, Text(sf.NumShapes) As String
        Dim i, k As Integer
        ProgressBar1.Maximum = sf.NumShapes
        For i = 0 To sf.NumShapes - 1
            st2 = sf.CellValue(Manhol_no, i)
            k = st2.IndexOf(".")
            If k > 1 Then
                Text(i) = Mid(st2, 1, k)
                st = Text(i)
                f = SearchStr(st, i - 1, Text)
                If f = False Then NetComboBox.Items.Add(st)
            End If
            ProgressBar1.Value = i

        Next

        sf.Close()
        sf = Nothing
    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click

        MyRefresh()

    End Sub
    Private Sub MyRefresh()
        If OpenExcel = True Then
            oWB.Save()
            oWB.Close()
            oXL.Quit()
            OpenExcel = False
        End If
        
        ProgressBar1.Value = 0
        Counter = 0
        NetComboBox.Items.Clear()

    End Sub


    Private Sub PipeConnection2()
        Dim sf As New MapWinGIS.Shapefile
        Dim i As Integer
        Dim Pt As New MapWinGIS.Point
        sf.Open(shpPoint2)
        Dim Man_No(sf.NumShapes + 1) As Double
        Dim x1(sf.NumShapes + 1) As Double
        Dim y1(sf.NumShapes + 1) As Double
        Dim XCoord(sf.NumShapes + 1) As Double
        Dim YCoord(sf.NumShapes + 1) As Double
        Dim GLevel(sf.NumShapes + 1) As Double
        Dim strVal As String
        Dim ln As Integer
        ReDim Man_Str(sf.NumShapes + 1)
        ProgressBar1.Maximum = sf.NumShapes

        For i = 0 To sf.NumShapes - 1
            strVal = GetStr(sf.CellValue(Manhol_no, i))
            Man_Str(i) = strVal
            Man2_Str(i) = strVal

            ln = GetStrlen(strVal)
            If ln = 2 Then strVal = Int(strVal) + GetNewStr(strVal)
            Man(i) = Val(strVal)
            Man_No(i) = Val(strVal)
            Pt = sf.QuickPoint(i, 0)
            XCoord(i) = Pt.x
            YCoord(i) = Pt.y
            xx(i) = XCoord(i)
            yy(i) = YCoord(i)
            GLevel(i) = sf.CellValue(Ground_Le, i)


A10:        ProgressBar1.Value = i
            Application.DoEvents()
        Next

        Sorting2(XCoord, YCoord, Man_No, GLevel, sf.NumShapes - 1)

        NetComboBox.Enabled = True
        sf.Close()
    End Sub

    Private Function GetNewStr(ByVal str As String) As String
        Dim k As Integer
        Dim st As String
        k = str.IndexOf(".")
        st = Mid(str, k + 2, Len(str))
        GetNewStr = "0.0" + st

    End Function



    Private Function GetStrlen(ByVal str As String) As Integer
        Dim k As Integer
        Dim st As String
        k = str.IndexOf(".")
        st = Mid(str, k + 2, Len(str))
        GetStrlen = Len(st)
    End Function


    Private Function GetStr(ByVal str As String) As String
        Dim k As Integer

        k = str.IndexOf(".")

        GetStr = Mid(str, k + 2, Len(str))
    End Function


    Private Sub Check_Man(ByVal inx() As Integer, ByVal k As Integer, ByVal ind As Integer)
        Dim i As Integer
        Dim T1, mMin, Y2 As Single

        Y2 = 10000.0
        For i = 0 To k
            T1 = Man(inx(i))
            If T1 < Y2 Then
                Y2 = T1
                mMin = Y2
                xx2 = xx(inx(i))
                yy2 = yy(inx(i))
            End If
        Next i

        If Man(ind) > mMin Then
            xxx = mMin

        Else
            xxx = 0


        End If


    End Sub
    Private Function PointInPipe(ByVal x1 As Double, ByVal y1 As Double) As Integer
        Dim i As Integer
        Dim pt(2) As Double
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpPoint)
        For i = 0 To sf.NumShapes - 1
            pt = sf.QuickPoints(i, 0)
            If x1 = pt(0) And y1 = pt(1) Then
                PointInPipe = i
                Exit For
            Else : PointInPipe = -1
            End If
        Next
        sf.Close()
    End Function


    Private Sub Sorting2(ByVal XCoord() As Double, ByVal YCoord() As Double, ByVal Man_no() As Double, ByVal Glevel() As Double, ByVal Cnt As Integer)
        Dim i, j As Integer
        Dim T1 As Double
        Dim ptX As Double
        Dim ptY As Double
        Dim pt0 As New MapWinGIS.Point
        Dim sf As New MapWinGIS.Shapefile
        Dim xd, yd As Double
        Dim Tstr As String
        Dim ff As Boolean
        sf.Open(shpLine)

        Dim sXL1 As Microsoft.Office.Interop.Excel.Application
        Dim sWB1 As Microsoft.Office.Interop.Excel.Workbook
        Dim sSheet1 As Microsoft.Office.Interop.Excel.Worksheet
        Dim range1 As Object
        Dim arrayst1(19) As String
        sXL1 = CreateObject("Excel.Application")
        ' Get a new blank workbook.
        sWB1 = sXL1.Workbooks.Add
        sSheet1 = sWB1.ActiveSheet
        range1 = sSheet1.Range("A1", "K1")
        arrayst1(0) = "Area"
        arrayst1(1) = "Length"
        arrayst1(2) = "Man_No1"
        arrayst1(3) = "Man_No2"
        arrayst1(4) = "Ground L"
        arrayst1(5) = "x1"
        arrayst1(6) = "y1"
        arrayst1(7) = "x2"
        arrayst1(8) = "y2"
        arrayst1(9) = "Mh1"
        arrayst1(10) = "Mh2"

        ff = False
        range1.Value2 = arrayst1
        ij = 0


        


        For i = 0 To Cnt
            For j = i + 1 To Cnt
                If Man_no(i) > Man_no(j) Then


                    T1 = Man_no(j)
                    Man_no(j) = Man_no(i)
                    Man_no(i) = T1

                    Tstr = Man_Str(j)
                    Man_Str(j) = Man_Str(i)
                    Man_Str(i) = Tstr
 


                    T1 = XCoord(j)
                    XCoord(j) = XCoord(i)
                    XCoord(i) = T1

                    T1 = YCoord(j)
                    YCoord(j) = YCoord(i)
                    YCoord(i) = T1


                    T1 = Glevel(j)
                    Glevel(j) = Glevel(i)
                    Glevel(i) = T1

                End If
            Next j
            ProgressBar1.Value = i
            Application.DoEvents()
        Next i




        Dim sv, kk As Integer
        sv = 0
        For kk = 0 To Cnt

            If Math.Abs(Int(Man_no(kk)) - Int(Man_no(kk + 1))) >= 1 Then

                For i = sv To kk
                    For j = i + 1 To kk

                        If Man_no(i) < Man_no(j) Then

                            T1 = Man_no(j)
                            Man_no(j) = Man_no(i)
                            Man_no(i) = T1

                            Tstr = Man_Str(j)
                            Man_Str(j) = Man_Str(i)
                            Man_Str(i) = Tstr

                            T1 = XCoord(j)
                            XCoord(j) = XCoord(i)
                            XCoord(i) = T1

                            T1 = YCoord(j)
                            YCoord(j) = YCoord(i)
                            YCoord(i) = T1

                            T1 = Glevel(j)
                            Glevel(j) = Glevel(i)
                            Glevel(i) = T1


                        End If


                        ProgressBar1.Value = j
                        Application.DoEvents()
                    Next j
                    ProgressBar1.Value = i
                    Application.DoEvents()
                Next i
                sv = kk + 1
            End If

            ProgressBar1.Value = kk
            Application.DoEvents()
        Next

        j = 0


 


        For i = 0 To Cnt

            range1 = sSheet1.Range("A" & j + 2, "K" & j + 2)
            xd = Math.Abs(XCoord(i) - XCoord(i + 1))
            yd = Math.Abs(YCoord(i) - YCoord(i + 1))
            Distance = Math.Sqrt(xd * xd + yd * yd)
            arrayst1(0) = "0.5"
            arrayst1(1) = Format(Distance, "##0.0")
            arrayst1(2) = Man_no(i)
            arrayst1(3) = Man_no(i + 1)
            arrayst1(4) = Glevel(i)
            arrayst1(5) = XCoord(i)
            arrayst1(6) = YCoord(i)
            arrayst1(7) = XCoord(i + 1)
            arrayst1(8) = YCoord(i + 1)

            arrayst1(9) = Man_Str(i)
            arrayst1(10) = Man_Str(i + 1)


            'If i = sf.NumShapes + 1 Then Exit For


            If Math.Abs(Int(Man_no(i)) - Int(Man_no(i + 1))) >= 1 Then

                'search what point connect to MAN_NO
                ptX = XCoord(i)
                ptY = YCoord(i)



                'LineCoord(ptX, ptY)
                Get_Intersection2(ptX, ptY)
                '    FindCoord(339535.5095, 4021896.7123)
                arrayst1(3) = xxx
                arrayst1(7) = xx2
                arrayst1(8) = yy2

                If xxx = 0 Then
                    arrayst1(10) = "0.0"
                Else
                    arrayst1(10) = xxxStr

                End If

                xd = Math.Abs(ptX - xx2)
                yd = Math.Abs(ptY - yy2)
                Distance = Math.Sqrt(xd * xd + yd * yd)
                arrayst1(1) = Format(Distance, "##0.0")

            End If
            'If Val(arrayst1(3)) <> 0 Then
            range1.Value2 = arrayst1
            j = j + 1
            'End If
            ProgressBar1.Value = i
            Application.DoEvents()
        Next


        'NetComboBox.Items.Clear()

        sXL1.Visible = True
        sXL1.WindowState = Microsoft.Office.Interop.Excel.XlWindowState.xlNormal
        
        Report(sWB1)

        sXL1 = Nothing
        sWB1 = Nothing
        sSheet1 = Nothing
        range1 = Nothing
    End Sub
    Private Sub Report(ByVal oWB2 As Microsoft.Office.Interop.Excel.Workbook)
        Dim L, i As Integer
        Dim Rng As Microsoft.Office.Interop.Excel.Range
        Dim oSheet2 As Microsoft.Office.Interop.Excel.Worksheet
        oSheet2 = oWB2.ActiveSheet
        Dim st As String = oSheet2.Name

        Rng = oWB2.Worksheets(st).Range("A2").CurrentRegion()
        L = Rng.Rows.Count - 1
        L = L - 1
        Dim Zn(L + 1) As Single
        Dim Err As Integer = 0
        Dim j As Integer
        Dim zErr As New ArrayList
        RichTextBox1.Text += "Network Name         : " + NetComboBox.Text + vbNewLine
        For i = 0 To L
            Zn(i) = Rng.Range("D" & i + 2).Value
            If Zn(i) = 0 Then Err = Err + 1
            If Zn(i) = 0 And Err > 1 Then

                zErr.Add(i + 2)

                j = j + 1

            End If
        Next
        If Err = 1 Then RichTextBox1.Text += "No Errors  " + vbNewLine
        zReport(zErr)
    End Sub

    Private Sub zReport(ByVal zErr As ArrayList)
        Dim i As Integer
        RichTextBox1.Visible = True

        If zErr.Count = 0 Then Exit Sub
        For i = 0 To zErr.Count - 1

            RichTextBox1.Text += "Error in Line No. = " + zErr.Item(i).ToString + vbNewLine
        Next
    End Sub

    Private Sub ReadLineCoord()
        Dim sf As New MapWinGIS.Shapefile
        Dim ff As Boolean
        Dim k As Integer
        Dim pL(4) As Double
        ff = sf.Open(shpLine)
        ReDim px1Coord(sf.NumShapes + 1)
        ReDim py1Coord(sf.NumShapes + 1)
        ReDim px2Coord(sf.NumShapes + 1)
        ReDim py2Coord(sf.NumShapes + 1)
        ProgressBar1.Maximum = sf.NumShapes
        If ff Then
            For k = 0 To sf.NumShapes - 1
                pL = sf.QuickPoints(k, 4)
                px1Coord(k) = pL(0)
                py1Coord(k) = pL(1)
                px2Coord(k) = pL(2)
                py2Coord(k) = pL(3)
                ProgressBar1.Value = k
            Next
        End If

        sf.Close()
    End Sub

    Private Sub FindCoord(ByVal ptx As Double, ByVal pty As Double)
        Dim sf As New MapWinGIS.Shapefile
        Dim ff As Boolean
        Dim k As Integer
        Dim pL(4) As Double
        Dim pt As New MapWinGIS.Point
        Dim f1, f2, f3, f4 As Boolean
        Dim xu1, yu1, xu2, yu2, xl1, yl1, xl2, yl2 As Double
        ff = sf.Open(shpLine)
        f1 = False
        f2 = False
        f3 = False
        f4 = False
        ProgressBar1.Maximum = sf.NumShapes
        If ff = True Then
            For k = 0 To sf.NumShapes - 1

                xu1 = px1Coord(k) + Tol
                xl1 = px1Coord(k) - Tol
                yu1 = py1Coord(k) + Tol
                yl1 = py1Coord(k) - Tol

                xu2 = px2Coord(k) + Tol
                xl2 = px2Coord(k) - Tol
                yu2 = py2Coord(k) + Tol
                yl2 = py2Coord(k) - Tol


                If (ptx >= xl1 And ptx <= xu1) Then f1 = True
                If (pty >= yl1 And pty <= yu1) Then f2 = True
                If f1 = True And f2 = True Then
                    f1 = False
                    f2 = False
                End If

                If (ptx >= xl2 And ptx <= xu2) Then f3 = True
                If (pty >= yl2 And pty <= yu2) Then f4 = True
                If f3 = True And f4 = True Then
                    f1 = False
                    f2 = False
                End If

                ProgressBar1.Value = k
            Next k
        End If

        sf.Close()
    End Sub




    Private Sub Get_Intersection2(ByVal ptx As Double, ByVal pty As Double)
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpLine)
        Dim i, n1, k, Index(10000) As Integer
        Dim Pot(4) As Double
        k = 0
        ProgressBar1.Maximum = sf.NumShapes
        For i = 0 To sf.NumShapes - 1
            Pot = sf.QuickPoints(i, n1)

            If (ptx >= (Pot(0) - Tol) And ptx <= (Pot(0) + Tol)) And (pty >= (Pot(1) - Tol) And pty <= (Pot(1) + Tol)) Then
                Index(k) = PointInPipe2(Pot(2), Pot(3))
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1

            End If
            If (ptx >= (Pot(2) - Tol) And ptx <= (Pot(2) + Tol)) And (pty >= (Pot(3) - Tol) And pty <= (Pot(3) + Tol)) Then
                Index(k) = PointInPipe2(Pot(0), Pot(1))
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1

            End If
AEnd:       ProgressBar1.Value = i
            Application.DoEvents()
        Next i

        Dim ind, Ct As Integer
        ind = PointInPipe2(ptx, pty)
        Ct = k - 1

        Check_Man2(Index, Ct, ind)

        sf.Close()

    End Sub

    Private Function LineCoord(ByVal x As Double, ByVal y As Double) As Boolean
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpLine)
        Dim i, k, Index(10000) As Integer
        Dim P1 As New MapWinGIS.Point
        Dim P2 As New MapWinGIS.Point
        k = 0
        ProgressBar1.Maximum = sf.NumShapes
        For i = 0 To sf.NumShapes - 1
            P1 = sf.QuickPoint(i, 0)
            P2 = sf.QuickPoint(i, 1)

            If x = P1.x And y = P1.y Then

                'MessageBox.Show(P1.x.ToString + "   " + P1.y.ToString + "  " + k.ToString)
                Index(k) = PointInPipe2(P2.x, P2.y)
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1
            End If


            If x = P2.x And x = P2.x Then
                'MessageBox.Show(P2.x.ToString + "   " + P2.y.ToString + "  " + k.ToString)
                Index(k) = PointInPipe2(P1.x, P1.y)
                If Index(k) = -1 Then GoTo AEnd
                k = k + 1
            End If

AEnd:
            ProgressBar1.Value = i
        Next i

 
        Dim ind, Ct As Integer
        ind = PointInPipe2(x, y)
        Ct = k - 1
        Check_Man2(Index, Ct, ind)
        sf.Close()

    End Function

    Private Sub Check_Man2(ByVal inx() As Integer, ByVal k As Integer, ByVal ind As Integer)
        Dim i As Integer
        Dim T1, mMin, Y2 As Double
        Dim Tstr, T5 As String
        Y2 = 10000.0

        For i = 0 To k
            T1 = Man(inx(i))
            Tstr = Man2_Str(inx(i))
            If T1 < Y2 Then
                Y2 = T1
                mMin = Y2
                xx2 = xx(inx(i))
                yy2 = yy(inx(i))
                T5 = Man2_Str(inx(i))

            End If
        Next i
        F = False
        If Man(ind) > mMin Then
            xxx = mMin
            xxxStr = T5
        Else
            xxx = 0

            xxxStr = "0.0"
        End If






    End Sub
    Private Function PointInPipe2(ByVal x1 As Double, ByVal y1 As Double) As Integer
        Dim i As Integer
        Dim pt(2) As Double
        Dim sf As New MapWinGIS.Shapefile
        sf.Open(shpPoint2)
        For i = 0 To sf.NumShapes - 1
            pt = sf.QuickPoints(i, 1)
            If x1 >= pt(0) - Tol And x1 <= pt(0) + Tol And y1 >= pt(1) - Tol And y1 <= pt(1) + Tol Then
                PointInPipe2 = i
                Exit For
            Else : PointInPipe2 = -1
            End If
        Next
        sf.Close()
    End Function


    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub



     

 
  
    Private Function Search2(ByVal zValue As Single, ByVal k As Integer, ByVal ZZZ() As Single) As Boolean
        Dim i As Integer
        For i = 0 To k
            If zValue = ZZZ(i) Then
                Search2 = True
                Exit For
            Else : Search2 = False
            End If
        Next
    End Function





    Private Function SearchStr(ByVal zValue As String, ByVal k As Integer, ByVal TempStr() As String) As Boolean
        Dim i As Integer
        For i = 0 To k
            If zValue = TempStr(i) Then
                SearchStr = True

                Exit Function
            Else : SearchStr = False
            End If
        Next
    End Function


    Private Sub NetComboBox_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles NetComboBox.KeyDown
      

    End Sub

    Private Sub Creat_Shp()

        Dim fld As New MapWinGIS.Field
        Dim shp As New MapWinGIS.Shape
        Dim sf As New MapWinGIS.Shapefile
        Dim i, k, j As Integer
        Dim sfpt As New MapWinGIS.Shapefile

        sfpt.Open(shpPoint)
        Dim StrVal As String
        Dim st2, st As String
        Dim st3(sfpt.NumShapes + 1) As String
        Dim f As Boolean
        Dim GLevel As String
        Dim pt As New MapWinGIS.Point
        j = 0

        sf.CreateNew("C:\Shapes\" + NetStr + "Junction" + ".shp", MapWinGIS.ShpfileType.SHP_POINT)

        fld = New MapWinGIS.Field
        fld.Name = "E"
        fld.Type = MapWinGIS.FieldType.STRING_FIELD
        fld.Width = 5
        sf.EditInsertField(fld, 0)

        fld = New MapWinGIS.Field
        fld.Name = "M"
        fld.Type = MapWinGIS.FieldType.STRING_FIELD
        fld.Width = 5
        sf.EditInsertField(fld, 1)
 
        fld = New MapWinGIS.Field
        fld.Name = "man_no"
        fld.Type = MapWinGIS.FieldType.STRING_FIELD
        fld.Width = 25
        sf.EditInsertField(fld, Manhol_no)

        fld = New MapWinGIS.Field
        fld.Name = "G_level"
        fld.Type = MapWinGIS.FieldType.STRING_FIELD
        fld.Width = 12
        sf.EditInsertField(fld, Ground_Le)


        ProgressBar1.Maximum = sfpt.NumShapes

        For i = 0 To sfpt.NumShapes - 1

            ProgressBar1.Value = i
            Application.DoEvents()
            st2 = UCase(sfpt.CellValue(Manhol_no, i))
            If st2 = "" Then GoTo A10
            k = st2.IndexOf(".")
            st = Mid(st2, 1, k)
            If st = UCase(NetStr) Then
                st3(i) = sfpt.CellValue(Manhol_no, i)
                StrVal = st3(i)

                pt = sfpt.QuickPoint(i, 0)
                GLevel = sfpt.CellValue(Ground_Le, i)

                f = SearchStr(StrVal, i - 1, st3)
                If (f = False) Then
                    shp = New MapWinGIS.Shape
                    shp.Create(MapWinGIS.ShpfileType.SHP_POINT)
                    shp.InsertPoint(pt, 0)
                    sf.EditInsertShape(shp, j)
                    sf.EditCellValue(0, j, "1")
                    sf.EditCellValue(1, j, "1")
                    sf.EditCellValue(2, j, StrVal)
                    sf.EditCellValue(3, j, GLevel)
                    j = j + 1
                End If

            End If
A10:
        Next i
        sf.StopEditingShapes(True, True)
        sfpt.Close()
        sf.Close()
        sfpt = Nothing
        shpPoint2 = "C:\Shapes\" + NetComboBox.Text + "Junction" + ".shp"




    End Sub


 

    Private Sub NetComboBox_TextUpdate(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NetComboBox.TextUpdate
        NetStr = NetComboBox.Text
    End Sub

    Private Sub NetComboBox_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NetComboBox.SelectedIndexChanged
        NetStr = NetComboBox.Text
        Button1.Enabled = True
    End Sub



    Private Sub NetComboBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NetComboBox.TextChanged
        NetStr = NetComboBox.Text
    End Sub


    Private Sub LoadShapeFileToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LoadShapeFileToolStripMenuItem.Click
        Dim i As Integer
        Dim strName(2) As String
        Dim sf As New MapWinGIS.Shapefile
        'OpenFileDialog1.InitialDirectory = "C:\"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.Filter = "Shapefiles: (*.shp)|*.shp"

        MyRefresh()
        If (OpenFileDialog1.ShowDialog = DialogResult.OK) Then
            strName = OpenFileDialog1.FileNames
            i = OpenFileDialog1.FileNames.Length

            Select Case i
                Case 1
                    MessageBox.Show("Must be 2 Layers")

                Case 2
                    sf.Open(strName(0))
                    If sf.ShapefileType = MapWinGIS.ShpfileType.SHP_POINT Then
                        shpPoint = strName(0)
                        shpLine = strName(1)

                        'DrawShpFile(shpPoint)
                        'DrawShpFile(shpLine)
                    Else
                        shpPoint = strName(1)
                        shpLine = strName(0)

                        sf.Close()
                        'DrawShpFile(shpPoint)
                        'DrawShpFile(shpLine)
                    End If

                    ReadShp(shpPoint)
                Case Is > 2
                    MessageBox.Show("Must be 2 Layers")

            End Select



        End If

    End Sub

   
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        NetComboBox.Enabled = False
        NetStr = NetComboBox.Text
        Button1.Enabled = False
        Creat_Shp()
        PipeConnection2()

        If File.Exists("C:\Shapes\" + NetStr + "Junction" + ".shp") Then
            File.Delete("C:\Shapes\" + NetStr + "Junction" + ".shp")
            File.Delete("C:\Shapes\" + NetStr + "Junction" + ".dbf")
            File.Delete("C:\Shapes\" + NetStr + "Junction" + ".shx")

        End If


    End Sub
End Class

