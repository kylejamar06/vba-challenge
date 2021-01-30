{\rtf1\ansi\ansicpg1252\cocoartf2513
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fmodern\fcharset0 Courier;}
{\colortbl;\red255\green255\blue255;\red0\green0\blue0;}
{\*\expandedcolortbl;;\cssrgb\c0\c0\c0;}
\margl1440\margr1440\vieww10800\viewh8400\viewkind0
\deftab720
\pard\pardeftab720\partightenfactor0

\f0\fs24 \cf2 \expnd0\expndtw0\kerning0
\outl0\strokewidth0 \strokec2 Sub VBA_challenge():\
  \
Dim WS As Worksheet\
    For Each WS In ActiveWorkbook.Worksheets\
    WS.Activate\
        \
        lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row\
\
        Cells(1, "I").Value = "Ticker"\
        Cells(1, "J").Value = "Yearly Change"\
        Cells(1, "K").Value = "Percent Change"\
        Cells(1, "L").Value = "Total Stock Volume"\
      \
        Dim Open_Price As Double\
        Dim Close_Price As Double\
        Dim Yearly_Change As Double\
        Dim Ticker_Name As String\
        Dim Percent_Change As Double\
        Dim Volume As Double\
        Volume = 0\
        Dim Row As Double\
        Row = 2\
        Dim Column As Integer\
        Column = 1\
        Dim i As Long\
        \
      \
        Open_Price = Cells(2, Column + 2).Value\
      \
        \
        For i = 2 To lastrow\
        \
            If Cells(i + 1, Column).Value <> Cells(i, Column).Value Then\
               \
                Ticker_Name = Cells(i, Column).Value\
                Cells(Row, Column + 8).Value = Ticker_Name\
                \
                Close_Price = Cells(i, Column + 5).Value\
               \
                Yearly_Change = Close_Price - Open_Price\
                Cells(Row, Column + 9).Value = Yearly_Change\
                \
                If (Open_Price = 0 And Close_Price = 0) Then\
                    Percent_Change = 0\
                ElseIf (Open_Price = 0 And Close_Price <> 0) Then\
                    Percent_Change = 1\
                Else\
                    Percent_Change = Yearly_Change / Open_Price\
                    Cells(Row, Column + 10).Value = Percent_Change\
                    Cells(Row, Column + 10).NumberFormat = "0.00%"\
                End If\
                \
                Volume = Volume + Cells(i, Column + 6).Value\
                Cells(Row, Column + 11).Value = Volume\
                \
                Row = Row + 1\
               \
                Open_Price = Cells(i + 1, Column + 2)\
                '\
                Volume = 0\
           \
            Else\
                Volume = Volume + Cells(i, Column + 6).Value\
            End If\
        Next i\
        \
        \
        YCLastRow = WS.Cells(Rows.Count, Column + 8).End(xlUp).Row\
       \
        For j = 2 To YCLastRow\
            If (Cells(j, Column + 9).Value > 0 Or Cells(j, Column + 9).Value = 0) Then\
                Cells(j, Column + 9).Interior.ColorIndex = 4\
            ElseIf Cells(j, Column + 9).Value < 0 Then\
                Cells(j, Column + 9).Interior.ColorIndex = 3\
            End If\
        Next j\
        \
    Next WS\
        \
End Sub\
}