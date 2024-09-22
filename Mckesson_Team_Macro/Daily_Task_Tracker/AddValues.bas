Attribute VB_Name = "AddValues"
Sub Addnames()

Dim Todaysdate As String
Todaysdate = Format(Date, "dd MMMM yyyy")
Dim Sheetcount As Integer
Sheetcount = Worksheets.Count
Sheetcount = Sheetcount - 1
Sheets(Sheetcount).Activate

'Adding Headers
[C180].Value = "Names"
[D180].Value = "Count Completed"

'Adding Names
[C181].Value = "Amit Kumar Jaiswal"
[C182].Value = "Deepak Vishwakarma"
[C183].Value = "Deepakshi Sharma"
[C184].Value = "Justin John Jacob"
[C185].Value = "Naved Afzal"
[C186].Value = "Nupur Aggarwal"
[C187].Value = "Hari Om Wadhera"
[C188].Value = "Rahul Gupta"
[C189].Value = "Ranadip Ghosh"
[C190].Value = "Jasmeet Kaur"

'Adding attributes to the coulumns

Range("c180:D190").Borders.Weight = 2
Range("c180:D190").Borders.LineStyle = xlContinuous
Range("c180:D180").Font.Bold = True
Range("c180:D180").Interior.Color = RGB(248, 203, 173)
Range("c181:c190").Interior.Color = RGB(226, 239, 218)

'Adding Values
    'Deepak Score
        Dim DeepakScore As Integer
        Dim DeepakTimeMin As Integer
        Dim DeepakTimeHr As Integer
        DeepakScore = [D17].Value + [D18].Value + [D19].Value + [D21].Value + [D22].Value + [D23].Value + [D24].Value + [D25].Value + [D26].Value
        
        [D182].Value = DeepakScore
        
       'Deepakshi Score
        Dim DeepakkshiScore As Integer
        Dim DeepakshiTimeMin As Integer
        Dim DeepakshiTimeHr As Integer
        DeepakshiScore = [D30].Value + [D31].Value + [D32].Value + [D34].Value + [D35].Value + [D36].Value + [D37].Value + [D38].Value + [D39].Value
        
        [D183].Value = DeepakshiScore

       'Justin Score
        Dim JustinScore As Integer
        Dim JustinTimeMin As Integer
        Dim JustinTimeHr As Integer
        JustinScore = [D43].Value + [D44].Value + [D45].Value + [D47].Value + [D48].Value + [D49].Value + [D50].Value + [D51].Value + [D52].Value

        [D184].Value = JustinScore

        'Nupur Score
        Dim NupurshiScore As Integer
        Dim NupurTimeMin As Integer
        Dim NupurTimeHr As Integer
        NupurScore = [D69].Value + [D70].Value + [D71].Value + [D73].Value + [D74].Value + [D75].Value + [D76].Value + [D77].Value + [D78].Value

        [D186].Value = NupurScore

        'Naved Score
        Dim NavedshiScore As Integer
        Dim NavedTimeMin As Integer
        Dim NavedTimeHr As Integer
        NavedScore = [D56].Value + [D57].Value + [D58].Value + [D60].Value + [D61].Value + [D62].Value + [D63].Value + [D64].Value + [D65].Value
        
        [D185].Value = NavedScore


        'Ranadip Score
        Dim RanadipshiScore As Integer
        Dim RanadipTimeMin As Integer
        Dim RanadipTimeHr As Integer
        RanadipScore = [D120].Value + [D121].Value + [D122].Value + [D124].Value + [D125].Value + [D126].Value + [D127].Value + [D128].Value + [D129].Value + [D130].Value + [D131].Value + [D132].Value + [D134].Value + [D135].Value + [D136].Value + [D137].Value + [D138].Value + [D139].Value
        [D189].Value = RanadipScore

        'HariOm Score
        Dim HariOmScore As Integer
        Dim HariOmTimeMin As Integer
        Dim HariOmTimeHr As Integer
        HariOmScore = [D83].Value + [D84].Value + [D85].Value + [D87].Value + [D88].Value + [D89].Value + [D90].Value + [D91].Value + [D92].Value

        [D187].Value = HariOmScore
        
        'Rahul Score
        Dim RahulScore As Integer
        Dim RahulTimeMin As Integer
        Dim RahulTimeHr As Integer
        
        RahulScore = [D96].Value + [D97].Value + [D98].Value + [D100].Value + [D101].Value + [D102].Value + [D103].Value + [D104].Value + [D105].Value + [D107].Value + [D108].Value + [D109].Value + [D111].Value + [D112].Value + [D113].Value + [D114].Value + [D115].Value + [D116].Value
        [D188].Value = RahulScore
        
        'Jasmeet Score
        Dim JasmeetScore As Integer
        Dim JasmeetTimeMin As Integer
        Dim JasmeetTimeHr As Integer
        JasmeetScore = [D143].Value + [D144].Value + [D145].Value + [D147].Value + [D148].Value + [D149].Value + [D150].Value + [D151].Value + [D152].Value
        [D190].Value = JasmeetScore
        
        'Amit Score
        Dim AmitScore As Integer
        Dim AmitTimeMin As Integer
        Dim AmitTimeHr As Integer
        AmitScore = [D4].Value + [D5].Value + [D6].Value + [D9].Value + [D10].Value + [D11].Value + [D12].Value + [D13].Value + [D8].Value
        
        [D181].Value = AmitScore
        

    Dim msg As Integer
    msg = MsgBox("Completed 2 of 2 Macro", vbInformation, "Done")
   'Worksheets("Sample").Visible = False

End Sub

