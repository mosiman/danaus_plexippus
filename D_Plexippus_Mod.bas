Attribute VB_Name = "D_Plexippus_Mod"
Global YearCount As Integer
Global PopulationCount As Double
Global Temperature As Single

Global ImgCount As Integer
Global ImgIndices As Integer
Option Explicit

'Public Sub CreateNewImage(TargetForm As Form)
    'ImgCount = ImgCount + 1
    'Dim NewPicBox(1 To 1000) As Image
    
    
    'TargetForm.NewPicBox.Picture = App.Path & "\dplexippus.jpg"
    'TargetForm.NewPicBox.Visible = True
    'TargetForm.NewPicBox.Width = 50
    'TargetForm.NewPicBox.Height = 50
    'TargetForm.NewPicBox.Left = 10
    
    
'End Sub

Public Sub UpdateForm(ByVal YearCount As Integer, ByVal PopulationCount As Long, ByVal Temperature As Single)
'Update the form
'Population/temperature/year and butterfly pic updates
    Const E = 2.71828
    
    
    'PopulationCount = 44717632190.974 - (22168383.74026 * YearCount)
    'PopulationCount = Str$((1 * 10 ^ 94) * (E ^ (-0.098 * YearCount)))
    PopulationCount = (1 * 10 ^ 94) * (E ^ (-0.098 * YearCount))

    frmSimMain.lblYear.Caption = "Year: " & Str$(YearCount)
    'frmSimMain.lblTemperature.Caption = "Temperature: " & Str$(Temperature) & " C"
    frmSimMain.lblPopulation.Caption = "Population: " & Str$(Int(PopulationCount))
    ButterflyPics PopulationCount
    
    If PopulationCount = 0 Then
        frmSimMain.YearPlusOne.Enabled = False
    ElseIf PopulationCount > 0 Then
        frmSimMain.YearPlusOne.Enabled = True
    End If
    
    If YearCount = 1990 Then
        frmSimMain.YearMinusOne.Enabled = False
    Else
        frmSimMain.YearMinusOne.Enabled = True
    End If
End Sub

Public Sub ButterflyPics(ByVal PopCount As Double)
    Const BFLY = 10000000
    Const HIGHX = 8775
    Const LOWX = 1800
    Const HIGHY = 6135
    Const LOWY = 240
    'X range between 1800 and 8775 (6975)
    'Y range between 240 and 6135 (5895)
    
    Dim K As Integer, X As Integer
    
    ImgCount = PopCount \ BFLY
    
    'do while imgcount<>imgindices
    If ImgCount <> ImgIndices Then
    
        If ImgIndices < ImgCount Then
        'increase butterfly pictures
            ImgIndices = ImgIndices + 1
            
            For K = ImgIndices To ImgCount
                Load frmSimMain.imgDPlexippus(K)
                frmSimMain.imgDPlexippus(K).Picture = frmSimMain.imgDPlexippus(0).Picture
                frmSimMain.imgDPlexippus(K).Height = frmSimMain.imgDPlexippus(0).Width
                frmSimMain.imgDPlexippus(K).Width = frmSimMain.imgDPlexippus(0).Width
                'rand position within range
                frmSimMain.imgDPlexippus(K).Top = Int(Rnd * (HIGHY - LOWY + 1) + LOWY) - frmSimMain.imgDPlexippus(K).Height
                frmSimMain.imgDPlexippus(K).Left = Int(Rnd * (HIGHX - frmSimMain.imgDPlexippus(K).Width - LOWX + 1) + LOWX)
                
                If frmSimMain.chkButterfly.Value = 0 Then
                    frmSimMain.imgDPlexippus(K).Visible = True
                End If
                'bring bfly in front of rest
                frmSimMain.imgDPlexippus(K).ZOrder vbBringToFront
            Next K
            ImgIndices = ImgCount
        Else
        'Decrease butterfly pictures
            
            For K = ImgIndices To (ImgCount + 1) Step -1
                If K > 1 Then
                    Unload frmSimMain.imgDPlexippus(K)
                End If
            Next K
            
            ImgIndices = ImgCount
            

        End If
    End If
    
    If ImgCount = 0 And PopCount = 0 Then
        Unload frmSimMain.imgDPlexippus(1)
    End If
    'loop
End Sub

Public Sub NextImg(imgtext As Object, imgImage As Object, NumImg As Integer, currimg As Integer, NextPrevious As Boolean)
'If Next is true, currimg = currimg + 1 else false, currimg --
    Dim K As Integer
    
    If NextPrevious = True Then
        currimg = currimg + 1
    Else
        currimg = currimg - 1
    End If
    
    For K = 0 To NumImg
        imgtext(K).Visible = False
        imgImage(K).Visible = False
    Next K

    imgtext(currimg).Visible = True
    imgImage(currimg).Visible = True
    

    
    
End Sub
