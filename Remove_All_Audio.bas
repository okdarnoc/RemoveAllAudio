Attribute VB_Name = "Remove_All_Audio"
Sub RemoveAllAudio()

    Dim oSlide As Slide
    Dim oShape As shape
    Dim slideNums As String
    Dim slideRange As Variant
    Dim slideIndex As Variant
    Dim slideNumber As Integer
    Dim i As Integer
    
    ' Request user input for the slide numbers
    slideNums = InputBox("Enter slide numbers to remove audio (separated by comma, use dash for range):", "Slide Numbers")
    slideRange = Split(slideNums, ",")
    
    ' Loop through each slide number or range provided
    For Each slideIndex In slideRange
        ' Check if the slideIndex is a range
        If InStr(slideIndex, "-") > 0 Then
            ' Split the range into start and end
            Dim rangeParts As Variant
            rangeParts = Split(slideIndex, "-")
            
            ' Loop through each slide in the range
            For slideNumber = CInt(Trim(rangeParts(0))) To CInt(Trim(rangeParts(1)))
                ' Call the function to remove audio from the slide
                RemoveAudioFromSlide slideNumber
            Next slideNumber
        Else
            ' Remove audio from the individual slide
            RemoveAudioFromSlide CInt(Trim(slideIndex))
        End If
    Next slideIndex
    
    MsgBox "Audio has been removed from the specified slides.", vbInformation

End Sub

Sub RemoveAudioFromSlide(slideNumber As Integer)

    Dim oSlide As Slide
    Dim oShape As shape
    Dim i As Integer
    
    ' Check if the slide number is valid
    If slideNumber > 0 And slideNumber <= ActivePresentation.Slides.Count Then
        Set oSlide = ActivePresentation.Slides(slideNumber)
        
        i = oSlide.Shapes.Count
        
        ' Loop through each shape on the slide
        For i = i To 1 Step -1
            Set oShape = oSlide.Shapes(i)
            
            ' Check if the shape is an audio file
            If oShape.Type = msoMedia Then
                If oShape.MediaType = ppMediaTypeSound Then
                    ' Delete the audio file
                    oShape.Delete
                End If
            End If
        Next i
    End If

End Sub

