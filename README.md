# PowerPoint Audio Remover

This VBA project is a PowerPoint audio remover, designed to remove audio files from specific slides within a PowerPoint presentation.

## Description

The main subroutine `RemoveAllAudio` prompts the user to input slide numbers from which they wish to remove audio. The user can input single slide numbers separated by commas or ranges specified by a dash. After all slides specified have been processed, a message box is displayed, indicating that the audio has been removed from the specified slides.

The second subroutine `RemoveAudioFromSlide` takes a slide number as an input, checks if the slide number is valid, and if valid, it goes through all the shapes on the slide. If it encounters a shape that is an audio file, it deletes the shape.

This tool can be extremely useful if you want to quickly remove all audio elements from certain slides in a PowerPoint presentation, especially for large presentations where doing so manually could be time-consuming.

## Code

```vba
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
