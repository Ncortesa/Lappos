'''''''''''''''''''''''''''
'LAPPOS
'Author: N.Cortesão
'Version : Alpha 0.1
'12/04/2016
'''''''''''''''''''''''''''

'''''''''''''''''
'Name:LetterMOD
'Description:This function allow the modification of the font type
'''''''''''''''''
Sub LetterMOD() ''''''''''''''''''
     Qtslidestotal = ActivePresentation.Slides.Count ' How many slides the presentation have
	 
    ' TipoLetra = ActivePresentation.Slides("Appos1").Shapes("MudaLetra").TextFrame.TextRange.Font.Name   
	
     For k = 3 To Qtslidestotal ' The first 2 slides belong to appos
        QtShapesTotal = ActivePresentation.Slides(k).Shapes.Count     ' Number of shapes per slide
            For i = 1 To QtShapesTotal                                ' Run all shapes in the slide to perform the font modification
            
            On Error GoTo ErrorHandler							      ' Some shapes in VBA don't have the text range feature so we need to try and catch errors
                 ActivePresentation.Slides(k).Shapes(i).TextFrame.TextRange.Font.Name = "Arial"   ' Change the font of the shape      
                 
ErrorHandler:  ' this handler allow the catch of bugs
   Resume Next
              
            Next i
    Next k
End Sub  ' End of LetterMOD
'+++++++++++++++++++++++++++++++++'
'''''''''''''''''''''''''''''''''''
