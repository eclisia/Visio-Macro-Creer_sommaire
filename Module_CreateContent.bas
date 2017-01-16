Attribute VB_Name = "Module_CreateContent"
Sub Create_ContentPage()

    Dim ContentPage As Visio.Page 'shape
    Dim textHyperlink As Visio.Shape
    Dim removeShape As Visio.Shape
    Dim myVSHyperlink As Visio.Hyperlink
    Dim myPageCollection As Visio.Pages
    Dim myPg As Visio.Page
    Dim i As Integer
    Dim pageExist As Boolean

    
    'Initial settings
    i = 0
    pageExiste = False
    

    Set myPageCollection = ActiveDocument.Pages
    
    
    '   Check if the sommaire page is already existing
    For Each myPg In myPageCollection
        If myPg.name = "Sommaire" Then
            pageExist = True
        End If
    Next myPg
    
    '   Test if the sommaire page is already existing
    If pageExist = True Then
        MsgBox "Page de Sommaire existe déjà. La macro va être annulée", vbOKOnly, "Page existante"
        Exit Sub
    End If
    
    Set ContentPage = ActiveDocument.Pages.Add
    ContentPage.name = "Sommaire"

    'Creation of the Hyperlink
    For Each myPg In myPageCollection
        Set textHyperlink = ContentPage.DrawRectangle(0, i * 0.3, 3, i * 0.3 + 0.2)
        textHyperlink.TextStyle = "Normal"
        textHyperlink.LineStyle = "Text Only"
        textHyperlink.FillStyle = "Text Only"
        textHyperlink.Text = myPg.name
        textHyperlink.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,255))"
        Set myVSHyperlink = textHyperlink.Hyperlinks.Add
        myVSHyperlink.name = "Row_" & i
        myVSHyperlink.IsDefaultLink = False
        myVSHyperlink.SubAddress = myPg.name
        Debug.Print myPg.name
        i = i + 1
    Next myPg
    
    'Remove previous hyperlink to the "sommaire" of all page
    For Each myPg In myPageCollection
        If GetShapeByName("Shape_Sommaire", myPg) = True Then   'Check if the shape exist
            Set removeShape = myPg.Shapes.Item("Shape_Sommaire")
            removeShape.Delete
        End If
    Next myPg
    
    'Creation of Hyperlink to the "sommaire" page for each page of the visio document.
    For Each myPg In myPageCollection
        Set textHyperlink = myPg.DrawRectangle(0, 0, 0.2, 6)
        textHyperlink.TextStyle = "Normal"
        textHyperlink.LineStyle = "Text Only"
        textHyperlink.FillStyle = "Text Only"
        textHyperlink.Text = "Retour vers le Sommaire"
        textHyperlink.CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,255))"
        textHyperlink.name = "Shape_Sommaire"
        Set myVSHyperlink = textHyperlink.Hyperlinks.Add
        myVSHyperlink.name = "Row_" & i
        myVSHyperlink.IsDefaultLink = False
        myVSHyperlink.SubAddress = ContentPage.name

    Next myPg
    
    'Move the Sommaire page as the first page of the Document
    ActiveDocument.Pages.Item("Sommaire").Index = 1

    MsgBox ContentPage.name & " has been created."



End Sub


Sub remove_Sommaire()
'Macro to remove only the "hyperlink" which have been added on each page by the CreateContent module

    Dim removeShape As Visio.Shape

    Dim myPageCollection As Visio.Pages
    Dim myPg As Visio.Page

    

    Set myPageCollection = ActiveDocument.Pages
    
    

    
    'Remove previous hyperlink to the "sommaire" of all page
    For Each myPg In myPageCollection
        If GetShapeByName("Shape_Sommaire", myPg) = True Then   'Check if the shape exist
            Set removeShape = myPg.Shapes.Item("Shape_Sommaire")
            removeShape.Delete
        End If
    Next myPg
    

    MsgBox "All 'sommaire' Hyperlink have been removed."


End Sub
    
    
    
Public Function GetShapeByName(name As String, pg As Visio.Page) As Boolean
'This private function permits to check if a shape exist
'If the shape does not exist, then on error is raised and the function return false

Dim ShapeTest As Visio.Shape

On Error GoTo Err
    Set ShapeTest = pg.Shapes(name)   'If the shape does not exist, then on error is raised and the function return false
    GetShapeByName = True
 
    Exit Function
 
Err:
    GetShapeByName = False
    Exit Function
End Function


    
    
    


