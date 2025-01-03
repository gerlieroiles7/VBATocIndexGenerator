Attribute VB_Name = "Module1"
Sub CreateToCAndIndex()
    Dim originalDoc As Document
    Dim newDoc As Document
    Dim toc As TableOfContents
    Dim rng As Range
    Dim paraCount As Integer
    Dim newFilePath As String

    ' Set the original document as the active document
    Set originalDoc = ActiveDocument

    ' Define the new file path
    newFilePath = "C:\Users\gerlie\Documents\Projects\MacroVBA\AutomatedToCAndIndexBuilder\DocumentWithToCAndIndex.docx"

    ' Ensure the directory exists
    On Error Resume Next
    MkDir "C:\Users\gerlie\Documents\Projects\MacroVBA\AutomatedToCAndIndexBuilder"
    On Error GoTo 0

    ' Save a copy of the original document
    originalDoc.SaveAs2 FileName:=newFilePath, FileFormat:=wdFormatXMLDocument
    Set newDoc = Documents.Open(newFilePath)

    ' Count paragraphs in the document
    paraCount = newDoc.Paragraphs.Count

    ' Generate text if there are insufficient paragraphs
    With newDoc.Content
        If paraCount < 6 Then
            .InsertAfter "Title of the Document" & vbCrLf
            .InsertAfter "Chapter 1 - Introduction to Automation" & vbCrLf
            .InsertAfter "Automation is transforming industries." & vbCrLf
            .InsertAfter "Chapter 2 - Implementation of Automation Tools" & vbCrLf
            .InsertAfter "Automation requires careful planning and strategy." & vbCrLf
            .InsertAfter "Chapter 3 - Benefits and Challenges" & vbCrLf
            .InsertAfter "The challenges of automation can include resistance and implementation hurdles." & vbCrLf
        End If
    End With

    ' Refresh paragraph count
    paraCount = newDoc.Paragraphs.Count

    ' Mark entries for the index
    On Error Resume Next
    If paraCount >= 2 Then
        Set rng = newDoc.Paragraphs(2).Range
        newDoc.Indexes.MarkEntry Range:=rng, Entry:="Automation"
    Else
        MsgBox "Paragraph 2 does not exist.", vbExclamation
    End If

    If paraCount >= 4 Then
        Set rng = newDoc.Paragraphs(4).Range
        newDoc.Indexes.MarkEntry Range:=rng, Entry:="Implementation"
    Else
        MsgBox "Paragraph 4 does not exist.", vbExclamation
    End If

    If paraCount >= 6 Then
        Set rng = newDoc.Paragraphs(6).Range
        newDoc.Indexes.MarkEntry Range:=rng, Entry:="Benefits"
    Else
        MsgBox "Paragraph 6 does not exist.", vbExclamation
    End If
    On Error GoTo 0

    ' Add Table of Contents
    Set toc = newDoc.TablesOfContents.Add(Range:=newDoc.Range(0, 0), _
        RightAlignPageNumbers:=True, UseHeadingStyles:=True, _
        UpperHeadingLevel:=1, LowerHeadingLevel:=3)

    ' Insert index at the end
    newDoc.Content.InsertAfter vbCrLf & "Index" & vbCrLf
    newDoc.Indexes.Add Range:=newDoc.Paragraphs(newDoc.Paragraphs.Count).Range

    ' Save the new document
    newDoc.Save

    ' Notify success
    MsgBox "Table of Contents and Index created in a new file!", vbInformation
End Sub




