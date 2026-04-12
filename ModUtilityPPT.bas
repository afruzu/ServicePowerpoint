' ============================================================
'  ModUtilityPPT.bas
'  PowerPoint VBA Utilities – v1.0
'
'  Scopo:
'    Raccolta di macro per l’editing rapido di presentazioni:
'    - Copia oggetti su tutte le slide
'    - Pulizia animazioni e transizioni
'    - Uniformazione font
'    - Eliminazione slide vuote
'
'  Uso:
'    Importare questo modulo in una presentazione .pptm
'    o in un componente aggiuntivo .ppam.
'
'  Autori:
'    Afruzu + Copilot
' ============================================================

Option Explicit

' ------------------------------------------------------------
' 1. Copia oggetto selezionato in tutte le diapositive
'
' Descrizione:
'   - Richiede che sia selezionata almeno una forma.
'   - Copia la prima forma selezionata.
'   - La incolla in tutte le slide, mantenendo posizione e dimensioni.
' ------------------------------------------------------------
Sub CopiaOggettoInTutteLeDiapositive()
    Dim sld As Slide
    Dim shp As Shape
    Dim copia As Shape

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Seleziona un oggetto prima di eseguire la macro.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    shp.Copy

    For Each sld In ActivePresentation.Slides
        Set copia = sld.Shapes.Paste(1)
        copia.Left = shp.Left
        copia.Top = shp.Top
        copia.Width = shp.Width
        copia.Height = shp.Height
    Next sld

    MsgBox "Oggetto copiato in tutte le diapositive.", vbInformation
End Sub


' ------------------------------------------------------------
' 2. Rimuovi oggetti con un certo nome da tutte le slide
'
' Descrizione:
'   - Chiede il nome dell’oggetto (Shape.Name).
'   - Elimina tutte le forme con quel nome in tutte le slide.
' ------------------------------------------------------------
Sub RimuoviOggettiPerNome()
    Dim nome As String
    Dim sld As Slide
    Dim i As Long
    Dim countDel As Long

    nome = InputBox("Nome dell'oggetto da eliminare:", "Rimuovi oggetti per nome")

    If nome = "" Then Exit Sub

    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(i).Name = nome Then
                sld.Shapes(i).Delete
                countDel = countDel + 1
            End If
        Next i
    Next sld

    MsgBox "Oggetti eliminati: " & countDel, vbInformation
End Sub


' ------------------------------------------------------------
' 3. Rimuovi tutte le animazioni
'
' Descrizione:
'   - Cancella la MainSequence di ogni slide.
' ------------------------------------------------------------
Sub RimuoviTutteLeAnimazioni()
    Dim sld As Slide

    For Each sld In ActivePresentation.Slides
        sld.TimeLine.MainSequence.Delete
    Next sld

    MsgBox "Animazioni rimosse da tutte le diapositive.", vbInformation
End Sub


' ------------------------------------------------------------
' 4. Rimuovi tutte le transizioni
'
' Descrizione:
'   - Imposta l’effetto di transizione a Nessuno per ogni slide.
' ------------------------------------------------------------
Sub RimuoviTutteLeTransizioni()
    Dim sld As Slide

    For Each sld In ActivePresentation.Slides
        sld.SlideShowTransition.EntryEffect = ppEffectNone
    Next sld

    MsgBox "Transizioni rimosse da tutte le diapositive.", vbInformation
End Sub


' ------------------------------------------------------------
' 5. Uniforma font in tutta la presentazione
'
' Descrizione:
'   - Chiede il nome del font.
'   - Applica il font a tutto il testo di tutte le forme.
' ------------------------------------------------------------
Sub UniformaFont()
    Dim fontName As String
    Dim sld As Slide
    Dim shp As Shape

    fontName = InputBox("Inserisci il nome del font da applicare:", "Uniforma font")

    If fontName = "" Then Exit Sub

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Name = fontName
                End If
            End If
        Next shp
    Next sld

    MsgBox "Font """ & fontName & """ applicato a tutta la presentazione.", vbInformation
End Sub


' ------------------------------------------------------------
' 6. Elimina le slide vuote
'
' Descrizione:
'   - Considera “vuota” una slide senza forme.
'   - Elimina tali slide.
' ------------------------------------------------------------
Sub EliminaSlideVuote()
    Dim sld As Slide
    Dim i As Long
    Dim countDel As Long

    For i = ActivePresentation.Slides.Count To 1 Step -1
        Set sld = ActivePresentation.Slides(i)
        If sld.Shapes.Count = 0 Then
            sld.Delete
            countDel = countDel + 1
        End If
    Next i

    MsgBox "Slide vuote eliminate: " & countDel, vbInformation
End Sub
