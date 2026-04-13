' ============================================================
'  ModUtilityPPT.bas
'  PowerPoint VBA Utilities – v1.1
'
'  Scopo:
'    Raccolta di macro per l’editing rapido di presentazioni:
'    - Copia oggetti su tutte le slide
'    - Pulizia animazioni e transizioni
'    - Uniformazione font, colori testo, dimensioni font
'    - Formattazione testo (grassetto, corsivo, etc.)
'    - Gestione numeri slide, layout, rinominazione
'    - Eliminazione slide vuote
'    - Cerca e sostituisci testo
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


' ------------------------------------------------------------
' 7. Cerca e sostituisci testo in tutte le slide
'
' Descrizione:
'   - Chiede il testo da cercare e quello di sostituzione.
'   - Sostituisce in tutto il testo delle forme di tutte le slide.
' ------------------------------------------------------------
Sub CercaSostituisciTesto()
    Dim findText As String
    Dim replaceText As String
    Dim sld As Slide
    Dim shp As Shape
    Dim countRepl As Long

    findText = InputBox("Testo da cercare:", "Cerca e sostituisci")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Testo di sostituzione:", "Cerca e sostituisci")
    If replaceText = "" Then Exit Sub

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    countRepl = countRepl + shp.TextFrame.TextRange.Replace(findText, replaceText)
                End If
            End If
        Next shp
    Next sld

    MsgBox "Sostituzioni effettuate: " & countRepl, vbInformation
End Sub


' ------------------------------------------------------------
' 8. Uniforma colore testo in tutta la presentazione
'
' Descrizione:
'   - Chiede il colore RGB come R,G,B (es. 255,0,0 per rosso).
'   - Applica il colore a tutto il testo di tutte le forme.
' ------------------------------------------------------------
Sub UniformaColoreTesto()
    Dim colorInput As String
    Dim rgbParts() As String
    Dim r As Integer, g As Integer, b As Integer
    Dim sld As Slide
    Dim shp As Shape

    colorInput = InputBox("Inserisci colore RGB come R,G,B (es. 0,0,0 per nero):", "Uniforma colore testo")
    If colorInput = "" Then Exit Sub

    rgbParts = Split(colorInput, ",")
    If UBound(rgbParts) <> 2 Then
        MsgBox "Formato non valido. Usa R,G,B.", vbExclamation
        Exit Sub
    End If

    r = Val(Trim(rgbParts(0)))
    g = Val(Trim(rgbParts(1)))
    b = Val(Trim(rgbParts(2)))

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Color.RGB = RGB(r, g, b)
                End If
            End If
        Next shp
    Next sld

    MsgBox "Colore applicato a tutto il testo.", vbInformation
End Sub


' ------------------------------------------------------------
' 9. Imposta dimensione font in tutta la presentazione
'
' Descrizione:
'   - Chiede la dimensione del font.
'   - Applica la dimensione a tutto il testo di tutte le forme.
' ------------------------------------------------------------
Sub ImpostaDimensioneFont()
    Dim fontSize As Single
    Dim sld As Slide
    Dim shp As Shape

    fontSize = Val(InputBox("Inserisci dimensione font (es. 24):", "Imposta dimensione font"))
    If fontSize <= 0 Then Exit Sub

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    shp.TextFrame.TextRange.Font.Size = fontSize
                End If
            End If
        Next shp
    Next sld

    MsgBox "Dimensione font " & fontSize & " applicata a tutto il testo.", vbInformation
End Sub


' ------------------------------------------------------------
' 10. Applica formattazione testo in tutta la presentazione
'
' Descrizione:
'   - Chiede il tipo di formattazione: bold, italic, underline, normal.
'   - Applica a tutto il testo di tutte le forme.
' ------------------------------------------------------------
Sub FormattaTesto()
    Dim formatType As String
    Dim sld As Slide
    Dim shp As Shape

    formatType = LCase(Trim(InputBox("Tipo di formattazione (bold, italic, underline, normal):", "Formattazione testo")))
    If formatType = "" Then Exit Sub

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                If shp.TextFrame.HasText Then
                    Select Case formatType
                        Case "bold"
                            shp.TextFrame.TextRange.Font.Bold = msoTrue
                        Case "italic"
                            shp.TextFrame.TextRange.Font.Italic = msoTrue
                        Case "underline"
                            shp.TextFrame.TextRange.Font.Underline = msoTrue
                        Case "normal"
                            shp.TextFrame.TextRange.Font.Bold = msoFalse
                            shp.TextFrame.TextRange.Font.Italic = msoFalse
                            shp.TextFrame.TextRange.Font.Underline = msoFalse
                        Case Else
                            MsgBox "Tipo non valido.", vbExclamation
                            Exit Sub
                    End Select
                End If
            End If
        Next shp
    Next sld

    MsgBox "Formattazione '" & formatType & "' applicata a tutto il testo.", vbInformation
End Sub


' ------------------------------------------------------------
' 11. Aggiungi numeri di pagina a tutte le slide
'
' Descrizione:
'   - Aggiunge il numero di pagina nel footer di ogni slide.
' ------------------------------------------------------------
Sub AggiungiNumeriSlide()
    Dim sld As Slide

    For Each sld In ActivePresentation.Slides
        sld.HeadersFooters.SlideNumber.Visible = msoTrue
        sld.HeadersFooters.Footer.Text = "&[SlideNumber]"
    Next sld

    MsgBox "Numeri di pagina aggiunti a tutte le slide.", vbInformation
End Sub


' ------------------------------------------------------------
' 12. Resetta layout delle slide
'
' Descrizione:
'   - Chiede il nome del layout da applicare.
'   - Applica il layout a tutte le slide.
' ------------------------------------------------------------
Sub ResettaLayoutSlide()
    Dim layoutName As String
    Dim layout As PpSlideLayout
    Dim sld As Slide

    layoutName = InputBox("Nome del layout da applicare (es. Title Slide, Blank):", "Resetta layout")
    If layoutName = "" Then Exit Sub

    ' Trova il layout corrispondente
    Select Case LCase(layoutName)
        Case "title slide"
            layout = ppLayoutTitle
        Case "title and content"
            layout = ppLayoutTitleAndContent
        Case "blank"
            layout = ppLayoutBlank
        Case "two content"
            layout = ppLayoutTwoContent
        Case Else
            MsgBox "Layout non riconosciuto.", vbExclamation
            Exit Sub
    End Select

    For Each sld In ActivePresentation.Slides
        sld.Layout = layout
    Next sld

    MsgBox "Layout '" & layoutName & "' applicato a tutte le slide.", vbInformation
End Sub


' ------------------------------------------------------------
' 13. Rinomina slide in batch
'
' Descrizione:
'   - Rinomina ogni slide come "Slide_X" dove X è il numero.
' ------------------------------------------------------------
Sub RinominaSlide()
    Dim sld As Slide

    For Each sld In ActivePresentation.Slides
        sld.Name = "Slide_" & sld.SlideIndex
    Next sld

    MsgBox "Slide rinominate.", vbInformation
End Sub

' ------------------------------------------------------------
' 14. Calcola area bianca disponibile nella slide
'
' Descrizione:
'   - Chiede i margini occupati da grafica/loghi (sinistro, destro, superiore, inferiore in punti).
'   - Calcola e mostra l'area libera centrale per testo e immagini.
'   - Restituisce le coordinate in variabili globali per uso futuro.
' ------------------------------------------------------------
Public AreaLiberaLeft As Single
Public AreaLiberaTop As Single
Public AreaLiberaWidth As Single
Public AreaLiberaHeight As Single

Sub CalcolaAreaBiancaDisponibile()
    Dim margineSinistro As Single
    Dim margineDestro As Single
    Dim margineSuperiore As Single
    Dim margineInferiore As Single
    Dim slideWidth As Single
    Dim slideHeight As Single

    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight

    margineSinistro = Val(InputBox("Margine sinistro occupato (punti):", "Calcola area bianca", "50"))
    If margineSinistro < 0 Then margineSinistro = 0

    margineDestro = Val(InputBox("Margine destro occupato (punti):", "Calcola area bianca", "50"))
    If margineDestro < 0 Then margineDestro = 0

    margineSuperiore = Val(InputBox("Margine superiore occupato (punti):", "Calcola area bianca", "50"))
    If margineSuperiore < 0 Then margineSuperiore = 0

    margineInferiore = Val(InputBox("Margine inferiore occupato (punti):", "Calcola area bianca", "50"))
    If margineInferiore < 0 Then margineInferiore = 0

    AreaLiberaLeft = margineSinistro
    AreaLiberaTop = margineSuperiore
    AreaLiberaWidth = slideWidth - margineSinistro - margineDestro
    AreaLiberaHeight = slideHeight - margineSuperiore - margineInferiore

    If AreaLiberaWidth <= 0 Or AreaLiberaHeight <= 0 Then
        MsgBox "Margini troppo grandi! Area libera nulla.", vbExclamation
        Exit Sub
    End If

    MsgBox "Area libera calcolata:" & vbCrLf & _
           "Left: " & AreaLiberaLeft & " pt" & vbCrLf & _
           "Top: " & AreaLiberaTop & " pt" & vbCrLf & _
           "Width: " & AreaLiberaWidth & " pt" & vbCrLf & _
           "Height: " & AreaLiberaHeight & " pt", vbInformation
End Sub

' ------------------------------------------------------------
' 15. Ridimensiona testo per contenere nell'area libera
'
' Descrizione:
'   - Usa l'area libera calcolata da CalcolaAreaBiancaDisponibile.
'   - Ridimensiona le caselle di testo per farle stare entro l'area libera.
' ------------------------------------------------------------
Sub RidimensionaTestoPerAreaLibera()
    Dim sld As Slide
    Dim shp As Shape
    Dim scaleX As Single
    Dim scaleY As Single
    Dim scale As Single

    ' Verifica se l'area libera è stata calcolata
    If AreaLiberaWidth <= 0 Or AreaLiberaHeight <= 0 Then
        MsgBox "Prima calcola l'area libera con CalcolaAreaBiancaDisponibile.", vbExclamation
        Exit Sub
    End If

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.HasTextFrame Then
                ' Ridimensiona la forma per contenere nell'area libera
                scaleX = AreaLiberaWidth / shp.Width
                scaleY = AreaLiberaHeight / shp.Height
                scale = IIf(scaleX < scaleY, scaleX, scaleY)
                If scale < 1 Then
                    shp.Width = shp.Width * scale
                    shp.Height = shp.Height * scale
                End If
                ' Posiziona nell'area libera
                If shp.Left < AreaLiberaLeft Then shp.Left = AreaLiberaLeft
                If shp.Top < AreaLiberaTop Then shp.Top = AreaLiberaTop
                If shp.Left + shp.Width > AreaLiberaLeft + AreaLiberaWidth Then shp.Left = AreaLiberaLeft + AreaLiberaWidth - shp.Width
                If shp.Top + shp.Height > AreaLiberaTop + AreaLiberaHeight Then shp.Top = AreaLiberaTop + AreaLiberaHeight - shp.Height
            End If
        Next shp
    Next sld

    MsgBox "Testo ridimensionato per contenere nell'area libera.", vbInformation
End Sub
