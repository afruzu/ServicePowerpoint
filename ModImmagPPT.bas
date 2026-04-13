' ============================================================
'  ModImmagPPT.bas
'  PowerPoint VBA Utilities – Gestione Immagini v1.2
'
'  Scopo:
'    Raccolta di macro per la gestione delle immagini in presentazioni PowerPoint:
'    - Ridimensionamento immagini
'    - Sostituzione immagini
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
' 1. Ridimensiona tutte le immagini nella presentazione
'
' Descrizione:
'   - Chiede la larghezza e altezza desiderate.
'   - Ridimensiona tutte le immagini (Shapes di tipo msoPicture) mantenendo le proporzioni se necessario.
' ------------------------------------------------------------
Sub RidimensionaImmagini()
    Dim larghezza As Single
    Dim altezza As Single
    Dim sld As Slide
    Dim shp As Shape
    Dim risposta As String

    larghezza = Val(InputBox("Inserisci la larghezza desiderata (in punti):", "Ridimensiona immagini"))
    If larghezza <= 0 Then Exit Sub

    altezza = Val(InputBox("Inserisci l'altezza desiderata (in punti):", "Ridimensiona immagini"))
    If altezza <= 0 Then Exit Sub

    risposta = MsgBox("Mantenere le proporzioni? (Sì per mantenere, No per forzare dimensioni)", vbYesNo, "Ridimensiona immagini")
    Dim mantieniProporzioni As Boolean
    mantieniProporzioni = (risposta = vbYes)

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                If mantieniProporzioni Then
                    ' Calcola il rapporto per mantenere proporzioni
                    Dim rapporto As Single
                    rapporto = shp.Width / shp.Height
                    If larghezza / altezza > rapporto Then
                        shp.Width = altezza * rapporto
                        shp.Height = altezza
                    Else
                        shp.Height = larghezza / rapporto
                        shp.Width = larghezza
                    End If
                Else
                    shp.Width = larghezza
                    shp.Height = altezza
                End If
            End If
        Next shp
    Next sld

    MsgBox "Immagini ridimensionate.", vbInformation
End Sub

' ------------------------------------------------------------
' 2. Sostituisci immagine selezionata
'
' Descrizione:
'   - Richiede che sia selezionata un'immagine.
'   - Apre una finestra di dialogo per selezionare un nuovo file immagine.
'   - Sostituisce l'immagine selezionata con quella nuova, mantenendo posizione e dimensioni.
' ------------------------------------------------------------
Sub SostituisciImmagine()
    Dim shp As Shape
    Dim filePath As String
    Dim fd As FileDialog

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Seleziona un'immagine prima di eseguire la macro.", vbExclamation
        Exit Sub
    End If

    Set shp = ActiveWindow.Selection.ShapeRange(1)
    If shp.Type <> msoPicture Then
        MsgBox "L'oggetto selezionato non è un'immagine.", vbExclamation
        Exit Sub
    End If

    ' Apri finestra di dialogo per selezionare file
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Seleziona nuova immagine"
        .Filters.Add "Immagini", "*.jpg;*.jpeg;*.png;*.gif;*.bmp;*.tiff"
        .AllowMultiSelect = False
        If .Show = -1 Then
            filePath = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With

    ' Sostituisci l'immagine
    shp.Fill.UserPicture filePath

    MsgBox "Immagine sostituita.", vbInformation
End Sub

' ------------------------------------------------------------
' 3. Allinea immagini al centro della slide
'
' Descrizione:
'   - Centra orizzontalmente e verticalmente tutte le immagini nella presentazione.
' ------------------------------------------------------------
Sub AllineaImmaginiCentro()
    Dim sld As Slide
    Dim shp As Shape

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                shp.Left = (ActivePresentation.PageSetup.SlideWidth - shp.Width) / 2
                shp.Top = (ActivePresentation.PageSetup.SlideHeight - shp.Height) / 2
            End If
        Next shp
    Next sld

    MsgBox "Immagini allineate al centro.", vbInformation
End Sub

' ------------------------------------------------------------
' 4. Rimuovi tutte le immagini dalla presentazione
'
' Descrizione:
'   - Elimina tutte le immagini (Shapes di tipo msoPicture) da tutte le slide.
' ------------------------------------------------------------
Sub RimuoviTutteLeImmagini()
    Dim sld As Slide
    Dim i As Long
    Dim countDel As Long

    For Each sld In ActivePresentation.Slides
        For i = sld.Shapes.Count To 1 Step -1
            If sld.Shapes(i).Type = msoPicture Then
                sld.Shapes(i).Delete
                countDel = countDel + 1
            End If
        Next i
    Next sld

    MsgBox "Immagini eliminate: " & countDel, vbInformation
End Sub

' ------------------------------------------------------------
' 5. Ridimensiona immagini per contenere nella slide
'
' Descrizione:
'   - Ridimensiona tutte le immagini per farle stare entro le dimensioni della slide,
'     mantenendo le proporzioni originali.
' ------------------------------------------------------------
Sub RidimensionaImmaginiPerSlide()
    Dim sld As Slide
    Dim shp As Shape
    Dim slideWidth As Single
    Dim slideHeight As Single
    Dim scaleX As Single
    Dim scaleY As Single
    Dim scale As Single

    slideWidth = ActivePresentation.PageSetup.SlideWidth
    slideHeight = ActivePresentation.PageSetup.SlideHeight

    For Each sld In ActivePresentation.Slides
        For Each shp In sld.Shapes
            If shp.Type = msoPicture Then
                ' Calcola i fattori di scala
                scaleX = slideWidth / shp.Width
                scaleY = slideHeight / shp.Height
                ' Usa il fattore più piccolo per mantenere proporzioni e contenere
                scale = IIf(scaleX < scaleY, scaleX, scaleY)
                If scale < 1 Then
                    shp.Width = shp.Width * scale
                    shp.Height = shp.Height * scale
                End If
            End If
        Next shp
    Next sld

    MsgBox "Immagini ridimensionate per contenere nella slide.", vbInformation
End Sub

' ------------------------------------------------------------'
' 6. Ridimensiona immagini per contenere nell'area libera
'
' Descrizione:
'   - Usa l'area libera calcolata da CalcolaAreaBiancaDisponibile.
'   - Ridimensiona tutte le immagini per farle stare entro l'area libera,
'     mantenendo le proporzioni originali.
' ------------------------------------------------------------
Sub RidimensionaImmaginiPerAreaLibera()
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
            If shp.Type = msoPicture Then
                ' Calcola i fattori di scala
                scaleX = AreaLiberaWidth / shp.Width
                scaleY = AreaLiberaHeight / shp.Height
                ' Usa il fattore più piccolo per mantenere proporzioni e contenere
                scale = IIf(scaleX < scaleY, scaleX, scaleY)
                If scale < 1 Then
                    shp.Width = shp.Width * scale
                    shp.Height = shp.Height * scale
                End If
                ' Posiziona nell'area libera se necessario (opzionale)
                If shp.Left < AreaLiberaLeft Then shp.Left = AreaLiberaLeft
                If shp.Top < AreaLiberaTop Then shp.Top = AreaLiberaTop
                If shp.Left + shp.Width > AreaLiberaLeft + AreaLiberaWidth Then shp.Left = AreaLiberaLeft + AreaLiberaWidth - shp.Width
                If shp.Top + shp.Height > AreaLiberaTop + AreaLiberaHeight Then shp.Top = AreaLiberaTop + AreaLiberaHeight - shp.Height
            End If
        Next shp
    Next sld

    MsgBox "Immagini ridimensionate per contenere nell'area libera.", vbInformation
End Sub</content>
<parameter name="filePath">c:\Codice Di Mia Produzione\CODICE VBA\Repository ServicePowerpoint\ModImmagPPT.bas