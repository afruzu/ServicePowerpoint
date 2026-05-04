# PowerPoint VBA Utilities

Una raccolta di macro VBA per velocizzare l’editing e l’adattamento di presentazioni PowerPoint:
applicare loghi, rimuovere animazioni, uniformare font, pulire slide vuote e molto altro.

## ✨ Funzionalità (v1.2)

### ModUtilityPPT.bas - Utilità Generali

- Copia un oggetto selezionato in tutte le diapositive
- Rimuove oggetti con un certo nome da tutte le slide
- Rimuove tutte le animazioni
- Rimuove tutte le transizioni
- Uniforma il font in tutta la presentazione
- Elimina le slide vuote
- Cerca e sostituisci testo in tutte le slide
- Uniforma colore testo in tutta la presentazione
- Imposta dimensione font in tutta la presentazione
- Applica formattazione testo (bold, italic, underline, normal)
- Aggiungi numeri di pagina a tutte le slide
- Resetta layout delle slide
- Rinomina slide in batch
- Calcola area bianca disponibile nella slide

- Ridimensiona testo per contenere nell'area libera

### ModImmagPPT.bas - Gestione Immagini
- Ridimensiona immagini in tutta la presentazione
- Sostituisci immagine selezionata con una nuova
- Allinea immagini al centro delle slide
- Rimuovi tutte le immagini dalla presentazione
- Ridimensiona immagini per contenere nella slide
- Ridimensiona immagini per contenere nell'area libera

## 📦 Installazione

1. Apri PowerPoint.
2. Premi `ALT + F11` per aprire l’editor VBA.
3. Vai su `File → Importa file…` e seleziona `ModUtilityPPT.bas` e `ModImmagPPT.bas`.
4. Salva la presentazione come `.pptm` oppure crea un componente aggiuntivo `.ppam`.

## ▶️ Utilizzo

### ModUtilityPPT.bas - Utilità Generali
- Seleziona una forma e lancia `CopiaOggettoInTutteLeDiapositive`.
- Lancia `UniformaFont` per impostare un font unico.
- Usa `RimuoviTutteLeAnimazioni` e `RimuoviTutteLeTransizioni` per “ripulire” una presentazione.
- Usa `EliminaSlideVuote` per rimuovere rapidamente le slide senza contenuto.
- Usa `CercaSostituisciTesto` per cercare e sostituire testo.
- Usa `UniformaColoreTesto` per applicare un colore uniforme al testo.
- Usa `ImpostaDimensioneFont` per impostare la dimensione font.
- Usa `FormattaTesto` per applicare grassetto, corsivo, etc.
- Usa `AggiungiNumeriSlide` per aggiungere numeri di pagina.
- Usa `ResettaLayoutSlide` per riapplicare un layout.
- Usa `RinominaSlide` per rinominare le slide.
- Usa `CalcolaAreaBiancaDisponibile` per calcolare l'area libera centrale della slide.
- Usa `RidimensionaTestoPerAreaLibera` per ridimensionare testo entro l'area libera.

### ModImmagPPT.bas - Gestione Immagini
- Usa `RidimensionaImmagini` per ridimensionare tutte le immagini.
- Usa `SostituisciImmagine` per sostituire un'immagine selezionata.
- Usa `AllineaImmaginiCentro` per centrare le immagini.
- Usa `RimuoviTutteLeImmagini` per eliminare tutte le immagini.
- Usa `RidimensionaImmaginiPerSlide` per ridimensionare le immagini entro le dimensioni della slide.
- Usa `RidimensionaImmaginiPerAreaLibera` per ridimensionare immagini entro l'area libera.

## 🛣️ Roadmap

- v2.0: Add-in `.ppam` con pulsanti dedicati nella Ribbon
- v3.0: GUI (UserForm) per selezionare le operazioni

## 📄 Licenza

MIT License — vedi file `LICENSE`.

---

## 🇬🇧 English Version

### Features (v1.2)

#### ModUtilityPPT.bas - General Utilities
- Copy selected object to all slides
- Remove objects with a certain name from all slides
- Remove all animations
- Remove all transitions
- Uniform font throughout the presentation
- Delete empty slides
- Search and replace text in all slides
- Uniform text color throughout the presentation
- Set font size throughout the presentation
- Apply text formatting (bold, italic, underline, normal)
- Add page numbers to all slides
- Reset slide layouts
- Rename slides in batch
- Calculate available white space in the slide
- Resize text to fit within the free area

#### ModImmagPPT.bas - Image Management
- Resize images throughout the presentation
- Replace selected image with a new one
- Align images to the center of slides
- Remove all images from the presentation
- Resize images to fit within the slide
- Resize images to fit within the free area

### Usage

#### ModUtilityPPT.bas - General Utilities
- Select a shape and run `CopiaOggettoInTutteLeDiapositive` (CopyObjectToAllSlides).
- Run `UniformaFont` (UniformFont) to set a unique font.
- Use `RimuoviTutteLeAnimazioni` (RemoveAllAnimations) and `RimuoviTutteLeTransizioni` (RemoveAllTransitions) to "clean" a presentation.
- Use `EliminaSlideVuote` (DeleteEmptySlides) to quickly remove slides without content.
- Use `CercaSostituisciTesto` (SearchReplaceText) to search and replace text.
- Use `UniformaColoreTesto` (UniformTextColor) to apply uniform text color.
- Use `ImpostaDimensioneFont` (SetFontSize) to set font size.
- Use `FormattaTesto` (FormatText) to apply bold, italic, etc.
- Use `AggiungiNumeriSlide` (AddSlideNumbers) to add page numbers.
- Use `ResettaLayoutSlide` (ResetSlideLayout) to reapply a layout.
- Use `RinominaSlide` (RenameSlides) to rename slides.
- Use `CalcolaAreaBiancaDisponibile` (CalculateAvailableWhiteSpace) to calculate the central free area of the slide.
- Use `RidimensionaTestoPerAreaLibera` (ResizeTextForFreeArea) to resize text within the free area.

#### ModImmagPPT.bas - Image Management
- Use `RidimensionaImmagini` (ResizeImages) to resize all images.
- Use `SostituisciImmagine` (ReplaceImage) to replace a selected image.
- Use `AllineaImmaginiCentro` (AlignImagesCenter) to center images.
- Use `RimuoviTutteLeImmagini` (RemoveAllImages) to delete all images.
- Use `RidimensionaImmaginiPerSlide` (ResizeImagesForSlide) to resize images within slide dimensions.
- Use `RidimensionaImmaginiPerAreaLibera` (ResizeImagesForFreeArea) to resize images within the free area.
