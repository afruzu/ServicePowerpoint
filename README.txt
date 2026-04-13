# PowerPoint VBA Utilities

Una raccolta di macro VBA per velocizzare l’editing e l’adattamento di presentazioni PowerPoint:
applicare loghi, rimuovere animazioni, uniformare font, pulire slide vuote e molto altro.
per comodità di lettura la libreria all'aumentare del numero di routines implementate è suddivisa in più moduli 
(attualmente due) - da caricare all'occorrenza.

## ✨ Funzionalità (v1.3)

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
- Ridimensiona immagini per contenere nell'area libera
- Ridimensiona testo per contenere nell'area libera
- Ridimensiona immagini in tutta la presentazione
- Sostituisci immagine selezionata con una nuova
- Allinea immagini al centro delle slide
- Rimuovi tutte le immagini dalla presentazione
- Ridimensiona immagini per contenere nella slide

## 📦 Installazione

1. Apri PowerPoint.
2. Premi `ALT + F11` per aprire l’editor VBA.
3. Vai su `File → Importa file…` e seleziona `ModUtilityPPT.bas` e `ModImmagPPT.bas`.
4. Salva la presentazione come `.pptm` oppure crea un componente aggiuntivo `.ppam`.

## ▶️ Utilizzo

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
- Usa `RidimensionaImmaginiPerAreaLibera` per ridimensionare immagini entro l'area libera.
- Usa `RidimensionaTestoPerAreaLibera` per ridimensionare testo entro l'area libera.
- Usa `RidimensionaImmagini` per ridimensionare tutte le immagini.
- Usa `SostituisciImmagine` per sostituire un'immagine selezionata.
- Usa `AllineaImmaginiCentro` per centrare le immagini.
- Usa `RimuoviTutteLeImmagini` per eliminare tutte le immagini.
- Usa `RidimensionaImmaginiPerSlide` per ridimensionare le immagini entro le dimensioni della slide.

## 🛣️ Roadmap

- v2.0: Add-in `.ppam` con pulsanti dedicati nella Ribbon
- v3.0: GUI (UserForm) per selezionare le operazioni

## 📄 Licenza

MIT License — vedi file `LICENSE`.
