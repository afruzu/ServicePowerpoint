# PowerPoint VBA Utilities

Una raccolta di macro VBA per velocizzare l’editing e l’adattamento di presentazioni PowerPoint:
applicare loghi, rimuovere animazioni, uniformare font, pulire slide vuote e molto altro.

## ✨ Funzionalità (v1.1)

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

## 📦 Installazione

1. Apri PowerPoint.
2. Premi `ALT + F11` per aprire l’editor VBA.
3. Vai su `File → Importa file…` e seleziona `src/ModUtilityPPT.bas`.
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

## 🛣️ Roadmap

- v1.2: gestione immagini (ridimensionamento, sostituzione)
- v2.0: Add-in `.ppam` con pulsanti dedicati nella Ribbon
- v3.0: GUI (UserForm) per selezionare le operazioni

## 📄 Licenza

MIT License — vedi file `LICENSE`.
