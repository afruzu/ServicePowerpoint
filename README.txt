# PowerPoint VBA Utilities

Una raccolta di macro VBA per velocizzare l’editing e l’adattamento di presentazioni PowerPoint:
applicare loghi, rimuovere animazioni, uniformare font, pulire slide vuote e molto altro.

## ✨ Funzionalità (v1.0)

- Copia un oggetto selezionato in tutte le diapositive
- Rimuove oggetti con un certo nome da tutte le slide
- Rimuove tutte le animazioni
- Rimuove tutte le transizioni
- Uniforma il font in tutta la presentazione
- Elimina le slide vuote

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

## 🛣️ Roadmap

- v1.1: gestione immagini (ridimensionamento, sostituzione)
- v1.2: normalizzazione layout e placeholder
- v2.0: Add-in `.ppam` con pulsanti dedicati nella Ribbon
- v3.0: GUI (UserForm) per selezionare le operazioni

## 📄 Licenza

MIT License — vedi file `LICENSE`.
