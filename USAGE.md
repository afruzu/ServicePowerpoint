# Utilizzo delle PowerPoint VBA Utilities

Questo documento descrive come importare, configurare e utilizzare le macro contenute nel modulo `ModUtilityPPT.bas` per velocizzare l’editing di presentazioni PowerPoint.

---

## 📦 Importazione del modulo

1. Apri PowerPoint.
2. Premi `ALT + F11` per aprire l’Editor VBA.
3. Vai su **File → Importa file…**
4. Seleziona `src/ModUtilityPPT.bas`.
5. Salva la presentazione come:
   - `.pptm` (presentazione con macro), oppure
   - `.ppam` (componente aggiuntivo PowerPoint).

---

## ▶️ Esecuzione delle macro

Puoi eseguire le macro in tre modi:

### **1. Da VBA**

- Premi `ALT + F11`
- Premi `F5` sulla macro desiderata

### **2. Da PowerPoint**

- Vai su **Visualizza → Macro**
- Seleziona la macro
- Clicca **Esegui**

### **3. Assegnandole a un pulsante**

- Inserisci una forma
- Tasto destro → **Assegna macro…**
- Seleziona la macro

---

## 🧰 Macro disponibili (v1.0)

### **1. CopiaOggettoInTutteLeDiapositive**

Copia l’oggetto selezionato e lo incolla in tutte le slide, mantenendo posizione e dimensioni.

**Uso:**

- Seleziona una forma
- Esegui la macro

---

### **2. RimuoviOggettiPerNome**

Elimina da tutte le slide le forme con un determinato nome (`Shape.Name`).

**Uso:**

- Esegui la macro
- Inserisci il nome richiesto nella finestra di dialogo

---

### **3. RimuoviTutteLeAnimazioni**

Rimuove tutte le animazioni dalla presentazione.

---

### **4. RimuoviTutteLeTransizioni**

Imposta la transizione di ogni slide su “Nessuna”.

---

### **5. UniformaFont**

Applica un font unico a tutto il testo della presentazione.

**Uso:**

- Esegui la macro
- Inserisci il nome del font (es. `Calibri`, `Arial`)

---

### **6. EliminaSlideVuote**

Rimuove automaticamente tutte le slide prive di forme.

---

### **7. CercaSostituisciTesto**

Cerca e sostituisce testo in tutte le forme di tutte le slide.

**Uso:**

- Esegui la macro
- Inserisci il testo da cercare
- Inserisci il testo di sostituzione

---

### **8. UniformaColoreTesto**

Applica un colore uniforme a tutto il testo della presentazione.

**Uso:**

- Esegui la macro
- Inserisci il colore RGB come R,G,B (es. 0,0,0 per nero)

---

### **9. ImpostaDimensioneFont**

Imposta una dimensione font uniforme a tutto il testo.

**Uso:**

- Esegui la macro
- Inserisci la dimensione (es. 24)

---

### **10. FormattaTesto**

Applica formattazione testo (grassetto, corsivo, sottolineato, normale) a tutto il testo.

**Uso:**

- Esegui la macro
- Inserisci il tipo: bold, italic, underline, normal

---

### **11. AggiungiNumeriSlide**

Aggiunge numeri di pagina nel footer di ogni slide.

---

### **12. ResettaLayoutSlide**

Riapplica un layout specifico a tutte le slide.

**Uso:**

- Esegui la macro
- Inserisci il nome del layout (es. Title Slide, Blank)

---

### **13. RinominaSlide**

Rinomina ogni slide come "Slide_X" dove X è il numero della slide.

---

## 🛠 Suggerimenti d’uso

- Usa `CopiaOggettoInTutteLeDiapositive` per loghi, watermark, numerazioni personalizzate.
- Usa `RimuoviTutteLeAnimazioni` + `RimuoviTutteLeTransizioni` per “ripulire” presentazioni ricevute da terzi.
- Usa `UniformaFont` per dare un aspetto coerente a materiali eterogenei.
- Usa `EliminaSlideVuote` dopo importazioni da PDF o Word.
- Usa `CercaSostituisciTesto` per aggiornare testi comuni come nomi o date.
- Usa `UniformaColoreTesto` e `ImpostaDimensioneFont` per uniformare lo stile testo.
- Usa `FormattaTesto` per applicare stili rapidi.
- Usa `AggiungiNumeriSlide` per aggiungere numerazione automatica.
- Usa `ResettaLayoutSlide` per standardizzare layout.
- Usa `RinominaSlide` per organizzare slide con nomi semplici.

---

## 🧭 Compatibilità

- PowerPoint per Windows (VBA abilitato)
- Versioni consigliate: Office 2016, 2019, 2021, Microsoft 365

---

## 📄 Licenza

Questo progetto è distribuito sotto licenza MIT.  
Vedi il file `LICENSE` per maggiori dettagli.
