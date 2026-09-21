# Specifiche e Linee Guida di Progetto APOffice

Questo documento definisce gli standard architetturali, grafici, funzionali e le regole operative per lo sviluppo e la modernizzazione delle form di **APOffice**.

---

## 1. Regola Continua di Compilazione e Versionamento
- **Ogni volta che si ricompila la soluzione**, incrementare la versione in [`Properties/AssemblyInfo.cs`](file:///c:/ApProject/APOffice/Properties/AssemblyInfo.cs) (campi `AssemblyVersion`, `AssemblyFileVersion` e `AssemblyInformationalVersion`).
- **Comando di compilazione standard (MSBuild)**:
  ```powershell
  Stop-Process -Name "APOffice" -Force -ErrorAction SilentlyContinue; & "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" "c:\ApProject\APOffice\APOffice.csproj" /t:Build /p:Configuration=Debug /v:m
  ```

---

## 2. Standard Grafici e UI Modernization

### A. Menu e Barre Superiori (`MenuStrip`)
- Applicare sempre `ModernMenuRenderer`:
  ```csharp
  menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
  ```
- Assegnare icone vettoriali ad alta definizione (`16x16`) a tutte le voci di menu tramite `clsUiIcons.GetIcon(...)` (es. `exit`, `excel`, `print`, `search`, `plus`, `trash`, `sync`, `tools_gear`, `barcode`, `send`, `smartphone`, `document_edit`).

### B. Pulsanti Glossy con Gradienti e Icone Vettoriali
Utilizzare `clsUiIcons.StyleStatButton(...)` per tutti i pulsanti operativi principali:
- **Tutti / Ricerca / Estrai**: Gradiente Sky Blue (`#EFF6FF` ➔ `#BFDBFE`, bordo `#60A5FA`, testo `#1E3A8A`) + icona `list` o `search`.
- **Nuovo / Aggiungi**: Gradiente Verde Smeraldo (`#ECFDF5` ➔ `#A7F3D0`, bordo `#34D399`, testo `#064E3B`) + icona `plus`.
- **Stampa / Report**: Gradiente Ambra (`#FEFCE8` ➔ `#FEF08A`, bordo `#FAC015`, testo `#713F12`) + icona `print`.
- **Invio / Trasmissione / Casse**: Gradiente Teal (`#F0FDF4` ➔ `#99F6E4`, bordo `#2DD4BF`, testo `#134E4A`) + icona `send`.
- **Elimina / Cancella**: Gradiente Rosso Soft (`#FEF2F2` ➔ `#FECACA`, bordo `#F87171`, testo `#7F1D1D`) + icona `trash`.
- **Sincronizzazione / Rettifiche / Raggruppa**: Gradiente Viola (`#F5F3FF` ➔ `#DDD6FE`, bordo `#A78BFA`, testo `#4C1D95`) + icona `sync` o `edit`.

### C. Griglie Dati (`APDataGridView`)
- Applicare sempre `clsUiIcons.StyleDataGridView(dgv)` all'avvio della Form (in `ApplyModernUi()`).
- Impostare la colonna descrittiva principale su `AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill`.
- Formattazione automatica importi numerici a due decimali (`#,##0.00` o `N2`) e allineamento `MiddleRight`.
- Codici, date e stati allineati al centro (`MiddleCenter`).

### D. Layout Responsive & Persistenza
- Sfondo form impostato a grigio chiaro moderno `#F3F4F6`.
- Ancoraggio elastico dei controlli (evitare controlli rigidi o resize con font stirati):
  - Griglia: `Anchor = Top | Bottom | Left | Right`
  - Controlli/Pulsanti superiori: `Top | Left` o `Top | Right`
  - Totali/Badges inferiori: `Bottom | Right`
- Persistenza automatica della finestra e delle larghezze delle colonne:
  - In `Form_Load`: `clsUiIcons.RestoreFormBounds(this); clsUiIcons.RestoreGridColumnWidths(dgv, "NomeForm_dgv");`
  - In `Form_Closing`: `clsUiIcons.SaveFormBounds(this); clsUiIcons.SaveGridColumnWidths(dgv, "NomeForm_dgv");`

---

## 3. Integrità del Codice & Funzionalità
- Mantenere rigorosamente tutti i binding, gli event handler (`Click`, `KeyDown`, `DoubleClick`, `SelectionChangeCommitted`), le chiamate SQL e i parametri.
- Assicurarsi sempre che ogni evento registrato nel `.Designer.cs` abbia il corrispondente metodo implementato nel file `.cs` per garantire 0 errori di compilazione.

---

## 4. Regole di Calcolo, Scorporo e Coefficiente di Ricarico
Per garantire la perfetta reversibilità bidirezionale tra Prezzo di Vendita Ivato, Costo Netto e Ricarico:
- **Relazione tra Prezzo Ivato, Prezzo Netto e Costo**:
  $$\text{Prezzo Netto} = \frac{P_{\text{vendita}}}{1 + \frac{\text{AliquotaIVA}}{100}}$$
- **Coefficiente di Ricarico ($K$)**:
  - Espresso come moltiplicatore (es. `1.30`) o percentuale (es. `30%`): $K = 1 + \frac{\text{PercRicarico}}{100}$
  - Calcolo del Costo Netto di Acquisto:
    $$\text{Costo Netto} = \frac{\text{Prezzo Netto}}{K} = \frac{P_{\text{vendita}}}{\left(1 + \frac{\text{AliquotaIVA}}{100}\right) \times K}$$
- **Verifica in Avanti (Reversibilità)**:
  $$P_{\text{vendita}} = \text{Costo Netto} \times K \times \left(1 + \frac{\text{AliquotaIVA}}{100}\right)$$
- **Parsing Stringhe Numeriche e Culture:**
  - **MAI** usare `Convert.ToDecimal(string)` o `decimal.Parse(string)` direttamente su stringhe formattate con `.` o provenienti da `CultureInfo.InvariantCulture`, poiché sotto cultura italiana (`it-IT`) il punto `.` viene trattato come separatore delle migliaia moltiplicando il numero per 100!
  - Utilizzare sempre `_clsFun.Txt2Dec(string)` che normalizza preventivamente la virgola e il punto.

---

## 5. Specifiche Tabelle Database & Best Practices
- **`GesLisAcquisto` (Listino Acquisto Fornitori)**:
  - `lia_day`: Colonna timestamp con vincolo **`NOT NULL`**. Valorizzare sempre con `GETDATE()` o data odierna sia nelle `INSERT` che nelle `UPDATE`.
  - `lia_arf`: Codice articolo fornitore; impostare uguale al codice articolo (`lia_arf = art_cod`).
  - `lia_tip`: `'L'` (Listino), `'M'` (Manuale), `'F'` (Fattura).
  - Valori predefiniti obbligatori: `lia_pxc = 1`, `lia_cxp = 0`, `lia_stf = 'A'`, `lia_ann = 0`, `lia_dtf = 2050-01-01` (`_clsDef.DAYOUT`).
- **`AnaArticoli`**:
  - Non interrogare campi inesistenti come `art_ann`; il filtro attivo si basa su `art_sta = 'A'`.
- **Parsing Sicuro Campi Numerici da DB**:
  - I campi anagrafici come `art_gsc` (giorni di scadenza), pesi o prezzi possono contenere stringhe vuote `""` o `null`.
  - Evitare conversioni dirette `Convert.ToDouble(...)` o unboxing diretti `(decimal)...`.
  - Utilizzare sempre `double.TryParse` e `decimal.TryParse` con fallback predefinito a zero (applicato in `frmAnaArtIngredienti`, `frmAnaArtIngredient2`, `frmAnaArticolo`).

---

## 6. Modulo Generazione Costi (`frmUtyGenCosti`)
- **Punto di accesso**: Menu Stampe ([`frmMnuStampe`](file:///c:/ApProject/APOffice/frmMnuStampe.cs)) -> Pulsante *Genera Costi*.
- **Funzionalità**:
  - Filtri: Fornitore, Data Registrazione, Coeff. / % di Ricarico (supporta sia `1.30` sia `30%`), Solo Articoli Attivi, Aggiorna costo base in anagrafica (`art_cos`).
  - Griglia interattiva con colonne: Codice, Descrizione, IVA %, P. Vendita, P. Netto, Coeff. Ric., Costo Calc., Costo Att. Click sulla colonna Descrizione o Codice (o doppio click riga) apre l'anagrafica prodotto ([`frmAnaArticolo`](file:///c:/ApProject/APOffice/frmAnaArticolo.cs)).
  - Salvataggio a blocchi di 300 istruzioni con progress bar e notifica esito.

---

## 7. Fatturazione Elettronica XML SdI (`frmGesFattElettroniche`)
- **Encoding**: Sempre `new UTF8Encoding(false)` (UTF-8 senza BOM), per evitare l'errore SdI `00200` causato da encoding ANSI o BOM non conforme.
- **Sanitizzazione Testi**: Utilizzare `XmlEsc()` (`System.Security.SecurityElement.Escape`) preservando lettere accentate e caratteri UTF-8 ammessi.
- **Progressivo Invio**: Max 5 caratteri alfanumerici (`IT<PIVA>_<Nid>.xml`) per rispettare la regola SdI `00102`.
- **Parametri Invio PEC SDI**: Indirizzi PEC in `frmAnaGestione` (`txtCnfPec`, `txtCnfAgp`), credenziali SMTP in `C:\ApProject\Temp\DivDocumenti\AgenziaEntrate.ini`.

