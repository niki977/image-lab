# Pubblicare The Ultraspeaker Image Lab sul Marketplace Microsoft

Il pacchetto è pronto per l’invio. Prima di iniziare completa i punti marcati **DA FARE**.

## 1. DA FARE prima dell’invio

1. **Titolare nelle pagine legali.** In `privacy.html` e `terms.html` sostituisci
   `[NOME E COGNOME DEL TITOLARE]` e `[FULL NAME OF THE OWNER]` con il nome (o la ragione sociale)
   dell’account con cui pubblichi in Partner Center. Deve coincidere con il nome dell’editore.
2. **Casella email.** Verifica che `info@theultraspeaker.com` riceva posta: il team Microsoft la usa per i test
   e gli utenti per il supporto.
3. **Carica tutto su GitHub** (repository `image-lab`), comprese le nuove pagine `support.html`, `privacy.html`,
   `terms.html`, `legal.css` e la cartella `store` (facoltativa: non serve al funzionamento).
4. **Controlla che rispondano** questi indirizzi:
   - `https://niki977.github.io/image-lab/index.html`
   - `https://niki977.github.io/image-lab/support.html`
   - `https://niki977.github.io/image-lab/privacy.html`
   - `https://niki977.github.io/image-lab/terms.html`
   - `https://niki977.github.io/image-lab/assets/image-lab-icon-80.png`
5. **Prova finale** su almeno due piattaforme tra Mac, Windows e PowerPoint sul web, con `test-presentation.pptx`:
   selezione, regolazione, Applica, nuova selezione (le regolazioni tornano), Originale.
6. **Convalida del manifest** sul tuo computer (serve Node.js):
   `npx office-addin-manifest validate image-lab-manifest.xml`
   Deve rispondere “The manifest is valid”. Partner Center ripete comunque la stessa verifica al caricamento.
7. **Ricarica il manifest** nella cartella `wef` del Mac: è cambiato (versione 1.1.0.0, requisiti, link al supporto).

## 2. Account Partner Center

1. Vai su **partner.microsoft.com** e iscriviti al programma **Microsoft 365 e Copilot** (account sviluppatore,
   individuale o aziendale). La verifica dell’identità può richiedere alcuni giorni.
2. Il nome dell’editore che scegli comparirà nello store: usa lo stesso delle pagine legali.

## 3. Creare l’offerta

In Partner Center: **Marketplace offers → Microsoft 365 e Copilot → + Nuova offerta → Componente aggiuntivo di Office**.

| Sezione | Cosa inserire |
|---|---|
| Nome offerta | The Ultraspeaker Image Lab |
| Pacchetto | `image-lab-manifest.xml` |
| Proprietà: categorie | Productivity, Design (o le più vicine proposte) |
| Proprietà: privacy | `https://niki977.github.io/image-lab/privacy.html` |
| Proprietà: contratto di licenza | Contratto standard Microsoft oppure `terms.html` |
| Scheda (per lingua) | Testi di `LISTING.md`, icona `assets/image-lab-icon-300.png`, screenshot `screenshots/<lingua>-1..3.png` |
| Supporto | `https://niki977.github.io/image-lab/support.html` |
| Disponibilità | Tutti i mercati, gratuito |
| Note per la certificazione | Testo qui sotto, in inglese, e allega `test-presentation.pptx` o un link per scaricarla |

## 4. Note per la certificazione (da incollare)

```
Image Lab adjusts photos on PowerPoint slides (histogram, levels, exposure, colour).
Requirements: PowerPointApi 1.10 (declared in the manifest). Tested on PowerPoint for Mac,
Windows and PowerPoint on the web. Not available on iPad (PowerPointApi is not supported there).

How to test:
1. Open the attached test-presentation.pptx (or any slide with a picture).
2. Select the photo, then Home tab > The Ultraspeaker > Image Lab.
3. Move the level sliders under the histogram or any adjustment slider; the preview updates.
4. Click "Apply to slide": the adjusted photo replaces the selected one with the same position,
   size and rotation. The original remains on the slide, hidden (visible in the Selection Pane).
5. Select the adjusted photo again: the pane restores the saved settings.
6. Click "Original" twice (confirm) to restore the original photo.
7. The language menu (top right) switches between EN, IT, ES, FR, DE.

No sign-in, no external services, no data collection. Images are processed locally in the task pane.
Support: info@theultraspeaker.com
```

## 5. Dopo l’invio

- La certificazione di Microsoft richiede in genere alcuni giorni. Se chiedono modifiche, arrivano per email con
  il riferimento alle regole del marketplace.
- Per gli aggiornamenti del codice basta ricaricare i file su GitHub. Se cambia il manifest (per esempio la versione),
  va ricaricato anche in Partner Center e l’offerta ripassa la certificazione.
