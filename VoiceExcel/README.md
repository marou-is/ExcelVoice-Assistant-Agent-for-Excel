# Voice to Excel

Speak a request in French, and the app writes and runs the VBA in your open Excel workbook.

Pipeline: microphone -> Whisper (local speech-to-text) -> Gemini (generates VBA) -> Excel (runs it via COM).

## Setup (Windows, Python 3.10+)

```powershell
cd C:\Users\elazz\Desktop\VoiceExcel
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env   # then put your Gemini key in .env
```

In Excel, enable once: **File > Options > Trust Center > Trust Center Settings > Macro Settings >
"Trust access to the VBA project object model"**.

## Run

```powershell
python main.py
```

1. Open Excel (or let the app start it).
2. Hold **Hold to talk**, say e.g. *"Crée un tableau avec les colonnes id et âge du client"*, release.
3. The transcript, the generated VBA, and the result in Excel appear in turn.

Uncheck "Generate and run automatically" to review the transcript and the VBA before running it.
You can also type a request in the text box instead of speaking.
Every macro that runs successfully is saved in `history/` as a `.bas` file.

## Example requests

- "Crée un tableau des chantiers avec numéro, nom du chantier, ville, date de début, budget"
- "Ajoute une colonne total égale à quantité fois prix unitaire"
- "Remplis le tableau avec dix lignes d'exemple"
- "Fais un graphique en barres du budget par chantier"
- "Colore en rouge les budgets supérieurs à 500 000"
