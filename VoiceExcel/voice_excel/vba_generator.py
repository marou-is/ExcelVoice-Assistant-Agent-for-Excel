import re

from google import genai
from google.genai import types

from . import config

ENTRY_POINT = "VoiceMacro"

SYSTEM_PROMPT = f"""You are an expert Excel VBA developer embedded in a voice assistant used by a
construction company. The user speaks French; their spoken request has been transcribed
(possibly with small speech-recognition mistakes, so infer the intended meaning).

Your job: translate the request into ONE standard VBA module that performs it in the
currently open Excel workbook.

Output rules:
- Reply with a single ```vba code block and nothing else.
- The module MUST contain `Public Sub {ENTRY_POINT}()` as its entry point. Helper Subs and
  Functions are allowed. Do not use `Option Explicit`.
- Work on ActiveWorkbook / ActiveSheet unless the user names a specific sheet or table.
- The macro runs unattended: never use MsgBox, InputBox, or UserForms.
- Do not use `On Error Resume Next` to hide failures; let real errors propagate so the
  assistant can report and repair them.

Excel conventions:
- "tableau" / "table" means a real Excel table (ListObject) with a header row, a table
  style (e.g. TableStyleMedium2) and auto-fitted columns.
- Use the column names the user says. Normalise them to readable headers
  (e.g. "âge du client" -> "Âge client", "id" -> "ID") unless the user spells an exact name
  such as "customer_age", in which case keep it verbatim.
- Give new tables a meaningful unique name (e.g. "tblClients"); if the name exists, append
  a number.
- Never overwrite existing data. If the active sheet is empty, start at A1; otherwise place
  new content two rows below the used range.
- Apply sensible number formats based on column meaning (dates, amounts with 2 decimals,
  percentages, integers for IDs and quantities).
- Only fill in sample data when the user asks for it; then make it realistic for a
  construction business (chantiers, matériaux, fournisseurs, clients...).
- For modifications ("ajoute une colonne", "trie", "filtre", "graphique", "formule",
  "total"), use the workbook context below to target the right sheet/table.

Safety: do not touch the file system, network, registry, other applications or other
workbooks, and do not save, close or delete workbooks, unless the user explicitly asks.
"""

RISKY_PATTERNS = {
    r"\bShell\s*\(": "runs an external program (Shell)",
    r"\bKill\b": "deletes files (Kill)",
    r"\bRmDir\b": "deletes folders (RmDir)",
    r"FileSystemObject": "accesses the file system",
    r"WScript\.": "uses Windows Script Host",
    r"\bCreateObject\s*\(": "creates an external COM object",
    r"\bDeclare\s+(PtrSafe\s+)?(Function|Sub)\b": "calls Windows API functions",
    r"\.SaveAs\b|\.Save\b": "saves a workbook",
    r"Workbooks\.Open|\.Close\b": "opens or closes workbooks",
    r"URLDownloadToFile|XMLHTTP|WinHttp": "accesses the network",
    r"\bSendKeys\b": "simulates keystrokes",
    r"\.Delete\b": "deletes Excel objects (sheets, rows, ranges...)",
}


class GenerationError(Exception):
    pass


class VBAGenerator:
    def __init__(self):
        if not config.GEMINI_API_KEY:
            raise GenerationError(
                "GEMINI_API_KEY is missing. Copy .env.example to .env and add your key."
            )
        self._client = genai.Client(api_key=config.GEMINI_API_KEY)

    def generate(self, request: str, workbook_context: str) -> str:
        prompt = (
            f"Workbook context:\n{workbook_context}\n\n"
            f"User request (transcribed French speech):\n{request}"
        )
        return self._ask(prompt)

    def fix(self, request: str, workbook_context: str, code: str, error: str) -> str:
        prompt = (
            f"Workbook context:\n{workbook_context}\n\n"
            f"User request (transcribed French speech):\n{request}\n\n"
            f"Your previous module failed in Excel with this error:\n{error}\n\n"
            f"Previous module:\n```vba\n{code}\n```\n\n"
            "Return a corrected full module. The previous attempt may have partially "
            "run, so make the fix tolerate objects it already created."
        )
        return self._ask(prompt)

    def _ask(self, prompt: str) -> str:
        response = self._client.models.generate_content(
            model=config.GEMINI_MODEL,
            contents=prompt,
            config=types.GenerateContentConfig(
                system_instruction=SYSTEM_PROMPT,
                temperature=0.2,
            ),
        )
        return extract_vba(response.text or "")


def extract_vba(text: str) -> str:
    match = re.search(r"```(?:vba|vb|vbnet|basic)?\s*\n(.*?)```", text, re.DOTALL | re.IGNORECASE)
    code = (match.group(1) if match else text).strip()
    if not re.search(rf"\bSub\s+{ENTRY_POINT}\s*\(", code, re.IGNORECASE):
        raise GenerationError(f"Gemini did not return a `Sub {ENTRY_POINT}()`:\n{text}")
    return code


def find_risky_operations(code: str) -> list[str]:
    return [reason for pattern, reason in RISKY_PATTERNS.items() if re.search(pattern, code, re.IGNORECASE)]
