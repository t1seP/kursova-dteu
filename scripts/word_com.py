"""Microsoft Word через COM з PowerShell (без pywin32).

Використовується, щоб оновити поле змісту, порахувати реальні сторінки
і за бажанням зберегти PDF. Працює лише на Windows з встановленим Word;
якщо Word недоступний, функції повертають None і скрипти працюють далі.
"""
import json
import os
import shutil
import subprocess
import tempfile

HEADINGS = ["ЗМІСТ", "ВСТУП", "ВИСНОВКИ ТА ПРОПОЗИЦІЇ", "ВИСНОВКИ І ПРОПОЗИЦІЇ",
            "СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ", "ДОДАТКИ"]

_PS = r'''
$ErrorActionPreference = "Stop"
$path = '__PATH__'
$out = '__OUT__'
$pdf = '__PDF__'
$targets = @(__TARGETS__)
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
$doc = $null
$res = @{ ok = $false }
try {
  $doc = $word.Documents.Open($path, $false, __READONLY__)
  if (__UPDATE__) {
    foreach ($t in $doc.TablesOfContents) { $t.Update() }
    $doc.Fields.Update() | Out-Null
    $doc.Repaginate()
    foreach ($t in $doc.TablesOfContents) { $t.Update() }
  }
  $doc.Repaginate()
  $res.pages = $doc.ComputeStatistics(2)
  $h = @{}
  foreach ($p in $doc.Paragraphs) {
    $t = $p.Range.Text.Trim()
    if ($t.Length -lt 40 -and $targets -contains $t -and -not $h.ContainsKey($t)) {
      $h[$t] = $p.Range.Information(3)
    }
  }
  $res.headings = $h
  if (__UPDATE__) { $doc.Save() }
  if ($pdf -ne '') { $doc.SaveAs2($pdf, 17) }
  $res.ok = $true
} catch {
  $res.error = $_.Exception.Message
} finally {
  if ($doc) { $doc.Close($false) }
  $word.Quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
  $res | ConvertTo-Json -Depth 4 | Out-File -FilePath $out -Encoding utf8
}
'''


def _q(s):
    """Рядок у одинарних лапках PowerShell."""
    return s.replace("'", "''")


def word_available():
    return os.name == "nt" and shutil.which("powershell") is not None


def run_word(docx_path, update=False, pdf_path=None, timeout=240):
    """Відкриває docx у Word. update=True оновлює зміст і поля та зберігає файл.

    Повертає dict {ok, pages, headings{назва: сторінка}, error} або None,
    якщо Word недоступний.
    """
    if not word_available():
        return None
    docx_path = os.path.abspath(docx_path)
    tmpdir = tempfile.mkdtemp(prefix="kursova_word_")
    out = os.path.join(tmpdir, "result.json")
    script = (_PS.replace("__PATH__", _q(docx_path))
                 .replace("__OUT__", _q(out))
                 .replace("__PDF__", _q(os.path.abspath(pdf_path)) if pdf_path else "")
                 .replace("__TARGETS__", ",".join("'" + _q(h) + "'" for h in HEADINGS))
                 .replace("__READONLY__", "$false" if update else "$true")
                 .replace("__UPDATE__", "$true" if update else "$false"))
    ps1 = os.path.join(tmpdir, "word.ps1")
    with open(ps1, "w", encoding="utf-8-sig") as f:
        f.write(script)
    try:
        subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", ps1],
                       capture_output=True, timeout=timeout)
    except (subprocess.TimeoutExpired, OSError) as e:
        return {"ok": False, "error": str(e)}
    if not os.path.exists(out):
        return {"ok": False, "error": "Word не повернув результат"}
    with open(out, encoding="utf-8-sig") as f:
        data = json.load(f)
    shutil.rmtree(tmpdir, ignore_errors=True)
    return data


def main_text_pages(headings):
    """Сторінки основного тексту: від ВСТУПУ до СПИСКУ ДЖЕРЕЛ (не включно)."""
    if not headings:
        return None
    start = headings.get("ВСТУП")
    end = headings.get("СПИСОК ВИКОРИСТАНИХ ДЖЕРЕЛ")
    if start and end:
        return end - start
    return None
