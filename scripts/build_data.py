# /// script
# requires-python = ">=3.10"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Gera programa.js a partir de data/programa.xlsx.

Uso:
    uv run scripts/build_data.py            # lê data/programa.xlsx, escreve programa.js
    uv run scripts/build_data.py caminho.xlsx [saida.js]

A planilha deve ter as abas "Sessões Técnicas" e "Sessões de Pôsteres" com as
colunas Sessão, Dia, Horário, [Tema], Ordem, ID, Trilha, Título, Autores.
Um asterisco ao final do título marca artigo concorrendo a best paper.
"""

import json
import re
import sys
import unicodedata
from datetime import date
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent.parent
YEAR = 2026

WEEKDAYS = {
    "seg": "Segunda-feira",
    "ter": "Terça-feira",
    "qua": "Quarta-feira",
    "qui": "Quinta-feira",
    "sex": "Sexta-feira",
    "sab": "Sábado",
    "sáb": "Sábado",
    "dom": "Domingo",
}

# Correções pontuais de digitação vindas da planilha, por ID do trabalho.
TITLE_FIXES = {
    35747: ("Recuperaçãoo", "Recuperação"),
}

# Correções pontuais em nomes de autores (texto exato -> correção).
AUTHOR_FIXES = {
    "Luis Felipe SIlva Toschi": "Luis Felipe Silva Toschi",
}

# Local de cada sessão: rótulo curto (badge), endereço completo (tooltip) e chave (cor).
# Se a planilha tiver uma coluna "Local", ela tem precedência.
VENUE = "Prédio 2"
ROOMS = {
    "teatro": {"key": "teatro", "label": "Teatro Moacyr Scliar", "full": "Teatro Moacyr Scliar, 4º andar, Prédio 2"},
    "s507": {"key": "s507", "label": "Sala 507", "full": "Sala 507, Prédio 2"},
    "s508": {"key": "s508", "label": "Sala 508", "full": "Sala 508, Prédio 2"},
    "hall": {"key": "hall", "label": "Hall 3º andar", "full": "Hall do 3º andar, Prédio 2"},
}
ROOM_POSTER = "hall"
ROOMS_ORAL = {1: "teatro", 5: "teatro", 8: "teatro", 2: "s507", 4: "s507", 7: "s507", 3: "s508", 6: "s508", 9: "s508"}

LOWER_PARTICLES = {"de", "da", "do", "das", "dos", "e", "di", "del", "van", "von"}


def slug(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]+", "-", s.lower()).strip("-")


def clean_text(s):
    if s is None:
        return ""
    s = str(s)
    s = s.replace("\\textit{", "").replace("\\textbf{", "")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def titlecase_name(name):
    """Converte nomes escritos todos em maiúsculas para capitalização normal."""
    if name != name.upper() or not any(c.isalpha() for c in name):
        return name
    parts = []
    for i, w in enumerate(name.lower().split(" ")):
        if i > 0 and w in LOWER_PARTICLES:
            parts.append(w)
        else:
            parts.append("-".join(p[:1].upper() + p[1:] for p in w.split("-")))
    return " ".join(parts)


def parse_day(raw):
    """'09/11 (seg)' -> dict com iso, label curto e nome do dia."""
    m = re.match(r"(\d{1,2})/(\d{1,2})\s*\(([^)]+)\)", raw.strip())
    if not m:
        raise ValueError(f"Dia em formato inesperado: {raw!r}")
    d, mo, wd = int(m.group(1)), int(m.group(2)), m.group(3).lower()
    iso = date(YEAR, mo, d).isoformat()
    return {
        "iso": iso,
        "short": f"{d:02d}/{mo:02d}",
        "weekday": WEEKDAYS.get(wd, wd),
        "weekdayShort": wd,
    }


def parse_authors(raw):
    names = [clean_text(a) for a in raw.split(",")]
    return [titlecase_name(AUTHOR_FIXES.get(n, n)) for n in names if n]


def room_for(name, kind, from_sheet=""):
    if from_sheet:
        return {"key": slug(from_sheet), "label": from_sheet, "full": from_sheet}
    if kind == "poster":
        return ROOMS[ROOM_POSTER]
    m = re.search(r"(\d+)\s*$", name)
    key = ROOMS_ORAL.get(int(m.group(1))) if m else None
    return ROOMS[key] if key else None


def read_sheet(ws, kind):
    rows = list(ws.iter_rows(values_only=True))
    header = [clean_text(h) for h in rows[0]]
    col = {h: i for i, h in enumerate(header) if h}
    required = ["Sessão", "Dia", "Horário", "Ordem", "ID", "Trilha", "Título", "Autores"]
    missing = [c for c in required if c not in col]
    if missing:
        raise ValueError(f"Aba {ws.title!r} sem colunas {missing}")

    sessions = {}
    for row in rows[1:]:
        if row is None or all(c is None for c in row):
            continue
        name = clean_text(row[col["Sessão"]])
        if not name:
            continue
        key = name
        if key not in sessions:
            day = parse_day(clean_text(row[col["Dia"]]))
            sessions[key] = {
                "id": slug(name),
                "name": name,
                "kind": kind,
                "day": day,
                "time": clean_text(row[col["Horário"]]).replace("-", "–"),
                "theme": clean_text(row[col["Tema"]]) if "Tema" in col else "",
                "room": room_for(name, kind, clean_text(row[col["Local"]]) if "Local" in col else ""),
                "papers": [],
            }
        title = clean_text(row[col["Título"]])
        best = title.endswith("*")
        title = title.rstrip("* ").strip()
        pid = int(float(row[col["ID"]]))
        if pid in TITLE_FIXES:
            old, new = TITLE_FIXES[pid]
            title = title.replace(old, new)
        sessions[key]["papers"].append({
            "order": int(float(row[col["Ordem"]])),
            "id": pid,
            "track": clean_text(row[col["Trilha"]]),
            "title": title,
            "authors": parse_authors(clean_text(row[col["Autores"]])),
            "bestPaper": best,
        })
    for s in sessions.values():
        s["papers"].sort(key=lambda p: p["order"])
    return list(sessions.values())


def main():
    src = Path(sys.argv[1]) if len(sys.argv) > 1 else ROOT / "data" / "programa.xlsx"
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else ROOT / "programa.js"

    wb = openpyxl.load_workbook(src, data_only=True, read_only=True)
    sessions = []
    sessions += read_sheet(wb["Sessões Técnicas"], "oral")
    sessions += read_sheet(wb["Sessões de Pôsteres"], "poster")

    ids = [p["id"] for s in sessions for p in s["papers"]]
    dup = sorted({i for i in ids if ids.count(i) > 1})
    if dup:
        print(f"aviso: IDs repetidos: {dup}", file=sys.stderr)

    days = sorted({json.dumps(s["day"], ensure_ascii=False) for s in sessions})
    data = {
        "event": {
            "name": "ERAMIA-RS 2026",
            "venue": VENUE,
            "subtitle": "Escola Regional de Aprendizado de Máquina e Inteligência Artificial do Rio Grande do Sul",
        },
        "generatedAt": date.today().isoformat(),
        "days": [json.loads(d) for d in days],
        "sessions": sessions,
    }
    payload = json.dumps(data, ensure_ascii=False, indent=1)
    out.write_text(
        "// Gerado por scripts/build_data.py a partir de data/programa.xlsx. Não edite à mão.\n"
        f"window.PROGRAMA = {payload};\n",
        encoding="utf-8",
    )
    n_oral = sum(len(s["papers"]) for s in sessions if s["kind"] == "oral")
    n_poster = sum(len(s["papers"]) for s in sessions if s["kind"] == "poster")
    print(f"{out.name}: {len(sessions)} sessões, {n_oral} orais, {n_poster} pôsteres")


if __name__ == "__main__":
    main()
