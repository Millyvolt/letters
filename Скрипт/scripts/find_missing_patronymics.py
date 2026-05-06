import argparse
import html
import os
import re
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import pandas as pd


PATRONYMIC_ENDINGS = (
    "ович",
    "евич",
    "ич",
    "овна",
    "евна",
    "ична",
    "инична",
)


@dataclass
class NameRecord:
    row_excel: int
    original: str
    first_name: str
    surname: str


@dataclass
class Suggestion:
    patronymic: str
    confidence: float
    evidence: str
    query: str
    source_hits: str
    sources_used: int


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _is_word(token: str) -> bool:
    return bool(re.fullmatch(r"[А-Яа-яЁё-]+", token))


def _looks_like_patronymic(word: str) -> bool:
    w = word.lower()
    return any(w.endswith(end) for end in PATRONYMIC_ENDINGS)


def parse_missing_patronymic_name(text: str, row_excel: int) -> Optional[NameRecord]:
    """
    Expected source format in column A: "Имя Отчество Фамилия" or "Имя Фамилия".
    We only return records where patronymic appears to be missing (2 words).
    """
    tokens = [t for t in text.split() if t]
    if len(tokens) != 2:
        return None

    t0, t1 = tokens[0], tokens[1]

    # Case 1: "Имя Фамилия"
    if _is_word(t0) and _is_word(t1):
        first_name, surname = t0, t1
    # Case 2: "Р.Е. Фамилия" (initials + surname): we can only search by surname.
    elif re.fullmatch(r"[А-ЯЁ]\.[А-ЯЁ]\.", t0) and _is_word(t1):
        first_name, surname = "", t1
    else:
        return None

    return NameRecord(
        row_excel=row_excel,
        original=text,
        first_name=first_name,
        surname=surname,
    )


def _http_get(url: str, timeout_sec: float = 20.0) -> str:
    request = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            )
        },
    )
    with urllib.request.urlopen(request, timeout=timeout_sec) as response:
        return response.read().decode("utf-8", errors="ignore")


def _http_get_via_jina(url: str, timeout_sec: float = 25.0) -> str:
    mirror_url = f"https://r.jina.ai/http://{url.replace('https://', '').replace('http://', '')}"
    return _http_get(mirror_url, timeout_sec=timeout_sec)


def search_duckduckgo(query: str) -> str:
    encoded = urllib.parse.quote_plus(query)
    url = f"https://duckduckgo.com/html/?q={encoded}"
    try:
        return _http_get(url)
    except Exception:
        return _http_get_via_jina(url)


def search_bing(query: str) -> str:
    encoded = urllib.parse.quote_plus(query)
    url = f"https://www.bing.com/search?q={encoded}&setlang=ru"
    try:
        return _http_get(url)
    except Exception:
        return _http_get_via_jina(url)


def search_yandex(query: str) -> str:
    encoded = urllib.parse.quote_plus(query)
    url = f"https://yandex.ru/search/?text={encoded}&lr=213"
    try:
        return _http_get(url)
    except Exception:
        return _http_get_via_jina(url)


def search_wikipedia(first_name: str, surname: str) -> str:
    # 1) Find relevant pages.
    person_query = f"{first_name} {surname}".strip()
    encoded = urllib.parse.quote_plus(person_query)
    search_url = (
        "https://ru.wikipedia.org/w/api.php?action=query&list=search"
        f"&srsearch={encoded}&format=json&utf8=1&srlimit=3"
    )
    raw_json = _http_get(search_url)
    try:
        import json

        data = json.loads(raw_json)
        titles = [item.get("title", "") for item in data.get("query", {}).get("search", [])]
    except Exception:
        return ""

    if not titles:
        return ""

    # 2) Fetch plain text extracts for found pages.
    all_text_parts: List[str] = []
    for title in titles:
        t = urllib.parse.quote_plus(title)
        extract_url = (
            "https://ru.wikipedia.org/w/api.php?action=query&prop=extracts"
            f"&explaintext=1&format=json&utf8=1&titles={t}"
        )
        try:
            extract_raw = _http_get(extract_url)
            data = json.loads(extract_raw)
            pages = data.get("query", {}).get("pages", {})
            for p in pages.values():
                txt = str(p.get("extract", ""))
                if txt:
                    all_text_parts.append(txt)
        except Exception:
            continue

    return "\n".join(all_text_parts)


def search_serpapi(query: str, api_key: str) -> str:
    encoded = urllib.parse.quote_plus(query)
    url = (
        "https://serpapi.com/search.json"
        f"?q={encoded}&hl=ru&gl=ru&engine=google&api_key={urllib.parse.quote_plus(api_key)}"
    )
    raw_json = _http_get(url)
    try:
        import json

        data = json.loads(raw_json)
    except Exception:
        return ""

    text_parts: List[str] = []
    for item in data.get("organic_results", [])[:10]:
        title = str(item.get("title", ""))
        snippet = str(item.get("snippet", ""))
        link = str(item.get("link", ""))
        block = " ".join([title, snippet, link]).strip()
        if block:
            text_parts.append(block)
    return "\n".join(text_parts)


def extract_patronymic_candidates(
    html_text: str,
    first_name: str,
    surname: str,
) -> List[Tuple[str, str]]:
    """
    Returns list of tuples: (candidate_patronymic, evidence_snippet).
    """
    plain = html.unescape(html_text)
    plain = re.sub(r"<[^>]+>", " ", plain)
    plain = re.sub(r"\s+", " ", plain)

    candidates: List[Tuple[str, str]] = []

    # Match variants:
    # 1) Имя Отчество Фамилия
    # 2) Фамилия Имя Отчество
    escaped_first = re.escape(first_name)
    escaped_surname = re.escape(surname)
    patr = r"([А-ЯЁ][а-яё]+)"

    patterns = [
        re.compile(rf"\b{escaped_first}\s+{patr}\s+{escaped_surname}\b"),
        re.compile(rf"\b{escaped_surname}\s+{escaped_first}\s+{patr}\b"),
    ]

    for pattern in patterns:
        for match in pattern.finditer(plain):
            candidate = match.group(1)
            if not _looks_like_patronymic(candidate):
                continue

            start = max(0, match.start() - 90)
            end = min(len(plain), match.end() + 90)
            snippet = plain[start:end].strip()
            candidates.append((candidate, snippet))

    return candidates


def pick_best_suggestion(
    source_candidates: Dict[str, List[Tuple[str, str]]],
    query: str,
) -> Optional[Suggestion]:
    candidates: List[Tuple[str, str]] = []
    source_hits_counter: Dict[str, int] = {}
    source_seen_for_patronymic: Dict[str, set] = {}

    for source_name, items in source_candidates.items():
        for patr, snippet in items:
            key = patr.strip()
            candidates.append((key, snippet))
            source_hits_counter[source_name] = source_hits_counter.get(source_name, 0) + 1
            source_seen_for_patronymic.setdefault(key, set()).add(source_name)

    if not candidates:
        return None

    freq = {}
    evidence = {}
    for patr, snippet in candidates:
        key = patr.strip()
        freq[key] = freq.get(key, 0) + 1
        evidence.setdefault(key, snippet)

    # Prefer consensus across multiple sources, then by total frequency.
    best_patronymic, count = max(
        freq.items(),
        key=lambda x: (len(source_seen_for_patronymic.get(x[0], set())), x[1]),
    )
    total = sum(freq.values())
    source_consensus = len(source_seen_for_patronymic.get(best_patronymic, set()))
    total_sources = 5.0
    confidence = ((count / total) * 0.65 + (source_consensus / total_sources) * 0.35) if total else 0.0
    confidence = min(1.0, confidence)
    source_hits_text = ", ".join(
        f"{src}:{source_hits_counter.get(src, 0)}"
        for src in ("duckduckgo", "bing", "yandex", "wikipedia", "serpapi")
    )

    return Suggestion(
        patronymic=best_patronymic,
        confidence=confidence,
        evidence=evidence[best_patronymic],
        query=query,
        source_hits=source_hits_text,
        sources_used=source_consensus,
    )


def build_query(record: NameRecord, context_values: List[str]) -> str:
    context = " ".join(v for v in context_values if v)[:140]
    if record.first_name:
        return (
            f'"{record.first_name} {record.surname}" "{record.surname} {record.first_name}" '
            f"отчество {context}"
        ).strip()
    return f'"{record.surname}" отчество {context}'.strip()


def infer_patronymic_for_record(record: NameRecord, context_values: List[str], delay_sec: float) -> Optional[Suggestion]:
    query = build_query(record, context_values)
    source_candidates: Dict[str, List[Tuple[str, str]]] = {}
    search_sources = {
        "duckduckgo": search_duckduckgo,
        "bing": search_bing,
        "yandex": search_yandex,
    }

    for source_name, search_fn in search_sources.items():
        try:
            html = search_fn(query)
            source_candidates[source_name] = extract_patronymic_candidates(
                html_text=html,
                first_name=record.first_name,
                surname=record.surname,
            )
        except Exception:
            # Keep processing if one source temporarily fails or blocks requests.
            source_candidates[source_name] = []

        if delay_sec > 0:
            time.sleep(delay_sec)

    serpapi_key = os.getenv("SERPAPI_KEY", "").strip()
    if serpapi_key:
        try:
            serp_text = search_serpapi(query, serpapi_key)
            source_candidates["serpapi"] = extract_patronymic_candidates(
                html_text=serp_text,
                first_name=record.first_name,
                surname=record.surname,
            )
        except Exception:
            source_candidates["serpapi"] = []

    # Additional source that provides richer person descriptions than SERP snippets.
    try:
        wiki_text = search_wikipedia(record.first_name, record.surname)
        source_candidates["wikipedia"] = extract_patronymic_candidates(
            html_text=wiki_text,
            first_name=record.first_name,
            surname=record.surname,
        )
    except Exception:
        source_candidates["wikipedia"] = []

    suggestion = pick_best_suggestion(source_candidates, query)
    return suggestion


def run(
    input_path: str,
    output_path: str,
    sheet_name: str,
    name_col_index: int,
    delay_sec: float,
) -> None:
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    if name_col_index < 0 or name_col_index >= df.shape[1]:
        raise IndexError(f"name_col_index={name_col_index} out of range for {df.shape[1]} columns")

    rows_output = []
    processed = 0

    for idx, row in df.iterrows():
        excel_row = idx + 2  # row 1 is header
        name_raw = _clean_text(row.iat[name_col_index])
        if not name_raw:
            continue

        record = parse_missing_patronymic_name(name_raw, excel_row)
        if record is None:
            continue

        context_values = [_clean_text(v) for col_i, v in enumerate(row.values) if col_i != name_col_index]
        suggestion = infer_patronymic_for_record(record, context_values, delay_sec=delay_sec)
        processed += 1

        rows_output.append(
            {
                "excel_row": record.row_excel,
                "original_name": record.original,
                "first_name": record.first_name,
                "surname": record.surname,
                "suggested_patronymic": suggestion.patronymic if suggestion else "",
                "confidence": round(suggestion.confidence, 3) if suggestion else 0.0,
                "suggested_full_name": (
                    f"{record.first_name} {suggestion.patronymic} {record.surname}"
                    if suggestion
                    else ""
                ),
                "search_query": suggestion.query if suggestion else build_query(record, context_values),
                "evidence_snippet": suggestion.evidence if suggestion else "",
                "sources_used_for_winner": suggestion.sources_used if suggestion else 0,
                "source_hits": (
                    suggestion.source_hits
                    if suggestion
                    else "duckduckgo:0, bing:0, yandex:0, wikipedia:0, serpapi:0"
                ),
            }
        )

        print(
            f"[{processed}] Row {record.row_excel}: {record.original} -> "
            f"{rows_output[-1]['suggested_patronymic'] or 'NOT FOUND'}"
        )

    out_df = pd.DataFrame(rows_output)
    out_df.to_excel(output_path, index=False)
    print(f"Done. Wrote {len(out_df)} rows to: {output_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find missing patronymics in Excel by searching the web."
    )
    parser.add_argument(
        "--input",
        default="Копия Копия CRM_industry_education_2026.xlsx",
        help="Path to input Excel workbook",
    )
    parser.add_argument(
        "--output",
        default="patronymic_suggestions.xlsx",
        help="Path to output Excel report with suggestions",
    )
    parser.add_argument(
        "--sheet",
        default="CRM_ЛПР",
        help="Sheet name",
    )
    parser.add_argument(
        "--name-col-index",
        type=int,
        default=1,
        help="Zero-based column index of full name (default 1 -> column B, ФИО)",
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.2,
        help="Delay between web requests in seconds",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    run(
        input_path=args.input,
        output_path=args.output,
        sheet_name=args.sheet,
        name_col_index=args.name_col_index,
        delay_sec=args.delay,
    )
