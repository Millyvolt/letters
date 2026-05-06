import argparse
import html
import os
import re
import time
import urllib.parse
import urllib.request
from dataclasses import dataclass
from typing import Dict, List, Optional, Set, Tuple

import pandas as pd


EMAIL_REGEX = re.compile(r"\b[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}\b")


@dataclass
class RowRecord:
    row_excel: int
    name: str
    company: str
    original_email: str


@dataclass
class Suggestion:
    email: str
    confidence: float
    evidence: str
    query: str
    source_hits: str
    sources_used: int


def _clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def _normalize_email(value: str) -> str:
    return value.strip().strip(".,;:()[]{}<>\"'").lower()


def _is_valid_email(value: str) -> bool:
    value = _normalize_email(value)
    return bool(EMAIL_REGEX.fullmatch(value))


def _extract_tokens_for_match(text: str) -> Set[str]:
    raw_tokens = re.split(r"[\s\.,;:!?\"'()\-_/\\]+", text.lower())
    return {tok for tok in raw_tokens if len(tok) >= 3}


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


def search_google(query: str) -> str:
    encoded = urllib.parse.quote_plus(query)
    url = f"https://www.google.com/search?q={encoded}&hl=ru&gl=ru"
    try:
        return _http_get(url)
    except Exception:
        return _http_get_via_jina(url)


def _strip_html(raw_text: str) -> str:
    plain = html.unescape(raw_text)
    plain = re.sub(r"<[^>]+>", " ", plain)
    return re.sub(r"\s+", " ", plain)


def extract_email_candidates(
    page_text: str,
    name: str,
    company: str,
) -> List[Tuple[str, str]]:
    plain = _strip_html(page_text)
    candidates: List[Tuple[str, str]] = []

    for match in EMAIL_REGEX.finditer(plain):
        email_text = _normalize_email(match.group(0))
        if not _is_valid_email(email_text):
            continue
        if email_text.endswith((".png", ".jpg", ".jpeg", ".svg", ".webp")):
            continue

        start = max(0, match.start() - 120)
        end = min(len(plain), match.end() + 120)
        snippet = plain[start:end].strip()
        candidates.append((email_text, snippet))

    # Also parse mailto links that can be present in search snippets.
    for mt in re.findall(r"mailto:([A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", plain, flags=re.IGNORECASE):
        email_text = _normalize_email(mt)
        if _is_valid_email(email_text):
            candidates.append((email_text, f"mailto:{email_text}"))

    return candidates


def _email_score(email_value: str, snippet: str, name: str, company: str) -> float:
    score = 0.0
    local_part, _, domain = email_value.partition("@")
    domain_base = domain.split(".")[0] if domain else ""

    # Prefer corporate addresses over common personal mailboxes.
    if domain.endswith((
        "gmail.com",
        "mail.ru",
        "yandex.ru",
        "bk.ru",
        "list.ru",
        "inbox.ru",
        "outlook.com",
        "hotmail.com",
    )):
        score += 0.1
    else:
        score += 0.35

    normalized_snippet = snippet.lower()
    for tok in _extract_tokens_for_match(name):
        if tok in local_part:
            score += 0.16
        if tok in normalized_snippet:
            score += 0.06

    for tok in _extract_tokens_for_match(company):
        if tok in domain_base:
            score += 0.2
        if tok in normalized_snippet:
            score += 0.08

    return min(score, 1.0)


def pick_best_suggestion(
    source_candidates: Dict[str, List[Tuple[str, str]]],
    query: str,
    name: str,
    company: str,
) -> Optional[Suggestion]:
    if not source_candidates:
        return None

    freq: Dict[str, int] = {}
    best_evidence: Dict[str, str] = {}
    weighted_score: Dict[str, float] = {}
    source_seen_for_email: Dict[str, Set[str]] = {}
    source_hits_counter: Dict[str, int] = {}

    for source_name, items in source_candidates.items():
        source_hits_counter[source_name] = len(items)
        for email_value, snippet in items:
            freq[email_value] = freq.get(email_value, 0) + 1
            best_evidence.setdefault(email_value, snippet)
            weighted_score[email_value] = max(
                weighted_score.get(email_value, 0.0),
                _email_score(email_value, snippet, name=name, company=company),
            )
            source_seen_for_email.setdefault(email_value, set()).add(source_name)

    if not freq:
        return None

    def _rank_key(email_value: str) -> Tuple[float, int, int]:
        source_consensus = len(source_seen_for_email.get(email_value, set()))
        return (
            weighted_score.get(email_value, 0.0),
            source_consensus,
            freq.get(email_value, 0),
        )

    best_email = max(freq.keys(), key=_rank_key)
    best_score = weighted_score.get(best_email, 0.0)
    source_consensus = len(source_seen_for_email.get(best_email, set()))
    confidence = min(1.0, best_score * 0.75 + (source_consensus / 5.0) * 0.25)

    source_hits_text = ", ".join(
        f"{src}:{source_hits_counter.get(src, 0)}"
        for src in ("duckduckgo", "bing", "yandex", "google")
    )

    return Suggestion(
        email=best_email,
        confidence=confidence,
        evidence=best_evidence[best_email],
        query=query,
        source_hits=source_hits_text,
        sources_used=source_consensus,
    )


def build_query(name: str, company: str) -> str:
    base = f"\"{name}\" \"{company}\" email контакты".strip()
    return re.sub(r"\s+", " ", base)


def infer_email_for_record(record: RowRecord, delay_sec: float) -> Optional[Suggestion]:
    query = build_query(record.name, record.company)
    source_candidates: Dict[str, List[Tuple[str, str]]] = {}
    search_sources = {
        "duckduckgo": search_duckduckgo,
        "bing": search_bing,
        "yandex": search_yandex,
        "google": search_google,
    }

    for source_name, search_fn in search_sources.items():
        try:
            page = search_fn(query)
            source_candidates[source_name] = extract_email_candidates(
                page_text=page,
                name=record.name,
                company=record.company,
            )
        except Exception:
            source_candidates[source_name] = []
        if delay_sec > 0:
            time.sleep(delay_sec)

    return pick_best_suggestion(
        source_candidates=source_candidates,
        query=query,
        name=record.name,
        company=record.company,
    )


def run(
    input_path: str,
    output_report_path: str,
    output_filled_path: str,
    sheet_name: str,
    email_col_index: int,
    name_col_index: int,
    company_col_index: int,
    delay_sec: float,
) -> None:
    df = pd.read_excel(input_path, sheet_name=sheet_name)
    if email_col_index < 0 or email_col_index >= df.shape[1]:
        raise IndexError(f"email_col_index={email_col_index} out of range for {df.shape[1]} columns")
    if name_col_index < 0 or name_col_index >= df.shape[1]:
        raise IndexError(f"name_col_index={name_col_index} out of range for {df.shape[1]} columns")
    if company_col_index < 0 or company_col_index >= df.shape[1]:
        raise IndexError(f"company_col_index={company_col_index} out of range for {df.shape[1]} columns")

    rows_output = []
    processed = 0
    filled_count = 0
    email_col_name = df.columns[email_col_index]

    for idx, row in df.iterrows():
        excel_row = idx + 2  # row 1 is headers
        current_email = _clean_text(row.iat[email_col_index])
        if current_email and _is_valid_email(current_email):
            continue

        name = _clean_text(row.iat[name_col_index])
        company = _clean_text(row.iat[company_col_index])
        if not name and not company:
            continue

        record = RowRecord(
            row_excel=excel_row,
            name=name,
            company=company,
            original_email=current_email,
        )
        suggestion = infer_email_for_record(record, delay_sec=delay_sec)
        processed += 1

        suggested_email = suggestion.email if suggestion else ""
        if suggested_email:
            df.at[idx, email_col_name] = suggested_email
            filled_count += 1

        rows_output.append(
            {
                "excel_row": record.row_excel,
                "name": record.name,
                "company": record.company,
                "original_email": record.original_email,
                "suggested_email": suggested_email,
                "confidence": round(suggestion.confidence, 3) if suggestion else 0.0,
                "search_query": suggestion.query if suggestion else build_query(record.name, record.company),
                "evidence_snippet": suggestion.evidence if suggestion else "",
                "sources_used_for_winner": suggestion.sources_used if suggestion else 0,
                "source_hits": suggestion.source_hits if suggestion else "duckduckgo:0, bing:0, yandex:0, google:0",
            }
        )

        print(
            f"[{processed}] Row {record.row_excel}: "
            f"{record.name or '(no name)'} -> {suggested_email or 'NOT FOUND'}"
        )

    os.makedirs(os.path.dirname(output_report_path) or ".", exist_ok=True)
    os.makedirs(os.path.dirname(output_filled_path) or ".", exist_ok=True)

    report_df = pd.DataFrame(rows_output)
    report_df.to_excel(output_report_path, index=False)
    df.to_excel(output_filled_path, index=False, sheet_name=sheet_name)

    print(f"Done. Processed missing emails: {processed}")
    print(f"Filled emails: {filled_count}")
    print(f"Report saved to: {output_report_path}")
    print(f"Updated workbook saved to: {output_filled_path}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Find missing emails in Excel by searching the web."
    )
    parser.add_argument(
        "--input",
        default="Копия Копия CRM_industry_education_2026.xlsx",
        help="Path to input Excel workbook",
    )
    parser.add_argument(
        "--output-report",
        default="email_suggestions.xlsx",
        help="Path to output Excel report with suggestions",
    )
    parser.add_argument(
        "--output-filled",
        default="Копия Копия CRM_industry_education_2026.filled_emails.xlsx",
        help="Path to output Excel workbook with inserted suggested emails",
    )
    parser.add_argument(
        "--sheet",
        default="CRM_ЛПР",
        help="Sheet name",
    )
    parser.add_argument(
        "--email-col-index",
        type=int,
        default=5,
        help="Zero-based email column index (default 5 => 6th column)",
    )
    parser.add_argument(
        "--name-col-index",
        type=int,
        default=1,
        help="Zero-based full-name column index (default 1 => column B)",
    )
    parser.add_argument(
        "--company-col-index",
        type=int,
        default=0,
        help="Zero-based company column index (default 0 => column A)",
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
        output_report_path=args.output_report,
        output_filled_path=args.output_filled,
        sheet_name=args.sheet,
        email_col_index=args.email_col_index,
        name_col_index=args.name_col_index,
        company_col_index=args.company_col_index,
        delay_sec=args.delay,
    )
