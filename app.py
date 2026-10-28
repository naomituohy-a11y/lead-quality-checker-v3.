import io
import re
import datetime
import unicodedata
import os
import json
import hashlib
from decimal import Decimal, InvalidOperation
from typing import Any, Dict, List, Tuple

import pandas as pd
import streamlit as st
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.styles import PatternFill

try:
    import phonenumbers
    from phonenumbers.phonenumberutil import NumberParseException
except Exception:
    phonenumbers = None
    NumberParseException = Exception

try:
    import psycopg2
    from psycopg2.extras import RealDictCursor
except Exception:
    psycopg2 = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="Wholesale Lead Quality Checker V2", page_icon="✅", layout="wide")

HEADER_YELLOW = PatternFill(start_color="FFF59D", end_color="FFF59D", fill_type="solid")
CELL_GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
CELL_BLUE = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
CELL_AMBER = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
CELL_RED = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
DASH_CHARS = "\u2010\u2011\u2012\u2013\u2014\u2212"

def get_secret(key: str, default: Any = None) -> Any:
    val = os.environ.get(key)
    if val:
        return val
    try:
        if hasattr(st, "secrets") and key in st.secrets:
            return st.secrets[key]
    except Exception:
        pass
    return default

def get_db_connection():
    db_url = get_secret("DATABASE_URL")
    if not db_url or not psycopg2:
        return None
    try:
        conn = psycopg2.connect(db_url)
        return conn
    except Exception:
        return None

def init_db():
    conn = get_db_connection()
    if not conn:
        return
    try:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS company_cache (
                    id SERIAL PRIMARY KEY,
                    normalized_company TEXT,
                    domain TEXT,
                    status TEXT,
                    explanation TEXT,
                    source_links TEXT,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    UNIQUE(normalized_company, domain)
                );
            """)
            conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

init_db()

def fetch_from_cache(company_name: str, domain: str) -> Tuple[str, str, List[str]]:
    conn = get_db_connection()
    if not conn:
        return "", "", []
    try:
        norm_c = unicodedata.normalize("NFKC", str(company_name)).strip().casefold()
        clean_d = str(domain).strip().casefold()
        with conn.cursor() as cur:
            cur.execute(
                "SELECT status, explanation, source_links FROM company_cache WHERE normalized_company = %s AND domain = %s;",
                (norm_c, clean_d)
            )
            row = cur.fetchone()
            if row:
                status, explanation, links_json = row
                links = json.loads(links_json) if links_json else []
                return status, explanation, links
    except Exception:
        pass
    finally:
        conn.close()
    return "", "", []

def save_to_cache(company_name: str, domain: str, status: str, explanation: str, links: List[str]):
    conn = get_db_connection()
    if not conn:
        return
    try:
        norm_c = unicodedata.normalize("NFKC", str(company_name)).strip().casefold()
        clean_d = str(domain).strip().casefold()
        links_json = json.dumps(links)
        with conn.cursor() as cur:
            cur.execute("""
                INSERT INTO company_cache (normalized_company, domain, status, explanation, source_links, updated_at)
                VALUES (%s, %s, %s, %s, %s, CURRENT_TIMESTAMP)
                ON CONFLICT (normalized_company, domain) 
                DO UPDATE SET status = EXCLUDED.status, explanation = EXCLUDED.explanation, 
                              source_links = EXCLUDED.source_links, updated_at = CURRENT_TIMESTAMP;
            """, (norm_c, clean_d, status, explanation, links_json))
            conn.commit()
    except Exception:
        pass
    finally:
        conn.close()

def norm_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00A0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text.casefold()

def norm_key(value: Any) -> str:
    text = norm_text(value)
    text = text.replace("&", " and ")
    for ch in DASH_CHARS:
        text = text.replace(ch, " ")
    text = re.sub(r"[\\/]", " ", text)
    text = text.replace("’", "'")
    text = re.sub(r"[^a-z0-9+#.\s]", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text

def clean_header(value: Any) -> str:
    text = "" if value is None else str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00A0", " ")
    text = text.strip().lower().replace("_", " ")
    return re.sub(r"[^a-z0-9]+", " ", text).strip()

PLACEHOLDER_VALUES = {"", "picklist", "leave blank", "blank", "integer", "text", "date", "dd mm yyyy", "mm dd yyyy", "https www", "http www", "do not map"}

def is_placeholder(value: Any) -> bool:
    key = norm_key(value)
    raw = norm_text(value)
    if key in PLACEHOLDER_VALUES:
        return True
    return raw.startswith(("all accepted", "target the below", "no proof", "no toll", "no fee", "please map"))

def looks_like_template_row(row: pd.Series) -> bool:
    values = [str(v).strip() for v in row.tolist() if str(v).strip()]
    if not values:
        return True
    joined = " | ".join(values[:40]).casefold()
    markers = ["picklist", "leave blank", "integer", "text", "dd/mm/yyyy", "mm/dd/yyyy", "https://", "no toll", "fee phone", "target the below", "all accepted", "do not map", "please map"]
    if sum(1 for marker in markers if marker in joined) >= 2:
        return True
    placeholder_count = sum(1 for value in values if is_placeholder(value))
    return bool(values and placeholder_count / len(values) >= 0.6)

SYNONYM_GROUPS = [
    ["us", "usa", "u s", "u s a", "united states", "united states of america", "america"],
    ["uk", "u k", "gb", "gbr", "great britain", "united kingdom", "england", "britain"],
    ["germany", "de", "deu", "deutschland"],
    ["it", "information technology", "technology"],
    ["it operations", "information technology operations", "it ops", "technology operations"],
    ["hr", "human resources", "people"],
    ["vp", "vice president"],
    ["svp", "senior vice president"],
    ["evp", "executive vice president"],
    ["c level", "c-level", "c suite", "c-suite", "chief"],
    ["ceo", "chief executive officer"],
    ["cfo", "chief financial officer"],
    ["cio", "chief information officer"],
    ["cto", "chief technology officer"],
    ["coo", "chief operating officer"],
    ["ciso", "chief information security officer"],
    ["biz dev", "business development", "bd"],
    ["ops", "operations"],
    ["manufacturing", "manufacturing and process industries", "manufacturing process industries"],
]

def canonical_key(value: Any) -> str:
    key = norm_key(value)
    if not key:
        return ""
    for group in SYNONYM_GROUPS:
        group_keys = [norm_key(item) for item in group]
        if key in group_keys:
            return group_keys[0]
    return key

def expanded_keys(value: Any) -> List[str]:
    key = norm_key(value)
    canonical = canonical_key(value)
    output: List[str] = []
    for item in [key, canonical]:
        if item and item not in output:
            output.append(item)
    for group in SYNONYM_GROUPS:
        group_keys = [norm_key(item) for item in group]
        if key in group_keys or canonical in group_keys:
            for item in group_keys:
                if item and item not in output:
                    output.append(item)
    return output

FIELD_ALIASES: Dict[str, List[str]] = {
    "company": ["company", "company name", "account", "account name", "organisation", "organization", "business name", "companyname"],
    "email": ["email", "email address", "work email", "contact email", "person email"],
    "country": ["country", "lead country", "company country", "country code", "cntry"],
    "region": ["region", "state", "province", "c state", "company state", "lead state"],
    "industry": ["industry", "companyindustry", "company industry", "main industry", "indcode1", "indcode1 main industry", "c industry"],
    "sub_industry": ["sub industry", "subindustry", "indcode2", "indcode2 sub industry"],
    "function": ["function", "department", "departments", "job function", "job area", "ocpcode1", "ocpcode1 job area", "ocpcode2", "ocpcode2 job function"],
    "job_level": ["position", "job level", "job_level", "seniority", "level", "ocpcode3", "ocpcode3 job level"],
    "job_role": ["job role", "role", "job_role", "job_role__c"],
    "company_size": ["companysize", "company size", "number of employees", "number_of_employees", "employees", "orgemp", "orgemp number of employees"],
    "phone": ["phone", "telephone", "mobile", "cell", "phonemain", "phone main", "phone_1"],
    "website": ["website", "domain", "url", "companyurl", "company url", "web", "company website"],
    "job_title": ["job title", "job_title", "jobtitle", "title", "jobtitletext", "job title text"],
}

def score_header(header: str, alias: str) -> int:
    header_clean = clean_header(header)
    alias_clean = clean_header(alias)
    if not header_clean or not alias_clean:
        return 0
    if header_clean == alias_clean:
        return 100
    if alias_clean in header_clean:
        return 85
    header_tokens = set(header_clean.split())
    alias_tokens = set(alias_clean.split())
    if alias_tokens and alias_tokens.issubset(header_tokens):
        return 80
    return int(fuzz.token_sort_ratio(header_clean, alias_clean))

def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
    result: Dict[str, str] = {}
    used_columns = set()
    for field, aliases in FIELD_ALIASES.items():
        best_col = ""
        best_score = 0
        for col in df.columns:
            if col in used_columns:
                continue
            header = clean_header(col)
            if field == "country" and "county" in header.split():
                continue
            for alias in aliases:
                score = score_header(str(col), alias)
                if score > best_score:
                    best_score = score
                    best_col = str(col)
        if best_col and best_score >= 72:
            result[field] = best_col
            used_columns.add(best_col)
    return result

def field_for_picklist_column(column_name: Any) -> str:
    header = clean_header(column_name)
    if not header or "county" in header.split():
        return ""
    best_field = ""
    best_score = 0
    for field, aliases in FIELD_ALIASES.items():
        for alias in aliases:
            score = score_header(str(column_name), alias)
            if score > best_score:
                best_score = score
                best_field = field
    return best_field if best_score >= 72 else ""

def is_value_or_code_column(column_name: Any) -> bool:
    header = clean_header(column_name)
    return header == "value" or header.startswith("value ") or header == "code" or header.endswith(" code") or " value " in f" {header} "

def add_allowed(allowed: Dict[str, Dict[str, str]], field: str, value: Any) -> None:
    if not field or is_placeholder(value):
        return
    text = str(value).strip()
    if not text or len(text) > 120:
        return
    allowed.setdefault(field, {})
    for key in expanded_keys(text):
        allowed[field].setdefault(key, text)

def extract_picklist_rules_from_df(df_pick: pd.DataFrame) -> Dict[str, Any]:
    allowed: Dict[str, Dict[str, str]] = {}
    mapping_pairs = []
    toll_free_prefixes = []
    columns = list(df_pick.columns)

    for col in columns:
        field = field_for_picklist_column(col)
        if not field:
            continue
        for value in df_pick[col].tolist():
            add_allowed(allowed, field, value)

    for index, col in enumerate(columns):
        field = field_for_picklist_column(col)
        if not field:
            continue
        nearby = columns[index + 1:index + 4]
        for partner in nearby:
            if is_value_or_code_column(partner):
                mapping_pairs.append((str(col), str(partner), field))
                for label, code in zip(df_pick[col].tolist(), df_pick[partner].tolist()):
                    add_allowed(allowed, field, label)
                    add_allowed(allowed, field, code)
                break

    for col in columns:
        for raw in df_pick[col].tolist():
            value = str(raw or "").strip()
            digits = re.sub(r"[^\d+]", "", value)
            if re.match(r"^(\+?1)?8(00|33|44|55|66|77|88)", digits):
                toll_free_prefixes.append(digits)

    return {
        "allowed": allowed,
        "allowed_counts": {field: len(values) for field, values in allowed.items()},
        "mapping_pairs": mapping_pairs,
        "samples": {field: list(values.values())[:10] for field, values in allowed.items()},
        "toll_free_prefixes": sorted(set(toll_free_prefixes)),
    }

def merge_rule_sets(rule_sets: List[Dict[str, Any]]) -> Dict[str, Any]:
    merged_allowed: Dict[str, Dict[str, str]] = {}
    mapping_pairs = []
    toll_free_prefixes = []
    for rules in rule_sets:
        for field, values in rules.get("allowed", {}).items():
            merged_allowed.setdefault(field, {})
            merged_allowed[field].update(values)
        mapping_pairs.extend(rules.get("mapping_pairs", []))
        toll_free_prefixes.extend(rules.get("toll_free_prefixes", []))
    return {
        "allowed": merged_allowed,
        "allowed_counts": {field: len(values) for field, values in merged_allowed.items()},
        "mapping_pairs": mapping_pairs,
        "samples": {field: list(values.values())[:10] for field, values in merged_allowed.items()},
        "toll_free_prefixes": sorted(set(toll_free_prefixes)),
    }

def match_value(value: Any, allowed_map: Dict[str, str], allow_fuzzy: bool = True) -> Tuple[str, str, int]:
    raw = str(value or "").strip()
    if not raw:
        return "Review", "blank value", 0
    if not allowed_map:
        return "Rule Missing", "no picklist values detected", 0
    for key in expanded_keys(raw):
        if key in allowed_map:
            return "Match", f"matched to {allowed_map[key]}", 100
    if allow_fuzzy:
        match = process.extractOne(norm_key(raw), list(allowed_map.keys()), scorer=fuzz.token_sort_ratio)
        if match:
            matched_key = match[0]
            score = int(match[1])
            if score >= 90:
                return "Match", f"strong fuzzy match to {allowed_map[matched_key]} ({score})", score
            if score >= 75:
                return "Review", f"possible fuzzy match to {allowed_map[matched_key]} ({score})", score
    return "No Match", "not found in picklist", 0

COMPANY_SUFFIXES = {"ltd", "limited", "co", "company", "corp", "corporation", "inc", "incorporated", "plc", "llc", "sa", "ag", "nv", "se", "bv", "oy", "ab", "aps", "as", "sarl", "sas", "spa", "gmbh", "pte", "pty", "sdn", "bhd", "holdings", "holding", "group"}

def email_domain(email: Any) -> str:
    text = str(email or "").strip()
    if "@" not in text:
        return ""
    return text.split("@", 1)[1].strip().lower()

def clean_domain(domain: Any) -> str:
    text = str(domain or "").strip().lower()
    text = re.sub(r"^https?://", "", text)
    text = re.sub(r"/.*$", "", text)
    return text.replace("www.", "")

def domain_base(domain: Any) -> str:
    text = clean_domain(domain)
    if not text:
        return ""
    return re.sub(r"[^a-z0-9]", "", text.split(".")[0])

def company_tokens(company: Any) -> List[str]:
    text = str(company or "")
    text = unicodedata.normalize("NFKC", text)
    text = re.sub(r"[^A-Za-z0-9\s]", " ", text)
    tokens = [token.lower() for token in text.split() if token.strip()]
    tokens = [token for token in tokens if token not in COMPANY_SUFFIXES]
    tokens = [token for token in tokens if token not in {"of", "and", "the", "for", "to", "a", "an"}]
    return tokens

def compare_company_domain(company: Any, email: Any, website: Any) -> Tuple[str, str, int]:
    company_text = str(company or "").strip()
    domain = email_domain(email) or clean_domain(website)
    if not company_text or not domain:
        return "Review", "missing company or domain", 0
    base = domain_base(domain)
    if not base:
        return "Review", "invalid domain", 0
    tokens = company_tokens(company_text)
    joined = "".join(tokens)
    if joined and joined in base:
        return "Match", "company tokens contained in domain", 95
    token_string = " ".join(tokens)
    score = int(max(fuzz.token_sort_ratio(token_string, base), fuzz.partial_ratio(token_string, base)))
    if score >= 85:
        return "Match", f"strong fuzzy company/domain match ({score})", score
    if score >= 70:
        return "Review", f"weak fuzzy company/domain match ({score})", score
    return "No Match", f"low company/domain similarity ({score})", score

def run_ai_web_research(company: str, domain: str, country: str, api_key: str) -> Tuple[str, str, List[str]]:
    if not OpenAI or not api_key:
        return "Review", "OpenAI key missing or library not installed", []
    
    cached_status, cached_expl, cached_links = fetch_from_cache(company, domain)
    if cached_status:
        return cached_status, cached_expl + " (Loaded from DB Cache)", cached_links

    try:
        client = OpenAI(api_key=api_key)
        response = client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[
                {
                    "role": "system",
                    "content": "You are a corporate data validation assistant. Analyze if a company name and domain match, or if they are connected through rebranding, acquisition, or parent companies. Return strict JSON with keys: 'status' (VERIFIED_MATCH, REBRAND, or UNVERIFIED), 'explanation' (short text), and 'sources' (list of URL strings)."
                },
                {
                    "role": "user",
                    "content": f"Company Name: {company}\nDomain/Website: {domain}\nCountry: {country}\n\nAre these connected?"
                }
            ],
            response_format={"type": "json_object"}
        )
        data = json.loads(response.choices[0].message.content)
        status = data.get("status", "UNVERIFIED")
        explanation = data.get("explanation", "AI research completed.")
        sources = data.get("sources", [])
        
        save_to_cache(company, domain, status, explanation, sources)
        return status, explanation, sources
    except Exception as e:
        return "Review", f"AI research failed: {str(e)}", []

def phone_to_string(raw_phone: Any) -> str:
    if raw_phone is None:
        return ""
    text = str(raw_phone).strip()
    if not text:
        return ""
    if "e" in text.lower():
        try:
            text = format(int(Decimal(text)), "d")
        except (InvalidOperation, ValueError):
            pass
    if re.fullmatch(r"\d+\.0+", text):
        text = text.split(".", 1)[0]
    return text

def normalise_phone(raw_phone: Any) -> str:
    phone = phone_to_string(raw_phone)
    phone = re.sub(r"[^\d+]", "", phone)
    if phone.startswith("00"):
        phone = "+" + phone[2:]
    return phone

def is_toll_free(phone: Any, extra_prefixes: List[str]) -> bool:
    cleaned = normalise_phone(phone)
    digits = re.sub(r"\D", "", cleaned)
    built_in = ["1800", "1888", "1877", "1866", "1855", "1844", "1833", "800", "888", "877", "866", "855", "844", "833"]
    for prefix in built_in + extra_prefixes:
        prefix_digits = re.sub(r"\D", "", str(prefix))
        if prefix_digits and digits.startswith(prefix_digits):
            return True
    return False

def phone_check(raw_phone: Any, toll_free_prefixes: List[str]) -> Tuple[str, str, int]:
    phone = normalise_phone(raw_phone)
    if not phone:
        return "Review", "missing phone", 0
    if is_toll_free(phone, toll_free_prefixes):
        return "No Match", "toll-free / excluded phone prefix", 0
    if phonenumbers is None:
        return "Review", "phonenumbers library not available", 0
    try:
        parsed = phonenumbers.parse(phone, None if phone.startswith("+") else "US")
        region = phonenumbers.region_code_for_number(parsed) or ""
        if not phonenumbers.is_possible_number(parsed):
            return "No Match", f"number not possible (parsed region={region})", 0
        if not phonenumbers.is_valid_number(parsed):
            return "Review", f"number not valid (parsed region={region})", 50
        return "Match", f"valid number (parsed region={region})", 100
    except Exception:
        return "Review", "could not parse phone", 0

TITLE_TERMS = ["chief", "ceo", "cfo", "cio", "cto", "coo", "ciso", "president", "vice president", "vp", "svp", "evp", "director", "head", "manager", "technology", "information technology", "operations", "digital", "data", "security", "infrastructure", "systems"]

def title_relevance(title: Any, function: Any, level: Any) -> Tuple[str, str, int]:
    combined = f" {norm_key(title)} {norm_key(function)} {norm_key(level)} "
    if not combined.strip():
        return "Review", "missing title/function/level", 0
    hits = []
    score = 0
    for term in TITLE_TERMS:
        if f" {norm_key(term)} " in combined:
            hits.append(term)
            if term in {"chief", "ceo", "cfo", "cio", "cto", "coo", "ciso"}:
                score = max(score, 100)
            elif term in {"president", "vice president", "vp", "svp", "evp"}:
                score = max(score, 90)
            elif term in {"director", "head"}:
                score = max(score, 75)
            elif term == "manager":
                score = max(score, 60)
            else:
                score = max(score, 65)
    if score >= 85:
        return "Match", "strong target signal: " + ", ".join(sorted(set(hits))), score
    if score >= 60:
        return "Review", "possible target signal: " + ", ".join(sorted(set(hits))), score
    return "No Match", "no clear target title signal", score

WEIGHTS = {"country": 15, "industry": 15, "company_size": 10, "function": 15, "job_level": 15, "title": 15, "domain": 10, "phone": 5}

def points_for_status(status: str, weight: int) -> float:
    status_clean = norm_key(status)
    if status_clean in {"match", "verified_match", "rebrand", "affiliated"}:
        return float(weight)
    if status_clean in {"review", "rule missing", "column missing", "unverified"}:
        return weight * 0.5
    return 0.0

def overall_status(score: float) -> str:
    if score >= 85:
        return "PASS"
    if score >= 60:
        return "REVIEW"
    return "FAIL"

QA_COLUMNS = [
    "QA_Country_Status", "QA_Country_Reason", "QA_Country_Score",
    "QA_Industry_Status", "QA_Industry_Reason", "QA_Industry_Score",
    "QA_Company_Size_Status", "QA_Company_Size_Reason", "QA_Company_Size_Score",
    "QA_Function_Status", "QA_Function_Reason", "QA_Function_Score",
    "QA_Job_Level_Status", "QA_Job_Level_Reason", "QA_Job_Level_Score",
    "QA_Region_Status", "QA_Region_Reason", "QA_Region_Score",
    "QA_Job_Role_Status", "QA_Job_Role_Reason", "QA_Job_Role_Score",
    "QA_Sub_Industry_Status", "QA_Sub_Industry_Reason", "QA_Sub_Industry_Score",
    "QA_Title_Relevance_Status", "QA_Title_Relevance_Reason", "QA_Title_Relevance_Score",
    "QA_Domain_Status", "QA_Domain_Reason", "QA_Domain_Score",
    "QA_Phone_Status", "QA_Phone_Reason", "QA_Phone_Score",
    "AI_Explanation", "AI_Source_Links",
    "QA_Missing_Fields", "QA_Score", "QA_Overall_Status", "QA_Issues", "QA_Debug_Notes",
]

def fill_for_value(value: Any) -> PatternFill:
    text = norm_key(value)
    if text in {"pass", "match", "yes", "verified_match", "rebrand", "affiliated"}:
        return CELL_GREEN
    if text in {"review", "rule missing", "column missing", "unverified"}:
        return CELL_AMBER
    if text in {"fail", "no match", "no"}:
        return CELL_RED
    if not text:
        return CELL_GREEN
    return CELL_BLUE

def write_results_to_workbook(master_bytes: bytes, sheet_name: str, results: pd.DataFrame, qa_columns: List[str], apply_colours: bool) -> bytes:
    workbook = load_workbook(io.BytesIO(master_bytes))
    worksheet = workbook[sheet_name]
    worksheet.title = "Results"
    start_col = worksheet.max_column + 1
    for index, column_name in enumerate(qa_columns):
        col_number = start_col + index
        worksheet.cell(row=1, column=col_number).value = column_name
        if apply_colours:
            worksheet.cell(row=1, column=col_number).fill = HEADER_YELLOW
    column_positions = {column_name: start_col + index for index, column_name in enumerate(qa_columns)}
    for _, result_row in results.iterrows():
        excel_row = int(result_row["_excel_row"])
        for column_name in qa_columns:
            col_number = column_positions[column_name]
            value = result_row.get(column_name, "")
            worksheet.cell(row=excel_row, column=col_number).value = value
            if apply_colours:
                if column_name.endswith("_Status") or column_name == "QA_Overall_Status":
                    worksheet.cell(row=excel_row, column=col_number).fill = fill_for_value(value)
                elif column_name == "QA_Issues":
                    worksheet.cell(row=excel_row, column=col_number).fill = CELL_GREEN if not value else CELL_AMBER
                elif column_name == "QA_Score":
                    try:
                        numeric_score = float(value)
                        if numeric_score >= 85:
                            worksheet.cell(row=excel_row, column=col_number).fill = CELL_GREEN
                        elif numeric_score >= 60:
                            worksheet.cell(row=excel_row, column=col_number).fill = CELL_AMBER
                        else:
                            worksheet.cell(row=excel_row, column=col_number).fill = CELL_RED
                    except Exception:
                        worksheet.cell(row=excel_row, column=col_number).fill = CELL_BLUE
    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return output.read()

def process_file(master_bytes: bytes, picklist_bytes: bytes, master_sheet: str, picklist_sheets: List[str], apply_colours: bool, use_ai: bool, openai_key: str, max_ai_calls: int) -> Tuple[bytes, Dict[str, Any]]:
    master_df = pd.read_excel(io.BytesIO(master_bytes), sheet_name=master_sheet, dtype=str, keep_default_na=False)
    rule_sets = []
    
    if picklist_bytes and picklist_sheets:
        for pick_sheet in picklist_sheets:
            picklist_df = pd.read_excel(io.BytesIO(picklist_bytes), sheet_name=pick_sheet, dtype=str, keep_default_na=False)
            rule_sets.append(extract_picklist_rules_from_df(picklist_df))
            
    rules = merge_rule_sets(rule_sets) if rule_sets else {"allowed": {}, "allowed_counts": {}, "mapping_pairs": [], "samples": {}, "toll_free_prefixes": []}
    colmap = detect_columns(master_df)
    results = []
    skipped_rows = []
    ai_calls_made = 0

    def get_value(row: pd.Series, field: str) -> str:
        column = colmap.get(field, "")
        if column and column in master_df.columns:
            return str(row.get(column, "") or "").strip()
        return ""

    for index, row in master_df.iterrows():
        excel_row = index + 2
        if looks_like_template_row(row):
            skipped_rows.append(excel_row)
            continue
        result: Dict[str, Any] = {"_excel_row": excel_row}
        issues = []
        total_score = 0.0
        field_checks = [("country", "QA_Country", WEIGHTS["country"], False), ("industry", "QA_Industry", WEIGHTS["industry"], True), ("company_size", "QA_Company_Size", WEIGHTS["company_size"], True), ("function", "QA_Function", WEIGHTS["function"], True), ("job_level", "QA_Job_Level", WEIGHTS["job_level"], True)]
        for field, prefix, weight, allow_fuzzy in field_checks:
            if field not in colmap:
                status, reason, score = "Column Missing", "master column not detected", 0
            else:
                allowed_map = rules["allowed"].get(field, {})
                status, reason, score = match_value(get_value(row, field), allowed_map, allow_fuzzy)
            result[f"{prefix}_Status"] = status
            result[f"{prefix}_Reason"] = reason
            result[f"{prefix}_Score"] = score
            total_score += points_for_status(status, weight)
            if status != "Match":
                issues.append(field)

        for optional_field, prefix in [("region", "QA_Region"), ("job_role", "QA_Job_Role"), ("sub_industry", "QA_Sub_Industry")]:
            if optional_field not in colmap:
                status, reason, score = "Column Missing", "master column not detected", 0
            else:
                allowed_map = rules["allowed"].get(optional_field, {})
                status, reason, score = match_value(get_value(row, optional_field), allowed_map, True)
            result[f"{prefix}_Status"] = status
            result[f"{prefix}_Reason"] = reason
            result[f"{prefix}_Score"] = score

        title_status, title_reason, title_score = title_relevance(get_value(row, "job_title"), get_value(row, "function"), get_value(row, "job_level"))
        result["QA_Title_Relevance_Status"] = title_status
        result["QA_Title_Relevance_Reason"] = title_reason
        result["QA_Title_Relevance_Score"] = title_score
        total_score += points_for_status(title_status, WEIGHTS["title"])
        if title_status != "Match":
            issues.append("title")

        domain_status, domain_reason, domain_score = compare_company_domain(get_value(row, "company"), get_value(row, "email"), get_value(row, "website"))
        ai_expl = ""
        ai_links = []

        if domain_status in ["No Match", "Review"] and use_ai and ai_calls_made < max_ai_calls:
            comp_val = get_value(row, "company")
            web_val = email_domain(get_value(row, "email")) or clean_domain(get_value(row, "website"))
            cntry_val = get_value(row, "country")
            if comp_val and web_val:
                ai_calls_made += 1
                ai_status, ai_expl, ai_links = run_ai_web_research(comp_val, web_val, cntry_val, openai_key)
                
                if ai_status in {"VERIFIED_MATCH", "REBRAND"}:
                    domain_status = "Affiliated"
                    domain_reason = f"AI Verified: {ai_expl}"
                    domain_score = 90

        result["QA_Domain_Status"] = domain_status
        result["QA_Domain_Reason"] = domain_reason
        result["QA_Domain_Score"] = domain_score
        result["AI_Explanation"] = ai_expl
        result["AI_Source_Links"] = ", ".join(ai_links)
        
        total_score += points_for_status(domain_status, WEIGHTS["domain"])
        
        if domain_status not in {"Match", "VERIFIED_MATCH", "REBRAND", "Affiliated"}:
            issues.append("domain")

        phone_status, phone_reason, phone_score = phone_check(get_value(row, "phone"), rules.get("toll_free_prefixes", []))
        result["QA_Phone_Status"] = phone_status
        result["QA_Phone_Reason"] = phone_reason
        result["QA_Phone_Score"] = phone_score
        total_score += points_for_status(phone_status, WEIGHTS["phone"])
        if phone_status not in {"Match", "Review"}:
            issues.append("phone")

        missing = []
        for required_field in ["email", "company", "country"]:
            if required_field not in colmap:
                missing.append(f"{required_field} column missing")
            elif not get_value(row, required_field):
                missing.append(required_field)
        result["QA_Missing_Fields"] = ", ".join(missing)
        if missing:
            issues.append("missing required")

        result["QA_Score"] = round(total_score, 1)
        result["QA_Overall_Status"] = overall_status(total_score)
        result["QA_Issues"] = "; ".join(sorted(set(issues)))
        result["QA_Debug_Notes"] = f"Detected columns: {colmap}"
        results.append(result)

    results_df = pd.DataFrame(results)
    if results_df.empty:
        results_df = pd.DataFrame(columns=["_excel_row"] + QA_COLUMNS)
    output_bytes = write_results_to_workbook(master_bytes=master_bytes, sheet_name=master_sheet, results=results_df, qa_columns=QA_COLUMNS, apply_colours=apply_colours)
    debug = {"detected_columns": colmap, "allowed_counts": rules.get("allowed_counts", {}), "mapping_pairs": rules.get("mapping_pairs", [])[:100], "picklist_samples": rules.get("samples", {}), "toll_free_prefixes": rules.get("toll_free_prefixes", []), "skipped_template_rows": skipped_rows, "processed_rows": len(results_df), "ai_calls_executed": ai_calls_made}
    return output_bytes, debug


# Industry mapping is deliberately independent of legacy QA rules and scoring.
# No network calls; no fuzzy industry assignment; original data is never replaced.
INDUSTRY_MAPPING_COLUMNS = [
    'Industry_Map_Source', 'Industry_Map_Value', 'Industry_Map_Status',
    'Industry_Map_Reason', 'Industry_Map_Candidates', 'Industry_Map_Selected_Candidates',
    'Industry_Map_Priority', 'Industry_Map_Rules_Version',
]


def industry_key(value):
    if value is None:
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        pass
    text = unicodedata.normalize('NFKC', str(value)).casefold().replace('&', ' and ')
    return re.sub(r'[^a-z0-9]+', ' ', text).strip()


def industry_flag(value, field, row_number):
    key = industry_key(value)
    if key in {'yes', 'true', '1', 'approved'}:
        return True
    if key in {'no', 'false', '0', '', 'draft', 'pending'}:
        return False
    raise ValueError(f'Mapping row {row_number}: {field} must be Yes or No.')


def compile_industry_rules(frame):
    required = ['Sub-industry', 'Leadscale industry', 'Priority', 'Approved', 'Context required']
    missing = [name for name in required if name not in frame.columns]
    if missing:
        raise ValueError('Mapping grid missing columns: ' + ', '.join(missing))
    index = {}
    seen = {}
    for offset, (_, row) in enumerate(frame.iterrows(), 2):
        if all(not industry_key(v) for v in row.tolist()):
            continue
        source = str(row.get('Sub-industry', '')).strip()
        target = str(row.get('Leadscale industry', '')).strip()
        if not industry_key(source) or not industry_key(target):
            raise ValueError(f'Mapping row {offset}: source and target are required.')
        if ';' in target:
            raise ValueError(f'Mapping row {offset}: use one target per row, not semicolons.')
        try:
            priority = float(row['Priority'])
            if not (priority >= 1 and priority < float('inf') and priority.is_integer()):
                raise ValueError()
        except (TypeError, ValueError, OverflowError):
            raise ValueError(f'Mapping row {offset}: Priority must be a positive whole number.')
        approved = industry_flag(row['Approved'], 'Approved', offset)
        context = industry_flag(row['Context required'], 'Context required', offset)
        record = {'source': source, 'target': target, 'priority': int(priority),
                  'approved': approved, 'context': context}
        pair = (industry_key(source), industry_key(target))
        if pair in seen:
            previous = seen[pair]
            if any(previous[k] != record[k] for k in ['priority', 'approved', 'context']):
                raise ValueError(f'Mapping row {offset}: conflicting duplicate rule for {source} / {target}.')
            continue
        seen[pair] = record
        index.setdefault(pair[0], []).append(record)
    if not index:
        raise ValueError('The mapping grid has no rules.')
    return index


def compile_campaign_industries(values):
    # Read the explicitly selected campaign column, not the legacy pooled QA rules.
    labels = {}
    placeholders = {'picklist', 'industry', 'industries', 'value', 'leave blank', 'blank'}
    for raw in values:
        key = industry_key(raw)
        if not key or key in placeholders:
            continue
        label = str(raw).strip()
        if ';' in label or '\n' in label or '\r' in label:
            raise ValueError('Campaign industries must have one industry per cell; split combined cells first.')
        if key.startswith(('all accepted', 'all industries', 'target the below', 'please map')):
            raise ValueError('Campaign industry instructions were found. Select a column containing only selected industry labels.')
        if key in labels and labels[key] != label:
            raise ValueError('Two campaign labels normalize to the same industry: ' + labels[key] + ' / ' + label)
        labels[key] = label
    if not labels:
        raise ValueError('No selected campaign industries were found.')
    return labels



INDUSTRY_LABEL_ALIASES = [
    ["Computer Software", "Software", "Software Development"],
    ["Information Technology & Services", "IT Services", "Information Technology Services", "IT Services and IT Consulting"],
    ["Hospital & Health Care", "Hospitals and Health Care", "Hospital and Healthcare", "Hospitals & Healthcare"],
    ["Marketing & Advertising", "Advertising and Marketing", "Marketing and Advertising"],
    ["Transportation/Trucking/Railroad", "Transportation, Trucking and Railroad"],
    ["Nonprofit Organization Management", "Non-profit Organization Management", "Nonprofit Organisation Management"],
    ["Electrical/Electronic Manufacturing", "Electrical and Electronic Manufacturing"],
    ["Health, Wellness & Fitness", "Health Wellness and Fitness"],
    ["Staffing & Recruiting", "Staffing and Recruitment", "Staffing and Recruiting"],
    ["E-learning", "E Learning", "E-Learning Providers"],
    ["Airlines/Aviation", "Airlines and Aviation"],
    ["Fishery", "Fisheries"],
    ["Medical Devices", "Medical Device"],
    ["Renewables & Environment", "Renewables and Environment"],
]
INDUSTRY_ALIAS_KEYS = {}
for _industry_group in INDUSTRY_LABEL_ALIASES:
    _industry_keys = {industry_key(v) for v in _industry_group}
    for _industry_alias in _industry_keys:
        INDUSTRY_ALIAS_KEYS[_industry_alias] = _industry_keys


def match_campaign_industry(label, campaign):
    """Return exact client label only for exact, approved alias or clear typo."""
    key = industry_key(label)
    if key in campaign:
        return campaign[key], "exact", []
    aliases = INDUSTRY_ALIAS_KEYS.get(key, set())
    hits = sorted({campaign[k] for k in aliases if k in campaign}, key=industry_key)
    if len(hits) == 1:
        return hits[0], "known alternative name", []
    if len(hits) > 1:
        return "", "ambiguous alternative names", hits
    scored = sorted([(fuzz.ratio(key, k), k) for k in campaign], reverse=True)
    if not scored:
        return "", "no match", []
    best, candidate = scored[0]
    runner = scored[1][0] if len(scored) > 1 else 0
    # Auto-accept only very small spelling differences with a clear lead.
    # Semantic equivalents belong in the explicit alias table, not fuzzy guessing.
    if min(len(key), len(candidate)) >= 8 and best >= 96 and best - runner >= 8:
        return campaign[candidate], "high-confidence spelling match", []
    suggestions = [campaign[k] for score,k in scored[:3] if score >= 75]
    return "", "uncertain wording" if suggestions else "no match", suggestions

def resolve_industry(source, rule_index, campaign, version=''):
    source = '' if source is None else str(source).strip()
    result = dict.fromkeys(INDUSTRY_MAPPING_COLUMNS, '')
    result.update(Industry_Map_Source=source, Industry_Map_Rules_Version=version)
    def finish(status, reason, target='', priority=''):
        result.update(Industry_Map_Status=status, Industry_Map_Reason=reason,
                      Industry_Map_Value=target, Industry_Map_Priority=priority)
        return result
    key = industry_key(source)
    if not key:
        return finish('Review', 'Source industry is blank or its column could not be identified.')
    if not campaign:
        return finish('Review', 'Campaign industry selection is missing or invalid; no assignment made.')
    # Preserve exact already-compatible source labels without remapping them.
    if key in campaign:
        result['Industry_Map_Selected_Candidates'] = campaign[key]
        return finish('Mapped', 'Exact label match after formatting normalization.', campaign[key])
    records = rule_index.get(key, [])
    result['Industry_Map_Candidates'] = '; '.join(sorted({r['target'] for r in records}, key=industry_key))
    if not records:
        target, method, suggestions = match_campaign_industry(source, campaign)
        if target:
            return finish('Mapped', 'Source matched by ' + method + '.', target)
        result['Industry_Map_Selected_Candidates'] = '; '.join(suggestions)
        return finish('Review', 'Unknown source industry; ' + method + '. Suggestions require review.')
    approved = [r for r in records if r['approved']]
    if not approved:
        return finish('Review', 'Mapping rules are awaiting approval.')
    eligible = []
    uncertain = []
    for r in approved:
        target, method, suggestions = match_campaign_industry(r['target'], campaign)
        if target:
            eligible.append(dict(r, client_target=target, method=method))
        elif suggestions:
            uncertain.extend(suggestions)
    result['Industry_Map_Selected_Candidates'] = '; '.join(sorted(
        {r['client_target'] for r in eligible} | set(uncertain), key=industry_key))
    if not eligible:
        if uncertain:
            return finish('Review', 'Possible wording matches found, but no confident mapping; review the suggestions.')
        return finish('No Match', 'No built-in mapping matches this campaign selection; retain the lead for review.')
    if any(r['context'] for r in eligible):
        return finish('Review', 'Company context is required before assigning one of the selected candidates.')
    if uncertain:
        return finish('Review', 'A candidate has uncertain or ambiguous wording; review before applying priority.')
    eligible.sort(key=lambda r: (r['priority'], industry_key(r['client_target']), industry_key(r['target'])))
    chosen = eligible[0]
    ties = len({r['client_target'] for r in eligible if r['priority'] == chosen['priority']})
    reason = ('Only valid match selected for this campaign.' if len(eligible) == 1 else
              'Multiple valid selected matches; used lowest priority number.')
    if ties > 1:
        reason += ' Equal priorities resolved by fixed alphabetical label order.'
    reason += ' Campaign label matched by ' + chosen['method'] + '.'
    return finish('Mapped', reason, chosen['client_target'], chosen['priority'])


def append_industry_mapping(qa_bytes, source_column, rule_frame, campaign_values, apply_colours=True, allow_missing_setup=False):
    import hashlib
    compiled = compile_industry_rules(rule_frame)
    setup_reason = ''
    try:
        campaign = compile_campaign_industries(campaign_values or [])
    except ValueError as exc:
        if not allow_missing_setup:
            raise
        campaign = {}
        setup_reason = str(exc)
    canonical_rules = sorted(
        (key, industry_key(r['target']), r['priority'], r['approved'], r['context'])
        for key, records in compiled.items() for r in records)
    version = hashlib.sha256(json.dumps(canonical_rules).encode()).hexdigest()[:12]
    book = load_workbook(io.BytesIO(qa_bytes))
    sheet = book['Results']
    headers = [cell.value for cell in sheet[1]]
    if headers.count(source_column) != 1 and not allow_missing_setup:
        raise ValueError('Select one unique source sub-industry column.')
    if any(name in headers for name in INDUSTRY_MAPPING_COLUMNS):
        raise ValueError('This file already contains industry mapping results. Use the original master file.')
    source_position = headers.index(source_column) + 1 if headers.count(source_column) == 1 else None
    qa_status_position = headers.index('QA_Overall_Status') + 1
    first = sheet.max_column + 1
    for j, name in enumerate(INDUSTRY_MAPPING_COLUMNS, first):
        sheet.cell(1, j, name)
        if apply_colours:
            sheet.cell(1, j).fill = HEADER_YELLOW
    # Cache classifications, not contact rows. Key includes source; conflicting source
    # classifications are never merged merely because contacts share a domain.
    cache = {}
    counts = {}
    for row_number in range(2, sheet.max_row + 1):
        # Respect legacy blank/template-row skipping without touching those rows.
        if sheet.cell(row_number, qa_status_position).value is None:
            continue
        source = sheet.cell(row_number, source_position).value if source_position else ''
        cache_key = str(source or '').strip()
        if cache_key not in cache:
            cache[cache_key] = resolve_industry(source, compiled, campaign, version)
        decision = dict(cache[cache_key])
        if setup_reason:
            decision['Industry_Map_Reason'] = 'Campaign setup: ' + setup_reason
        for j, name in enumerate(INDUSTRY_MAPPING_COLUMNS, first):
            cell = sheet.cell(row_number, j, decision[name])
            if apply_colours and name == 'Industry_Map_Status':
                cell.fill = {'Mapped': CELL_GREEN, 'Review': CELL_AMBER, 'No Match': CELL_RED}[decision[name]]
        status = decision['Industry_Map_Status']
        counts[status] = counts.get(status, 0) + 1
    output = io.BytesIO()
    book.save(output)
    return output.getvalue(), {'status_counts': counts, 'rules_version': version,
                               'unique_source_values_resolved': len(cache),
                               'campaign_industries': list(campaign.values()), 'setup_issue': setup_reason}


def industry_grid_bytes(frame):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        frame.to_excel(writer, sheet_name='Industry Rules', index=False)
        sheet = writer.book['Industry Rules']
        sheet.freeze_panes = 'C2'
        sheet.auto_filter.ref = sheet.dimensions
        widths = {'A': 35, 'B': 65, 'C': 48, 'D': 14, 'E': 14, 'F': 20, 'G': 85}
        for column, width in widths.items():
            sheet.column_dimensions[column].width = width
    return output.getvalue()

BUNDLED_INDUSTRY_RULES = json.loads('[{"Main industry": "Aerospace & Defense", "Sub-industry": "Armed Forces", "Leadscale industry": "Military", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Aviation and Aerospace Component Manufacturing", "Leadscale industry": "Aviation & Aerospace", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Defense and Space Manufacturing", "Leadscale industry": "Defense & Space", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Military and International Affairs", "Leadscale industry": "Military", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined military and international affairs; use International Affairs for diplomatic bodies."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Military and International Affairs", "Leadscale industry": "International Affairs", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined military and international affairs; use International Affairs for diplomatic bodies."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Space Research and Technology", "Leadscale industry": "Aviation & Aerospace", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Space research may instead belong to Research or Defense & Space depending on activity."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Space Research and Technology", "Leadscale industry": "Research", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Space research may instead belong to Research or Defense & Space depending on activity."}, {"Main industry": "Aerospace & Defense", "Sub-industry": "Space Research and Technology", "Leadscale industry": "Defense & Space", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Space research may instead belong to Research or Defense & Space depending on activity."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Farming", "Leadscale industry": "Farming", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Farming, Ranching, Forestry", "Leadscale industry": "Agriculture", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined farming, ranching and forestry; Agriculture is the broad fallback. Additional candidate: Paper & Forest Products."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Farming, Ranching, Forestry", "Leadscale industry": "Farming", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined farming, ranching and forestry; Agriculture is the broad fallback. Additional candidate: Paper & Forest Products."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Farming, Ranching, Forestry", "Leadscale industry": "Ranching", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined farming, ranching and forestry; Agriculture is the broad fallback. Additional candidate: Paper & Forest Products."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Farming, Ranching, Forestry", "Leadscale industry": "Paper & Forest Products", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined farming, ranching and forestry; Agriculture is the broad fallback. Additional candidate: Paper & Forest Products."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Fisheries", "Leadscale industry": "Fishery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Forestry and Logging", "Leadscale industry": "Paper & Forest Products", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Horticulture", "Leadscale industry": "Farming", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Metal Ore Mining", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Mining", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Nonmetallic Mineral Mining", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Ranching", "Leadscale industry": "Ranching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Ranching and Fisheries", "Leadscale industry": "Ranching", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined ranching and fisheries; verify the dominant activity."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Ranching and Fisheries", "Leadscale industry": "Fishery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined ranching and fisheries; verify the dominant activity."}, {"Main industry": "Agriculture & Natural Resources", "Sub-industry": "Wholesale Raw Farm Products", "Leadscale industry": "Agriculture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Automotive", "Sub-industry": "Alternative Fuel Vehicle Manufacturing", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Automotive", "Sub-industry": "Motor Vehicle Manufacturing", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Automotive", "Sub-industry": "Motor Vehicle Parts Manufacturing", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Automotive", "Sub-industry": "Vehicle Repair and Maintenance", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Automotive", "Sub-industry": "Wholesale Motor Vehicles and Parts", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Accounting", "Leadscale industry": "Accounting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Administrative and Support Services", "Leadscale industry": "Facilities Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad support-services category; Facilities Services is a provisional fallback."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Administrative and Support Services", "Leadscale industry": "Outsourcing/Offshoring", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad support-services category; Facilities Services is a provisional fallback."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Business Consulting and Services", "Leadscale industry": "Management Consulting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Child Day Care Services", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Commercial and Industrial Equipment Rental", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Machinery rental assumed; category depends on equipment."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Commercial and Industrial Equipment Rental", "Leadscale industry": "Business Supplies & Equipment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Machinery rental assumed; category depends on equipment."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Consumer Goods Rental", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Consumer rental is a service; verify asset type."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Consumer Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Death Care Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Design Services", "Leadscale industry": "Design", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Engineering Services", "Leadscale industry": "Mechanical or Industrial Engineering", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "General engineering spans civil, software and industrial disciplines."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Engineering Services", "Leadscale industry": "Civil Engineering", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "General engineering spans civil, software and industrial disciplines."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Environmental Services", "Leadscale industry": "Environmental Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Equipment Rental Services", "Leadscale industry": "Business Supplies & Equipment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Equipment type unspecified; Business Supplies & Equipment is provisional."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Equipment Rental Services", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Equipment type unspecified; Business Supplies & Equipment is provisional."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Events Services", "Leadscale industry": "Events Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Executive Search Services", "Leadscale industry": "Staffing & Recruiting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Facilities Services", "Leadscale industry": "Facilities Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Footwear and Leather Goods Repair", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Funeral Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Household Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Human Resources Services", "Leadscale industry": "Human Resources", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Information Services", "Leadscale industry": "Information Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Janitorial Services", "Leadscale industry": "Facilities Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Landscaping Services", "Leadscale industry": "Facilities Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Laundry and Drycleaning Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Office Administration", "Leadscale industry": "Executive Office", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Operations Consulting", "Leadscale industry": "Management Consulting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Outsourcing and Offshoring Consulting", "Leadscale industry": "Outsourcing/Offshoring", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Personal and Laundry Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Personal Care Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Pet Services", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Non-clinical pet services; veterinary care belongs to Veterinary."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Photography", "Leadscale industry": "Photography", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Printing Services", "Leadscale industry": "Printing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Professional Services", "Leadscale industry": "Management Consulting", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No general Professional Services target; Management Consulting is a provisional fallback. Additional candidate: Design."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Professional Services", "Leadscale industry": "Legal Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No general Professional Services target; Management Consulting is a provisional fallback. Additional candidate: Design."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Professional Services", "Leadscale industry": "Accounting", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No general Professional Services target; Management Consulting is a provisional fallback. Additional candidate: Design."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Professional Services", "Leadscale industry": "Design", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No general Professional Services target; Management Consulting is a provisional fallback. Additional candidate: Design."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Repair and Maintenance", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Repair type unspecified; specialist equipment may need a different category."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Research Services", "Leadscale industry": "Research", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Reupholstery and Furniture Repair", "Leadscale industry": "Consumer Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Security and Investigations", "Leadscale industry": "Security & Investigations", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Security Guards and Patrol Services", "Leadscale industry": "Security & Investigations", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Security Systems Services", "Leadscale industry": "Security & Investigations", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Staffing and Recruiting", "Leadscale industry": "Staffing & Recruiting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Strategic Management Services", "Leadscale industry": "Management Consulting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Telephone Call Centers", "Leadscale industry": "Outsourcing/Offshoring", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Temporary Help Services", "Leadscale industry": "Staffing & Recruiting", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Translation and Localization", "Leadscale industry": "Translation & Localization", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Wellness and Fitness Services", "Leadscale industry": "Health, Wellness & Fitness", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Wholesale", "Leadscale industry": "Wholesale", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Business and Consumer Services", "Sub-industry": "Wholesale Import and Export", "Leadscale industry": "Import & Export", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Hardware", "Sub-industry": "Accessible Hardware Manufacturing", "Leadscale industry": "Computer Hardware", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Hardware type unspecified; computer hardware assumed."}, {"Main industry": "Computer Hardware", "Sub-industry": "Accessible Hardware Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Hardware type unspecified; computer hardware assumed."}, {"Main industry": "Computer Hardware", "Sub-industry": "Computer Hardware Manufacturing", "Leadscale industry": "Computer Hardware", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Hardware", "Sub-industry": "Computer Networking Products", "Leadscale industry": "Computer Networking", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Hardware", "Sub-industry": "Computers and Electronics Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad computers/electronics manufacturing; choose Computer Hardware for computer-only firms."}, {"Main industry": "Computer Hardware", "Sub-industry": "Computers and Electronics Manufacturing", "Leadscale industry": "Computer Hardware", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad computers/electronics manufacturing; choose Computer Hardware for computer-only firms."}, {"Main industry": "Computer Hardware", "Sub-industry": "Computers and Electronics Manufacturing", "Leadscale industry": "Consumer Electronics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad computers/electronics manufacturing; choose Computer Hardware for computer-only firms."}, {"Main industry": "Computer Hardware", "Sub-industry": "Semiconductor Manufacturing", "Leadscale industry": "Semiconductors", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Hardware", "Sub-industry": "Wholesale Computer Equipment", "Leadscale industry": "Computer Hardware", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "Digital Accessibility Services", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Digital accessibility can include consulting, software and design."}, {"Main industry": "Computer Services", "Sub-industry": "Digital Accessibility Services", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Digital accessibility can include consulting, software and design."}, {"Main industry": "Computer Services", "Sub-industry": "Digital Accessibility Services", "Leadscale industry": "Design", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Digital accessibility can include consulting, software and design."}, {"Main industry": "Computer Services", "Sub-industry": "IT Services and IT Consulting", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Custom Software Development", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Closest available Leadscale category based on the industry activity."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Custom Software Development", "Leadscale industry": "Information Technology & Services", "Priority": 2, "Approved": "Yes", "Context required": "No", "Notes": "Closest available Leadscale category based on the industry activity."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Data Services", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Design Services", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Installation and Disposal", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Operations and Maintenance", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Testing and Evaluation", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "IT System Training and Support", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Internet", "Leadscale industry": "Internet", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad technology/internet category; verify primary activity."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Internet", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad technology/internet category; verify primary activity."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Media", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combines technology, information and media; no single equivalent. Additional candidate: Media Production."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Media", "Leadscale industry": "Internet", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combines technology, information and media; no single equivalent. Additional candidate: Media Production."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Media", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combines technology, information and media; no single equivalent. Additional candidate: Media Production."}, {"Main industry": "Computer Services", "Sub-industry": "Technology, Information and Media", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combines technology, information and media; no single equivalent. Additional candidate: Media Production."}, {"Main industry": "Computer Software", "Sub-industry": "Blockchain Services", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Blockchain services may be software, finance or IT consulting."}, {"Main industry": "Computer Software", "Sub-industry": "Blockchain Services", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Blockchain services may be software, finance or IT consulting."}, {"Main industry": "Computer Software", "Sub-industry": "Blockchain Services", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Blockchain services may be software, finance or IT consulting."}, {"Main industry": "Computer Software", "Sub-industry": "Business Intelligence Platforms", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Computer and Network Security", "Leadscale industry": "Computer & Network Security", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Computer Games", "Leadscale industry": "Computer Games", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Data Infrastructure and Analytics", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Closest available Leadscale category based on the industry activity."}, {"Main industry": "Computer Software", "Sub-industry": "Data Infrastructure and Analytics", "Leadscale industry": "Computer Software", "Priority": 2, "Approved": "Yes", "Context required": "No", "Notes": "Closest available Leadscale category based on the industry activity."}, {"Main industry": "Computer Software", "Sub-industry": "Data Security Software Products", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Desktop Computing Software Products", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Embedded Software Products", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Internet Marketplace Platforms", "Leadscale industry": "Internet", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Mobile Computing Software Products", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Mobile Gaming Apps", "Leadscale industry": "Computer Games", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Social Networking Platforms", "Leadscale industry": "Internet", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Computer Software", "Sub-industry": "Software Development", "Leadscale industry": "Computer Software", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Apparel Manufacturing", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Baked Goods Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Beverage Manufacturing", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Breweries", "Leadscale industry": "Wine & Spirits", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Dairy Product Manufacturing", "Leadscale industry": "Dairy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Distilleries", "Leadscale industry": "Wine & Spirits", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Fashion Accessories Manufacturing", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Food and Beverage Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Footwear Manufacturing", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Fruit and Vegetable Preserves Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Furniture and Home Furnishings Manufacturing", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Household and Institutional Furniture Manufacturing", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Household Appliance Manufacturing", "Leadscale industry": "Consumer Electronics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Consumer electrical appliances assumed; industrial appliances differ."}, {"Main industry": "Consumer Products", "Sub-industry": "Household Appliance Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Consumer electrical appliances assumed; industrial appliances differ."}, {"Main industry": "Consumer Products", "Sub-industry": "Leather Product Manufacturing", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Mattress and Blinds Manufacturing", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined mattresses and blinds; Furniture is the nearest broad category."}, {"Main industry": "Consumer Products", "Sub-industry": "Meat Products Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Office Furniture and Fixtures Manufacturing", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Personal Care Product Manufacturing", "Leadscale industry": "Cosmetics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Includes personal care beyond cosmetics; check product type."}, {"Main industry": "Consumer Products", "Sub-industry": "Personal Care Product Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Includes personal care beyond cosmetics; check product type."}, {"Main industry": "Consumer Products", "Sub-industry": "Seafood Product Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Soap and Cleaning Product Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No cleaning-products category; Consumer Goods is a fallback."}, {"Main industry": "Consumer Products", "Sub-industry": "Soap and Cleaning Product Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No cleaning-products category; Consumer Goods is a fallback."}, {"Main industry": "Consumer Products", "Sub-industry": "Sporting Goods Manufacturing", "Leadscale industry": "Sporting Goods", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Sugar and Confectionery Product Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Textile Manufacturing", "Leadscale industry": "Textiles", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Tobacco Manufacturing", "Leadscale industry": "Tobacco", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Alcoholic Beverages", "Leadscale industry": "Wine & Spirits", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Apparel and Sewing Supplies", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Appliances, Electrical, and Electronics", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Food and Beverage", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Footwear", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Furniture and Home Furnishings", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Luxury Goods and Jewelry", "Leadscale industry": "Luxury Goods & Jewelry", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wholesale Photography Equipment and Supplies", "Leadscale industry": "Photography", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Wineries", "Leadscale industry": "Wine & Spirits", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Consumer Products", "Sub-industry": "Women\'s Handbag Manufacturing", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Cosmetology and Barber Schools", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "E-Learning Providers", "Leadscale industry": "E-learning", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Education", "Leadscale industry": "Education Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad education category; verify level and delivery model. Additional candidate: Professional Training & Coaching."}, {"Main industry": "Education", "Sub-industry": "Education", "Leadscale industry": "Higher Education", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad education category; verify level and delivery model. Additional candidate: Professional Training & Coaching."}, {"Main industry": "Education", "Sub-industry": "Education", "Leadscale industry": "Primary/Secondary Education", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad education category; verify level and delivery model. Additional candidate: Professional Training & Coaching."}, {"Main industry": "Education", "Sub-industry": "Education", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad education category; verify level and delivery model. Additional candidate: Professional Training & Coaching."}, {"Main industry": "Education", "Sub-industry": "Education Administration Programs", "Leadscale industry": "Education Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Fine Arts Schools", "Leadscale industry": "Education Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "A school is mapped by education activity rather than art production."}, {"Main industry": "Education", "Sub-industry": "Fine Arts Schools", "Leadscale industry": "Fine Art", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "A school is mapped by education activity rather than art production."}, {"Main industry": "Education", "Sub-industry": "Flight Training", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Higher Education", "Leadscale industry": "Higher Education", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Language Schools", "Leadscale industry": "Education Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Primary and Secondary Education", "Leadscale industry": "Primary/Secondary Education", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Professional Training and Coaching", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Secretarial Schools", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Education", "Sub-industry": "Technical and Vocational Training", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Biomass Electric Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Data and Analytics", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Climate analytics may fit Research or IT Services instead."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Data and Analytics", "Leadscale industry": "Research", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Climate analytics may fit Research or IT Services instead."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Data and Analytics", "Leadscale industry": "Information Technology & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Climate analytics may fit Research or IT Services instead."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Technology Product Manufacturing", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Product type unspecified; climate focus used."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Technology Product Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Product type unspecified; climate focus used."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Climate Technology Product Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Product type unspecified; climate focus used."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Coal Mining", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Electric Power Generation", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Electric Power Transmission, Control, and Distribution", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Energy Technology", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Energy technology spans renewable and fossil energy; review before use."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Energy Technology", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Energy technology spans renewable and fossil energy; review before use."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Fossil Fuel Electric Power Generation", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Fuel Cell Manufacturing", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Fuel cells may use non-renewable fuels; verify technology."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Fuel Cell Manufacturing", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Fuel cells may use non-renewable fuels; verify technology."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Geothermal Electric Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Hydroelectric Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Natural Gas Distribution", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Natural Gas Extraction", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Nuclear Electric Power Generation", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Oil and Coal Product Manufacturing", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Oil and Gas", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Oil Extraction", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Oil, Gas, and Mining", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined oil, gas and mining; Mining & Metals may be preferable."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Oil, Gas, and Mining", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined oil, gas and mining; Mining & Metals may be preferable."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Pipeline Transportation", "Leadscale industry": "Logistics & Supply Chain", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Renewable Energy Equipment Manufacturing", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Renewable Energy Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Renewable Energy Semiconductor Manufacturing", "Leadscale industry": "Semiconductors", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Mapped by semiconductor manufacturing activity rather than end-use."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Renewable Energy Semiconductor Manufacturing", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Mapped by semiconductor manufacturing activity rather than end-use."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Services for Renewable Energy", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Smart Meter Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Solar Electric Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Steam and Air-Conditioning Supply", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Utilities", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Waste Collection", "Leadscale industry": "Environmental Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Waste Treatment and Disposal", "Leadscale industry": "Environmental Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Water Supply and Irrigation Systems", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Water, Waste, Steam, and Air Conditioning Services", "Leadscale industry": "Utilities", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Wholesale Petroleum and Petroleum Products", "Leadscale industry": "Oil & Energy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Energy & Utilities", "Sub-industry": "Wind Electric Power Generation", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Banking", "Leadscale industry": "Banking", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Capital Markets", "Leadscale industry": "Capital Markets", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Collection Agencies", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Debt collection assumed; some agencies fit Legal Services."}, {"Main industry": "Financial Services", "Sub-industry": "Collection Agencies", "Leadscale industry": "Legal Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Debt collection assumed; some agencies fit Legal Services."}, {"Main industry": "Financial Services", "Sub-industry": "Credit Intermediation", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Financial Services", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Funds and Trusts", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Holding Companies", "Leadscale industry": "Management Consulting", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Holding companies are not necessarily investment managers; no direct equivalent."}, {"Main industry": "Financial Services", "Sub-industry": "Holding Companies", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Holding companies are not necessarily investment managers; no direct equivalent."}, {"Main industry": "Financial Services", "Sub-industry": "Holding Companies", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Holding companies are not necessarily investment managers; no direct equivalent."}, {"Main industry": "Financial Services", "Sub-industry": "Investment Advice", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Investment Banking", "Leadscale industry": "Investment Banking", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Investment Management", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Loan Brokers", "Leadscale industry": "Financial Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Pension Funds", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Savings Institutions", "Leadscale industry": "Banking", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Securities and Commodity Exchanges", "Leadscale industry": "Capital Markets", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Financial Services", "Sub-industry": "Venture Capital and Private Equity Principals", "Leadscale industry": "Venture Capital & Private Equity", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Administration of Justice", "Leadscale industry": "Judiciary", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad justice administration; Government Administration may apply."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Administration of Justice", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad justice administration; Government Administration may apply."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Air, Water, and Waste Program Management", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Community Development and Urban Planning", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Conservation Programs", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Government programme administration assumed; verify operator type."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Conservation Programs", "Leadscale industry": "Environmental Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Government programme administration assumed; verify operator type."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Conservation Programs", "Leadscale industry": "Renewables & Environment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Government programme administration assumed; verify operator type."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Correctional Institutions", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Courts of Law", "Leadscale industry": "Judiciary", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Economic Programs", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Environmental Quality Programs", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Executive Offices", "Leadscale industry": "Executive Office", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Fire Protection", "Leadscale industry": "Public Safety", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Government Administration", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Government Relations Services", "Leadscale industry": "Government Relations", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Health and Human Services", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Public administration assumed; healthcare providers differ."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Health and Human Services", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Public administration assumed; healthcare providers differ."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Health and Human Services", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Public administration assumed; healthcare providers differ."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Housing and Community Development", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Housing Programs", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "International Affairs", "Leadscale industry": "International Affairs", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "International Trade and Development", "Leadscale industry": "International Trade & Development", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Law Enforcement", "Leadscale industry": "Law Enforcement", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Legislative Offices", "Leadscale industry": "Legislative Office", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Libraries", "Leadscale industry": "Libraries", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Assistance Programs", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Health", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Public health services assumed; government agencies may fit Government Administration."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Health", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Public health services assumed; government agencies may fit Government Administration."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Policy Offices", "Leadscale industry": "Public Policy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Safety", "Leadscale industry": "Public Safety", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Public Works", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Transportation Programs", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Government & Public Sector", "Sub-industry": "Utilities Administration", "Leadscale industry": "Government Administration", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Alternative Medicine", "Leadscale industry": "Alternative Medicine", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Ambulance Services", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Chiropractors", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Dentists", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Family Planning Centers", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Home Health Care Services", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Hospitals", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Hospitals and Health Care", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Medical and Diagnostic Laboratories", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Medical Practices", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Mental Health Care", "Leadscale industry": "Mental Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Nursing Homes and Residential Care Facilities", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Optometrists", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Outpatient Care Centers", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Physical, Occupational and Speech Therapists", "Leadscale industry": "Hospital & Health Care", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Physicians", "Leadscale industry": "Medical Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Hospitals & Healthcare", "Sub-industry": "Veterinary Services", "Leadscale industry": "Veterinary", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Insurance", "Sub-industry": "Claims Adjusting, Actuarial Services", "Leadscale industry": "Insurance", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Insurance", "Sub-industry": "Insurance", "Leadscale industry": "Insurance", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Insurance", "Sub-industry": "Insurance Agencies and Brokerages", "Leadscale industry": "Insurance", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Insurance", "Sub-industry": "Insurance and Employee Benefit Funds", "Leadscale industry": "Insurance", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Insurance", "Sub-industry": "Insurance Carriers", "Leadscale industry": "Insurance", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Legal Services", "Sub-industry": "Alternative Dispute Resolution", "Leadscale industry": "Alternative Dispute Resolution", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Legal Services", "Sub-industry": "Law Practice", "Leadscale industry": "Law Practice", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Legal Services", "Sub-industry": "Legal Services", "Leadscale industry": "Legal Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Legal Services", "Sub-industry": "Trusts and Estates", "Leadscale industry": "Legal Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Legal trusts/estates services assumed; investment trusts fit Investment Management."}, {"Main industry": "Legal Services", "Sub-industry": "Trusts and Estates", "Leadscale industry": "Investment Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Legal trusts/estates services assumed; investment trusts fit Investment Management."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Abrasives and Nonmetallic Minerals Manufacturing", "Leadscale industry": "Glass, Ceramics & Concrete", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Agricultural Chemical Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Agriculture, Construction, Mining Machinery Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Animal Feed Manufacturing", "Leadscale industry": "Food Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Appliances, Electrical, and Electronics Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Architectural and Structural Metal Manufacturing", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Artificial Rubber and Synthetic Fiber Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Synthetic rubber/fibre material production; verify Plastics or Textiles if appropriate."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Artificial Rubber and Synthetic Fiber Manufacturing", "Leadscale industry": "Plastics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Synthetic rubber/fibre material production; verify Plastics or Textiles if appropriate."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Artificial Rubber and Synthetic Fiber Manufacturing", "Leadscale industry": "Textiles", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Synthetic rubber/fibre material production; verify Plastics or Textiles if appropriate."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Audio and Video Equipment Manufacturing", "Leadscale industry": "Consumer Electronics", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Automation Machinery Manufacturing", "Leadscale industry": "Industrial Automation", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Boilers, Tanks, and Shipping Container Manufacturing", "Leadscale industry": "Packaging & Containers", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined boilers, tanks and containers; Machinery may be more suitable."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Boilers, Tanks, and Shipping Container Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined boilers, tanks and containers; Machinery may be more suitable."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Chemical Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Chemical Raw Materials Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Clay and Refractory Products Manufacturing", "Leadscale industry": "Glass, Ceramics & Concrete", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Commercial and Industrial Machinery Maintenance", "Leadscale industry": "Mechanical or Industrial Engineering", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Industrial maintenance rather than engineering; no direct maintenance target."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Commercial and Industrial Machinery Maintenance", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Industrial maintenance rather than engineering; no direct maintenance target."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Commercial and Service Industry Machinery Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Construction Hardware Manufacturing", "Leadscale industry": "Building Materials", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Cutlery and Handtool Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Consumer hand tools assumed; industrial tool makers may fit Machinery."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Cutlery and Handtool Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Consumer hand tools assumed; industrial tool makers may fit Machinery."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Electric Lighting Equipment Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Electrical Equipment Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Electronic and Precision Equipment Maintenance", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Engines and Power Transmission Equipment Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Fabricated Metal Products", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Glass Product Manufacturing", "Leadscale industry": "Glass, Ceramics & Concrete", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Glass, Ceramics and Concrete Manufacturing", "Leadscale industry": "Glass, Ceramics & Concrete", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "HVAC and Refrigeration Equipment Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Industrial Machinery Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Lime and Gypsum Products Manufacturing", "Leadscale industry": "Glass, Ceramics & Concrete", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Machinery Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Magnetic and Optical Media Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Manufacture of storage media; no direct target."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Magnetic and Optical Media Manufacturing", "Leadscale industry": "Computer Hardware", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Manufacture of storage media; no direct target."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Manufacturing", "Leadscale industry": "Mechanical or Industrial Engineering", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Manufacturing target; industrial engineering is a provisional fallback. Additional candidate: Electrical/Electronic Manufacturing."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Manufacturing target; industrial engineering is a provisional fallback. Additional candidate: Electrical/Electronic Manufacturing."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Manufacturing target; industrial engineering is a provisional fallback. Additional candidate: Electrical/Electronic Manufacturing."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Manufacturing target; industrial engineering is a provisional fallback. Additional candidate: Electrical/Electronic Manufacturing."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Measuring and Control Instrument Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Metal Treatments", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Metal Valve, Ball, and Roller Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Metalworking Machinery Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Packaging and Containers Manufacturing", "Leadscale industry": "Packaging & Containers", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Paint, Coating, and Adhesive Manufacturing", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Paper and Forest Product Manufacturing", "Leadscale industry": "Paper & Forest Products", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Plastics and Rubber Product Manufacturing", "Leadscale industry": "Plastics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined plastics and rubber; Plastics is the nearest available category."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Plastics and Rubber Product Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined plastics and rubber; Plastics is the nearest available category."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Plastics Manufacturing", "Leadscale industry": "Plastics", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Primary Metal Manufacturing", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Railroad Equipment Manufacturing", "Leadscale industry": "Railroad Manufacture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Robot Manufacturing", "Leadscale industry": "Industrial Automation", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Robotics Engineering", "Leadscale industry": "Mechanical or Industrial Engineering", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Rubber Products Manufacturing", "Leadscale industry": "Consumer Goods", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Rubber target; Consumer Goods is a provisional fallback dependent on product."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Rubber Products Manufacturing", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Rubber target; Consumer Goods is a provisional fallback dependent on product."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Rubber Products Manufacturing", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No Rubber target; Consumer Goods is a provisional fallback dependent on product."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Shipbuilding", "Leadscale industry": "Shipbuilding", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Spring and Wire Product Manufacturing", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Transportation Equipment Manufacturing", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Transport equipment is broader than Automotive; verify vehicle type. Additional candidate: Aviation & Aerospace."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Transportation Equipment Manufacturing", "Leadscale industry": "Railroad Manufacture", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Transport equipment is broader than Automotive; verify vehicle type. Additional candidate: Aviation & Aerospace."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Transportation Equipment Manufacturing", "Leadscale industry": "Shipbuilding", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Transport equipment is broader than Automotive; verify vehicle type. Additional candidate: Aviation & Aerospace."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Transportation Equipment Manufacturing", "Leadscale industry": "Aviation & Aerospace", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Transport equipment is broader than Automotive; verify vehicle type. Additional candidate: Aviation & Aerospace."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Turned Products and Fastener Manufacturing", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Chemical and Allied Products", "Leadscale industry": "Chemicals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Machinery", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Metals and Minerals", "Leadscale industry": "Mining & Metals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Paper Products", "Leadscale industry": "Paper & Forest Products", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Recyclable Materials", "Leadscale industry": "Environmental Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Recycling activity assumed; pure traders may fit Wholesale."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wholesale Recyclable Materials", "Leadscale industry": "Wholesale", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Recycling activity assumed; pure traders may fit Wholesale."}, {"Main industry": "Manufacturing & Industrial", "Sub-industry": "Wood Product Manufacturing", "Leadscale industry": "Paper & Forest Products", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Advertising Services", "Leadscale industry": "Marketing & Advertising", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Animation and Post-production", "Leadscale industry": "Animation", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Includes post-production beyond animation; Media Production may apply."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Animation and Post-production", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Includes post-production beyond animation; Media Production may apply."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Artists and Writers", "Leadscale industry": "Fine Art", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined artists and writers; writers should map to Writing & Editing."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Artists and Writers", "Leadscale industry": "Writing & Editing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined artists and writers; writers should map to Writing & Editing."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Blogs", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Book and Periodical Publishing", "Leadscale industry": "Publishing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Book Publishing", "Leadscale industry": "Publishing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Broadcast Media Production and Distribution", "Leadscale industry": "Broadcast Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Business Content", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Format unspecified; may fit Publishing or Marketing & Advertising."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Business Content", "Leadscale industry": "Publishing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Format unspecified; may fit Publishing or Marketing & Advertising."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Business Content", "Leadscale industry": "Marketing & Advertising", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Format unspecified; may fit Publishing or Marketing & Advertising."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Graphic Design", "Leadscale industry": "Graphic Design", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Internet News", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Internet Publishing", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Market Research", "Leadscale industry": "Market Research", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Marketing Services", "Leadscale industry": "Marketing & Advertising", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Media Production", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies and Sound Recording", "Leadscale industry": "Motion Pictures & Film", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film and audio; music-only companies should map to Music."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies and Sound Recording", "Leadscale industry": "Music", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film and audio; music-only companies should map to Music."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies and Sound Recording", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film and audio; music-only companies should map to Music."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies, Videos and Sound", "Leadscale industry": "Motion Pictures & Film", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film, video and sound; Music may apply to audio-only firms."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies, Videos and Sound", "Leadscale industry": "Music", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film, video and sound; Music may apply to audio-only firms."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Movies, Videos and Sound", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined film, video and sound; Music may apply to audio-only firms."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Musicians", "Leadscale industry": "Music", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Newspaper Publishing", "Leadscale industry": "Newspapers", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Online Audio and Video Media", "Leadscale industry": "Online Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Periodical Publishing", "Leadscale industry": "Publishing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Public Relations and Communications Services", "Leadscale industry": "Public Relations & Communications", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Radio and Television Broadcasting", "Leadscale industry": "Broadcast Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Sheet Music Publishing", "Leadscale industry": "Music", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Sound Recording", "Leadscale industry": "Music", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Marketing, Media & Advertising", "Sub-industry": "Writing and Editing", "Leadscale industry": "Writing & Editing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Civic and Social Organizations", "Leadscale industry": "Civic & Social Organization", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Community Services", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Emergency and Relief Services", "Leadscale industry": "Nonprofit Organization Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Fundraising", "Leadscale industry": "Fund-Raising", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Individual and Family Services", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Industry Associations", "Leadscale industry": "Civic & Social Organization", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Trade association assumed; verify organisation purpose."}, {"Main industry": "Nonprofit", "Sub-industry": "Industry Associations", "Leadscale industry": "Nonprofit Organization Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Trade association assumed; verify organisation purpose."}, {"Main industry": "Nonprofit", "Sub-industry": "Non-profit Organizations", "Leadscale industry": "Nonprofit Organization Management", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Philanthropic Fundraising Services", "Leadscale industry": "Philanthropy", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Political Organizations", "Leadscale industry": "Political Organization", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Professional Organizations", "Leadscale industry": "Civic & Social Organization", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Professional membership body assumed; verify organisation purpose."}, {"Main industry": "Nonprofit", "Sub-industry": "Professional Organizations", "Leadscale industry": "Nonprofit Organization Management", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Professional membership body assumed; verify organisation purpose."}, {"Main industry": "Nonprofit", "Sub-industry": "Religious Institutions", "Leadscale industry": "Religious Institutions", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Services for the Elderly and Disabled", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Think Tanks", "Leadscale industry": "Think Tanks", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Nonprofit", "Sub-industry": "Vocational Rehabilitation Services", "Leadscale industry": "Individual & Family Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Biotechnology Research", "Leadscale industry": "Biotechnology", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Medical Device", "Leadscale industry": "Medical Devices", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Medical Equipment Manufacturing", "Leadscale industry": "Medical Devices", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Nanotechnology Research", "Leadscale industry": "Nanotechnology", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Pharmaceutical Manufacturing", "Leadscale industry": "Pharmaceuticals", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Wholesale Drugs and Sundries", "Leadscale industry": "Pharmaceuticals", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Wholesale pharmaceuticals assumed; sundries-only distributors may fit Wholesale."}, {"Main industry": "Pharma & Life Sciences", "Sub-industry": "Wholesale Drugs and Sundries", "Leadscale industry": "Wholesale", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Wholesale pharmaceuticals assumed; sundries-only distributors may fit Wholesale."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Accessible Architecture and Design", "Leadscale industry": "Architecture & Planning", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Architecture and Planning", "Leadscale industry": "Architecture & Planning", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Building Construction", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Building Equipment Contractors", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Building Finishing Contractors", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Building Structure and Exterior Contractors", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Civil Engineering", "Leadscale industry": "Civil Engineering", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Construction", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Highway, Street, and Bridge Construction", "Leadscale industry": "Civil Engineering", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Interior Design", "Leadscale industry": "Design", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Leasing Non-residential Real Estate", "Leadscale industry": "Commercial Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Leasing Residential Real Estate", "Leadscale industry": "Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Nonresidential Building Construction", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Real Estate", "Leadscale industry": "Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Real Estate Agents and Brokers", "Leadscale industry": "Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Real Estate and Equipment Rental Services", "Leadscale industry": "Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Real estate and equipment rental combined; verify principal activity."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Real Estate and Equipment Rental Services", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Real estate and equipment rental combined; verify principal activity."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Real Estate and Equipment Rental Services", "Leadscale industry": "Business Supplies & Equipment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Real estate and equipment rental combined; verify principal activity."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Regenerative Design", "Leadscale industry": "Design", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Regenerative design spans architecture, engineering and product design."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Regenerative Design", "Leadscale industry": "Architecture & Planning", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Regenerative design spans architecture, engineering and product design."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Regenerative Design", "Leadscale industry": "Mechanical or Industrial Engineering", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Regenerative design spans architecture, engineering and product design."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Residential Building Construction", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Specialty Trade Contractors", "Leadscale industry": "Construction", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Subdivision of Land", "Leadscale industry": "Real Estate", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Surveying and Mapping Services", "Leadscale industry": "Civil Engineering", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Utility System Construction", "Leadscale industry": "Civil Engineering", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Wholesale Building Materials", "Leadscale industry": "Building Materials", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Real Estate & Construction", "Sub-industry": "Wholesale Hardware, Plumbing, Heating Equipment", "Leadscale industry": "Building Materials", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Food and Beverage Retail", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Food retail spans supermarkets, specialist shops and beverage retail."}, {"Main industry": "Retail", "Sub-industry": "Food and Beverage Retail", "Leadscale industry": "Supermarkets", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Food retail spans supermarkets, specialist shops and beverage retail."}, {"Main industry": "Retail", "Sub-industry": "Food and Beverage Retail", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Food retail spans supermarkets, specialist shops and beverage retail."}, {"Main industry": "Retail", "Sub-industry": "Online and Mail Order Retail", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Apparel and Fashion", "Leadscale industry": "Apparel & Fashion", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Appliances, Electrical, and Electronic Equipment", "Leadscale industry": "Consumer Electronics", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Art Dealers", "Leadscale industry": "Fine Art", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Art Supplies", "Leadscale industry": "Arts & Crafts", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Books and Printed News", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Building Materials and Garden Equipment", "Leadscale industry": "Building Materials", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Florists", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Furniture and Home Furnishings", "Leadscale industry": "Furniture", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Gasoline", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Groceries", "Leadscale industry": "Supermarkets", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Health and Personal Care Products", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Retail activity retained; pharmacies and manufacturers differ."}, {"Main industry": "Retail", "Sub-industry": "Retail Health and Personal Care Products", "Leadscale industry": "Pharmaceuticals", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Retail activity retained; pharmacies and manufacturers differ."}, {"Main industry": "Retail", "Sub-industry": "Retail Health and Personal Care Products", "Leadscale industry": "Cosmetics", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Retail activity retained; pharmacies and manufacturers differ."}, {"Main industry": "Retail", "Sub-industry": "Retail Luxury Goods and Jewelry", "Leadscale industry": "Luxury Goods & Jewelry", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Motor Vehicles", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Musical Instruments", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Office Equipment", "Leadscale industry": "Business Supplies & Equipment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Office Supplies and Gifts", "Leadscale industry": "Business Supplies & Equipment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Pharmacies", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Retail", "Sub-industry": "Retail Recyclable Materials & Used Merchandise", "Leadscale industry": "Retail", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Cable and Satellite Programming", "Leadscale industry": "Broadcast Media", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Communications Equipment Manufacturing", "Leadscale industry": "Electrical/Electronic Manufacturing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Media & Telecommunications", "Leadscale industry": "Telecommunications", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined media and telecommunications; verify principal activity."}, {"Main industry": "Telecommunications", "Sub-industry": "Media & Telecommunications", "Leadscale industry": "Broadcast Media", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined media and telecommunications; verify principal activity."}, {"Main industry": "Telecommunications", "Sub-industry": "Media & Telecommunications", "Leadscale industry": "Media Production", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined media and telecommunications; verify principal activity."}, {"Main industry": "Telecommunications", "Sub-industry": "Satellite Telecommunications", "Leadscale industry": "Telecommunications", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Telecommunications", "Leadscale industry": "Telecommunications", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Telecommunications Carriers", "Leadscale industry": "Telecommunications", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Telecommunications", "Sub-industry": "Wireless Services", "Leadscale industry": "Wireless", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Freight and Package Transportation", "Leadscale industry": "Package/Freight Delivery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Ground Passenger Transportation", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Interurban and Rural Bus Services", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Maritime Transportation", "Leadscale industry": "Maritime", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Parts Distribution", "Leadscale industry": "Wholesale", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Parts type unspecified; Wholesale is provisional."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Parts Distribution", "Leadscale industry": "Automotive", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Parts type unspecified; Wholesale is provisional."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Parts Distribution", "Leadscale industry": "Machinery", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Parts type unspecified; Wholesale is provisional."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Postal Services", "Leadscale industry": "Package/Freight Delivery", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Rail Transportation", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "School and Employee Bus Services", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Shuttles and Special Needs Transportation Services", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Taxi and Limousine Services", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Transportation, Logistics, Supply Chain and Storage", "Leadscale industry": "Logistics & Supply Chain", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad logistics/storage category; Warehousing or transport may be more precise."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Transportation, Logistics, Supply Chain and Storage", "Leadscale industry": "Warehousing", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad logistics/storage category; Warehousing or transport may be more precise."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Transportation, Logistics, Supply Chain and Storage", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Broad logistics/storage category; Warehousing or transport may be more precise."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Truck Transportation", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Urban Transit Services", "Leadscale industry": "Transportation/Trucking/Railroad", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Transportation & Logistics", "Sub-industry": "Warehousing and Storage", "Leadscale industry": "Warehousing", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Accommodation and Food Services", "Leadscale industry": "Hospitality", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined accommodation and food services; verify principal activity."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Accommodation and Food Services", "Leadscale industry": "Restaurants", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined accommodation and food services; verify principal activity."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Accommodation and Food Services", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined accommodation and food services; verify principal activity."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Airlines and Aviation", "Leadscale industry": "Airlines/Aviation", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Amusement Parks and Arcades", "Leadscale industry": "Entertainment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Bars, Taverns, and Nightclubs", "Leadscale industry": "Entertainment", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Nightclubs may fit Hospitality; Entertainment used for combined category."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Bars, Taverns, and Nightclubs", "Leadscale industry": "Hospitality", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Nightclubs may fit Hospitality; Entertainment used for combined category."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Bed-and-Breakfasts, Hostels, Homestays", "Leadscale industry": "Hospitality", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Caterers", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Circuses and Magic Shows", "Leadscale industry": "Entertainment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Dance Companies", "Leadscale industry": "Performing Arts", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Entertainment Providers", "Leadscale industry": "Entertainment", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Food and Beverage Services", "Leadscale industry": "Food & Beverages", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Gambling Facilities and Casinos", "Leadscale industry": "Gambling & Casinos", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Golf Courses and Country Clubs", "Leadscale industry": "Recreational Facilities & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Historical Sites", "Leadscale industry": "Museums & Institutions", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Hospitality", "Leadscale industry": "Hospitality", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Hotels and Motels", "Leadscale industry": "Hospitality", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Mobile Food Services", "Leadscale industry": "Restaurants", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Museums", "Leadscale industry": "Museums & Institutions", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Museums, Historical Sites, and Zoos", "Leadscale industry": "Museums & Institutions", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Performing Arts", "Leadscale industry": "Performing Arts", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Performing Arts and Spectator Sports", "Leadscale industry": "Performing Arts", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined performing arts and sports; Sports may apply."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Performing Arts and Spectator Sports", "Leadscale industry": "Sports", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Combined performing arts and sports; Sports may apply."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Racetracks", "Leadscale industry": "Sports", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Sports venue assumed; betting operators belong to Gambling & Casinos."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Racetracks", "Leadscale industry": "Gambling & Casinos", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "Sports venue assumed; betting operators belong to Gambling & Casinos."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Recreational Facilities", "Leadscale industry": "Recreational Facilities & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Restaurants", "Leadscale industry": "Restaurants", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Sightseeing Transportation", "Leadscale industry": "Leisure, Travel & Tourism", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Skiing Facilities", "Leadscale industry": "Recreational Facilities & Services", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Spectator Sports", "Leadscale industry": "Sports", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Sports and Recreation Instruction", "Leadscale industry": "Professional Training & Coaching", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Sports Teams and Clubs", "Leadscale industry": "Sports", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Theater Companies", "Leadscale industry": "Performing Arts", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Travel Arrangements", "Leadscale industry": "Leisure, Travel & Tourism", "Priority": 1, "Approved": "Yes", "Context required": "No", "Notes": "Use this exact Leadscale label when selected for the campaign."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Zoos and Botanical Gardens", "Leadscale industry": "Museums & Institutions", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No direct zoo/botanical-garden category; Museums & Institutions is a fallback."}, {"Main industry": "Travel & Hospitality", "Sub-industry": "Zoos and Botanical Gardens", "Leadscale industry": "Recreational Facilities & Services", "Priority": 1, "Approved": "Yes", "Context required": "Yes", "Notes": "No direct zoo/botanical-garden category; Museums & Institutions is a fallback."}]')

# --- STREAMLIT USER INTERFACE ---
st.title("Wholesale Lead Quality Checker V2")
st.caption("Enhanced with PostgreSQL Caching, Pre-Flight Cost Controls, and Optional AI Web Search Verification.")

with st.sidebar:
    st.header("Configuration & Cost Control")
    admin_password = st.text_input("Admin Password (Access Protection)", type="password")
    use_ai = st.toggle("Enable AI Web Research", value=False, help="Uses OpenAI API with web search to verify domain/company mismatches.")
    
    env_key = get_secret("OPENAI_API_KEY")
    if env_key:
        openai_key = env_key
    else:
        openai_key = st.text_input("OpenAI API Key", type="password", help="Enter your secret API key. Never shared.")
        
    max_ai_calls = st.slider("Max AI Lookups Per Run (Budget Cap)", min_value=5, max_value=100, value=25, step=5, help="Hard stop to prevent runaway API costs.")

master_file = st.file_uploader("Upload Wholesale Master (.xlsx)", type=["xlsx"])
picklist_file = st.file_uploader("Upload Wholesale Picklist (.xlsx) [Optional]", type=["xlsx"])

master_sheet = None
picklist_sheets: List[str] = []
master_bytes = None
picklist_bytes = None

if master_file is not None:
    master_bytes = master_file.getvalue()
    master_workbook = load_workbook(io.BytesIO(master_bytes), read_only=True)
    master_sheet = st.selectbox("Master sheet to process", master_workbook.sheetnames)

if picklist_file is not None:
    picklist_bytes = picklist_file.getvalue()
    picklist_workbook = load_workbook(io.BytesIO(picklist_bytes), read_only=True)
    default_sheets = []
    for sheet in picklist_workbook.sheetnames:
        sheet_key = clean_header(sheet)
        if sheet_key in {"importtemplate", "leads", "sheet1"} or "template" in sheet_key:
            default_sheets.append(sheet)
    if not default_sheets and picklist_workbook.sheetnames:
        default_sheets = [picklist_workbook.sheetnames[0]]
    picklist_sheets = st.multiselect("Picklist sheet(s) to use", picklist_workbook.sheetnames, default=default_sheets, help="Select multiple sheets if phone exclusions or mappings are on separate tabs.")


# Standard industry mapping: built in, runs automatically with normal QA.
industry_rule_frame = pd.DataFrame(BUNDLED_INDUSTRY_RULES)
industry_campaign_values = []
industry_source_column = None
industry_campaign_sheet = None
industry_campaign_column = None
industry_setup_note = ''
if master_bytes and master_sheet:
    source_headers = list(pd.read_excel(io.BytesIO(master_bytes), sheet_name=master_sheet, nrows=0).columns)
    sub_aliases = {clean_header(v) for v in FIELD_ALIASES['sub_industry']} | {'sub industries', 'sub industry name', 'company sub industry'}
    source_matches = [h for h in source_headers if clean_header(h) in sub_aliases]
    if len(source_matches) == 1:
        industry_source_column = source_matches[0]
    elif not source_matches:
        main_matches = [h for h in source_headers if clean_header(h) in {'industry', 'company industry', 'c industry'}]
        if len(main_matches) == 1:
            industry_source_column = main_matches[0]
if picklist_bytes:
    candidate_sheets = picklist_sheets or load_workbook(io.BytesIO(picklist_bytes), read_only=True).sheetnames
    campaign_frames = {name: pd.read_excel(io.BytesIO(picklist_bytes), sheet_name=name, dtype=str, keep_default_na=False) for name in candidate_sheets}
    industry_headers = {'industry', 'industries', 'company industry', 'leadscale industry', 'target industry', 'target industries', 'industry value', 'industry values', 'accepted industries'}
    candidate_columns = [(name, col) for name, frame in campaign_frames.items() for col in frame.columns if clean_header(col) in industry_headers]
    if len(candidate_columns) == 1:
        industry_campaign_sheet, industry_campaign_column = candidate_columns[0]
    else:
        industry_setup_note = 'Campaign industry column is ambiguous or missing. Choose it below; otherwise industry rows will be flagged for review.'
with st.expander('Industry mapping — automatic setup', expanded=bool(industry_setup_note)):
    st.caption('Industry mapping runs as standard using the built-in grid. Exact labels, known alternative names and clear spelling matches are checked; uncertain matches stay in review. Existing QA scores are unchanged. Use one campaign per run.')
    if master_bytes and master_sheet:
        industry_source_column = st.selectbox('Returned industry / sub-industry column', source_headers,
            index=source_headers.index(industry_source_column) if industry_source_column in source_headers else None,
            key='standard_industry_source')
    if picklist_bytes:
        all_sheets = load_workbook(io.BytesIO(picklist_bytes), read_only=True).sheetnames
        industry_campaign_sheet = st.selectbox('Campaign industry picklist sheet', all_sheets,
            index=all_sheets.index(industry_campaign_sheet) if industry_campaign_sheet in all_sheets else None,
            key='standard_industry_sheet')
        if industry_campaign_sheet is not None:
            selected_frame = pd.read_excel(io.BytesIO(picklist_bytes), sheet_name=industry_campaign_sheet, dtype=str, keep_default_na=False)
            columns = list(selected_frame.columns)
            industry_campaign_column = st.selectbox('Campaign selected industry labels', columns,
                index=columns.index(industry_campaign_column) if industry_campaign_column in columns else None,
                key='standard_industry_column')
            if industry_campaign_column is not None:
                industry_campaign_values = selected_frame[industry_campaign_column].tolist()
                try:
                    st.write('Selected campaign industry labels:', list(compile_campaign_industries(industry_campaign_values).values()))
                except ValueError as exc:
                    st.warning(str(exc))
    else:
        st.caption('No campaign picklist uploaded: industry mapping will be flagged for review; the other checks still run.')
    st.download_button('Download built-in industry mapping for reference', industry_grid_bytes(industry_rule_frame),
        file_name='built_in_industry_rules.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

apply_colours = st.toggle("Colour-code Excel results", value=True)
show_debug = st.toggle("Show debug panel", value=True)

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "debug_info" not in st.session_state:
    st.session_state.debug_info = None

# Every output is tied to the current master, campaign and options.
signature_payload = {
    'source': industry_source_column, 'campaign_values': industry_campaign_values,
    'master_sheet': master_sheet, 'picklist_sheets': picklist_sheets,
    'colours': apply_colours, 'ai': use_ai, 'ai_cap': max_ai_calls,
}
digest = hashlib.sha256(json.dumps(signature_payload, sort_keys=True).encode())
digest.update(master_bytes or b'')
digest.update(picklist_bytes or b'')
mapping_signature = digest.hexdigest()

can_run = master_bytes is not None and master_sheet is not None

if use_ai and not openai_key:
    st.warning("Please enter your OpenAI API key in the sidebar to use AI Web Research.")

if st.button("Run Wholesale QA V2", type="primary", use_container_width=True, disabled=not can_run):
    try:
        with st.spinner("Processing local rules, database cache, and AI lookups..."):
            output_bytes, debug_info = process_file(
                master_bytes=master_bytes,
                picklist_bytes=picklist_bytes,
                master_sheet=master_sheet,
                picklist_sheets=picklist_sheets,
                apply_colours=apply_colours,
                use_ai=use_ai,
                openai_key=openai_key,
                max_ai_calls=max_ai_calls
            )
            output_bytes, industry_mapping_summary = append_industry_mapping(
                output_bytes, industry_source_column, industry_rule_frame, industry_campaign_values, apply_colours, allow_missing_setup=True)
            debug_info["industry_mapping"] = industry_mapping_summary
        st.session_state.output_bytes = output_bytes
        st.session_state.completed_mapping_signature = mapping_signature
        st.session_state.debug_info = debug_info
        st.success("Processing complete!")
    except Exception as error:
        st.session_state.output_bytes = None
        st.session_state.debug_info = None
        st.error(f"Error: {error}")

if st.session_state.debug_info and show_debug:
    with st.expander("Debug & Execution Summary", expanded=True):
        st.json(st.session_state.debug_info)

if st.session_state.output_bytes and st.session_state.get("completed_mapping_signature") == mapping_signature:
    original_name = master_file.name if master_file is not None else "wholesale_master.xlsx"
    base_name = original_name.rsplit(".", 1)[0]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"{base_name}_v2_results_{timestamp}.xlsx"
    st.download_button(label="Download Enriched Wholesale QA Results", data=st.session_state.output_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
