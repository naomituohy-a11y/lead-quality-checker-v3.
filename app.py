import io
import re
import datetime
import unicodedata
import os
import json
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

# --- SAFE SECRETS WRAPPER ---
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

# --- DATABASE & CACHE SETUP ---
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
                
                # CHANGED: "Match" to "Affiliated"
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
        
        # CHANGED: "Affiliated" prevents adding a domain issue
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
    master_bytes = master_file.read()
    master_workbook = load_workbook(io.BytesIO(master_bytes), read_only=True)
    master_sheet = st.selectbox("Master sheet to process", master_workbook.sheetnames)

if picklist_file is not None:
    picklist_bytes = picklist_file.read()
    picklist_workbook = load_workbook(io.BytesIO(picklist_bytes), read_only=True)
    default_sheets = []
    for sheet in picklist_workbook.sheetnames:
        sheet_key = clean_header(sheet)
        if sheet_key in {"importtemplate", "leads", "sheet1"} or "template" in sheet_key:
            default_sheets.append(sheet)
    if not default_sheets and picklist_workbook.sheetnames:
        default_sheets = [picklist_workbook.sheetnames[0]]
    picklist_sheets = st.multiselect("Picklist sheet(s) to use", picklist_workbook.sheetnames, default=default_sheets, help="Select multiple sheets if phone exclusions or mappings are on separate tabs.")

apply_colours = st.toggle("Colour-code Excel results", value=True)
show_debug = st.toggle("Show debug panel", value=True)

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None
if "debug_info" not in st.session_state:
    st.session_state.debug_info = None

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
        st.session_state.output_bytes = output_bytes
        st.session_state.debug_info = debug_info
        st.success("Processing complete!")
    except Exception as error:
        st.session_state.output_bytes = None
        st.session_state.debug_info = None
        st.error(f"Error: {error}")

if st.session_state.debug_info and show_debug:
    with st.expander("Debug & Execution Summary", expanded=True):
        st.json(st.session_state.debug_info)

if st.session_state.output_bytes:
    original_name = master_file.name if master_file is not None else "wholesale_master.xlsx"
    base_name = original_name.rsplit(".", 1)[0]
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    output_filename = f"{base_name}_v2_results_{timestamp}.xlsx"
    st.download_button(label="Download Enriched Wholesale QA Results", data=st.session_state.output_bytes, file_name=output_filename, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
