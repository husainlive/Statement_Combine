import argparse
import hashlib
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
from decimal import Decimal, InvalidOperation
from datetime import datetime, timedelta
from email import policy
from email.parser import BytesParser
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

try:
    import pdfplumber
except ModuleNotFoundError as error:
    missing_package = error.name or "pdfplumber"
    raise SystemExit(
        f"Missing Python package: {missing_package}\n"
        f"Current interpreter: {sys.executable}\n"
        "Install it in this same interpreter with:\n"
        f'"{sys.executable}" -m pip install pdfplumber'
    ) from error

try:
    from pypdf import PdfReader, PdfWriter
except ModuleNotFoundError as error:
    missing_package = error.name or "pypdf"
    raise SystemExit(
        f"Missing Python package: {missing_package}\n"
        f"Current interpreter: {sys.executable}\n"
        "Install it in this same interpreter with:\n"
        f'"{sys.executable}" -m pip install pypdf'
    ) from error

try:
    import extract_msg
except ModuleNotFoundError:
    extract_msg = None

try:
    import pythoncom
    import win32com.client
except ModuleNotFoundError:
    pythoncom = None
    win32com = None

# =========================
# SETTINGS
# =========================

BASE_FOLDER = Path(__file__).resolve().parent
SETTINGS_FILE = BASE_FOLDER / "pdf_classifier_settings.json"

DEFAULT_PATH_SETTINGS = {
    "mail": BASE_FOLDER / "input_mails",
    "input": BASE_FOLDER / "input_pdfs",
    "output": BASE_FOLDER / "classified_pdfs",
    "combined": BASE_FOLDER / "combined_pdfs",
    "accounts": BASE_FOLDER / "account_numbers.txt",
}


def load_path_settings() -> dict[str, Path]:
    path_settings = DEFAULT_PATH_SETTINGS.copy()
    if not SETTINGS_FILE.exists():
        return path_settings

    try:
        raw_settings = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return path_settings

    if not isinstance(raw_settings, dict):
        return path_settings

    for key in path_settings:
        raw_value = raw_settings.get(key)
        if isinstance(raw_value, str) and raw_value.strip():
            saved_path = Path(raw_value.strip()).expanduser()
            default_path = DEFAULT_PATH_SETTINGS[key]
            if key == "accounts":
                is_usable = saved_path.exists()
            else:
                is_usable = saved_path.exists() or saved_path.parent.exists()

            if is_usable:
                path_settings[key] = saved_path
            else:
                path_settings[key] = default_path

    return path_settings


def write_path_settings(path_settings: dict[str, Path]) -> None:
    payload = {
        key: str(path_settings.get(key, default_path).expanduser())
        for key, default_path in DEFAULT_PATH_SETTINGS.items()
    }
    SETTINGS_FILE.write_text(json.dumps(payload, indent=2), encoding="utf-8")


PATH_SETTINGS = load_path_settings()
MAIL_SOURCE_FOLDER = PATH_SETTINGS["mail"]
SOURCE_FOLDER = PATH_SETTINGS["input"]
OUTPUT_FOLDER = PATH_SETTINGS["output"]
COMBINED_FOLDER = PATH_SETTINGS["combined"]
ACCOUNT_NUMBERS_FILE = PATH_SETTINGS["accounts"]

# Choose True to copy files, False to move files
COPY_INSTEAD_OF_MOVE = True

# Choose True to clear input PDFs after a successful classification run
CLEAR_INPUT_AFTER_RUN = True

# Choose True to clear processed mail files after a successful classification run
CLEAR_MAIL_AFTER_RUN = True

MT940_SUBFOLDER_NAME = "mt940"


def normalize_account(value: str) -> str:
    return re.sub(r"[\s\-/]", "", value).upper()


ACCOUNT_SEPARATOR_PATTERN = r"[\s\-/]*"
PRIMARY_ACCOUNT_LABEL_PATTERN = re.compile(
    r"(?i)(?:"
    r"\baccount\s*(?:number|no\.?|#)\b"
    r"|\bacct\s*(?:number|no\.?|#)\b"
    r"|\ba/c\s*(?:number|no\.?)?\b"
    r"|\biban\b"
    r")"
)
GENERAL_ACCOUNT_LABEL_PATTERN = re.compile(r"(?i)(?:\baccount\b|\bacct\b|\ba/c\b)")


def parse_account_line(line: str) -> dict[str, str | bool] | None:
    cleaned = line.strip()
    if not cleaned or cleaned.startswith("#"):
        return None

    parts = [part.strip() for part in cleaned.split("|")]
    if len(parts) == 1:
        account_number = parts[0]
        return {
            "account_number": account_number,
            "normalized_account": normalize_account(account_number),
            "password_protected": False,
            "password": "",
        }

    account_number = parts[0]
    protection_value = parts[1].lower() if len(parts) > 1 else "no"
    password = parts[2] if len(parts) > 2 else ""
    password_protected = protection_value in {"yes", "true", "1", "y"}
    return {
        "account_number": account_number,
        "normalized_account": normalize_account(account_number),
        "password_protected": password_protected,
        "password": password if password_protected else "",
    }


def format_account_line(account_number: str, password_protected: bool, password: str) -> str:
    return f"{account_number}|{'yes' if password_protected else 'no'}|{password if password_protected else ''}"


def load_account_records(txt_path: Path) -> list[dict[str, str | bool]]:
    if not txt_path.exists():
        print(f"Account list file not found: {txt_path}")
        return []

    account_records: list[dict[str, str | bool]] = []
    for line in txt_path.read_text(encoding="utf-8").splitlines():
        parsed = parse_account_line(line)
        if parsed:
            account_records.append(parsed)
    return account_records


def load_account_numbers(txt_path: Path) -> list[str]:
    return [str(record["account_number"]) for record in load_account_records(txt_path)]


def save_account_number(
    account_number: str,
    password_protected: bool,
    password: str,
    txt_path: Path,
) -> tuple[bool, str]:
    cleaned = account_number.strip()
    if not cleaned:
        return False, "Enter an account number first."

    cleaned_password = password.strip()
    if password_protected and not cleaned_password:
        return False, "Enter the password for the protected account."

    existing_records = load_account_records(txt_path)
    normalized_existing = {str(record["normalized_account"]) for record in existing_records}

    if normalize_account(cleaned) in normalized_existing:
        return False, f"Account number already exists: {cleaned}"

    txt_path.parent.mkdir(parents=True, exist_ok=True)
    with txt_path.open("a", encoding="utf-8") as file_handle:
        if txt_path.stat().st_size > 0:
            file_handle.write("\n")
        file_handle.write(format_account_line(cleaned, password_protected, cleaned_password))

    return True, f"Saved account number: {cleaned}"


KNOWN_ACCOUNT_RECORDS = load_account_records(ACCOUNT_NUMBERS_FILE)
KNOWN_ACCOUNTS = [str(record["account_number"]) for record in KNOWN_ACCOUNT_RECORDS]


def normalize_text(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def build_password_candidates() -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()

    for record in KNOWN_ACCOUNT_RECORDS:
        if not bool(record["password_protected"]):
            continue
        password = str(record["password"]).strip()
        if password and password not in seen:
            candidates.append(password)
            seen.add(password)

    return candidates


def extract_text_with_password(pdf_path: Path, password: str | None) -> str:
    all_text = []

    with pdfplumber.open(pdf_path, password=password) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                all_text.append(text)

    return "\n".join(all_text)


def extract_text_from_pdf(pdf_path: Path) -> tuple[str, str | None]:
    try:
        return extract_text_with_password(pdf_path, None), None
    except Exception:
        pass

    for password in build_password_candidates():
        try:
            return extract_text_with_password(pdf_path, password), password
        except Exception:
            continue

    return "", None


def account_value_pattern(normalized_account: str) -> re.Pattern:
    separated_account = ACCOUNT_SEPARATOR_PATTERN.join(
        re.escape(character)
        for character in normalized_account
    )
    return re.compile(
        rf"(?<![A-Z0-9]){separated_account}(?![A-Z0-9]|[ \t]*[\-/][ \t]*[A-Z0-9])",
        flags=re.IGNORECASE,
    )


def find_known_account_occurrences(text: str) -> list[tuple[int, int, str]]:
    occurrences: list[tuple[int, int, str]] = []
    for record in KNOWN_ACCOUNT_RECORDS:
        account_number = str(record["account_number"])
        normalized_account = str(record["normalized_account"])
        if not normalized_account:
            continue
        for match in account_value_pattern(normalized_account).finditer(text):
            occurrences.append((match.start(), match.end(), account_number))

    return sorted(occurrences, key=lambda occurrence: occurrence[0])


def append_account_candidate(candidates: list[str], seen: set[str], candidate: str) -> None:
    normalized_candidate = normalize_account(candidate)
    if not normalized_candidate or normalized_candidate in seen:
        return
    seen.add(normalized_candidate)
    candidates.append(candidate)


def find_possible_accounts(text: str) -> list[str]:
    candidates: list[str] = []
    seen: set[str] = set()
    detected_matches: list[tuple[int, int, str]] = []

    for start, _end, account_number in find_known_account_occurrences(text):
        detected_matches.append((start, 0, account_number))

    for match in re.finditer(r"\b\d{10,24}\b", text):
        detected_matches.append((match.start(), 1, match.group(0)))

    for match in re.finditer(r"\b(?:\d{2,8}[\s\-/]?){2,7}\d{2,8}\b", text):
        cleaned = normalize_account(match.group(0))
        if 10 <= len(cleaned) <= 24 and cleaned.isdigit():
            detected_matches.append((match.start(), 1, cleaned))

    for match in re.finditer(r"\b[A-Z]{2}\d{2}[A-Z0-9]{10,30}\b", text, flags=re.IGNORECASE):
        detected_matches.append((match.start(), 1, normalize_account(match.group(0))))

    for _start, _priority, candidate in sorted(detected_matches, key=lambda item: (item[0], item[1])):
        append_account_candidate(candidates, seen, candidate)

    return candidates


def account_context_priority(text: str, start: int, end: int) -> int:
    line_start = text.rfind("\n", 0, start) + 1
    line_end = text.find("\n", end)
    if line_end == -1:
        line_end = len(text)

    line_text = text[line_start:line_end]
    nearby_before = text[max(line_start, start - 120):start]
    context = f"{nearby_before} {line_text}"

    if PRIMARY_ACCOUNT_LABEL_PATTERN.search(context):
        return 3
    if GENERAL_ACCOUNT_LABEL_PATTERN.search(context):
        return 2
    if start < 1500:
        return 1
    return 0


def preferred_known_account_from_text(text: str) -> str | None:
    best_by_account: dict[str, tuple[int, int]] = {}
    for start, end, account_number in find_known_account_occurrences(text):
        priority = account_context_priority(text, start, end)
        current_best = best_by_account.get(account_number)
        if current_best is None or (priority, -start) > (current_best[0], -current_best[1]):
            best_by_account[account_number] = (priority, start)

    if not best_by_account:
        return None

    ranked_accounts = sorted(
        (
            (priority, start, account_number)
            for account_number, (priority, start) in best_by_account.items()
        ),
        key=lambda item: (-item[0], item[1]),
    )
    return ranked_accounts[0][2]


def classify_account(
    found_accounts: list[str],
    allow_unlisted: bool = False,
    text: str | None = None,
) -> str:
    if text:
        preferred_account = preferred_known_account_from_text(text)
        if preferred_account:
            return preferred_account

    known_accounts_by_normalized = {
        str(record["normalized_account"]): str(record["account_number"])
        for record in KNOWN_ACCOUNT_RECORDS
    }

    for found in found_accounts:
        normalized_found = normalize_account(found)
        if normalized_found in known_accounts_by_normalized:
            return known_accounts_by_normalized[normalized_found]

    if allow_unlisted:
        for found in found_accounts:
            cleaned = found.strip()
            if cleaned:
                return cleaned

    return "Unclassified"


def decode_text_payload(payload: bytes) -> str:
    for encoding in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            return payload.decode(encoding)
        except UnicodeDecodeError:
            continue
    return payload.decode("latin-1", errors="ignore")


def append_unique_accounts(
    destination: list[str],
    seen: set[str],
    candidates: list[str],
) -> None:
    for candidate in candidates:
        normalized_candidate = normalize_account(candidate)
        if not normalized_candidate or normalized_candidate in seen:
            continue
        seen.add(normalized_candidate)
        destination.append(candidate)


def find_mt940_accounts(text: str) -> list[str]:
    ordered_candidates: list[str] = []
    seen: set[str] = set()

    for match in re.finditer(r"(?im)^:25[A-Z]?:\s*(.+)$", text):
        raw_value = match.group(1).strip().replace("/", " ")
        append_unique_accounts(ordered_candidates, seen, find_possible_accounts(raw_value))

    append_unique_accounts(ordered_candidates, seen, find_possible_accounts(text))
    return ordered_candidates


def looks_like_mt940_text(text: str) -> bool:
    tag_matches = re.findall(r"(?m)^:(?:20|25[A-Z]?|28C?|60[FM]?|61|62[FM]?):", text)
    return len(tag_matches) >= 2


def is_supported_dat_attachment(
    filename: str,
    content_type: str,
    payload: bytes,
) -> bool:
    if filename.lower().endswith(".dat"):
        return True

    if content_type.lower() not in {
        "application/octet-stream",
        "application/x-unknown-content-type",
        "text/plain",
        "application/dat",
        "application/x-dat",
    }:
        return False

    sample_text = decode_text_payload(payload[:8192])
    return looks_like_mt940_text(sample_text)


def extract_mt940_date(text: str, fallback_file_date: datetime) -> datetime:
    for pattern in (
        r"(?im)^:62[FM]?:[CD](\d{6})",
        r"(?im)^:60[FM]?:[CD](\d{6})",
    ):
        match = re.search(pattern, text)
        if not match:
            continue
        try:
            return datetime.strptime(match.group(1), "%y%m%d")
        except ValueError:
            continue

    transaction_dates = re.findall(r"(?im)^:61:(\d{6})", text)
    for raw_value in reversed(transaction_dates):
        try:
            return datetime.strptime(raw_value, "%y%m%d")
        except ValueError:
            continue

    return fallback_file_date


def detect_bank_date_preference(text: str) -> str:
    upper_text = text.upper()
    if "ALRAYAN BANK" in upper_text or "AL RAYAN" in upper_text:
        return "day_first"
    if "QATAR NATIONAL BANK (Q.P.S.C.)" in upper_text:
        return "month_first"
    return "day_first"


def is_qnb_statement(text: str) -> bool:
    return "QATAR NATIONAL BANK (Q.P.S.C.)" in text.upper()


def adjust_detected_statement_date(text: str, detected_date: datetime) -> datetime:
    if is_qnb_statement(text):
        return detected_date - timedelta(days=1)
    return detected_date


def parse_flexible_date(raw: str, preference: str = "day_first") -> datetime | None:
    cleaned = re.sub(r"[\u200e\u200f]", "", raw).strip()
    cleaned = re.sub(r"(?i)\b(st|nd|rd|th)\b", "", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned)
    cleaned = cleaned.replace(".", "/")

    day_first_formats = (
        "%d/%m/%Y",
        "%d/%m/%y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%d %m %Y",
        "%d %m %y",
        "%d %b %Y",
        "%d %b %y",
        "%d %B %Y",
        "%d %B %y",
        "%d-%b-%Y",
        "%d-%b-%y",
        "%d-%B-%Y",
        "%d-%B-%y",
        "%d/%b/%Y",
        "%d/%b/%y",
        "%d/%B/%Y",
        "%d/%B/%y",
    )
    month_first_formats = (
        "%m/%d/%Y",
        "%m/%d/%y",
        "%m-%d-%Y",
        "%m-%d-%y",
        "%m %d %Y",
        "%m %d %y",
        "%b %d %Y",
        "%b %d %y",
        "%B %d %Y",
        "%B %d %y",
        "%b %d, %Y",
        "%b %d, %y",
        "%B %d, %Y",
        "%B %d, %y",
    )
    year_first_formats = (
        "%Y-%m-%d",
        "%Y/%m/%d",
        "%Y.%m.%d",
        "%Y%m%d",
        "%Y %b %d",
        "%Y %B %d",
    )

    format_groups = [day_first_formats, month_first_formats, year_first_formats]
    if preference == "month_first":
        format_groups = [month_first_formats, day_first_formats, year_first_formats]

    candidates = [cleaned]
    if "," in cleaned:
        candidates.append(cleaned.replace(",", ""))
    if "/" in cleaned:
        candidates.append(cleaned.replace("/", "-"))
        candidates.append(cleaned.replace("/", " "))
    if "-" in cleaned:
        candidates.append(cleaned.replace("-", "/"))
        candidates.append(cleaned.replace("-", " "))

    seen_candidates: set[str] = set()
    for candidate in candidates:
        candidate = candidate.strip()
        if not candidate or candidate in seen_candidates:
            continue
        seen_candidates.add(candidate)
        for format_group in format_groups:
            for fmt in format_group:
                try:
                    return datetime.strptime(candidate, fmt)
                except ValueError:
                    pass
    return None


def extract_date(text: str, fallback_file_date: datetime) -> datetime:
    date_preference = detect_bank_date_preference(text)
    statement_range = extract_statement_range(text, date_preference)
    if statement_range:
        _start_date, end_date = statement_range
        return adjust_detected_statement_date(text, end_date)

    date_token = (
        r"(?:"
        r"\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}"
        r"|"
        r"\d{4}[./\-]\d{1,2}[./\-]\d{1,2}"
        r"|"
        r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}"
        r"|"
        r"[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}"
        r")"
    )

    search_texts = [text]
    normalized_text = normalize_text(text)
    if normalized_text != text:
        search_texts.append(normalized_text)

    prioritized_patterns = [
        r"(?im)^\s*Statement\s+Date\s*[:\-]?\s*([A-Za-z0-9,./\- ]{6,30})$",
        r"(?im)^\s*Date\s*[:\-]?\s*([A-Za-z0-9,./\- ]{6,30})$",
        r"(?im)^\s*As\s+of\s*[:\-]?\s*([A-Za-z0-9,./\- ]{6,30})$",
        r"(?im)^\s*Statement\s+Period\s*[:\-]?\s*[A-Za-z0-9,./\- ]+\s+To\s+([A-Za-z0-9,./\- ]{6,30})$",
        r"(?im)^\s*Period\s*[:\-]?\s*[A-Za-z0-9,./\- ]+\s+To\s+([A-Za-z0-9,./\- ]{6,30})$",
        rf"(?i)\bStatement\s+Date\s*[:\-]?\s*({date_token})\b",
        rf"(?i)\bDate\s*[:\-]?\s*({date_token})\b",
        rf"(?i)\bAs\s+of\s*[:\-]?\s*({date_token})\b",
    ]

    for search_text in search_texts:
        for pattern in prioritized_patterns:
            match = re.search(pattern, search_text, flags=re.IGNORECASE)
            if not match:
                continue
            parsed = parse_flexible_date(match.group(1), date_preference)
            if parsed:
                return adjust_detected_statement_date(text, parsed)

    return fallback_file_date


def safe_filename(name: str) -> str:
    return re.sub(r'[<>:"/\\|?*]', "_", name)


def unique_path(path: Path) -> Path:
    if not path.exists():
        return path

    counter = 1
    while True:
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        if not candidate.exists():
            return candidate
        counter += 1


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as file_handle:
        for chunk in iter(lambda: file_handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def bytes_sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def make_writable(path: Path) -> None:
    try:
        os.chmod(path, 0o666)
    except OSError:
        pass


def folder_pdf_hashes(folder_path: Path) -> dict[str, Path]:
    hashes: dict[str, Path] = {}
    if not folder_path.exists():
        return hashes

    for pdf_file in sorted(folder_path.glob("*.pdf")):
        if pdf_file.name.lower() == "combined_statements.pdf":
            continue
        try:
            hashes[file_sha256(pdf_file)] = pdf_file
        except OSError:
            continue

    return hashes


def save_payload_file(destination_folder: Path, filename: str, payload: bytes) -> tuple[Path, str]:
    destination_folder.mkdir(parents=True, exist_ok=True)
    payload_hash = bytes_sha256(payload)

    for existing_file in sorted(destination_folder.iterdir()):
        if not existing_file.is_file():
            continue
        if file_sha256(existing_file) == payload_hash:
            return existing_file, f"Skipped duplicate content (matches {existing_file.name})"

    output_name = safe_filename(filename)
    output_path = destination_folder / output_name

    if output_path.exists():
        if file_sha256(output_path) == payload_hash:
            return output_path, "Skipped existing"
        output_path = unique_path(output_path)

    output_path.write_bytes(payload)
    return output_path, "Saved"


def save_attachment_bytes(destination_folder: Path, filename: str, payload: bytes) -> tuple[Path | None, bool]:
    output_path, action = save_payload_file(destination_folder, filename, payload)
    return output_path, action == "Saved"


def iter_supported_message_parts(message) -> list[tuple[str, bytes]]:
    supported_parts: list[tuple[str, bytes]] = []
    for part in message.walk():
        if part.is_multipart():
            continue

        filename = part.get_filename() or ""
        payload = part.get_payload(decode=True)
        if not payload:
            continue

        content_type = part.get_content_type()
        lower_filename = filename.lower()
        if lower_filename.endswith(".pdf") or part.get_content_type() == "application/pdf":
            output_name = filename or "attachment.pdf"
            supported_parts.append((output_name, payload))
        elif is_supported_dat_attachment(filename, content_type, payload):
            output_name = filename if lower_filename.endswith(".dat") else f"{Path(filename).stem or 'attachment'}.dat"
            supported_parts.append((output_name, payload))

    return supported_parts


def iter_pdf_message_parts(message) -> list[tuple[str, bytes]]:
    return [
        (filename, payload)
        for filename, payload in iter_supported_message_parts(message)
        if filename.lower().endswith(".pdf")
    ]


def decrypt_smime_payload(payload: bytes) -> bytes:
    input_path = BASE_FOLDER / f"smime_{next(tempfile._get_candidate_names())}.p7m"
    try:
        input_path.write_bytes(payload)
        escaped_input_path = str(input_path).replace("'", "''")
        command = [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            (
                "$ErrorActionPreference='Stop'; "
                "Add-Type -AssemblyName System.Security; "
                "$cms=New-Object System.Security.Cryptography.Pkcs.EnvelopedCms; "
                f"$cms.Decode([System.IO.File]::ReadAllBytes('{escaped_input_path}')); "
                "$cms.Decrypt(); "
                "[Console]::OpenStandardOutput().Write($cms.ContentInfo.Content,0,$cms.ContentInfo.Content.Length)"
            ),
        ]
        result = subprocess.run(
            command,
            capture_output=True,
            check=True,
            timeout=60,
        )
        return result.stdout
    finally:
        if input_path.exists():
            try:
                delete_file(input_path)
            except OSError:
                pass


def extract_supported_attachments_from_smime_payload(smime_payload: bytes) -> list[tuple[str, bytes]]:
    decrypted_message = BytesParser(policy=policy.default).parsebytes(
        decrypt_smime_payload(smime_payload)
    )
    return iter_supported_message_parts(decrypted_message)


def extract_pdfs_from_smime_payload(smime_payload: bytes, destination_folder: Path) -> list[Path]:
    extracted_files: list[Path] = []
    for filename, payload in extract_supported_attachments_from_smime_payload(smime_payload):
        if not filename.lower().endswith(".pdf"):
            continue
        output_path, was_created = save_attachment_bytes(destination_folder, filename, payload)
        if was_created and output_path is not None:
            extracted_files.append(output_path)
    return extracted_files


def write_unprotected_pdf(source_path: Path, destination_path: Path, password: str) -> None:
    reader = PdfReader(str(source_path))
    if reader.is_encrypted:
        decrypt_result = reader.decrypt(password)
        if decrypt_result == 0:
            raise ValueError(f"Could not decrypt PDF for export: {source_path.name}")

    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)

    with destination_path.open("wb") as output_handle:
        writer.write(output_handle)


def merge_pdfs_in_folder(folder_path: Path) -> str | None:
    pdf_files = sorted(
        (
            path
            for path in folder_path.glob("*.pdf")
            if path.name.lower() != "combined_statements.pdf"
        ),
        key=lambda path: (path.name[:10], path.name.lower()),
    )

    if not pdf_files:
        return None

    writer = PdfWriter()
    seen_hashes: set[str] = set()
    merged_count = 0
    for pdf_file in pdf_files:
        pdf_hash = file_sha256(pdf_file)
        if pdf_hash in seen_hashes:
            continue
        seen_hashes.add(pdf_hash)
        reader = PdfReader(str(pdf_file))
        for page in reader.pages:
            writer.add_page(page)
        merged_count += 1

    if merged_count == 0:
        return None

    combined_path = folder_path / "combined_statements.pdf"
    try:
        with combined_path.open("wb") as output_handle:
            writer.write(output_handle)
    except PermissionError:
        return f"Could not update combined PDF because it is open: {combined_path}"

    skipped_count = len(pdf_files) - merged_count
    if skipped_count > 0:
        return f"Combined {merged_count} PDFs ({skipped_count} duplicates skipped) -> {combined_path}"
    return f"Combined {merged_count} PDFs -> {combined_path}"


def build_combined_output_name(account_name: str, period_name: str) -> str:
    safe_account = safe_filename(account_name)
    return f"{safe_account}_{period_name.replace('-', '_')}.pdf"


def combined_period_folder(period_name: str) -> Path:
    return COMBINED_FOLDER / period_name


def combine_classified_pdfs() -> list[str]:
    results: list[str] = []
    if not OUTPUT_FOLDER.exists():
        return results

    COMBINED_FOLDER.mkdir(parents=True, exist_ok=True)
    for account_dir in sorted(path for path in OUTPUT_FOLDER.iterdir() if path.is_dir()):
        for period_dir in sorted(path for path in account_dir.iterdir() if path.is_dir()):
            result = merge_pdfs_in_folder(period_dir)
            if result:
                results.append(result)

            combined_source = period_dir / "combined_statements.pdf"
            if combined_source.exists():
                combined_destination_dir = combined_period_folder(period_dir.name)
                combined_destination_dir.mkdir(parents=True, exist_ok=True)
                combined_destination = combined_destination_dir / build_combined_output_name(
                    account_dir.name,
                    period_dir.name,
                )
                shutil.copy2(combined_source, combined_destination)
                results.append(f"Saved combined PDF -> {combined_destination}")

    return results


def get_statement_dates_from_folder(folder_path: Path) -> list[datetime]:
    statement_dates: set[datetime] = set()
    for pdf_file in folder_path.glob("*.pdf"):
        if pdf_file.name.lower() == "combined_statements.pdf":
            continue
        match = re.match(r"^(\d{4}-\d{2}-\d{2})_", pdf_file.name)
        if not match:
            continue
        try:
            statement_dates.add(datetime.strptime(match.group(1), "%Y-%m-%d"))
        except ValueError:
            continue

    return sorted(statement_dates)


def parse_money_value(raw_value: str) -> Decimal | None:
    cleaned = raw_value.strip().replace(",", "")
    if not cleaned:
        return None

    negative = cleaned.endswith("-")
    if negative:
        cleaned = cleaned[:-1]

    try:
        amount = Decimal(cleaned)
    except InvalidOperation:
        return None

    return -amount if negative else amount


def extract_statement_range(
    text: str,
    date_preference: str = "day_first",
) -> tuple[datetime, datetime] | None:
    date_token = (
        r"(?:"
        r"\d{1,2}[./\-]\d{1,2}[./\-]\d{2,4}"
        r"|"
        r"\d{4}[./\-]\d{1,2}[./\-]\d{1,2}"
        r"|"
        r"\d{1,2}\s+[A-Za-z]{3,9}\s+\d{2,4}"
        r"|"
        r"[A-Za-z]{3,9}\s+\d{1,2},?\s+\d{2,4}"
        r")"
    )

    match = None
    for pattern in (
        rf"\b({date_token})\s+To\s+({date_token})\b",
        rf"\b({date_token})\s*-\s*({date_token})\b",
    ):
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            break

    if not match:
        return None

    parsed_dates: list[datetime] = []
    for raw_value in match.groups():
        parsed_date = parse_flexible_date(raw_value, date_preference)
        if parsed_date is None:
            return None
        parsed_dates.append(parsed_date)

    return parsed_dates[0], parsed_dates[1]


def extract_statement_balances(text: str) -> tuple[Decimal | None, Decimal | None]:
    brought_match = re.search(
        r"\bBrought\s+Forward\s+([0-9,]+\.\d{2}-?)\b",
        text,
        flags=re.IGNORECASE,
    )
    carried_match = re.search(
        r"\bCarried\s+Forward\s+([0-9,]+\.\d{2}-?)\b",
        text,
        flags=re.IGNORECASE,
    )

    opening_balance = parse_money_value(brought_match.group(1)) if brought_match else None
    closing_balance = parse_money_value(carried_match.group(1)) if carried_match else None
    return opening_balance, closing_balance


def extract_statement_metadata(
    pdf_file: Path,
) -> tuple[Path, datetime, datetime, Decimal | None, Decimal | None] | None:
    try:
        statement_text = extract_text_with_password(pdf_file, None)
    except Exception:
        return None

    statement_range = extract_statement_range(statement_text)
    if not statement_range:
        return None

    opening_balance, closing_balance = extract_statement_balances(statement_text)
    start_date, end_date = statement_range
    return pdf_file, start_date, end_date, opening_balance, closing_balance


def find_missing_statement_dates() -> list[str]:
    warnings: list[str] = []
    if not OUTPUT_FOLDER.exists():
        return warnings

    for account_dir in sorted(path for path in OUTPUT_FOLDER.iterdir() if path.is_dir()):
        for period_dir in sorted(path for path in account_dir.iterdir() if path.is_dir()):
            statement_entries: list[tuple[Path, datetime, datetime, Decimal | None, Decimal | None]] = []
            for pdf_file in sorted(period_dir.glob("*.pdf")):
                if pdf_file.name.lower() == "combined_statements.pdf":
                    continue

                metadata = extract_statement_metadata(pdf_file)
                if metadata is not None:
                    statement_entries.append(metadata)

            statement_entries.sort(key=lambda entry: (entry[2], entry[1], entry[0].name.lower()))
            missing_dates: list[str] = []
            for previous_entry, current_entry in zip(statement_entries, statement_entries[1:]):
                _previous_file, _previous_start, previous_end, _previous_opening, previous_closing = previous_entry
                _current_file, current_start, _current_end, current_opening, _current_closing = current_entry

                gap_days = (current_start - previous_end).days - 1
                if gap_days <= 0:
                    continue

                if (
                    previous_closing is not None
                    and current_opening is not None
                    and previous_closing == current_opening
                ):
                    continue

                missing_day = previous_end + timedelta(days=1)
                while missing_day < current_start:
                    missing_dates.append(missing_day.strftime("%Y-%m-%d"))
                    missing_day = missing_day + timedelta(days=1)

            if missing_dates:
                warnings.append(
                    f"Possible missing transaction dates for {account_dir.name} / {period_dir.name}: "
                    + ", ".join(sorted(set(missing_dates)))
                )

    return warnings


def export_pdf(source_path: Path, destination_path: Path, used_password: str | None) -> str:
    destination_dir = destination_path.parent
    destination_dir.mkdir(parents=True, exist_ok=True)

    source_hash = file_sha256(source_path)
    existing_match = folder_pdf_hashes(destination_dir).get(source_hash)
    if existing_match is not None:
        return f"Skipped duplicate content (matches {existing_match.name})"

    if destination_path.exists():
        return "Skipped existing"

    if used_password:
        write_unprotected_pdf(source_path, destination_path, used_password)
        if not COPY_INSTEAD_OF_MOVE:
            source_path.unlink()
        return "Decrypted copy saved" if COPY_INSTEAD_OF_MOVE else "Decrypted file moved"

    if COPY_INSTEAD_OF_MOVE:
        shutil.copy2(source_path, destination_path)
        return "Copied"

    shutil.move(str(source_path), str(destination_path))
    return "Moved"


def save_mt940_attachment(
    filename: str,
    payload: bytes,
    fallback_file_date: datetime,
) -> tuple[str, bool]:
    statement_text = decode_text_payload(payload)
    found_accounts = find_mt940_accounts(statement_text)
    account_folder = classify_account(found_accounts, allow_unlisted=True)
    doc_date = extract_mt940_date(statement_text, fallback_file_date)

    period_folder = doc_date.strftime("%m-%Y")
    day_prefix = doc_date.strftime("%Y-%m-%d")
    destination_dir = OUTPUT_FOLDER / safe_filename(account_folder) / period_folder / MT940_SUBFOLDER_NAME
    destination_name = f"{day_prefix}_{filename}"
    destination_path, action = save_payload_file(destination_dir, destination_name, payload)
    accounts_text = ", ".join(found_accounts) if found_accounts else "None"

    status_text = "Saved MT940 DAT" if action == "Saved" else action
    return (
        f"{status_text}: {filename} -> {destination_path} | Account: {account_folder} | Accounts found: {accounts_text}",
        action == "Saved",
    )


def extract_supported_attachments_from_eml(eml_path: Path) -> list[tuple[str, bytes]]:
    with eml_path.open("rb") as file_handle:
        message = BytesParser(policy=policy.default).parse(file_handle)

    supported_attachments = iter_supported_message_parts(message)
    for part in message.walk():
        payload = part.get_payload(decode=True)
        if not payload:
            continue

        filename = (part.get_filename() or "").lower()
        if part.get_content_type() == "application/pkcs7-mime" or filename.endswith(".p7m"):
            supported_attachments.extend(extract_supported_attachments_from_smime_payload(payload))

    return supported_attachments


def extract_pdfs_from_eml(eml_path: Path, destination_folder: Path) -> list[Path]:
    extracted_files: list[Path] = []
    for filename, payload in extract_supported_attachments_from_eml(eml_path):
        if not filename.lower().endswith(".pdf"):
            continue
        output_path, was_created = save_attachment_bytes(destination_folder, filename, payload)
        if was_created and output_path is not None:
            extracted_files.append(output_path)

    return extracted_files


def extract_supported_attachments_from_msg(msg_path: Path) -> list[tuple[str, bytes]]:
    if extract_msg is None:
        return extract_supported_attachments_from_msg_with_outlook(msg_path)

    supported_attachments: list[tuple[str, bytes]] = []
    message = extract_msg.openMsg(str(msg_path))

    try:
        for attachment in message.attachments:
            filename = getattr(attachment, "longFilename", None) or getattr(
                attachment,
                "shortFilename",
                None,
            ) or ""
            payload = getattr(attachment, "data", None)
            if not isinstance(payload, (bytes, bytearray)):
                continue

            lower_filename = filename.lower()
            if lower_filename.endswith(".pdf"):
                supported_attachments.append((filename or "attachment.pdf", bytes(payload)))
            elif is_supported_dat_attachment(filename, "", bytes(payload)):
                output_name = filename if lower_filename.endswith(".dat") else f"{Path(filename).stem or 'attachment'}.dat"
                supported_attachments.append((output_name, bytes(payload)))

        if str(getattr(message, "classType", "")).upper() == "IPM.NOTE.SMIME":
            for attachment in getattr(message, "rawAttachments", []):
                filename = str(getattr(attachment, "longFilename", "") or getattr(attachment, "name", ""))
                mimetype = str(getattr(attachment, "mimetype", "") or "")
                payload = getattr(attachment, "data", None)
                if not isinstance(payload, (bytes, bytearray)):
                    continue
                if mimetype == "application/pkcs7-mime" or filename.lower().endswith(".p7m"):
                    supported_attachments.extend(
                        extract_supported_attachments_from_smime_payload(bytes(payload))
                    )
    finally:
        close_method = getattr(message, "close", None)
        if callable(close_method):
            close_method()

    return supported_attachments


def extract_supported_attachments_from_msg_with_outlook(msg_path: Path) -> list[tuple[str, bytes]]:
    if win32com is None:
        raise RuntimeError(
            "Outlook .msg support needs pywin32. Run run_pdf_classifier.bat again "
            "so it can install the default requirements."
        )

    supported_attachments: list[tuple[str, bytes]] = []
    com_initialized = False
    message = None

    if pythoncom is not None:
        try:
            pythoncom.CoInitialize()
            com_initialized = True
        except Exception:
            com_initialized = False

    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        message = outlook.Session.OpenSharedItem(str(msg_path.resolve()))

        with tempfile.TemporaryDirectory(prefix="pdf_classifier_msg_") as temp_dir:
            temp_root = Path(temp_dir)
            attachments = getattr(message, "Attachments", None)
            attachment_count = int(getattr(attachments, "Count", 0) or 0)

            for index in range(1, attachment_count + 1):
                attachment = attachments.Item(index)
                filename = (
                    str(getattr(attachment, "FileName", "") or "")
                    or str(getattr(attachment, "DisplayName", "") or "")
                    or f"attachment_{index}"
                )
                temp_path = unique_path(temp_root / safe_filename(filename))
                attachment.SaveAsFile(str(temp_path))
                payload = temp_path.read_bytes()
                lower_filename = filename.lower()

                if lower_filename.endswith(".pdf"):
                    supported_attachments.append((filename or "attachment.pdf", payload))
                elif is_supported_dat_attachment(filename, "", payload):
                    output_name = filename if lower_filename.endswith(".dat") else f"{Path(filename).stem or 'attachment'}.dat"
                    supported_attachments.append((output_name, payload))
                elif lower_filename.endswith(".p7m"):
                    supported_attachments.extend(extract_supported_attachments_from_smime_payload(payload))
    finally:
        if message is not None:
            try:
                message.Close(1)
            except Exception:
                pass
        if com_initialized and pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    return supported_attachments


def extract_pdfs_from_msg(msg_path: Path, destination_folder: Path) -> list[Path]:
    extracted_files: list[Path] = []
    for filename, payload in extract_supported_attachments_from_msg(msg_path):
        if not filename.lower().endswith(".pdf"):
            continue
        output_path, was_created = save_attachment_bytes(destination_folder, filename, payload)
        if was_created and output_path is not None:
            extracted_files.append(output_path)

    return extracted_files


def import_mail_attachments() -> tuple[list[str], list[Path], int, int]:
    MAIL_SOURCE_FOLDER.mkdir(parents=True, exist_ok=True)
    SOURCE_FOLDER.mkdir(parents=True, exist_ok=True)

    results: list[str] = []
    processed_mail_files: list[Path] = []
    eml_files = sorted(MAIL_SOURCE_FOLDER.glob("*.eml"))
    msg_files = sorted(MAIL_SOURCE_FOLDER.glob("*.msg"))

    if not eml_files and not msg_files:
        results.append("No saved mail files found in input_mails.")
        return results, processed_mail_files, 0, 0

    imported_pdf_count = 0
    imported_mt940_count = 0

    def import_supported_attachments(
        mail_path: Path,
        attachments: list[tuple[str, bytes]],
    ) -> None:
        nonlocal imported_pdf_count, imported_mt940_count

        if not attachments:
            results.append(f"No PDF or MT940 DAT attachments found in {mail_path.name}")
            return

        fallback_date = datetime.fromtimestamp(mail_path.stat().st_mtime)
        for filename, payload in attachments:
            lower_filename = filename.lower()
            if lower_filename.endswith(".pdf"):
                output_path, was_created = save_attachment_bytes(SOURCE_FOLDER, filename, payload)
                if output_path is None:
                    continue
                if was_created:
                    results.append(f"Saved PDF from {mail_path.name} -> {output_path.name}")
                    imported_pdf_count += 1
                else:
                    results.append(f"Skipped duplicate PDF from {mail_path.name} -> {output_path.name}")
                continue

            if lower_filename.endswith(".dat"):
                status_line, was_saved = save_mt940_attachment(filename, payload, fallback_date)
                results.append(f"{status_line} | Mail: {mail_path.name}")
                if was_saved:
                    imported_mt940_count += 1

    for eml_file in eml_files:
        try:
            import_supported_attachments(eml_file, extract_supported_attachments_from_eml(eml_file))
            processed_mail_files.append(eml_file)
        except Exception as error:
            results.append(f"Could not read {eml_file.name}: {error}")

    for msg_file in msg_files:
        try:
            import_supported_attachments(msg_file, extract_supported_attachments_from_msg(msg_file))
            processed_mail_files.append(msg_file)
        except Exception as error:
            results.append(f"Could not read {msg_file.name}: {error}")

    if imported_pdf_count == 0 and imported_mt940_count == 0 and (eml_files or msg_files):
        results.append("No PDF or MT940 DAT attachments were imported from saved emails.")

    return results, processed_mail_files, imported_pdf_count, imported_mt940_count


def process_pdf(pdf_path: Path) -> None:
    print(f"Processing: {pdf_path.name}")

    text, used_password = extract_text_from_pdf(pdf_path)
    normalized_text = normalize_text(text)
    account_search_text = text or normalized_text

    found_accounts = find_possible_accounts(account_search_text)
    account_folder = classify_account(found_accounts, text=account_search_text)

    fallback_date = datetime.fromtimestamp(pdf_path.stat().st_mtime)
    doc_date = extract_date(text or normalized_text, fallback_date)

    period_folder = doc_date.strftime("%m-%Y")
    day_prefix = doc_date.strftime("%Y-%m-%d")

    # Create the account folder only when the matched account number is needed.
    destination_dir = OUTPUT_FOLDER / account_folder / period_folder
    destination_dir.mkdir(parents=True, exist_ok=True)

    new_name = f"{day_prefix}_{safe_filename(pdf_path.name)}"
    destination_file = destination_dir / new_name

    action = export_pdf(pdf_path, destination_file, used_password)

    print(f"{action} -> {destination_file}")
    print(f"Accounts found: {found_accounts if found_accounts else 'None'}")
    if used_password:
        print(f"Unlocked with password: {used_password}")
    elif not text:
        print("Could not unlock or read PDF text.")
    print("-" * 60)


def process_pdf_with_status(pdf_path: Path) -> str:
    text, used_password = extract_text_from_pdf(pdf_path)
    normalized_text = normalize_text(text)
    account_search_text = text or normalized_text

    found_accounts = find_possible_accounts(account_search_text)
    account_folder = classify_account(found_accounts, text=account_search_text)

    fallback_date = datetime.fromtimestamp(pdf_path.stat().st_mtime)
    doc_date = extract_date(text or normalized_text, fallback_date)

    period_folder = doc_date.strftime("%m-%Y")
    day_prefix = doc_date.strftime("%Y-%m-%d")

    destination_dir = OUTPUT_FOLDER / account_folder / period_folder
    destination_dir.mkdir(parents=True, exist_ok=True)

    new_name = f"{day_prefix}_{safe_filename(pdf_path.name)}"
    destination_file = destination_dir / new_name

    action = export_pdf(pdf_path, destination_file, used_password)

    accounts_text = ", ".join(found_accounts) if found_accounts else "None"
    password_text = used_password if used_password else "Not needed or not found"
    return (
        f"{action}: {pdf_path.name} -> {destination_file} | "
        f"Accounts: {accounts_text} | Password: {password_text}"
    )


def detect_pdf_target(pdf_path: Path) -> tuple[str, str, str] | None:
    text, _used_password = extract_text_from_pdf(pdf_path)
    normalized_text = normalize_text(text)
    account_search_text = text or normalized_text
    found_accounts = find_possible_accounts(account_search_text)
    account_folder = classify_account(found_accounts, text=account_search_text)
    fallback_date = datetime.fromtimestamp(pdf_path.stat().st_mtime)
    doc_date = extract_date(text or normalized_text, fallback_date)
    return account_folder, doc_date.strftime("%m-%Y"), doc_date.strftime("%Y-%m-%d")


def normalize_period_input(period_text: str) -> str:
    cleaned = period_text.strip().replace("/", "-").replace("_", "-")
    if not re.fullmatch(r"\d{2}-\d{4}", cleaned):
        raise ValueError("Month must be in MM-YYYY format, for example 04-2026.")

    month_value = int(cleaned[:2])
    if month_value < 1 or month_value > 12:
        raise ValueError("Month must be between 01 and 12.")

    return cleaned


def detect_pdf_period(pdf_path: Path) -> str | None:
    detected_target = detect_pdf_target(pdf_path)
    if detected_target is None:
        return None
    return detected_target[1]


def detect_periods_from_pdfs(pdf_files: list[Path]) -> list[str]:
    detected_periods: set[str] = set()
    for pdf_file in pdf_files:
        try:
            period_value = detect_pdf_period(pdf_file)
        except Exception:
            continue
        if period_value:
            detected_periods.add(period_value)

    return sorted(detected_periods)


def detect_targets_from_pdfs(pdf_files: list[Path]) -> list[tuple[str, str, str]]:
    detected_targets: set[tuple[str, str, str]] = set()
    for pdf_file in pdf_files:
        try:
            detected_target = detect_pdf_target(pdf_file)
        except Exception:
            continue
        if detected_target:
            detected_targets.add(detected_target)

    return sorted(detected_targets)


def remove_empty_parent_dirs(path: Path, stop_at: Path) -> None:
    current_path = path
    while current_path.exists() and current_path != stop_at:
        if not is_directory_empty(current_path):
            break

        try:
            remove_directory(current_path)
        except OSError:
            break
        current_path = current_path.parent


def clear_output_month(period_text: str, account_names: list[str] | None = None) -> list[str]:
    normalized_period = normalize_period_input(period_text)
    results: list[str] = []
    selected_accounts = set(account_names or [])

    if OUTPUT_FOLDER.exists():
        for account_dir in sorted(path for path in OUTPUT_FOLDER.iterdir() if path.is_dir()):
            if selected_accounts and account_dir.name not in selected_accounts:
                continue
            target_dir = account_dir / normalized_period
            if target_dir.exists():
                try:
                    delete_tree(target_dir)
                    results.append(f"Cleared output folder -> {target_dir}")
                    remove_empty_parent_dirs(account_dir, OUTPUT_FOLDER)
                except PermissionError as error:
                    blocked_path = getattr(error, "filename", None) or str(target_dir)
                    results.append(
                        f"Could not fully clear output folder because a file is in use -> {blocked_path}"
                    )

    combined_period_dir = combined_period_folder(normalized_period)
    if combined_period_dir.exists():
        combined_names = None
        if selected_accounts:
            combined_names = {
                build_combined_output_name(account_name, normalized_period)
                for account_name in selected_accounts
            }

        for combined_file in sorted(combined_period_dir.glob("*.pdf")):
            if combined_names and combined_file.name not in combined_names:
                continue
            try:
                delete_file(combined_file)
                results.append(f"Removed combined PDF -> {combined_file}")
            except PermissionError:
                results.append(f"Could not remove combined PDF because it is in use -> {combined_file}")
        if combined_period_dir.exists():
            remove_empty_parent_dirs(combined_period_dir, COMBINED_FOLDER)

    if not results:
        results.append(f"No output folders found for {normalized_period}.")

    return results


def statement_pdfs_in_folder(folder_path: Path) -> list[Path]:
    return sorted(
        path
        for path in folder_path.glob("*.pdf")
        if path.name.lower() != "combined_statements.pdf"
    )


def refresh_combined_pdf_for_target(account_name: str, period_text: str) -> list[str]:
    normalized_period = normalize_period_input(period_text)
    results: list[str] = []
    target_dir = OUTPUT_FOLDER / account_name / normalized_period
    combined_source = target_dir / "combined_statements.pdf"
    combined_destination_dir = combined_period_folder(normalized_period)
    combined_destination = combined_destination_dir / build_combined_output_name(
        account_name,
        normalized_period,
    )

    if target_dir.exists() and statement_pdfs_in_folder(target_dir):
        merge_result = merge_pdfs_in_folder(target_dir)
        if merge_result:
            results.append(merge_result)
        if combined_source.exists():
            combined_destination_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(combined_source, combined_destination)
            results.append(f"Saved combined PDF -> {combined_destination}")
        return results

    if combined_source.exists():
        try:
            delete_file(combined_source)
            results.append(f"Removed combined PDF -> {combined_source}")
        except PermissionError:
            results.append(f"Could not remove combined PDF because it is in use -> {combined_source}")

    if target_dir.exists():
        remove_empty_parent_dirs(target_dir, OUTPUT_FOLDER)

    if combined_destination.exists():
        try:
            delete_file(combined_destination)
            results.append(f"Removed combined PDF -> {combined_destination}")
        except PermissionError:
            results.append(f"Could not remove combined PDF because it is in use -> {combined_destination}")

    if combined_destination_dir.exists():
        remove_empty_parent_dirs(combined_destination_dir, COMBINED_FOLDER)

    return results


def clear_output_statements(detected_targets: list[tuple[str, str, str]]) -> list[str]:
    results: list[str] = []
    affected_targets: set[tuple[str, str]] = set()
    unique_targets = sorted(set(detected_targets))

    for account_name, period_value, day_prefix in unique_targets:
        normalized_period = normalize_period_input(period_value)
        target_dir = OUTPUT_FOLDER / account_name / normalized_period
        matched_count = 0
        removed_count = 0

        if target_dir.exists():
            for statement_file in sorted(target_dir.glob(f"{day_prefix}_*.pdf")):
                if statement_file.name.lower() == "combined_statements.pdf":
                    continue
                matched_count += 1
                try:
                    delete_file(statement_file)
                    removed_count += 1
                    results.append(f"Removed existing statement -> {statement_file}")
                except PermissionError:
                    results.append(
                        f"Could not remove existing statement because it is in use -> {statement_file}"
                    )

        if removed_count > 0:
            affected_targets.add((account_name, normalized_period))
        elif matched_count == 0:
            results.append(
                f"No existing statement found to replace for {account_name} / {normalized_period} / {day_prefix}."
            )

    for account_name, period_value in sorted(affected_targets):
        results.extend(refresh_combined_pdf_for_target(account_name, period_value))

    if not results:
        results.append("No matching output statements found to clear.")

    return results


def clear_detected_output_months(pdf_files: list[Path]) -> list[str]:
    detected_targets = detect_targets_from_pdfs(pdf_files)
    if not detected_targets:
        return ["Could not detect any statement targets to clear from the current PDFs."]

    target_summary = ", ".join(
        f"{account_name}/{period_value}/{day_prefix}"
        for account_name, period_value, day_prefix in detected_targets
    )
    return [f"Detected cleanup targets: {target_summary}"] + clear_output_statements(detected_targets)



def delete_file(path: Path) -> None:
    last_error: OSError | None = None
    for _attempt in range(5):
        try:
            path.unlink()
            return
        except PermissionError as error:
            last_error = error
            make_writable(path)
            time.sleep(0.2)

    if last_error is not None:
        raise last_error


def is_directory_empty(path: Path) -> bool:
    try:
        next(path.iterdir())
        return False
    except StopIteration:
        return True
    except OSError:
        return False


def remove_directory(path: Path) -> None:
    last_error: OSError | None = None
    for _attempt in range(5):
        try:
            path.rmdir()
            return
        except PermissionError as error:
            last_error = error
            make_writable(path)
            time.sleep(0.2)
        except OSError as error:
            last_error = error
            time.sleep(0.2)

    if last_error is not None:
        raise last_error


def delete_tree(folder_path: Path) -> None:
    for child in sorted(folder_path.iterdir(), reverse=True):
        if child.is_dir():
            delete_tree(child)
        else:
            delete_file(child)

    try:
        remove_directory(folder_path)
    except OSError:
        if folder_path.exists() and not is_directory_empty(folder_path):
            raise


def clear_input_pdfs(pdf_files: list[Path]) -> list[str]:
    results: list[str] = []
    for pdf_file in pdf_files:
        if not pdf_file.exists():
            continue
        try:
            delete_file(pdf_file)
            results.append(f"Removed input PDF -> {pdf_file}")
        except PermissionError:
            results.append(f"Could not remove input PDF because it is in use -> {pdf_file}")
    return results


def clear_mail_files(mail_files: list[Path]) -> list[str]:
    results: list[str] = []
    for mail_file in mail_files:
        if not mail_file.exists():
            continue
        try:
            delete_file(mail_file)
            results.append(f"Removed mail file -> {mail_file}")
        except PermissionError:
            results.append(f"Could not remove mail file because it is in use -> {mail_file}")
    return results


def run_classification(auto_clear_detected_periods: bool = False) -> list[str]:
    OUTPUT_FOLDER.mkdir(parents=True, exist_ok=True)
    global KNOWN_ACCOUNT_RECORDS, KNOWN_ACCOUNTS
    KNOWN_ACCOUNT_RECORDS = load_account_records(ACCOUNT_NUMBERS_FILE)
    KNOWN_ACCOUNTS = [str(record["account_number"]) for record in KNOWN_ACCOUNT_RECORDS]

    results, processed_mail_files, _imported_pdf_count, imported_mt940_count = import_mail_attachments()
    pdf_files = sorted(SOURCE_FOLDER.glob("*.pdf"))
    if not pdf_files:
        if CLEAR_MAIL_AFTER_RUN:
            results.extend(clear_mail_files(processed_mail_files))

        if imported_mt940_count > 0:
            results.append("Done.")
            return results

        if not results:
            return ["No PDF or MT940 DAT files found."]
        results.append("No PDF or MT940 DAT files found.")
        return results

    if auto_clear_detected_periods:
        results.extend(clear_detected_output_months(pdf_files))

    seen_hashes: dict[str, Path] = {}
    processed_input_files: list[Path] = []
    for pdf_file in pdf_files:
        pdf_hash = file_sha256(pdf_file)
        original_file = seen_hashes.get(pdf_hash)
        if original_file is not None:
            results.append(f"Skipped duplicate source PDF: {pdf_file.name} matches {original_file.name}")
            processed_input_files.append(pdf_file)
            continue
        seen_hashes[pdf_hash] = pdf_file
        results.append(process_pdf_with_status(pdf_file))
        processed_input_files.append(pdf_file)

    results.extend(combine_classified_pdfs())
    results.extend(find_missing_statement_dates())
    if CLEAR_INPUT_AFTER_RUN:
        results.extend(clear_input_pdfs(processed_input_files))
    if CLEAR_MAIL_AFTER_RUN:
        results.extend(clear_mail_files(processed_mail_files))
    results.append("Done.")
    return results


def convert_combined_statements(converter_output_folder: Path | None = None) -> list[str]:
    try:
        from statement_converter import BUILD_STATEMENT
    except Exception as error:
        return [f"Could not load statement converter: {error}"]

    combined_files = sorted(OUTPUT_FOLDER.rglob("combined_statements.pdf"))
    if not combined_files:
        return ["No combined_statements.pdf files were found in the classified output folder."]

    default_output_description = "same folder as each combined_statements.pdf, under statement_converter"
    results = [
        f"Converting {len(combined_files)} combined PDF(s).",
        f"Converter output folder: {converter_output_folder or default_output_description}",
    ]
    success_count = 0
    failed_count = 0

    for combined_file in combined_files:
        output_root = converter_output_folder or (combined_file.parent / "statement_converter")
        conversion = BUILD_STATEMENT(combined_file, output_root=output_root, write_debug=True)
        relative_name = str(combined_file.relative_to(OUTPUT_FOLDER))
        if conversion.success:
            success_count += 1
            results.append(
                f"Converted: {relative_name} -> DAT: {conversion.mt940_path} | XLSX: {conversion.xlsx_path}"
            )
        else:
            failed_count += 1
            results.append(f"Failed: {relative_name} -> {conversion.error or 'Unknown converter error'}")

    results.append(f"Conversion finished. Success: {success_count}. Failed: {failed_count}.")
    return results


def launch_ui() -> None:
    root = tk.Tk()
    root.title("PDF Classifier")
    root.geometry("940x700")
    root.minsize(820, 620)
    root.configure(bg="#f4f1ea")

    title = tk.Label(
        root,
        text="PDF Classifier",
        font=("Segoe UI Semibold", 20),
        bg="#f4f1ea",
        fg="#1f2a44",
    )
    title.pack(pady=(18, 6))

    subtitle = tk.Label(
        root,
        text="Classify PDFs and MT940 DAT files by account number, with monthly PDF combines",
        font=("Segoe UI", 10),
        bg="#f4f1ea",
        fg="#5b6474",
        justify="center",
    )
    subtitle.pack(pady=(0, 14))

    def sync_header_wrap(_event=None) -> None:
        subtitle.config(wraplength=max(480, root.winfo_width() - 80))

    root.bind("<Configure>", sync_header_wrap)

    is_running = False
    animation_job: str | None = None
    worker_error: Exception | None = None
    worker_results: list[str] | None = None
    action_buttons: list[tk.Widget] = []
    loading_frames = [
        "Processing   ",
        "Processing.  ",
        "Processing.. ",
        "Processing...",
    ]
    loading_index = 0
    path_settings_window: tk.Toplevel | None = None

    path_vars = {
        "mail": tk.StringVar(value=str(MAIL_SOURCE_FOLDER)),
        "input": tk.StringVar(value=str(SOURCE_FOLDER)),
        "output": tk.StringVar(value=str(OUTPUT_FOLDER)),
        "combined": tk.StringVar(value=str(COMBINED_FOLDER)),
        "accounts": tk.StringVar(value=str(ACCOUNT_NUMBERS_FILE)),
    }
    path_rows = [
        ("Mail folder", "mail", False),
        ("Input folder", "input", False),
        ("Output folder", "output", False),
        ("Combined folder", "combined", False),
        ("Account file", "accounts", True),
    ]
    path_summary_var = tk.StringVar()

    def path_display_name(path_key: str) -> str:
        raw_value = path_vars[path_key].get().strip()
        if not raw_value:
            return "(not set)"
        path_value = Path(raw_value)
        return path_value.name or str(path_value)

    def refresh_path_summary() -> None:
        path_summary_var.set(
            " | ".join(
                [
                    f"Mail: {path_display_name('mail')}",
                    f"Input: {path_display_name('input')}",
                    f"Output: {path_display_name('output')}",
                    f"Combined: {path_display_name('combined')}",
                    f"Accounts: {path_display_name('accounts')}",
                ]
            )
        )

    def update_runtime_paths() -> None:
        global PATH_SETTINGS, MAIL_SOURCE_FOLDER, SOURCE_FOLDER, OUTPUT_FOLDER, COMBINED_FOLDER, ACCOUNT_NUMBERS_FILE
        PATH_SETTINGS = {
            "mail": Path(path_vars["mail"].get().strip()).expanduser(),
            "input": Path(path_vars["input"].get().strip()).expanduser(),
            "output": Path(path_vars["output"].get().strip()).expanduser(),
            "combined": Path(path_vars["combined"].get().strip()).expanduser(),
            "accounts": Path(path_vars["accounts"].get().strip()).expanduser(),
        }
        MAIL_SOURCE_FOLDER = PATH_SETTINGS["mail"]
        SOURCE_FOLDER = PATH_SETTINGS["input"]
        OUTPUT_FOLDER = PATH_SETTINGS["output"]
        COMBINED_FOLDER = PATH_SETTINGS["combined"]
        ACCOUNT_NUMBERS_FILE = PATH_SETTINGS["accounts"]
        refresh_path_summary()

    def browse_for_path(
        path_key: str,
        select_file: bool,
        target_vars: dict[str, tk.StringVar],
    ) -> None:
        initial_path = target_vars[path_key].get().strip() or str(BASE_FOLDER)
        if select_file:
            selected_path = filedialog.asksaveasfilename(
                title="Select Account Numbers File",
                initialfile=Path(initial_path).name if initial_path else "account_numbers.txt",
                initialdir=str(Path(initial_path).parent if initial_path else BASE_FOLDER),
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            )
        else:
            selected_path = filedialog.askdirectory(
                title="Select Folder",
                initialdir=initial_path,
            )

        if selected_path:
            target_vars[path_key].set(selected_path)

    def open_path_settings() -> None:
        nonlocal path_settings_window
        if path_settings_window is not None and path_settings_window.winfo_exists():
            path_settings_window.lift()
            path_settings_window.focus_force()
            return

        dialog = tk.Toplevel(root)
        path_settings_window = dialog
        dialog.title("Path Settings")
        dialog.geometry("860x330")
        dialog.minsize(720, 280)
        dialog.configure(bg="#f4f1ea")
        dialog.transient(root)

        dialog_vars = {
            key: tk.StringVar(value=path_vars[key].get())
            for key in path_vars
        }

        def close_path_settings() -> None:
            nonlocal path_settings_window
            if path_settings_window is not None and path_settings_window.winfo_exists():
                path_settings_window.destroy()
            path_settings_window = None

        def save_path_settings() -> None:
            for key, value_var in dialog_vars.items():
                path_vars[key].set(value_var.get().strip())
            update_runtime_paths()
            try:
                write_path_settings(PATH_SETTINGS)
            except OSError as error:
                messagebox.showerror(
                    "PDF Classifier",
                    f"Could not save path settings to {SETTINGS_FILE}.\n{error}",
                )
                return
            refresh_account_list()
            close_path_settings()

        dialog.protocol("WM_DELETE_WINDOW", close_path_settings)
        dialog.grid_columnconfigure(0, weight=1)
        dialog.grid_rowconfigure(1, weight=1)

        header_frame = tk.Frame(dialog, bg="#f4f1ea")
        header_frame.grid(row=0, column=0, sticky="ew", padx=20, pady=(18, 6))
        header_frame.grid_columnconfigure(1, weight=1)

        dialog_title = tk.Label(
            header_frame,
            text="Path Settings",
            font=("Segoe UI Semibold", 14),
            bg="#f4f1ea",
            fg="#1f2a44",
        )
        dialog_title.grid(row=0, column=0, sticky="w")

        dialog_hint = tk.Label(
            header_frame,
            text="Keep the main window clean by editing folders here.",
            font=("Segoe UI", 10),
            bg="#f4f1ea",
            fg="#5b6474",
        )
        dialog_hint.grid(row=0, column=1, sticky="e")

        form_frame = tk.Frame(dialog, bg="#fffdf8", bd=1, relief="solid")
        form_frame.grid(row=1, column=0, sticky="nsew", padx=20, pady=(0, 12))
        form_frame.grid_columnconfigure(1, weight=1)

        for row_index, (label_text, path_key, select_file) in enumerate(path_rows):
            label = tk.Label(
                form_frame,
                text=label_text,
                anchor="w",
                font=("Segoe UI", 10),
                bg="#fffdf8",
                fg="#24324a",
            )
            label.grid(row=row_index, column=0, sticky="w", padx=(12, 8), pady=8)

            entry = tk.Entry(
                form_frame,
                textvariable=dialog_vars[path_key],
                font=("Consolas", 10),
                bg="#fffdf8",
                fg="#1f2a44",
                insertbackground="#1f2a44",
                relief="solid",
                bd=1,
            )
            entry.grid(row=row_index, column=1, sticky="ew", padx=(0, 8), pady=8, ipady=5)

            browse_button = tk.Button(
                form_frame,
                text="Browse",
                command=lambda current_key=path_key, is_file=select_file: browse_for_path(
                    current_key,
                    is_file,
                    dialog_vars,
                ),
                font=("Segoe UI", 9),
                bg="#d6e6f2",
                fg="#1f2a44",
                activebackground="#bdd5e7",
                activeforeground="#1f2a44",
                padx=10,
                pady=5,
                bd=0,
            )
            browse_button.grid(row=row_index, column=2, sticky="ew", padx=(0, 12), pady=8)

        footer = tk.Frame(dialog, bg="#f4f1ea")
        footer.grid(row=2, column=0, sticky="ew", padx=20, pady=(0, 18))
        footer.grid_columnconfigure(0, weight=1)

        cancel_button = tk.Button(
            footer,
            text="Cancel",
            command=close_path_settings,
            font=("Segoe UI", 10),
            bg="#d6e6f2",
            fg="#1f2a44",
            activebackground="#bdd5e7",
            activeforeground="#1f2a44",
            padx=12,
            pady=7,
            bd=0,
        )
        cancel_button.grid(row=0, column=1, padx=(8, 0))

        save_button = tk.Button(
            footer,
            text="Save Paths",
            command=save_path_settings,
            font=("Segoe UI Semibold", 10),
            bg="#22577a",
            fg="white",
            activebackground="#16384f",
            activeforeground="white",
            padx=12,
            pady=7,
            bd=0,
        )
        save_button.grid(row=0, column=2, padx=(8, 0))

        dialog.grab_set()
        dialog.focus_force()
        dialog.bind("<Escape>", lambda _event: close_path_settings())

    path_bar = tk.Frame(root, bg="#fffdf8", bd=1, relief="solid")
    path_bar.pack(fill="x", padx=20)
    path_bar.grid_columnconfigure(1, weight=1)

    path_bar_label = tk.Label(
        path_bar,
        text="Paths",
        font=("Segoe UI Semibold", 10),
        bg="#fffdf8",
        fg="#1f2a44",
    )
    path_bar_label.grid(row=0, column=0, sticky="w", padx=(12, 10), pady=10)

    path_summary_label = tk.Label(
        path_bar,
        textvariable=path_summary_var,
        font=("Segoe UI", 9),
        bg="#fffdf8",
        fg="#5b6474",
        anchor="w",
        justify="left",
    )
    path_summary_label.grid(row=0, column=1, sticky="ew", padx=(0, 10), pady=10)

    def sync_path_summary_wrap(_event=None) -> None:
        available_width = max(240, path_bar.winfo_width() - 220)
        path_summary_label.config(wraplength=available_width)

    path_settings_button = tk.Button(
        path_bar,
        text="Manage Paths",
        command=open_path_settings,
        font=("Segoe UI", 9),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=10,
        pady=6,
        bd=0,
    )
    path_settings_button.grid(row=0, column=2, sticky="e", padx=(0, 12), pady=8)
    action_buttons.append(path_settings_button)
    path_bar.bind("<Configure>", sync_path_summary_wrap)

    accounts_frame = tk.Frame(root, bg="#f4f1ea")
    accounts_frame.pack(fill="x", padx=20, pady=(14, 0))

    account_label = tk.Label(
        accounts_frame,
        text="Store Account Number",
        font=("Segoe UI Semibold", 11),
        bg="#f4f1ea",
        fg="#1f2a44",
    )
    account_label.pack(anchor="w", pady=(0, 8))

    account_entry_frame = tk.Frame(accounts_frame, bg="#f4f1ea")
    account_entry_frame.pack(fill="x")

    account_number_var = tk.StringVar()
    account_entry = tk.Entry(
        account_entry_frame,
        textvariable=account_number_var,
        font=("Consolas", 11),
        bg="#fffdf8",
        fg="#1f2a44",
        insertbackground="#1f2a44",
        relief="solid",
        bd=1,
    )
    account_entry.pack(side="left", fill="x", expand=True, ipady=6)

    password_row = tk.Frame(accounts_frame, bg="#f4f1ea")
    password_row.pack(fill="x", pady=(10, 0))

    password_protected_var = tk.BooleanVar(value=False)
    password_check = tk.Checkbutton(
        password_row,
        text="Password Protected",
        variable=password_protected_var,
        font=("Segoe UI", 10),
        bg="#f4f1ea",
        fg="#1f2a44",
        activebackground="#f4f1ea",
        activeforeground="#1f2a44",
        selectcolor="#fffdf8",
    )
    password_check.pack(side="left")

    password_var = tk.StringVar()
    password_entry = tk.Entry(
        password_row,
        textvariable=password_var,
        font=("Consolas", 11),
        bg="#fffdf8",
        fg="#1f2a44",
        insertbackground="#1f2a44",
        relief="solid",
        bd=1,
        show="*",
        state="disabled",
    )
    password_entry.pack(side="left", fill="x", expand=True, padx=(12, 0), ipady=6)

    account_list_box = tk.Listbox(
        accounts_frame,
        height=4,
        font=("Consolas", 10),
        bg="#fffdf8",
        fg="#1f2a44",
        relief="solid",
        bd=1,
    )
    account_list_box.pack(fill="x", pady=(10, 0))

    maintenance_frame = tk.Frame(root, bg="#f4f1ea")
    maintenance_frame.pack(fill="x", padx=20, pady=(12, 0))

    maintenance_label = tk.Label(
        maintenance_frame,
        text="Statement Maintenance",
        font=("Segoe UI Semibold", 11),
        bg="#f4f1ea",
        fg="#1f2a44",
    )
    maintenance_label.pack(anchor="w", pady=(0, 8))

    maintenance_row = tk.Frame(maintenance_frame, bg="#f4f1ea")
    maintenance_row.pack(fill="x")

    clear_month_hint = tk.Label(
        maintenance_row,
        text="The app will detect statement account/date targets from the PDFs automatically.",
        font=("Segoe UI", 10),
        bg="#f4f1ea",
        fg="#5b6474",
    )
    clear_month_hint.pack(side="left")

    clear_before_run_var = tk.BooleanVar(value=True)
    clear_before_run_check = tk.Checkbutton(
        maintenance_frame,
        text="Automatically replace matching output statements before classification",
        variable=clear_before_run_var,
        font=("Segoe UI", 10),
        bg="#f4f1ea",
        fg="#1f2a44",
        activebackground="#f4f1ea",
        activeforeground="#1f2a44",
        selectcolor="#fffdf8",
    )
    clear_before_run_check.pack(anchor="w", pady=(8, 0))

    status_frame = tk.Frame(root, bg="#f4f1ea")
    status_frame.pack(fill="x", padx=20, pady=(12, 0))

    status_card = tk.Frame(status_frame, bg="#16384f", padx=14, pady=10)
    status_card.pack(anchor="w")

    status_dot = tk.Label(
        status_card,
        text="●",
        font=("Segoe UI Symbol", 12),
        bg="#16384f",
        fg="#7aa874",
    )
    status_dot.pack(side="left")

    status_var = tk.StringVar(value="Ready")
    status_label = tk.Label(
        status_card,
        textvariable=status_var,
        font=("Segoe UI Semibold", 10),
        bg="#16384f",
        fg="white",
        padx=8,
    )
    status_label.pack(side="left")

    def refresh_account_list() -> None:
        update_runtime_paths()
        global KNOWN_ACCOUNT_RECORDS, KNOWN_ACCOUNTS
        KNOWN_ACCOUNT_RECORDS = load_account_records(ACCOUNT_NUMBERS_FILE)
        KNOWN_ACCOUNTS = [str(record["account_number"]) for record in KNOWN_ACCOUNT_RECORDS]
        account_list_box.delete(0, tk.END)
        for record in KNOWN_ACCOUNT_RECORDS:
            protection_text = "Yes" if bool(record["password_protected"]) else "No"
            account_list_box.insert(
                tk.END,
                f"{record['account_number']} | Password Protected: {protection_text}",
            )

    def sync_password_state() -> None:
        if password_protected_var.get():
            password_entry.config(state="normal")
        else:
            password_var.set("")
            password_entry.config(state="disabled")

    def append_log(lines: list[str]) -> None:
        log_box.delete("1.0", tk.END)
        log_box.insert(tk.END, "\n".join(lines))
        log_box.see(tk.END)

    def set_controls_enabled(enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for widget in action_buttons:
            widget.config(state=state)

    def animate_loading() -> None:
        nonlocal animation_job, loading_index
        if not is_running:
            animation_job = None
            return

        status_var.set(loading_frames[loading_index])
        status_dot.config(fg="#ffd166" if loading_index % 2 == 0 else "#7bd389")
        loading_index = (loading_index + 1) % len(loading_frames)
        animation_job = root.after(220, animate_loading)

    def finish_run() -> None:
        nonlocal is_running, animation_job, worker_error, worker_results
        if worker_results is None and worker_error is None:
            root.after(120, finish_run)
            return

        is_running = False
        if animation_job is not None:
            root.after_cancel(animation_job)
            animation_job = None

        set_controls_enabled(True)

        if worker_error is not None:
            status_var.set("Run failed")
            status_dot.config(fg="#ff6b6b")
            messagebox.showerror("PDF Classifier", str(worker_error))
            worker_error = None
            return

        results = worker_results or []
        worker_results = None
        status_var.set("Completed")
        status_dot.config(fg="#7aa874")
        refresh_account_list()
        append_log(results)
        missing_warnings = [
            line for line in results if line.startswith("Possible missing transaction dates")
        ]
        if missing_warnings:
            messagebox.showwarning(
                "PDF Classifier",
                "\n".join(missing_warnings),
            )
        if results and results[-1] == "Done.":
            messagebox.showinfo("PDF Classifier", "Classification completed successfully.")
        elif results and results[-1].startswith("Conversion finished."):
            messagebox.showinfo("PDF Classifier", results[-1])
        elif results and results[-1] == "Combining finished.":
            messagebox.showinfo("PDF Classifier", "Combined classified PDFs successfully.")

    def run_classification_worker() -> None:
        nonlocal worker_error, worker_results
        try:
            results = run_classification(auto_clear_detected_periods=clear_before_run_var.get())
            worker_results = results
            worker_error = None
        except Exception as error:
            worker_error = error
            worker_results = None

    def run_converter_worker() -> None:
        nonlocal worker_error, worker_results
        try:
            results = convert_combined_statements()
            worker_results = results
            worker_error = None
        except Exception as error:
            worker_error = error
            worker_results = None

    def run_combine_worker() -> None:
        nonlocal worker_error, worker_results
        try:
            results = combine_classified_pdfs()
            if results:
                results.append("Combining finished.")
            else:
                results = ["No classified PDF folders were found to combine.", "Combining finished."]
            worker_results = results
            worker_error = None
        except Exception as error:
            worker_error = error
            worker_results = None

    def on_run() -> None:
        nonlocal is_running, worker_error, worker_results, loading_index
        if is_running:
            return

        try:
            update_runtime_paths()
        except Exception as error:
            messagebox.showerror("PDF Classifier", str(error))
            return

        is_running = True
        worker_error = None
        worker_results = None
        loading_index = 0
        status_dot.config(fg="#ffd166")
        status_var.set("Starting...")
        append_log(["Working on your files. Please wait..."])
        set_controls_enabled(False)
        animate_loading()
        threading.Thread(target=run_classification_worker, daemon=True).start()
        finish_run()

    def on_convert_combined() -> None:
        nonlocal is_running, worker_error, worker_results, loading_index
        if is_running:
            return

        try:
            update_runtime_paths()
        except Exception as error:
            messagebox.showerror("PDF Classifier", str(error))
            return

        is_running = True
        worker_error = None
        worker_results = None
        loading_index = 0
        status_dot.config(fg="#ffd166")
        status_var.set("Converting...")
        append_log(["Converting combined PDFs. Please wait..."])
        set_controls_enabled(False)
        animate_loading()
        threading.Thread(target=run_converter_worker, daemon=True).start()
        finish_run()

    def on_combine_classified() -> None:
        nonlocal is_running, worker_error, worker_results, loading_index
        if is_running:
            return

        try:
            update_runtime_paths()
        except Exception as error:
            messagebox.showerror("PDF Classifier", str(error))
            return

        is_running = True
        worker_error = None
        worker_results = None
        loading_index = 0
        status_dot.config(fg="#ffd166")
        status_var.set("Combining...")
        append_log(["Combining classified PDFs only. Please wait..."])
        set_controls_enabled(False)
        animate_loading()
        threading.Thread(target=run_combine_worker, daemon=True).start()
        finish_run()

    def on_save_account() -> None:
        update_runtime_paths()
        success, message = save_account_number(
            account_number_var.get(),
            password_protected_var.get(),
            password_var.get(),
            ACCOUNT_NUMBERS_FILE,
        )
        if success:
            account_number_var.set("")
            password_protected_var.set(False)
            password_var.set("")
            sync_password_state()
            refresh_account_list()
            append_log([message, "Account list refreshed."])
        else:
            messagebox.showwarning("PDF Classifier", message)

    def on_clear_month() -> None:
        update_runtime_paths()

        pdf_files = sorted(SOURCE_FOLDER.glob("*.pdf"))
        if not pdf_files:
            messagebox.showwarning(
                "PDF Classifier",
                "No PDFs were found in the input folder to detect months for cleanup.",
            )
            return

        confirmed = messagebox.askyesno(
            "PDF Classifier",
            "Clear only the matching classified statements detected from the current input PDFs?",
        )
        if not confirmed:
            return

        try:
            results = clear_detected_output_months(pdf_files)
            append_log(results)
            messagebox.showinfo(
                "PDF Classifier",
                "Finished clearing the detected output statements.",
            )
        except PermissionError as error:
            blocked_path = getattr(error, "filename", None) or "a file in the selected month"
            messagebox.showerror(
                "PDF Classifier",
                f"Could not clear the detected month because this file is in use or locked:\n{blocked_path}\n\nClose the PDF if it is open, then try again.",
            )
        except Exception as error:
            messagebox.showerror("PDF Classifier", str(error))

    def open_path(path: Path) -> None:
        try:
            path.mkdir(parents=True, exist_ok=True) if path.suffix == "" else path.parent.mkdir(parents=True, exist_ok=True)
            os.startfile(path)
        except Exception as error:
            messagebox.showerror("PDF Classifier", f"Could not open {path}.\n{error}")

    button_frame = tk.Frame(root, bg="#f4f1ea")
    button_frame.pack(fill="x", padx=20, pady=(14, 10))
    for column_index in range(3):
        button_frame.grid_columnconfigure(column_index, weight=1, uniform="actions")

    run_button = tk.Button(
        button_frame,
        text="Run Classification",
        command=lambda: on_run(),
        font=("Segoe UI Semibold", 11),
        bg="#22577a",
        fg="white",
        activebackground="#16384f",
        activeforeground="white",
        padx=14,
        pady=8,
        bd=0,
    )
    run_button.grid(row=0, column=0, sticky="ew", padx=4, pady=4)
    action_buttons.append(run_button)

    save_account_button = tk.Button(
        button_frame,
        text="Save Account Number",
        command=lambda: on_save_account(),
        font=("Segoe UI", 10),
        bg="#7aa874",
        fg="white",
        activebackground="#5b8456",
        activeforeground="white",
        padx=12,
        pady=8,
        bd=0,
    )
    save_account_button.grid(row=0, column=1, sticky="ew", padx=4, pady=4)
    action_buttons.append(save_account_button)

    clear_month_button = tk.Button(
        button_frame,
        text="Clear Detected Statements",
        command=lambda: on_clear_month(),
        font=("Segoe UI", 10),
        bg="#c97c5d",
        fg="white",
        activebackground="#a85d42",
        activeforeground="white",
        padx=12,
        pady=8,
        bd=0,
    )
    clear_month_button.grid(row=0, column=2, sticky="ew", padx=4, pady=4)
    action_buttons.append(clear_month_button)

    mail_button = tk.Button(
        button_frame,
        text="Open Mail Folder",
        command=lambda: (update_runtime_paths(), open_path(MAIL_SOURCE_FOLDER)),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    mail_button.grid(row=1, column=0, sticky="ew", padx=4, pady=4)
    action_buttons.append(mail_button)

    input_button = tk.Button(
        button_frame,
        text="Open Input Folder",
        command=lambda: (update_runtime_paths(), open_path(SOURCE_FOLDER)),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    input_button.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
    action_buttons.append(input_button)

    output_button = tk.Button(
        button_frame,
        text="Open Output Folder",
        command=lambda: (update_runtime_paths(), open_path(OUTPUT_FOLDER)),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    output_button.grid(row=1, column=2, sticky="ew", padx=4, pady=4)
    action_buttons.append(output_button)

    combined_button = tk.Button(
        button_frame,
        text="Open Combined Folder",
        command=lambda: (update_runtime_paths(), open_path(COMBINED_FOLDER)),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    combined_button.grid(row=2, column=0, sticky="ew", padx=4, pady=4)
    action_buttons.append(combined_button)

    accounts_button = tk.Button(
        button_frame,
        text="Open Account File",
        command=lambda: (update_runtime_paths(), open_path(ACCOUNT_NUMBERS_FILE)),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    accounts_button.grid(row=2, column=1, sticky="ew", padx=4, pady=4)
    action_buttons.append(accounts_button)

    convert_combined_button = tk.Button(
        button_frame,
        text="Convert Combined PDFs",
        command=lambda: on_convert_combined(),
        font=("Segoe UI", 10),
        bg="#22577a",
        fg="white",
        activebackground="#16384f",
        activeforeground="white",
        padx=12,
        pady=8,
        bd=0,
    )
    convert_combined_button.grid(row=2, column=2, sticky="ew", padx=4, pady=4)
    action_buttons.append(convert_combined_button)

    converter_settings_button = tk.Button(
        button_frame,
        text="Open Converter Settings",
        command=lambda: open_path(BASE_FOLDER / "statement_converter_settings.json"),
        font=("Segoe UI", 10),
        bg="#d6e6f2",
        fg="#1f2a44",
        activebackground="#bdd5e7",
        activeforeground="#1f2a44",
        padx=12,
        pady=8,
        bd=0,
    )
    converter_settings_button.grid(row=3, column=0, sticky="ew", padx=4, pady=4)
    action_buttons.append(converter_settings_button)

    combine_classified_button = tk.Button(
        button_frame,
        text="Combine Classified PDFs",
        command=lambda: on_combine_classified(),
        font=("Segoe UI", 10),
        bg="#22577a",
        fg="white",
        activebackground="#16384f",
        activeforeground="white",
        padx=12,
        pady=8,
        bd=0,
    )
    combine_classified_button.grid(row=3, column=1, sticky="ew", padx=4, pady=4)
    action_buttons.append(combine_classified_button)

    log_label = tk.Label(
        root,
        text="Execution Log",
        font=("Segoe UI Semibold", 11),
        bg="#f4f1ea",
        fg="#1f2a44",
    )
    log_label.pack(anchor="w", padx=20, pady=(8, 6))

    log_box = scrolledtext.ScrolledText(
        root,
        wrap=tk.WORD,
        font=("Consolas", 10),
        bg="#fffdf8",
        fg="#1f2a44",
        insertbackground="#1f2a44",
        height=14,
    )
    log_box.pack(fill="both", expand=True, padx=20, pady=(0, 20))

    account_entry.bind("<Return>", lambda _event: on_save_account())
    password_entry.bind("<Return>", lambda _event: on_save_account())
    password_check.config(command=sync_password_state)
    update_runtime_paths()
    sync_password_state()
    refresh_account_list()
    sync_header_wrap()
    sync_path_summary_wrap()
    account_entry.focus_set()

    root.mainloop()


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="PDF and MT940 mail classifier with desktop UI.")
    parser.add_argument(
        "--cli",
        action="store_true",
        help="Run in terminal mode instead of opening the UI.",
    )
    parser.add_argument(
        "--convert-combined",
        action="store_true",
        help="Convert every classified combined_statements.pdf into MT940 DAT and final-statement XLSX outputs.",
    )
    parser.add_argument(
        "--combine-classified",
        action="store_true",
        help="Only rebuild combined_statements.pdf files from already-classified PDF folders.",
    )
    parser.add_argument(
        "--converter-output-folder",
        type=Path,
        default=None,
        help="Optional output folder for converted DAT/XLSX files. Defaults to a statement_converter folder beside each combined_statements.pdf.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    if args.combine_classified:
        results = combine_classified_pdfs()
        if results:
            results.append("Combining finished.")
        else:
            results = ["No classified PDF folders were found to combine.", "Combining finished."]
        for line in results:
            print(line)
        return

    if args.convert_combined:
        results = convert_combined_statements(args.converter_output_folder)
        for line in results:
            print(line)
        return

    if args.cli:
        results = run_classification()
        for line in results:
            print(line)
        return

    launch_ui()


if __name__ == "__main__":
    main()
