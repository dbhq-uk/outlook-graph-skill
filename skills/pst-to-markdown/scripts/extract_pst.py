#!/usr/bin/env python3
"""
PST Email Extraction Tool

Extract emails from Outlook PST files into an organized archive of markdown
files, raw email backups, and attachments with integrity verification.

Usage:
    python extract_pst.py <pst_file> <output_dir>           # Full extraction
    python extract_pst.py <pst_file> <output_dir> --append  # Append new emails only

The --append flag enables incremental extraction: it loads the existing index.csv
and skips any emails whose Message-ID is already in the archive. This lets you
update the PST file and re-run extraction to add only new emails.
"""

import argparse
import csv
import hashlib
import os
import re
import subprocess
import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional

# Optional dependencies with fallbacks
try:
    from dateutil import parser as date_parser
    HAS_DATEUTIL = True
except ImportError:
    HAS_DATEUTIL = False

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    HAS_TQDM = False
    # Simple fallback for tqdm
    def tqdm(iterable, desc=None, **kwargs):
        if desc:
            print(f"{desc}...")
        return iterable

try:
    import html2text
    HAS_HTML2TEXT = True
except ImportError:
    HAS_HTML2TEXT = False

# Try libratom/pypff first, fall back to readpst
USE_LIBRATOM = False
try:
    from libratom.lib.pff import PffArchive
    USE_LIBRATOM = True
except ImportError:
    pass


def sanitize_filename(text: str, max_length: int = 50) -> str:
    """Sanitize text for use in filenames."""
    if not text:
        return "unknown"
    # Replace spaces with hyphens
    text = text.replace(" ", "-")
    # Remove problematic characters
    text = re.sub(r'[<>:"/\\|?*@\[\]]', '', text)
    # Collapse multiple hyphens
    text = re.sub(r'-+', '-', text)
    # Remove leading/trailing hyphens
    text = text.strip('-')
    # Truncate
    if len(text) > max_length:
        text = text[:max_length].rstrip('-')
    return text or "unknown"


def sanitize_email(email: str) -> str:
    """Sanitize email address for filename use."""
    if not email:
        return "unknown"
    # Extract just the email part if in "Name <email>" format
    match = re.search(r'<([^>]+)>', email)
    if match:
        email = match.group(1)
    # Remove @ and other special chars but keep dots
    email = re.sub(r'[<>:"/\\|?*\[\]@]', '', email)
    return sanitize_filename(email, max_length=40)


def compute_sha256(filepath: Path) -> str:
    """Compute SHA256 hash of a file."""
    sha256 = hashlib.sha256()
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            sha256.update(chunk)
    return sha256.hexdigest()


def format_size(size_bytes: int) -> str:
    """Format file size in human-readable format."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}" if unit != 'B' else f"{size_bytes} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"


def html_to_markdown(html_content: str) -> str:
    """Convert HTML to Markdown."""
    if not html_content:
        return ""
    if HAS_HTML2TEXT:
        h = html2text.HTML2Text()
        h.ignore_links = False
        h.ignore_images = False
        h.body_width = 0  # Don't wrap lines
        return h.handle(html_content)
    else:
        # Simple fallback: strip HTML tags
        import re
        text = re.sub(r'<br\s*/?>', '\n', html_content, flags=re.IGNORECASE)
        text = re.sub(r'<p\s*/?>', '\n\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'<[^>]+>', '', text)
        text = re.sub(r'&nbsp;', ' ', text)
        text = re.sub(r'&lt;', '<', text)
        text = re.sub(r'&gt;', '>', text)
        text = re.sub(r'&amp;', '&', text)
        text = re.sub(r'&quot;', '"', text)
        return text.strip()


def parse_email_address(addr_str: str) -> tuple[str, str]:
    """Parse email address into (name, email) tuple."""
    if not addr_str:
        return ("", "")
    match = re.match(r'^"?([^"<]*)"?\s*<?([^>]*)>?$', addr_str.strip())
    if match:
        name = match.group(1).strip().strip('"')
        email = match.group(2).strip() or addr_str.strip()
        return (name, email)
    return ("", addr_str.strip())


def format_date_human(dt: datetime) -> str:
    """Format datetime for human-readable display."""
    return dt.strftime("%B %d, %Y at %I:%M %p")


class EmailExtractor:
    """Extract emails from PST file."""

    def __init__(self, pst_path: Path, output_dir: Path,
                 include_deleted: bool = False,
                 target_timezone: str = "UTC",
                 verbose: bool = False,
                 append: bool = False,
                 owner_email: str = None):
        self.pst_path = pst_path
        self.output_dir = output_dir
        self.emails_dir = output_dir / "emails"
        self.include_deleted = include_deleted
        self.target_timezone = target_timezone
        self.verbose = verbose
        self.append = append
        self.owner_email = owner_email

        self.stats = {
            'total': 0,
            'processed': 0,
            'errors': 0,
            'attachments': 0,
            'skipped': 0
        }
        self.index_data = []
        self.existing_message_ids = set()  # For append mode
        self.error_log = []
        self.folder_counts = {}
        self.date_range = {'min': None, 'max': None}

    def log(self, message: str):
        """Log message if verbose mode is on."""
        if self.verbose:
            print(message)

    def log_error(self, message: str):
        """Log error message."""
        self.error_log.append(f"{datetime.now().isoformat()} - {message}")
        print(f"ERROR: {message}", file=sys.stderr)

    def setup_directories(self):
        """Create output directory structure."""
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.emails_dir.mkdir(exist_ok=True)

    def _load_existing_index(self):
        """Load existing index.csv to get already-extracted message IDs."""
        index_path = self.output_dir / "index.csv"
        if not index_path.exists():
            print("No existing index.csv found - will extract all emails")
            return

        print(f"Loading existing index from {index_path}...")
        try:
            with open(index_path, 'r', newline='', encoding='utf-8') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    msg_id = row.get('message_id', '').strip()
                    if msg_id:
                        self.existing_message_ids.add(msg_id)
                    # Also load existing index data so we can merge later
                    self.index_data.append(row)

            print(f"Found {len(self.existing_message_ids)} existing emails (by message ID)")

            # Update stats from existing data
            for row in self.index_data:
                pst_folder = row.get('pst_folder', 'Unknown')
                self.folder_counts[pst_folder] = self.folder_counts.get(pst_folder, 0) + 1

                # Track date range (make timezone-aware for consistent comparisons)
                date_str = row.get('date', '')
                if date_str:
                    try:
                        date = datetime.strptime(date_str, '%Y-%m-%d').replace(tzinfo=timezone.utc)
                        if self.date_range['min'] is None or date < self.date_range['min']:
                            self.date_range['min'] = date
                        if self.date_range['max'] is None or date > self.date_range['max']:
                            self.date_range['max'] = date
                    except ValueError:
                        pass

        except Exception as e:
            self.log_error(f"Failed to load existing index: {e}")
            self.existing_message_ids.clear()
            self.index_data.clear()

    def extract(self):
        """Main extraction method."""
        self.setup_directories()

        print(f"Processing: {self.pst_path}")
        print(f"Output: {self.output_dir}")
        if self.append:
            print("Mode: APPEND (skipping existing emails)")
            self._load_existing_index()
        else:
            print("Mode: OVERWRITE (replacing existing emails)")
        print()

        if USE_LIBRATOM:
            self._extract_with_libratom()
        else:
            self._extract_with_readpst()

        self._generate_index_files()
        self._generate_manifest()
        self._write_extraction_log()

        self._print_summary()

    def _extract_with_libratom(self):
        """Extract using libratom/pypff."""
        print("Using libratom for extraction...")

        with PffArchive(self.pst_path) as archive:
            messages = list(archive.messages())
            self.stats['total'] = len(messages)

            for message in tqdm(messages, desc="Extracting emails"):
                try:
                    self._process_libratom_message(message)
                except Exception as e:
                    self.stats['errors'] += 1
                    self.log_error(f"Failed to process message: {e}")

    def _process_libratom_message(self, message):
        """Process a single message from libratom."""
        # Extract metadata
        sent_date = message.delivery_time or message.creation_time
        if sent_date is None:
            sent_date = datetime.now(timezone.utc)

        subject = message.subject or "(No Subject)"
        sender = message.sender_name or ""
        if message.sender_email_address:
            if sender:
                sender = f"{sender} <{message.sender_email_address}>"
            else:
                sender = message.sender_email_address

        # Get recipients
        to_list = []
        cc_list = []
        bcc_list = []

        # libratom message properties
        if hasattr(message, 'plain_text_body'):
            body_text = message.plain_text_body or ""
        else:
            body_text = ""

        if hasattr(message, 'html_body'):
            body_html = message.html_body or ""
        else:
            body_html = ""

        # Convert HTML to markdown if available
        if body_html:
            body_md = html_to_markdown(body_html)
        else:
            body_md = body_text

        # Get headers
        headers = ""
        if hasattr(message, 'transport_headers'):
            headers = message.transport_headers or ""

        # Parse recipients from headers if available
        if headers:
            to_match = re.search(r'^To:\s*(.+?)(?=\n\S|\Z)', headers, re.MULTILINE | re.DOTALL)
            if to_match:
                to_list = [addr.strip() for addr in to_match.group(1).replace('\n', '').split(',')]
            cc_match = re.search(r'^Cc:\s*(.+?)(?=\n\S|\Z)', headers, re.MULTILINE | re.DOTALL)
            if cc_match:
                cc_list = [addr.strip() for addr in cc_match.group(1).replace('\n', '').split(',')]

        # Get message ID
        message_id = ""
        if headers:
            mid_match = re.search(r'^Message-ID:\s*(.+)$', headers, re.MULTILINE | re.IGNORECASE)
            if mid_match:
                message_id = mid_match.group(1).strip()

        # Folder path
        folder_path = "Unknown"

        # Process attachments
        attachments = []
        if hasattr(message, 'attachments'):
            for i, attachment in enumerate(message.attachments):
                att_info = self._save_attachment(attachment, i + 1, None)  # Will set path later
                if att_info:
                    attachments.append(att_info)

        # Create email data dict
        email_data = {
            'sent_date': sent_date,
            'subject': subject,
            'sender': sender,
            'to_list': to_list,
            'cc_list': cc_list,
            'bcc_list': bcc_list,
            'body_md': body_md,
            'body_text': body_text,
            'headers': headers,
            'message_id': message_id,
            'folder_path': folder_path,
            'attachments': attachments,
            'raw_message': message,
        }

        self._save_email(email_data)

    def _extract_with_readpst(self):
        """Extract using readpst command-line tool."""
        print("Using readpst for extraction...")

        # Check if readpst is available
        readpst_available = False
        try:
            subprocess.run(['readpst', '-V'], capture_output=True, check=True)
            readpst_available = True
        except (subprocess.CalledProcessError, FileNotFoundError):
            pass

        if not readpst_available:
            # Check if input is a directory of pre-extracted emails
            if self.pst_path.is_dir():
                print(f"Processing pre-extracted emails from: {self.pst_path}")
                self._process_eml_directory(self.pst_path)
                return

            print("ERROR: readpst not found and input is not a directory.")
            print()
            print("Options:")
            print("  1. Install pst-utils (recommended):")
            print("     Ubuntu/Debian: sudo apt install pst-utils")
            print("     macOS: brew install libpst")
            print()
            print("  2. Install libratom:")
            print("     pip install libratom")
            print()
            print("  3. Pre-extract emails using readpst on another machine:")
            print(f"     readpst -e -o extracted_emails/ {self.pst_path.name}")
            print("     Then run this tool on the extracted_emails/ directory")
            sys.exit(1)

        # Create temp directory for readpst output
        with tempfile.TemporaryDirectory() as tmpdir:
            tmppath = Path(tmpdir)

            # Run readpst with EML output
            print("Extracting emails from PST (this may take a while)...")
            result = subprocess.run([
                'readpst',
                '-e',  # Extract each message to separate file
                '-o', str(tmppath),
                str(self.pst_path)
            ], capture_output=True, text=True)

            if result.returncode != 0:
                print(f"readpst failed: {result.stderr}")
                sys.exit(1)

            self._process_eml_directory(tmppath)

    def _process_eml_directory(self, eml_dir: Path):
        """Process a directory containing .eml files."""
        # Find all extracted .eml files
        eml_files = list(eml_dir.rglob('*.eml'))
        self.stats['total'] = len(eml_files)

        print(f"Found {len(eml_files)} emails")

        for eml_path in tqdm(eml_files, desc="Processing emails"):
            try:
                self._process_eml_file(eml_path, eml_dir)
            except Exception as e:
                self.stats['errors'] += 1
                self.log_error(f"Failed to process {eml_path.name}: {e}")

    def _process_eml_file(self, eml_path: Path, base_dir: Path):
        """Process a single .eml file."""
        import email
        from email import policy

        with open(eml_path, 'rb') as f:
            msg = email.message_from_binary_file(f, policy=policy.default)

        # Extract date
        date_str = msg.get('Date', '')
        try:
            if HAS_DATEUTIL and date_str:
                sent_date = date_parser.parse(date_str)
            elif date_str:
                # Try standard library email.utils
                from email.utils import parsedate_to_datetime
                sent_date = parsedate_to_datetime(date_str)
            else:
                sent_date = datetime.now(timezone.utc)
        except Exception:
            sent_date = datetime.now(timezone.utc)

        # Ensure timezone aware
        if sent_date.tzinfo is None:
            sent_date = sent_date.replace(tzinfo=timezone.utc)

        # Extract basic fields
        subject = msg.get('Subject', '(No Subject)')
        sender = msg.get('From', '')
        to_str = msg.get('To', '')
        cc_str = msg.get('Cc', '')
        bcc_str = msg.get('Bcc', '')

        # Fix MAILER-DAEMON sent emails - these are actually emails sent by the user
        # readpst/libpst exports sent emails with MAILER-DAEMON as the email address
        # but includes the real sender info in X-libpst-forensic-sender header
        if 'MAILER-DAEMON' in sender and self.owner_email:
            forensic_sender = msg.get('X-libpst-forensic-sender', '')
            # Extract display name from the From field
            sender_name_match = re.match(r'^"?([^"<]+)"?\s*<', sender)
            display_name = sender_name_match.group(1).strip() if sender_name_match else ''

            # Use owner email to reconstruct the real sender
            if forensic_sender:
                if display_name:
                    sender = f'"{display_name}" <{self.owner_email}>'
                else:
                    sender = self.owner_email

        to_list = [addr.strip() for addr in to_str.split(',') if addr.strip()] if to_str else []
        cc_list = [addr.strip() for addr in cc_str.split(',') if addr.strip()] if cc_str else []
        bcc_list = [addr.strip() for addr in bcc_str.split(',') if addr.strip()] if bcc_str else []

        message_id = msg.get('Message-ID', '')

        # Get folder path from relative path
        rel_path = eml_path.relative_to(base_dir)
        folder_path = str(rel_path.parent) if rel_path.parent != Path('.') else "Root"

        # Extract body
        body_text = ""
        body_html = ""

        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == 'text/plain' and not body_text:
                    try:
                        body_text = part.get_content()
                    except Exception:
                        body_text = str(part.get_payload(decode=True) or b'', errors='replace')
                elif content_type == 'text/html' and not body_html:
                    try:
                        body_html = part.get_content()
                    except Exception:
                        body_html = str(part.get_payload(decode=True) or b'', errors='replace')
        else:
            content_type = msg.get_content_type()
            try:
                content = msg.get_content()
            except Exception:
                content = str(msg.get_payload(decode=True) or b'', errors='replace')
            if content_type == 'text/html':
                body_html = content
            else:
                body_text = content

        # Convert to markdown
        if body_html:
            body_md = html_to_markdown(body_html)
        else:
            body_md = body_text

        # Get headers
        headers = "\n".join(f"{k}: {v}" for k, v in msg.items())

        # Create email data
        email_data = {
            'sent_date': sent_date,
            'subject': subject,
            'sender': sender,
            'to_list': to_list,
            'cc_list': cc_list,
            'bcc_list': bcc_list,
            'body_md': body_md,
            'body_text': body_text,
            'headers': headers,
            'message_id': message_id,
            'folder_path': folder_path,
            'attachments': [],
            'raw_eml_path': eml_path,
        }

        # Extract attachments
        if msg.is_multipart():
            att_index = 1
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    att_data = part.get_payload(decode=True)
                    if att_data:
                        filename = part.get_filename() or f"attachment_{att_index}"
                        content_type = part.get_content_type()
                        email_data['attachments'].append({
                            'original_name': filename,
                            'data': att_data,
                            'content_type': content_type,
                            'index': att_index,
                        })
                        att_index += 1

        self._save_email(email_data)

    def _save_email(self, email_data: dict):
        """Save email to output directory."""
        # In append mode, skip emails we've already extracted
        message_id = email_data.get('message_id', '').strip()
        if self.append and message_id and message_id in self.existing_message_ids:
            self.stats['skipped'] += 1
            self.log(f"Skipping existing email: {email_data.get('subject', '(No Subject)')[:50]}")
            return

        sent_date = email_data['sent_date']
        subject = email_data['subject']
        sender = email_data['sender']
        to_list = email_data['to_list']

        # Parse sender
        sender_name, sender_email = parse_email_address(sender)

        # Parse primary recipient
        if to_list:
            to_name, to_email = parse_email_address(to_list[0])
        else:
            to_name, to_email = "", "unknown"

        # Create folder name - prefer name over email when available
        date_str = sent_date.strftime("%Y-%m-%d_%H%M%S")
        sender_sanitized = sanitize_filename(sender_name, max_length=40) if sender_name else sanitize_email(sender_email)
        to_sanitized = sanitize_filename(to_name, max_length=40) if to_name else sanitize_email(to_email)
        subject_sanitized = sanitize_filename(subject, max_length=50)

        folder_name = f"{date_str}_from-{sender_sanitized}_to-{to_sanitized}_{subject_sanitized}"

        # Ensure path isn't too long (Windows safety)
        if len(folder_name) > 150:
            folder_name = folder_name[:150]

        # Build the full path including PST folder structure
        pst_folder = email_data.get('folder_path', 'Unknown')
        if pst_folder and pst_folder not in ('Unknown', 'Root', '.'):
            # Sanitize each path component
            folder_parts = [sanitize_filename(part, max_length=50) for part in pst_folder.split('/') if part]
            subfolder_path = Path(*folder_parts) if folder_parts else Path()
            base_dir = self.emails_dir / subfolder_path
        else:
            base_dir = self.emails_dir

        # Handle duplicates
        email_folder = base_dir / folder_name
        counter = 1
        original_folder_name = folder_name
        while email_folder.exists():
            folder_name = f"{original_folder_name}-{counter:03d}"
            email_folder = base_dir / folder_name
            counter += 1

        email_folder.mkdir(parents=True, exist_ok=True)

        # Calculate relative path from emails_dir for index
        relative_folder_path = str(email_folder.relative_to(self.emails_dir))

        # Save attachments
        saved_attachments = []
        for att in email_data.get('attachments', []):
            att_info = self._write_attachment(email_folder, att)
            if att_info:
                saved_attachments.append(att_info)
                self.stats['attachments'] += 1

        # Save raw .eml file
        eml_path = email_folder / "email.eml"
        if 'raw_eml_path' in email_data:
            # Copy the original eml file
            import shutil
            shutil.copy2(email_data['raw_eml_path'], eml_path)
        else:
            # Generate .eml from message data
            self._generate_eml(email_folder, email_data)

        # Generate email.md
        self._generate_email_md(email_folder, email_data, saved_attachments)

        # Generate checksums
        self._generate_checksums(email_folder)

        # Update stats and index
        self.stats['processed'] += 1

        # Track date range (ensure timezone-aware for consistent comparisons)
        if sent_date.tzinfo is None:
            sent_date = sent_date.replace(tzinfo=timezone.utc)
            email_data['sent_date'] = sent_date  # Update for consistency
        if self.date_range['min'] is None or sent_date < self.date_range['min']:
            self.date_range['min'] = sent_date
        if self.date_range['max'] is None or sent_date > self.date_range['max']:
            self.date_range['max'] = sent_date

        # Track folder counts
        pst_folder = email_data.get('folder_path', 'Unknown')
        self.folder_counts[pst_folder] = self.folder_counts.get(pst_folder, 0) + 1

        # Add to index
        sender_name, sender_email = parse_email_address(email_data['sender'])
        to_name, to_email = parse_email_address(to_list[0]) if to_list else ("", "")

        self.index_data.append({
            'folder_name': relative_folder_path,
            'date': sent_date.strftime("%Y-%m-%d"),
            'time': sent_date.strftime("%H:%M:%S"),
            'from_email': sender_email,
            'from_name': sender_name,
            'to_email': to_email,
            'to_name': to_name,
            'cc': ', '.join(email_data.get('cc_list', [])),
            'subject': subject,
            'attachment_count': len(saved_attachments),
            'has_body': bool(email_data.get('body_md', '').strip()),
            'pst_folder': pst_folder,
            'message_id': email_data.get('message_id', ''),
        })

    def _write_attachment(self, email_folder: Path, att: dict) -> Optional[dict]:
        """Write attachment to disk."""
        try:
            original_name = att.get('original_name', 'attachment')
            index = att.get('index', 1)

            # Get file extension
            ext = Path(original_name).suffix or ''
            base_name = Path(original_name).stem

            # Sanitize and create filename
            sanitized_base = sanitize_filename(base_name, max_length=40)
            filename = f"attachment_{index:03d}_{sanitized_base}{ext}"

            filepath = email_folder / filename

            # Write data
            data = att.get('data', b'')
            with open(filepath, 'wb') as f:
                f.write(data)

            return {
                'filename': filename,
                'original_name': original_name,
                'size_bytes': len(data),
                'content_type': att.get('content_type', 'application/octet-stream'),
                'sha256': compute_sha256(filepath),
            }
        except Exception as e:
            self.log_error(f"Failed to save attachment {att.get('original_name', 'unknown')}: {e}")
            return None

    def _generate_eml(self, email_folder: Path, email_data: dict):
        """Generate .eml file from email data."""
        eml_path = email_folder / "email.eml"

        # Build a basic RFC 822 message
        lines = []

        if email_data.get('headers'):
            lines.append(email_data['headers'])
        else:
            lines.append(f"Date: {email_data['sent_date'].strftime('%a, %d %b %Y %H:%M:%S %z')}")
            lines.append(f"From: {email_data['sender']}")
            if email_data['to_list']:
                lines.append(f"To: {', '.join(email_data['to_list'])}")
            if email_data.get('cc_list'):
                lines.append(f"Cc: {', '.join(email_data['cc_list'])}")
            lines.append(f"Subject: {email_data['subject']}")
            if email_data.get('message_id'):
                lines.append(f"Message-ID: {email_data['message_id']}")
            lines.append("MIME-Version: 1.0")
            lines.append("Content-Type: text/plain; charset=UTF-8")

        lines.append("")
        lines.append(email_data.get('body_text', '') or email_data.get('body_md', ''))

        with open(eml_path, 'w', encoding='utf-8', errors='replace') as f:
            f.write('\n'.join(lines))

    def _generate_email_md(self, email_folder: Path, email_data: dict, attachments: list):
        """Generate email.md file."""
        sent_date = email_data['sent_date']
        subject = email_data['subject']

        sender_name, sender_email = parse_email_address(email_data['sender'])
        sender_display = f"{sender_name} <{sender_email}>" if sender_name else sender_email

        to_display = ', '.join(email_data.get('to_list', []))
        cc_display = ', '.join(email_data.get('cc_list', []))

        # Build YAML frontmatter
        yaml_lines = [
            "---",
            f'message_id: "{email_data.get("message_id", "")}"',
            f'date: "{sent_date.isoformat()}"',
            f'from: "{email_data["sender"]}"',
            "to:",
        ]
        for addr in email_data.get('to_list', []):
            yaml_lines.append(f'  - "{addr}"')
        if not email_data.get('to_list'):
            yaml_lines.append("  []")

        yaml_lines.append("cc:")
        for addr in email_data.get('cc_list', []):
            yaml_lines.append(f'  - "{addr}"')
        if not email_data.get('cc_list'):
            yaml_lines.append("  []")

        yaml_lines.append("bcc: []")
        yaml_lines.append(f'subject: "{subject}"')
        yaml_lines.append(f'has_attachments: {str(bool(attachments)).lower()}')
        yaml_lines.append(f'attachment_count: {len(attachments)}')

        if attachments:
            yaml_lines.append("attachments:")
            for att in attachments:
                yaml_lines.append(f'  - filename: "{att["filename"]}"')
                yaml_lines.append(f'    original_name: "{att["original_name"]}"')
                yaml_lines.append(f'    size_bytes: {att["size_bytes"]}')
                yaml_lines.append(f'    content_type: "{att["content_type"]}"')
                yaml_lines.append(f'    sha256: "{att["sha256"]}"')

        yaml_lines.append(f'pst_folder: "{email_data.get("folder_path", "Unknown")}"')
        yaml_lines.append(f'extraction_date: "{datetime.now(timezone.utc).isoformat()}"')
        yaml_lines.append(f'source_file: "{self.pst_path.name}"')
        yaml_lines.append("---")

        # Build markdown content
        md_lines = [
            "",
            f"# {subject}",
            "",
            f"**From:** {sender_display}",
            f"**To:** {to_display}",
        ]

        if cc_display:
            md_lines.append(f"**CC:** {cc_display}")

        md_lines.append(f"**Date:** {format_date_human(sent_date)}")
        md_lines.append(f"**Subject:** {subject}")
        md_lines.append("")
        md_lines.append("---")
        md_lines.append("")
        md_lines.append("## Body")
        md_lines.append("")
        md_lines.append(email_data.get('body_md', '') or "(No body content)")
        md_lines.append("")

        if attachments:
            md_lines.append("---")
            md_lines.append("")
            md_lines.append("## Attachments")
            md_lines.append("")
            for i, att in enumerate(attachments, 1):
                size_str = format_size(att['size_bytes'])
                md_lines.append(f"{i}. [{att['original_name']}](./{att['filename']}) ({size_str})")
            md_lines.append("")

        md_lines.append("---")
        md_lines.append("")
        md_lines.append("## Original Headers")
        md_lines.append("")
        md_lines.append("```")
        md_lines.append(email_data.get('headers', '(Headers not available)'))
        md_lines.append("```")

        # Write file
        md_path = email_folder / "email.md"
        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(yaml_lines))
            f.write('\n'.join(md_lines))

    def _generate_checksums(self, email_folder: Path):
        """Generate checksums.sha256 for all files in folder."""
        checksums_path = email_folder / "checksums.sha256"
        lines = []

        for filepath in sorted(email_folder.iterdir()):
            if filepath.name != "checksums.sha256" and filepath.is_file():
                sha256 = compute_sha256(filepath)
                lines.append(f"{sha256}  {filepath.name}")

        with open(checksums_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _generate_index_files(self):
        """Generate index.csv and index.md."""
        # Sort by date
        self.index_data.sort(key=lambda x: (x['date'], x['time']))

        # Generate CSV
        csv_path = self.output_dir / "index.csv"
        with open(csv_path, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=[
                'folder_name', 'date', 'time', 'from_email', 'from_name',
                'to_email', 'to_name', 'cc', 'subject', 'attachment_count',
                'has_body', 'pst_folder', 'message_id'
            ])
            writer.writeheader()
            writer.writerows(self.index_data)

        # Generate MD
        md_path = self.output_dir / "index.md"

        date_range_str = "N/A"
        if self.date_range['min'] and self.date_range['max']:
            date_range_str = f"{self.date_range['min'].strftime('%Y-%m-%d')} to {self.date_range['max'].strftime('%Y-%m-%d')}"

        source_size = format_size(self.pst_path.stat().st_size) if self.pst_path.is_file() else "N/A"

        lines = [
            "# Email Archive Index",
            "",
            f"**Source:** {self.pst_path.name} ({source_size})",
            f"**Extracted:** {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M:%S')} UTC",
            f"**Total Emails:** {self.stats['processed']:,}",
            f"**Total Attachments:** {self.stats['attachments']:,}",
            f"**Date Range:** {date_range_str}",
            "",
            "## Statistics",
            "",
            "| Folder | Count |",
            "|--------|-------|",
        ]

        for folder, count in sorted(self.folder_counts.items()):
            lines.append(f"| {folder} | {count} |")

        lines.append("")
        lines.append("## Timeline")
        lines.append("")

        # Group by year/month
        current_year = None
        current_month = None

        for item in self.index_data:
            date = datetime.strptime(item['date'], '%Y-%m-%d')
            year = date.year
            month = date.strftime('%B %Y')

            if year != current_year:
                lines.append(f"### {year}")
                lines.append("")
                current_year = year
                current_month = None

            if month != current_month:
                lines.append(f"#### {month}")
                lines.append("")
                current_month = month

            att_str = f" ({item['attachment_count']} attachments)" if item['attachment_count'] else ""
            from_display = item['from_name'] or item['from_email']
            to_display = item['to_name'] or item['to_email']

            lines.append(f"- [{item['date']} {item['time']}](./emails/{item['folder_name']}/) - "
                        f"**{from_display}** \u2192 {to_display} - \"{item['subject']}\"{att_str}")

        with open(md_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')

    def _generate_manifest(self):
        """Generate master manifest.sha256."""
        manifest_path = self.output_dir / "manifest.sha256"
        lines = [
            f"# Generated: {datetime.now(timezone.utc).isoformat()}",
            f"# Source: {self.pst_path.name}",
        ]

        if self.pst_path.is_file():
            lines.append(f"# SHA256 of source: {compute_sha256(self.pst_path)}")

        lines.append("")

        # Hash all checksums files
        for checksums_file in sorted(self.emails_dir.rglob("checksums.sha256")):
            rel_path = checksums_file.relative_to(self.output_dir)
            sha256 = compute_sha256(checksums_file)
            lines.append(f"{sha256}  {rel_path}")

        # Hash index files
        for index_file in ['index.csv', 'index.md']:
            filepath = self.output_dir / index_file
            if filepath.exists():
                sha256 = compute_sha256(filepath)
                lines.append(f"{sha256}  {index_file}")

        with open(manifest_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _write_extraction_log(self):
        """Write extraction log."""
        log_path = self.output_dir / "extraction_log.txt"

        lines = [
            "=" * 60,
            "PST Email Extraction Log",
            "=" * 60,
            "",
            f"Source: {self.pst_path}",
            f"Output: {self.output_dir}",
            f"Mode: {'APPEND' if self.append else 'OVERWRITE'}",
            f"Started: {datetime.now().isoformat()}",
            "",
            "Statistics:",
            f"  Total messages found: {self.stats['total']}",
            f"  Successfully processed: {self.stats['processed']}",
            f"  Skipped (already exist): {self.stats['skipped']}",
            f"  Errors: {self.stats['errors']}",
            f"  Attachments extracted: {self.stats['attachments']}",
            "",
        ]

        if self.error_log:
            lines.append("Errors:")
            lines.extend(f"  {error}" for error in self.error_log)
            lines.append("")

        lines.append("=" * 60)

        with open(log_path, 'w') as f:
            f.write('\n'.join(lines) + '\n')

    def _print_summary(self):
        """Print extraction summary."""
        print()
        print("=" * 60)
        print("Extraction Complete")
        print("=" * 60)
        print(f"  Emails processed: {self.stats['processed']:,}")
        if self.append:
            print(f"  Emails skipped (already exist): {self.stats['skipped']:,}")
        print(f"  Attachments: {self.stats['attachments']:,}")
        print(f"  Errors: {self.stats['errors']:,}")
        print()
        print(f"Output: {self.output_dir}")
        print("=" * 60)


def main():
    parser = argparse.ArgumentParser(
        description="Extract emails from Outlook PST files into organized markdown archive"
    )
    parser.add_argument("pst_file", help="Path to PST file (or directory of .eml files)")
    parser.add_argument("output_dir", help="Output directory")
    parser.add_argument("--include-deleted", action="store_true",
                       help="Include deleted items")
    parser.add_argument("--timezone", default="UTC",
                       help="Target timezone for dates (default: UTC)")
    parser.add_argument("--verbose", "-v", action="store_true",
                       help="Verbose output")
    parser.add_argument("--append", action="store_true",
                       help="Append mode: skip emails already in the archive (by message ID)")
    parser.add_argument("--owner-email",
                       help="PST owner's email address (used to fix MAILER-DAEMON sent items)")

    args = parser.parse_args()

    pst_path = Path(args.pst_file)
    output_dir = Path(args.output_dir)

    if not pst_path.exists():
        print(f"Error: PST file not found: {pst_path}")
        sys.exit(1)

    extractor = EmailExtractor(
        pst_path=pst_path,
        output_dir=output_dir,
        include_deleted=args.include_deleted,
        target_timezone=args.timezone,
        verbose=args.verbose,
        append=args.append,
        owner_email=args.owner_email,
    )

    extractor.extract()


if __name__ == "__main__":
    main()
