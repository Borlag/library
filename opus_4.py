#!/usr/bin/env python3
"""
Spec Sync - Aviation Specification Organizer
Fully autonomous tool for organizing PDF specifications and drawings
Version: 1.1.0
"""

__version__ = "1.1.0"

import sys
import os
import re
import hashlib
import shutil
import json
import sqlite3
import logging
import smtplib
import traceback
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Set
from dataclasses import dataclass, asdict
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Third-party imports
import yaml
import pandas as pd
import pytesseract
from PIL import Image
import PyPDF2
import win32com.client
from tqdm import tqdm
from dotenv import load_dotenv

# GUI imports
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QListWidget, QLineEdit, QComboBox,
    QCheckBox, QRadioButton, QButtonGroup, QProgressBar,
    QTextEdit, QFileDialog, QMessageBox, QMenuBar, QStatusBar,
    QListWidgetItem, QGroupBox, QTabWidget, QSpinBox
)
from PySide6.QtCore import Qt, QThread, Signal, QTimer
from PySide6.QtGui import QAction, QFont

# OpenAI imports (optional)
try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

# Load environment variables
load_dotenv()

# Configure logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# File handler for detailed logs
fh = logging.FileHandler('sync.log', encoding='utf-8')
fh.setLevel(logging.DEBUG)

# Console handler for GUI
ch = logging.StreamHandler()
ch.setLevel(logging.INFO)

# Formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
fh.setFormatter(formatter)
ch.setFormatter(formatter)

logger.addHandler(fh)
logger.addHandler(ch)


@dataclass
class FileInfo:
    """Domain model for specification file"""
    path: str
    abbreviation: str
    sha256: str
    size_mb: float
    modified: str
    category: str = ""
    revision: str = ""
    tags: List[str] = None
    is_duplicate: bool = False
    original_path: str = ""

    def __post_init__(self):
        if self.tags is None:
            self.tags = []


class SpecDatabase:
    """SQLite database for inventory management"""
    
    def __init__(self, db_path: str = "inventory.db"):
        self.db_path = db_path
        self.conn = sqlite3.connect(db_path)
        self._create_tables()
    
    def _create_tables(self):
        """Create necessary database tables"""
        cursor = self.conn.cursor()
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS files (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sha256 TEXT UNIQUE NOT NULL,
                abbreviation TEXT NOT NULL,
                filename TEXT NOT NULL,
                original_path TEXT NOT NULL,
                target_path TEXT,
                size_mb REAL,
                modified TEXT,
                category TEXT,
                revision TEXT,
                tags TEXT,
                processed_at TEXT DEFAULT CURRENT_TIMESTAMP
            )
        ''')
        
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS duplicates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                sha256 TEXT NOT NULL,
                path TEXT NOT NULL,
                link_path TEXT,
                FOREIGN KEY (sha256) REFERENCES files(sha256)
            )
        ''')
        
        self.conn.commit()
    
    def file_exists(self, sha256: str) -> bool:
        """Check if file with given SHA256 already exists"""
        cursor = self.conn.cursor()
        try:
            cursor.execute("SELECT COUNT(*) FROM files WHERE sha256 = ?", (sha256,))
            return cursor.fetchone()[0] > 0
        except sqlite3.OperationalError:
            # Table might not exist yet or column missing
            return False
    
    def add_file(self, file_info: FileInfo, target_path: str):
        """Add file to inventory"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT OR REPLACE INTO files 
            (sha256, abbreviation, filename, original_path, target_path, 
             size_mb, modified, category, revision, tags)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (
            file_info.sha256,
            file_info.abbreviation,
            os.path.basename(file_info.path),
            file_info.path,
            target_path,
            file_info.size_mb,
            file_info.modified,
            file_info.category,
            file_info.revision,
            json.dumps(file_info.tags)
        ))
        self.conn.commit()
    
    def add_duplicate(self, sha256: str, path: str, link_path: str = None):
        """Record duplicate file"""
        cursor = self.conn.cursor()
        cursor.execute('''
            INSERT INTO duplicates (sha256, path, link_path)
            VALUES (?, ?, ?)
        ''', (sha256, path, link_path))
        self.conn.commit()
    
    def get_all_files(self) -> List[Dict]:
        """Get all processed files"""
        cursor = self.conn.cursor()
        cursor.execute("SELECT * FROM files")
        columns = [description[0] for description in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]
    
    def close(self):
        """Close database connection"""
        self.conn.close()


class AbbreviationExtractor:
    """Extract and validate abbreviations from filenames"""
    
    def __init__(self, abbrev_file: str, case_sensitive: bool = True):
        self.case_sensitive = case_sensitive
        self.abbreviations = self._load_abbreviations(abbrev_file)
        self.pattern = re.compile(r'[A-Z0-9\-]{3,15}')
    
    def _load_abbreviations(self, file_path: str) -> Set[str]:
        """Load valid abbreviations from file"""
        abbrevs = set()
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    abbrev = line.strip()
                    if abbrev:
                        if self.case_sensitive:
                            abbrevs.add(abbrev)
                        else:
                            abbrevs.add(abbrev.upper())
        except Exception as e:
            logger.error(f"Error loading abbreviations: {e}")
        return abbrevs
    
    def extract(self, filename: str) -> Optional[str]:
        """Extract abbreviation from filename"""
        if self.case_sensitive:
            # For case-sensitive mode, look for exact matches
            # First try to find uppercase patterns
            matches = self.pattern.findall(filename)
            for match in matches:
                if match in self.abbreviations:
                    return match
            
            # Also check lowercase patterns if abbreviations contain lowercase
            pattern_lower = re.compile(r'[a-z0-9\-]{3,15}')
            matches_lower = pattern_lower.findall(filename)
            for match in matches_lower:
                if match in self.abbreviations:
                    return match
        else:
            # Case-insensitive mode (original behavior)
            matches = self.pattern.findall(filename.upper())
            for match in matches:
                if match in self.abbreviations:
                    return match
        return None


class RevisionExtractor:
    """Extract revision information from filenames"""
    
    PATTERNS = [
        r'[-_]REV\.?\s*([A-Z0-9]+)',
        r'[-_]R([0-9]+)',
        r'ISSUE\s*([A-Z0-9]+)',
        r'VERSION\s*([A-Z0-9]+)',
        r'V([0-9]+(?:\.[0-9]+)?)'
    ]
    
    @classmethod
    def extract(cls, filename: str) -> str:
        """Extract revision from filename"""
        for pattern in cls.PATTERNS:
            match = re.search(pattern, filename, re.IGNORECASE)
            if match:
                return match.group(1).upper()
        return ""


class FileHasher:
    """Calculate SHA-256 hash of files"""
    
    @staticmethod
    def calculate_sha256(file_path: str, chunk_size: int = 8192) -> str:
        """Calculate SHA-256 hash of a file"""
        sha256_hash = hashlib.sha256()
        try:
            with open(file_path, "rb") as f:
                while chunk := f.read(chunk_size):
                    sha256_hash.update(chunk)
            return sha256_hash.hexdigest()
        except Exception as e:
            logger.error(f"Error hashing file {file_path}: {e}")
            return ""


class Categorizer:
    """Three-tier categorization system"""
    
    def __init__(self, rules_file: str, use_llm: bool = False, llm_model: str = "gpt-4"):
        self.rules = self._load_rules(rules_file)
        self.use_llm = use_llm
        self.llm_model = llm_model
    
    def _load_rules(self, file_path: str) -> Dict:
        """Load categorization rules from YAML"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return yaml.safe_load(f)
        except Exception as e:
            logger.error(f"Error loading rules: {e}")
            return {}
    
    def categorize(self, file_info: FileInfo) -> Tuple[str, List[str]]:
        """Categorize file using 3-tier approach"""
        # Tier 1: Keyword matching
        category, tags = self._keyword_match(file_info)
        if category:
            return category, tags
        
        # Tier 2: Content parsing
        category, tags = self._content_parse(file_info)
        if category:
            return category, tags
        
        # Tier 3: LLM fallback
        if self.use_llm:
            category, tags = self._llm_categorize(file_info)
            if category:
                return category, tags
        
        return "Uncategorized", []
    
    def _keyword_match(self, file_info: FileInfo) -> Tuple[str, List[str]]:
        """Match keywords in filename"""
        filename = os.path.basename(file_info.path).upper()
        
        for category, keywords in self.rules.get('keywords', {}).items():
            for keyword in keywords:
                if keyword.upper() in filename:
                    tags = [kw for kw in keywords if kw.upper() in filename]
                    return category, tags
        
        return "", []
    
    def _content_parse(self, file_info: FileInfo) -> Tuple[str, List[str]]:
        """Parse file content for categorization"""
        try:
            ext = os.path.splitext(file_info.path)[1].lower()
            
            if ext == '.pdf':
                text = self._extract_pdf_text(file_info.path)
            elif ext in ['.tiff', '.tif', '.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp']:
                text = self._ocr_image(file_info.path)
            elif ext in ['.doc', '.docx']:
                # For Word documents, use python-docx if available
                try:
                    import docx
                    doc = docx.Document(file_info.path)
                    text = '\n'.join([para.text for para in doc.paragraphs])
                except ImportError:
                    logger.warning("python-docx not installed, skipping Word file content parsing")
                    return "", []
            elif ext in ['.txt', '.xml', '.json', '.yaml', '.yml']:
                # Plain text files
                with open(file_info.path, 'r', encoding='utf-8', errors='ignore') as f:
                    text = f.read()
            else:
                # For other formats, try to extract text if possible
                return "", []
            
            # Search for category indicators in content
            text_upper = text.upper()
            for category, indicators in self.rules.get('content_indicators', {}).items():
                for indicator in indicators:
                    if indicator.upper() in text_upper:
                        return category, [indicator]
            
        except Exception as e:
            logger.error(f"Error parsing content of {file_info.path}: {e}")
        
        return "", []
    
    def _extract_pdf_text(self, pdf_path: str) -> str:
        """Extract text from PDF"""
        text = ""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page in pdf_reader.pages[:3]:  # First 3 pages only
                    text += page.extract_text()
        except Exception as e:
            logger.error(f"Error extracting PDF text: {e}")
        return text
    
    def _ocr_image(self, image_path: str) -> str:
        """OCR image file"""
        try:
            image = Image.open(image_path)
            text = pytesseract.image_to_string(image, lang='eng')
            return text
        except Exception as e:
            logger.error(f"Error OCR processing image: {e}")
            return ""
    
    def _llm_categorize(self, file_info: FileInfo) -> Tuple[str, List[str]]:
        """Use LLM for categorization"""
        if not OPENAI_AVAILABLE or not os.getenv('OPENAI_API_KEY'):
            logger.warning("OpenAI not available for LLM categorization")
            return "", []
        
        try:
            openai.api_key = os.getenv('OPENAI_API_KEY')
            
            prompt = f"""Categorize this aviation specification file:
            Filename: {os.path.basename(file_info.path)}
            Abbreviation: {file_info.abbreviation}
            
            Categories: {list(self.rules.get('keywords', {}).keys())}
            
            Return JSON: {{"category": "...", "tags": ["...", "..."]}}"""
            
            response = openai.ChatCompletion.create(
                model=self.llm_model,
                messages=[{"role": "user", "content": prompt}],
                temperature=0.3,
                max_tokens=100
            )
            
            result = json.loads(response.choices[0].message.content)
            return result.get('category', ''), result.get('tags', [])
            
        except Exception as e:
            logger.error(f"LLM categorization error: {e}")
            return "", []


class WindowsLinkCreator:
    """Create Windows shortcut (.lnk) files"""
    
    @staticmethod
    def create_shortcut(target_path: str, shortcut_path: str):
        """Create Windows shortcut"""
        try:
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(shortcut_path)
            shortcut.Targetpath = target_path
            shortcut.save()
            logger.info(f"Created shortcut: {shortcut_path} -> {target_path}")
        except Exception as e:
            logger.error(f"Error creating shortcut: {e}")


class ExcelReporter:
    """Generate INDEX.xlsx report"""
    
    @staticmethod
    def generate_report(files: List[FileInfo], output_path: str, target_root: str):
        """Generate Excel report with metadata"""
        data = []
        
        for file_info in files:
            # Create hyperlink path
            relative_path = os.path.relpath(
                os.path.join(target_root, file_info.abbreviation, os.path.basename(file_info.path)),
                os.path.dirname(output_path)
            )
            
            data.append({
                'Abbrev': file_info.abbreviation,
                'File_Name': os.path.basename(file_info.path),
                'Revision': file_info.revision,
                'Category': file_info.category,
                'Tags': ', '.join(file_info.tags),
                'Size_MB': round(file_info.size_mb, 2),
                'Modified': file_info.modified,
                'SHA256': file_info.sha256,
                'Link': f'=HYPERLINK("{relative_path}", "Open")'
            })
        
        df = pd.DataFrame(data)
        
        # Write to Excel with formatting
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Specifications')
            
            # Get worksheet
            worksheet = writer.sheets['Specifications']
            
            # Add filters
            worksheet.auto_filter.ref = worksheet.dimensions
            
            # Adjust column widths
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                
                adjusted_width = min(max_length + 2, 50)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        logger.info(f"Generated Excel report: {output_path}")


class SafetyManager:
    """Handle dumps and restoration"""
    
    def __init__(self, db: SpecDatabase):
        self.db = db
    
    def create_dump(self, dump_dir: str = "dumps") -> str:
        """Create safety dump"""
        os.makedirs(dump_dir, exist_ok=True)
        
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        dump_file = os.path.join(dump_dir, f"dump_{timestamp}.json")
        
        dump_data = {
            'timestamp': timestamp,
            'files': self.db.get_all_files(),
            'version': '1.0'
        }
        
        with open(dump_file, 'w', encoding='utf-8') as f:
            json.dump(dump_data, f, indent=2, ensure_ascii=False)
        
        logger.info(f"Created dump: {dump_file}")
        return dump_file
    
    def restore_from_dump(self, dump_file: str) -> bool:
        """Restore system from dump"""
        try:
            with open(dump_file, 'r', encoding='utf-8') as f:
                dump_data = json.load(f)
            
            # Clear current database
            self.db.conn.execute("DELETE FROM files")
            self.db.conn.execute("DELETE FROM duplicates")
            
            # Restore files
            for file_data in dump_data['files']:
                # Convert back to FileInfo
                file_info = FileInfo(
                    path=file_data['original_path'],
                    abbreviation=file_data['abbreviation'],
                    sha256=file_data['sha256'],
                    size_mb=file_data['size_mb'],
                    modified=file_data['modified'],
                    category=file_data.get('category', ''),
                    revision=file_data.get('revision', ''),
                    tags=json.loads(file_data.get('tags', '[]'))
                )
                
                self.db.add_file(file_info, file_data['target_path'])
            
            logger.info(f"Restored from dump: {dump_file}")
            return True
            
        except Exception as e:
            logger.error(f"Error restoring from dump: {e}")
            return False


class EmailNotifier:
    """Send email notifications for critical errors"""
    
    @staticmethod
    def send_error_notification(error_msg: str):
        """Send error notification via email"""
        smtp_server = os.getenv('SMTP_SERVER')
        smtp_port = int(os.getenv('SMTP_PORT', 587))
        smtp_user = os.getenv('SMTP_USER')
        smtp_password = os.getenv('SMTP_PASSWORD')
        recipient = os.getenv('ERROR_EMAIL_RECIPIENT')
        
        if not all([smtp_server, smtp_user, smtp_password, recipient]):
            logger.warning("Email configuration incomplete, skipping notification")
            return
        
        try:
            msg = MIMEMultipart()
            msg['From'] = smtp_user
            msg['To'] = recipient
            msg['Subject'] = 'Spec Sync - Critical Error'
            
            body = f"""
            Critical error occurred in Spec Sync:
            
            {error_msg}
            
            Timestamp: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}
            """
            
            msg.attach(MIMEText(body, 'plain'))
            
            with smtplib.SMTP(smtp_server, smtp_port) as server:
                server.starttls()
                server.login(smtp_user, smtp_password)
                server.send_message(msg)
            
            logger.info("Error notification sent via email")
            
        except Exception as e:
            logger.error(f"Failed to send email notification: {e}")


class SyncWorker(QThread):
    """Background worker thread for sync operation"""
    
    progress = Signal(int)
    status = Signal(str)
    log = Signal(str)
    finished = Signal(dict)
    error = Signal(str)
    
    def __init__(self, config: Dict):
        super().__init__()
        self.config = config
        self.is_cancelled = False
        self.db = None
    
    def run(self):
        """Main sync process"""
        try:
            self.db = SpecDatabase()
            safety_manager = SafetyManager(self.db)
            
            # Create initial dump if not dry-run
            if not self.config['dry_run']:
                safety_manager.create_dump()
            
            # Initialize components
            extractor = AbbreviationExtractor(
                self.config['abbrev_file'],
                case_sensitive=self.config.get('case_sensitive', True)
            )
            categorizer = Categorizer(
                self.config['rules_file'],
                use_llm=self.config['use_llm'],
                llm_model=self.config['llm_model']
            )
            
            # Check existing specification journal
            existing_specs = self._load_existing_specs()
            
            # Scan files
            all_files = []
            for source_root in self.config['source_roots']:
                if self.is_cancelled:
                    break
                
                self.status.emit(f"Scanning {source_root}...")
                files = self._scan_directory(source_root, extractor)
                all_files.extend(files)
            
            self.log.emit(f"Found {len(all_files)} files with valid abbreviations")
            
            # Process files
            unique_files = []
            duplicate_count = 0
            error_count = 0
            
            for i, file_info in enumerate(all_files):
                if self.is_cancelled:
                    break
                
                self.progress.emit(int((i + 1) / len(all_files) * 100))
                self.status.emit(f"Processing {os.path.basename(file_info.path)}...")
                
                try:
                    # Check if duplicate
                    if self.db.file_exists(file_info.sha256):
                        file_info.is_duplicate = True
                        duplicate_count += 1
                        
                        if self.config['mode'] == 'move_link' and not self.config['dry_run']:
                            # Create link in original location
                            target_file = self._get_target_path(file_info)
                            link_path = file_info.path + ".lnk"
                            WindowsLinkCreator.create_shortcut(target_file, link_path)
                            self.db.add_duplicate(file_info.sha256, file_info.path, link_path)
                    else:
                        # Check against existing journal
                        if self._is_in_existing_journal(file_info, existing_specs):
                            self.log.emit(f"File already in specification journal: {file_info.path}")
                            if self.config['skip_existing']:
                                continue
                        
                        # Categorize
                        file_info.category, file_info.tags = categorizer.categorize(file_info)
                        
                        # Extract revision
                        file_info.revision = RevisionExtractor.extract(os.path.basename(file_info.path))
                        
                        # Process file
                        if not self.config['dry_run']:
                            target_path = self._process_file(file_info)
                            self.db.add_file(file_info, target_path)
                        
                        unique_files.append(file_info)
                
                except sqlite3.OperationalError as e:
                    # Database error - recreate tables and retry
                    logger.warning(f"Database error, recreating tables: {e}")
                    self.db._create_tables()
                    # Retry the operation
                    try:
                        if not self.db.file_exists(file_info.sha256):
                            # Process as new file
                            file_info.category, file_info.tags = categorizer.categorize(file_info)
                            file_info.revision = RevisionExtractor.extract(os.path.basename(file_info.path))
                            
                            if not self.config['dry_run']:
                                target_path = self._process_file(file_info)
                                self.db.add_file(file_info, target_path)
                            
                            unique_files.append(file_info)
                    except Exception as e2:
                        error_count += 1
                        self.log.emit(f"Error processing {file_info.path} after retry: {e2}")
                        logger.error(f"Error processing {file_info.path} after retry: {e2}")
                        
                except Exception as e:
                    error_count += 1
                    self.log.emit(f"Error processing {file_info.path}: {e}")
                    logger.error(f"Error processing {file_info.path}: {e}")
            
            # Generate report
            if unique_files and not self.config['dry_run']:
                report_path = os.path.join(self.config['target_root'], 'INDEX.xlsx')
                ExcelReporter.generate_report(unique_files, report_path, self.config['target_root'])
            
            # Final statistics
            stats = {
                'total_scanned': len(all_files),
                'unique': len(unique_files),
                'duplicates': duplicate_count,
                'errors': error_count
            }
            
            self.finished.emit(stats)
            
        except Exception as e:
            error_msg = f"Critical error: {traceback.format_exc()}"
            self.error.emit(error_msg)
            logger.critical(error_msg)
            
            # Send email notification
            EmailNotifier.send_error_notification(error_msg)
        
        finally:
            if self.db:
                self.db.close()
    
    def _scan_directory(self, root_path: str, extractor: AbbreviationExtractor) -> List[FileInfo]:
        """Scan directory for files with valid abbreviations"""
        files = []
        
        # Extended list of supported file formats
        supported_extensions = (
            # Documents
            '.pdf', '.doc', '.docx', '.xls', '.xlsx', '.ppt', '.pptx',
            '.odt', '.ods', '.odp', '.rtf', '.txt',
            # Images
            '.tiff', '.tif', '.jpg', '.jpeg', '.png', '.bmp', '.gif',
            '.webp', '.svg', '.ico', '.heic', '.heif',
            # CAD/Technical
            '.dwg', '.dxf', '.step', '.stp', '.iges', '.igs',
            '.stl', '.obj', '.3ds', '.dae',
            # Archives (for future expansion)
            '.zip', '.rar', '.7z', '.tar', '.gz',
            # Other technical formats
            '.xml', '.json', '.yaml', '.yml'
        )
        
        for dirpath, _, filenames in os.walk(root_path):
            for filename in filenames:
                if filename.lower().endswith(supported_extensions):
                    file_path = os.path.join(dirpath, filename)
                    
                    # Extract abbreviation
                    abbrev = extractor.extract(filename)
                    if not abbrev:
                        continue
                    
                    # Get file info
                    try:
                        stat = os.stat(file_path)
                        sha256 = FileHasher.calculate_sha256(file_path)
                        
                        file_info = FileInfo(
                            path=file_path,
                            abbreviation=abbrev,
                            sha256=sha256,
                            size_mb=stat.st_size / (1024 * 1024),
                            modified=datetime.fromtimestamp(stat.st_mtime).strftime('%Y-%m-%d %H:%M:%S')
                        )
                        
                        files.append(file_info)
                        
                    except Exception as e:
                        self.log.emit(f"Error scanning {file_path}: {e}")
        
        return files
    
    def _load_existing_specs(self) -> pd.DataFrame:
        """Load existing specification journal"""
        if self.config.get('existing_journal'):
            try:
                return pd.read_excel(self.config['existing_journal'])
            except Exception as e:
                self.log.emit(f"Error loading existing journal: {e}")
        return pd.DataFrame()
    
    def _is_in_existing_journal(self, file_info: FileInfo, journal: pd.DataFrame) -> bool:
        """Check if file exists in journal"""
        if journal.empty:
            return False
        
        # Check by SHA256 if column exists
        if 'SHA256' in journal.columns:
            return file_info.sha256 in journal['SHA256'].values
        
        # Otherwise check by filename
        filename = os.path.basename(file_info.path)
        if 'File_Name' in journal.columns:
            return filename in journal['File_Name'].values
        
        return False
    
    def _get_target_path(self, file_info: FileInfo) -> str:
        """Get target path for file"""
        target_dir = os.path.join(self.config['target_root'], file_info.abbreviation)
        os.makedirs(target_dir, exist_ok=True)
        return os.path.join(target_dir, os.path.basename(file_info.path))
    
    def _process_file(self, file_info: FileInfo) -> str:
        """Process file according to mode"""
        target_path = self._get_target_path(file_info)
        
        if self.config['mode'] == 'copy':
            shutil.copy2(file_info.path, target_path)
        elif self.config['mode'] == 'move_link':
            shutil.move(file_info.path, target_path)
        
        return target_path
    
    def cancel(self):
        """Cancel the sync operation"""
        self.is_cancelled = True


class SpecSyncGUI(QMainWindow):
    """Main GUI application"""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.init_ui()
    
    def init_ui(self):
        """Initialize user interface"""
        self.setWindowTitle("Spec Sync v1.1.0 - Aviation Specification Organizer")
        self.setGeometry(100, 100, 1000, 800)
        
        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Main layout
        main_layout = QVBoxLayout(central_widget)
        
        # Create tabs
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        # Configuration tab
        config_tab = QWidget()
        config_layout = QVBoxLayout(config_tab)
        tabs.addTab(config_tab, "Configuration")
        
        # Source roots
        source_group = QGroupBox("Source Directories")
        source_layout = QVBoxLayout()
        
        self.source_list = QListWidget()
        self.source_list.setSelectionMode(QListWidget.MultiSelection)
        source_layout.addWidget(self.source_list)
        
        source_buttons = QHBoxLayout()
        add_source_btn = QPushButton("Add Directory")
        add_source_btn.clicked.connect(self.add_source_directory)
        remove_source_btn = QPushButton("Remove Selected")
        remove_source_btn.clicked.connect(self.remove_source_directory)
        source_buttons.addWidget(add_source_btn)
        source_buttons.addWidget(remove_source_btn)
        source_layout.addLayout(source_buttons)
        
        source_group.setLayout(source_layout)
        config_layout.addWidget(source_group)
        
        # Target root
        target_group = QGroupBox("Target Directory")
        target_layout = QHBoxLayout()
        
        self.target_input = QLineEdit()
        target_browse_btn = QPushButton("Browse")
        target_browse_btn.clicked.connect(self.browse_target_directory)
        target_layout.addWidget(self.target_input)
        target_layout.addWidget(target_browse_btn)
        
        target_group.setLayout(target_layout)
        config_layout.addWidget(target_group)
        
        # Abbreviations file
        abbrev_group = QGroupBox("Abbreviations File")
        abbrev_layout = QHBoxLayout()
        
        self.abbrev_input = QLineEdit()
        abbrev_browse_btn = QPushButton("Browse")
        abbrev_browse_btn.clicked.connect(self.browse_abbrev_file)
        abbrev_layout.addWidget(self.abbrev_input)
        abbrev_layout.addWidget(abbrev_browse_btn)
        
        abbrev_group.setLayout(abbrev_layout)
        config_layout.addWidget(abbrev_group)
        
        # Mode selection
        mode_group = QGroupBox("Processing Mode")
        mode_layout = QHBoxLayout()
        
        self.mode_group = QButtonGroup()
        copy_radio = QRadioButton("Copy Mode")
        copy_radio.setChecked(True)
        move_link_radio = QRadioButton("Move & Link Mode")
        
        self.mode_group.addButton(copy_radio, 0)
        self.mode_group.addButton(move_link_radio, 1)
        
        mode_layout.addWidget(copy_radio)
        mode_layout.addWidget(move_link_radio)
        
        mode_group.setLayout(mode_layout)
        config_layout.addWidget(mode_group)
        
        # Options
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        
        self.use_llm_check = QCheckBox("Use LLM fallback for categorization")
        self.dry_run_check = QCheckBox("Dry Run (simulate without changes)")
        self.skip_existing_check = QCheckBox("Skip files in existing journal")
        self.case_sensitive_check = QCheckBox("Case-sensitive abbreviation matching")
        self.case_sensitive_check.setChecked(True)  # Default to case-sensitive
        
        options_layout.addWidget(self.use_llm_check)
        options_layout.addWidget(self.dry_run_check)
        options_layout.addWidget(self.skip_existing_check)
        options_layout.addWidget(self.case_sensitive_check)
        
        # LLM settings
        llm_layout = QHBoxLayout()
        llm_layout.addWidget(QLabel("LLM Model:"))
        self.llm_combo = QComboBox()
        self.llm_combo.addItems(["gpt-4", "gpt-3.5-turbo"])
        llm_layout.addWidget(self.llm_combo)
        options_layout.addLayout(llm_layout)
        
        options_group.setLayout(options_layout)
        config_layout.addWidget(options_group)
        
        # Existing journal
        journal_group = QGroupBox("Existing Specification Journal (Optional)")
        journal_layout = QHBoxLayout()
        
        self.journal_input = QLineEdit()
        journal_browse_btn = QPushButton("Browse")
        journal_browse_btn.clicked.connect(self.browse_journal_file)
        journal_layout.addWidget(self.journal_input)
        journal_layout.addWidget(journal_browse_btn)
        
        journal_group.setLayout(journal_layout)
        config_layout.addWidget(journal_group)
        
        # Progress tab
        progress_tab = QWidget()
        progress_layout = QVBoxLayout(progress_tab)
        tabs.addTab(progress_tab, "Progress")
        
        # Progress bar
        self.progress_bar = QProgressBar()
        progress_layout.addWidget(self.progress_bar)
        
        # Status label
        self.status_label = QLabel("Ready")
        progress_layout.addWidget(self.status_label)
        
        # Log console
        self.log_console = QTextEdit()
        self.log_console.setReadOnly(True)
        self.log_console.setFont(QFont("Consolas", 9))
        progress_layout.addWidget(self.log_console)
        
        # Control buttons
        control_layout = QHBoxLayout()
        
        self.start_btn = QPushButton("Start Sync")
        self.start_btn.clicked.connect(self.start_sync)
        self.cancel_btn = QPushButton("Cancel")
        self.cancel_btn.clicked.connect(self.cancel_sync)
        self.cancel_btn.setEnabled(False)
        
        control_layout.addWidget(self.start_btn)
        control_layout.addWidget(self.cancel_btn)
        main_layout.addLayout(control_layout)
        
        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status_bar(0, 0, 0, 0)
        
        # Menu bar
        self.create_menu_bar()
        
        # Load default values
        self.load_defaults()
    
    def create_menu_bar(self):
        """Create menu bar"""
        menubar = self.menuBar()
        
        # Tools menu
        tools_menu = menubar.addMenu("Tools")
        
        # Safety submenu
        safety_menu = tools_menu.addMenu("Safety")
        
        create_dump_action = QAction("Create Dump", self)
        create_dump_action.triggered.connect(self.create_dump)
        safety_menu.addAction(create_dump_action)
        
        restore_dump_action = QAction("Restore from Dump", self)
        restore_dump_action.triggered.connect(self.restore_from_dump)
        safety_menu.addAction(restore_dump_action)
    
    def load_defaults(self):
        """Load default configuration values"""
        # Check for abbreviations.txt in current directory
        if os.path.exists("abbreviations.txt"):
            self.abbrev_input.setText("abbreviations.txt")
        
        # Set default target
        self.target_input.setText(r"\\fileserver\specs_sorted")
    
    def add_source_directory(self):
        """Add source directory"""
        directory = QFileDialog.getExistingDirectory(self, "Select Source Directory")
        if directory:
            self.source_list.addItem(directory)
    
    def remove_source_directory(self):
        """Remove selected source directories"""
        for item in self.source_list.selectedItems():
            self.source_list.takeItem(self.source_list.row(item))
    
    def browse_target_directory(self):
        """Browse for target directory"""
        directory = QFileDialog.getExistingDirectory(self, "Select Target Directory")
        if directory:
            self.target_input.setText(directory)
    
    def browse_abbrev_file(self):
        """Browse for abbreviations file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Abbreviations File", "", "Text Files (*.txt)"
        )
        if file_path:
            self.abbrev_input.setText(file_path)
    
    def browse_journal_file(self):
        """Browse for existing journal file"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Specification Journal", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.journal_input.setText(file_path)
    
    def start_sync(self):
        """Start synchronization process"""
        # Validate inputs
        if self.source_list.count() == 0:
            QMessageBox.warning(self, "Warning", "Please add at least one source directory")
            return
        
        if not self.target_input.text():
            QMessageBox.warning(self, "Warning", "Please specify target directory")
            return
        
        if not self.abbrev_input.text():
            QMessageBox.warning(self, "Warning", "Please specify abbreviations file")
            return
        
        # Check if abbreviations file exists
        if not os.path.exists(self.abbrev_input.text()):
            QMessageBox.warning(
                self, 
                "Warning", 
                f"Abbreviations file not found: {self.abbrev_input.text()}\n\n"
                "Please create this file with one abbreviation per line."
            )
            return
        
        # Check if rules.yaml exists
        if not os.path.exists('rules.yaml'):
            QMessageBox.warning(
                self,
                "Warning",
                "rules.yaml not found!\n\n"
                "Please ensure rules.yaml is in the same directory as the script."
            )
            return
        
        # Prepare configuration
        config = {
            'source_roots': [
                self.source_list.item(i).text() 
                for i in range(self.source_list.count())
            ],
            'target_root': self.target_input.text(),
            'abbrev_file': self.abbrev_input.text(),
            'rules_file': 'rules.yaml',
            'mode': 'copy' if self.mode_group.checkedId() == 0 else 'move_link',
            'use_llm': self.use_llm_check.isChecked(),
            'llm_model': self.llm_combo.currentText(),
            'dry_run': self.dry_run_check.isChecked(),
            'skip_existing': self.skip_existing_check.isChecked(),
            'case_sensitive': self.case_sensitive_check.isChecked(),
            'existing_journal': self.journal_input.text() if self.journal_input.text() else None
        }
        
        # Create worker thread
        self.worker = SyncWorker(config)
        self.worker.progress.connect(self.update_progress)
        self.worker.status.connect(self.update_status)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.sync_finished)
        self.worker.error.connect(self.sync_error)
        
        # Update UI
        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.log_console.clear()
        
        # Start worker
        self.worker.start()
    
    def cancel_sync(self):
        """Cancel synchronization"""
        if self.worker:
            self.worker.cancel()
            self.append_log("Cancelling sync operation...")
    
    def update_progress(self, value):
        """Update progress bar"""
        self.progress_bar.setValue(value)
    
    def update_status(self, status):
        """Update status label"""
        self.status_label.setText(status)
    
    def append_log(self, message):
        """Append message to log console"""
        self.log_console.append(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")
    
    def update_status_bar(self, scanned, unique, duplicates, errors):
        """Update status bar with statistics"""
        self.status_bar.showMessage(
            f"Scanned: {scanned} | Unique: {unique} | Duplicates: {duplicates} | Errors: {errors}"
        )
    
    def sync_finished(self, stats):
        """Handle sync completion"""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        
        self.update_status_bar(
            stats['total_scanned'],
            stats['unique'],
            stats['duplicates'],
            stats['errors']
        )
        
        QMessageBox.information(
            self,
            "Sync Complete",
            f"Synchronization completed!\n\n"
            f"Total scanned: {stats['total_scanned']}\n"
            f"Unique files: {stats['unique']}\n"
            f"Duplicates: {stats['duplicates']}\n"
            f"Errors: {stats['errors']}"
        )
    
    def sync_error(self, error_msg):
        """Handle sync error"""
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        
        QMessageBox.critical(self, "Error", f"Sync failed:\n\n{error_msg}")
    
    def create_dump(self):
        """Create safety dump"""
        try:
            db = SpecDatabase()
            safety_manager = SafetyManager(db)
            dump_file = safety_manager.create_dump()
            db.close()
            
            QMessageBox.information(
                self,
                "Dump Created",
                f"Safety dump created successfully:\n{dump_file}"
            )
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to create dump:\n{e}")
    
    def restore_from_dump(self):
        """Restore from safety dump"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Dump File", "dumps", "JSON Files (*.json)"
        )
        
        if file_path:
            reply = QMessageBox.question(
                self,
                "Confirm Restore",
                "Are you sure you want to restore from this dump?\n"
                "This will replace the current inventory.",
                QMessageBox.Yes | QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                try:
                    db = SpecDatabase()
                    safety_manager = SafetyManager(db)
                    success = safety_manager.restore_from_dump(file_path)
                    db.close()
                    
                    if success:
                        QMessageBox.information(
                            self,
                            "Restore Complete",
                            "Successfully restored from dump"
                        )
                    else:
                        QMessageBox.warning(
                            self,
                            "Restore Failed",
                            "Failed to restore from dump"
                        )
                except Exception as e:
                    QMessageBox.critical(self, "Error", f"Restore error:\n{e}")


def main():
    """Main entry point"""
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle('Fusion')
    
    # Create and show main window
    window = SpecSyncGUI()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()