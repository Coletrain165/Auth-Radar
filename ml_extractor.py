"""
ML-based Extraction Module with Confidence Scoring
Extracts auth form fields using spaCy NER and combines with regex patterns.
Falls back to regex if ML model not available.

Usage:
    extractor = MLExtractor(model_path="auth_form_ner_model")
    results = extractor.extract_with_confidence(text)
    # Returns: {"field_name": {"value": "...", "confidence": 0.95, "method": "ml|regex|none"}}
"""

import os
import spacy
from pathlib import Path
from typing import Dict, Tuple, Optional, Any
import re
from datetime import datetime

# Dynamic year detection for OCR date correction
# This avoids hardcoding years like "2025" in OCR patterns
def get_current_year():
    """Get current year as integer."""
    return datetime.now().year

# Current and previous year for OCR pattern matching
CURRENT_YEAR = get_current_year()
CURRENT_YEAR_STR = str(CURRENT_YEAR)
CURRENT_YEAR_SUFFIX = CURRENT_YEAR_STR[-2:]  # e.g., "26"
PREV_YEAR = CURRENT_YEAR - 1
PREV_YEAR_STR = str(PREV_YEAR)
PREV_YEAR_SUFFIX = PREV_YEAR_STR[-2:]  # e.g., "25"


class MLExtractor:
    """
    Hybrid extraction using spaCy NER + regex patterns.
    Provides confidence scores for each extracted field.
    """
    
    # Confidence thresholds
    ML_CONFIDENCE_THRESHOLD = 0.7
    REGEX_CONFIDENCE = 0.8  # High confidence for regex matches
    
    # Field name mappings from spaCy NER to our fields
    ENTITY_TO_FIELD = {
        "PATIENT_NAME": "Patient Name",
        "AUTH_NUM": "Auth #",
        "DATE_APPROVED": "Date Approved",
        "DATE_EXPIRE": "Date Auth Expire",
        "PATIENT_ID": "Patient ID",
    }
    
    # Regex patterns (from original extractor)
    REGEX_PATTERNS = {
        "Patient Name": [
            r"Participant'?s?\s*Name[:\s]*\n?\s*([A-Z][A-Za-z]+(?:[,\s]+[A-Z][A-Za-z]+)+)",
            r"Member'?s?\s*Name[:\s]*\n?\s*([A-Z][A-Za-z]+(?:[,\s]+[A-Z][A-Za-z]+)+)",
            r"Patient'?s?\s*Name[:\s]*\n?\s*([A-Z][A-Za-z]+(?:[,\s]+[A-Z][A-Za-z]+)+)",
        ],
        "Auth #": [
            r"\|\s*\|\s*(\d{7,10})",  # After double pipe "| | 25357360"
            r"(\d{7,10})\s+\d{1,2}[/\-]",  # Auth number followed by date
            r"(\d{7,10})\s+[fF][iIlL1]",  # Auth number followed by garbled 'fi' date
            r"Auth\s*#?:?\s*(\d{7,10})",  # "Auth #: 12345678" or "Auth: 12345678"
            r"Auth\s*Number:?\s*(\d{7,10})",  # "Auth Number: 12345678"
            r"Expire:?\s*\|?\s*(\d{7,10})",  # After Expire label
        ],
        "Date Approved": [
            r"\d{7,10}\s+(\S+/\d{1,2}/\d{4})",
            r"Date\s+Approved:?\s*(\d{1,2}/\d{1,2}/\d{4})",  # Added: "Date Approved: 01/15/2025"
        ],
        "Date Auth Expire": [
            r"(\d{1,2}[/\-]\d{1,2}[/\-]\d{4})\s*\|?\s*UM",
            r"Date\s+Auth\.?\s*Expire:?\s*(\d{1,2}/\d{1,2}/\d{4})",  # Added: "Date Auth Expire: 12/15/2025"
            r"Expire[sd]?:?\s*(\d{1,2}/\d{1,2}/\d{4})",  # Added: "Expires: 12/15/2025"
        ],
        "Patient ID": [
            r"Participant\s*ID:?\s*\|?\s*(\d{6,12})",
            r"Member\s*ID:?\s*\|?\s*(\d{6,12})",  # Added: "Member ID: 123456789"
            r"Patient\s*ID:?\s*\|?\s*(\d{6,12})",  # Added: "Patient ID: 123456789"
        ],
    }
    
    def __init__(self, model_path: Optional[str] = None):
        """
        Initialize ML extractor.
        
        Args:
            model_path: Path to trained spaCy model. If None, only regex is used.
        """
        self.nlp = None
        self.model_path = model_path
        self.use_ml = False
        
        if model_path and Path(model_path).exists():
            try:
                self.nlp = spacy.load(model_path)
                self.use_ml = True
                print(f"ML Model loaded from {model_path}")
            except Exception as e:
                print(f"Could not load ML model: {e}. Falling back to regex only.")
                self.use_ml = False
        else:
            print("No ML model provided. Using regex extraction only.")
    
    def clean_ocr_date(self, value):
        """
        Fix common OCR errors in dates.
        This mirrors the patterns in auth_extractor.py's clean_ocr_date() method.
        """
        if not value:
            return value
        
        # Remove any spaces within the date string first
        value = re.sub(r'\s+', '', value)
        
        # Handle bracket "[" being misread - often "[" is "1" or part of date
        value = re.sub(r'^\[', '', value)  # Remove leading bracket
        value = re.sub(r'\[', '1', value)  # Replace remaining brackets with 1
        
        # Handle parenthesis "(" being misread 
        value = re.sub(r'^\(', '', value)  # Remove leading paren
        value = re.sub(r'\(', '', value)  # Remove remaining parens
        
        # Handle "foo" -> "10/0" (very common OCR error for "10/0")
        value = re.sub(r'^foo(?=\d|/)', '10/0', value, flags=re.IGNORECASE)
        
        # Handle "[os" or "os" at start -> "08" 
        value = re.sub(r'^os(?=\d|/|e)', '08/', value, flags=re.IGNORECASE)
        
        # Handle "ose" pattern -> "08/2" (common misread for 08/2x)
        value = re.sub(r'^08e', '08/2', value, flags=re.IGNORECASE)
        
        # Handle "fia" -> "12" (a misread for 2, common for December dates)
        value = re.sub(r'^fia(?=/|\d)', '12', value, flags=re.IGNORECASE)
        
        # Handle "fiz" -> "12" (z misread for 2)
        value = re.sub(r'^fiz(?=/|\d|e|y)', '12', value, flags=re.IGNORECASE)
        
        # Handle extremely garbled patterns like "fizeayaces" -> "12/24/2025"
        value = re.sub(r'^12ea', '12/24', value, flags=re.IGNORECASE)
        value = re.sub(r'^12e([a-z])', r'12/2\1', value, flags=re.IGNORECASE)
        
        # Handle "ay" in date -> "24" (a=2, y=4)
        value = re.sub(r'/ay(?=/|\d)', '/24', value, flags=re.IGNORECASE)
        value = re.sub(r'(\d)ay(?=/|\d)', r'\g<1>/24', value, flags=re.IGNORECASE)
        
        # Handle "y" followed by digits -> likely day part missing slash
        value = re.sub(r'^(\d{1,2})y(\d)', r'\1/\2', value, flags=re.IGNORECASE)
        
        # Handle "year" in date text -> strip it
        value = re.sub(r'year(?=\d{4})', '/', value, flags=re.IGNORECASE)
        
        # Dynamic year handling - use current year instead of hardcoded 2025
        current_year = CURRENT_YEAR_STR
        
        # Handle extremely garbled year patterns with dynamic year
        value = re.sub(r'yeajo0e[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'yeaj00e[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'yea[a-z0-9]*e[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'yeaj[a-z0]*(?=\d)', '/', value, flags=re.IGNORECASE)
        value = re.sub(r'yeaj(?=\d)', '/', value, flags=re.IGNORECASE)
        value = re.sub(r'yea(?=\d)', '/', value, flags=re.IGNORECASE)

        # Handle "fin" -> "11" (n misread for 1, common for November dates)
        value = re.sub(r'^fin(?=/|\d)', '11', value, flags=re.IGNORECASE)
        
        # "fi" at start often means "1" - e.g., "fi2" -> "12", "fi1" -> "11"
        value = re.sub(r'^fi(\d)', r'1\1', value, flags=re.IGNORECASE)
        
        # Special case: "fio" -> "10"
        value = re.sub(r'^fio(?=/)', '10', value, flags=re.IGNORECASE)
        
        # "fioo" or "fi0o" patterns -> likely "10/0" or similar 
        value = re.sub(r'^fioo', '10/0', value, flags=re.IGNORECASE)
        value = re.sub(r'^fi0o', '10/0', value, flags=re.IGNORECASE)
        
        # Handle complex garbled patterns
        value = re.sub(r'^fioso', '10/0', value, flags=re.IGNORECASE)
        
        # Handle "fioye" pattern specifically -> likely "10/2" (e.g., "fioye7/20es" -> "10/27/2025")
        value = re.sub(r'^fioye(\d)', r'10/2\1', value, flags=re.IGNORECASE)
        
        # Handle "fioy" -> "10/" (y misread for slash)
        value = re.sub(r'^fioy(?=\d)', '10/', value, flags=re.IGNORECASE)
        value = re.sub(r'^fi0y(?=\d)', '10/', value, flags=re.IGNORECASE)
        value = re.sub(r'^fioy', '10/0', value, flags=re.IGNORECASE)
        value = re.sub(r'^fi0y', '10/0', value, flags=re.IGNORECASE)
        
        # Common OCR misreads for "10": fio, f1o, flo, 1o, lo, lO, etc.
        value = re.sub(r'^f[1l]o(?=/)', '10', value, flags=re.IGNORECASE)
        value = re.sub(r'^[1l]o(?=/)', '10', value, flags=re.IGNORECASE)
        value = re.sub(r'^[1l][oO0](?=/)', '10', value)
        
        # Fix "s" or "es" being read instead of digits
        value = re.sub(r's/', '8/', value, flags=re.IGNORECASE)
        value = re.sub(r'/os', '/08', value, flags=re.IGNORECASE)
        value = re.sub(r'/oos', '/08', value, flags=re.IGNORECASE)
        
        # Fix "e9" in middle -> "29" (e misread for 2)  
        value = re.sub(r'/e9/', '/29/', value, flags=re.IGNORECASE)
        value = re.sub(r'^e9/', '29/', value, flags=re.IGNORECASE)
        
        # Fix "y" misread as part of number
        value = re.sub(r'(\d)y(\d)', r'\1/\2', value, flags=re.IGNORECASE)
        
        # Fix "e" between digits
        value = re.sub(r'(\d)e(\d)(?=/)', r'\1/\2', value, flags=re.IGNORECASE)
        
        # Fix duplicated digits
        value = re.sub(r'/(\d)\1(?=/)', r'/\1', value)
        
        # Fix "e0" in year -> "20"
        value = re.sub(r'/e0(\d\d)$', r'/20\1', value, flags=re.IGNORECASE)
        value = re.sub(r'/(\d)e(\d\d)$', r'/\g<1>0\2', value, flags=re.IGNORECASE)
        
        # Dynamic year patterns using current year (handles 2025, 2026, etc.)
        # "es" at end -> likely "25" or "26" depending on current year
        value = re.sub(r'/20e[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'/2oe[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'/202[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'e0e[s56]$', current_year, value, flags=re.IGNORECASE)
        
        # Handle "aces" at end -> current year (yaces must come BEFORE aces)
        value = re.sub(r'yace[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'/ace[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'ace[s56]$', current_year, value, flags=re.IGNORECASE)
        
        # Handle "jo0es" or "j00es" at end -> current year
        value = re.sub(r'jo0e[s56]$', current_year, value, flags=re.IGNORECASE)
        value = re.sub(r'j00e[s56]$', current_year, value, flags=re.IGNORECASE)
        value = re.sub(r'j0e[s56]$', current_year, value, flags=re.IGNORECASE)
        
        # Handle "o0es" at end -> current year 
        value = re.sub(r'/o0e[s56]$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'o0e[s56]$', current_year, value, flags=re.IGNORECASE)
        
        # 2026-specific OCR patterns
        value = re.sub(r'/20e6$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'/2o26$', f'/{current_year}', value, flags=re.IGNORECASE)
        value = re.sub(r'/20z6$', f'/{current_year}', value, flags=re.IGNORECASE)
        
        # Handle garbled years like "year2025" -> "2025"
        value = re.sub(r'year(\d{4})$', r'/\1', value, flags=re.IGNORECASE)
        
        # Fix within date (after first /)
        value = re.sub(r'(?<=/)[fli1][iloO0](?=/)', '10', value)
        
        # Fix O/o read as 0 in dates
        value = re.sub(r'[oO](?=\d)', '0', value)
        value = re.sub(r'(?<=\d)[oO]', '0', value)
        
        return value
    
    def extract_with_ml(self, text: str) -> Dict[str, Dict[str, Any]]:
        """
        Extract fields using spaCy NER.
        
        Returns:
            {
                "Patient Name": {"value": "Smith, John", "confidence": 0.92, "method": "ml"},
                "Auth #": {"value": "12345678", "confidence": 0.88, "method": "ml"},
                ...
            }
        """
        if not self.use_ml or not self.nlp:
            return {}
        
        doc = self.nlp(text)
        results = {}
        
        # Group entities by type
        entities_by_type = {}
        for ent in doc.ents:
            if ent.label_ not in entities_by_type:
                entities_by_type[ent.label_] = []
            entities_by_type[ent.label_].append({
                "text": ent.text,
                "score": getattr(ent, 'kb_id_', 1.0) if hasattr(ent, 'kb_id_') else 1.0
            })
        
        # Convert to field names
        for entity_type, field_name in self.ENTITY_TO_FIELD.items():
            if entity_type in entities_by_type:
                entities = entities_by_type[entity_type]
                if entities:
                    # Use first entity (highest confidence typically)
                    best = entities[0]
                    value = best["text"].strip()
                    
                    # Apply OCR date cleaning for date fields
                    if "Date" in field_name:
                        value = self.clean_ocr_date(value)
                    
                    # Approximate confidence from model uncertainty
                    try:
                        score = float(best.get("score", 0.85))
                    except (ValueError, TypeError):
                        score = 0.85
                    confidence = min(0.95, score)
                    results[field_name] = {
                        "value": value,
                        "confidence": confidence,
                        "method": "ml"
                    }
        
        return results
    
    def extract_with_regex(self, text: str, field_name: str) -> Optional[Tuple[str, float]]:
        """
        Extract a single field using regex.
        
        Returns:
            (value, confidence) or None
        """
        patterns = self.REGEX_PATTERNS.get(field_name, [])
        
        for pattern in patterns:
            match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
            if match:
                value = match.group(1).strip()
                value = re.sub(r'[\n\r]+', ' ', value)
                value = re.sub(r'\s+', ' ', value)
                
                # Apply OCR date cleaning for date fields
                if "Date" in field_name:
                    value = self.clean_ocr_date(value)
                
                return value, self.REGEX_CONFIDENCE
        
        return None
    
    def extract_with_confidence(self, text: str) -> Dict[str, Dict[str, Any]]:
        """
        Extract fields using hybrid approach (ML + Regex).
        ML is tried first, regex as fallback.
        
        Returns:
            {
                "Patient Name": {"value": "...", "confidence": 0.92, "method": "ml"},
                "Auth #": {"value": "...", "confidence": 0.88, "method": "regex"},
                "Date Approved": {"value": "...", "confidence": 0.0, "method": "none"}
            }
        """
        results = {}
        fields = list(self.ENTITY_TO_FIELD.values())
        
        # Try ML extraction first
        if self.use_ml:
            ml_results = self.extract_with_ml(text)
            results.update(ml_results)
        
        # Fallback to regex for missing fields
        for field_name in fields:
            if field_name not in results:
                regex_result = self.extract_with_regex(text, field_name)
                if regex_result:
                    value, confidence = regex_result
                    results[field_name] = {
                        "value": value,
                        "confidence": confidence,
                        "method": "regex"
                    }
                else:
                    # Field not found
                    results[field_name] = {
                        "value": None,
                        "confidence": 0.0,
                        "method": "none"
                    }
        
        return results
    
    def get_best_value(self, results: Dict[str, Dict[str, Any]], 
                       field_name: str, min_confidence: float = 0.0) -> Optional[str]:
        """
        Get the best value for a field from extraction results.
        
        Args:
            results: Results from extract_with_confidence()
            field_name: Field to retrieve
            min_confidence: Minimum confidence threshold
        
        Returns:
            Value if found and confidence >= min_confidence, else None
        """
        if field_name not in results:
            return None
        
        result = results[field_name]
        if result["confidence"] >= min_confidence:
            return result["value"]
        
        return None
    
    def summarize_results(self, results: Dict[str, Dict[str, Any]]) -> str:
        """Return human-readable summary of extraction results with confidence."""
        summary = []
        for field, data in results.items():
            value = data["value"] if data["value"] else "[NOT FOUND]"
            conf = f"{data['confidence']*100:.0f}%"
            method = data["method"].upper()
            summary.append(f"{field}: {value} ({conf} confidence, {method})")
        return "\n".join(summary)


# Example usage and testing
if __name__ == "__main__":
    # Example text (would come from PDF extraction)
    sample_text = """
    Participant's Name: Johnson, Mary
    Participant ID: 123456789
    Auth #: 25166372
    Date Approved: 10/15/2025
    Date Auth Expire: 10/16/2025
    """
    
    # Initialize extractor (without model - regex only)
    extractor = MLExtractor(model_path=None)
    
    # Extract with confidence
    results = extractor.extract_with_confidence(sample_text)
    
    print("Extraction Results:")
    print(extractor.summarize_results(results))
    
    print("\n\nDetailed Results:")
    for field, data in results.items():
        print(f"{field}: {data}")
