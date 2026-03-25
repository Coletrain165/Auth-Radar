"""
PDF Labeling Tool for Creating Training Data
Helps manually label PDF text for training the spaCy NER model.

Usage:
    python labeling_tool.py --pdf sample.pdf
    - Displays extracted text
    - Allows you to mark entity spans
    - Generates spaCy training data JSON
"""

import json
import argparse
from pathlib import Path
from typing import List, Dict, Any, Optional
import pdfplumber
import re


class PDFLabeler:
    """Interactive tool for labeling PDF text."""
    
    ENTITIES = [
        "PATIENT_NAME",
        "AUTH_NUM",
        "DATE_APPROVED",
        "DATE_EXPIRE",
        "PATIENT_ID"
    ]
    
    def __init__(self, pdf_path: str):
        """Initialize with a PDF file."""
        self.pdf_path = Path(pdf_path)
        self.text = self.extract_text()
        self.entities = []
    
    def extract_text(self) -> str:
        """Extract text from PDF."""
        try:
            with pdfplumber.open(self.pdf_path) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""
                    text += "\n---PAGE BREAK---\n"
            return text
        except Exception as e:
            print(f"Error extracting PDF: {e}")
            return ""
    
    def display_text_with_positions(self):
        """Display text with character positions for easy reference."""
        lines = self.text.split('\n')
        pos = 0
        print("\nText with character positions:\n")
        print("=" * 80)
        
        for i, line in enumerate(lines[:50], 1):  # Show first 50 lines
            print(f"[{pos:5d}] {line}")
            pos += len(line) + 1
        
        if len(lines) > 50:
            print(f"\n... ({len(lines) - 50} more lines)")
        
        print("=" * 80)
        print(f"\nTotal text length: {len(self.text)} characters\n")
    
    def find_entity_in_text(self, entity_type: str, search_term: str) -> Optional[Dict[str, Any]]:
        """
        Find an entity by search term and return span info.
        
        Args:
            entity_type: Type of entity (PATIENT_NAME, AUTH_NUM, etc)
            search_term: Text to search for (e.g., "John Smith", "12345678")
        
        Returns:
            {"start": int, "end": int, "text": str} or None
        """
        # Case-insensitive search
        text_lower = self.text.lower()
        search_lower = search_term.lower()
        
        start = text_lower.find(search_lower)
        if start == -1:
            return None
        
        end = start + len(search_term)
        return {
            "label": entity_type,
            "start": start,
            "end": end,
            "text": self.text[start:end]
        }
    
    def add_entity(self, entity_type: str, start: int, end: int) -> bool:
        """
        Add an entity with character positions.
        
        Args:
            entity_type: Type of entity
            start: Start character position
            end: End character position
        
        Returns:
            True if valid, False otherwise
        """
        if start < 0 or end > len(self.text) or start >= end:
            print(f"Invalid range: {start}-{end} (text length: {len(self.text)})")
            return False
        
        entity = {
            "label": entity_type,
            "start": start,
            "end": end,
            "text": self.text[start:end]
        }
        
        # Check for overlaps
        for existing in self.entities:
            if not (end <= existing["start"] or start >= existing["end"]):
                print(f"Overlap with existing entity: {existing}")
                return False
        
        self.entities.append(entity)
        print(f"✓ Added {entity_type}: '{entity['text']}'")
        return True
    
    def add_entity_by_search(self, entity_type: str, search_term: str) -> bool:
        """Add entity by searching for text."""
        entity = self.find_entity_in_text(entity_type, search_term)
        if not entity:
            print(f"Could not find '{search_term}' in text")
            return False
        
        return self.add_entity(entity_type, entity["start"], entity["end"])
    
    def list_entities(self):
        """Show all labeled entities."""
        if not self.entities:
            print("No entities labeled yet.")
            return
        
        print("\nLabeled Entities:")
        print("-" * 80)
        for i, ent in enumerate(self.entities, 1):
            context = self.text[max(0, ent["start"]-20):min(len(self.text), ent["end"]+20)]
            context = context.replace('\n', ' ')
            print(f"{i}. {ent['label']:15} | '{ent['text']}'")
            print(f"   Position: {ent['start']}-{ent['end']} | Context: ...{context}...")
        print("-" * 80)
    
    def clear_entity(self, index: int) -> bool:
        """Remove an entity by index."""
        if 0 <= index < len(self.entities):
            removed = self.entities.pop(index)
            print(f"Removed: {removed['label']}")
            return True
        return False
    
    def export_training_data(self, output_path: str) -> bool:
        """Export labeled data in spaCy format."""
        training_example = {
            "text": self.text.strip(),
            "entities": [
                {
                    "label": ent["label"],
                    "start": ent["start"],
                    "end": ent["end"]
                }
                for ent in self.entities
            ]
        }
        
        # Load existing training data if file exists
        out_path = Path(output_path)
        if out_path.exists():
            with open(out_path, 'r') as f:
                all_data = json.load(f)
        else:
            all_data = []
        
        # Add new example
        all_data.append(training_example)
        
        # Save
        with open(out_path, 'w') as f:
            json.dump(all_data, f, indent=2)
        
        print(f"✓ Exported {len(all_data)} total examples to {output_path}")
        return True
    
    def interactive_label(self, output_file: str = "data/training_data.json"):
        """Interactive labeling session."""
        print(f"\n{'='*80}")
        print("PDF LABELING TOOL")
        print(f"{'='*80}")
        print(f"File: {self.pdf_path}")
        print(f"Text length: {len(self.text)} characters")
        print(f"\nCommands:")
        print("  search <entity_type> '<text>' - Search for text and label it")
        print("                                   Entity types: PATIENT_NAME, AUTH_NUM, DATE_APPROVED, DATE_EXPIRE, PATIENT_ID")
        print("  add <entity_type> <start> <end> - Add entity by position")
        print("  list                             - Show all labeled entities")
        print("  remove <index>                   - Remove entity by index")
        print("  show                             - Show text with positions")
        print("  export                           - Export to training data")
        print("  quit                             - Exit")
        print(f"{'='*80}\n")
        
        self.display_text_with_positions()
        
        while True:
            cmd = input("labeler> ").strip()
            
            if not cmd:
                continue
            
            parts = cmd.split(maxsplit=1)
            action = parts[0].lower()
            args = parts[1] if len(parts) > 1 else ""
            
            if action == "quit":
                print("Exiting...")
                break
            
            elif action == "show":
                self.display_text_with_positions()
            
            elif action == "list":
                self.list_entities()
            
            elif action == "search":
                # Parse: search PATIENT_NAME 'John Smith'
                match = re.match(r"(\w+)\s+['\"](.+?)['\"]", args)
                if match:
                    entity_type, search_term = match.groups()
                    self.add_entity_by_search(entity_type, search_term)
                else:
                    print("Usage: search <entity_type> '<text>'")
            
            elif action == "add":
                # Parse: add PATIENT_NAME 100 120
                parts_args = args.split()
                if len(parts_args) >= 3:
                    entity_type = parts_args[0]
                    try:
                        start = int(parts_args[1])
                        end = int(parts_args[2])
                        self.add_entity(entity_type, start, end)
                    except ValueError:
                        print("Usage: add <entity_type> <start_pos> <end_pos>")
                else:
                    print("Usage: add <entity_type> <start_pos> <end_pos>")
            
            elif action == "remove":
                try:
                    idx = int(args) - 1
                    if self.clear_entity(idx):
                        pass
                    else:
                        print(f"Invalid index")
                except ValueError:
                    print("Usage: remove <index>")
            
            elif action == "export":
                self.export_training_data(output_file)
            
            else:
                print(f"Unknown command: {action}")


def main():
    parser = argparse.ArgumentParser(description="Label PDF text for training data")
    parser.add_argument("--pdf", type=str, required=True, help="Path to PDF file")
    parser.add_argument("--output", type=str, default="data/training_data.json", 
                       help="Output JSON file for training data")
    
    args = parser.parse_args()
    
    if not Path(args.pdf).exists():
        print(f"Error: PDF file not found: {args.pdf}")
        return
    
    labeler = PDFLabeler(args.pdf)
    labeler.interactive_label(output_file=args.output)


if __name__ == "__main__":
    main()
