"""
spaCy NER Trainer for Auth Form Extraction
Trains a Named Entity Recognition model to identify:
- Patient Name
- Auth #
- Date Approved
- Date Auth Expire
- Patient ID

Usage:
1. Label some PDFs using labeling_tool.py or manually create training data
2. Run: python ml_trainer.py --train training_data.json --output model
3. Model is saved and ready for inference
"""

import json
import random
import spacy
from pathlib import Path
import argparse
from typing import List, Tuple, Dict, Any


class AuthFormNERTrainer:
    """Train spaCy NER model for auth form data extraction."""
    
    # Entity labels for auth forms
    LABELS = [
        "PATIENT_NAME",
        "AUTH_NUM",
        "DATE_APPROVED",
        "DATE_EXPIRE",
        "PATIENT_ID"
    ]
    
    def __init__(self, model_name: str = "en_core_web_sm"):
        """Initialize trainer with spaCy model."""
        try:
            self.nlp = spacy.load(model_name)
        except OSError:
            print(f"Model {model_name} not found. Installing...")
            import os
            os.system(f"python -m spacy download {model_name}")
            self.nlp = spacy.load(model_name)
        
        # Create NER pipeline if not exists
        if "ner" not in self.nlp.pipe_names:
            ner = self.nlp.add_pipe("ner", last=True)
        else:
            ner = self.nlp.get_pipe("ner")
        
        # Add labels
        for label in self.LABELS:
            ner.add_label(label)
    
    @staticmethod
    def convert_to_spacy_format(labeled_data: List[Dict[str, Any]]) -> List[Tuple[str, Dict]]:
        """
        Convert from simple labeling format to spaCy training format.
        
        Input format:
        [
            {
                "text": "Participant's Name: Johnson, Mary Auth #: 12345678 Date Approved: 10/15/2025",
                "entities": [
                    {"label": "PATIENT_NAME", "start": 22, "end": 34},
                    {"label": "AUTH_NUM", "start": 48, "end": 56},
                    {"label": "DATE_APPROVED", "start": 73, "end": 83}
                ]
            }
        ]
        
        Output: spaCy training format
        [
            (text, {"entities": [(start, end, label), ...]})
        ]
        """
        training_data = []
        
        for doc in labeled_data:
            text = doc["text"]
            entities = []
            
            for ent in doc.get("entities", []):
                entities.append((ent["start"], ent["end"], ent["label"]))
            
            training_data.append((text, {"entities": entities}))
        
        return training_data
    
    def train(self, training_data: List[Tuple[str, Dict]], 
              iterations: int = 30, drop_rate: float = 0.5):
        """
        Train the NER model.
        
        Args:
            training_data: List of (text, annotations) tuples
            iterations: Number of training iterations
            drop_rate: Dropout rate for regularization
        """
        ner = self.nlp.get_pipe("ner")
        
        # Disable other pipes for efficiency
        other_pipes = [pipe for pipe in self.nlp.pipe_names if pipe != "ner"]
        with self.nlp.disable_pipes(*other_pipes):
            optimizer = self.nlp.create_optimizer()
            
            for iteration in range(iterations):
                random.shuffle(training_data)
                losses = {}
                
                for text, annotations in training_data:
                    doc = self.nlp.make_doc(text)
                    example = spacy.training.Example.from_dict(doc, annotations)
                    self.nlp.update([example], drop=drop_rate, sgd=optimizer, losses=losses)
                
                if (iteration + 1) % 10 == 0:
                    print(f"Iteration {iteration + 1}/{iterations}, Loss: {losses.get('ner', 0):.4f}")
    
    def evaluate(self, test_data: List[Tuple[str, Dict]]) -> Dict[str, float]:
        """Evaluate model on test data."""
        from spacy.scorer import Scorer
        
        scorer = Scorer()
        examples = []
        
        for text, annotations in test_data:
            doc = self.nlp.make_doc(text)
            example = spacy.training.Example.from_dict(doc, annotations)
            examples.append(example)
        
        scores = scorer.score(self.nlp.pipe([ex.reference for ex in examples]), examples)
        return scores
    
    def save_model(self, output_path: str):
        """Save trained model."""
        path = Path(output_path)
        self.nlp.to_disk(path)
        print(f"Model saved to {path}")
    
    def load_model(self, model_path: str):
        """Load a saved model."""
        self.nlp = spacy.load(model_path)
        print(f"Model loaded from {model_path}")


def create_sample_training_data() -> List[Dict[str, Any]]:
    """Create sample training data for demonstration."""
    samples = [
        {
            "text": "Participant's Name: Johnson, Mary Auth #: 25166372 Date Approved: 10/15/2025 Date Auth Expire: 10/16/2025 Participant ID: 123456789",
            "entities": [
                {"label": "PATIENT_NAME", "start": 22, "end": 34},
                {"label": "AUTH_NUM", "start": 49, "end": 57},
                {"label": "DATE_APPROVED", "start": 75, "end": 85},
                {"label": "DATE_EXPIRE", "start": 108, "end": 118},
                {"label": "PATIENT_ID", "start": 134, "end": 143}
            ]
        },
        {
            "text": "Member's Name: Smith, Robert Auth #: 25846304 Date Approved: 01/05/2025 Date Auth Expire: 11/07/2025 Member ID: 987654321",
            "entities": [
                {"label": "PATIENT_NAME", "start": 15, "end": 27},
                {"label": "AUTH_NUM", "start": 42, "end": 50},
                {"label": "DATE_APPROVED", "start": 68, "end": 78},
                {"label": "DATE_EXPIRE", "start": 101, "end": 111},
                {"label": "PATIENT_ID", "start": 122, "end": 131}
            ]
        },
    ]
    return samples


def main():
    """CLI for training."""
    parser = argparse.ArgumentParser(description="Train spaCy NER model for auth forms")
    parser.add_argument("--train", type=str, help="Path to training data JSON file")
    parser.add_argument("--output", type=str, default="auth_form_ner_model", 
                       help="Output directory for trained model")
    parser.add_argument("--iterations", type=int, default=30, help="Training iterations")
    parser.add_argument("--sample", action="store_true", help="Create and train on sample data")
    
    args = parser.parse_args()
    
    trainer = AuthFormNERTrainer()
    
    # Load training data
    if args.sample:
        print("Using sample training data...")
        training_data_raw = create_sample_training_data()
    elif args.train:
        print(f"Loading training data from {args.train}...")
        with open(args.train, 'r') as f:
            training_data_raw = json.load(f)
    else:
        print("No training data provided. Use --train <file> or --sample")
        return
    
    # Convert to spaCy format
    training_data = trainer.convert_to_spacy_format(training_data_raw)
    print(f"Loaded {len(training_data)} training examples")
    
    # Train
    print(f"Training for {args.iterations} iterations...")
    trainer.train(training_data, iterations=args.iterations)
    
    # Save
    trainer.save_model(args.output)
    print(f"Training complete! Model saved to {args.output}")


if __name__ == "__main__":
    main()
