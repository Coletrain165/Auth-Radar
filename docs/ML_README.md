# Hybrid ML + Regex Auth PDF Extraction System

## Overview

This project has been enhanced with **machine learning** for better data extraction from authorization PDFs. The system uses a **hybrid approach**:

1. **Primary: spaCy NER (Named Entity Recognition)** - Learns patterns from labeled examples
2. **Fallback: Regex patterns** - Reliable pattern matching as backup
3. **Confidence scores** - Know how confident each extraction is

## What's New

### New Files Created

| File | Purpose |
|------|---------|
| `ml_trainer.py` | Train spaCy NER model from labeled data |
| `ml_extractor.py` | Extract fields using ML + confidence scoring |
| `labeling_tool.py` | Interactive tool to label PDFs for training |
| `ML_QUICKSTART.py` | Quick start guide |
| `auth_extractor.py` (modified) | Now uses ML extraction when model is available |

### How It Works

```
PDF Text
   ↓
┌─────────────────────────┐
│  Try ML Extraction      │  ← Uses trained spaCy model
│  (if model exists)      │     Returns: value + confidence
└─────────────────────────┘
   ↓ (if field found)
Return ML result
   ↓ (if field NOT found)
┌─────────────────────────┐
│  Try Smart Regex        │  ← Context-aware pattern matching
└─────────────────────────┘
   ↓ (if field found)
Return Regex result
   ↓ (if field NOT found)
┌─────────────────────────┐
│  Try Pattern Matching   │  ← Legacy regex patterns
└─────────────────────────┘
   ↓ (if field found)
Return Regex result
   ↓ (if field NOT found)
Return: NOT FOUND (confidence: 0.0)
```

## Quick Start (3 Steps)

### Step 1: Install ML Dependencies

```bash
pip install spacy
python -m spacy download en_core_web_sm
```

### Step 2: Train a Model (Quick - 30 seconds)

```bash
python ml_trainer.py --sample --output auth_form_ner_model
```

This trains on sample data. For production, see "Training with Real Data" below.

### Step 3: Run Extraction

```bash
python auth_extractor.py
```

The system will automatically:
- Load the trained model
- Extract with ML
- Fall back to regex if needed
- Show confidence scores

## Output Format

Extraction results now include confidence information:

```python
{
    "file": "auth_12345.pdf",
    "Patient Name": "Johnson, Mary",
    "Patient Name_confidence": 0.92,        # 0.0-1.0 confidence score
    "Patient Name_method": "ml",            # Where it came from: ml, regex_smart, regex_pattern, filename, none
    "Auth #": "25166372",
    "Auth #_confidence": 0.88,
    "Auth #_method": "ml",
    "Date Approved": "10/15/2025",
    "Date Approved_confidence": 0.85,
    "Date Approved_method": "regex_smart",  # Fell back to regex
    ...
}
```

## Training with Real Data

For best results, train on your actual PDFs:

### Option A: Interactive Labeling (Recommended)

```bash
# Step 1: Label a PDF
python labeling_tool.py --pdf "path/to/sample_auth.pdf" --output training_data.json
```

Interactive commands:
```
labeler> search PATIENT_NAME 'Johnson, Mary'
labeler> search AUTH_NUM '25166372'
labeler> search DATE_APPROVED '10/15/2025'
labeler> search DATE_EXPIRE '10/16/2025'
labeler> search PATIENT_ID '123456789'
labeler> list              # Review what you labeled
labeler> export            # Save to training_data.json
labeler> quit
```

Repeat for 10-20 representative PDFs to build a good training set.

### Option B: Programmatic Labeling

```python
import json
from labeling_tool import PDFLabeler

labeler = PDFLabeler("path/to/pdf.pdf")

# Add entities by search
labeler.add_entity_by_search("PATIENT_NAME", "Johnson, Mary")
labeler.add_entity_by_search("AUTH_NUM", "25166372")

# Or by character position
labeler.add_entity("DATE_APPROVED", 1250, 1260)

# Export
labeler.export_training_data("training_data.json")
```

### Step 2: Train on Your Data

```bash
python ml_trainer.py --train training_data.json --output auth_form_ner_model --iterations 50
```

### Step 3: Evaluate Results

Run extraction and check:
- Confidence scores (aim for > 0.85)
- Extraction methods (prefer "ml" over "regex")
- If confidence is low, label more examples

## Confidence Score Interpretation

| Confidence | Interpretation | Action |
|---|---|---|
| 0.90-1.00 | Highly confident, good to use | ✓ Use as-is |
| 0.75-0.90 | Confident, generally reliable | ✓ Usually OK |
| 0.50-0.75 | Moderate, may need review | ⚠️ Check results |
| 0.00-0.50 | Low confidence or regex fallback | ❌ Review carefully |

## Monitoring Performance

Check extraction methods in your results:

```python
# Count by method
methods = {}
for result in results:
    method = result.get("Patient Name_method", "unknown")
    methods[method] = methods.get(method, 0) + 1

print(methods)
# Output example: {'ml': 45, 'regex_smart': 12, 'regex_pattern': 3, 'filename': 2}
```

**Goal**: Most fields from "ml", minimal "regex_pattern"

## Improving the Model

If accuracy is low:

1. **Get more training data**
   - Label 20-50 PDFs instead of 10
   - Focus on problematic examples

2. **Balance entity types**
   - Each PDF should have all 5 entity types labeled
   - Don't just label one or two fields

3. **Retrain**
   ```bash
   python ml_trainer.py --train training_data.json --iterations 50
   ```

4. **Test incrementally**
   - Train with 10 examples → test
   - Add 10 more → retrain → test
   - Find the sweet spot

## Troubleshooting

### Q: "ML model not found" message

**A**: Create a model first:
```bash
python ml_trainer.py --sample --output auth_form_ner_model
# or
python ml_trainer.py --train training_data.json --output auth_form_ner_model
```

### Q: Confidence scores are too low

**A**: This is normal with sample data. Label real examples:
```bash
python labeling_tool.py --pdf "real_auth.pdf" --output training_data.json
python ml_trainer.py --train training_data.json --output auth_form_ner_model
```

### Q: Model not improving with more data

**A**: Check training data quality:
```bash
import json
with open("training_data.json") as f:
    data = json.load(f)
    print(f"Examples: {len(data)}")
    for ex in data[:3]:
        print(f"Entities: {len(ex['entities'])} - {[e['label'] for e in ex['entities']]}")
```

Each example should have 5 entities (PATIENT_NAME, AUTH_NUM, DATE_APPROVED, DATE_EXPIRE, PATIENT_ID).

### Q: Some fields always fail

**A**: They might be missing in training data:
1. Ensure labeled PDFs have all 5 fields
2. Check labeling tool sees the text correctly:
   ```bash
   python labeling_tool.py --pdf "auth.pdf"
   ```
   Then use `show` command to see all text

## API Reference

### MLExtractor Class

```python
from ml_extractor import MLExtractor

# Initialize
extractor = MLExtractor(model_path="auth_form_ner_model")

# Extract with confidence
results = extractor.extract_with_confidence(text)

# Get a specific field
value = extractor.get_best_value(results, "Patient Name", min_confidence=0.8)

# Human-readable summary
print(extractor.summarize_results(results))
```

### AuthFormNERTrainer Class

```python
from ml_trainer import AuthFormNERTrainer

# Create trainer
trainer = AuthFormNERTrainer()

# Load data
with open("training_data.json") as f:
    data = json.load(f)

# Convert and train
training_data = trainer.convert_to_spacy_format(data)
trainer.train(training_data, iterations=50)

# Save model
trainer.save_model("auth_form_ner_model")
```

## File Structure

```
c:\Auth Radar\
├── auth_extractor.py              # Main extraction tool (now with ML)
├── ml_trainer.py                  # Model training script
├── ml_extractor.py                # ML extraction logic
├── labeling_tool.py               # PDF labeling utility
├── ML_QUICKSTART.py               # Quick start guide
├── auth_form_ner_model/           # Trained model (created after training)
├── training_data.json             # Training data (created by labeling tool)
├── patient_names.json             # Patient database
├── Extract_Auths.bat              # Windows batch runner
└── ...other files...
```

## Next Steps

1. **For Testing**: Use sample model
   ```bash
   python ml_trainer.py --sample --output auth_form_ner_model
   ```

2. **For Production**: Label real data and train
   ```bash
   # Label 20 PDFs
   python labeling_tool.py --pdf "pdf1.pdf" --output training_data.json
   python labeling_tool.py --pdf "pdf2.pdf" --output training_data.json
   # ... repeat for more PDFs
   
   # Train production model
   python ml_trainer.py --train training_data.json --output auth_form_ner_model --iterations 50
   ```

3. **Monitor Results**: Check confidence scores and methods in output

4. **Iterate**: As you get new PDFs, add them to training data and retrain

## Performance Notes

- **Training time**: ~1 minute for 30 iterations with 20 examples
- **Inference time**: <100ms per PDF page
- **Model size**: ~100-200 MB on disk
- **Memory**: ~500 MB during training

## Limitations

- ML model works best with consistent PDF formats
- First 10 training examples are crucial
- Low training data = lower confidence
- Scanned/OCR text may need more examples

## Support

For issues with:
- **PDF extraction**: Check `raw_text_preview` in results
- **ML accuracy**: Verify training data with `labeling_tool.py --pdf pdf.pdf`
- **Confidence scores**: Ensure model exists and is in correct location

## License

Same as original project.
