# Universal Data Extractor

Universal file processor that extracts structured JSON data from any file type using **Gemini 2.5 Pro**.

## Architecture

```
┌─────────────────┐    ┌──────────────────┐    ┌─────────────────┐
│   Input Files   │    │   File Processor │    │ Gemini 2.5 Flash│
│                 │    │                  │    │                 │
│ • PDF           │───▶│ • Type Detection │───▶│ • Native Support│
│ • Images        │    │ • MIME Analysis  │    │ • JSON Output   │
│ • Text/CSV/JSON │    │ • Smart Routing  │    │ • Thinking Mode │
│ • Excel (XLSX)  │    │                  │    │                 │
│ • Word (DOCX)   │    │                  │    │                 │
└─────────────────┘    └──────────────────┘    └─────────────────┘
                                │                        │
                                ▼                        ▼
                    ┌──────────────────┐    ┌─────────────────┐
                    │   PDF Converter  │    │  JSON Output    │
                    │                  │    │                 │
                    │ • Excel → PDF    │    │ • Structured    │
                    │ • Word → PDF     │    │ • Validated     │
                    │ • Auto Cleanup   │    │ • Clean Data    │
                    └──────────────────┘    └─────────────────┘
```

## Processing Flow

```
File Input → Type Detection → Processing Path → Gemini 2.5 Pro → JSON Output

Native Support (Fast):           Convert to PDF (When Needed):
├── PDF                         ├── Excel (.xlsx, .xls)
├── Images (.jpg, .png, etc.)   └── Word (.docx, .doc)  
└── Text (.txt, .csv, .json)    
```

## Installation

```bash
# Core dependency
pip install google-genai>=1.0.0

# Excel support (optional)
pip install aspose-cells-python>=24.1.0

# Word support (optional)
pip install reportlab>=4.0.0 docx2txt>=0.8 python-docx>=1.1.0

# All dependencies
pip install -r requirements.txt
```

## Configuration

```bash
# Set API key
export GOOGLE_API_KEY="your-api-key-here"

# Get API key from: https://aistudio.google.com/app/apikey
```

## Usage

### Basic Commands

```bash
# Single file
python main.py input/document.pdf

# Multiple files
python main.py input/data.xlsx input/report.pdf input/image.png

# Pattern matching
python main.py "input/*.*"

# Custom prompt
python main.py input/financial.xlsx --prompt "Extract revenue and expenses by quarter"

# Verbose output
python main.py input/file.pdf -v
```

### Python API

```python
from main import UniversalDataExtractor

# Initialize
extractor = UniversalDataExtractor(api_key="your-key")

# Process single file
result = extractor.process("input/document.pdf")

# Process multiple files
results = extractor.batch_process(["file1.pdf", "file2.xlsx"])
```

## File Support Matrix

| File Type | Extension | Native Support | Conversion | Performance |
|-----------|-----------|---------------|------------|-------------|
| PDF | `.pdf` | ✅ | None | ⚡ ~10s |
| Images | `.jpg`, `.png`, `.gif`, etc. | ✅ | None | ⚡ ~8s |
| Text | `.txt`, `.csv`, `.json`, `.md` | ✅ | None | ⚡ ~5s |
| Excel | `.xlsx`, `.xls`, `.xlsm` | ❌ | PDF | 🔄 ~30s |
| Word | `.docx`, `.doc` | ❌ | PDF | 🔄 ~25s |

## Output Structure

### JSON Schema
```json
{
  "metadata": {
    "source_file": "document.pdf",
    "file_type": "pdf", 
    "processed_at": "2025-01-XX"
  },
  "content": {
    "tables": [...],
    "paragraphs": [...],
    "key_data": {...}
  }
}
```

### Output Location
- **Directory**: `output/`
- **Format**: `{filename}_extracted.json`
- **Encoding**: UTF-8
- **Structure**: Pretty-printed JSON

## Technical Specifications

### Gemini 2.5 Pro Configuration
```python
config=types.GenerateContentConfig(
    response_mime_type='application/json',  # Force JSON
    temperature=0.1,                        # Consistent output
    thinking_config=types.ThinkingConfig(   # Enhanced reasoning
        thinking_budget=1024
    )
)
```

### Dependencies
- **Core**: `google-genai>=1.0.0`
- **Excel**: `aspose-cells-python>=24.1.0` 
- **Word**: `reportlab>=4.0.0`, `docx2txt>=0.8`, `python-docx>=1.1.0`
- **Optional**: `pydantic>=2.0.0`, `chardet>=5.0.0`

### Limits
- **File size**: No hard limit (API dependent)
- **Batch size**: Unlimited
- **Formats**: All common business file types
- **Rate limits**: Per Google API quotas

## Error Handling

### Common Issues

| Error | Cause | Solution |
|-------|-------|----------|
| `No API key found` | Missing environment variable | `export GOOGLE_API_KEY="key"` |
| `Excel processing disabled` | Missing dependency | `pip install aspose-cells-python` |
| `Word processing disabled` | Missing dependency | `pip install reportlab docx2txt` |
| `JSON parsing failed` | Invalid response | Check file format, retry |
| `File not found` | Invalid path | Verify file exists |

### Debug Mode
```bash
python main.py input/file.pdf -v
```

## Performance Benchmarks

### Processing Times (10MB files)
- **PDF**: 10-15 seconds
- **Image**: 8-12 seconds  
- **Text**: 5-8 seconds
- **Excel**: 25-35 seconds (includes conversion)
- **Word**: 20-30 seconds (includes conversion)

### Memory Usage
- **Native processing**: Low (< 100MB)
- **With conversion**: Medium (100-300MB)
- **Batch processing**: Linear scaling

## Project Structure

```
universal-extractor/
├── main.py                 # Main processor
├── excel_to_pdf.py         # Excel conversion
├── word_to_pdf.py          # Word conversion  
├── text_to_pdf.py          # Text conversion (fallback)
├── requirements.txt        # Dependencies
├── test_extractor.py       # Test suite
├── input/                  # Input files
├── output/                 # JSON outputs
└── temp/                   # Temporary files (auto-cleaned)
```

## Testing

```bash
# Run test suite
python test_extractor.py

# Test specific file types
python main.py input/test.txt input/test.json -v
```

## API Reference

### UniversalDataExtractor Class

```python
class UniversalDataExtractor:
    def __init__(self, api_key: str = None)
    def process(self, file_path: str, prompt: str = None) -> str | None
    def batch_process(self, file_paths: List[str], prompt: str = None) -> Dict
```

### Command Line Arguments

```
positional arguments:
  files              Files to process

options:
  -h, --help         Show help message
  -v, --verbose      Verbose logging
  --prompt PROMPT    Custom extraction prompt
  --api-key KEY      Gemini API key
```