import os
import pathlib
import logging
import sys
from typing import Optional, Dict, Any, List

# PDF creation dependencies
try:
    import chardet
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Preformatted, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
    logging.info("ReportLab available - Text to PDF conversion enabled")
except ImportError:
    REPORTLAB_AVAILABLE = False
    logging.warning("ReportLab not available. Install: pip install reportlab")
    chardet = None
    # Create proper mock classes with required methods
    class _SimpleDocTemplate:
        def __init__(self, filename, **kwargs): pass
        def build(self, story): pass
    class _Paragraph:
        def __init__(self, text, style): pass
    class _Preformatted:
        def __init__(self, text, style): pass
    class _Spacer:
        def __init__(self, width, height): pass
    class _ParagraphStyle:
        def __init__(self, name, **kwargs): pass
    class _StyleSheet(dict):
        def add(self, style): pass
    
    # Replace imports with mocks
    SimpleDocTemplate = _SimpleDocTemplate  # type: ignore
    Paragraph = _Paragraph  # type: ignore
    Preformatted = _Preformatted  # type: ignore
    Spacer = _Spacer  # type: ignore
    ParagraphStyle = _ParagraphStyle  # type: ignore
    def getSampleStyleSheet():  # type: ignore[override]
        return _StyleSheet()
    A4 = None  # type: ignore

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class TextToPdfService:
    """
    Text file to PDF conversion service using ReportLab.
    Note: This is mainly for edge cases - Gemini can process text files natively.
    """
    
    def __init__(self):
        """Initialize the Text to PDF service."""
        if not REPORTLAB_AVAILABLE:
            raise ImportError(
                "ReportLab is required for PDF creation. "
                "Install with: pip install reportlab"
            )
        
        self.supported_formats = {'.txt', '.csv', '.tsv', '.json', '.xml', '.html', '.md', '.rtf'}
    
    def detect_encoding(self, file_path: str) -> str:
        """Detect file encoding."""
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read()
                if chardet and hasattr(chardet, 'detect'):
                    result = chardet.detect(raw_data)
                    return result.get('encoding', 'utf-8') or 'utf-8'
                else:
                    return 'utf-8'
        except Exception as e:
            logger.warning(f"Could not detect encoding for {file_path}: {e}")
            return 'utf-8'
    
    def read_text_file(self, file_path: str) -> Optional[str]:
        """Read text file with encoding detection."""
        try:
            encoding = self.detect_encoding(file_path)
            with open(file_path, 'r', encoding=encoding) as f:
                content = f.read()
            return content
        except Exception as e:
            logger.error(f"Failed to read text file {file_path}: {e}")
            return None
    
    def convert_text_to_pdf(self, text_path: str, pdf_path: str, 
                           title: Optional[str] = None) -> bool:
        """
        Convert text file to PDF.
        
        Args:
            text_path: Path to input text file
            pdf_path: Path for output PDF file
            title: Custom title for the document
        
        Returns:
            bool: True if conversion successful, False otherwise
        """
        try:
            logger.info(f"Converting text to PDF: {text_path} -> {pdf_path}")
            
            # Read text content
            content = self.read_text_file(text_path)
            if content is None:
                return False
            
            # Create PDF document
            pdf_doc = SimpleDocTemplate(
                pdf_path,
                pagesize=A4,
                rightMargin=50,
                leftMargin=50,
                topMargin=50,
                bottomMargin=50
            )
            
            story = []
            styles = getSampleStyleSheet()
            
            # Add custom styles
            styles.add(ParagraphStyle(
                'TextNormal',
                parent=styles['Normal'],
                fontSize=10,
                leading=14,
                spaceAfter=6,
                fontName='Courier'  # Monospace for better text formatting
            ))
            
            # Add title
            doc_title = title or pathlib.Path(text_path).stem
            story.append(Paragraph(doc_title, styles['Title']))
            story.append(Spacer(1, 20))
            
            # Process content based on file type
            file_extension = pathlib.Path(text_path).suffix.lower()
            
            if file_extension in ['.json', '.xml']:
                # Format structured data
                story.append(Preformatted(content, styles['Code']))
            else:
                # Regular text processing
                paragraphs = content.split('\n\n')  # Split on double newlines
                
                for para in paragraphs:
                    para = para.strip()
                    if not para:
                        story.append(Spacer(1, 6))
                        continue
                    
                    # Handle different content types
                    if file_extension == '.md':
                        # Basic Markdown processing
                        if para.startswith('#'):
                            # Heading
                            level = len(para) - len(para.lstrip('#'))
                            heading_text = para.lstrip('#').strip()
                            if level == 1:
                                story.append(Paragraph(heading_text, styles['Heading1']))
                            elif level == 2:
                                story.append(Paragraph(heading_text, styles['Heading2']))
                            else:
                                story.append(Paragraph(heading_text, styles['Heading3']))
                        else:
                            story.append(Paragraph(para, styles['TextNormal']))
                    else:
                        # Regular text
                        story.append(Paragraph(para, styles['TextNormal']))
            
            # Build PDF
            pdf_doc.build(story)
            
            logger.info(f"Successfully converted text to PDF: {pdf_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to convert text to PDF: {e}")
            return False
    
    def get_text_metadata(self, text_path: str) -> Dict[str, Any]:
        """Extract metadata from text file."""
        try:
            file_stat = os.stat(text_path)
            file_size_mb = file_stat.st_size / (1024 * 1024)
            
            content = self.read_text_file(text_path)
            
            metadata = {
                "file_size_mb": round(file_size_mb, 3),
                "file_format": pathlib.Path(text_path).suffix.lower(),
                "encoding": self.detect_encoding(text_path)
            }
            
            if content:
                lines = content.split('\n')
                words = content.split()
                
                metadata.update({
                    "line_count": len(lines),
                    "word_count": len(words),
                    "character_count": len(content)
                })
                
                # Detect content type
                if pathlib.Path(text_path).suffix.lower() == '.csv':
                    # CSV analysis
                    if lines:
                        first_line = lines[0]
                        delimiter_count = {
                            ',': first_line.count(','),
                            ';': first_line.count(';'),
                            '\t': first_line.count('\t'),
                            '|': first_line.count('|')
                        }
                        likely_delimiter = max(delimiter_count.items(), key=lambda x: x[1])[0]
                        metadata.update({
                            "csv_columns": delimiter_count[likely_delimiter] + 1,
                            "csv_delimiter": likely_delimiter
                        })
            
            return metadata
            
        except Exception as e:
            logger.error(f"Failed to extract text metadata: {e}")
            return {"error": str(e)}


if __name__ == "__main__":
    """Command-line interface for text to PDF conversion."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Text to PDF Conversion Service")
    parser.add_argument("input", help="Input text file")
    parser.add_argument("-o", "--output", help="Output PDF file")
    parser.add_argument("--title", help="Custom document title")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging")
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        service = TextToPdfService()
        
        input_path = pathlib.Path(args.input)
        if not input_path.exists():
            print(f"❌ File not found: {args.input}")
            sys.exit(1)
        
        # Determine output path
        if args.output:
            output_path = args.output
        else:
            output_path = f"{input_path.stem}.pdf"
        
        print(f"Converting: {args.input} -> {output_path}")
        
        success = service.convert_text_to_pdf(
            args.input, output_path, title=args.title
        )
        
        if success:
            print(f"✅ Success: {output_path}")
        else:
            print("❌ Conversion failed")
            
    except ImportError as e:
        print(f"❌ Missing dependency: {e}")
        print("Install required package: pip install reportlab chardet")
    except Exception as e:
        print(f"❌ Error: {e}")
        logger.error(f"Unexpected error: {e}", exc_info=True)