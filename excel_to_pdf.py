import os
import pathlib
import logging
from typing import Optional, Dict, Any, List

# Aspose.Cells for production-grade Excel processing
try:
    import aspose.cells as cells
    from aspose.cells import Workbook, SaveFormat
    ASPOSE_AVAILABLE = True
    logging.info("Aspose.Cells available - production Excel processing enabled")
except ImportError:
    ASPOSE_AVAILABLE = False
    logging.warning("Aspose.Cells not available. Install: pip install aspose-cells-python")

# Setup logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class ExcelToPdfService:
    """
    Production-ready Excel to PDF conversion service using Aspose.Cells.
    Preserves formulas, calculations, and Excel formatting for financial documents.
    """
    
    def __init__(self):
        """Initialize the Excel to PDF service."""
        if not ASPOSE_AVAILABLE:
            raise ImportError(
                "Aspose.Cells is required for production use. "
                "Install with: pip install aspose-cells-python"
            )
    
    def get_excel_metadata(self, excel_path: str) -> Dict[str, Any]:
        """Extract metadata from Excel file."""
        try:
            workbook = Workbook(excel_path) # type: ignore
            
            # Get file info
            file_stat = os.stat(excel_path)
            file_size_mb = file_stat.st_size / (1024 * 1024)
            
            # Get worksheet count safely
            try:
                sheets_count = workbook.worksheets.count # type: ignore
            except Exception:
                sheets_count = 0
            
            # Get worksheet names safely
            worksheet_names = []
            try:
                for i in range(sheets_count):
                    worksheet_names.append(workbook.worksheets[i].name)
            except:
                worksheet_names = []
            
            # Extract built-in document properties with safe attribute access
            try:
                props = workbook.built_in_document_properties
                title = props.title if hasattr(props, 'title') and props.title else None
                author = props.author if hasattr(props, 'author') and props.author else None
            except:
                title = None
                author = None
            
            # Check for macros safely
            try:
                has_macros = workbook.has_macro if hasattr(workbook, 'has_macro') else False
            except:
                has_macros = False
            
            metadata = {
                "title": title,
                "sheets_count": sheets_count,
                "file_size_mb": round(file_size_mb, 2),
                "author": author,
                "has_macros": has_macros,
                "worksheet_names": worksheet_names
            }
            
            # Try to get creation and modification times with different possible attribute names
            try:
                props = workbook.built_in_document_properties
                if hasattr(props, 'created_time'):
                    metadata["date_created"] = str(props.created_time)
                elif hasattr(props, 'creation_date'):
                    metadata["date_created"] = str(props.creation_date) # type: ignore
                else:
                    metadata["date_created"] = None
            except:
                metadata["date_created"] = None
                
            try:
                props = workbook.built_in_document_properties
                if hasattr(props, 'last_modified_time'):
                    metadata["last_modified"] = str(props.last_modified_time) # type: ignore
                elif hasattr(props, 'last_save_time'):
                    metadata["last_modified"] = str(props.last_save_time) # type: ignore
                elif hasattr(props, 'modified'):
                    metadata["last_modified"] = str(props.modified) # type: ignore
                else:
                    metadata["last_modified"] = None
            except:
                metadata["last_modified"] = None
            
            return metadata
            
        except Exception as e:
            logger.error(f"Failed to extract Excel metadata: {e}")
            return {}
    
    def convert_excel_to_pdf(self, excel_path: str, pdf_path: str, 
                           sheet_indices: Optional[List[int]] = None,
                           recalculate_formulas: bool = True) -> bool:
        """
        Convert Excel file to PDF with formula recalculation.
        
        Args:
            excel_path: Path to input Excel file
            pdf_path: Path for output PDF file
            sheet_indices: List of sheet indices to convert (None = all sheets)
            recalculate_formulas: Whether to recalculate formulas before export
        
        Returns:
            bool: True if conversion successful, False otherwise
        """
        try:
            logger.info(f"Converting Excel to PDF: {excel_path} -> {pdf_path}")
            
            # Load workbook
            workbook = Workbook(excel_path) # type: ignore
            
            # Critical: Calculate all formulas before export
            # This ensures P&L and BS calculations are current
            if recalculate_formulas:
                workbook.calculate_formula()
                logger.info("All formulas recalculated successfully")
            
            # Get worksheet count safely
            try:
                total_sheets = workbook.worksheets.count # type: ignore
            except Exception:
                total_sheets = 1
            
            # If specific sheets are requested, create a new workbook with only those sheets
            if sheet_indices is not None:
                logger.info(f"Converting specific sheets: {sheet_indices}")
                
                # Create new workbook for selected sheets
                new_workbook = Workbook() # type: ignore
                # Remove default sheet
                new_workbook.worksheets.remove_at(0) # type: ignore
                
                for i, sheet_index in enumerate(sheet_indices):
                    if 0 <= sheet_index < total_sheets:
                        # Copy worksheet
                        source_sheet = workbook.worksheets[sheet_index]
                        new_workbook.worksheets.add(source_sheet.name)
                        target_sheet = new_workbook.worksheets[i]
                        target_sheet.copy(source_sheet)
                
                workbook = new_workbook
                if recalculate_formulas:
                    workbook.calculate_formula()
            
            # Save as PDF using SaveFormat
            workbook.save(pdf_path, SaveFormat.PDF) # type: ignore
            
            logger.info(f"Successfully converted Excel to PDF: {pdf_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to convert Excel to PDF: {e}")
            return False
    
    def convert_sheet_to_pdf(self, excel_path: str, pdf_path: str, 
                           sheet_name: str, recalculate_formulas: bool = True) -> bool:
        """
        Convert a specific worksheet to PDF by name.
        
        Args:
            excel_path: Path to input Excel file
            pdf_path: Path for output PDF file
            sheet_name: Name of the worksheet to convert
            recalculate_formulas: Whether to recalculate formulas before export
        
        Returns:
            bool: True if conversion successful, False otherwise
        """
        try:
            workbook = Workbook(excel_path) # type: ignore
            
            # Find sheet by name
            sheet_index = None
            for i in range(workbook.worksheets.count): # type: ignore
                if workbook.worksheets[i].name == sheet_name:
                    sheet_index = i
                    break
            
            if sheet_index is None:
                logger.error(f"Sheet '{sheet_name}' not found in workbook")
                return False
            
            return self.convert_excel_to_pdf(excel_path, pdf_path, 
                                           sheet_indices=[sheet_index], 
                                           recalculate_formulas=recalculate_formulas)
            
        except Exception as e:
            logger.error(f"Failed to convert sheet '{sheet_name}' to PDF: {e}")
            return False
    
    def batch_convert(self, input_files: List[str], output_dir: str = "output") -> Dict[str, Any]:
        """
        Convert multiple Excel files to PDF.
        
        Args:
            input_files: List of Excel file paths
            output_dir: Output directory for PDF files
        
        Returns:
            Dict containing conversion results
        """
        # Create output directory
        output_path = pathlib.Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)
        
        results = {
            "successful_conversions": [],
            "failed_conversions": [],
            "total_files": len(input_files),
            "success_count": 0,
            "failure_count": 0
        }
        
        for file_path in input_files:
            input_path = pathlib.Path(file_path)
            
            if not input_path.exists():
                logger.error(f"File not found: {file_path}")
                results["failed_conversions"].append({
                    "file": file_path,
                    "error": "File not found"
                })
                results["failure_count"] += 1
                continue
            
            # Check if it's an Excel file
            if input_path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm']:
                logger.error(f"Not an Excel file: {file_path}")
                results["failed_conversions"].append({
                    "file": file_path,
                    "error": "Not an Excel file"
                })
                results["failure_count"] += 1
                continue
            
            # Generate PDF path
            pdf_filename = f"{input_path.stem}.pdf"
            pdf_path = output_path / pdf_filename
            
            # Get metadata
            metadata = self.get_excel_metadata(file_path)
            
            # Convert to PDF
            success = self.convert_excel_to_pdf(str(input_path), str(pdf_path))
            
            if success:
                results["successful_conversions"].append({
                    "input_file": file_path,
                    "output_file": str(pdf_path),
                    "metadata": metadata
                })
                results["success_count"] += 1
            else:
                results["failed_conversions"].append({
                    "file": file_path,
                    "error": "Conversion failed"
                })
                results["failure_count"] += 1
        
        return results


def main():
    """Main function for command-line usage."""
    import argparse
    
    parser = argparse.ArgumentParser(description="Excel to PDF Conversion Service")
    parser.add_argument("input", nargs="+", help="Input Excel file(s)")
    parser.add_argument("-o", "--output", default="output", help="Output directory (default: output)")
    parser.add_argument("-s", "--sheet", help="Convert specific sheet by name")
    parser.add_argument("--no-recalc", action="store_true", help="Skip formula recalculation")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging")
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        service = ExcelToPdfService()
        
        if len(args.input) == 1 and args.sheet:
            # Single file, specific sheet
            input_file = args.input[0]
            input_path = pathlib.Path(input_file)
            output_path = pathlib.Path(args.output)
            output_path.mkdir(parents=True, exist_ok=True)
            
            pdf_filename = f"{input_path.stem}_{args.sheet}.pdf"
            pdf_path = output_path / pdf_filename
            
            print(f"Converting sheet '{args.sheet}' from {input_file}")
            success = service.convert_sheet_to_pdf(
                input_file, str(pdf_path), args.sheet, 
                recalculate_formulas=not args.no_recalc
            )
            
            if success:
                print(f"✅ Successfully converted to: {pdf_path}")
            else:
                print("❌ Conversion failed")
                
        else:
            # Batch conversion
            print(f"Converting {len(args.input)} file(s)...")
            results = service.batch_convert(args.input, args.output)
            
            print(f"\n📊 Conversion Results:")
            print(f"Total files: {results['total_files']}")
            print(f"Successful: {results['success_count']}")
            print(f"Failed: {results['failure_count']}")
            
            if results["successful_conversions"]:
                print(f"\n✅ Successful conversions:")
                for conversion in results["successful_conversions"]:
                    print(f"  {conversion['input_file']} -> {conversion['output_file']}")
                    metadata = conversion["metadata"]
                    if metadata.get("sheets_count"):
                        print(f"    Sheets: {metadata['sheets_count']}, Size: {metadata.get('file_size_mb', 'N/A')} MB")
            
            if results["failed_conversions"]:
                print(f"\n❌ Failed conversions:")
                for failure in results["failed_conversions"]:
                    print(f"  {failure['file']}: {failure['error']}")
                    
    except ImportError as e:
        print(f"❌ Missing required dependency: {e}")
        print("Install Aspose.Cells: pip install aspose-cells-python")
    except Exception as e:
        print(f"❌ Error: {e}")
        logger.error(f"Unexpected error: {e}", exc_info=True)


if __name__ == "__main__":
    # Example usage for your specific case
    test_files = [
        "input/file_example_XLS_100.xls",
    ]
    
    # Direct service usage
    try:
        service = ExcelToPdfService()
        
        for file_path in test_files:
            if os.path.exists(file_path):
                print(f"\n{'='*60}")
                print(f"Processing: {file_path}")
                print(f"{'='*60}")
                
                # Get metadata
                metadata = service.get_excel_metadata(file_path)
                print(f"Excel Info:")
                print(f"  Sheets: {metadata.get('sheets_count', 'N/A')}")
                print(f"  Size: {metadata.get('file_size_mb', 'N/A')} MB")
                print(f"  Worksheets: {metadata.get('worksheet_names', [])}")
                
                # Convert to PDF
                input_path = pathlib.Path(file_path)
                pdf_filename = f"{input_path.stem}.pdf"
                pdf_path = pathlib.Path("output") / pdf_filename
                
                # Create output directory
                pdf_path.parent.mkdir(parents=True, exist_ok=True)
                
                success = service.convert_excel_to_pdf(file_path, str(pdf_path))
                
                if success:
                    print(f"✅ SUCCESS!")
                    print(f"PDF saved to: {pdf_path}")
                else:
                    print("❌ FAILED!")
                    
            else:
                print(f"⚠️  File not found: {file_path}")
                
    except ImportError as e:
        print(f"❌ Missing required dependency: {e}")
        print("Install Aspose.Cells: pip install aspose-cells-python")
    except Exception as e:
        print(f"❌ Error: {e}")