import os
import pathlib
import logging
from typing import Optional, Dict, Any, List

# Aspose.Cells for production-grade Excel processing
try:
    import aspose.cells as cells
    from aspose.cells import Workbook, SaveFormat, PageOrientationType, PaperSizeType
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
    Enhanced with proper scaling and page fitting options.
    """
    
    def __init__(self, default_fit_to_pages_wide: int = 1, default_orientation: str = "auto"):
        """
        Initialize the Excel to PDF service.
        
        Args:
            default_fit_to_pages_wide: Default number of pages wide to fit content (1 = fit all columns on one page width)
            default_orientation: Default page orientation ("auto", "portrait", "landscape")
        """
        if not ASPOSE_AVAILABLE:
            raise ImportError(
                "Aspose.Cells is required for production use. "
                "Install with: pip install aspose-cells-python"
            )
        
        self.default_fit_to_pages_wide = default_fit_to_pages_wide
        self.default_orientation = default_orientation
    
    def get_excel_metadata(self, excel_path: str) -> Dict[str, Any]:
        """Extract metadata from Excel file."""
        try:
            workbook = Workbook(excel_path) # type: ignore
            
            # Get file info
            file_stat = os.stat(excel_path)
            file_size_mb = file_stat.st_size / (1024 * 1024)
            
            # Get worksheet count safely
            try:
                sheets_count = len(workbook.worksheets) # type: ignore
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
    
    def _configure_page_setup(self, worksheet, fit_to_pages_wide: Optional[int] = None, 
                            fit_to_pages_tall: Optional[int] = None, 
                            orientation: str = "auto", zoom_scale: Optional[int] = None):
        """
        Configure page setup for optimal PDF output.
        
        Args:
            worksheet: The worksheet to configure
            fit_to_pages_wide: Number of pages wide to fit content (None for default)
            fit_to_pages_tall: Number of pages tall to fit content (None for auto)
            orientation: Page orientation ("auto", "portrait", "landscape")
            zoom_scale: Manual zoom percentage (overrides fit_to_pages if set)
        """
        try:
            page_setup = worksheet.page_setup
            
            # Set paper size to A4
            page_setup.paper_size = PaperSizeType.PAPER_A4
            
            # Configure scaling options
            if zoom_scale is not None:
                # Manual zoom scaling
                page_setup.zoom = zoom_scale
                logger.info(f"Applied manual zoom: {zoom_scale}%")
            else:
                # Fit-to-pages scaling
                fit_wide = fit_to_pages_wide if fit_to_pages_wide is not None else self.default_fit_to_pages_wide
                
                # Set fit to pages options
                page_setup.fit_to_pages_wide = fit_wide
                if fit_to_pages_tall is not None:
                    page_setup.fit_to_pages_tall = fit_to_pages_tall
                else:
                    # Let height auto-adjust
                    page_setup.fit_to_pages_tall = 0  # 0 means auto-adjust height
                
                logger.info(f"Applied fit-to-pages: {fit_wide} pages wide, auto height")
            
            # Set orientation
            if orientation == "landscape":
                page_setup.orientation = PageOrientationType.LANDSCAPE
                logger.info("Set orientation: Landscape")
            elif orientation == "portrait":
                page_setup.orientation = PageOrientationType.PORTRAIT
                logger.info("Set orientation: Portrait")
            else:  # auto
                # Auto-detect based on used range
                try:
                    used_range = worksheet.cells.max_display_range
                    if used_range:
                        cols = used_range.column_count
                        rows = used_range.row_count
                        if cols > rows * 1.5:  # More columns than rows, use landscape
                            page_setup.orientation = PageOrientationType.LANDSCAPE
                            logger.info("Auto-detected orientation: Landscape")
                        else:
                            page_setup.orientation = PageOrientationType.PORTRAIT
                            logger.info("Auto-detected orientation: Portrait")
                except:
                    # Default to landscape for better column fitting
                    page_setup.orientation = PageOrientationType.LANDSCAPE
                    logger.info("Default orientation: Landscape")
            
            # Set margins for better space utilization
            page_setup.left_margin = 0.5
            page_setup.right_margin = 0.5
            page_setup.top_margin = 0.5
            page_setup.bottom_margin = 0.5
            page_setup.header_margin = 0.3
            page_setup.footer_margin = 0.3
            
            # Print quality and other options
            page_setup.print_quality = 300
            page_setup.print_draft = False
            page_setup.print_gridlines = True  # Show gridlines in PDF
            page_setup.print_headings = False  # Don't show row/column headers
            
        except Exception as e:
            logger.error(f"Failed to configure page setup: {e}")
    
    def convert_excel_to_pdf(self, excel_path: str, pdf_path: str, 
                           sheet_indices: Optional[List[int]] = None,
                           recalculate_formulas: bool = True,
                           fit_to_pages_wide: Optional[int] = None,
                           fit_to_pages_tall: Optional[int] = None,
                           orientation: str = "auto",
                           zoom_scale: Optional[int] = None) -> bool:
        """
        Convert Excel file to PDF with formula recalculation and proper scaling.
        
        Args:
            excel_path: Path to input Excel file
            pdf_path: Path for output PDF file
            sheet_indices: List of sheet indices to convert (None = all sheets)
            recalculate_formulas: Whether to recalculate formulas before export
            fit_to_pages_wide: Number of pages wide to fit content (None for default)
            fit_to_pages_tall: Number of pages tall to fit content (None for auto)
            orientation: Page orientation ("auto", "portrait", "landscape")
            zoom_scale: Manual zoom percentage (overrides fit_to_pages if set)
        
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
                total_sheets = len(workbook.worksheets) # type: ignore
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
            
            # Configure page setup for all worksheets
            logger.info("Configuring page setup for optimal PDF scaling...")
            for i in range(len(workbook.worksheets)):
                worksheet = workbook.worksheets[i]
                self._configure_page_setup(
                    worksheet, 
                    fit_to_pages_wide=fit_to_pages_wide,
                    fit_to_pages_tall=fit_to_pages_tall,
                    orientation=orientation,
                    zoom_scale=zoom_scale
                )
            
            # Save as PDF using SaveFormat
            workbook.save(pdf_path, SaveFormat.PDF) # type: ignore
            
            logger.info(f"Successfully converted Excel to PDF: {pdf_path}")
            return True
            
        except Exception as e:
            logger.error(f"Failed to convert Excel to PDF: {e}")
            return False
    
    def convert_sheet_to_pdf(self, excel_path: str, pdf_path: str, 
                           sheet_name: str, recalculate_formulas: bool = True,
                           fit_to_pages_wide: Optional[int] = None,
                           fit_to_pages_tall: Optional[int] = None,
                           orientation: str = "auto",
                           zoom_scale: Optional[int] = None) -> bool:
        """
        Convert a specific worksheet to PDF by name with scaling options.
        
        Args:
            excel_path: Path to input Excel file
            pdf_path: Path for output PDF file
            sheet_name: Name of the worksheet to convert
            recalculate_formulas: Whether to recalculate formulas before export
            fit_to_pages_wide: Number of pages wide to fit content (None for default)
            fit_to_pages_tall: Number of pages tall to fit content (None for auto)
            orientation: Page orientation ("auto", "portrait", "landscape")
            zoom_scale: Manual zoom percentage (overrides fit_to_pages if set)
        
        Returns:
            bool: True if conversion successful, False otherwise
        """
        try:
            workbook = Workbook(excel_path) # type: ignore
            
            # Find sheet by name
            sheet_index = None
            for i in range(len(workbook.worksheets)): # type: ignore
                if workbook.worksheets[i].name == sheet_name:
                    sheet_index = i
                    break
            
            if sheet_index is None:
                logger.error(f"Sheet '{sheet_name}' not found in workbook")
                return False
            
            return self.convert_excel_to_pdf(
                excel_path, pdf_path, 
                sheet_indices=[sheet_index], 
                recalculate_formulas=recalculate_formulas,
                fit_to_pages_wide=fit_to_pages_wide,
                fit_to_pages_tall=fit_to_pages_tall,
                orientation=orientation,
                zoom_scale=zoom_scale
            )
            
        except Exception as e:
            logger.error(f"Failed to convert sheet '{sheet_name}' to PDF: {e}")
            return False
    
    def batch_convert(self, input_files: List[str], output_dir: str = "output",
                     fit_to_pages_wide: Optional[int] = None,
                     orientation: str = "auto") -> Dict[str, Any]:
        """
        Convert multiple Excel files to PDF with scaling options.
        
        Args:
            input_files: List of Excel file paths
            output_dir: Output directory for PDF files
            fit_to_pages_wide: Number of pages wide to fit content (None for default)
            orientation: Page orientation ("auto", "portrait", "landscape")
        
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
            
            # Convert to PDF with scaling options
            success = self.convert_excel_to_pdf(
                str(input_path), str(pdf_path),
                fit_to_pages_wide=fit_to_pages_wide,
                orientation=orientation
            )
            
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
    
    parser = argparse.ArgumentParser(description="Excel to PDF Conversion Service with Scaling")
    parser.add_argument("input", nargs="+", help="Input Excel file(s)")
    parser.add_argument("-o", "--output", default="temp", help="Output directory (default: temp)")
    parser.add_argument("-s", "--sheet", help="Convert specific sheet by name")
    parser.add_argument("--no-recalc", action="store_true", help="Skip formula recalculation")
    parser.add_argument("-w", "--fit-width", type=int, default=1, 
                       help="Fit content to N pages wide (default: 1)")
    parser.add_argument("-h", "--fit-height", type=int, 
                       help="Fit content to N pages tall (default: auto)")
    parser.add_argument("--orientation", choices=["auto", "portrait", "landscape"], 
                       default="auto", help="Page orientation (default: auto)")
    parser.add_argument("--zoom", type=int, help="Manual zoom percentage (overrides fit options)")
    parser.add_argument("-v", "--verbose", action="store_true", help="Verbose logging")
    
    args = parser.parse_args()
    
    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)
    
    try:
        service = ExcelToPdfService(
            default_fit_to_pages_wide=args.fit_width,
            default_orientation=args.orientation
        )
        
        if len(args.input) == 1 and args.sheet:
            # Single file, specific sheet
            input_file = args.input[0]
            input_path = pathlib.Path(input_file)
            output_path = pathlib.Path(args.output)
            output_path.mkdir(parents=True, exist_ok=True)
            
            pdf_filename = f"{input_path.stem}_{args.sheet}.pdf"
            pdf_path = output_path / pdf_filename
            
            print(f"Converting sheet '{args.sheet}' from {input_file}")
            print(f"Scaling: {args.fit_width} pages wide, orientation: {args.orientation}")
            
            success = service.convert_sheet_to_pdf(
                input_file, str(pdf_path), args.sheet, 
                recalculate_formulas=not args.no_recalc,
                fit_to_pages_wide=args.fit_width,
                fit_to_pages_tall=args.fit_height,
                orientation=args.orientation,
                zoom_scale=args.zoom
            )
            
            if success:
                print(f"✅ Successfully converted to: {pdf_path}")
            else:
                print("❌ Conversion failed")
                
        else:
            # Batch conversion
            print(f"Converting {len(args.input)} file(s)...")
            print(f"Scaling: {args.fit_width} pages wide, orientation: {args.orientation}")
            
            results = service.batch_convert(
                args.input, args.output,
                fit_to_pages_wide=args.fit_width,
                orientation=args.orientation
            )
            
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
    # Example usage for your specific case with enhanced scaling
    test_files = [
        "input/file_example_XLS_100.xls",
    ]
    
    # Direct service usage with scaling options
    try:
        # Initialize service with default settings for better column fitting
        service = ExcelToPdfService(
            default_fit_to_pages_wide=1,  # Fit all columns to 1 page width
            default_orientation="auto"    # Auto-detect best orientation
        )
        
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
                
                # Convert to PDF with enhanced scaling
                input_path = pathlib.Path(file_path)
                pdf_filename = f"{input_path.stem}_scaled.pdf"
                pdf_path = pathlib.Path("temp") / pdf_filename
                
                # Create output directory
                pdf_path.parent.mkdir(parents=True, exist_ok=True)
                
                print(f"Converting with scaling options...")
                print(f"  - Fit to 1 page wide")
                print(f"  - Auto orientation detection")
                print(f"  - Reduced margins for better space utilization")
                
                success = service.convert_excel_to_pdf(
                    file_path, str(pdf_path),
                    fit_to_pages_wide=1,      # Fit all columns to 1 page width
                    orientation="auto",       # Auto-detect best orientation
                    recalculate_formulas=True
                )
                
                if success:
                    print(f"✅ SUCCESS!")
                    print(f"PDF saved to: {pdf_path}")
                    print(f"All columns should now fit on the page width!")
                else:
                    print("❌ FAILED!")
                    
                # Alternative: Convert with manual zoom for more control
                pdf_filename_zoom = f"{input_path.stem}_zoom70.pdf"
                pdf_path_zoom = pathlib.Path("temp") / pdf_filename_zoom
                
                print(f"\nTrying alternative with 70% zoom...")
                success_zoom = service.convert_excel_to_pdf(
                    file_path, str(pdf_path_zoom),
                    zoom_scale=70,            # Manual 70% zoom
                    orientation="landscape",  # Force landscape
                    recalculate_formulas=True
                )
                
                if success_zoom:
                    print(f"✅ Zoom version created: {pdf_path_zoom}")
                    
            else:
                print(f"⚠️  File not found: {file_path}")
                
    except ImportError as e:
        print(f"❌ Missing required dependency: {e}")
        print("Install Aspose.Cells: pip install aspose-cells-python")
    except Exception as e:
        print(f"❌ Error: {e}")