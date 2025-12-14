"""
PDF utilities for page numbering and other operations using ReportLab and PyPDF
"""

import io
import subprocess
import tempfile
import os
from enum import Enum
from typing import List, Dict, Tuple

from PIL import Image
from pypdf import PdfReader, PdfWriter
from reportlab.lib.colors import Color
from reportlab.pdfgen import canvas


class PageNumberPosition(str, Enum):
    """Available positions for page numbers"""
    TOP_LEFT = "top_left"
    TOP_CENTER = "top_center"
    TOP_RIGHT = "top_right"
    MIDDLE_LEFT = "middle_left"
    MIDDLE_CENTER = "middle_center"
    MIDDLE_RIGHT = "middle_right"
    BOTTOM_LEFT = "bottom_left"
    BOTTOM_CENTER = "bottom_center"
    BOTTOM_RIGHT = "bottom_right"


def add_page_numbers(
    pdf_bytes: bytes,
    position: PageNumberPosition = PageNumberPosition.BOTTOM_CENTER,
    font_size: int = 12,
    font_color: tuple = (0, 0, 0),  # RGB 0-255
    margin: int = 30,
    start_page: int = 1,
    format_string: str = "{page}"
) -> bytes:
    """
    Add page numbers to a PDF document using ReportLab and PyPDF.

    Args:
        pdf_bytes: PDF file as bytes
        position: Position of page numbers (9 positions available)
        font_size: Font size for page numbers
        font_color: RGB color tuple (0-255 range)
        margin: Margin from edge in points
        start_page: Starting page number
        format_string: Format string for page numbers (use {page} as placeholder, {total} for total pages)

    Returns:
        Modified PDF as bytes

    Example:
        >>> pdf_bytes = open('input.pdf', 'rb').read()
        >>> result = add_page_numbers(pdf_bytes, position=PageNumberPosition.BOTTOM_CENTER)
        >>> open('output.pdf', 'wb').write(result)
    """
    # Read the input PDF
    input_pdf = PdfReader(io.BytesIO(pdf_bytes))
    output_pdf = PdfWriter()

    # Convert RGB from 0-255 to 0-1 range for ReportLab
    color = Color(
        font_color[0] / 255.0,
        font_color[1] / 255.0,
        font_color[2] / 255.0
    )

    # Process each page
    for page_num in range(len(input_pdf.pages)):
        page = input_pdf.pages[page_num]

        # Get page dimensions
        page_width = float(page.mediabox.width)
        page_height = float(page.mediabox.height)

        # Create a new PDF with the page number
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(page_width, page_height))

        # Format the page number text
        page_number_text = format_string.format(
            page=page_num + start_page,
            total=len(input_pdf.pages)
        )

        # Set font and color
        can.setFont("Helvetica", font_size)
        can.setFillColor(color)

        # Calculate position and draw text
        if position == PageNumberPosition.TOP_LEFT:
            x = margin
            y = page_height - margin
            can.drawString(x, y, page_number_text)
        elif position == PageNumberPosition.TOP_CENTER:
            x = page_width / 2
            y = page_height - margin
            can.drawCentredString(x, y, page_number_text)
        elif position == PageNumberPosition.TOP_RIGHT:
            x = page_width - margin
            y = page_height - margin
            can.drawRightString(x, y, page_number_text)
        elif position == PageNumberPosition.MIDDLE_LEFT:
            x = margin
            y = page_height / 2
            can.drawString(x, y, page_number_text)
        elif position == PageNumberPosition.MIDDLE_CENTER:
            x = page_width / 2
            y = page_height / 2
            can.drawCentredString(x, y, page_number_text)
        elif position == PageNumberPosition.MIDDLE_RIGHT:
            x = page_width - margin
            y = page_height / 2
            can.drawRightString(x, y, page_number_text)
        elif position == PageNumberPosition.BOTTOM_LEFT:
            x = margin
            y = margin
            can.drawString(x, y, page_number_text)
        elif position == PageNumberPosition.BOTTOM_CENTER:
            x = page_width / 2
            y = margin
            can.drawCentredString(x, y, page_number_text)
        else:  # BOTTOM_RIGHT
            x = page_width - margin
            y = margin
            can.drawRightString(x, y, page_number_text)

        can.save()

        # Move to the beginning of the BytesIO buffer
        packet.seek(0)
        overlay_pdf = PdfReader(packet)

        # Merge the overlay with the original page
        page.merge_page(overlay_pdf.pages[0])
        output_pdf.add_page(page)

    # Save to bytes
    output_bytes = io.BytesIO()
    output_pdf.write(output_bytes)
    output_bytes.seek(0)

    return output_bytes.getvalue()


def extract_images_from_pdf(pdf_bytes: bytes) -> List[Dict[str, any]]:
    """
    Extract all images from a PDF document using PyPDF (fast and lightweight).

    Args:
        pdf_bytes: PDF file as bytes

    Returns:
        List of dictionaries containing image data with metadata:
        [
            {
                "image_bytes": bytes (original format or PNG),
                "page": int (page number, 1-indexed),
                "index": int (global image index),
                "width": int,
                "height": int,
                "format": str (JPEG, PNG, etc.)
            }
        ]

    Example:
        >>> pdf_bytes = open('input.pdf', 'rb').read()
        >>> images = extract_images_from_pdf(pdf_bytes)
        >>> for i, img_data in enumerate(images):
        ...     with open(f'image_{i}.{img_data["format"].lower()}', 'wb') as f:
        ...         f.write(img_data['image_bytes'])
    """
    # Read the PDF
    pdf_reader = PdfReader(io.BytesIO(pdf_bytes))
    images = []
    global_index = 0

    # Iterate through each page
    for page_num, page in enumerate(pdf_reader.pages, start=1):
        # Extract images from the page
        if '/Resources' in page and '/XObject' in page['/Resources']:
            xObject = page['/Resources']['/XObject'].get_object()

            for obj_name in xObject:
                obj = xObject[obj_name]

                # Check if it's an image
                if obj['/Subtype'] == '/Image':
                    try:
                        # Get image properties
                        width = obj['/Width']
                        height = obj['/Height']

                        # Try to determine format
                        if '/Filter' in obj:
                            filter_type = obj['/Filter']
                            if filter_type == '/DCTDecode':
                                img_format = 'JPEG'
                                ext = 'jpg'
                            elif filter_type == '/FlateDecode':
                                img_format = 'PNG'
                                ext = 'png'
                            elif filter_type == '/JPXDecode':
                                img_format = 'JPEG2000'
                                ext = 'jp2'
                            else:
                                img_format = 'PNG'  # Default
                                ext = 'png'
                        else:
                            img_format = 'PNG'
                            ext = 'png'

                        # Extract image data
                        img_data = obj.get_data()

                        # Convert to PIL Image and then to PNG for consistency
                        try:
                            pil_image = Image.open(io.BytesIO(img_data))

                            # Convert to PNG bytes
                            img_bytes_io = io.BytesIO()
                            pil_image.save(img_bytes_io, format='PNG')
                            img_bytes = img_bytes_io.getvalue()

                            images.append({
                                "image_bytes": img_bytes,
                                "page": page_num,
                                "index": global_index,
                                "width": pil_image.width,
                                "height": pil_image.height,
                                "format": "PNG",
                                "original_format": img_format
                            })
                            global_index += 1
                        except Exception as e:
                            # If PIL fails, use raw data
                            print(f"Warning: Could not convert image to PNG on page {page_num}: {e}")
                            images.append({
                                "image_bytes": img_data,
                                "page": page_num,
                                "index": global_index,
                                "width": width,
                                "height": height,
                                "format": img_format,
                                "original_format": img_format
                            })
                            global_index += 1

                    except Exception as e:
                        print(f"Warning: Could not extract image from page {page_num}: {e}")
                        continue

    return images


class CompressionQuality(str, Enum):
    """Available compression quality levels"""
    HIGH = "high"
    MEDIUM = "medium"
    LOW = "low"


def compress_pdf(
    pdf_bytes: bytes,
    quality: CompressionQuality = CompressionQuality.MEDIUM
) -> Tuple[bytes, int, int]:
    """
    Compress a PDF using Ghostscript.

    Args:
        pdf_bytes: PDF file as bytes
        quality: Compression quality level (high, medium, low)

    Returns:
        Tuple of (compressed_pdf_bytes, original_size, compressed_size)

    Raises:
        RuntimeError: If Ghostscript is not available or compression fails
    """
    # Ghostscript quality settings
    quality_settings = {
        CompressionQuality.HIGH: '/ebook',      # 150 DPI, good quality
        CompressionQuality.MEDIUM: '/screen',   # 72 DPI, medium quality
        CompressionQuality.LOW: '/screen',      # 72 DPI with more aggressive settings
    }

    dpi_settings = {
        CompressionQuality.HIGH: '150',
        CompressionQuality.MEDIUM: '100',
        CompressionQuality.LOW: '72',
    }

    original_size = len(pdf_bytes)

    # Create temporary files
    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as input_file:
        input_path = input_file.name
        input_file.write(pdf_bytes)

    with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as output_file:
        output_path = output_file.name

    try:
        # First check if Ghostscript is available
        try:
            gs_check = subprocess.run(['gs', '--version'], capture_output=True, timeout=5)
            if gs_check.returncode != 0:
                raise RuntimeError("Ghostscript not available")
            print(f"[Compress PDF] Ghostscript version: {gs_check.stdout.decode().strip()}")
        except (subprocess.SubprocessError, FileNotFoundError) as e:
            raise RuntimeError(f"Ghostscript not installed: {e}")

        # Build Ghostscript command
        gs_command = [
            'gs',
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.4',
            f'-dPDFSETTINGS={quality_settings[quality]}',
            '-dNOPAUSE',
            '-dQUIET',
            '-dBATCH',
            '-dDownsampleColorImages=true',
            f'-dColorImageResolution={dpi_settings[quality]}',
            '-dDownsampleGrayImages=true',
            f'-dGrayImageResolution={dpi_settings[quality]}',
            '-dDownsampleMonoImages=true',
            f'-dMonoImageResolution={dpi_settings[quality]}',
            f'-sOutputFile={output_path}',
            input_path
        ]

        # Run Ghostscript
        print(f"[Compress PDF] Running Ghostscript with quality: {quality}")
        result = subprocess.run(
            gs_command,
            capture_output=True,
            text=True,
            timeout=55  # 55 second timeout
        )

        if result.returncode != 0:
            print(f"[Compress PDF] Ghostscript stderr: {result.stderr}")
            print(f"[Compress PDF] Ghostscript stdout: {result.stdout}")
            raise RuntimeError(f"Ghostscript failed with code {result.returncode}: {result.stderr}")

        # Read compressed file
        with open(output_path, 'rb') as f:
            compressed_bytes = f.read()

        compressed_size = len(compressed_bytes)
        print(f"[Compress PDF] Compressed from {original_size / 1024 / 1024:.2f}MB to {compressed_size / 1024 / 1024:.2f}MB")

        return compressed_bytes, original_size, compressed_size

    finally:
        # Clean up temp files
        try:
            os.unlink(input_path)
        except:
            pass
        try:
            os.unlink(output_path)
        except:
            pass


def compress_pdf_pypdf(pdf_bytes: bytes) -> Tuple[bytes, int, int]:
    """
    Compress a PDF using PyPDF (fallback when Ghostscript is unavailable).

    Args:
        pdf_bytes: PDF file as bytes

    Returns:
        Tuple of (compressed_pdf_bytes, original_size, compressed_size)
    """
    original_size = len(pdf_bytes)

    # Read and write with PyPDF compression
    reader = PdfReader(io.BytesIO(pdf_bytes))
    writer = PdfWriter()

    # Copy pages and compress
    for page in reader.pages:
        page.compress_content_streams()  # Compress content streams
        writer.add_page(page)

    # Add compression options
    writer.add_metadata(reader.metadata)

    # Write to bytes with compression
    output = io.BytesIO()
    writer.write(output)
    compressed_bytes = output.getvalue()
    compressed_size = len(compressed_bytes)

    print(f"[PyPDF Compress] Compressed from {original_size / 1024 / 1024:.2f}MB to {compressed_size / 1024 / 1024:.2f}MB")

    return compressed_bytes, original_size, compressed_size


def compress_pdf_all_qualities(pdf_bytes: bytes) -> Dict[str, Dict[str, any]]:
    """
    Compress a PDF to all three quality levels.
    Uses Ghostscript if available, falls back to PyPDF compression.

    Args:
        pdf_bytes: PDF file as bytes

    Returns:
        Dictionary with compression results for each quality:
        {
            "high": {"compressed_bytes": bytes, "size": int, "ratio": int},
            "medium": {"compressed_bytes": bytes, "size": int, "ratio": int},
            "low": {"compressed_bytes": bytes, "size": int, "ratio": int},
            "original_size": int
        }
    """
    original_size = len(pdf_bytes)
    results = {
        "original_size": original_size
    }

    # Check if Ghostscript is available
    ghostscript_available = False
    try:
        gs_check = subprocess.run(['gs', '--version'], capture_output=True, timeout=2)
        ghostscript_available = gs_check.returncode == 0
        print(f"[Compress PDF] Ghostscript available: {ghostscript_available}")
    except:
        print("[Compress PDF] Ghostscript not available, using PyPDF fallback")

    if ghostscript_available:
        # Use Ghostscript for better compression
        for quality in [CompressionQuality.HIGH, CompressionQuality.MEDIUM, CompressionQuality.LOW]:
            try:
                compressed_bytes, _, compressed_size = compress_pdf(pdf_bytes, quality)
                ratio = round(((original_size - compressed_size) / original_size) * 100)

                results[quality.value] = {
                    "compressed_bytes": compressed_bytes,
                    "size": compressed_size,
                    "ratio": ratio
                }
            except Exception as e:
                print(f"[Compress PDF] Ghostscript error for {quality.value}: {e}")
                # Fall back to PyPDF
                compressed_bytes, _, compressed_size = compress_pdf_pypdf(pdf_bytes)
                ratio = round(((original_size - compressed_size) / original_size) * 100)
                results[quality.value] = {
                    "compressed_bytes": compressed_bytes,
                    "size": compressed_size,
                    "ratio": ratio
                }
    else:
        # Use PyPDF fallback for all qualities
        compressed_bytes, _, compressed_size = compress_pdf_pypdf(pdf_bytes)
        ratio = round(((original_size - compressed_size) / original_size) * 100)

        # For PyPDF, all qualities return the same result (can't do quality levels without Ghostscript)
        for quality in [CompressionQuality.HIGH, CompressionQuality.MEDIUM, CompressionQuality.LOW]:
            results[quality.value] = {
                "compressed_bytes": compressed_bytes,
                "size": compressed_size,
                "ratio": ratio
            }

    return results
