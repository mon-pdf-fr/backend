"""
PDF utilities for page numbering and other operations using ReportLab and PyPDF
"""

import io
from enum import Enum
from typing import List, Dict

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
