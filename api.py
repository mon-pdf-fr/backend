"""
FastAPI application for PDF operations
"""

from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import StreamingResponse, Response
from fastapi.middleware.cors import CORSMiddleware
from pdf_utils import (
    add_page_numbers,
    PageNumberPosition,
    extract_images_from_pdf,
    compress_pdf_all_qualities,
    CompressionQuality
)
import io
import zipfile
import json
import base64


app = FastAPI(
    title="Docling API",
    description="PDF manipulation API using open-source tools",
    version="1.0.0"
)

# Add CORS middleware to allow requests from your Next.js app
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # In production, replace with your actual domain
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    max_age=3600,
)


@app.post("/add-page-numbers", summary="Add page numbers to PDF")
async def add_page_numbers_endpoint(
    file: UploadFile = File(..., description="PDF file to add page numbers to"),
    position: PageNumberPosition = Form(
        PageNumberPosition.BOTTOM_CENTER,
        description="Position of page numbers on the page"
    ),
    font_size: int = Form(12, description="Font size for page numbers", ge=6, le=72),
    margin: int = Form(30, description="Margin from edge in points", ge=0),
    start_page: int = Form(1, description="Starting page number", ge=1),
    format_string: str = Form(
        "{page}",
        description="Format string (use {page} for page number, {total} for total pages)"
    ),
    font_color_r: int = Form(0, description="Red component (0-255)", ge=0, le=255),
    font_color_g: int = Form(0, description="Green component (0-255)", ge=0, le=255),
    font_color_b: int = Form(0, description="Blue component (0-255)", ge=0, le=255)
):
    """
    Add page numbers to a PDF document using ReportLab and PyPDF (fully open-source).

    **9 Positions Available:**
    - top_left, top_center, top_right
    - middle_left, middle_center, middle_right
    - bottom_left, bottom_center, bottom_right

    **Parameters:**
    - **file**: PDF file to process
    - **position**: Where to place page numbers
    - **font_size**: Size of the page number text (6-72)
    - **margin**: Distance from edge in points
    - **start_page**: What number to start counting from
    - **format_string**: How to format page numbers (e.g., "Page {page}", "{page}/{total}")
    - **font_color_r/g/b**: RGB color components (0-255)
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")

    try:
        # Read file
        pdf_bytes = await file.read()

        # Add page numbers
        result_bytes = add_page_numbers(
            pdf_bytes=pdf_bytes,
            position=position,
            font_size=font_size,
            font_color=(font_color_r, font_color_g, font_color_b),
            margin=margin,
            start_page=start_page,
            format_string=format_string
        )

        # Return modified PDF
        return StreamingResponse(
            io.BytesIO(result_bytes),
            media_type="application/pdf",
            headers={
                "Content-Disposition": f"attachment; filename=numbered_{file.filename}"
            }
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error processing PDF: {str(e)}")


@app.post("/extract-images", summary="Extract images from PDF")
async def extract_images_endpoint(
    file: UploadFile = File(..., description="PDF file to extract images from")
):
    """
    Extract all images from a PDF document (fast PyPDF extraction).

    **Features:**
    - Fast extraction using PyPDF
    - All images converted to PNG format
    - Provides metadata (page number, dimensions, format)
    - Returns base64-encoded images for frontend preview and download

    **Returns:**
    JSON with list of images and metadata. Each image includes:
    - image_base64: Base64-encoded PNG image data
    - page: Page number where image appears
    - index: Unique index for the image
    - width/height: Image dimensions
    - format: Output format (PNG)
    - original_format: Original format in PDF (JPEG, PNG, etc.)
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")

    try:
        # Read file
        pdf_bytes = await file.read()

        # Extract images
        images = extract_images_from_pdf(pdf_bytes)

        if not images:
            return {
                "total_images": 0,
                "images": []
            }

        # Convert to JSON response with base64
        import base64
        result = []
        for img_data in images:
            result.append({
                "image_base64": base64.b64encode(img_data['image_bytes']).decode('utf-8'),
                "page": img_data['page'],
                "index": img_data['index'],
                "width": img_data['width'],
                "height": img_data['height'],
                "format": img_data['format'],
                "original_format": img_data.get('original_format', img_data['format'])
            })

        return {
            "total_images": len(result),
            "images": result
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error extracting images: {str(e)}")


@app.post("/compress-pdf", summary="Compress PDF to all quality levels")
async def compress_pdf_endpoint(
    file: UploadFile = File(..., description="PDF file to compress")
):
    """
    Compress a PDF to all three quality levels using Ghostscript.

    **Features:**
    - High quality: 150 DPI, ~25% size reduction
    - Medium quality: 100 DPI, ~45% size reduction
    - Low quality: 72 DPI, ~65% size reduction
    - Returns base64-encoded PDFs for all quality levels

    **Returns:**
    JSON with compressed PDFs in base64 format for each quality level:
    - originalSize: Original file size in bytes
    - qualities.high/medium/low: Compressed data with size and ratio
    """
    # Validate file type
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")

    try:
        # Read file
        pdf_bytes = await file.read()

        file_size_mb = len(pdf_bytes) / 1024 / 1024
        print(f"[Compress PDF] Processing {file.filename} ({file_size_mb:.2f}MB)")

        # Compress to all quality levels
        results = compress_pdf_all_qualities(pdf_bytes)

        # Convert to response format with base64
        response = {
            "success": True,
            "originalSize": results["original_size"],
            "qualities": {
                "high": {
                    "size": results["high"]["size"],
                    "ratio": results["high"]["ratio"],
                    "blob": base64.b64encode(results["high"]["compressed_bytes"]).decode('utf-8')
                },
                "medium": {
                    "size": results["medium"]["size"],
                    "ratio": results["medium"]["ratio"],
                    "blob": base64.b64encode(results["medium"]["compressed_bytes"]).decode('utf-8')
                },
                "low": {
                    "size": results["low"]["size"],
                    "ratio": results["low"]["ratio"],
                    "blob": base64.b64encode(results["low"]["compressed_bytes"]).decode('utf-8')
                }
            }
        }

        print(f"[Compress PDF] Compression complete for {file.filename}")
        return response

    except HTTPException:
        raise
    except Exception as e:
        print(f"[Compress PDF] Error: {e}")
        raise HTTPException(status_code=500, detail=f"Error compressing PDF: {str(e)}")


@app.get("/health/ghostscript", summary="Check Ghostscript availability")
async def check_ghostscript():
    """Check if Ghostscript is installed and available"""
    import subprocess
    try:
        result = subprocess.run(['gs', '--version'], capture_output=True, text=True, timeout=5)
        if result.returncode == 0:
            return {
                "available": True,
                "version": result.stdout.strip(),
                "message": "Ghostscript is installed and working"
            }
        else:
            return {
                "available": False,
                "error": "Ghostscript command failed",
                "stderr": result.stderr
            }
    except FileNotFoundError:
        return {
            "available": False,
            "error": "Ghostscript not found - compression will not work",
            "message": "Install Ghostscript or deploy to a platform that supports it"
        }
    except Exception as e:
        return {
            "available": False,
            "error": str(e)
        }


@app.get("/", summary="API Health Check")
async def root():
    """Health check endpoint"""
    return {
        "status": "ok",
        "message": "Docling API is running",
        "version": "1.0.0",
        "endpoints": [
            {
                "path": "/add-page-numbers",
                "method": "POST",
                "description": "Add page numbers to PDF (9 positions available)"
            },
            {
                "path": "/extract-images",
                "method": "POST",
                "description": "Extract all images from PDF (fast)"
            },
            {
                "path": "/compress-pdf",
                "method": "POST",
                "description": "Compress PDF to all quality levels (high, medium, low)"
            },
            {
                "path": "/health/ghostscript",
                "method": "GET",
                "description": "Check if Ghostscript is available for compression"
            }
        ]
    }
