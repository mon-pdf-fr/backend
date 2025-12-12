"""
FastAPI application for PDF operations
"""

from fastapi import FastAPI, File, UploadFile, HTTPException, Form
from fastapi.responses import StreamingResponse
from pdf_utils import add_page_numbers, PageNumberPosition
import io


app = FastAPI(
    title="Docling API",
    description="PDF manipulation API using open-source tools",
    version="1.0.0"
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
            }
        ]
    }
