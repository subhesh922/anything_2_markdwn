from fastapi import FastAPI, UploadFile, HTTPException
from fastapi.responses import JSONResponse
import os
import tempfile
from agent_prog import (
    agent_file_processor,
    check_pdf_type,
    normal_pdf_processor,
    extract_text_to_markdown,
    convert_docx_to_temp_pdf,
    ppt_to_pdf_win32com,
    xlsx_to_mrkdwn,
    csv_to_mrkdwn,
    txt_to_mrkdwn,
    extract_text_to_tempfile
)

app = FastAPI(
    title="Any File to Markdown Converter",
    description="API to convert various file formats to markdown text",
    version="1.0.0"
)

@app.post("/convert-to-markdown/")
async def convert_to_markdown(file: UploadFile):
    try:
        # Create a temporary file to store the uploaded content
        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(file.filename)[1]) as temp_file:
            content = await file.read()
            temp_file.write(content)
            temp_file_path = temp_file.name

        # Get the file extension and process accordingly
        file_type = agent_file_processor(temp_file_path)
        
        if file_type == '.pdf':
            if check_pdf_type(temp_file_path) == "scanned":
                markdown_text = extract_text_to_markdown(temp_file_path, lang="eng", dpi=300)
            else:
                markdown_text = normal_pdf_processor(temp_file_path)
        elif file_type == '.docx':
            pdf_file_path = convert_docx_to_temp_pdf(temp_file_path)
            markdown_text = extract_text_to_markdown(pdf_file_path, lang="eng", dpi=300)
            os.unlink(pdf_file_path)  # Clean up temporary PDF
        elif file_type == '.pptx':
            pptx_file_path = ppt_to_pdf_win32com(temp_file_path)
            markdown_text = extract_text_to_markdown(pptx_file_path, lang="eng", dpi=300)
            os.unlink(pptx_file_path)  # Clean up temporary PDF
        elif file_type == '.xlsx':
            markdown_text = xlsx_to_mrkdwn(temp_file_path)
        elif file_type == '.csv':
            markdown_text = csv_to_mrkdwn(temp_file_path)
        elif file_type == '.txt':
            markdown_text = txt_to_mrkdwn(temp_file_path)
        elif file_type == '.png':
            temp_txt_path = extract_text_to_tempfile(temp_file_path)
            markdown_text = txt_to_mrkdwn(temp_txt_path)
            os.unlink(temp_txt_path)  # Clean up temporary text file
        else:
            raise HTTPException(status_code=400, detail="Unsupported file format")

        # Clean up the temporary uploaded file
        os.unlink(temp_file_path)

        return JSONResponse(
            content={
                "markdown_text": markdown_text,
                "file_type": file_type
            }
        )

    except Exception as e:
        # Clean up temporary file in case of error
        if 'temp_file_path' in locals():
            os.unlink(temp_file_path)
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/")
async def root():
    return {"message": "Welcome to File to Markdown Converter API"} 
