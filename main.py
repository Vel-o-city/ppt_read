from fastapi import FastAPI, HTTPException, Form
from pydantic import BaseModel
import requests
from io import BytesIO
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
import pytesseract
import boto3
from datetime import datetime
import os
from dotenv import load_dotenv

load_dotenv()
aws_access_key_id = os.getenv("AWS_ACCESS_KEY_ID")
aws_secret_access_key = os.getenv("AWS_SECRET_ACCESS_KEY")
aws_region = os.getenv("AWS_REGION")
aws_s3_bucket_name = os.getenv("AWS_S3_BUCKET_NAME")

s3 = boto3.client('s3',
                  aws_access_key_id=aws_access_key_id,
                  aws_secret_access_key=aws_secret_access_key,
                  region_name=aws_region)

app = FastAPI()

class ExtractImagesRequest(BaseModel):
    url: str

@app.post("/ppt_kaushik")
async def extract_and_process(request: ExtractImagesRequest):
    try:
        result = extract_images_from_pptx_and_upload_to_s3(request.url)
        return result
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

def extract_images_from_pptx_and_upload_to_s3(pptx_s3_url, crop_percentage=7):
    response = requests.get(pptx_s3_url)
    pptx_bytes = BytesIO(response.content)
    pptx_bytes.seek(0)
    presentation = Presentation(pptx_bytes)
    result_list = []
    for slide_number, slide in enumerate(presentation.slides, start=1):
        slide_title = None
        for shape in slide.shapes:
            if shape.has_text_frame:
                slide_title = shape.text_frame.text.strip().replace(" ", "_")
                break
        if not slide_title:
            slide_title = f"slide_{slide_number}"
        for shape_number, shape in enumerate(slide.shapes, start=1):
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image_stream = BytesIO(shape.image.blob)
                image = Image.open(image_stream)
                width, height = image.size
                crop_height = int(height * crop_percentage / 100)
                crop_box = (0, height - crop_height, width, height)
                cropped_image = image.crop(crop_box)
                res = pytesseract.image_to_string(cropped_image)
                text = "*" + res + "*"
                remaining_image = image.crop((0, 0, width, height - crop_height))
                remaining_image_bytes = BytesIO()
                remaining_image.save(remaining_image_bytes, format=image.format)
                remaining_image_bytes.seek(0)
                folder_structure = 'prod/ldoc/inventoryManagement'
                timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                s3_key = f"{folder_structure}/{slide_title}_shape_{shape_number}_{timestamp}.{image.format.lower()}"
                s3.upload_fileobj(remaining_image_bytes, aws_s3_bucket_name, s3_key)
                s3_url = f"https://{aws_s3_bucket_name}.s3.{aws_region}.amazonaws.com/{s3_key}"
                result_list.append({text: s3_url})
    return result_list

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
