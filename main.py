from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from docx import Document
import requests
from supabase import create_client
import json
from datetime import datetime
import os
import tempfile
from typing import List

# ثوابت الأعمدة
TABLE_COLUMNS = {
    'INDEX': 0,
    'GIVEN_NAME': 1,
    'FATHER_NAME': 2,
    'SURNAME': 3,
    'PASSPORT_NO': 4,
    'DATE_OF_BIRTH': 5,
    'NATIONALITY': 6,
    'RESIDENCY_COUNTRY': 7
}

# التحقق من متغيرات البيئة
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")

if not SUPABASE_URL or not SUPABASE_KEY:
    raise ValueError("SUPABASE_URL و SUPABASE_KEY مطلوبة")

supabase = create_client(SUPABASE_URL, SUPABASE_KEY)
app = FastAPI()

class Traveler(BaseModel):
    given_name: str
    father_name: str
    surname: str
    passport_no: str
    date_of_birth: str
    nationality: str
    residency_country: str

class FillApprovalFormRequest(BaseModel):
    travel_group_id: str
    travelers: List[Traveler]

@app.post("/fill-approval-form")
async def fill_approval_form(request: FillApprovalFormRequest):
    try:
        # تحميل template من Supabase
        template_url = "https://hifkuvyvhrxmcgkbvgqo.supabase.co/storage/v1/object/public/form-templates/template-1787915478648-310160860.docx"
        template_response = requests.get(template_url, timeout=10)
        template_response.raise_for_status()
        
        # فتح المستند
        doc = Document(template_response.content)
        
        # البحث عن الجدول (الجدول الثاني)
        if len(doc.tables) < 2:
            raise ValueError("الـ template لا يحتوي على جدول")
        
        table = doc.tables[1]
        
        # ملء الجدول
        for idx, traveler in enumerate(request.travelers, 1):
            if idx > len(table.rows) - 1:
                break
            
            row = table.rows[idx]
            
            row.cells[TABLE_COLUMNS['INDEX']].text = str(idx)
            row.cells[TABLE_COLUMNS['GIVEN_NAME']].text = traveler.given_name
            row.cells[TABLE_COLUMNS['FATHER_NAME']].text = traveler.father_name
            row.cells[TABLE_COLUMNS['SURNAME']].text = traveler.surname
            row.cells[TABLE_COLUMNS['PASSPORT_NO']].text = traveler.passport_no
            row.cells[TABLE_COLUMNS['DATE_OF_BIRTH']].text = traveler.date_of_birth
            row.cells[TABLE_COLUMNS['NATIONALITY']].text = traveler.nationality
            row.cells[TABLE_COLUMNS['RESIDENCY_COUNTRY']].text = traveler.residency_country
        
        # حفظ مؤقتاً
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = tmp.name
            doc.save(tmp_path)
        
        # رفع على Supabase Storage
        timestamp = datetime.now().timestamp()
        file_name = f"forms/approval_form_{timestamp}.docx"
        
        try:
            with open(tmp_path, "rb") as f:
                supabase.storage.from_("form-templates").upload(file_name, f.read())
        except Exception as upload_error:
            raise RuntimeError(f"فشل رفع الملف: {str(upload_error)}")
        finally:
            os.remove(tmp_path)
        
        # الرابط العام
        file_url = f"{SUPABASE_URL}/storage/v1/object/public/form-templates/{file_name}"
        
        return {
            "status": "success",
            "file_url": file_url,
            "travelers_count": len(request.travelers)
        }
    
    except requests.RequestException as req_error:
        raise HTTPException(status_code=500, detail=f"خطأ في تحميل الـ template: {str(req_error)}")
    except ValueError as val_error:
        raise HTTPException(status_code=400, detail=str(val_error))
    except RuntimeError as runtime_error:
        raise HTTPException(status_code=500, detail=str(runtime_error))
    except Exception as error:
        raise HTTPException(status_code=500, detail=f"خطأ عام: {str(error)}")

@app.get("/health")
async def health():
    return {"status": "ok"}
