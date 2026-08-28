from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from docx import Document
from docx.shared import Pt, RGBColor
import requests
from supabase import create_client
import json
from datetime import datetime
import os
import tempfile
from typing import List, Dict, Any

# Supabase
SUPABASE_URL = os.getenv("SUPABASE_URL")
SUPABASE_KEY = os.getenv("SUPABASE_KEY")
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
        
        if template_response.status_code != 200:
            raise Exception("فشل تحميل الـ template")
        
        # فتح المستند
        doc = Document(template_response.content)
        
        # البحث عن الجدول (الجدول الثاني)
        if len(doc.tables) < 2:
            raise Exception("الـ template لا يحتوي على جدول")
        
        table = doc.tables[1]
        
        # ملء الجدول
        for idx, traveler in enumerate(request.travelers, 1):
            if idx > len(table.rows) - 1:
                break
            
            row = table.rows[idx]
            
            # الخلايا:
            # 0: ت (الرقم)
            # 1: GIVEN NAME
            # 2: THE FATHER NAME
            # 3: SURNAME
            # 4: PASSPORT NO
            # 5: DATE OF BIRTH
            # 6: NATIONALITY
            # 7: RESIDENGE COUNTRY
            
            row.cells[0].text = str(idx)
            row.cells[1].text = traveler.given_name
            row.cells[2].text = traveler.father_name
            row.cells[3].text = traveler.surname
            row.cells[4].text = traveler.passport_no
            row.cells[5].text = traveler.date_of_birth
            row.cells[6].text = traveler.nationality
            row.cells[7].text = traveler.residency_country
        
        # حفظ مؤقتاً
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
            tmp_path = tmp.name
            doc.save(tmp_path)
        
        # رفع على Supabase Storage
        timestamp = datetime.now().timestamp()
        file_name = f"forms/approval_form_{timestamp}.docx"
        
        with open(tmp_path, "rb") as f:
            supabase.storage.from_("form-templates").upload(file_name, f.read())
        
        # الرابط العام
        file_url = f"{SUPABASE_URL}/storage/v1/object/public/form-templates/{file_name}"
        
        # حذف المؤقت
        os.remove(tmp_path)
        
        return {
            "status": "success",
            "file_url": file_url,
            "travelers_count": len(request.travelers)
        }
    
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/health")
async def health():
    return {"status": "ok"}
