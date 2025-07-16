import os
import shutil
from fastapi import FastAPI, UploadFile, File, Form
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from conversor import (
    converter_ics_para_csv,
    converter_ics_para_json,
    converter_mbox_para_csv,
    converter_mbox_para_json
)

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Ajuste para seu domínio Vercel se necessário
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.post("/convert")
async def convert_file(file: UploadFile = File(...), origem: str = Form(...), destino: str = Form(...)):
    os.makedirs("temp", exist_ok=True)
    input_path = os.path.join("temp", file.filename)

    with open(input_path, "wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    if origem == "ICS" and destino == "CSV":
        sucesso, output_path = converter_ics_para_csv(input_path, "temp")
    elif origem == "ICS" and destino == "JSON":
        sucesso, output_path = converter_ics_para_json(input_path, "temp")
    elif origem == "MBOX" and destino == "CSV":
        sucesso, output_path = converter_mbox_para_csv(input_path, "temp")
    elif origem == "MBOX" and destino == "JSON":
        sucesso, output_path = converter_mbox_para_json(input_path, "temp")
    else:
        return {"erro": "Conversão não suportada"}

    if sucesso:
        return FileResponse(output_path, filename=os.path.basename(output_path), media_type="application/octet-stream")
    else:
        return {"erro": output_path}
