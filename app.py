import os
import shutil
import subprocess
import tempfile
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse, JSONResponse

app = FastAPI()

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/normalize-xlsx")
async def normalize_xlsx(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported")

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        return JSONResponse(status_code=500, content={"error": "LibreOffice binary not found"})

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input.xlsx")
        out_path = os.path.join(tmp, "normalized.xlsx")

        data = await file.read()
        with open(in_path, "wb") as f:
            f.write(data)

        cmd = [
            soffice,
            "--headless",
            "--nologo",
            "--nodefault",
            "--nolockcheck",
            "--norestore",
            "--convert-to", "xlsx",
            "--outdir", tmp,
            in_path,
        ]

        proc = subprocess.run(cmd, capture_output=True, text=True)

        # LibreOffice часто пишет output в stdout, even on success
        # И обычно имя = input.xlsx, поэтому учитываем оба варианта
        produced_default = os.path.join(tmp, "input.xlsx")

        if proc.returncode != 0:
            return JSONResponse(
                status_code=500,
                content={
                    "error": "Conversion failed",
                    "returncode": proc.returncode,
                    "stdout": proc.stdout,
                    "stderr": proc.stderr,
                },
            )

        if os.path.exists(out_path):
            final_path = out_path
        elif os.path.exists(produced_default):
            final_path = produced_default
        else:
            return JSONResponse(
                status_code=500,
                content={
                    "error": "Converted file not found",
                    "stdout": proc.stdout,
                    "stderr": proc.stderr,
                },
            )

        return FileResponse(
            final_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=file.filename.replace(".xlsx", "_fixed.xlsx"),
        )
