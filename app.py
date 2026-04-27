import os
import shutil
import subprocess
import tempfile
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response, JSONResponse

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

        with open(in_path, "wb") as f:
            f.write(await file.read())

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

        # LibreOffice обычно пишет output с тем же именем
        candidates = [
            os.path.join(tmp, "input.xlsx"),
            os.path.join(tmp, "normalized.xlsx"),
        ]
        out_path = next((p for p in candidates if os.path.exists(p)), None)

        if not out_path:
            return JSONResponse(
                status_code=500,
                content={
                    "error": "Converted file not found",
                    "stdout": proc.stdout,
                    "stderr": proc.stderr,
                },
            )

        with open(out_path, "rb") as f:
            content = f.read()

    out_name = (file.filename[:-5] if file.filename.lower().endswith(".xlsx") else file.filename) + "_fixed.xlsx"

    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={
            "Content-Disposition": f'attachment; filename="{out_name}"'
        },
    )
