import os
import shutil
import subprocess
import tempfile
import traceback
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import Response, JSONResponse

app = FastAPI()

@app.exception_handler(Exception)
async def all_exceptions_handler(request, exc):
    return JSONResponse(
        status_code=500,
        content={
            "error": str(exc),
            "traceback": traceback.format_exc(),
        },
    )

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/normalize-xlsx")
async def normalize_xlsx(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported")

    soffice = shutil.which("libreoffice") or shutil.which("soffice")
    if not soffice:
        raise RuntimeError("LibreOffice binary not found")

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input.xlsx")

        with open(in_path, "wb") as f:
            f.write(await file.read())

        proc = subprocess.run(
            [
                soffice, "--headless", "--nologo", "--nodefault",
                "--nolockcheck", "--norestore",
                "--convert-to", "xlsx", "--outdir", tmp, in_path
            ],
            capture_output=True, text=True
        )

        if proc.returncode != 0:
            raise RuntimeError(f"LibreOffice failed. stdout={proc.stdout} stderr={proc.stderr}")

        candidates = [os.path.join(tmp, "input.xlsx"), os.path.join(tmp, "normalized.xlsx")]
        out_path = next((p for p in candidates if os.path.exists(p)), None)
        if not out_path:
            raise RuntimeError(f"Converted file not found. stdout={proc.stdout} stderr={proc.stderr}")

        with open(out_path, "rb") as f:
            content = f.read()

    out_name = file.filename[:-5] + "_fixed.xlsx"
    return Response(
        content=content,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename="{out_name}"'},
    )
