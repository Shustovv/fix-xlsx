import os
import subprocess
import tempfile
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse

app = FastAPI()

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/normalize-xlsx")
async def normalize_xlsx(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx is supported")

    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input.xlsx")
        out_path = os.path.join(tmp, "input.xlsx")

        with open(in_path, "wb") as f:
            f.write(await file.read())

        cmd = [
            "libreoffice",
            "--headless",
            "--convert-to", "xlsx",
            "--outdir", tmp,
            in_path
        ]
        proc = subprocess.run(cmd, capture_output=True, text=True)

        if proc.returncode != 0 or not os.path.exists(out_path):
            raise HTTPException(
                status_code=500,
                detail=f"Conversion failed: {proc.stderr or proc.stdout}"
            )

        return FileResponse(
            out_path,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=file.filename.replace(".xlsx", "_fixed.xlsx")
        )
