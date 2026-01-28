from fastapi import FastAPI, HTTPException
from pydantic import BaseModel
from typing import Optional, Dict, Any
import worker_sap
import sys
import io

# Setup stdout for Windows
if sys.platform == 'win32':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

app = FastAPI()

class Order(BaseModel):
    botId: str
    params: Dict[str, Any] = {}
    filePath: Optional[str] = None

@app.get("/status")
def get_status():
    return {"status": "online", "mode": "local_bridge"}

@app.post("/execute")
def execute_bot(order: Order):
    print(f"📥 [LOCAL SERVER] Recibida orden: {order.botId}")
    try:
        # Call the refactored logic from worker_sap
        # We pass params as the third arg
        result = worker_sap.run_automation(order.botId, order.filePath, order.params)
        return {"status": "success", "result": result}
    except Exception as e:
        print(f"❌ [LOCAL SERVER] Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    # Run on port 8000
    uvicorn.run(app, host="0.0.0.0", port=8000)
