from fastapi import APIRouter, HTTPException
from pydantic import BaseModel
from app.services.ai_service import AIService
from app.services.excel_interpreter import ExcelInterpreter

router = APIRouter()
ai_service = AIService()
excel_interpreter = ExcelInterpreter()

class QueryRequest(BaseModel):
    query: str
    context: dict  # Excel context (selected range, sheet data, etc.)

class QueryResponse(BaseModel):
    action: str
    parameters: dict
    explanation: str
    office_js_code: str

@router.post("/query", response_model=QueryResponse)
async def process_query(request: QueryRequest):
    try:
        ai_response = await ai_service.interpret_query(
            request.query, 
            request.context
        )

        excel_action = excel_interpreter.generate_action(ai_response)
        
        return excel_action
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/create-chart")
async def create_chart(request: QueryRequest):
    """Generate chart configuration"""
    try:
        # Get chart config from AI
        chart_config = await ai_service.generate_chart(
            request.query,
            request.context
        )
        
        # Wrap it in the proper response format
        ai_response = {
            "action": "chart",
            "parameters": chart_config,
            "explanation": f"Creating a {chart_config.get('chartType', 'chart')} chart with the specified data"
        }
        
        # Process through interpreter to get Office.js code
        excel_action = excel_interpreter.generate_action(ai_response)
        
        return excel_action
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/generate-formula")
async def generate_formula(request: QueryRequest):
    """Generate Excel formula from natural language"""
    try:
        formula = await ai_service.generate_formula(
            request.query,
            request.context
        )
        return {"formula": formula}
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@router.post("/create-pivot-table")
async def create_pivot_table(request: QueryRequest):
    """Generate pivot table configuration"""
    try:
        pivot_config = await ai_service.generate_pivot_table(
            request.query,
            request.context
        )
        return pivot_config
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))