from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.routers.ai_routers import router

app = FastAPI(title="Excel AI Agent API")

# CORS for Excel Add-in
app.add_middleware(
    CORSMiddleware,
    allow_origins=["https://localhost:3000"],  # Your add-in URL
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

app.include_router(router, prefix="/api/v1", tags=["ai"])

@app.get("/")
async def root():
    return {"message": "Excel AI Agent API"}