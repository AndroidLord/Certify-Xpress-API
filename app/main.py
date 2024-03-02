from fastapi import FastAPI
from routers import certification

app = FastAPI()


@app.get("/")
async def root():
    return {"message": "This is the root of the CertifyXpresAPI, please use the /docs endpoint to see the API documentation."}


app.include_router(certification.router, prefix="/api/v1")

