from fastapi import FastAPI 
from fastapi.middleware.cors import CORSMiddleware

from api.generate import gen_router
from api.upload import upl_router
from core.logger import setup_logger

#from core.configurations import settings

setup_logger()

def create_app() -> FastAPI:
    # setup_logging()

    app = FastAPI(
        title="Docx-Generation",
        version="1.0.0"
    )

    # Middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=True,
        allow_methods=["*"],
        allow_headers=["*"],
    )

    # Routers
    app.include_router(gen_router)#prefix=settings.API_PREFIX)
    app.include_router(upl_router)

    return app

app = create_app()

if __name__ == "__main__":
    app.run("0.0.0.0",debug=True)
