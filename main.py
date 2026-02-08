from fastapi import FastAPI 
from fastapi.middleware.cors import CORSMiddleware

from api.generate import gen_router
from api.upload import upl_router



def create_app() -> FastAPI:
    # setup_logging()

    app = FastAPI(
        title="Docx-Generation",
        version=None
    )

    # Middleware
    app.add_middleware(
        CORSMiddleware,
        allow_origins=settings.CORS_ORIGINS,
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
    app.run(debug=True, port="0.0.0.0")
