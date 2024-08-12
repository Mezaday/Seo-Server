import uvicorn
from core.settings import settings

if __name__ == "__main__":
    try:
        uvicorn.run(
            app='core.app:app',
            host=settings.UVICORN_HOST,
            port=settings.UVICORN_PORT,
            reload=settings.UVICORN_RELOAD
        )
    except Exception as e:
        print(e)