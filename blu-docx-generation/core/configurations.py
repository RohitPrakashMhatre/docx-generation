from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    # App
    APP_NAME: str = "Docx-Generation"
    VERSION: str = "1.0.0"
    DEBUG: bool = False

    # Database (hosted)
    DATABASE_HOST: str
    DATABASE_PORT: int = 3306
    DATABASE_NAME: str
    DATABASE_USERNAME: str
    DATABASE_PASSWORD: str

    # Static
    TABLE_NAME: str = "docx_json"

    model_config = SettingsConfigDict(
        env_file=".env",
        env_file_encoding="utf-8",
    )


settings = Settings()
