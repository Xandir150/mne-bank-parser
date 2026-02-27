from pathlib import Path

from pydantic_settings import BaseSettings


class Settings(BaseSettings):
    data_dir: Path = Path("/data")
    input_dir: Path = Path("/data/input")
    processed_dir: Path = Path("/data/processed")
    output_dir: Path = Path("/data/output")
    db_path: Path = Path("/data/db/statements.db")

    scan_interval: int = 60  # seconds

    # Bank code -> bank name mapping
    bank_names: dict[str, str] = {
        "520": "Hipotekarna Banka",
        "530": "NLB Banka",
        "535": "Prva Banka CG",
        "540": "Erste Bank",
        "560": "Universal Capital Bank",
        "565": "Lovćen Banka",
        "570": "Zapad Banka",
        "575": "Ziraat Bank",
        "580": "Adriatic Bank",
    }

    # Supported file extensions per bank code
    # 540 (Erste) also accepts .htm/.html
    supported_extensions: dict[str, list[str]] = {
        "520": [".pdf"],
        "530": [".pdf"],
        "535": [".pdf"],
        "540": [".pdf", ".htm", ".html"],
        "560": [".pdf"],
        "565": [".pdf"],
        "570": [".pdf"],
        "575": [".pdf"],
        "580": [".pdf"],
    }

    class Config:
        env_prefix = "IZVOD_"


settings = Settings()
