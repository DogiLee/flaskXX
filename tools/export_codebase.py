"""Export the reviewable PDGM source tree as one Markdown document."""

from __future__ import annotations

import argparse
import re
from datetime import datetime
from pathlib import Path


INCLUDED_SUFFIXES = {
    ".py", ".html", ".css", ".js", ".txt", ".json", ".toml",
    ".yaml", ".yml", ".ini", ".cfg", ".bat", ".ps1", ".sh",
}
INCLUDED_NAMES = {".gitignore", "requirements.txt"}
EXCLUDED_DIRECTORIES = {
    ".git", ".idea", ".investigation", ".pytest_cache", ".venv", "venv", "__pycache__",
    "artifact_work", "data", "node_modules", "outputs", "tests", "uploads", "yedekler",
}
LANGUAGES = {
    ".py": "python", ".html": "html", ".css": "css", ".js": "javascript",
    ".json": "json", ".ps1": "powershell", ".bat": "bat",
    ".sh": "bash", ".yaml": "yaml", ".yml": "yaml", ".toml": "toml",
}


def dahil_mi(path: Path, root: Path, output: Path) -> bool:
    """Return whether a single source/documentation file belongs in the export."""
    if path.resolve() == output:
        return False
    relative = path.relative_to(root)
    if any(part in EXCLUDED_DIRECTORIES for part in relative.parts[:-1]):
        return False
    if path.name.startswith("pdgm_codebase_export."):
        return False
    if path.name.startswith(".env"):
        return False
    return path.suffix.lower() in INCLUDED_SUFFIXES or path.name in INCLUDED_NAMES


def kod_blogu(metin: str) -> str:
    """Pick a Markdown fence that cannot be closed by source content."""
    uzunluk = max((len(parca) for parca in re.findall(r"`+", metin)), default=0)
    return "`" * max(3, uzunluk + 1)


def dil(path: Path) -> str:
    return LANGUAGES.get(path.suffix.lower(), "text")


def export_et(root: Path, output: Path) -> int:
    root = root.resolve()
    output = output.resolve()
    dosyalar = sorted(
        (path for path in root.rglob("*") if path.is_file() and dahil_mi(path, root, output)),
        key=lambda path: path.relative_to(root).as_posix().lower(),
    )

    satirlar = [
        "# PDGM Codebase Export",
        "",
        f"- Üretim zamanı: {datetime.now().astimezone().isoformat(timespec='seconds')}",
        f"- Kök dizin: `{root}`",
        f"- Dahil edilen dosya sayısı: {len(dosyalar)}",
        "- Hariç tutulanlar: testler, yardımcı çalışma araçları, Markdown dokümanları, çalışma verileri, yüklemeler, yedekler, çıktı raporları, inceleme kopyaları, sanal ortamlar, önbellekler ve export dosyaları.",
        "",
        "## Dosya listesi",
        "",
    ]
    satirlar.extend(f"- `{path.relative_to(root).as_posix()}`" for path in dosyalar)
    satirlar.extend(["", "## Dosya içerikleri", ""])

    for path in dosyalar:
        relative = path.relative_to(root).as_posix()
        try:
            metin = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            satirlar.extend([f"## `{relative}`", "", "_UTF-8 olmayan/binary dosya: içerik dışa aktarılmadı._", ""])
            continue
        fence = kod_blogu(metin)
        satirlar.extend([f"## `{relative}`", "", f"{fence}{dil(path)}", metin.rstrip("\n"), fence, ""])

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text("\n".join(satirlar), encoding="utf-8")
    return len(dosyalar)


def main() -> None:
    varsayilan_kok = Path(__file__).resolve().parents[1]
    parser = argparse.ArgumentParser(description="PDGM kaynak ağacını tek Markdown dosyasına aktarır.")
    parser.add_argument("--root", type=Path, default=varsayilan_kok, help="Aktarılacak proje kökü")
    parser.add_argument("--output", type=Path, help="Üretilecek Markdown yolu")
    args = parser.parse_args()

    root = args.root.resolve()
    if not root.is_dir():
        parser.error(f"Kök dizin bulunamadı: {root}")
    output = (args.output or root / "pdgm_codebase_export.md").resolve()
    sayi = export_et(root, output)
    print(f"{sayi} dosya dışa aktarıldı: {output}")


if __name__ == "__main__":
    main()
