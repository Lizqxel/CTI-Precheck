import hashlib
from pathlib import Path


def sha256_of_file(file_path: Path) -> str:
    hash_obj = hashlib.sha256()
    with file_path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 256), b""):
            hash_obj.update(chunk)
    return hash_obj.hexdigest()


def main() -> None:
    project_root = Path(__file__).resolve().parent
    dist_dir = project_root / "dist"
    if not dist_dir.exists():
        raise SystemExit("dist フォルダが見つかりません。先に build_release.py を実行してください。")

    exe_files = sorted([p for p in dist_dir.glob("*.exe") if p.is_file()])
    if not exe_files:
        raise SystemExit("dist に EXE が見つかりません。")

    checksum_lines: list[str] = []
    for exe_path in exe_files:
        checksum = sha256_of_file(exe_path)
        checksum_lines.append(f"{checksum}  {exe_path.name}")

    checksum_path = dist_dir / "checksums.txt"
    checksum_path.write_text("\n".join(checksum_lines) + "\n", encoding="utf-8")

    print(f"Generated: {checksum_path}")
    for line in checksum_lines:
        print(line)


if __name__ == "__main__":
    main()
