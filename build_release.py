import shutil
import subprocess
import sys
from pathlib import Path


def cleanup(paths: list[Path]) -> None:
    for path in paths:
        if path.exists():
            if path.is_dir():
                shutil.rmtree(path)
            else:
                path.unlink()


def main() -> None:
    project_root = Path(__file__).resolve().parent
    cleanup([project_root / "build", project_root / "dist"])

    result = subprocess.run(
        [sys.executable, "-m", "PyInstaller", "build_release.spec"],
        cwd=project_root,
        capture_output=True,
        text=True,
    )

    if result.returncode != 0:
        print(result.stdout)
        print(result.stderr)
        raise SystemExit(1)

    print(result.stdout)
    checksum_result = subprocess.run(
        [sys.executable, "generate_checksums.py"],
        cwd=project_root,
        capture_output=True,
        text=True,
    )

    if checksum_result.returncode != 0:
        print(checksum_result.stdout)
        print(checksum_result.stderr)
        raise SystemExit(1)

    print(checksum_result.stdout)
    print("Release build completed successfully")


if __name__ == "__main__":
    main()
