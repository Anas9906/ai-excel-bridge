import os
import shutil
import subprocess


def main():
    print("Building standalone executable with PyInstaller...")

    subprocess.run([
        "python", "-m", "PyInstaller",
        "-y",
        "--noconsole",
        "--onedir",
        "--name", "GIECO_Insurance_Sync",
        # Windows COM / pywin32 hidden imports
        "--hidden-import", "win32timezone",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        # New OCR pipeline modules
        "--hidden-import", "ocr_provider",
        "--hidden-import", "cache_manager",
        # Google GenAI
        "--hidden-import", "google.genai",
        "--hidden-import", "google.genai.types",
        # Image handling
        "--hidden-import", "PIL",
        "--hidden-import", "PIL.Image",
        "gui_app.py"
    ], check=True)

    print("Build successful. Copying .env file to dist folder...")

    dist_dir = os.path.join("dist", "GIECO_Insurance_Sync")
    env_src = ".env"
    env_dst = os.path.join(dist_dir, ".env")

    if os.path.exists(env_src):
        shutil.copy2(env_src, env_dst)
        print(f"Copied {env_src} to {env_dst}")
    else:
        print(f"Warning: {env_src} not found. Create it from .env.example before running the exe.")

    print("Packaging complete! Executable at:")
    print(os.path.abspath(dist_dir))


if __name__ == "__main__":
    main()
