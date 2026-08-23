import os
import shutil
import subprocess

def main():
    print("Building standalone executable with PyInstaller...")
    
    # Run PyInstaller
    subprocess.run([
        "python", "-m", "PyInstaller",
        "-y",
        "--noconsole",
        "--onedir",
        "--hidden-import", "win32timezone",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--name", "GIECO_Insurance_Sync",
        "gui_app.py"
    ], check=True)
    
    print("Build successful. Copying .env file to dist folder...")
    
    # Copy .env to dist/GIECO_Insurance_Sync
    dist_dir = os.path.join("dist", "GIECO_Insurance_Sync")
    env_src = ".env"
    env_dst = os.path.join(dist_dir, ".env")
    
    if os.path.exists(env_src):
        shutil.copy2(env_src, env_dst)
        print(f"Copied {env_src} to {env_dst}")
    else:
        print(f"Warning: {env_src} not found in the current directory.")
        
    print("Packaging complete! You can find the executable at:")
    print(os.path.abspath(dist_dir))

if __name__ == "__main__":
    main()
