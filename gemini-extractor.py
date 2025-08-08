import argparse
import subprocess
import sys
import shutil
from pathlib import Path
import pyzipper as zipfile
import rarfile
from google import genai
from rich.console import Console
from rich.live import Live
import winreg

console = Console()

def info(msg):
    console.print(f"[cyan][!][/cyan] {msg}")

def success(msg):
    console.print(f"[green][✓][/green] {msg}")

def warning(msg):
    console.print(f"[yellow][!][/yellow] {msg}")

def error(msg):
    console.print(f"[red][✗][/red] {msg}")

def find_winrar_path():
    """Find WinRAR installation path from registry."""
    reg_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\WinRAR.exe"
    ]
    for reg_path in reg_paths:
        try:
            with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path) as key:
                value, _ = winreg.QueryValueEx(key, None)  # (default) value
                if Path(value).exists():
                    return str(Path(value))
        except FileNotFoundError:
            continue
    return None

def archive_requires_password(archive_path: Path) -> bool:
    """Check if archive is password-protected."""
    try:
        if archive_path.suffix.lower() == ".zip":
            with zipfile.ZipFile(archive_path) as zf:
                for info in zf.infolist():
                    if info.flag_bits & 0x1:
                        return True
            return False
        elif archive_path.suffix.lower() == ".rar":
            with rarfile.RarFile(archive_path) as rf:
                return rf.needs_password()
    except Exception as e:
        error(f"Failed to check password protection: {e}")
        sys.exit(1)

def get_password_from_comment(comment_text: str):
    """Send the archive's comment text to Gemini and try to extract the password."""
    if not comment_text:
        warning("No comment found in archive.")
        return None

    prompt = (
        "You are given a text extracted from the comment section of a compressed file. "
        "The comment may contain a password. "
        "Carefully read the text and identify the password exactly as it appears. "
        "If no password is present, say 'NO_PASSWORD_FOUND'. "
        "Output only the password without any extra words or formatting."
        f"\n\nText:\n---\n{comment_text}\n---"
    )

    try:
        client = genai.Client()
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt
        )
        password_candidate = response.text.strip()
    except Exception as e:
        error(f"Gemini request failed: {e}")
        return None

    if password_candidate and password_candidate != "NO_PASSWORD_FOUND":
        success(f"Password detected: {password_candidate}")
        return password_candidate
    else:
        warning("No password found in archive comment.")
        return None

def run_winrar_with_progress(winrar_path, cmd_args, output_folder):
    """Run WinRAR command and show progress live."""
    try:
        cmd = [winrar_path] + cmd_args
        process = subprocess.Popen(
            cmd, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            text=True, bufsize=1
        )
        with Live(console=console, refresh_per_second=4) as live:
            for line in process.stdout:
                live.update(f"[cyan]{line.strip()}[/cyan]")
        process.wait()
        if process.returncode != 0:
            raise subprocess.CalledProcessError(process.returncode, cmd)
    except subprocess.CalledProcessError:
        if output_folder.exists():
            shutil.rmtree(output_folder, ignore_errors=True)
        raise

def extract_archive(archive_path: str):
    archive_path = Path(archive_path)

    if not archive_path.exists():
        error(f"File not found: {archive_path}")
        return
    if archive_path.suffix.lower() not in ['.zip', '.rar']:
        error("Only .zip and .rar files are supported.")
        return

    winrar_path = find_winrar_path()
    print(winrar_path)
    if not winrar_path:
        error("WinRAR not found. Please install WinRAR.")
        return
    else:
        info(f"Using WinRAR from: {winrar_path}")

    output_folder = archive_path.with_suffix('')

    info(f"Checking if '{archive_path.name}' requires a password...")
    if not archive_requires_password(archive_path):
        info("No password required. Extracting directly...")
        try:
            run_winrar_with_progress(
                winrar_path,
                ['x', str(archive_path), f'{output_folder}\\'],
                output_folder
            )
            success("Extraction completed successfully (no password).")
        except Exception:
            error("Extraction failed.")
        return

    info("Password protection detected. Reading comment...")
    comment_text = None
    try:
        if archive_path.suffix.lower() == '.zip':
            with zipfile.ZipFile(archive_path) as zip_file:
                if zip_file.comment:
                    comment_text = zip_file.comment.decode('utf-8', 'ignore')
        elif archive_path.suffix.lower() == '.rar':
            with rarfile.RarFile(archive_path) as rar_file:
                comment_text = rar_file.comment
    except Exception as e:
        error(f"Failed to read archive comment: {e}")
        return

    if not comment_text:
        warning("No comment found — skipping extraction.")
        return

    password = get_password_from_comment(comment_text)
    if not password:
        warning("Password not detected — skipping extraction.")
        return

    info(f"Extracting '{archive_path.name}' with detected password...")
    try:
        run_winrar_with_progress(
            winrar_path,
            ['x', f'-p{password}', str(archive_path), f'{output_folder}\\'],
            output_folder
        )
        success("Extraction completed successfully.")
    except Exception:
        error("Extraction failed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Extract archive with or without password")
    parser.add_argument("filename", help="Path to the archive file (.zip or .rar)")
    args = parser.parse_args()

    extract_archive(args.filename)
    input("\nPress Enter to exit...")
