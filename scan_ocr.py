"""Read text locally with Tesseract or Windows OCR."""
import base64
import os
from pathlib import Path
import re
import shutil
import subprocess

from image_reader import scan_bytes
from kyc import FIELDS


WINDOWS_OCR = r"""
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false)
Add-Type -AssemblyName System.Runtime.WindowsRuntime
$null = [Windows.Media.Ocr.OcrEngine, Windows.Foundation, ContentType=WindowsRuntime]
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($null -eq $engine) { throw 'No Windows OCR language is available.' }
$inputText = [Console]::In.ReadToEnd().Trim()
if (-not $inputText) { exit 0 }

$null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Storage.Streams.DataWriter, Windows.Storage.Streams, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Graphics.Imaging.SoftwareBitmap, Windows.Graphics.Imaging, ContentType=WindowsRuntime]
$null = [Windows.Media.Ocr.OcrResult, Windows.Foundation, ContentType=WindowsRuntime]
$asTask = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethodDefinition -and
    $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq ('IAsyncOperation' + [char]96 + '1')
} | Select-Object -First 1
function Wait-WinRT($operation, [Type] $type) {
    $task = $asTask.MakeGenericMethod($type).Invoke($null, @($operation))
    $task.Wait()
    $task.Result
}

$stream = [Windows.Storage.Streams.InMemoryRandomAccessStream]::new()
$writer = [Windows.Storage.Streams.DataWriter]::new($stream)
try {
    $writer.WriteBytes([Convert]::FromBase64String($inputText))
    $null = Wait-WinRT ($writer.StoreAsync()) ([uint32])
    $null = $writer.DetachStream()
    $stream.Seek(0)
    $decoder = Wait-WinRT ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
    $bitmap = Wait-WinRT ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
    try {
        $result = Wait-WinRT ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
        foreach ($line in $result.Lines) { [Console]::WriteLine($line.Text) }
    } finally {
        $bitmap.Dispose()
    }
} finally {
    $writer.Dispose()
    $stream.Dispose()
}
"""


def run_windows_ocr(executable, payload):
    command = base64.b64encode(WINDOWS_OCR.encode("utf-16-le")).decode("ascii")
    return subprocess.run(
        [executable, "-NoProfile", "-NonInteractive", "-EncodedCommand", command],
        input=base64.b64encode(payload), capture_output=True, timeout=60,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )


def windows_ocr_path():
    if os.name == "nt":
        powershell = Path(os.environ.get("SystemRoot", r"C:\Windows")) / (
            "System32/WindowsPowerShell/v1.0/powershell.exe"
        )
        if powershell.is_file():
            return str(powershell)
    return None


def find_ocr_engine():
    executable = shutil.which("tesseract")
    if executable:
        return executable
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "LOCALAPPDATA"):
        directory = os.environ.get(variable)
        if directory:
            candidate = Path(directory) / "Tesseract-OCR" / "tesseract.exe"
            if candidate.is_file():
                return str(candidate)
    if os.name == "nt":
        powershell = windows_ocr_path()
        if powershell:
            try:
                result = run_windows_ocr(powershell, b"")
            except (OSError, subprocess.TimeoutExpired):
                pass
            else:
                if result.returncode == 0:
                    return powershell
        raise ValueError("Windows OCR is unavailable. Enable a Windows OCR language or install Tesseract.")
    raise ValueError("Install local Tesseract OCR to read record IDs from the decrypted images.")


tesseract_path = find_ocr_engine


def ocr_text(executable, payload, page_mode, *, preserve_spacing=False):
    if Path(executable).name.casefold() == "powershell.exe":
        result = run_windows_ocr(executable, payload)
        if result.returncode:
            raise ValueError("Windows OCR could not read this image.")
        return result.stdout.decode("utf-8", errors="replace")
    command = [executable, "stdin", "stdout", "--psm", str(page_mode)]
    if preserve_spacing:
        command.extend(["-c", "preserve_interword_spaces=1"])
    result = subprocess.run(
        command,
        input=payload, capture_output=True, timeout=60,
        env={**os.environ, "OMP_THREAD_LIMIT": "1"},
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if result.returncode:
        raise ValueError("OCR could not read this image.")
    return result.stdout.decode("utf-8", errors="replace")


def read_fields(path, key):
    return parse_fields(ocr_text(find_ocr_engine(), scan_bytes(path, key), 6))


def parse_fields(text):
    labels = {field.casefold(): field for field in FIELDS if field not in {"Image Name", "Remark"}}
    labels.update({"record no.": "Record No", "registration address": "Rag Address",
                   "father name": "Father/Husband Name", "husband name": "Father/Husband Name",
                   "dob": "Date Of Birth", "city": "City/District/Town"})
    fields = {}
    for line in text.splitlines():
        match = re.fullmatch(r"\s*([^:]+):\s*(\S.*?)\s*", line)
        if match:
            label, value = match.groups()
            field = labels.get(label.strip().casefold())
            if field:
                fields.setdefault(field, value)
    return fields
