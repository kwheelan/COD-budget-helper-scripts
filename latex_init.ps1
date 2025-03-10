Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process

# Add the path to the MiKTeX binaries to the current session's PATH
$miktexPath = "C:\Users\katrina.wheelan\AppData\Local\Programs\MiKTeX\miktex\bin\x64"
$env:PATH = "$miktexPath;$env:PATH"

# Optional: Output the updated PATH to confirm the addition
Write-Output "Updated PATH: $env:PATH"

# You can also test if pdflatex is now recognized
$result = & pdflatex --version
if ($?) {
    Write-Output "pdflatex is now available in PATH."
} else {
    Write-Output "Failed to recognize pdflatex. Check the PATH and try again."
}

Write-Output "Path setup script completed successfully."