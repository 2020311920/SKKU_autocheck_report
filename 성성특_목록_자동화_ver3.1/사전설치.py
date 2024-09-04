import subprocess
import sys

# List of packages to be installed
# 의미 없는 수정입니다.
packages = [
    "altgraph==0.17.4",
    "Cython==3.0.10",
    "packaging==24.0",
    "pefile==2023.2.7",
    "pip==24.0",
    "pyinstaller==6.7.0",
    "pyinstaller-hooks-contrib==2024.6",
    "pywin32-ctypes==0.2.2",
    "setuptools==70.0.0",
    "olefile",
    "pdfplumber",
    "pywin32"
]

def install_packages():
    for package in packages:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        except subprocess.CalledProcessError as e:
            print(f"Failed to install package: {package}. Error: {e}")

# Call the install_packages function
install_packages()

# Your main code starts here
print("All packages installed successfully!")
input("아무키나 눌러주세요")
