"""Install packages as defined in this file into the Python environment."""
from setuptools import setup, find_packages

setup(
    name="prefillqr",
    url="https://github.com/IonAir1/qrcode",
    description="creates a qr code to a pre filled g forms",
    version="0.0.1",
    packages=find_packages(where=".", exclude=["tests"]),
    install_requires=[
        "setuptools>=45.0",
        "bitlyshortener>=0.6.1",
        "pandas>=1.4.2",
        "Pillow>=9.1.1",
        "qrcode",
        "openpyxl",
        "pyautogui",
    ],
)
