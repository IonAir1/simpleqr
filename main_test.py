import pytest
import main

def test_simpleQR():
    text = "AA AA\nBB BB\nCC CC"
    url = "google.com/=name"
    assert main.SimpleQR(text,url).generate() == True

test_simpleQR()