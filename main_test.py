import pytest
import main
from configparser import ConfigParser


def test_simpleqr():
    #generate function
    text = "AA AA\nBB BB\nCC CC"
    url = "google.com/=name"
    assert main.SimpleQR(text,url).generate(replace=False) == True
    assert main.SimpleQR(text,url).generate(invert=True) == True

def test_config():

    cfg = ConfigParser()
    keys = ['text','url','invert']
    values = ["A A\nB B\nC C", 'google.com/=name', True]
    values2 = []


    #saving config
    for n in range(len(keys)):
        main.Config('cfg.ini').save(keys[n], values[n])
    for n in range(len(keys)):
        cfg.read('cfg.ini')
        values2.append(cfg.get('main', keys[n]))
        if values2[-1] in ['True', 'False']:
            if values2[-1] == 'True':
                values2[-1] = True
            else:
                values2[-1] = False
    assert values == values2


    #loading config
    config = main.Config('cfg.ini')
    config.load()
    values2 = [config.text, config.url, config.invert]
    assert values == values2

test_simpleqr()
test_config()
