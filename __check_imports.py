mods = ['main','downloader','login','email_outlook','email_smtp','utils']
import importlib, sys
for m in mods:
    try:
        importlib.import_module(m)
        print(m, 'OK')
    except Exception as e:
        print(m, 'ERR', type(e).__name__, e)
