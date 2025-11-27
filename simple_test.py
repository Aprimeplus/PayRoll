import sys
print("Python OK")
try:
    import pandas
    print("Pandas OK")
except ImportError:
    print("Pandas Missing")

try:
    import tkinter
    print("Tkinter OK")
except ImportError:
    print("Tkinter Missing")

try:
    import psycopg2
    print("Psycopg2 OK")
except ImportError:
    print("Psycopg2 Missing")
