from cx_Freeze import setup, Executable

# Iekļaujam papildu failus
include_files = ["config.json", "Temp.xlsx"]

# Norādām galveno skriptu
target = Executable(
    script="PDFtoINVOICE.py",  # Jūsu galvenais skripts
    base="Win32GUI",          # Izmantojiet "Win32GUI" GUI programmām, vai atstājiet tukšu konsoles aplikācijām
    target_name="AutoOgre Rekins.exe",  # Izveidotā .exe faila nosaukums
    icon="C:/Users/janis/Documents/Programmas/dist/attels.ico"   # Norādiet ikonas failu
)

# cx_Freeze konfigurācija
setup(
    name="AutoOgre Rekins",
    version="1.2",
    description="AutoOgre Rēķins",
    options={
        "build_exe": {
            "include_files": include_files,  # Papildu faili, kas jāiekļauj
            "packages": ["openpyxl"],        # Papildu bibliotēkas
        }
    },
    executables=[target]
)