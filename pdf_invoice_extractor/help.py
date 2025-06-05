import os

def nolasit_rekina_numuru(fails="rekina_numurs.txt"):
    """
    Nolasa pēdējo rēķina numuru no faila. Ja fails neeksistē, atgriež sākotnējo vērtību.
    """
    if os.path.exists(fails):
        with open(fails, "r") as f:
            try:
                return int(f.read().strip())
            except ValueError:
                return 1  # Ja fails ir tukšs vai bojāts, sāk no 1
    return 1  # Ja fails neeksistē, sāk no 1

def saglabat_rekina_numuru(rekina_numurs, fails="rekina_numurs.txt"):
    """
    Saglabā jauno rēķina numuru failā.
    """
    with open(fails, "w") as f:
        f.write(str(rekina_numurs))