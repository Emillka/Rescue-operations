import datetime
import openpyxl
from openpyxl import Workbook

class AkcjaRatownicza:
    def __init__(self, rodzaj_zdarzenia, pojazdy, miejscowosc, data, warunki_atmosferyczne, temperatura, pora_roku):
        self.rodzaj_zdarzenia = rodzaj_zdarzenia
        self.pojazdy = pojazdy
        self.miejscowosc = miejscowosc
        self.data = data
        self.warunki_atmosferyczne = warunki_atmosferyczne
        self.temperatura = temperatura
        self.pora_roku = pora_roku

    def wyswietl_informacje(self):
        print(f"Rodzaj zdarzenia: {self.rodzaj_zdarzenia}")
        print(f"Pojazdy: {', '.join(self.pojazdy)}")
        print(f"Miejscowość: {self.miejscowosc}")
        print(f"Data: {self.data}")
        print(f"Warunki atmosferyczne: {', '.join(self.warunki_atmosferyczne)}")
        print(f"Temperatura: {self.temperatura}")
        print(f"Pora roku: {self.pora_roku}")

def wprowadz_date():
    while True:
        data_input = input("Podaj datę (RRRR-MM-DD): ")
        try:
            data = datetime.datetime.strptime(data_input, "%Y-%m-%d").date()
            return data
        except ValueError:
            print("Błędny format daty. Spróbuj ponownie.")

def zapisz_do_excel(akcje, nazwa_pliku="akcje_ratownicze.xlsx"):
    try:
        # Spróbuj otworzyć istniejący plik Excel
        workbook = openpyxl.load_workbook(nazwa_pliku)
        sheet = workbook.active
    except FileNotFoundError:
        # Jeśli plik nie istnieje, utwórz nowy plik
        workbook = Workbook()
        sheet = workbook.active
        # Nagłówki kolumn (tylko jeśli plik został utworzony nowy)
        sheet.append(["Rodzaj zdarzenia", "Pojazdy", "Miejscowość", "Data", "Warunki atmosferyczne", "Temperatura", "Pora_roku"])

    for akcja in akcje:
        # Dostosuj formatowanie pojazdów i warunków atmosferycznych
        pojazdy = ', '.join(akcja.pojazdy)
        warunki_atmosferyczne = ', '.join(akcja.warunki_atmosferyczne)
        sheet.append([akcja.rodzaj_zdarzenia, pojazdy, akcja.miejscowosc, akcja.data, warunki_atmosferyczne, akcja.temperatura, akcja.pora_roku])

    workbook.save(nazwa_pliku)


def wczytaj_z_excel(nazwa_pliku="akcje_ratownicze.xlsx"):
    akcje = []
    try:
        workbook = openpyxl.load_workbook(nazwa_pliku)
        sheet = workbook.active

        for row in sheet.iter_rows(min_row=2, values_only=True):
            akcje.append(AkcjaRatownicza(*row))

    except FileNotFoundError:
        print("Plik nie istnieje. Tworzony nowy plik.")
    return akcje

if __name__ == "__main__":
    akcje_ratownicze = wczytaj_z_excel()

    try:
        while True:
            print("\n1. Dodaj nową akcję ratowniczą")
            print("2. Wyświetl wszystkie akcje ratownicze")
            print("3. Zapisz do pliku Excel")
            print("4. Zakończ")

            wybor = input("Wybierz opcję: ")

            if wybor == "1":
                print("Wybierz rodzaj zdarzenia:")
                print("1 - Wypadek")
                print("2 - Pożar wewnętrzny")
                print("3 - Pożar zewnętrzny")
                print("4 - Poszukiwanie osoby zaginionej")
                print("5 - Sprawdzenie monitoringu")
                print("6 - Zabezpieczenie rejonu działań PSP")
                print("7 - Ćwiczenia")
                print("8 - Inne")

                wybor_zdarzenia = input("Wybierz numer rodzaju zdarzenia: ")

                if wybor_zdarzenia == "8":
                    rodzaj_zdarzenia = input("Podaj inny rodzaj zdarzenia: ")
                    komentarz = input("Dodaj komentarz: ")
                    rodzaj_zdarzenia += f" ({komentarz})"
                elif wybor_zdarzenia in ["1", "2", "3", "4", "5", "6", "7"]:
                    rodzaj_zdarzenia = {
                        "1": "Wypadek",
                        "2": "Pożar wewnętrzny",
                        "3": "Pożar zewnętrzny",
                        "4": "Poszukiwanie osoby zaginionej",
                        "5": "Sprawdzenie monitoringu",
                        "6": "Zabezpieczenie rejonu działań PSP",
                        "7": "Ćwiczenia"
                    }[wybor_zdarzenia]
                else:
                    print("Niepoprawny wybór. Spróbuj ponownie.")
                    continue

                miejscowosc = input("Podaj miejscowość: ")
                data = wprowadz_date()

                print("Wybierz pojazdy:")
                print("1 - GLBA 0,4/2")
                print("2 - GBA 2,5/16")
                print("3 - GCBA 5/32")
                print("4 - D (drabina)")
                print("5 - H (podnośnik hydrauliczny)")

                wybrane_pojazdy = []
                while True:
                    wybor_pojazdu = input("Wybierz numer pojazdu (wpisz 'koniec' jeśli zakończyłeś wybór): ")
                    if wybor_pojazdu.lower() == "koniec":
                        break
                    elif wybor_pojazdu in ["1", "2", "3", "4", "5"]:
                        if wybor_pojazdu == "1":
                            wybrane_pojazdy.append("GLBA 0,4/2")
                        elif wybor_pojazdu == "2":
                            wybrane_pojazdy.append("GBA 2,5/16")
                        elif wybor_pojazdu == "3":
                            wybrane_pojazdy.append("GCBA 5/32")
                        elif wybor_pojazdu == "4":
                            wybrane_pojazdy.append("D (drabina)")
                        elif wybor_pojazdu == "5":
                            wybrane_pojazdy.append("H (podnośnik hydrauliczny)")
                    else:
                        print("Niepoprawny wybór. Spróbuj ponownie.")

                print("Wybierz warunki atmosferyczne:")
                print("1 - Wiatr")
                print("2 - Deszcz")
                print("3 - Śnieg")
                print("4 - Mróz")
                print("5 - Warunki ok")

                wybrane_warunki = []
                while True:
                    wybor_warunkow = input("Wybierz numer warunków atmosferycznych (wpisz 'koniec' jeśli zakończyłeś wybór): ")
                    if wybor_warunkow.lower() == "koniec":
                        break
                    elif wybor_warunkow == "5":
                        wybrane_warunki = ["Warunki ok"]
                        break
                    elif wybor_warunkow in ["1", "2", "3", "4"]:
                        if wybor_warunkow == "1":
                            wybrane_warunki.append("Wiatr")
                        elif wybor_warunkow == "2":
                            wybrane_warunki.append("Deszcz")
                        elif wybor_warunkow == "3":
                            wybrane_warunki.append("Śnieg")
                        elif wybor_warunkow == "4":
                            wybrane_warunki.append("Mróz")
                    else:
                        print("Niepoprawny wybór. Spróbuj ponownie.")

                temperatura = input("Podaj temperaturę: ")

                print("Podsumowanie wprowadzonych danych:")
                print(f"Rodzaj zdarzenia: {rodzaj_zdarzenia}")
                print(f"Miejscowość: {miejscowosc}")
                print(f"Data: {data}")
                print(f"Pojazdy: {', '.join(wybrane_pojazdy)}")
                print(f"Warunki atmosferyczne: {', '.join(wybrane_warunki)}")

                nowa_akcja = AkcjaRatownicza(rodzaj_zdarzenia, wybrane_pojazdy, miejscowosc, data, wybrane_warunki, temperatura, "")
                akcje_ratownicze.append(nowa_akcja)
                print("Dodano nową akcję ratowniczą.")

            elif wybor == "2":
                for akcja in akcje_ratownicze:
                    akcja.wyswietl_informacje()

            elif wybor == "3":
                zapisz_do_excel(akcje_ratownicze)
                print("Dane zapisane do pliku Excel.")

            elif wybor == "4":
                break

            else:
                print("Niepoprawny wybór. Spróbuj ponownie.")

    except KeyboardInterrupt:
        print("\nPrzerwano przez użytkownika. Zamykanie programu.")
