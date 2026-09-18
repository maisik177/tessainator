# Tessainator V6

Tessainator analizuje zamówienia formatek z arkusza i proponuje ich łączenie we wspólne grupy zakupowe. Uwzględnia format, FEFCO, terminy dostaw, minimalną wielkość partii, dopuszczalną przycinkę, orientację formatki oraz odpad. Wynik zapisuje do nowego pliku Excel, nie zmieniając danych wejściowych.

## Możliwości

- import XLSX, XLS, XLSM, ODS i CSV;
- grupowanie według zgodności technicznej i terminów;
- minima liczby sztuk i powierzchni w m²;
- kontrola przycinki według pola, boków albo obu kryteriów;
- opcjonalny obrót o 90° i łączenie różnych FEFCO;
- specjalna obsługa alternatywnych formatów FEFCO 201;
- mnożnik dla FEFCO 300 i 301;
- podział grup, jeśli ogranicza łączny odpad;
- raport odrzuconych pozycji, ostrzeżeń i ustawień.

## Wymagania i instalacja

- Windows 10/11;
- Python 3.10+ z `tkinter`;
- `pandas` i `openpyxl`;
- opcjonalnie `odfpy` dla wejścia ODS.

```powershell
py -m pip install pandas openpyxl odfpy
```

## Uruchomienie

Kliknij `uruchom_tessainator_v6.bat` albo wykonaj:

```powershell
py laczenie_zamowien_formatki.py
```

Następnie ustaw zasady, kliknij **Zapisz i wybierz plik**, wskaż arkusz i otwórz utworzony plik `<nazwa>_polaczone_formatki.xlsx`. Program działa lokalnie, nie łączy się z bazą danych ani Internetem.

## Dane wejściowe

Nazwy kolumn nie uwzględniają wielkości liter.

| Kolumna | Znaczenie |
|---|---|
| `fefco` | kod konstrukcji FEFCO |
| `dimension_x` | pierwszy wymiar formatki w mm |
| `dimension_y` | drugi wymiar formatki w mm |
| `quantity` | ilość bazowa |
| `delivery_date` | preferowana data dostawy |

Jeżeli nie ma `delivery_date`, program próbuje użyć m.in. `data_dostawy`, `termin_dostawy`, `order_created_at`, `created_at` lub `order_date` i zapisuje ostrzeżenie.

Opcjonalne kolumny FEFCO 201:

- `has_alternative_dimensions`, `alternative_dimension_x`, `alternative_dimension_y`;
- `has_alternative_dimensions_2`, `alternative_dimension_2_x`, `alternative_dimension_2_y`.

Błędne daty, FEFCO, wymiary i ilości trafiają do arkusza `ODRZUCONE` z podanym powodem, zamiast zatrzymywać cały przebieg.

## Ustawienia

Ustawienia są zapamiętywane w `ustawienia_laczenia.json`.

| Ustawienie | Domyślnie | Znaczenie |
|---|---:|---|
| maksymalna różnica dat | 16 dni | największy odstęp w grupie |
| maksymalna przycinka | 10% | dopuszczalna przycinka/odpad |
| minimum sztuk | 90 | minimum efektywnych formatek |
| minimum m² | 290 | minimum powierzchni grupy |
| warunek minimum | `OR` | wystarczy minimum sztuk albo m²; `AND` wymaga obu |
| tryb przycinki | `area` | dostępne również `side` i `both` |
| obrót 90° | wyłączony | pozwala łączyć długość z szerokością |
| różne FEFCO | wyłączone | pozwala grupować różne konstrukcje |
| optymalizacja odpadu | włączona | dzieli grupę, gdy zmniejsza odpad |

Obrót wymaga sprawdzenia kierunku fali, druku i wykrojnika. Grupy z różnymi FEFCO wymagają potwierdzenia technologicznego.

## Wynik

1. `PODSUMOWANIE` — wynik przebiegu i wartości kontrolne.
2. `GRUPY DO ZAMÓWIENIA` — sugerowane wspólne zamówienia.
3. `POZYCJE` — przypisanie każdej formatki.
4. `ODRZUCONE` — błędne wiersze z przyczyną.
5. `OSTRZEŻENIA` — problemy niewstrzymujące obliczeń.
6. `USTAWIENIA` — parametry użyte w przebiegu.

Arkusze `OPTYMALIZACJA`, `DOPASOWANIA TECHNICZNE` i `DANE TECHNICZNE` pokazują szczegóły algorytmu. Wynik jest propozycją optymalizacji, a nie automatyczną decyzją produkcyjną.

## Testy

```powershell
py -m unittest -v test_laczenie_zamowien_formatki.py
```

Testy obejmują obrót, różne FEFCO, warianty FEFCO 201, przycinkę, odrzucanie danych, optymalizację odpadu i awaryjny zapis.

## Pliki

- `laczenie_zamowien_formatki.py` — aplikacja i algorytm;
- `uruchom_tessainator_v6.bat` — start w Windows;
- `ustawienia_laczenia.json` — zapamiętane parametry;
- `test_laczenie_zamowien_formatki.py` — testy.

Plik wejściowy pozostaje bez zmian. Gdy poprzedni wynik jest otwarty w Excelu, nowy rezultat otrzymuje nazwę z datą i godziną.
