# Tessainator V6

Program łączy zamówienia formatek na podstawie wspólnego formatu, dat dostawy, kodów FEFCO, minimalnej partii i dopuszczalnej przycinki.

## Uruchomienie

1. Kliknij dwukrotnie `uruchom_tessainator_v6.bat`.
2. Ustaw zasady na kartach **Podstawowe** i **Zaawansowane**.
3. Kliknij **Zapisz i wybierz plik**.
4. Wskaż plik Excel z zamówieniami.
5. Otwórz utworzony plik `_polaczone_formatki.xlsx`.

Program nie zmienia pliku wejściowego i nie łączy się z bazą danych.

## Kolejność pracy z wynikiem

1. `PODSUMOWANIE` — wynik całego przebiegu.
2. `GRUPY DO ZAMÓWIENIA` — jeden wiersz oznacza jedno sugerowane wspólne zamówienie.
3. `POZYCJE` — formatki należące do grup oraz formaty niepołączone.
4. `ODRZUCONE` i `OSTRZEŻENIA` — dane wymagające sprawdzenia.
5. `USTAWIENIA` — wartości użyte podczas obliczeń i opis zasady działania każdej opcji.

Arkusze `OPTYMALIZACJA`, `DOPASOWANIA TECHNICZNE` i `DANE TECHNICZNE` służą do kontroli szczegółów algorytmu.

## Ważne opcje

- Obrót o 90° należy włączać tylko wtedy, gdy pozwala na to kierunek fali, druku i wykrojnika.
- Łączenie różnych FEFCO może utworzyć grupę z kilku konstrukcji. Zgodność technologiczną należy potwierdzić przed zamówieniem.
