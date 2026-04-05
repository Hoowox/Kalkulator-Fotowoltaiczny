# Symulator produkcji energii fotowoltaicznej

Aplikacja desktopowa w Pythonie służąca do symulacji produkcji energii elektrycznej oraz analizy efektywności energetycznej.


## Funkcje
* **Symulacja produkcji:** Obliczanie uzysków energii na podstawie zadanych parametrów.
* **Przetwarzanie danych:** Obsługa plików wejściowych w formatach **Excel** oraz **JSON**.
* **Wizualizacje:** 
     * Interaktywne wykresy słupkowe.
     * Zaawansowane wizualizacje danych w formacie **3D**.
* **GUI:** Przyjazny interfejs użytkownika zbudowany w Tkinter.


## Technologia
Program został zbudowany z wykorzystaniem:
* **Python** (rdzeń aplikacji)
* **Tkinter** – interfejs graficzny.
* **Pandas** – zaawansowana analiza i manipulacja danymi.
* **Matplotlib** – generowanie wykresów i wizualizacji 2D/3D.


## Wymagania systemowe

- System operacyjny: Windows 7 / 8 / 10 / 11 (64-bit lub 32-bit)
- Brak dodatkowego oprogramowania (Python, biblioteki itp.)
- Ok. 250 MB wolnego miejsca na dysku (głównie na plik wykonywalny)

  
## Instalacja i uruchomienie (Windows)

Program jest dostarczany jako gotowy plik wykonywalny `.exe`. **Nie wymaga instalacji Pythona ani dodatkowych bibliotek**.

1. Pobierz repozytorium na swój komputer:
   - Kliknij zielony przycisk **`<> Code`** na stronie GitHub, a następnie wybierz **`Download ZIP`**.
   - Lub jeśli masz zainstalowanego gita:  
     `git clone https://github.com/TwojaNazwaUzytkownika/nazwa-repo.git`

2. Wypakuj pobrane archiwum ZIP (jeśli pobierałeś jako ZIP).

3. Otwórz folder `dist` – to tam znajduje się plik wykonywalny.

4. Uruchom program **`Kalkulator Fotowoltaiczny.exe`**:
   - Kliknij na plik dwukrotnie lewym przyciskiem myszy.
   - Jeśli system Windows wyświetli ostrzeżenie o nieznanym wydawcy, kliknij **„Więcej informacji”**, a następnie **„Uruchom mimo to”** (program jest bezpieczny).


## Prezentacja autorskiego programu

Po uruchomieniu programu otrzymujemy interfejs główny przedstawiony na Rys. 1.1. Składa się on z kilku zdefiniowanych sekcji, z których każda pełni określoną rolę. W lewym panelu użytkownik ma możliwość wprowadzania kluczowych danych wejściowych, takich jak liczba paneli, kąt nachylenia paneli, azymut, moc pojedynczego panelu, sprawność paneli, temperatura otoczenia, liczba godzin słonecznych w skali roku, wartość albedo, poziom zacienienia oraz zanieczyszczeń. Po wprowadzeniu danych, użytkownik może skorzystać z przycisku "Oblicz", który uruchamia skrypt dokonujący obliczeń. W rezultacie, uzyskany wynik jest prezentowany na dolnym panelu po lewej stronie interfejsu. Dzięki temu użytkownik ma natychmiastowy dostęp do rezultatów swoich działań. Prawa sekcja interfejsu zajmuje miejsce na prezentację wyników w formie graficznej, w postaci wykresu słupkowego lub wizualizacji 3D. W przypadku skorzystania z opcji wykresu użytkownik może analizować ilość energii wygenerowanej w poszczególnych miesiącach, przedstawionej za pomocą słupków.

![Główny interfejs programu – lewy panel z danymi wejściowymi, prawy panel z wykresami, dolny panel z wynikiem](https://github.com/user-attachments/assets/6d9e2fcc-69d7-4d32-8200-e19c6456f350)

**Rys. 1.1** Interfejs użytkownika

W sekcji „Ustawienia” przedstawionej na Rys. 1.2 użytkownik może wprowadzić dodatkowe dane w postaci średnich miesięcznych temperatur oraz ilości godzin słonecznych, dzięki czemu możliwe jest uzyskanie dokładniejszych wyników. W sytuacji, gdy użytkownik nie skorzysta z funkcji ustawienia temperatur miesięcznych, konieczne będzie podanie jednej ogólnej temperatury dla całego roku.

![Sekcja ustawień programu – formularz do wprowadzania średnich miesięcznych temperatur i godzin słonecznych](https://github.com/user-attachments/assets/f9d52897-cbaa-453e-b25a-d37bd2b32cdf)

**Rys. 1.2** Sekcja ustawień programu

Kolejną zakładkę reprezentowaną przez Rys. 1.3 stanowi wybór widoku. Po wprowadzeniu danych oraz obliczeniu wyniku użytkownik będzie mógł zobaczyć wykres, a także wizualizację budynku w 3D wraz z planowanymi modułami PV. Warto zaznaczyć, że całość jest w pełni interaktywna. Domek ukazany na wizualizacji obraca się samoczynnie. Ponadto możemy go zatrzymać, bądź obracać ręcznie za pomocą myszki.

![Wizualizacja 3D budynku z modułami PV oraz okno z wykresem wyników symulacji](https://github.com/user-attachments/assets/49194baf-6dca-4364-b956-6a561fc257db)

**Rys. 1.3** Wizualizacja projektu instalacji oraz wyników symulacji

W sekcji „Zapisz” oraz „Importuj” przedstawionej na poniższym Rys. 1.4 możemy skorzystać z wymienionych funkcji w celu zapisania, bądź zaimportowania wprowadzanych danych, ustawień lub uzyskanych wykresów. Możliwość ta znacznie ułatwia pracę w przypadku kilku projektów, różnych danych, bądź konieczności dzielenia się wynikami z innymi użytkownikami.

![Okno programu z przyciskami do zapisu i importu danych, ustawień oraz wykresów](https://github.com/user-attachments/assets/4d0ceb51-1433-4ba8-82fe-ca8ede3611b7)

**Rys. 1.4** Zapis danych

Poza wyżej wymienionymi funkcjami program posiada także inne oczywiste funkcje, m.in.:

- zakładkę „Pomoc”, za pośrednictwem której można otworzyć okno z podstawowymi informacjami dotyczącymi obsługi programu (można je także wywołać, klikając w dowolnym momencie klawisz F1 na klawiaturze),
- podstawową funkcję częściowego sprawdzania poprawności danych, która uniemożliwia stworzenie wykresu na podstawie danych, które są niepoprawne (np. zacienienie o wartości powyżej 99%) lub w przypadku próby wprowadzenia wyrazów w miejscu, gdzie program oczekuje cyfr.


## Przykładowe dane i import

W folderze `pomocne pliki` (w repozytorium) znajdują się pliki, które możesz wykorzystać do testowania i nauki obsługi programu.

### Import danych pomiarowych (Excel)

1. W programie przejdź do zakładki **„importuj”**.
2. Kliknij przycisk **„importuj dane”**.
3. Wskaż plik `DANEZ2023.xlsx` z folderu `pomocne pliki`.
4. Dane zostaną wczytane i będą dostępne do dalszej analizy.



### Import ustawień lub wykresów

Analogicznie postępuj w przypadku importu wcześniej zapisanych ustawień symulacji lub wykresów:
- Wybierz odpowiednią opcję w zakładce **„importuj”** (np. „importuj ustawienia”, „importuj wykres”).
- Wskaż plik z folderu `pomocne pliki`.

## Autor
Bartosz Pawlak

## Licencja
MIT

