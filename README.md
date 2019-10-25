Program do zbierania informacji o wydatkach z banku (aktualnie dostępny tylko iPKO) i kontroli budżetu domowego.

Skrzynce pocztowej gmail należy ustawić dostęp logowania dla mniej zaufanych aplikacji.
Konfigurujemy plik config.txt (instrukcja poniżej).
Program pobiera wszystkie maile z Gmaila, które odpowiadają za powiadomienia wydatków z banku.
Z maili wyłuskiwane są dane na temat kwoty, daty oraz godziny.
Dane wrzucane są do pliku xlsx "Fundusz_Domowy", w którym tworzone są arkusze na każdy miesiąc.
Plik xlsx jest zapisywany i na bieżąco aktualizowany.
Pobrane maile z powiadomieniami usuwane są ze skrzynki gmail (powiadomienia zaśmiecają niepotrzebnie inbox gmail)

Jak aktualizować plik xlsx Fundusz_Domowy?
Polecam ustawić w Task Schedluer (Harmonogram zadań) cykliczne odpalanie się programu raz na dobę. Należy pamiętać, by w takim przypadku przekopiować plik config.txt do miejsca, z którego wywoływany jest task (defaultowo jest to dysk systemowy, ścieżka C:\config.txt).

Konfiguracja pliku config.txt:

Password - hasło do gmaila
Mail - adres email
Bank - adres email banku, z którego przychodzą powiadomienia (dla iPKO jest to 'powiadomienia@pkobp.pl')
ExcelLang - PL dla polskiej wersji, ENG dla angielskiej (różnica we wstawianiu formuł do spreadsheeta)
SaveDir - ścieżka do miejsca, gdzie zapisywany i aktualizowany ma być plik "Fundusz_Domowy"

Stack technologiczny:
- .NET Framework
- GemBox.Spreadsheet
- EAGetMail
