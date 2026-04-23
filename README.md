# LabelMate – Doplňek Microsoft Office pro klasifikaci dokumentů

**LabelMate** je jednoduchý doplněk pro Microsoft Office, který uživatelům umožňuje snadno klasifikovat dokumenty pomocí volitelných klasifikačních značek (např. TLP). Pomáhá tak zavádět a dodržovat pravidla vnitřní klasifikace dat.

## Funkce
- Flexibilní klasifikační schéma: značení dokumentů podle libovolně definovaných úrovní klasifikace.
- Jednoduché vložení klasifikační značky přímo do záhlaví dokumentu (Word) nebo do zvolených polí (Excel, PowerPoint).
- Banner v taskpane při neklasifikovaném dokumentu s výzvou k označení, který se po klasifikaci automaticky odstraní.
- Podpora více jazyků (autodetekce jazyka systému, možnost ručního přepnutí).
- Deployment přes Microsoft 365 Admin Center.
- Hostování je navrženo přes GitHub Pages. 

## Instalace
1. Nasazení přes Admin Centrum Microsoft 365
   - Použijte soubor [`manifest.xml`](./manifest.xml)
   - Nahrajte ho v části **Integrated apps**
2. Nebo ruční aktivace pro vývoj/testování
   - Postupujte dle [oficiálního návodu Microsoftu](https://learn.microsoft.com/cs-cz/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins)

## Aktuální stav prototypu
Tento repozitář obsahuje funkční verzi doplňku **LabelMate**.

### Podporované aplikace:
- **Microsoft Word** – klasifikační značka je vložena do záhlaví jako needitovatelný content control.
- **Microsoft Excel** – značka je vložena do středu záhlaví stránky (page header) pro všechny listy.
- **Microsoft PowerPoint** – značka je vložena jako textový rámec v zápatí každého slidu; při kontrole klasifikace se automaticky sjednotí napříč slidy (doplní chybějící footery a opraví překlepy podle první validní klasifikace z databáze).

### Podporované jazyky:
- Čeština (CZ), Angličtina (EN), Slovenština (SK)
- Autodetekce podle jazyka Office, možnost ručního přepnutí v taskpane.
- Tři přednastavené klasifikační úrovně v každém jazyce (`TLP:Interní` / `TLP:Chráněný` / `TLP:PřísněChráněný` a jejich ekvivalenty).

### Uživatelské rozhraní:
- Světlý a tmavý režim s autodetekcí dle Office / OS a možností ručního přepnutí.
- Tooltipy s popisem jednotlivých úrovní a odkazy na interní dokumentaci.

## Plánovaný rozvoj
Budoucí vývoj se zaměří na dokončení zbývajících funkcí a zlepšení UX.

### Funkce k implementaci:
- Automatické zobrazení taskpane při otevření neklasifikovaného dokumentu
- Uživatelská a administrátorská dokumentace

## Struktura souborů
```
|-- manifest.xml                    # manifest doplňku
|-- taskpane.html                   # UI taskpane
|-- app.js                          # bootstrap aplikace a jazykové UI
|-- core.js                         # i18n a UI utility
|-- i18n.js                         # jazykové mutace a definice labelů
|-- theme.js                        # světlý/tmavý režim
|-- banner.js                       # banner pro neklasifikovaný dokument
|-- labels.js                       # render klasifikačních tlačítek a tooltipů
|-- classification.host.js          # router pro Word / Excel / PowerPoint
|-- classification.word.js          # implementace pro Word
|-- classification.excel.js         # implementace pro Excel
|-- classification.powerpoint.js    # implementace pro PowerPoint
|-- styles.core.css                 # základní styly a proměnné
|-- styles.components.css           # styly komponent
|-- styles.ui.css                   # styly UI prvků
|-- assets/
    |-- icon.png                    # ikona doplňku
```

## Dokumentace
[Office JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office)