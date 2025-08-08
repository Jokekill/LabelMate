# LabelMate – Doplňek Microsoft Office pro klasifikaci dokumentů

**LabelMate** je jednoduchý doplněk pro Microsoft Office, který uživatelům umožňuje snadno klasifikovat dokumenty pomocí volitelných klasifikačních značek (např. TLP). Pomáhá tak zavádět a dodržovat pravidla vnitřní klasifikace dat.

## Funkce
- Flexibilní klasifikační schéma: značení dokumentů podle libovolně definovaných úrovní klasifikace.
- Jednoduché vložení klasifikační značky přímo do záhlaví dokumentu (Word) nebo do zvolených polí (Excel, PowerPoint).
- Banner v taskpane při neklasifikovaném dokumentu s výzvou k označení, který se po klasifikaci automaticky odstraní.
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
Tento repozitář obsahuje první funkční verzi doplňku **LabelMate**.

### Omezení prototypu:
- Funguje pouze pro **Microsoft Word**
- Základní, neupravené uživatelské rozhraní
- Tři přednastavené klasifikační značky v češtině:
  - `TLP:Interní`
  - `TLP:Chráněný`
  - `TLP:PřísněChráněný`

## Plánovaný rozvoj
Budoucí vývoj se zaměří na rozšíření funkcí, větší integraci a zlepšení UX.

### Funkce k implementaci:
- Rozšíření podpory na Excel a PowerPoint
- Automatická detekce jazyka dokumentu a přizpůsobení značek (EN, CZ, SK)
- Možnost manuální volby jazyka klasifikace
- Banner v taskpane, který upozorní na chybějící klasifikaci a po jejím vložení zmizí
- Modernizované uživatelské rozhraní s podporou světlého/tmavého režimu (TailwindCSS)
- Vkládání klasifikační značky jako needitovatelný content control
- Automatické zobrazení taskpane při otevření neklasifikovaného dokumentu
- Tooltipy a odkazy na interní dokumentaci u klasifikačních voleb
- Uživatelská a administrátorská dokumentace

## Struktura souborů
```
|-- manifest.xml      # manifest doplňku
|-- taskpane.html     # UI taskpane + JavaScript logika
|-- assets/
    |-- icon.png      # ikona doplňku
```

## Dokumentace
[Office JavaScript API](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/javascript-api-for-office) 
