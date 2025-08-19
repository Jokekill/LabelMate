// i18n.js — UI texty + definice klasifikací (s help/helpHtml)
window.LM_I18N = {
  EN: {
    appTitle: "LabelMate for Word",
    themeLabel: "Theme",
    langLabel: "Language",
    choosePrompt: "Choose classification level:",
    banner: {
      title: "No classification set",
      desc: "Choose a classification below. The banner will disappear after it is set."
    },
    statusOk: (label) => `Classification “${label}” was set successfully.`,
    statusErrVerify: (label, found, why) =>
      `Couldn’t confirm classification “${label}”.` +
      (found ? ` Found: “${found}”.` : "") + (why ? `\nReason: ${why}` : ""),
    themeOptions: { light: "Light", dark: "Dark" },
    docLinkText: "More in documentation",
    labels: [
      {
        text: "TLP:Internal",
        tip: "For internal use only.",
        docUrl: "https://intra/docs/tlp-internal",
        helpHtml:
          "For use inside the organization. Don’t share externally. " +
          `<a href="https://intra/docs/tlp-internal" target="_blank" rel="noopener">More in documentation</a>.`
      },
      {
        text: "TLP:Protected",
        tip: "Contains sensitive info.",
        docUrl: "https://intra/docs/tlp-protected",
        helpHtml:
          "Sensitive data (e.g., personal, contractual, financial). Share on a need-to-know basis only. " +
          `<a href="https://intra/docs/tlp-protected" target="_blank" rel="noopener">More in documentation</a>.`
      },
      {
        text: "TLP:StrictlyProtected",
        tip: "Highly restricted data.",
        docUrl: "https://intra/docs/tlp-strict",
        helpHtml:
          "Highly restricted (secrets, credentials, regulated data). Use approved systems, encryption required. " +
          `<a href="https://intra/docs/tlp-strict" target="_blank" rel="noopener">More in documentation</a>.`
      }
    ],
  },

  CZ: {
    appTitle: "LabelMate pro Word",
    themeLabel: "Motiv",
    langLabel: "Jazyk",
    choosePrompt: "Zvolte úroveň klasifikace:",
    banner: {
      title: "Dokument nemá nastavenou klasifikaci",
      desc: "Zvolte úroveň klasifikace níže. Po nastavení banner zmizí."
    },
    statusOk: (label) => `Klasifikace „${label}” byla úspěšně nastavena.`,
    statusErrVerify: (label, found, why) =>
      `Nepodařilo se potvrdit nastavení klasifikace „${label}”.` +
      (found ? ` Nalezený text: „${found}”.` : "") + (why ? `\nDůvod: ${why}` : ""),
    themeOptions: { light: "Světlý", dark: "Tmavý" },
    docLinkText: "Více v dokumentaci",
    labels: [
      {
        text: "TLP:Interní",
        tip: "Interní informace.",
        docUrl: "https://intra/docs/tlp-internal-cs",
        helpHtml:
          "Určeno pro použití uvnitř organizace. Nešířit mimo firmu. " +
          `<a href="https://intra/docs/tlp-internal-cs" target="_blank" rel="noopener">Více v dokumentaci</a>.`
      },
      {
        text: "TLP:Chráněný",
        tip: "Citlivé údaje, omezené sdílení.",
        docUrl: "https://intra/docs/tlp-protected-cs",
        helpHtml:
          "Citlivá data (např. osobní, smluvní, finanční). Sdílení pouze pro nezbytně nutné osoby. " +
          `<a href="https://intra/docs/tlp-protected-cs" target="_blank" rel="noopener">Více v dokumentaci</a>.`
      },
      {
        text: "TLP:PřísněChráněný",
        tip: "Vysoce citlivá data.",
        docUrl: "https://intra/docs/tlp-strict-cs",
        helpHtml:
          "Vysoce citlivé (tajemství, přístupy, regulovaná data). Používejte schválené systémy, vyžadováno šifrování. " +
          `<a href="https://intra/docs/tlp-strict-cs" target="_blank" rel="noopener">Více v dokumentaci</a>.`
      }
    ],
  },

  SK: {
    appTitle: "LabelMate pre Word",
    themeLabel: "Motív",
    langLabel: "Jazyk",
    choosePrompt: "Zvoľte úroveň klasifikácie:",
    banner: {
      title: "Dokument nemá nastavenú klasifikáciu",
      desc: "Zvoľte klasifikáciu nižšie. Po nastavení banner zmizne."
    },
    statusOk: (label) => `Klasifikácia „${label}” bola úspešne nastavená.`,
    statusErrVerify: (label, found, why) =>
      `Nepodarilo sa potvrdiť klasifikáciu „${label}”.` +
      (found ? ` Nájdený text: „${found}”.` : "") + (why ? `\nDôvod: ${why}` : ""),
    themeOptions: { light: "Svetlý", dark: "Tmavý" },
    docLinkText: "Viac v dokumentácii",
    labels: [
      {
        text: "TLP:Interné",
        tip: "Interné informácie.",
        docUrl: "https://intra/docs/tlp-internal-sk",
        helpHtml:
          "Určené na použitie v rámci organizácie. Nezdieľať externe. " +
          `<a href="https://intra/docs/tlp-internal-sk" target="_blank" rel="noopener">Viac v dokumentácii</a>.`
      },
      {
        text: "TLP:Chránené",
        tip: "Citlivé údaje, obmedzené zdieľanie.",
        docUrl: "https://intra/docs/tlp-protected-sk",
        helpHtml:
          "Citlivé dáta (osobné, zmluvné, finančné). Zdieľať len pre nevyhnutných. " +
          `<a href="https://intra/docs/tlp-protected-sk" target="_blank" rel="noopener">Viac v dokumentácii</a>.`
      },
      {
        text: "TLP:PrísneChránené",
        tip: "Vysoko citlivé dáta.",
        docUrl: "https://intra/docs/tlp-strict-sk",
        helpHtml:
          "Vysoko citlivé (tajomstvá, prístupy, regulované). Používajte schválené systémy, vyžadované šifrovanie. " +
          `<a href="https://intra/docs/tlp-strict-sk" target="_blank" rel="noopener">Viac v dokumentácii</a>.`
      }
    ],
  }
};

window.LM_DEFAULT_LANG = "EN";
