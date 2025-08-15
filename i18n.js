// i18n.js — UI texty + definice klasifikací
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
    labels: [
      { text: "TLP:Internal",           tip: "For internal use only.", docUrl: "https://intra/docs/tlp-internal" },
      { text: "TLP:Protected",          tip: "Contains sensitive info.", docUrl: "https://intra/docs/tlp-protected" },
      { text: "TLP:StrictlyProtected",  tip: "Highly restricted data.", docUrl: "https://intra/docs/tlp-strict" }
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
    labels: [
      { text: "TLP:Interní",          tip: "Interní informace.",                docUrl: "https://intra/docs/tlp-internal-cs" },
      { text: "TLP:Chráněný",         tip: "Citlivé údaje, omezené sdílení.",   docUrl: "https://intra/docs/tlp-protected-cs" },
      { text: "TLP:PřísněChráněný",   tip: "Vysoce citlivá data.",              docUrl: "https://intra/docs/tlp-strict-cs" }
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
    labels: [
      { text: "TLP:Interné",          tip: "Interné informácie.",                 docUrl: "https://intra/docs/tlp-internal-sk" },
      { text: "TLP:Chránené",         tip: "Citlivé údaje, obmedzené zdieľanie.", docUrl: "https://intra/docs/tlp-protected-sk" },
      { text: "TLP:PrísneChránené",   tip: "Vysoko citlivé dáta.",                docUrl: "https://intra/docs/tlp-strict-sk" }
    ],
  }
};

window.LM_DEFAULT_LANG = "EN";
