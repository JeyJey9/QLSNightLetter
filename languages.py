LANG_CODES = {
    "American":            "en_US",
    "English":             "en_GB",
    "Spanish":             "es_ES",
    "Mexican Spanish":     "es_MX",
    "Brazilian Portuguese":"pt_BR",
    "German":              "de_DE",
    "Cologne German":      "de_KO",
    "Dutch":               "nl_NL",
    "Turkish":             "tr_TR",
    "Afrikaans":           "af_ZA",
    "Romanian":            "ro_RO",
    "Russian":             "ru_RU",
    "Simplified Chinese":  "zh_CN",
    "Traditional Chinese": "zh_TW",
    "Thai":                "th_TH",
    "Venezuelan Spanish":  "es_VE",
    "Argentinian Spanish": "es_AR",
    "Vietnamese":          "vi_VN",
    "Australian English":  "en_AU",
}

ALL_KEYS = [
    *LANG_CODES.keys(),
    "Language",
    "Select language",
    "OK",
    "Cancel",
    "Exit",
    "Apply",
    "Yes",
    "No",
    # GUI-specific
    "Root Folder:",
    "Mapping File:",
    "Browse",
    "Last Monthly  (F)",
    "Prev Monthly  (E)",
    "Last Weekly   (M)",
    "Prev Weekly   (L)",
    "Last Daily    (Q)",
    "Prev Daily    (P)",
    "Shift Values:",
    "Monthly (D→C–F→E)",
    "Weekly  (I→H–M→L)",
    "Daily   (P→O–Q→P)",
    "Run",
    "Done",
    "Complete",
    "skipped",
]

TRANSLATIONS = {
    # ---- ENGLISH (American) ----
    "en_US": {
        **{name: name for name in LANG_CODES},
        "Language": "Language",
        "Select language": "Select language",
        "OK": "OK",
        "Cancel": "Cancel",
        "Exit": "Exit",
        "Apply": "Apply",
        "Yes": "Yes",
        "No": "No",
        "Root Folder:": "Root Folder:",
        "Mapping File:": "Mapping File:",
        "Browse": "Browse",
        "Last Monthly  (F)": "Last Monthly  (F)",
        "Prev Monthly  (E)": "Prev Monthly  (E)",
        "Last Weekly   (M)": "Last Weekly   (M)",
        "Prev Weekly   (L)": "Prev Weekly   (L)",
        "Last Daily    (Q)": "Last Daily    (Q)",
        "Prev Daily    (P)": "Prev Daily    (P)",
        "Shift Values:": "Shift Values:",
        "Monthly (D→C–F→E)": "Monthly (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Weekly  (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Daily   (P→O–Q→P)",
        "Run": "Run",
        "Done": "Done",
        "Complete": "Complete",
        "skipped": "skipped",
        "Column\nUpdate\nOptions": "Column\nUpdate\nOptions",
        "Help" : "Help",
        "Help_Message":
    "QLS Night Letter Tool Help\n\n"
    "Root Folder:\n"
    "  Select the main folder containing your plant subfolders (e.g., CAL, WO CAL, FCPA).\n\n"
    "Mapping File:\n"
    "  Browse for the Sticker_Mapping.xlsx file to map PDF content to Excel columns.\n\n"
    "Column Update Options:\n"
    "  Select which metrics you want to update in the Excel master files:\n"
    "    - Last/Prev Monthly (F/E)\n"
    "    - Last/Prev Weekly (M/L)\n"
    "    - Last/Prev Daily (Q/P)\n\n"
    "Shift Values:\n"
    "  When checked, shifts old values before writing new ones in the selected columns.\n\n"
    "Text Log:\n"
    "  Displays progress, errors, and skipped files during processing.\n\n"
    "If you need further support, contact grece@ford.com.tr"

    },
    # ---- ENGLISH (British) ----
    "en_GB": {
        **{name: name for name in LANG_CODES},
        "Language": "Language",
        "Select language": "Select language",
        "OK": "OK",
        "Cancel": "Cancel",
        "Exit": "Exit",
        "Apply": "Apply",
        "Yes": "Yes",
        "No": "No",
        "Root Folder:": "Root Folder:",
        "Mapping File:": "Mapping File:",
        "Browse": "Browse",
        "Last Monthly  (F)": "Last Monthly  (F)",
        "Prev Monthly  (E)": "Prev Monthly  (E)",
        "Last Weekly   (M)": "Last Weekly   (M)",
        "Prev Weekly   (L)": "Prev Weekly   (L)",
        "Last Daily    (Q)": "Last Daily    (Q)",
        "Prev Daily    (P)": "Prev Daily    (P)",
        "Shift Values:": "Shift Values:",
        "Monthly (D→C–F→E)": "Monthly (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Weekly  (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Daily   (P→O–Q→P)",
        "Run": "Run",
        "Done": "Done",
        "Complete": "Complete",
        "skipped": "skipped",
        "Column\nUpdate\nOptions": "Column\nUpdate\nOptions",
        "Help" : "Help",
        "Help_Message":
    "QLS Night Letter Tool Help\n\n"
    "Root Folder:\n"
    "  Select the main folder containing your plant subfolders (e.g., CAL, WO CAL, FCPA).\n\n"
    "Mapping File:\n"
    "  Browse for the Sticker_Mapping.xlsx file to map PDF content to Excel columns.\n\n"
    "Column Update Options:\n"
    "  Select which metrics you want to update in the Excel master files:\n"
    "    - Last/Prev Monthly (F/E)\n"
    "    - Last/Prev Weekly (M/L)\n"
    "    - Last/Prev Daily (Q/P)\n\n"
    "Shift Values:\n"
    "  When checked, shifts old values before writing new ones in the selected columns.\n\n"
    "Text Log:\n"
    "  Displays progress, errors, and skipped files during processing.\n\n"
    "If you need further support, contact grece@ford.com.tr"

    },
    # ---- ENGLISH (Australian) ----
    "en_AU": {
        **{name: name for name in LANG_CODES},
        "Language": "Language",
        "Select language": "Select language",
        "OK": "OK",
        "Cancel": "Cancel",
        "Exit": "Exit",
        "Apply": "Apply",
        "Yes": "Yes",
        "No": "No",
        "Root Folder:": "Root Folder:",
        "Mapping File:": "Mapping File:",
        "Browse": "Browse",
        "Last Monthly  (F)": "Last Monthly  (F)",
        "Prev Monthly  (E)": "Prev Monthly  (E)",
        "Last Weekly   (M)": "Last Weekly   (M)",
        "Prev Weekly   (L)": "Prev Weekly   (L)",
        "Last Daily    (Q)": "Last Daily    (Q)",
        "Prev Daily    (P)": "Prev Daily    (P)",
        "Shift Values:": "Shift Values:",
        "Monthly (D→C–F→E)": "Monthly (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Weekly  (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Daily   (P→O–Q→P)",
        "Run": "Run",
        "Done": "Done",
        "Complete": "Complete",
        "skipped": "skipped",
        "Column\nUpdate\nOptions": "Column\nUpdate\nOptions",
        "Help" : "Help",
        "Help_Message":
    "QLS Night Letter Tool Help\n\n"
    "Root Folder:\n"
    "  Select the main folder containing your plant subfolders (e.g., CAL, WO CAL, FCPA).\n\n"
    "Mapping File:\n"
    "  Browse for the Sticker_Mapping.xlsx file to map PDF content to Excel columns.\n\n"
    "Column Update Options:\n"
    "  Select which metrics you want to update in the Excel master files:\n"
    "    - Last/Prev Monthly (F/E)\n"
    "    - Last/Prev Weekly (M/L)\n"
    "    - Last/Prev Daily (Q/P)\n\n"
    "Shift Values:\n"
    "  When checked, shifts old values before writing new ones in the selected columns.\n\n"
    "Text Log:\n"
    "  Displays progress, errors, and skipped files during processing.\n\n"
    "If you need further support, contact grece@ford.com.tr"

    },
    # ---- SPANISH (Spain) ----
    "es_ES": {
        "American": "Inglés americano",
        "English": "Inglés británico",
        "Spanish": "Español",
        "Mexican Spanish": "Español mexicano",
        "Brazilian Portuguese": "Portugués brasileño",
        "German": "Alemán",
        "Cologne German": "Alemán de Colonia",
        "Dutch": "Neerlandés",
        "Turkish": "Turco",
        "Afrikaans": "Afrikáans",
        "Romanian": "Rumano",
        "Russian": "Ruso",
        "Simplified Chinese": "Chino simplificado",
        "Traditional Chinese": "Chino tradicional",
        "Thai": "Tailandés",
        "Venezuelan Spanish": "Español venezolano",
        "Argentinian Spanish": "Español argentino",
        "Vietnamese": "Vietnamita",
        "Australian English": "Inglés australiano",
        "Language": "Idioma",
        "Select language": "Seleccionar idioma",
        "OK": "Aceptar",
        "Cancel": "Cancelar",
        "Exit": "Salir",
        "Apply": "Aplicar",
        "Yes": "Sí",
        "No": "No",
        "Root Folder:": "Carpeta raíz:",
        "Mapping File:": "Archivo de mapeo:",
        "Browse": "Examinar",
        "Last Monthly  (F)": "Último mensual (F)",
        "Prev Monthly  (E)": "Mensual anterior (E)",
        "Last Weekly   (M)": "Último semanal (M)",
        "Prev Weekly   (L)": "Semanal anterior (L)",
        "Last Daily    (Q)": "Último diario (Q)",
        "Prev Daily    (P)": "Diario anterior (P)",
        "Shift Values:": "Desplazar valores:",
        "Monthly (D→C–F→E)": "Mensual (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Semanal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Diario (P→O–Q→P)",
        "Run": "Ejecutar",
        "Done": "Hecho",
        "Complete": "Completo",
        "skipped": "omitido",
        "Column\nUpdate\nOptions": "Columna\nActualizar\nOpciones",
        "Help": "Ayuda",
        "Help_Message":
    "Ayuda QLS Night Letter Tool\n\n"
    "Carpeta raíz:\n"
    "  Seleccione la carpeta principal que contiene las subcarpetas de planta (ej: CAL, WO CAL, FCPA).\n\n"
    "Archivo de mapeo:\n"
    "  Seleccione Sticker_Mapping.xlsx para mapear el contenido del PDF a columnas de Excel.\n\n"
    "Opciones de actualización de columna:\n"
    "  Seleccione qué métricas desea actualizar en los archivos maestros de Excel:\n"
    "    - Mensual último/anterior (F/E)\n"
    "    - Semanal último/anterior (M/L)\n"
    "    - Diario último/anterior (Q/P)\n\n"
    "Desplazar valores:\n"
    "  Si está marcado, desplaza los valores antiguos antes de escribir los nuevos en las columnas seleccionadas.\n\n"
    "Registro de texto:\n"
    "  Muestra el progreso, errores y archivos omitidos durante el procesamiento.\n\n"
    "Si necesita más ayuda, contacte a grece@ford.com.tr"

    },
    # ---- MEXICAN SPANISH ----
    "es_MX": {
        "American": "Inglés americano",
        "English": "Inglés británico",
        "Spanish": "Español",
        "Mexican Spanish": "Español mexicano",
        "Brazilian Portuguese": "Portugués brasileño",
        "German": "Alemán",
        "Cologne German": "Alemán de Colonia",
        "Dutch": "Neerlandés",
        "Turkish": "Turco",
        "Afrikaans": "Afrikáans",
        "Romanian": "Rumano",
        "Russian": "Ruso",
        "Simplified Chinese": "Chino simplificado",
        "Traditional Chinese": "Chino tradicional",
        "Thai": "Tailandés",
        "Venezuelan Spanish": "Español venezolano",
        "Argentinian Spanish": "Español argentino",
        "Vietnamese": "Vietnamita",
        "Australian English": "Inglés australiano",
        "Language": "Idioma",
        "Select language": "Seleccionar idioma",
        "OK": "Aceptar",
        "Cancel": "Cancelar",
        "Exit": "Salir",
        "Apply": "Aplicar",
        "Yes": "Sí",
        "No": "No",
        "Root Folder:": "Carpeta raíz:",
        "Mapping File:": "Archivo de mapeo:",
        "Browse": "Examinar",
        "Last Monthly  (F)": "Último mensual (F)",
        "Prev Monthly  (E)": "Mensual anterior (E)",
        "Last Weekly   (M)": "Último semanal (M)",
        "Prev Weekly   (L)": "Semanal anterior (L)",
        "Last Daily    (Q)": "Último diario (Q)",
        "Prev Daily    (P)": "Diario anterior (P)",
        "Shift Values:": "Desplazar valores:",
        "Monthly (D→C–F→E)": "Mensual (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Semanal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Diario (P→O–Q→P)",
        "Run": "Ejecutar",
        "Done": "Hecho",
        "Complete": "Completo",
        "skipped": "omitido",
        "Column\nUpdate\nOptions": "Columna\nActualizar\nOpciones",
        "Help": "Ayuda",
        "Help_Message":
        "Ayuda QLS Night Letter Tool\n\n"
        "Carpeta raíz:\n"
        "  Seleccione la carpeta principal que contiene las subcarpetas de planta (ej: CAL, WO CAL, FCPA).\n\n"
        "Archivo de mapeo:\n"
        "  Seleccione Sticker_Mapping.xlsx para mapear el contenido del PDF a columnas de Excel.\n\n"
        "Opciones de actualización de columna:\n"
        "  Seleccione qué métricas desea actualizar en los archivos maestros de Excel:\n"
        "    - Mensual último/anterior (F/E)\n"
        "    - Semanal último/anterior (M/L)\n"
        "    - Diario último/anterior (Q/P)\n\n"
        "Desplazar valores:\n"
        "  Si está marcado, desplaza los valores antiguos antes de escribir los nuevos en las columnas seleccionadas.\n\n"
        "Registro de texto:\n"
        "  Muestra el progreso, errores y archivos omitidos durante el procesamiento.\n\n"
        "Si necesita más ayuda, contacte a grece@ford.com.tr"

    },
    # ---- BRAZILIAN PORTUGUESE ----
    "pt_BR": {
        "American": "Inglês americano",
        "English": "Inglês britânico",
        "Spanish": "Espanhol",
        "Mexican Spanish": "Espanhol mexicano",
        "Brazilian Portuguese": "Português brasileiro",
        "German": "Alemão",
        "Cologne German": "Alemão de Colônia",
        "Dutch": "Holandês",
        "Turkish": "Turco",
        "Afrikaans": "Africâner",
        "Romanian": "Romeno",
        "Russian": "Russo",
        "Simplified Chinese": "Chinês simplificado",
        "Traditional Chinese": "Chinês tradicional",
        "Thai": "Tailandês",
        "Venezuelan Spanish": "Espanhol venezuelano",
        "Argentinian Spanish": "Espanhol argentino",
        "Vietnamese": "Vietnamita",
        "Australian English": "Inglês australiano",
        "Language": "Idioma",
        "Select language": "Selecionar idioma",
        "OK": "OK",
        "Cancel": "Cancelar",
        "Exit": "Sair",
        "Apply": "Aplicar",
        "Yes": "Sim",
        "No": "Não",
        "Root Folder:": "Pasta raiz:",
        "Mapping File:": "Arquivo de mapeamento:",
        "Browse": "Procurar",
        "Last Monthly  (F)": "Último mensal (F)",
        "Prev Monthly  (E)": "Mensal anterior (E)",
        "Last Weekly   (M)": "Último semanal (M)",
        "Prev Weekly   (L)": "Semanal anterior (L)",
        "Last Daily    (Q)": "Último diário (Q)",
        "Prev Daily    (P)": "Diário anterior (P)",
        "Shift Values:": "Deslocar valores:",
        "Monthly (D→C–F→E)": "Mensal (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Semanal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Diário (P→O–Q→P)",
        "Run": "Executar",
        "Done": "Concluído",
        "Complete": "Completo",
        "skipped": "ignorado",
        "Column\nUpdate\nOptions": "Coluna\nAtualizar\nOpções",
        "Help": "Ajuda",
        "Help_Message":
    "Ajuda da Ferramenta QLS Night Letter\n\n"
    "Pasta raiz:\n"
    "  Selecione a pasta principal que contém as subpastas das plantas (ex: CAL, WO CAL, FCPA).\n\n"
    "Arquivo de mapeamento:\n"
    "  Selecione o arquivo Sticker_Mapping.xlsx para mapear o conteúdo do PDF para colunas do Excel.\n\n"
    "Opções de atualização de coluna:\n"
    "  Selecione quais métricas deseja atualizar nos arquivos mestres do Excel:\n"
    "    - Último/Anterior Mensal (F/E)\n"
    "    - Último/Anterior Semanal (M/L)\n"
    "    - Último/Anterior Diário (Q/P)\n\n"
    "Deslocar valores:\n"
    "  Quando marcado, desloca os valores antigos antes de gravar os novos nas colunas selecionadas.\n\n"
    "Log de texto:\n"
    "  Exibe o progresso, erros e arquivos ignorados durante o processamento.\n\n"
    "Se precisar de mais suporte, entre em contato: grece@ford.com.tr"

    },
    # ---- GERMAN ----
    "de_DE": {
        "American": "Amerikanisches Englisch",
        "English": "Englisch",
        "Spanish": "Spanisch",
        "Mexican Spanish": "Mexikanisches Spanisch",
        "Brazilian Portuguese": "Brasilianisches Portugiesisch",
        "German": "Deutsch",
        "Cologne German": "Kölsch",
        "Dutch": "Niederländisch",
        "Turkish": "Türkisch",
        "Afrikaans": "Afrikaans",
        "Romanian": "Rumänisch",
        "Russian": "Russisch",
        "Simplified Chinese": "Vereinfachtes Chinesisch",
        "Traditional Chinese": "Traditionelles Chinesisch",
        "Thai": "Thailändisch",
        "Venezuelan Spanish": "Venezolanisches Spanisch",
        "Argentinian Spanish": "Argentinisches Spanisch",
        "Vietnamese": "Vietnamesisch",
        "Australian English": "Australisches Englisch",
        "Language": "Sprache",
        "Select language": "Sprache auswählen",
        "OK": "OK",
        "Cancel": "Abbrechen",
        "Exit": "Beenden",
        "Apply": "Übernehmen",
        "Yes": "Ja",
        "No": "Nein",
        "Root Folder:": "Stammordner:",
        "Mapping File:": "Mapping-Datei:",
        "Browse": "Durchsuchen",
        "Last Monthly  (F)": "Letzter Monat (F)",
        "Prev Monthly  (E)": "Vorheriger Monat (E)",
        "Last Weekly   (M)": "Letzte Woche (M)",
        "Prev Weekly   (L)": "Vorherige Woche (L)",
        "Last Daily    (Q)": "Letzter Tag (Q)",
        "Prev Daily    (P)": "Vorheriger Tag (P)",
        "Shift Values:": "Werte verschieben:",
        "Monthly (D→C–F→E)": "Monatlich (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Wöchentlich (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Täglich (P→O–Q→P)",
        "Run": "Ausführen",
        "Done": "Fertig",
        "Complete": "Vollständig",
        "skipped": "übersprungen",
        "Column\nUpdate\nOptions": "Spalte\nAktualisieren\nOptionen",
        "Help": "Hilfe",
        "Help_Message":
    "QLS Night Letter Tool Hilfe\n\n"
    "Stammordner:\n"
    "  Wählen Sie den Hauptordner mit den Werk-Unterordnern (z. B. CAL, WO CAL, FCPA).\n\n"
    "Mapping-Datei:\n"
    "  Wählen Sie die Sticker_Mapping.xlsx-Datei, um PDF-Inhalte auf Excel-Spalten zuzuordnen.\n\n"
    "Spalte Aktualisieren Optionen:\n"
    "  Wählen Sie, welche Kennzahlen Sie in den Excel-Masterdateien aktualisieren möchten:\n"
    "    - Letzter/Vorheriger Monat (F/E)\n"
    "    - Letzte/Vorherige Woche (M/L)\n"
    "    - Letzter/Vorheriger Tag (Q/P)\n\n"
    "Werte verschieben:\n"
    "  Wenn aktiviert, werden alte Werte verschoben, bevor neue in die ausgewählten Spalten geschrieben werden.\n\n"
    "Textprotokoll:\n"
    "  Zeigt Fortschritt, Fehler und übersprungene Dateien während der Verarbeitung an.\n\n"
    "Bei weiteren Fragen kontaktieren Sie grece@ford.com.tr"

    },
    # ---- COLOGNE GERMAN ----
    "de_KO": {
        "American": "Amerikanisches Englisch",
        "English": "Englisch",
        "Spanish": "Spanisch",
        "Mexican Spanish": "Mexikanisches Spanisch",
        "Brazilian Portuguese": "Brasilianisches Portugiesisch",
        "German": "Deutsch",
        "Cologne German": "Kölsch",
        "Dutch": "Niederländisch",
        "Turkish": "Türkisch",
        "Afrikaans": "Afrikaans",
        "Romanian": "Rumänisch",
        "Russian": "Russisch",
        "Simplified Chinese": "Vereinfachtes Chinesisch",
        "Traditional Chinese": "Traditionelles Chinesisch",
        "Thai": "Thailändisch",
        "Venezuelan Spanish": "Venezolanisches Spanisch",
        "Argentinian Spanish": "Argentinisches Spanisch",
        "Vietnamese": "Vietnamesisch",
        "Australian English": "Australisches Englisch",
        "Language": "Sprache",
        "Select language": "Sprache auswählen",
        "OK": "OK",
        "Cancel": "Abbrechen",
        "Exit": "Beenden",
        "Apply": "Übernehmen",
        "Yes": "Ja",
        "No": "Nein",
        "Root Folder:": "Stammordner:",
        "Mapping File:": "Mapping-Datei:",
        "Browse": "Durchsuchen",
        "Last Monthly  (F)": "Letzter Monat (F)",
        "Prev Monthly  (E)": "Vorheriger Monat (E)",
        "Last Weekly   (M)": "Letzte Woche (M)",
        "Prev Weekly   (L)": "Vorherige Woche (L)",
        "Last Daily    (Q)": "Letzter Tag (Q)",
        "Prev Daily    (P)": "Vorheriger Tag (P)",
        "Shift Values:": "Werte verschieben:",
        "Monthly (D→C–F→E)": "Monatlich (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Wöchentlich (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Täglich (P→O–Q→P)",
        "Run": "Ausführen",
        "Done": "Fertig",
        "Complete": "Vollständig",
        "skipped": "übersprungen",
        "Column\nUpdate\nOptions": "Spalte\nAktualisieren\nOptionen",
        "Help": "Hilfe",
        "Help_Message":
    "QLS Night Letter Tool Hilfe\n\n"
    "Stammordner:\n"
    "  Wählen Sie den Hauptordner mit den Werk-Unterordnern (z. B. CAL, WO CAL, FCPA).\n\n"
    "Mapping-Datei:\n"
    "  Wählen Sie die Sticker_Mapping.xlsx-Datei, um PDF-Inhalte auf Excel-Spalten zuzuordnen.\n\n"
    "Spalte Aktualisieren Optionen:\n"
    "  Wählen Sie, welche Kennzahlen Sie in den Excel-Masterdateien aktualisieren möchten:\n"
    "    - Letzter/Vorheriger Monat (F/E)\n"
    "    - Letzte/Vorherige Woche (M/L)\n"
    "    - Letzter/Vorheriger Tag (Q/P)\n\n"
    "Werte verschieben:\n"
    "  Wenn aktiviert, werden alte Werte verschoben, bevor neue in die ausgewählten Spalten geschrieben werden.\n\n"
    "Textprotokoll:\n"
    "  Zeigt Fortschritt, Fehler und übersprungene Dateien während der Verarbeitung an.\n\n"
    "Bei weiteren Fragen kontaktieren Sie grece@ford.com.tr"

    },
    # ---- DUTCH ----
    "nl_NL": {
        "American": "Amerikaans Engels",
        "English": "Brits Engels",
        "Spanish": "Spaans",
        "Mexican Spanish": "Mexicaans Spaans",
        "Brazilian Portuguese": "Braziliaans Portugees",
        "German": "Duits",
        "Cologne German": "Keuls",
        "Dutch": "Nederlands",
        "Turkish": "Turks",
        "Afrikaans": "Afrikaans",
        "Romanian": "Roemeens",
        "Russian": "Russisch",
        "Simplified Chinese": "Vereenvoudigd Chinees",
        "Traditional Chinese": "Traditioneel Chinees",
        "Thai": "Thais",
        "Venezuelan Spanish": "Venezolaans Spaans",
        "Argentinian Spanish": "Argentijns Spaans",
        "Vietnamese": "Vietnamees",
        "Australian English": "Australisch Engels",
        "Language": "Taal",
        "Select language": "Taal selecteren",
        "OK": "OK",
        "Cancel": "Annuleren",
        "Exit": "Afsluiten",
        "Apply": "Toepassen",
        "Yes": "Ja",
        "No": "Nee",
        "Root Folder:": "Hoofdmap:",
        "Mapping File:": "Koppelingsbestand:",
        "Browse": "Bladeren",
        "Last Monthly  (F)": "Laatste maand (F)",
        "Prev Monthly  (E)": "Vorige maand (E)",
        "Last Weekly   (M)": "Laatste week (M)",
        "Prev Weekly   (L)": "Vorige week (L)",
        "Last Daily    (Q)": "Laatste dag (Q)",
        "Prev Daily    (P)": "Vorige dag (P)",
        "Shift Values:": "Waarden verschuiven:",
        "Monthly (D→C–F→E)": "Maandelijks (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Wekelijks (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Dagelijks (P→O–Q→P)",
        "Run": "Uitvoeren",
        "Done": "Klaar",
        "Complete": "Voltooid",
        "skipped": "overgeslagen",
        "Column\nUpdate\nOptions": "Kolom\nBijwerken\nOpties",
        "Help": "Help",
        "Help_Message":
    "QLS Night Letter Tool Help\n\n"
    "Hoofdmap:\n"
    "  Selecteer de hoofdmap met de submappen per fabriek (bijv. CAL, WO CAL, FCPA).\n\n"
    "Koppelingsbestand:\n"
    "  Kies het bestand Sticker_Mapping.xlsx om PDF-inhoud naar Excel-kolommen te koppelen.\n\n"
    "Kolom update opties:\n"
    "  Selecteer welke metingen je in de Excel-hoofdbestanden wilt bijwerken:\n"
    "    - Laatste/vorige maand (F/E)\n"
    "    - Laatste/vorige week (M/L)\n"
    "    - Laatste/vorige dag (Q/P)\n\n"
    "Waarden verschuiven:\n"
    "  Indien aangevinkt worden oude waarden verschoven voordat nieuwe worden ingevuld.\n\n"
    "Tekstlogboek:\n"
    "  Toont voortgang, fouten en overgeslagen bestanden tijdens het verwerken.\n\n"
    "Voor verdere hulp: grece@ford.com.tr"

    },
    # ---- TURKISH ----
    "tr_TR": {
        "American": "Amerikan İngilizcesi",
        "English": "İngilizce",
        "Spanish": "İspanyolca",
        "Mexican Spanish": "Meksika İspanyolcası",
        "Brazilian Portuguese": "Brezilya Portekizcesi",
        "German": "Almanca",
        "Cologne German": "Köln Almancası",
        "Dutch": "Felemenkçe",
        "Turkish": "Türkçe",
        "Afrikaans": "Afrikanca",
        "Romanian": "Rumence",
        "Russian": "Rusça",
        "Simplified Chinese": "Basitleştirilmiş Çince",
        "Traditional Chinese": "Geleneksel Çince",
        "Thai": "Tayca",
        "Venezuelan Spanish": "Venezuela İspanyolcası",
        "Argentinian Spanish": "Arjantin İspanyolcası",
        "Vietnamese": "Vietnamca",
        "Australian English": "Avustralya İngilizcesi",
        "Language": "Dil",
        "Select language": "Dil seçin",
        "OK": "Tamam",
        "Cancel": "İptal",
        "Exit": "Çıkış",
        "Apply": "Uygula",
        "Yes": "Evet",
        "No": "Hayır",
        "Root Folder:": "Kök Dizin:",
        "Mapping File:": "Eşleme Dosyası:",
        "Browse": "Gözat",
        "Last Monthly  (F)": "Son Aylık (F)",
        "Prev Monthly  (E)": "Önceki Aylık (E)",
        "Last Weekly   (M)": "Son Haftalık (M)",
        "Prev Weekly   (L)": "Önceki Haftalık (L)",
        "Last Daily    (Q)": "Son Günlük (Q)",
        "Prev Daily    (P)": "Önceki Günlük (P)",
        "Shift Values:": "Değerleri Kaydır:",
        "Monthly (D→C–F→E)": "Aylık (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Haftalık (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Günlük (P→O–Q→P)",
        "Run": "Çalıştır",
        "Done": "Bitti",
        "Complete": "Tamamlandı",
        "skipped": "atlandı",
        "Column\nUpdate\nOptions": "Sütun\nGüncelle\nSeçenekler",
        "Help": "Yardım",
        "Help_Message":
    "QLS Night Letter Tool Yardım\n\n"
    "Kök Klasör:\n"
    "  Ana klasörü seçin, alt klasörler fabrika klasörleri olmalı (örn: CAL, WO CAL, FCPA).\n\n"
    "Eşleme Dosyası:\n"
    "  PDF içeriğini Excel sütunlarına eşlemek için Sticker_Mapping.xlsx dosyasını seçin.\n\n"
    "Sütun Güncelleme Seçenekleri:\n"
    "  Excel ana dosyalarında hangi metriklerin güncelleneceğini seçin:\n"
    "    - Son/Önceki Aylık (F/E)\n"
    "    - Son/Önceki Haftalık (M/L)\n"
    "    - Son/Önceki Günlük (Q/P)\n\n"
    "Değerleri Kaydır:\n"
    "  İşaretliyse, yeni değerler yazılmadan önce eski değerler kaydırılır.\n\n"
    "Kayıt günlüğü:\n"
    "  İşlem sırasında ilerlemeyi, hataları ve atlanan dosyaları gösterir.\n\n"
    "Daha fazla destek için: grece@ford.com.tr"

    },
    # ---- AFRIKAANS ----
    "af_ZA": {
        "American": "Amerikaanse Engels",
        "English": "Britse Engels",
        "Spanish": "Spaans",
        "Mexican Spanish": "Meksikaanse Spaans",
        "Brazilian Portuguese": "Brasiliaanse Portugees",
        "German": "Duits",
        "Cologne German": "Keuls",
        "Dutch": "Nederlands",
        "Turkish": "Turks",
        "Afrikaans": "Afrikaans",
        "Romanian": "Roemeens",
        "Russian": "Russies",
        "Simplified Chinese": "Vereenvoudigde Sjinees",
        "Traditional Chinese": "Tradisionele Sjinees",
        "Thai": "Thai",
        "Venezuelan Spanish": "Venezolaanse Spaans",
        "Argentinian Spanish": "Argentynse Spaans",
        "Vietnamese": "Viëtnamees",
        "Australian English": "Australiese Engels",
        "Language": "Taal",
        "Select language": "Kies taal",
        "OK": "OK",
        "Cancel": "Kanselleer",
        "Exit": "Verlaat",
        "Apply": "Pas toe",
        "Yes": "Ja",
        "No": "Nee",
        "Root Folder:": "Gids:",
        "Mapping File:": "Koppel-lêer:",
        "Browse": "Blaai",
        "Last Monthly  (F)": "Laaste maandeliks (F)",
        "Prev Monthly  (E)": "Vorige maandeliks (E)",
        "Last Weekly   (M)": "Laaste weekliks (M)",
        "Prev Weekly   (L)": "Vorige weekliks (L)",
        "Last Daily    (Q)": "Laaste daagliks (Q)",
        "Prev Daily    (P)": "Vorige daagliks (P)",
        "Shift Values:": "Verskuif waardes:",
        "Monthly (D→C–F→E)": "Maandeliks (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Weekliks (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Daagliks (P→O–Q→P)",
        "Run": "Begin",
        "Done": "Klaar",
        "Complete": "Voltooi",
        "skipped": "oorgegaan",
        "Column\nUpdate\nOptions": "Kolom\nOpdateer\nOpsies",
        "Help": "Hulp",
        "Help_Message":
    "QLS Night Letter Tool Hulp\n\n"
    "Gids:\n"
    "  Kies die hoofgids wat die fabriek-subgidse bevat (bv. CAL, WO CAL, FCPA).\n\n"
    "Koppel-lêer:\n"
    "  Kies Sticker_Mapping.xlsx om PDF-inhoud na Excel-kolomme te koppel.\n\n"
    "Kolom-opdateringsopsies:\n"
    "  Kies watter waardes jy in die Excel-meesterlêers wil opdateer:\n"
    "    - Laaste/Vorige maand (F/E)\n"
    "    - Laaste/Vorige week (M/L)\n"
    "    - Laaste/Vorige dag (Q/P)\n\n"
    "Verskuif waardes:\n"
    "  Indien gekies, word ou waardes verskuif voordat nuwe waardes in die kolomme geskryf word.\n\n"
    "Tekslogboek:\n"
    "  Wys vordering, foute en oorgeslaande lêers tydens verwerking.\n\n"
    "Vir verdere hulp, kontak grece@ford.com.tr"

    },
    # ---- ROMANIAN ----
    "ro_RO": {
        "American": "Americană",
        "English": "Engleză",
        "Spanish": "Spaniolă",
        "Mexican Spanish": "Spaniolă mexicană",
        "Brazilian Portuguese": "Portugheză braziliană",
        "German": "Germană",
        "Cologne German": "Germană din Köln",
        "Dutch": "Olandeză",
        "Turkish": "Turcă",
        "Afrikaans": "Afrikaans",
        "Romanian": "Română",
        "Russian": "Rusă",
        "Simplified Chinese": "Chineză simplificată",
        "Traditional Chinese": "Chineză tradițională",
        "Thai": "Thailandeză",
        "Venezuelan Spanish": "Spaniolă venezueleană",
        "Argentinian Spanish": "Spaniolă argentiniană",
        "Vietnamese": "Vietnameză",
        "Australian English": "Engleză australiană",
        "Language": "Limbă",
        "Select language": "Selectează limba",
        "OK": "OK",
        "Cancel": "Anulează",
        "Exit": "Ieșire",
        "Apply": "Aplică",
        "Yes": "Da",
        "No": "Nu",
        "Root Folder:": "Director rădăcină:",
        "Mapping File:": "Fișier de mapare:",
        "Browse": "Răsfoiește",
        "Last Monthly  (F)": "Ultima lunară (F)",
        "Prev Monthly  (E)": "Lunară precedentă (E)",
        "Last Weekly   (M)": "Ultima săptămânală (M)",
        "Prev Weekly   (L)": "Săptămânală precedentă (L)",
        "Last Daily    (Q)": "Ultima zilnică (Q)",
        "Prev Daily    (P)": "Zilnică precedentă (P)",
        "Shift Values:": "Valori de mutat:",
        "Monthly (D→C–F→E)": "Lunar (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Săptămânal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Zilnic (P→O–Q→P)",
        "Run": "Rulează",
        "Done": "Gata",
        "Complete": "Complet",
        "skipped": "sărite",
        "Column\nUpdate\nOptions": "Opțiuni\nActualizare\nColoană",
        "Help": "Ajutor",
        "Help_Message":
    "Ghid QLS Night Letter Tool\n\n"
    "Director rădăcină:\n"
    "  Selectați dosarul principal care conține subfolderele fabricii (ex: CAL, WO CAL, FCPA).\n\n"
    "Fișier de mapare:\n"
    "  Selectați Sticker_Mapping.xlsx pentru a mapa conținutul PDF către coloane Excel.\n\n"
    "Opțiuni actualizare coloană:\n"
    "  Selectați ce metrici doriți să actualizați în fișierele master Excel:\n"
    "    - Lunar Ultim/Precedent (F/E)\n"
    "    - Săptămânal Ultim/Precedent (M/L)\n"
    "    - Zilnic Ultim/Precedent (Q/P)\n\n"
    "Valori de mutat:\n"
    "  Dacă este bifat, valorile vechi sunt mutate înainte de a scrie valorile noi în coloanele selectate.\n\n"
    "Jurnal text:\n"
    "  Afișează progresul, erorile și fișierele sărite în timpul procesării.\n\n"
    "Pentru suport suplimentar, contactați grece@ford.com.tr"

    },
    # ---- RUSSIAN ----
    "ru_RU": {
        "American": "Американский английский",
        "English": "Британский английский",
        "Spanish": "Испанский",
        "Mexican Spanish": "Мексиканский испанский",
        "Brazilian Portuguese": "Бразильский португальский",
        "German": "Немецкий",
        "Cologne German": "Кёльнский диалект",
        "Dutch": "Голландский",
        "Turkish": "Турецкий",
        "Afrikaans": "Африкаанс",
        "Romanian": "Румынский",
        "Russian": "Русский",
        "Simplified Chinese": "Упрощённый китайский",
        "Traditional Chinese": "Традиционный китайский",
        "Thai": "Тайский",
        "Venezuelan Spanish": "Венесуэльский испанский",
        "Argentinian Spanish": "Аргентинский испанский",
        "Vietnamese": "Вьетнамский",
        "Australian English": "Австралийский английский",
        "Language": "Язык",
        "Select language": "Выберите язык",
        "OK": "ОК",
        "Cancel": "Отмена",
        "Exit": "Выход",
        "Apply": "Применить",
        "Yes": "Да",
        "No": "Нет",
        "Root Folder:": "Корневая папка:",
        "Mapping File:": "Файл сопоставления:",
        "Browse": "Обзор",
        "Last Monthly  (F)": "Последний месяц (F)",
        "Prev Monthly  (E)": "Предыдущий месяц (E)",
        "Last Weekly   (M)": "Последняя неделя (M)",
        "Prev Weekly   (L)": "Предыдущая неделя (L)",
        "Last Daily    (Q)": "Последний день (Q)",
        "Prev Daily    (P)": "Предыдущий день (P)",
        "Shift Values:": "Сдвиг значений:",
        "Monthly (D→C–F→E)": "Ежемесячно (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Еженедельно (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Ежедневно (P→O–Q→P)",
        "Run": "Запуск",
        "Done": "Готово",
        "Complete": "Выполнено",
        "skipped": "пропущено",
        "Column\nUpdate\nOptions": "Колонка\nОбновить\nОпции",
        "Help": "Помощь",
        "Help_Message":
    "Справка QLS Night Letter Tool\n\n"
    "Корневая папка:\n"
    "  Выберите основную папку, содержащую подпапки заводов (например: CAL, WO CAL, FCPA).\n\n"
    "Файл сопоставления:\n"
    "  Выберите файл Sticker_Mapping.xlsx для сопоставления содержимого PDF со столбцами Excel.\n\n"
    "Опции обновления столбцов:\n"
    "  Выберите, какие метрики обновлять в мастер-файлах Excel:\n"
    "    - Последний/Предыдущий месяц (F/E)\n"
    "    - Последняя/Предыдущая неделя (M/L)\n"
    "    - Последний/Предыдущий день (Q/P)\n\n"
    "Сдвиг значений:\n"
    "  Если отмечено, старые значения смещаются перед записью новых в выбранные столбцы.\n\n"
    "Журнал:\n"
    "  Показывает ход выполнения, ошибки и пропущенные файлы.\n\n"
    "Если нужна помощь, пишите на grece@ford.com.tr"

    },
    # ---- SIMPLIFIED CHINESE ----
    "zh_CN": {
        "American": "美式英语",
        "English": "英式英语",
        "Spanish": "西班牙语",
        "Mexican Spanish": "墨西哥西班牙语",
        "Brazilian Portuguese": "巴西葡萄牙语",
        "German": "德语",
        "Cologne German": "科隆德语",
        "Dutch": "荷兰语",
        "Turkish": "土耳其语",
        "Afrikaans": "南非荷兰语",
        "Romanian": "罗马尼亚语",
        "Russian": "俄语",
        "Simplified Chinese": "简体中文",
        "Traditional Chinese": "繁体中文",
        "Thai": "泰语",
        "Venezuelan Spanish": "委内瑞拉西班牙语",
        "Argentinian Spanish": "阿根廷西班牙语",
        "Vietnamese": "越南语",
        "Australian English": "澳大利亚英语",
        "Language": "语言",
        "Select language": "选择语言",
        "OK": "确定",
        "Cancel": "取消",
        "Exit": "退出",
        "Apply": "应用",
        "Yes": "是",
        "No": "否",
        "Root Folder:": "根文件夹：",
        "Mapping File:": "映射文件：",
        "Browse": "浏览",
        "Last Monthly  (F)": "最近月度 (F)",
        "Prev Monthly  (E)": "上次月度 (E)",
        "Last Weekly   (M)": "最近周度 (M)",
        "Prev Weekly   (L)": "上次周度 (L)",
        "Last Daily    (Q)": "最近每日 (Q)",
        "Prev Daily    (P)": "上次每日 (P)",
        "Shift Values:": "移动值：",
        "Monthly (D→C–F→E)": "月度 (D→C–F→E)",
        "Weekly  (I→H–M→L)": "周度 (I→H–M→L)",
        "Daily   (P→O–Q→P)": "每日 (P→O–Q→P)",
        "Run": "运行",
        "Done": "完成",
        "Complete": "完成",
        "skipped": "已跳过",
        "Column\nUpdate\nOptions": "列\n更新\n选项",
        "Help": "帮助",
        "Help_Message":
    "QLS Night Letter Tool 帮助\n\n"
    "根文件夹：\n"
    "  选择包含工厂子文件夹的主文件夹（如：CAL、WO CAL、FCPA）。\n\n"
    "映射文件：\n"
    "  选择 Sticker_Mapping.xlsx 文件，将 PDF 内容映射到 Excel 列。\n\n"
    "列更新选项：\n"
    "  选择要在 Excel 主文件中更新的指标：\n"
    "    - 最近/上次月度 (F/E)\n"
    "    - 最近/上次周度 (M/L)\n"
    "    - 最近/上次每日 (Q/P)\n\n"
    "移动值：\n"
    "  若勾选，会在写入新值前将旧值下移。\n\n"
    "日志：\n"
    "  显示进度、错误以及处理时被跳过的文件。\n\n"
    "如需更多支持，请联系 grece@ford.com.tr"

    },
    # ---- TRADITIONAL CHINESE ----
    "zh_TW": {
        "American": "美式英語",
        "English": "英式英語",
        "Spanish": "西班牙語",
        "Mexican Spanish": "墨西哥西班牙語",
        "Brazilian Portuguese": "巴西葡萄牙語",
        "German": "德語",
        "Cologne German": "科隆德語",
        "Dutch": "荷蘭語",
        "Turkish": "土耳其語",
        "Afrikaans": "南非荷蘭語",
        "Romanian": "羅馬尼亞語",
        "Russian": "俄語",
        "Simplified Chinese": "簡體中文",
        "Traditional Chinese": "繁體中文",
        "Thai": "泰語",
        "Venezuelan Spanish": "委內瑞拉西班牙語",
        "Argentinian Spanish": "阿根廷西班牙語",
        "Vietnamese": "越南語",
        "Australian English": "澳大利亞英語",
        "Language": "語言",
        "Select language": "選擇語言",
        "OK": "確定",
        "Cancel": "取消",
        "Exit": "退出",
        "Apply": "應用",
        "Yes": "是",
        "No": "否",
        "Root Folder:": "根目錄：",
        "Mapping File:": "映射文件：",
        "Browse": "瀏覽",
        "Last Monthly  (F)": "最近每月 (F)",
        "Prev Monthly  (E)": "上次每月 (E)",
        "Last Weekly   (M)": "最近每週 (M)",
        "Prev Weekly   (L)": "上次每週 (L)",
        "Last Daily    (Q)": "最近每日 (Q)",
        "Prev Daily    (P)": "上次每日 (P)",
        "Shift Values:": "移動值：",
        "Monthly (D→C–F→E)": "每月 (D→C–F→E)",
        "Weekly  (I→H–M→L)": "每週 (I→H–M→L)",
        "Daily   (P→O–Q→P)": "每日 (P→O–Q→P)",
        "Run": "執行",
        "Done": "完成",
        "Complete": "完成",
        "skipped": "已跳過",
        "Column\nUpdate\nOptions": "欄\n更新\n選項",
        "Help": "幫助",
        "Help_Message":
    "QLS Night Letter Tool 使用說明\n\n"
    "根目錄：\n"
    "  選擇包含各廠區子資料夾的主要資料夾（例如：CAL、WO CAL、FCPA）。\n\n"
    "映射文件：\n"
    "  選擇 Sticker_Mapping.xlsx 檔案，以將 PDF 內容對應到 Excel 欄位。\n\n"
    "欄更新選項：\n"
    "  選擇要在 Excel 主檔中更新的指標：\n"
    "    - 最近/前次每月 (F/E)\n"
    "    - 最近/前次每週 (M/L)\n"
    "    - 最近/前次每日 (Q/P)\n\n"
    "移動值：\n"
    "  若勾選，會先將舊值下移再寫入新值。\n\n"
    "文字日誌：\n"
    "  顯示處理進度、錯誤與被跳過的檔案。\n\n"
    "若需進一步協助，請聯繫 grece@ford.com.tr"

    },
    # ---- THAI ----
    "th_TH": {
        "American": "อังกฤษแบบอเมริกัน",
        "English": "อังกฤษแบบบริติช",
        "Spanish": "สเปน",
        "Mexican Spanish": "สเปนเม็กซิกัน",
        "Brazilian Portuguese": "โปรตุเกสบราซิล",
        "German": "เยอรมัน",
        "Cologne German": "เยอรมันโคโลญ",
        "Dutch": "ดัตช์",
        "Turkish": "ตุรกี",
        "Afrikaans": "แอฟริกันส์",
        "Romanian": "โรมาเนีย",
        "Russian": "รัสเซีย",
        "Simplified Chinese": "จีนตัวย่อ",
        "Traditional Chinese": "จีนตัวเต็ม",
        "Thai": "ไทย",
        "Venezuelan Spanish": "สเปนเวเนซุเอลา",
        "Argentinian Spanish": "สเปนอาร์เจนตินา",
        "Vietnamese": "เวียดนาม",
        "Australian English": "อังกฤษออสเตรเลีย",
        "Language": "ภาษา",
        "Select language": "เลือกภาษา",
        "OK": "ตกลง",
        "Cancel": "ยกเลิก",
        "Exit": "ออก",
        "Apply": "นำไปใช้",
        "Yes": "ใช่",
        "No": "ไม่",
        "Root Folder:": "โฟลเดอร์หลัก:",
        "Mapping File:": "ไฟล์แมปปิ้ง:",
        "Browse": "เรียกดู",
        "Last Monthly  (F)": "เดือนล่าสุด (F)",
        "Prev Monthly  (E)": "เดือนก่อนหน้า (E)",
        "Last Weekly   (M)": "สัปดาห์ล่าสุด (M)",
        "Prev Weekly   (L)": "สัปดาห์ก่อนหน้า (L)",
        "Last Daily    (Q)": "วันล่าสุด (Q)",
        "Prev Daily    (P)": "วันก่อนหน้า (P)",
        "Shift Values:": "เลื่อนค่า:",
        "Monthly (D→C–F→E)": "รายเดือน (D→C–F→E)",
        "Weekly  (I→H–M→L)": "รายสัปดาห์ (I→H–M→L)",
        "Daily   (P→O–Q→P)": "รายวัน (P→O–Q→P)",
        "Run": "รัน",
        "Done": "เสร็จสิ้น",
        "Complete": "สมบูรณ์",
        "skipped": "ข้าม",
        "Column\nUpdate\nOptions": "คอลัมน์\nอัปเดต\nตัวเลือก",
        "Help": "ช่วยเหลือ",
        "Help_Message":
    "คู่มือการใช้งาน QLS Night Letter Tool\n\n"
    "โฟลเดอร์หลัก:\n"
    "  เลือกโฟลเดอร์หลักที่มีโฟลเดอร์ย่อยของโรงงาน (เช่น CAL, WO CAL, FCPA)\n\n"
    "ไฟล์แมปปิ้ง:\n"
    "  เลือกไฟล์ Sticker_Mapping.xlsx เพื่อแมปเนื้อหา PDF กับคอลัมน์ Excel\n\n"
    "ตัวเลือกการอัปเดตคอลัมน์:\n"
    "  เลือกข้อมูลที่ต้องการอัปเดตในไฟล์ Excel หลัก:\n"
    "    - รายเดือนล่าสุด/ก่อนหน้า (F/E)\n"
    "    - รายสัปดาห์ล่าสุด/ก่อนหน้า (M/L)\n"
    "    - รายวันล่าสุด/ก่อนหน้า (Q/P)\n\n"
    "เลื่อนค่า:\n"
    "  หากเลือกไว้ จะเลื่อนค่าก่อนหน้าก่อนบันทึกค่าลงคอลัมน์ที่เลือก\n\n"
    "บันทึกข้อความ:\n"
    "  แสดงความคืบหน้า ข้อผิดพลาด และไฟล์ที่ข้ามระหว่างการประมวลผล\n\n"
    "หากต้องการความช่วยเหลือเพิ่มเติม กรุณาติดต่อ grece@ford.com.tr"

    },
    # ---- VENEZUELAN SPANISH ----
    "es_VE": {
        # Venezuela Spanish = Latin American Spanish, close to es_ES
        "American": "Inglés americano",
        "English": "Inglés británico",
        "Spanish": "Español",
        "Mexican Spanish": "Español mexicano",
        "Brazilian Portuguese": "Portugués brasileño",
        "German": "Alemán",
        "Cologne German": "Alemán de Colonia",
        "Dutch": "Neerlandés",
        "Turkish": "Turco",
        "Afrikaans": "Afrikáans",
        "Romanian": "Rumano",
        "Russian": "Ruso",
        "Simplified Chinese": "Chino simplificado",
        "Traditional Chinese": "Chino tradicional",
        "Thai": "Tailandés",
        "Venezuelan Spanish": "Español venezolano",
        "Argentinian Spanish": "Español argentino",
        "Vietnamese": "Vietnamita",
        "Australian English": "Inglés australiano",
        "Language": "Idioma",
        "Select language": "Seleccionar idioma",
        "OK": "Aceptar",
        "Cancel": "Cancelar",
        "Exit": "Salir",
        "Apply": "Aplicar",
        "Yes": "Sí",
        "No": "No",
        "Root Folder:": "Carpeta raíz:",
        "Mapping File:": "Archivo de mapeo:",
        "Browse": "Examinar",
        "Last Monthly  (F)": "Último mensual (F)",
        "Prev Monthly  (E)": "Mensual anterior (E)",
        "Last Weekly   (M)": "Último semanal (M)",
        "Prev Weekly   (L)": "Semanal anterior (L)",
        "Last Daily    (Q)": "Último diario (Q)",
        "Prev Daily    (P)": "Diario anterior (P)",
        "Shift Values:": "Desplazar valores:",
        "Monthly (D→C–F→E)": "Mensual (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Semanal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Diario (P→O–Q→P)",
        "Run": "Ejecutar",
        "Done": "Hecho",
        "Complete": "Completo",
        "skipped": "omitido",
        "Column\nUpdate\nOptions": "Columna\nActualizar\nOpciones",
        "Help": "Ayuda",
        "Help_Message":
    "Ayuda QLS Night Letter Tool\n\n"
    "Carpeta raíz:\n"
    "  Seleccione la carpeta principal que contiene las subcarpetas de planta (ej: CAL, WO CAL, FCPA).\n\n"
    "Archivo de mapeo:\n"
    "  Seleccione Sticker_Mapping.xlsx para mapear el contenido del PDF a columnas de Excel.\n\n"
    "Opciones de actualización de columna:\n"
    "  Seleccione qué métricas desea actualizar en los archivos maestros de Excel:\n"
    "    - Mensual último/anterior (F/E)\n"
    "    - Semanal último/anterior (M/L)\n"
    "    - Diario último/anterior (Q/P)\n\n"
    "Desplazar valores:\n"
    "  Si está marcado, desplaza los valores antiguos antes de escribir los nuevos en las columnas seleccionadas.\n\n"
    "Registro de texto:\n"
    "  Muestra el progreso, errores y archivos omitidos durante el procesamiento.\n\n"
    "Si necesita más ayuda, contacte a grece@ford.com.tr"

    },
    # ---- ARGENTINIAN SPANISH ----
    "es_AR": {
        # Argentinian Spanish = close to es_ES
        "American": "Inglés americano",
        "English": "Inglés británico",
        "Spanish": "Español",
        "Mexican Spanish": "Español mexicano",
        "Brazilian Portuguese": "Portugués brasileño",
        "German": "Alemán",
        "Cologne German": "Alemán de Colonia",
        "Dutch": "Neerlandés",
        "Turkish": "Turco",
        "Afrikaans": "Afrikáans",
        "Romanian": "Rumano",
        "Russian": "Ruso",
        "Simplified Chinese": "Chino simplificado",
        "Traditional Chinese": "Chino tradicional",
        "Thai": "Tailandés",
        "Venezuelan Spanish": "Español venezolano",
        "Argentinian Spanish": "Español argentino",
        "Vietnamese": "Vietnamita",
        "Australian English": "Inglés australiano",
        "Language": "Idioma",
        "Select language": "Seleccionar idioma",
        "OK": "Aceptar",
        "Cancel": "Cancelar",
        "Exit": "Salir",
        "Apply": "Aplicar",
        "Yes": "Sí",
        "No": "No",
        "Root Folder:": "Carpeta raíz:",
        "Mapping File:": "Archivo de mapeo:",
        "Browse": "Examinar",
        "Last Monthly  (F)": "Último mensual (F)",
        "Prev Monthly  (E)": "Mensual anterior (E)",
        "Last Weekly   (M)": "Último semanal (M)",
        "Prev Weekly   (L)": "Semanal anterior (L)",
        "Last Daily    (Q)": "Último diario (Q)",
        "Prev Daily    (P)": "Diario anterior (P)",
        "Shift Values:": "Desplazar valores:",
        "Monthly (D→C–F→E)": "Mensual (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Semanal (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Diario (P→O–Q→P)",
        "Run": "Ejecutar",
        "Done": "Hecho",
        "Complete": "Completo",
        "skipped": "omitido",
        "Column\nUpdate\nOptions": "Columna\nActualizar\nOpciones",
        "Help": "Ayuda",
        "Help_Message":
    "Ayuda QLS Night Letter Tool\n\n"
    "Carpeta raíz:\n"
    "  Seleccione la carpeta principal que contiene las subcarpetas de planta (ej: CAL, WO CAL, FCPA).\n\n"
    "Archivo de mapeo:\n"
    "  Seleccione Sticker_Mapping.xlsx para mapear el contenido del PDF a columnas de Excel.\n\n"
    "Opciones de actualización de columna:\n"
    "  Seleccione qué métricas desea actualizar en los archivos maestros de Excel:\n"
    "    - Mensual último/anterior (F/E)\n"
    "    - Semanal último/anterior (M/L)\n"
    "    - Diario último/anterior (Q/P)\n\n"
    "Desplazar valores:\n"
    "  Si está marcado, desplaza los valores antiguos antes de escribir los nuevos en las columnas seleccionadas.\n\n"
    "Registro de texto:\n"
    "  Muestra el progreso, errores y archivos omitidos durante el procesamiento.\n\n"
    "Si necesita más ayuda, contacte a grece@ford.com.tr"

    },
    # ---- VIETNAMESE ----
    "vi_VN": {
        "American": "Tiếng Anh Mỹ",
        "English": "Tiếng Anh Anh",
        "Spanish": "Tiếng Tây Ban Nha",
        "Mexican Spanish": "Tiếng Tây Ban Nha Mexico",
        "Brazilian Portuguese": "Tiếng Bồ Đào Nha Brazil",
        "German": "Tiếng Đức",
        "Cologne German": "Tiếng Đức Cologne",
        "Dutch": "Tiếng Hà Lan",
        "Turkish": "Tiếng Thổ Nhĩ Kỳ",
        "Afrikaans": "Tiếng Afrikaans",
        "Romanian": "Tiếng Rumani",
        "Russian": "Tiếng Nga",
        "Simplified Chinese": "Tiếng Trung giản thể",
        "Traditional Chinese": "Tiếng Trung phồn thể",
        "Thai": "Tiếng Thái",
        "Venezuelan Spanish": "Tiếng Tây Ban Nha Venezuela",
        "Argentinian Spanish": "Tiếng Tây Ban Nha Argentina",
        "Vietnamese": "Tiếng Việt",
        "Australian English": "Tiếng Anh Úc",
        "Language": "Ngôn ngữ",
        "Select language": "Chọn ngôn ngữ",
        "OK": "OK",
        "Cancel": "Hủy",
        "Exit": "Thoát",
        "Apply": "Áp dụng",
        "Yes": "Có",
        "No": "Không",
        "Root Folder:": "Thư mục gốc:",
        "Mapping File:": "Tệp ánh xạ:",
        "Browse": "Duyệt",
        "Last Monthly  (F)": "Tháng gần nhất (F)",
        "Prev Monthly  (E)": "Tháng trước (E)",
        "Last Weekly   (M)": "Tuần gần nhất (M)",
        "Prev Weekly   (L)": "Tuần trước (L)",
        "Last Daily    (Q)": "Ngày gần nhất (Q)",
        "Prev Daily    (P)": "Ngày trước (P)",
        "Shift Values:": "Dịch chuyển giá trị:",
        "Monthly (D→C–F→E)": "Hàng tháng (D→C–F→E)",
        "Weekly  (I→H–M→L)": "Hàng tuần (I→H–M→L)",
        "Daily   (P→O–Q→P)": "Hàng ngày (P→O–Q→P)",
        "Run": "Chạy",
        "Done": "Xong",
        "Complete": "Hoàn tất",
        "skipped": "bỏ qua",
        "Column\nUpdate\nOptions": "Cột\nCập nhật\nTùy chọn",
        "Help": "Trợ giúp",
        "Help_Message":
    "Hướng dẫn sử dụng QLS Night Letter Tool\n\n"
    "Thư mục gốc:\n"
    "  Chọn thư mục chính chứa các thư mục con của nhà máy (ví dụ: CAL, WO CAL, FCPA).\n\n"
    "Tệp ánh xạ:\n"
    "  Duyệt đến tệp Sticker_Mapping.xlsx để ánh xạ nội dung PDF sang các cột Excel.\n\n"
    "Tùy chọn cập nhật cột:\n"
    "  Chọn các chỉ số bạn muốn cập nhật trong các tệp Excel chính:\n"
    "    - Tháng cuối/trước (F/E)\n"
    "    - Tuần cuối/trước (M/L)\n"
    "    - Ngày cuối/trước (Q/P)\n\n"
    "Dịch chuyển giá trị:\n"
    "  Nếu được chọn, các giá trị cũ sẽ được dịch chuyển trước khi ghi giá trị mới vào các cột đã chọn.\n\n"
    "Nhật ký văn bản:\n"
    "  Hiển thị tiến trình, lỗi và các tệp bị bỏ qua trong quá trình xử lý.\n\n"
    "Nếu cần hỗ trợ thêm, hãy liên hệ grece@ford.com.tr"

    },
}
