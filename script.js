const baseShortcuts = [
  { id: "s1", type: "shortcut", cat: "mov", syntax: "Win: Ctrl + Arrows | Mac: Cmd + Arrows", 
    en: { name: "Jump to Edge", shortDesc: "Move selection to data region edge", desc: "Mastering these commands allows you to work without touching the mouse." },
    es: { name: "Ir al borde de la región", shortDesc: "Mueve la selección al borde de datos", desc: "Dominar estos comandos te permitirá trabajar sin tocar el ratón." },
    pt: { name: "Pular para a Borda", shortDesc: "Mover seleção para a borda da região de dados", desc: "Dominar esses comandos permite trabalhar sem tocar no mouse." },
    id: { name: "Lompat ke Tepi", shortDesc: "Pindahkan pilihan ke tepi area data", desc: "Menguasai perintah ini memungkinkan Anda bekerja tanpa menyentuh mouse." }
  },
  { id: "s2", type: "shortcut", cat: "mov", syntax: "Win: Ctrl + Shift + Arrows | Mac: Cmd + Shift + Arrows",
    en: { name: "Select to End", shortDesc: "Select to the edge of data", desc: "Instantly highlights all data in a row or column up to the first empty cell." },
    es: { name: "Seleccionar hasta el final", shortDesc: "Selecciona hasta el final de los datos", desc: "Resalta instantáneamente todos los datos en una fila o columna hasta la primera celda vacía." },
    pt: { name: "Selecionar até o Fim", shortDesc: "Selecionar até a borda dos dados", desc: "Destaca instantaneamente todos os dados em uma linha ou coluna até a primeira célula vazia." },
    id: { name: "Pilih hingga Akhir", shortDesc: "Pilih hingga tepi data", desc: "Langsung menyorot semua data dalam baris atau kolom hingga sel kosong pertama." }
  },
  { id: "s3", type: "shortcut", cat: "mov", syntax: "Win: Ctrl + E / Ctrl + A | Mac: Cmd + A",
    en: { name: "Select Table/All", shortDesc: "Select the current data block", desc: "Selects the entire contiguous data range." },
    es: { name: "Seleccionar toda la tabla", shortDesc: "Selecciona la tabla actual", desc: "Selecciona todo el rango de datos contiguos." },
    pt: { name: "Selecionar Tabela", shortDesc: "Selecionar o bloco de dados atual", desc: "Seleciona todo o intervalo de dados contíguos." },
    id: { name: "Pilih Tabel", shortDesc: "Pilih blok data saat ini", desc: "Memilih seluruh rentang data yang berbatasan." }
  },
  { id: "s4", type: "shortcut", cat: "mov", syntax: "Win: Ctrl + PgUp/PgDn | Mac: Opt + Left/Right",
    en: { name: "Switch Tabs", shortDesc: "Move between worksheet tabs", desc: "Quickly navigate through different sheets." },
    es: { name: "Cambiar pestañas", shortDesc: "Mueve entre pestañas del libro", desc: "Navega rápidamente a través de diferentes hojas." },
    pt: { name: "Alternar Abas", shortDesc: "Mover entre as abas da planilha", desc: "Navegue rapidamente por diferentes planilhas." },
    id: { name: "Ganti Tab", shortDesc: "Pindah antar tab lembar kerja", desc: "Navigasi cepat melalui berbagai lembar kerja." }
  },
  { id: "s5", type: "shortcut", cat: "mov", syntax: "Shift + Arrows",
    en: { name: "Extend Selection", shortDesc: "Extend selection by one cell", desc: "Fine-tune your highlighted area one cell at a time." },
    es: { name: "Extender selección", shortDesc: "Extiende la selección una celda", desc: "Ajusta tu área seleccionada celda por celda." },
    pt: { name: "Estender Seleção", shortDesc: "Estender a seleção por uma célula", desc: "Ajuste sua área destacada uma célula por vez." },
    id: { name: "Perluas Pilihan", shortDesc: "Perluas pilihan sebanyak satu sel", desc: "Sesuaikan area yang disorot sel per sel." }
  },
  { id: "s16", type: "shortcut", cat: "mov", syntax: "Shift+Space / Ctrl+Space",
    en: { name: "Row/Column Select", shortDesc: "Select entire row or column", desc: "Shift+Space selects the row, Ctrl+Space selects the column." },
    es: { name: "Selección Fila/Columna", shortDesc: "Seleccionar fila o columna entera", desc: "Shift+Espacio selecciona la fila, Ctrl+Espacio selecciona la columna." },
    pt: { name: "Selecionar Linha/Coluna", shortDesc: "Selecionar linha ou coluna inteira", desc: "Shift+Espaço seleciona a linha, Ctrl+Espaço seleciona a coluna." },
    id: { name: "Pilih Baris/Kolom", shortDesc: "Pilih seluruh baris atau kolom", desc: "Shift+Spasi memilih baris, Ctrl+Spasi memilih kolom." }
  },

  { id: "s6", type: "shortcut", cat: "form", syntax: "Ctrl + Shift + $",
    en: { name: "Currency Format", shortDesc: "Apply currency format", desc: "Formats selected numbers as currency." },
    es: { name: "Formato Moneda", shortDesc: "Aplica formato Moneda", desc: "Aplica formato de moneda." },
    pt: { name: "Formato de Moeda", shortDesc: "Aplicar formato de moeda", desc: "Formata os números selecionados como moeda." },
    id: { name: "Format Mata Uang", shortDesc: "Terapkan format mata uang", desc: "Memformat angka yang dipilih sebagai mata uang." }
  },
  { id: "s7", type: "shortcut", cat: "form", syntax: "Ctrl + Shift + %",
    en: { name: "Percent Format", shortDesc: "Apply percentage format", desc: "Converts decimals to neat percentages." },
    es: { name: "Formato Porcentaje", shortDesc: "Aplica formato Porcentaje", desc: "Convierte decimales a porcentajes." },
    pt: { name: "Formato de Porcentagem", shortDesc: "Aplicar formato de porcentagem", desc: "Converte decimais p/ porcentagens." },
    id: { name: "Format Persentase", shortDesc: "Terapkan format persentase", desc: "Mengubah desimal menjadi persentase." }
  },
  { id: "s8", type: "shortcut", cat: "form", syntax: "Ctrl + B / Cmd + B",
    en: { name: "Bold", shortDesc: "Toggle bold text", desc: "Make your text stand out." },
    es: { name: "Negrita", shortDesc: "Poner/Quitar Negrita", desc: "Haz que tu texto destaque." },
    pt: { name: "Negrito", shortDesc: "Alternar texto em negrito", desc: "Destaque seu texto." },
    id: { name: "Tebal", shortDesc: "Beralih teks tebal", desc: "Buat teks Anda menonjol." }
  },
  { id: "s9", type: "shortcut", cat: "form", syntax: "Ctrl + 1 / Cmd + 1",
    en: { name: "Format Cells", shortDesc: "Open Format Cells dialog", desc: "The master menu for cell formatting." },
    es: { name: "Menú Formato", shortDesc: "Abrir menú de Formato", desc: "El menú maestro para todo el formato." },
    pt: { name: "Formatar Células", shortDesc: "Abre diálogo Formatar Células", desc: "O menu mestre de formatação." },
    id: { name: "Format Sel", shortDesc: "Buka dialog Format Sel", desc: "Menu utama untuk pemformatan sel." }
  },
  { id: "s10", type: "shortcut", cat: "form", syntax: "Alt + = / Cmd + Shift + T",
    en: { name: "AutoSum", shortDesc: "Insert AutoSum formula", desc: "Automatically writes a SUM formula." },
    es: { name: "AutoSuma", shortDesc: "Insertar AutoSuma", desc: "Escribe automáticamente una fórmula SUMA." },
    pt: { name: "AutoSoma", shortDesc: "Inserir fórmula de AutoSoma", desc: "Escreve automaticamente uma fórmula SOMA." },
    id: { name: "Jumlah Otomatis", shortDesc: "Sisipkan rumus AutoSum", desc: "Secara otomatis menulis rumus SUM (Jumlah)." }
  },
  { id: "s22", type: "shortcut", cat: "form", syntax: "F4",
    en: { name: "Absolute Reference", shortDesc: "Toggle absolute/relative ref", desc: "Locks cell references with $ signs." },
    es: { name: "Referencia Absoluta", shortDesc: "Alternar ref absoluta", desc: "Fija referencias de celda con signos $." },
    pt: { name: "Referência Absoluta", shortDesc: "Alternar ref absoluta", desc: "Bloqueia referências de células com cifrões $." },
    id: { name: "Referensi Absolut", shortDesc: "Beralih referensi absolut", desc: "Mengunci referensi sel dengan tanda $." }
  },
  { id: "s25", type: "shortcut", cat: "form", syntax: "Alt + Enter",
    en: { name: "Insert Line Break", shortDesc: "Line break inside cell", desc: "Starts a new line within the same active cell." },
    es: { name: "Salto de Línea", shortDesc: "Salto de línea dentro de celda", desc: "Inicia una nueva línea dentro de la celda activa." },
    pt: { name: "Quebra de Linha", shortDesc: "Quebra de linha na célula", desc: "Inicia uma nova linha dentro da mesma célula." },
    id: { name: "Sisipkan Jeda Baris", shortDesc: "Hentikan baris dalam sel", desc: "Memulai baris baru dalam sel yang sama." }
  },

  { id: "s17", type: "shortcut", cat: "bas", syntax: "Ctrl + C / X / V",
    en: { name: "Copy/Cut/Paste", shortDesc: "Standard clipboard tools", desc: "Essential key commands for moving data." },
    es: { name: "Copiar/Cortar/Pegar", shortDesc: "Herramientas de portapapeles", desc: "Comandos esenciales para mover datos." },
    pt: { name: "Copiar/Recortar/Colar", shortDesc: "Ferramentas da área de transferência", desc: "Comandos essenciais p/ mover dados." },
    id: { name: "Salin/Potong/Tempel", shortDesc: "Alat papan klip standar", desc: "Perintah tombol penting untuk memindahkan data." }
  },
  { id: "s18", type: "shortcut", cat: "bas", syntax: "Ctrl + Z / Ctrl + Y",
    en: { name: "Undo/Redo", shortDesc: "Reverse or repeat actions", desc: "Quickly fix mistakes or re-apply an action." },
    es: { name: "Deshacer/Rehacer", shortDesc: "Revertir o repetir", desc: "Comandos para arreglar errores rápidamente." },
    pt: { name: "Desfazer/Refazer", shortDesc: "Reverter ou repetir", desc: "Corrija rapidamente erros ou repita uma ação." },
    id: { name: "Urungkan/Ulangi", shortDesc: "Pembalikan atau ulangi aksi", desc: "Memperbaiki kesalahan dengan cepat." }
  },
  { id: "s21", type: "shortcut", cat: "bas", syntax: "F2",
    en: { name: "Edit Cell", shortDesc: "Jump into cell edit mode", desc: "Places the cursor at the end of the cell content." },
    es: { name: "Editar Celda", shortDesc: "Modo edición en celda", desc: "Pone el cursor al final de la celda para editar." },
    pt: { name: "Editar Célula", shortDesc: "Modo de edição da célula", desc: "Coloca o cursor no final do conteúdo da célula." },
    id: { name: "Edit Sel", shortDesc: "Masuk ke mode edit sel", desc: "Menempatkan kursor di akhir konten sel." }
  },
  { id: "s26", type: "shortcut", cat: "bas", syntax: "Ctrl + ;",
    en: { name: "Current Date", shortDesc: "Insert today's date", desc: "Stamps the current date as a static value." },
    es: { name: "Fecha Actual", shortDesc: "Insertar fecha de hoy", desc: "Inserta la fecha actual como valor estático." },
    pt: { name: "Data Atual", shortDesc: "Inserir a data de hoje", desc: "Insere a data atual como um valor estático." },
    id: { name: "Tanggal Saat Ini", shortDesc: "Sisipkan tanggal hari ini", desc: "Memasukkan tanggal saat ini sebagai nilai statis." }
  },
  { id: "s27", type: "shortcut", cat: "bas", syntax: "Ctrl + Shift + :",
    en: { name: "Current Time", shortDesc: "Insert current time", desc: "Stamps the current time as a static value." },
    es: { name: "Hora Actual", shortDesc: "Insertar hora actual", desc: "Inserta la hora actual como valor estático." },
    pt: { name: "Hora Atual", shortDesc: "Inserir a hora atual", desc: "Insere a hora atual como um valor estático." },
    id: { name: "Waktu Saat Ini", shortDesc: "Sisipkan waktu saat ini", desc: "Memasukkan waktu saat ini sebagai nilai statis." }
  },

  { id: "s11", type: "shortcut", cat: "gest", syntax: "Ctrl + Shift + L",
    en: { name: "Toggle Filters", shortDesc: "Turn filters on or off", desc: "Add or remove dropdown filters." },
    es: { name: "Filtros", shortDesc: "Activar/Desactivar Filtros", desc: "Añade o quita instantáneamente filtros." },
    pt: { name: "Alternar Filtros", shortDesc: "Ativa e desativa filtros", desc: "Adiciona ou remove filtros de menu suspenso." },
    id: { name: "Alihkan Filter", shortDesc: "Nyalakan/matikan filter", desc: "Pasang atau lepas filter tarik-turun." }
  },
  { id: "s12", type: "shortcut", cat: "gest", syntax: "Ctrl + T",
    en: { name: "Create Table", shortDesc: "Create Excel Table", desc: "Converts raw data into a Table." },
    es: { name: "Crear Tabla", shortDesc: "Crear Tabla oficial", desc: "Convierte datos sin procesar en Tabla oficial." },
    pt: { name: "Criar Tabela", shortDesc: "Criar uma Tabela do Excel", desc: "Converte dados brutos numa Tabela." },
    id: { name: "Buat Tabel", shortDesc: "Buat Tabel Excel", desc: "Konversikan data mentah menjadi Tabel resmi." }
  },
  { id: "s13", type: "shortcut", cat: "gest", syntax: "Ctrl + Alt + V",
    en: { name: "Paste Special", shortDesc: "Open Paste Special", desc: "Allows pasting values, formats, etc." },
    es: { name: "Pegado Especial", shortDesc: "Abrir Pegado Especial", desc: "Permite pegar solo valores, o formatos." },
    pt: { name: "Colar Especial", shortDesc: "Abre Colar Especial", desc: "Permite colar valores, formatos, etc." },
    id: { name: "Tempel Spesial", shortDesc: "Buka Tempel Spesial", desc: "Izinkan menempelkan hanya nilai, format dll." }
  },
  { id: "s14", type: "shortcut", cat: "gest", syntax: "Delete",
    en: { name: "Clear Content", shortDesc: "Delete cell contents", desc: "Removes text without deleting formats." },
    es: { name: "Borrar todo", shortDesc: "Borrar el contenido", desc: "Elimina contenido sin borrar el formato." },
    pt: { name: "Limpar Conteúdo", shortDesc: "Apagar conteúdo da célula", desc: "Remove texto sem apagar a formatação." },
    id: { name: "Hapus Konten", shortDesc: "Hapus isi sel", desc: "Menghapus teks tanpa jadwal pemformatan." }
  },
  { id: "s19", type: "shortcut", cat: "gest", syntax: "Ctrl + S",
    en: { name: "Save", shortDesc: "Save the file", desc: "Quickly save your progress." },
    es: { name: "Guardar", shortDesc: "Guardar archivo", desc: "Guarda tu progreso rápidamente." },
    pt: { name: "Salvar", shortDesc: "Salvar o arquivo", desc: "Salve seu progresso rapidamente." },
    id: { name: "Simpan", shortDesc: "Simpan file", desc: "Simpan progres Anda dengan cepat." }
  },
  { id: "s20", type: "shortcut", cat: "gest", syntax: "Ctrl + F",
    en: { name: "Search/Find", shortDesc: "Find data in workbook", desc: "Opens the Find tool to search for text/values." },
    es: { name: "Buscar", shortDesc: "Buscar datos en el libro", desc: "Abre la herramienta para buscar texto/valores." },
    pt: { name: "Localizar", shortDesc: "Localizar na pasta de trabalho", desc: "Abre a ferramenta Localizar dados/texto." },
    id: { name: "Cari/Temukan", shortDesc: "Temukan data di buku kerja", desc: "Buka Cari untuk menemukan teks/nilai." }
  },
  { id: "s23", type: "shortcut", cat: "gest", syntax: "Ctrl + N",
    en: { name: "New Workbook", shortDesc: "Open a blank workbook", desc: "Creates a new blank Excel file." },
    es: { name: "Nuevo Libro", shortDesc: "Abrir libro en blanco", desc: "Crea un archivo de Excel nuevo y en blanco." },
    pt: { name: "Nova Pasta de Trabalho", shortDesc: "Abrir pasta em branco", desc: "Cria um novo arquivo Excel em branco." },
    id: { name: "Buku Kerja Baru", shortDesc: "Buka buku kerja kosong", desc: "Buat file Excel kosong baru." }
  },
  { id: "s24", type: "shortcut", cat: "gest", syntax: "Ctrl + P",
    en: { name: "Print", shortDesc: "Open print menu", desc: "Prepares your current worksheet for printing." },
    es: { name: "Imprimir", shortDesc: "Abrir menú de impresión", desc: "Prepara la hoja para imprimir." },
    pt: { name: "Imprimir", shortDesc: "Abrir menu de impressão", desc: "Prepara sua planilha para impressão." },
    id: { name: "Cetak", shortDesc: "Buka menu cetak", desc: "Siapkan lembar cetak Anda." }
  }
];

const enFuncNames = {
  SUMIFS: "SUMIFS", COUNTIFS: "COUNTIFS", AVERAGEIF: "AVERAGEIF", AVERAGEIFS: "AVERAGEIFS", 
  MINIFS: "MINIFS", MAXIFS: "MAXIFS", AND: "AND", OR: "OR", CLEAN: "CLEAN", SUBSTITUTE: "SUBSTITUTE",
  TEXTSPLIT: "TEXTSPLIT", TEXTBEFORE: "TEXTBEFORE", TEXTAFTER: "TEXTAFTER", LEFT: "LEFT", MID: "MID", RIGHT: "RIGHT",
  TODAY: "TODAY", NOW: "NOW", EOMONTH: "EOMONTH", EDATE: "EDATE", WORKDAY: "WORKDAY",
  VLOOKUP: "VLOOKUP", HLOOKUP: "HLOOKUP", MATCH: "MATCH", INDEX: "INDEX", ROUND: "ROUND", IFERROR: "IFERROR", 
  ISBLANK: "ISBLANK", CONCATENATE: "CONCATENATE", LEN: "LEN", FIND: "FIND", SUM: "SUM", AVERAGE: "AVERAGE",
  COUNT: "COUNT", COUNTA: "COUNTA", "MAX / MIN": "MAX/MIN", "MODE.SNGL": "MODE.SNGL", TRIM: "TRIM",
  PROPER: "PROPER", "UPPER / LOWER": "UPPER/LOWER", VALUE: "VALUE", TEXTJOIN: "TEXTJOIN",
  IF: "IF", COUNTIF: "COUNTIF", XLOOKUP: "XLOOKUP", FILTER: "FILTER", UNIQUE: "UNIQUE", LET: "LET", "INDEX & MATCH": "INDEX & MATCH"
};

const esFuncNames = {
  SUMIFS: "SUMAR.SI.CONJUNTO", COUNTIFS: "CONTAR.SI.CONJUNTO", AVERAGEIF: "PROMEDIO.SI", AVERAGEIFS: "PROMEDIO.SI.CONJUNTO", MINIFS: "MIN.SI.CONJUNTO", MAXIFS: "MAX.SI.CONJUNTO",
  AND: "Y", OR: "O", CLEAN: "LIMPIAR", SUBSTITUTE: "SUSTITUIR", TEXTSPLIT: "DIVIDIRTEXTO", TEXTBEFORE: "TEXTOANTES", TEXTAFTER: "TEXTODESPUES", LEFT: "IZQUIERDA", MID: "EXTRAE", RIGHT: "DERECHA",
  TODAY: "HOY", NOW: "AHORA", EOMONTH: "FIN.MES", EDATE: "FECHA.MES", WORKDAY: "DIA.LAB",
  VLOOKUP: "BUSCARV", HLOOKUP: "BUSCARH", MATCH: "COINCIDIR", INDEX: "INDICE", ROUND: "REDONDEAR", IFERROR: "SI.ERROR",
  ISBLANK: "ESBLANCO", CONCATENATE: "CONCATENAR", LEN: "LARGO", FIND: "ENCONTRAR", SUM: "SUMA", AVERAGE: "PROMEDIO",
  COUNT: "CONTAR", COUNTA: "CONTARA", "MAX / MIN": "MAX/MIN", "MODE.SNGL": "MODA.UNO", TRIM: "ESPACIOS",
  PROPER: "NOMPROPIO", "UPPER / LOWER": "MAYUS/MINUS", VALUE: "VALOR", TEXTJOIN: "UNIRCADENAS",
  IF: "SI", COUNTIF: "CONTAR.SI", XLOOKUP: "BUSCARX", FILTER: "FILTRAR", UNIQUE: "UNICOS", LET: "LET", "INDEX & MATCH": "INDICE & COINCIDIR"
};

const ptFuncNames = {
  SUMIFS: "SOMASES", COUNTIFS: "CONT.SES", AVERAGEIF: "MÉDIASE", AVERAGEIFS: "MÉDIASES", MINIFS: "MÍNSES", MAXIFS: "MÁXSES",
  AND: "E", OR: "OU", CLEAN: "TIRAR", SUBSTITUTE: "SUBSTITUIR", TEXTSPLIT: "DIVIDIRTEXTO", TEXTBEFORE: "TEXTOANTES", TEXTAFTER: "TEXTODEPOIS", LEFT: "ESQUERDA", MID: "EXT.TEXTO", RIGHT: "DIREITA",
  TODAY: "HOJE", NOW: "AGORA", EOMONTH: "FIMMÊS", EDATE: "DATAM", WORKDAY: "DIATRABALHO",
  VLOOKUP: "PROCV", HLOOKUP: "PROCH", MATCH: "CORRESP", INDEX: "ÍNDICE", ROUND: "ARRED", IFERROR: "SEERRO",
  ISBLANK: "ÉCÉL.VAZIA", CONCATENATE: "CONCATENAR", LEN: "NÚM.CARACT", FIND: "PROCURAR", SUM: "SOMA", AVERAGE: "MÉDIA",
  COUNT: "CONT.NÚM", COUNTA: "CONT.VALORES", "MAX / MIN": "MÁX/MÍN", "MODE.SNGL": "MODA.ÚNICO", TRIM: "ARRUMAR",
  PROPER: "PRI.MAIÚSCULA", "UPPER / LOWER": "MAIÚSCULA/MINÚSCULA", VALUE: "VALOR", TEXTJOIN: "UNIRTEXTO",
  IF: "SE", COUNTIF: "CONT.SE", XLOOKUP: "PESQUISAX", FILTER: "FILTRO", UNIQUE: "ÚNICO", LET: "LET", "INDEX & MATCH": "ÍNDICE & CORRESP"
};

const idFuncNames = { ...enFuncNames };

const formulaDefs = [
  { id: "f1", cat: "bas", fn: "SUM", syn: "(range)" }, { id: "f2", cat: "bas", fn: "AVERAGE", syn: "(range)" },
  { id: "f3", cat: "bas", fn: "COUNT", syn: "(range)" }, { id: "f4", cat: "bas", fn: "COUNTA", syn: "(range)" },
  { id: "f5", cat: "bas", fn: "MAX / MIN", syn: "(range)" }, { id: "f6", cat: "bas", fn: "MODE.SNGL", syn: "(range)" },
  { id: "f7", cat: "limp", fn: "TRIM", syn: "(text)" }, { id: "f8", cat: "limp", fn: "PROPER", syn: "(text)" },
  { id: "f9", cat: "limp", fn: "UPPER / LOWER", syn: "(text)" }, { id: "f10", cat: "limp", fn: "VALUE", syn: "(text)" },
  { id: "f11", cat: "limp", fn: "TEXTJOIN", syn: "(delim, ignore, rng)" }, { id: "f12", cat: "log", fn: "IF", syn: "(test, t, f)" },
  { id: "f13", cat: "log", fn: "IFERROR", syn: "(val, err)" }, { id: "f14", cat: "log", fn: "COUNTIF", syn: "(rng, crit)" },
  { id: "f15", cat: "log", fn: "SUMIFS", syn: "(sum_rng, cr1, ...)" }, { id: "f16", cat: "avz", fn: "XLOOKUP", syn: "(val, arr1, arr2)" },
  { id: "f17", cat: "avz", fn: "INDEX & MATCH", syn: "(rng, MATCH)" }, { id: "f18", cat: "avz", fn: "FILTER", syn: "(rng, crit)" },
  { id: "f19", cat: "avz", fn: "UNIQUE", syn: "(rng)" }, { id: "f20", cat: "avz", fn: "LET", syn: "(n, v, calc)" },
  { id: "f21", cat: "log", fn: "AND", syn: "({arg1}, {arg2}, ...)", desc: { en: "Returns TRUE if all conditions are TRUE.", es: "Devuelve VERDADERO si todos son VERDADEROS.", pt: "Retorna VERDADEIRO se todos forem VERDADEIROS.", id: "Kembalikan BENAR jika semua argumen BENAR." } },
  { id: "f21b", cat: "log", fn: "OR", syn: "({arg1}, {arg2}, ...)", desc: { en: "Returns TRUE if any condition is TRUE.", es: "Devuelve VERDADERO si alguno es VERDADERO.", pt: "Retorna VERDADEIRO se algum for VERDADEIRO.", id: "Kembalikan BENAR jika ada argumen yang BENAR." } },
  { id: "f22", cat: "limp", fn: "CLEAN", syn: "({text})", desc: { en: "Removes non-printable characters.", es: "Elimina caracteres no imprimibles.", pt: "Remove caracteres não imprimíveis.", id: "Menghapus karakter yang tidak dapat dicetak." } },
  { id: "f23", cat: "limp", fn: "SUBSTITUTE", syn: "({text}, {old}, {new})", desc: { en: "Replaces specific text in a string.", es: "Reemplaza texto específico.", pt: "Substitui um texto específico.", id: "Mengganti teks tertentu dengan yang baru." } },
  { id: "f24", cat: "avz", fn: "TEXTSPLIT", syn: "({text}, {delimit})", desc: { en: "Splits text across rows/columns.", es: "Divide texto en filas/columnas.", pt: "Divide texto por linhas/colunas.", id: "Membagi teks ke kolom atau baris." } },
  { id: "f25", cat: "limp", fn: "TEXTBEFORE", syn: "({text}, {delimit})", desc: { en: "Returns text before a character.", es: "Devuelve texto antes de un delimitador.", pt: "Retorna o texto antes de um caractere.", id: "Kembalikan teks sebelum karakter tertentu." } },
  { id: "f26", cat: "limp", fn: "TEXTAFTER", syn: "({text}, {delimit})", desc: { en: "Returns text after a character.", es: "Devuelve texto después de un delimitador.", pt: "Retorna o texto depois de um caractere.", id: "Kembalikan teks setelah karakter tertentu." } },
  { id: "f27", cat: "limp", fn: "LEFT", syn: "({text}, {num})", desc: { en: "Extracts characters from the left.", es: "Extrae caracteres por la izquierda.", pt: "Extrai caracteres da esquerda.", id: "Ekstrak karakter dari sisi kiri." } },
  { id: "f28", cat: "limp", fn: "MID", syn: "({text}, {start}, {num})", desc: { en: "Extracts characters from the middle.", es: "Extrae texto del medio.", pt: "Extrai caracteres do meio da string.", id: "Ekstrak karakter dari tengah teks." } },
  { id: "f29", cat: "limp", fn: "RIGHT", syn: "({text}, {num})", desc: { en: "Extracts characters from the right.", es: "Extrae caracteres por la derecha.", pt: "Extrai caracteres da direita.", id: "Ekstrak karakter dari sisi kanan." } },
  { id: "f30", cat: "log", fn: "COUNTIFS", syn: "({range1}, {crit1}, ...)", desc: { en: "Counts cells meeting multiple criteria.", es: "Cuenta con múltiples criterios.", pt: "Conta com base em vários critérios.", id: "Menghitung dengan berbagai kriteria." } },
  { id: "f31", cat: "log", fn: "AVERAGEIF", syn: "({range}, {crit})", desc: { en: "Average based on one condition.", es: "Promedio con una condición.", pt: "Média baseada numa condição.", id: "Rata-rata dengan satu kondisi." } },
  { id: "f32", cat: "log", fn: "AVERAGEIFS", syn: "({avg_rng}, {crit_rng}, ...)", desc: { en: "Average with multiple conditions.", es: "Promedio con múltiples condiciones.", pt: "Média com várias condições.", id: "Rata-rata dengan multi kondisi." } },
  { id: "f33", cat: "log", fn: "MINIFS", syn: "({min_rng}, {crit_rng}, ...)", desc: { en: "Minimum value meeting conditions.", es: "Mínimo con condiciones.", pt: "Valor mínimo baseando-se em condições.", id: "Nilai minimum dengan kondisi." } },
  { id: "f34", cat: "log", fn: "MAXIFS", syn: "({max_rng}, {crit_rng}, ...)", desc: { en: "Maximum value meeting conditions.", es: "Máximo con condiciones.", pt: "Valor máximo baseando-se em condições.", id: "Nilai maksimal dengan kondisi." } },
  { id: "f35", cat: "bas", fn: "TODAY", syn: "()", desc: { en: "Returns the current date.", es: "Devuelve la fecha de hoy.", pt: "Retorna a data atual.", id: "Kembalikan tanggal hari ini." } },
  { id: "f36", cat: "bas", fn: "NOW", syn: "()", desc: { en: "Returns current date and time.", es: "Devuelve fecha y hora actual.", pt: "Retorna data e hora atuais.", id: "Kembalikan tanggal dan waktu saat ini." } },
  { id: "f37", cat: "avz", fn: "EOMONTH", syn: "({date}, {months})", desc: { en: "End of month after X months.", es: "Fin de mes tras X meses.", pt: "Fim do mês após X meses.", id: "Akhir bulan setelah X bulan." } },
  { id: "f38", cat: "avz", fn: "EDATE", syn: "({date}, {months})", desc: { en: "Exact date after X months.", es: "Mes exacto tras X meses.", pt: "Data exata após X meses.", id: "Tanggal pasti setelah X bulan." } },
  { id: "f39", cat: "avz", fn: "WORKDAY", syn: "({date}, {days}, [hols])", desc: { en: "Date after working days passed.", es: "Día laborable tras X días.", pt: "Data após dias úteis passados.", id: "Tanggal setelah hari kerja berlalu." } },
  { id: "f40", cat: "avz", fn: "VLOOKUP", syn: "({val}, {table}, {col}, 0)", desc: { en: "Vertical lookup in a table.", es: "Búsqueda vertical.", pt: "Pesquisa vertical numa tabela.", id: "Pencarian vertikal dalam tabel." } },
  { id: "f41", cat: "avz", fn: "HLOOKUP", syn: "({val}, {table}, {row}, 0)", desc: { en: "Horizontal lookup.", es: "Búsqueda horizontal.", pt: "Pesquisa horizontal.", id: "Pencarian horizontal." } },
  { id: "f42", cat: "avz", fn: "MATCH", syn: "({val}, {array}, 0)", desc: { en: "Returns relative position.", es: "Devuelve posición relativa.", pt: "Retorna posição relativa.", id: "Kembalikan posisi relatif." } },
  { id: "f43", cat: "avz", fn: "INDEX", syn: "({array}, {row}, {col})", desc: { en: "Value at a given row/column.", es: "Valor en fila/columna dada.", pt: "Valor numa linha/coluna.", id: "Nilai di baris/kolom tertentu." } },
  { id: "f44", cat: "bas", fn: "ROUND", syn: "({num}, {digits})", desc: { en: "Rounds to specified digits.", es: "Redondea a N dígitos.", pt: "Arredonda p/ casas especificadas.", id: "Membulatkan ke digit yang ditentukan." } },
  { id: "f45", cat: "log", fn: "ISBLANK", syn: "({cell})", desc: { en: "Checks if cell is empty.", es: "Verifica si la celda está vacía.", pt: "Verifica se a célula está vazia.", id: "Cek apakah sel kosong." } },
  { id: "f46", cat: "limp", fn: "CONCATENATE", syn: "({arg1}, {arg2}, ...)", desc: { en: "Joins multiple text strings.", es: "Une varias cadenas de texto.", pt: "Junta várias cadeias de texto.", id: "Menggabungkan beberapa string teks." } },
  { id: "f47", cat: "limp", fn: "LEN", syn: "({text})", desc: { en: "Returns the length of a string.", es: "Devuelve longitud de texto.", pt: "Retorna o tamanho da string.", id: "Mengembalikan panjang string." } },
  { id: "f48", cat: "limp", fn: "FIND", syn: "({find}, {within})", desc: { en: "Finds text within a string.", es: "Busca texto en una cadena.", pt: "Encontra texto numa string.", id: "Menemukan teks dalam string." } }
];

const basicDesc = {
  f1: { en: "Calculates the total sum of the selected range.", es: "Suma total del rango.", pt: "Soma total do intervalo.", id: "Jumlah total rentang." },
  f2: { en: "Arithmetic mean.", es: "Media aritmética.", pt: "Média aritmética.", id: "Rata-rata aritmatika." },
  f3: { en: "Count numbers.", es: "Cuenta números.", pt: "Conta números.", id: "Kalkulasi semua angka." },
  f4: { en: "Count non-empty.", es: "Cuenta no vacías.", pt: "Conta as células não vazias.", id: "Menghitung sel tidak kosong." },
  f5: { en: "Extreme values.", es: "Valores extremos.", pt: "Retorna valores extremos.", id: "Dapatkan maksimal minimum." },
  f6: { en: "Most frequent value.", es: "Valor más repetido.", pt: "Valor mais repetido.", id: "Nilai yang paling sering muncul." },
  f7: { en: "Remove extra spaces.", es: "Elimina espacios dobles.", pt: "Remove espaços extras.", id: "Singkirkan spasi dobel." },
  f8: { en: "Capitalize first letters.", es: "Capitaliza palabras.", pt: "Primeiras letras maiúsculas.", id: "Huruf awalan jadi besar." },
  f9: { en: "Convert text case.", es: "Convierte mayús/minús.", pt: "Converte caixa de texto.", id: "Tukar format besar huruf." },
  f10: { en: "Text to number.", es: "Texto a número.", pt: "Texto para número.", id: "Teks menjadi format nomor." },
  f11: { en: "Combine text.", es: "Concatenar moderno.", pt: "Juntar texto moderno.", id: "Gabungkan satu teks utuh." },
  f12: { en: "Conditional logic.", es: "Condición lógica.", pt: "Lógica condicional de planilhas.", id: "Fungsi logis bersyarat." },
  f13: { en: "Handle errors.", es: "Manejo de errores.", pt: "Tratamento de erros.", id: "Tangani error yang terjadi." },
  f14: { en: "Conditional count.", es: "Conteo condicional.", pt: "Contagem condicional da área.", id: "Hitung sesuai kondisi logis." },
  f15: { en: "Multi-conditional sum.", es: "Suma multi-condicional.", pt: "Soma multi-condicional.", id: "Penjumlahan menurut multi kriteria." },
  f16: { en: "Modern lookup.", es: "Búsqueda moderna.", pt: "Pesquisa moderna segura.", id: "Pencarian lanjutan akurat." },
  f17: { en: "2D Lookup.", es: "Búsqueda 2D.", pt: "Pesquisa 2D.", id: "Pencarian bentuk matriks 2D." },
  f18: { en: "Dynamic filtering.", es: "Filtrado dinámico.", pt: "Filtragem dinâmica dos dados.", id: "Menyaring info dinamis." },
  f19: { en: "Extract unique values.", es: "Extraer únicos.", pt: "Extrair valores únicos.", id: "Dapatkan nilai tidak ganda." },
  f20: { en: "Define variables.", es: "Definir variables.", pt: "Definir variáveis da fórmula.", id: "Menentukan variabel khusus." }
};

const UI = {
  en: { title: "Excel Cheat Sheet", subtitle: "Master shortcuts, formulas, and data analysis", searchPlaceholder: "Search shortcuts and formulas...", footerText: "All rights reserved.", resultsText: "Showing {count} results", shortcutsTitle: "Keyboard Shortcuts", formulasTitle: "Essential Formulas", filters: { all: "All", mov: "Movement", form: "Formatting", gest: "Management", bas: "Basics", limp: "Cleaning", log: "Logic", avz: "Advanced" } },
  es: { title: "Hoja de Referencia de Excel", subtitle: "Domina atajos, fórmulas y análisis de datos", searchPlaceholder: "Buscar atajos y fórmulas...", footerText: "Todos los derechos reservados.", resultsText: "Mostrando {count} resultados", shortcutsTitle: "Atajos de Teclado", formulasTitle: "Fórmulas Imprescindibles", filters: { all: "Todo", mov: "Movimiento", form: "Formateo", gest: "Gestión", bas: "Básicas", limp: "Limpieza", log: "Lógicas", avz: "Avanzadas" } },
  pt: { title: "Planilha de Referência de Excel", subtitle: "Domine atalhos, fórmulas e análise de dados", searchPlaceholder: "Pesquisar atalhos e fórmulas...", footerText: "Todos os direitos reservados.", resultsText: "Mostrando {count} resultados", shortcutsTitle: "Atalhos de Teclado", formulasTitle: "Fórmulas Essenciais", filters: { all: "Tudo", mov: "Movimento", form: "Formatação", gest: "Gestão", bas: "Básicas", limp: "Limpeza", log: "Lógicas", avz: "Avançadas" } },
  id: { title: "Lembar Contekan Excel", subtitle: "Kuasai pintasan, rumus, dan analisis data", searchPlaceholder: "Cari pintasan dan rumus...", footerText: "Hak cipta dilindungi undang-undang.", resultsText: "Menampilkan {count} hasil", shortcutsTitle: "Pintasan Keyboard", formulasTitle: "Rumus Penting", filters: { all: "Semua", mov: "Gerakan", form: "Format", gest: "Manajemen", bas: "Dasar", limp: "Pembersihan", log: "Logika", avz: "Lanjutan" } }
};

const funcDicts = { en: enFuncNames, es: esFuncNames, pt: ptFuncNames, id: idFuncNames };
const langs = ['en', 'es', 'pt', 'id'];
const locales = {};

langs.forEach(ls => {
  let langItems = [];
  
  // Build shortcuts
  baseShortcuts.forEach(sc => {
    langItems.push({
      id: sc.id, type: "shortcut", cat: sc.cat, syntax: sc.syntax,
      name: sc[ls].name, shortDesc: sc[ls].shortDesc, desc: sc[ls].desc
    });
  });

  // Build formulas
  formulaDefs.forEach(fd => {
    let fnName = funcDicts[ls][fd.fn] || fd.fn;
    let descObj = fd.desc ? fd.desc : basicDesc[fd.id];
    let fullDesc = descObj ? descObj[ls] : "(No description)";
    let shortDesc = fullDesc.substring(0, 30) + (fullDesc.length > 30 ? '...' : '');
    
    langItems.push({
      id: fd.id, type: "formula", cat: fd.cat,
      name: fnName, syntax: "=" + fnName + (fd.syn || "(...)"),
      shortDesc: shortDesc,
      desc: fullDesc
    });
  });

  locales[ls] = { ...UI[ls], items: langItems };
});

// Global State
let currentLang = 'en'; 
let currentExcelLang = 'en';
let currentFilter = 'all';
let searchQuery = '';

// DOM Elements
const _title = document.getElementById('page-title');
const _subtitle = document.getElementById('page-subtitle');
const _searchInput = document.getElementById('search-input');
const _langSelect = document.getElementById('lang-select');
const _excelLangToggle = document.getElementById('excel-lang-toggle');
const _excelLangText = document.getElementById('excel-lang-text');
const _themeToggle = document.getElementById('theme-toggle');
const _themeIcon = document.getElementById('theme-icon');
const _filtersContainer = document.getElementById('filters-container');
const _shortcutsContainer = document.getElementById('shortcuts-container');
const _formulasContainer = document.getElementById('formulas-container');
const _shortcutsWrapper = document.getElementById('shortcuts-wrapper');
const _formulasWrapper = document.getElementById('formulas-wrapper');
const _shortcutsTitle = document.getElementById('shortcuts-title');
const _formulasTitle = document.getElementById('formulas-title');
const _resultsStats = document.getElementById('results-stats');
const _footerText = document.getElementById('footer-text');
const _fabFilterMenu = document.getElementById('fab-filter-menu');
const _fabFilter = document.getElementById('fab-filter');
const _fabTop = document.getElementById('fab-top');

function init() {
    const savedTheme = localStorage.getItem('theme') || 'light';
    const savedLang = localStorage.getItem('lang') || 'en';
    
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    if (_langSelect) {
        _langSelect.value = savedLang;
        _langSelect.addEventListener('change', (e) => {
            setLanguage(e.target.value);
        });
    }

    setLanguage(savedLang);

    _excelLangToggle.addEventListener('click', () => {
        if (currentLang !== 'en') {
            currentExcelLang = currentExcelLang === 'en' ? currentLang : 'en';
            updateExcelLangButton();
            renderItems();
        }
    });

    _themeToggle.addEventListener('click', () => {
        const currentTheme = document.documentElement.getAttribute('data-theme');
        const newTheme = currentTheme === 'light' ? 'dark' : 'light';
        document.documentElement.setAttribute('data-theme', newTheme);
        localStorage.setItem('theme', newTheme);
        updateThemeIcon(newTheme);
    });

    _searchInput.addEventListener('input', (e) => {
        searchQuery = e.target.value.toLowerCase().trim();
        renderItems();
    });

    _fabTop.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    _fabFilter.addEventListener('click', (e) => {
        e.stopPropagation();
        _fabFilterMenu.classList.toggle('active');
    });

    document.addEventListener('click', (e) => {
        if (!e.target.closest('.fab-container')) {
            _fabFilterMenu.classList.remove('active');
        }
    });

    window.addEventListener('scroll', () => {
        const yPos = window.scrollY;
        if (yPos > 150) {
            _fabTop.classList.add('visible');
            _fabFilter.classList.add('visible');
        } else {
            _fabTop.classList.remove('visible');
            _fabFilter.classList.remove('visible');
        }
    });

    document.getElementById('year').textContent = new Date().getFullYear();
}

function updateThemeIcon(theme) {
    _themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
}

function setLanguage(langCode) {
    currentLang = langCode;
    currentExcelLang = langCode === 'en' ? 'en' : langCode; 
    localStorage.setItem('lang', langCode);
    
    const textData = locales[langCode];

    _title.textContent = textData.title;
    _subtitle.textContent = textData.subtitle;
    _searchInput.placeholder = textData.searchPlaceholder;
    _footerText.textContent = textData.footerText;
    _shortcutsTitle.textContent = textData.shortcutsTitle;
    _formulasTitle.textContent = textData.formulasTitle;
    
    updateExcelLangButton();

    if (!textData.filters[currentFilter]) {
        currentFilter = 'all';
    }

    renderFilters();
    renderItems();
}

function updateExcelLangButton() {
    if (currentLang !== 'en') {
        _excelLangToggle.style.display = 'flex';
        const displayLang = currentExcelLang === 'en' ? 'INGLÉS' : currentLang.toUpperCase();
        _excelLangText.textContent = `Versión ${displayLang}`;
    } else {
        _excelLangToggle.style.display = 'none';
        currentExcelLang = 'en';
    }
}

function renderFilters() {
    _filtersContainer.innerHTML = '';
    _fabFilterMenu.innerHTML = '';
    const filters = locales[currentLang].filters;
    
    for (const [key, name] of Object.entries(filters)) {
        const btnMain = document.createElement('button');
        btnMain.className = `filter-btn ${key === currentFilter ? 'active' : ''}`;
        btnMain.textContent = name;
        btnMain.dataset.filter = key;
        
        const btnFab = document.createElement('button');
        btnFab.className = `filter-btn ${key === currentFilter ? 'active' : ''}`;
        btnFab.textContent = name;
        btnFab.dataset.filter = key;

        const clickHandler = () => {
            currentFilter = key;
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll(`[data-filter="${key}"]`).forEach(b => b.classList.add('active'));
            _fabFilterMenu.classList.remove('active');
            renderItems();
        };

        btnMain.addEventListener('click', clickHandler);
        btnFab.addEventListener('click', clickHandler);
        
        _filtersContainer.appendChild(btnMain);
        _fabFilterMenu.appendChild(btnFab);
    }
}

function renderItems() {
    _shortcutsContainer.innerHTML = '';
    _formulasContainer.innerHTML = '';
    
    _shortcutsWrapper.style.display = 'block';
    _formulasWrapper.style.display = 'block';

    const items = locales[currentLang].items;
    let totalCount = 0;
    let shortcutsCount = 0;
    let formulasCount = 0;
    
    items.forEach(item => {
        const matchesSearch = item.name.toLowerCase().includes(searchQuery) || 
                              item.shortDesc.toLowerCase().includes(searchQuery) || 
                              item.desc.toLowerCase().includes(searchQuery);
        const matchesCategory = currentFilter === 'all' || item.cat === currentFilter;
        
        if (matchesSearch && matchesCategory) {
            totalCount++;
            
            let renderItemInfo = item;
            if (currentLang !== 'en' && currentExcelLang === 'en') {
                const enItem = locales['en'].items.find(i => i.id === item.id);
                if (enItem) {
                    renderItemInfo = { ...item, name: enItem.name, syntax: enItem.syntax };
                }
            }

            const template = createAccordionTemplate(renderItemInfo);
            
            if (item.type === 'shortcut') {
                _shortcutsContainer.innerHTML += template;
                shortcutsCount++;
            } else {
                _formulasContainer.innerHTML += template;
                formulasCount++;
            }
        }
    });

    if (shortcutsCount === 0) _shortcutsWrapper.style.display = 'none';
    if (formulasCount === 0) _formulasWrapper.style.display = 'none';

    const statsTemplate = locales[currentLang].resultsText;
    _resultsStats.textContent = statsTemplate.replace('{count}', totalCount);

    attachAccordionEvents();
}

function createAccordionTemplate(item) {
    const detailsLabel = locales[currentLang].formulasTitle ? (currentLang === 'es' ? 'Detalles:' : currentLang === 'pt' ? 'Detalhes:' : currentLang === 'id' ? 'Detail:' : 'Details:') : 'Details:';
    const syntaxLabel = item.type === 'formula' ? 
        (currentLang === 'es' ? 'Fórmula / Sintaxis:' : currentLang === 'pt' ? 'Fórmula / Sintaxe:' : currentLang === 'id' ? 'Rumus / Sintaks:' : 'Formula / Syntax:') : 
        (currentLang === 'es' ? 'Teclas:' : currentLang === 'pt' ? 'Teclas:' : currentLang === 'id' ? 'Tombol:' : 'Keys:');
    
    return `
        <div class="accordion-item">
            <button class="accordion-header" aria-expanded="false">
                <div class="accordion-title-wrapper">
                    <span class="btn-tag">${item.name}</span>
                    <span class="short-desc">${item.shortDesc}</span>
                </div>
                <svg class="accordion-icon" viewBox="0 0 24 24">
                    <polyline points="6 9 12 15 18 9"></polyline>
                </svg>
            </button>
            <div class="accordion-content">
                <div class="accordion-body">
                    <p><strong>${detailsLabel}</strong> ${item.desc}</p>
                    <div class="syntax-box">
                        <p><strong>${syntaxLabel}</strong> <br><code>${item.syntax}</code></p>
                    </div>
                </div>
            </div>
        </div>
    `;
}

function attachAccordionEvents() {
    const accordionHeaders = document.querySelectorAll('.accordion-header');

    accordionHeaders.forEach(header => {
        header.addEventListener('click', function() {
            const item = this.parentElement;
            const content = item.querySelector('.accordion-content');
            const isActive = item.classList.contains('active');

            if (!isActive) {
                item.classList.add('active');
                this.setAttribute('aria-expanded', 'true');
                content.style.maxHeight = content.scrollHeight + "px";
            } else {
                item.classList.remove('active');
                this.setAttribute('aria-expanded', 'false');
                content.style.maxHeight = null;
            }
        });
    });
}

document.addEventListener('DOMContentLoaded', init);
