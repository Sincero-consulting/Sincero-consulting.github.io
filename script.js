const locales = {
    en: {
        title: "Excel Cheat Sheet",
        subtitle: "Master shortcuts, formulas, and data analysis",
        searchPlaceholder: "Search shortcuts and formulas...",
        footerText: "All rights reserved.",
        resultsText: "Showing {count} results",
        shortcutsTitle: "Keyboard Shortcuts",
        formulasTitle: "Essential Formulas",
        filters: {
            all: "All",
            mov: "Movement",
            form: "Formatting",
            gest: "Management",
            bas: "Basics",
            limp: "Cleaning",
            log: "Logic",
            avz: "Advanced"
        },
        items: [
            // Movement
            { id: "s1", type: "shortcut", cat: "mov", name: "Jump to Edge", shortDesc: "Move selection to data region edge", syntax: "Windows: Ctrl + Arrows | Mac: Cmd + Arrows", desc: "Mastering these commands allows you to work without touching the mouse, which is the mark of a true expert." },
            { id: "s2", type: "shortcut", cat: "mov", name: "Select to End", shortDesc: "Select to the edge of data", syntax: "Windows: Ctrl + Shift + Arrows | Mac: Cmd + Shift + Arrows", desc: "Instantly highlights all data in a row or column up to the first empty cell." },
            { id: "s3", type: "shortcut", cat: "mov", name: "Select Table", shortDesc: "Select the current data block", syntax: "Windows: Ctrl + E | Mac: Cmd + A", desc: "Selects the entire contiguous data range (the current region)." },
            { id: "s4", type: "shortcut", cat: "mov", name: "Switch Tabs", shortDesc: "Move between worksheet tabs", syntax: "Windows: Ctrl + PgUp/PgDn | Mac: Option + Left/Right Arrow", desc: "Quickly navigate through different sheets in your workbook." },
            { id: "s5", type: "shortcut", cat: "mov", name: "Extend Selection", shortDesc: "Extend selection by one cell", syntax: "Windows: Shift + Arrows | Mac: Shift + Arrows", desc: "Fine-tune your highlighted area one cell at a time." },
            // Formatting
            { id: "s6", type: "shortcut", cat: "form", name: "Currency Format", shortDesc: "Apply currency format", syntax: "Windows: Ctrl + Shift + $ | Mac: Control + Shift + $", desc: "Instantly formats selected numbers as currency with two decimal places." },
            { id: "s7", type: "shortcut", cat: "form", name: "Percent Format", shortDesc: "Apply percentage format", syntax: "Windows: Ctrl + Shift + % | Mac: Control + Shift + %", desc: "Converts decimals to neat percentages instantly." },
            { id: "s8", type: "shortcut", cat: "form", name: "Bold", shortDesc: "Toggle bold text", syntax: "Windows: Ctrl + B | Mac: Cmd + B", desc: "Make your headers or important data stand out." },
            { id: "s9", type: "shortcut", cat: "form", name: "Format Cells", shortDesc: "Open Format Cells dialog", syntax: "Windows: Ctrl + 1 | Mac: Cmd + 1", desc: "The master menu for all cell visual and data formatting." },
            { id: "s10", type: "shortcut", cat: "form", name: "AutoSum", shortDesc: "Insert AutoSum formula", syntax: "Windows: Alt + = | Mac: Cmd + Shift + T", desc: "Automatically writes a SUM formula for adjacent numbers." },
            // Management
            { id: "s11", type: "shortcut", cat: "gest", name: "Toggle Filters", shortDesc: "Turn filters on or off", syntax: "Windows: Ctrl + Shift + L | Mac: Cmd + Shift + F", desc: "Instantly add or remove dropdown filters to your headers." },
            { id: "s12", type: "shortcut", cat: "gest", name: "Create Table", shortDesc: "Create an official Excel Table", syntax: "Windows: Ctrl + T | Mac: Cmd + T", desc: "Converts raw data into an official Table with auto-formatting and dynamic ranges." },
            { id: "s13", type: "shortcut", cat: "gest", name: "Paste Special", shortDesc: "Open Paste Special menu", syntax: "Windows: Ctrl + Alt + V | Mac: Cmd + Control + V", desc: "Allows pasting values only, formats only, or transposing data." },
            { id: "s14", type: "shortcut", cat: "gest", name: "Clear Content", shortDesc: "Delete all cell contents", syntax: "Windows: Delete | Mac: Fn + Delete", desc: "Removes text and formulas without deleting the cell formatting." },
            // Basics
            { id: "f1", type: "formula", cat: "bas", name: "SUM", shortDesc: "Total sum", syntax: "=SUM(range)", desc: "Calculates the total sum of the selected range. Ideal for quick descriptive analysis." },
            { id: "f2", type: "formula", cat: "bas", name: "AVERAGE", shortDesc: "Arithmetic mean", syntax: "=AVERAGE(range)", desc: "Calculates the arithmetic mean of the selected numbers." },
            { id: "f3", type: "formula", cat: "bas", name: "COUNT", shortDesc: "Count numbers", syntax: "=COUNT(range)", desc: "Counts the number of cells in a range that contain numbers." },
            { id: "f4", type: "formula", cat: "bas", name: "COUNTA", shortDesc: "Count non-empty", syntax: "=COUNTA(range)", desc: "Counts all cells that are not empty (useful for text and mixed data)." },
            { id: "f5", type: "formula", cat: "bas", name: "MAX / MIN", shortDesc: "Extreme values", syntax: "=MAX(range) / =MIN(range)", desc: "Returns the maximum or minimum value in the specified range." },
            { id: "f6", type: "formula", cat: "bas", name: "MODE.SNGL", shortDesc: "Most frequent value", syntax: "=MODE.SNGL(range)", desc: "Returns the most frequently occurring, or repetitive, value in an array or range of data." },
            // Cleaning
            { id: "f7", type: "formula", cat: "limp", name: "TRIM", shortDesc: "Remove extra spaces", syntax: "=TRIM(text)", desc: "Removes all spaces from text except for single spaces between words. Essential for cleaning imported data." },
            { id: "f8", type: "formula", cat: "limp", name: "PROPER", shortDesc: "Capitalize first letters", syntax: "=PROPER(text)", desc: "Capitalizes the first letter of each word (ideal for formatting names consistently)." },
            { id: "f9", type: "formula", cat: "limp", name: "UPPER / LOWER", shortDesc: "Convert text case", syntax: "=UPPER(text) / =LOWER(text)", desc: "Converts text to all uppercase or all lowercase letters." },
            { id: "f10", type: "formula", cat: "limp", name: "VALUE", shortDesc: "Text to number", syntax: "=VALUE(text)", desc: "Converts a text string that represents a number to a real number so it can be used in math equations." },
            { id: "f11", type: "formula", cat: "limp", name: "TEXTJOIN", shortDesc: "Combine text modernly", syntax: "=TEXTJOIN(delimiter, ignore_empty, range)", desc: "The modern and superior way to concatenate strings, allowing you to specify a delimiter and ignore empty cells." },
            // Logic
            { id: "f12", type: "formula", cat: "log", name: "IF", shortDesc: "Conditional logic", syntax: "=IF(logical_test, value_if_true, value_if_false)", desc: "The queen of functions. Performs a logical test and returns one value for a TRUE result, and another for a FALSE result." },
            { id: "f13", type: "formula", cat: "log", name: "IFERROR", shortDesc: "Handle errors smoothly", syntax: "=IFERROR(value, value_if_error)", desc: "Returns a custom value if a formula evaluates to an error, preventing ugly #N/A marks in your polished reports." },
            { id: "f14", type: "formula", cat: "log", name: "COUNTIF", shortDesc: "Conditional count", syntax: "=COUNTIF(range, criteria)", desc: "Counts the number of cells within a range that meet a single specific condition." },
            { id: "f15", type: "formula", cat: "log", name: "SUMIFS", shortDesc: "Multi-conditional sum", syntax: "=SUMIFS(sum_range, criteria_range1, criteria1, ...)", desc: "Adds cells based on multiple filters or criteria. Indispensable for dashboards." },
            // Advanced
            { id: "f16", type: "formula", cat: "avz", name: "XLOOKUP", shortDesc: "Modern lookup", syntax: "=XLOOKUP(lookup_val, lookup_array, return_array)", desc: "The modern successor to VLOOKUP. Safer, faster, searches in any direction, and handles missing data natively." },
            { id: "f17", type: "formula", cat: "avz", name: "INDEX & MATCH", shortDesc: "2D Lookup power couple", syntax: "=INDEX(range, MATCH(...))", desc: "The power couple for complex, two-dimensional lookups when XLOOKUP isn't available or suitable." },
            { id: "f18", type: "formula", cat: "avz", name: "FILTER", shortDesc: "Dynamic filtering", syntax: "=FILTER(range, criteria)", desc: "Creates a dynamic array that automatically updates and spills results based on your defined filters." },
            { id: "f19", type: "formula", cat: "avz", name: "UNIQUE", shortDesc: "Extract unique values", syntax: "=UNIQUE(range)", desc: "Instantly extracts a list of unique values from a range without needing manual duplication removal." },
            { id: "f20", type: "formula", cat: "avz", name: "LET", shortDesc: "Define variables", syntax: "=LET(name, value, calculation)", desc: "Allows assigning names to calculation results inside a formula to improve performance, readability, and logic flow." }
        ]
    },
    es: {
        title: "Hoja de Referencia de Excel",
        subtitle: "Domina atajos, fórmulas y análisis de datos",
        searchPlaceholder: "Buscar atajos y fórmulas...",
        footerText: "Todos los derechos reservados.",
        resultsText: "Mostrando {count} resultados",
        shortcutsTitle: "Atajos de Teclado",
        formulasTitle: "Fórmulas Imprescindibles",
        filters: {
            all: "Todo",
            mov: "Movimiento",
            form: "Formateo",
            gest: "Gestión",
            bas: "Básicas",
            limp: "Limpieza",
            log: "Lógicas",
            avz: "Avanzadas"
        },
        items: [
            // Movement
            { id: "s1", type: "shortcut", cat: "mov", name: "Ir al borde de la región", shortDesc: "Mueve la selección al borde de datos", syntax: "Windows: Ctrl + Flechas | Mac: Cmd + Flechas", desc: "Dominar estos comandos te permitirá trabajar sin tocar el ratón, lo cual es la marca de un verdadero experto." },
            { id: "s2", type: "shortcut", cat: "mov", name: "Seleccionar hasta el final", shortDesc: "Selecciona hasta el final de los datos", syntax: "Windows: Ctrl + Shift + Flechas | Mac: Cmd + Shift + Flechas", desc: "Resalta instantáneamente todos los datos en una fila o columna hasta la primera celda vacía." },
            { id: "s3", type: "shortcut", cat: "mov", name: "Seleccionar toda la tabla", shortDesc: "Selecciona la tabla actual", syntax: "Windows: Ctrl + E | Mac: Cmd + A", desc: "Selecciona todo el rango de datos contiguos (la región actual)." },
            { id: "s4", type: "shortcut", cat: "mov", name: "Cambiar entra pestañas", shortDesc: "Mueve entre pestañas del libro", syntax: "Windows: Ctrl + AvPág / RePág | Mac: Option + Flecha Der/Izq", desc: "Navega rápidamente a través de diferentes hojas en tu libro." },
            { id: "s5", type: "shortcut", cat: "mov", name: "Extender selección", shortDesc: "Extiende la selección una celda", syntax: "Windows: Shift + Flechas | Mac: Shift + Flechas", desc: "Ajusta tu área seleccionada celda por celda." },
            // Formatting
            { id: "s6", type: "shortcut", cat: "form", name: "Formato Moneda", shortDesc: "Aplica formato Moneda", syntax: "Windows: Ctrl + Shift + $ | Mac: Control + Shift + $", desc: "Aplica instantáneamente formato de moneda con dos decimales a los números seleccionados." },
            { id: "s7", type: "shortcut", cat: "form", name: "Formato Porcentaje", shortDesc: "Aplica formato Porcentaje", syntax: "Windows: Ctrl + Shift + % | Mac: Control + Shift + %", desc: "Convierte decimales a porcentajes limpios al instante." },
            { id: "s8", type: "shortcut", cat: "form", name: "Negrita", shortDesc: "Poner/Quitar Negrita", syntax: "Windows: Ctrl + N | Mac: Cmd + B", desc: "Haz que tus encabezados o datos importantes destaquen." },
            { id: "s9", type: "shortcut", cat: "form", name: "Menú Formato", shortDesc: "Abrir menú de Formato de Celda", syntax: "Windows: Ctrl + 1 | Mac: Cmd + 1", desc: "El menú maestro para todo el formato visual y de datos de las celdas." },
            { id: "s10", type: "shortcut", cat: "form", name: "AutoSuma", shortDesc: "Insertar AutoSuma", syntax: "Windows: Alt + = | Mac: Cmd + Shift + T", desc: "Escribe automáticamente una fórmula SUMA para los números adyacentes." },
            // Management
            { id: "s11", type: "shortcut", cat: "gest", name: "Filtros", shortDesc: "Activar/Desactivar Filtros", syntax: "Windows: Ctrl + Shift + L | Mac: Cmd + Shift + F", desc: "Añade o quita instantáneamente filtros desplegables a tus encabezados." },
            { id: "s12", type: "shortcut", cat: "gest", name: "Crear Tabla", shortDesc: "Crear una Tabla oficial", syntax: "Windows: Ctrl + T | Mac: Cmd + T", desc: "Convierte datos sin procesar en una Tabla oficial con formato automático y rangos dinámicos." },
            { id: "s13", type: "shortcut", cat: "gest", name: "Pegado Especial", shortDesc: "Abrir Pegado Especial", syntax: "Windows: Ctrl + Alt + V | Mac: Cmd + Control + V", desc: "Permite pegar solo valores, solo formatos o transponer datos." },
            { id: "s14", type: "shortcut", cat: "gest", name: "Borrar todo", shortDesc: "Borrar todo el contenido", syntax: "Windows: Suprimir | Mac: Fn + Delete", desc: "Elimina texto y fórmulas sin borrar el formato de la celda." },
            // Basics
            { id: "f1", type: "formula", cat: "bas", name: "SUMA", shortDesc: "Suma total", syntax: "=SUMA(rango)", desc: "Suma total. Ideales para el análisis descriptivo rápido." },
            { id: "f2", type: "formula", cat: "bas", name: "PROMEDIO", shortDesc: "Media aritmética", syntax: "=PROMEDIO(rango)", desc: "Lleva a cabo la media aritmética de los valores en el rango." },
            { id: "f3", type: "formula", cat: "bas", name: "CONTAR", shortDesc: "Cuenta números", syntax: "=CONTAR(rango)", desc: "Cuenta el número de celdas en un rango que contienen números." },
            { id: "f4", type: "formula", cat: "bas", name: "CONTARA", shortDesc: "Cuenta no vacías", syntax: "=CONTARA(rango)", desc: "Cuenta todas las celdas que no están vacías (muy útil para contar texto)." },
            { id: "f5", type: "formula", cat: "bas", name: "MAX / MIN", shortDesc: "Valores extremos", syntax: "=MAX(rango) / =MIN(rango)", desc: "Devuelve el valor máximo o mínimo en un conjunto de valores o rango." },
            { id: "f6", type: "formula", cat: "bas", name: "MODA.UNO", shortDesc: "Valor más repetido", syntax: "=MODA.UNO(rango)", desc: "Devuelve el valor que ocurre con mayor frecuencia o que más se repite en una matriz o rango de datos." },
            // Cleaning
            { id: "f7", type: "formula", cat: "limp", name: "ESPACIOS", shortDesc: "Elimina espacios dobles", syntax: "=ESPACIOS(texto)", desc: "Elimina espacios dobles o espacios al inicio y al final. Esencial antes de procesar datos." },
            { id: "f8", type: "formula", cat: "limp", name: "NOMPROPIO", shortDesc: "Capitaliza palabras", syntax: "=NOMPROPIO(texto)", desc: "Pone la primera letra en mayúscula de cada palabra (ideal para limpiar nombres)." },
            { id: "f9", type: "formula", cat: "limp", name: "MAYUS / MINUS", shortDesc: "Convierte mayús/minús", syntax: "=MAYUS(texto) / =MINUS(texto)", desc: "Convierte todo el texto a letras mayúsculas o minúsculas." },
            { id: "f10", type: "formula", cat: "limp", name: "VALOR", shortDesc: "Texto a número", syntax: "=VALOR(texto)", desc: "Convierte un número guardado como texto en un número real que Excel puede calcular." },
            { id: "f11", type: "formula", cat: "limp", name: "UNIRCADENAS", shortDesc: "Concatenar moderno", syntax: "=UNIRCADENAS(delimitador, ignorar_vacios, rango)", desc: "La forma moderna y superior de concatenar textos, permitiendo ignorar vacíos e incluir un delimitador constante." },
            // Logic
            { id: "f12", type: "formula", cat: "log", name: "SI", shortDesc: "Condición lógica", syntax: "=SI(prueba_lógica, valor_si_verdadero, valor_si_falso)", desc: "La función reina de Excel. Realiza una prueba lógica y devuelve un valor si es verdadero, y otro si es falso." },
            { id: "f13", type: "formula", cat: "log", name: "SI.ERROR", shortDesc: "Manejo de errores", syntax: "=SI.ERROR(fórmula, valor_si_error)", desc: "Para que tus reportes no muestren el feo #N/A o #¡DIV/0!. Reemplaza el error con texto personalizado o un cero." },
            { id: "f14", type: "formula", cat: "log", name: "CONTAR.SI", shortDesc: "Conteo condicional", syntax: "=CONTAR.SI(rango, criterio)", desc: "Cuenta solo si las celdas cumplen con una condición o criterio específico." },
            { id: "f15", type: "formula", cat: "log", name: "SUMAR.SI.CONJUNTO", shortDesc: "Suma multi-condicional", syntax: "=SUMAR.SI.CONJUNTO(rango_suma, rango_criterio1, ...)", desc: "Suma valores basándose en múltiples filtros y criterios a la vez. Fundamental para análisis de datos avanzado." },
            // Advanced
            { id: "f16", type: "formula", cat: "avz", name: "BUSCARX", shortDesc: "Búsqueda moderna", syntax: "=BUSCARX(valor_buscado, matriz_buscada, matriz_devuelta)", desc: "El sucesor moderno de BUSCARV. Es más seguro, rápido, busca en cualquier dirección y maneja errores por defecto." },
            { id: "f17", type: "formula", cat: "avz", name: "INDICE & COINCIDIR", shortDesc: "Poder de búsqueda 2D", syntax: "=INDICE(rango, COINCIDIR(...))", desc: "La pareja de poder para búsquedas complejas bidimensionales, especialmente cuando XLOOKUP no está disponible o requiere mapeos complejos." },
            { id: "f18", type: "formula", cat: "avz", name: "FILTRAR", shortDesc: "Filtrado dinámico", syntax: "=FILTRAR(rango, criterios)", desc: "Crea automáticamente una tabla dinámica extraída de otra más grande, y se actualiza sola según cambien los datos o tus filtros." },
            { id: "f19", type: "formula", cat: "avz", name: "UNICOS", shortDesc: "Extraer valores únicos", syntax: "=UNICOS(rango)", desc: "Extrae de forma instántanea una matriz con una lista de valores sin duplicados. Excelente para crear listas desplegables." },
            { id: "f20", type: "formula", cat: "avz", name: "LET", shortDesc: "Definir variables", syntax: "=LET(nombre1, valor1, cálculo)", desc: "Permite definir y asignar variables dentro de una misma fórmula para evitar calcular el mismo dato varias veces y mejorar el rendimiento de la hoja." }
        ]
    }
};

// Global State
let currentLang = 'en'; // Default UI language ('en' or 'es')
let currentExcelLang = 'en'; // Default Excel syntax language ('en' or 'es')
let currentFilter = 'all'; // Default category layout
let searchQuery = '';

// DOM Elements
const _title = document.getElementById('page-title');
const _subtitle = document.getElementById('page-subtitle');
const _searchInput = document.getElementById('search-input');
const _langToggle = document.getElementById('lang-toggle');
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

// Initialize application
function init() {
    // Check local storage for preferences
    const savedTheme = localStorage.getItem('theme') || 'light';
    const savedLang = localStorage.getItem('lang') || 'en'; // Changed initial preference to EN or default storage
    
    // Set theme
    document.documentElement.setAttribute('data-theme', savedTheme);
    updateThemeIcon(savedTheme);

    // Set language
    setLanguage(savedLang);

    // Event Listeners
    _langToggle.addEventListener('click', () => {
        setLanguage(currentLang === 'en' ? 'es' : 'en');
    });

    _excelLangToggle.addEventListener('click', () => {
        currentExcelLang = currentExcelLang === 'en' ? 'es' : 'en';
        updateExcelLangButton();
        renderItems();
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

    // FAB Listeners
    _fabTop.addEventListener('click', () => {
        window.scrollTo({ top: 0, behavior: 'smooth' });
    });

    _fabFilter.addEventListener('click', (e) => {
        e.stopPropagation();
        _fabFilterMenu.classList.toggle('active');
    });

    // Close menu if clicked outside
    document.addEventListener('click', (e) => {
        if (!e.target.closest('.fab-container')) {
            _fabFilterMenu.classList.remove('active');
        }
    });

    // Scroll listener for FAB visibility
    window.addEventListener('scroll', () => {
        const yPos = window.scrollY;
        // Lowered threshold to 150 so it shows up sooner
        if (yPos > 150) {
            _fabTop.classList.add('visible');
            _fabFilter.classList.add('visible');
        } else {
            _fabTop.classList.remove('visible');
            _fabFilter.classList.remove('visible');
        }
    });

    // Set Copyright Year
    document.getElementById('year').textContent = new Date().getFullYear();
}

function updateThemeIcon(theme) {
    _themeIcon.textContent = theme === 'dark' ? '☀️' : '🌙';
}

function setLanguage(langCode) {
    currentLang = langCode;
    // Default Excel syntax behavior: if UI is EN, Excel is EN. If UI is ES, we set Excel to ES initially, but allow toggle.
    currentExcelLang = langCode === 'en' ? 'en' : 'es'; 
    localStorage.setItem('lang', langCode);
    
    const textData = locales[langCode];

    // Update static HTML texts
    _title.textContent = textData.title;
    _subtitle.textContent = textData.subtitle;
    _searchInput.placeholder = textData.searchPlaceholder;
    _footerText.textContent = textData.footerText;
    _shortcutsTitle.textContent = textData.shortcutsTitle;
    _formulasTitle.textContent = textData.formulasTitle;
    
    // Update Button Text to show OPPOSITE language 
    _langToggle.querySelector('.lang-text').textContent = langCode === 'en' ? 'ES' : 'EN';
    
    updateExcelLangButton();

    // Maintain filter state if possible, otherwise reset to 'all'
    if (!textData.filters[currentFilter]) {
        currentFilter = 'all';
    }

    renderFilters();
    renderItems();
}

function updateExcelLangButton() {
    if (currentLang === 'es') {
        _excelLangToggle.style.display = 'flex';
        _excelLangText.textContent = currentExcelLang === 'es' ? 'Versión INGLÉS' : 'Versión ESPAÑOL';
    } else {
        _excelLangToggle.style.display = 'none';
        currentExcelLang = 'en'; // Force English syntax if UI is English
    }
}

function renderFilters() {
    _filtersContainer.innerHTML = '';
    _fabFilterMenu.innerHTML = '';
    const filters = locales[currentLang].filters;
    
    for (const [key, name] of Object.entries(filters)) {
        // Build for main filter bar
        const btnMain = document.createElement('button');
        btnMain.className = `filter-btn ${key === currentFilter ? 'active' : ''}`;
        btnMain.textContent = name;
        btnMain.dataset.filter = key;
        
        // Build for FAB floating menu
        const btnFab = document.createElement('button');
        btnFab.className = `filter-btn ${key === currentFilter ? 'active' : ''}`;
        btnFab.textContent = name;
        btnFab.dataset.filter = key;

        const clickHandler = () => {
            currentFilter = key;
            // Update active styling across both containers
            document.querySelectorAll('.filter-btn').forEach(b => b.classList.remove('active'));
            document.querySelectorAll(`[data-filter="${key}"]`).forEach(b => b.classList.add('active'));
            
            _fabFilterMenu.classList.remove('active'); // Close menu after selection
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
    
    // Reset display
    _shortcutsWrapper.style.display = 'block';
    _formulasWrapper.style.display = 'block';

    const items = locales[currentLang].items;
    let totalCount = 0;
    let shortcutsCount = 0;
    let formulasCount = 0;
    
    // Filter logic
    items.forEach(item => {
        const matchesSearch = item.name.toLowerCase().includes(searchQuery) || 
                              item.shortDesc.toLowerCase().includes(searchQuery) || 
                              item.desc.toLowerCase().includes(searchQuery);
        const matchesCategory = currentFilter === 'all' || item.cat === currentFilter;
        
        if (matchesSearch && matchesCategory) {
            totalCount++;
            
            // Sub-language toggle logic: If UI is ES but Excel Version is EN
            let renderItemInfo = item;
            if (currentLang === 'es' && currentExcelLang === 'en') {
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

    // Hide sections if empty
    if (shortcutsCount === 0) _shortcutsWrapper.style.display = 'none';
    if (formulasCount === 0) _formulasWrapper.style.display = 'none';

    // Update Stats
    const statsTemplate = locales[currentLang].resultsText;
    _resultsStats.textContent = statsTemplate.replace('{count}', totalCount);

    // Reattach Accordion Events dynamically generated elements
    attachAccordionEvents();
}

function createAccordionTemplate(item) {
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
                    <p><strong>${currentLang === 'en' ? 'Details:' : 'Detalles:'}</strong> ${item.desc}</p>
                    <div class="syntax-box">
                        <p><strong>${item.type === 'formula' ? (currentLang === 'en' ? 'Formula / Syntax:' : 'Fórmula / Sintaxis:') : (currentLang === 'en' ? 'Keys:' : 'Teclas:')}</strong> <br><code>${item.syntax}</code></p>
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

// Start sequence
document.addEventListener('DOMContentLoaded', init);
