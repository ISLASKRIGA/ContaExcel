// ===== State Management =====
let currentPdfData = null;
let currentFileName = '';
let extractedTables = [];

// ===== Table Styles =====
const tableStyles = {
    modern: {
        headerBg: 'FF667EEA',
        headerFont: 'FFFFFF',
        rowBg1: 'FFF8F9FF',
        rowBg2: 'FFFFFFFF',
        borderColor: 'FFE2E8F0'
    },
    classic: {
        headerBg: 'FF2C3E50',
        headerFont: 'FFFFFF',
        rowBg1: 'FFECF0F1',
        rowBg2: 'FFFFFFFF',
        borderColor: 'FFBDC3C7'
    },
    minimal: {
        headerBg: 'FF000000',
        headerFont: 'FFFFFF',
        rowBg1: 'FFFFFFFF',
        rowBg2: 'FFFFFFFF',
        borderColor: 'FFE0E0E0'
    },
    colorful: {
        headerBg: 'FFFF6B6B',
        headerFont: 'FFFFFF',
        rowBg1: 'FFFFE66D',
        rowBg2: 'FF4ECDC4',
        borderColor: 'FF95E1D3'
    }
};

// ===== DOM Elements =====
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const removeFile = document.getElementById('removeFile');
const previewArea = document.getElementById('previewArea');
const previewContent = document.getElementById('previewContent');
const sheetSelect = document.getElementById('sheetSelect');
const customizationOptions = document.getElementById('customizationOptions');
const convertButton = document.getElementById('convertButton');
const loadingState = document.getElementById('loadingState');

// ===== Event Listeners =====
uploadArea.addEventListener('click', () => fileInput.click());
fileInput.addEventListener('change', handleFileSelect);
removeFile.addEventListener('click', resetUpload);
sheetSelect.addEventListener('change', handleTableChange);
convertButton.addEventListener('click', generateExcel);

// Drag and Drop
uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('drag-over');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

// ===== File Handling =====
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

async function handleFile(file) {
    // Validate file type
    if (file.type !== 'application/pdf' && !file.name.match(/\.pdf$/i)) {
        alert('Por favor, selecciona un archivo PDF válido');
        return;
    }

    currentFileName = file.name;

    // Update UI
    fileName.textContent = file.name;
    fileSize.textContent = formatFileSize(file.size);
    uploadArea.style.display = 'none';
    fileInfo.style.display = 'flex';

    // Show loading
    loadingState.style.display = 'block';

    try {
        // Read PDF file
        const arrayBuffer = await file.arrayBuffer();
        await extractPdfData(arrayBuffer);

        loadingState.style.display = 'none';

        if (extractedTables.length === 0) {
            alert('No se encontraron tablas en el PDF. Intenta con otro archivo.');
            resetUpload();
            return;
        }

        displayTables();
        displayPreview(0);
        previewArea.style.display = 'block';
        customizationOptions.style.display = 'block';
        convertButton.style.display = 'flex';

    } catch (error) {
        console.error('Error processing PDF:', error);
        alert('Error al procesar el PDF. Por favor, intenta con otro archivo.');
        loadingState.style.display = 'none';
        resetUpload();
    }
}

function formatFileSize(bytes) {
    if (bytes === 0) return '0 Bytes';
    const k = 1024;
    const sizes = ['Bytes', 'KB', 'MB', 'GB'];
    const i = Math.floor(Math.log(bytes) / Math.log(k));
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

// ===== PDF Processing =====
async function extractPdfData(arrayBuffer) {
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
    extractedTables = [];

    // Process each page
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
        const page = await pdf.getPage(pageNum);
        const textContent = await page.getTextContent();

        // Extract text items with positions
        const items = textContent.items.map(item => ({
            text: item.str,
            x: item.transform[4],
            y: item.transform[5],
            height: item.height
        }));

        // Group items into rows based on Y position
        const rows = groupIntoRows(items);

        if (rows.length > 0) {
            extractedTables.push({
                name: `Página ${pageNum}`,
                data: rows
            });
        }
    }
}

function groupIntoRows(items) {
    if (items.length === 0) return [];

    // Sort by Y position (descending, as PDF coordinates start from bottom)
    items.sort((a, b) => b.y - a.y);

    const rows = [];
    let currentRow = [];
    let currentY = items[0].y;
    const yThreshold = 5; // Tolerance for same row

    items.forEach(item => {
        if (Math.abs(item.y - currentY) < yThreshold) {
            currentRow.push(item);
        } else {
            if (currentRow.length > 0) {
                // Sort row items by X position
                currentRow.sort((a, b) => a.x - b.x);
                rows.push(currentRow.map(i => i.text));
            }
            currentRow = [item];
            currentY = item.y;
        }
    });

    // Add last row
    if (currentRow.length > 0) {
        currentRow.sort((a, b) => a.x - b.x);
        rows.push(currentRow.map(i => i.text));
    }

    return rows;
}

function displayTables() {
    sheetSelect.innerHTML = '';
    extractedTables.forEach((table, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = table.name;
        sheetSelect.appendChild(option);
    });
}

function handleTableChange() {
    const selectedIndex = parseInt(sheetSelect.value);
    displayPreview(selectedIndex);
}

function displayPreview(tableIndex) {
    const tableData = extractedTables[tableIndex].data;

    if (!tableData || tableData.length === 0) {
        previewContent.innerHTML = '<p style="padding: 2rem; text-align: center; color: #718096;">No hay datos para mostrar</p>';
        return;
    }

    let html = '<table>';

    // Limit preview to first 15 rows
    const previewRows = tableData.slice(0, 15);

    previewRows.forEach((row, index) => {
        html += '<tr>';
        row.forEach(cell => {
            const tag = index === 0 ? 'th' : 'td';
            const value = cell !== undefined && cell !== null ? cell : '';
            html += `<${tag}>${value}</${tag}>`;
        });
        html += '</tr>';
    });

    html += '</table>';

    if (tableData.length > 15) {
        html += `<p style="padding: 1rem; text-align: center; color: #718096; font-size: 0.875rem;">Mostrando 15 de ${tableData.length} filas</p>`;
    }

    previewContent.innerHTML = html;
}

function resetUpload() {
    currentPdfData = null;
    currentFileName = '';
    extractedTables = [];
    fileInput.value = '';

    uploadArea.style.display = 'block';
    fileInfo.style.display = 'none';
    previewArea.style.display = 'none';
    customizationOptions.style.display = 'none';
    convertButton.style.display = 'none';
}

// ===== Excel Generation =====
async function generateExcel() {
    if (extractedTables.length === 0) {
        alert('No hay datos para convertir');
        return;
    }

    // Show loading state
    convertButton.style.display = 'none';
    loadingState.style.display = 'block';

    try {
        // Get customization options
        const sheetName = document.getElementById('sheetName').value || 'Estado de Cuenta';
        const style = document.getElementById('styleSelect').value;
        const autoWidth = document.getElementById('autoWidth').checked;
        const freezeHeader = document.getElementById('freezeHeader').checked;

        // Create workbook
        const wb = XLSX.utils.book_new();

        // Get selected table or use all tables
        const selectedIndex = parseInt(sheetSelect.value);
        const tablesToExport = extractedTables;

        tablesToExport.forEach((table, index) => {
            const wsName = tablesToExport.length === 1 ? sheetName : `${sheetName} ${index + 1}`;
            const ws = XLSX.utils.aoa_to_sheet(table.data);

            // Apply styling
            applyExcelStyle(ws, table.data, style);

            // Auto-width columns
            if (autoWidth) {
                const colWidths = calculateColumnWidths(table.data);
                ws['!cols'] = colWidths;
            }

            // Freeze header row
            if (freezeHeader && table.data.length > 0) {
                ws['!freeze'] = { xSplit: 0, ySplit: 1 };
            }

            XLSX.utils.book_append_sheet(wb, ws, wsName);
        });

        // Generate Excel file
        const excelFileName = currentFileName.replace(/\.pdf$/i, '.xlsx');
        XLSX.writeFile(wb, excelFileName);

        // Show success message
        setTimeout(() => {
            loadingState.style.display = 'none';
            convertButton.style.display = 'flex';
            showSuccessMessage();
        }, 500);

    } catch (error) {
        console.error('Error generating Excel:', error);
        alert('Hubo un error al generar el Excel. Por favor, intenta nuevamente.');
        loadingState.style.display = 'none';
        convertButton.style.display = 'flex';
    }
}

function applyExcelStyle(ws, data, styleName) {
    const style = tableStyles[styleName];
    const range = XLSX.utils.decode_range(ws['!ref']);

    for (let R = range.s.r; R <= range.e.r; ++R) {
        for (let C = range.s.c; C <= range.e.c; ++C) {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            if (!ws[cellAddress]) continue;

            // Initialize cell style
            if (!ws[cellAddress].s) ws[cellAddress].s = {};

            // Header row styling
            if (R === 0) {
                ws[cellAddress].s = {
                    fill: { fgColor: { rgb: style.headerBg } },
                    font: { color: { rgb: style.headerFont }, bold: true },
                    alignment: { horizontal: 'center', vertical: 'center' },
                    border: {
                        top: { style: 'thin', color: { rgb: style.borderColor } },
                        bottom: { style: 'thin', color: { rgb: style.borderColor } },
                        left: { style: 'thin', color: { rgb: style.borderColor } },
                        right: { style: 'thin', color: { rgb: style.borderColor } }
                    }
                };
            } else {
                // Alternating row colors
                const bgColor = R % 2 === 0 ? style.rowBg1 : style.rowBg2;
                ws[cellAddress].s = {
                    fill: { fgColor: { rgb: bgColor } },
                    border: {
                        top: { style: 'thin', color: { rgb: style.borderColor } },
                        bottom: { style: 'thin', color: { rgb: style.borderColor } },
                        left: { style: 'thin', color: { rgb: style.borderColor } },
                        right: { style: 'thin', color: { rgb: style.borderColor } }
                    }
                };
            }
        }
    }
}

function calculateColumnWidths(data) {
    if (data.length === 0) return [];

    const maxCols = Math.max(...data.map(row => row.length));
    const widths = [];

    for (let col = 0; col < maxCols; col++) {
        let maxWidth = 10;
        data.forEach(row => {
            if (row[col]) {
                const cellLength = String(row[col]).length;
                maxWidth = Math.max(maxWidth, cellLength);
            }
        });
        widths.push({ wch: Math.min(maxWidth + 2, 50) });
    }

    return widths;
}

function showSuccessMessage() {
    // Create success notification
    const notification = document.createElement('div');
    notification.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        background: linear-gradient(135deg, #10b981 0%, #059669 100%);
        color: white;
        padding: 1rem 1.5rem;
        border-radius: 12px;
        box-shadow: 0 12px 24px rgba(16, 185, 129, 0.3);
        z-index: 1000;
        animation: slideInRight 0.3s ease;
        display: flex;
        align-items: center;
        gap: 0.75rem;
        font-weight: 600;
    `;
    notification.innerHTML = `
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none">
            <circle cx="12" cy="12" r="10" fill="rgba(255,255,255,0.2)"/>
            <path d="M9 12l2 2 4-4" stroke="white" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"/>
        </svg>
        <span>¡Excel generado exitosamente!</span>
    `;

    document.body.appendChild(notification);

    // Remove after 3 seconds
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease';
        setTimeout(() => notification.remove(), 300);
    }, 3000);
}

// Add animation styles
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            opacity: 0;
            transform: translateX(100px);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    @keyframes slideOutRight {
        from {
            opacity: 1;
            transform: translateX(0);
        }
        to {
            opacity: 0;
            transform: translateX(100px);
        }
    }
`;
document.head.appendChild(style);

// ===== Navigation Menu Enhancement =====
const navLinks = document.querySelectorAll('.nav-link');
const sections = document.querySelectorAll('section[id]');

// Smooth Scrolling with offset for sticky header
document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function (e) {
        e.preventDefault();
        const targetId = this.getAttribute('href');
        const target = document.querySelector(targetId);

        if (target) {
            const headerHeight = document.querySelector('.header').offsetHeight;
            const targetPosition = target.offsetTop - headerHeight;

            window.scrollTo({
                top: targetPosition,
                behavior: 'smooth'
            });

            // Update active link
            updateActiveLink(targetId);
        }
    });
});

// Update active link on scroll
function updateActiveLink(activeId) {
    navLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href') === activeId) {
            link.classList.add('active');
        }
    });
}

// Highlight active section on scroll
let scrollTimeout;
window.addEventListener('scroll', () => {
    clearTimeout(scrollTimeout);
    scrollTimeout = setTimeout(() => {
        let current = '';
        const headerHeight = document.querySelector('.header').offsetHeight;

        sections.forEach(section => {
            const sectionTop = section.offsetTop - headerHeight - 100;
            const sectionHeight = section.offsetHeight;

            if (window.pageYOffset >= sectionTop &&
                window.pageYOffset < sectionTop + sectionHeight) {
                current = '#' + section.getAttribute('id');
            }
        });

        if (current) {
            updateActiveLink(current);
        }
    }, 100);
});

