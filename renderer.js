window.onload = function() {
    // =========================================
    // 1. ELEMENT SELECTORS
    // =========================================
    const newFileButton = document.getElementById('btn-new-script');
    const openFileButton = document.getElementById('btn-open-file');
    const toggleRunlistButton = document.getElementById('btn-toggle-runlist');
    const fileOpener = document.getElementById('file-opener');
    const runlistContainer = document.querySelector('.runlist-files');
    const teleprompterText = document.getElementById('teleprompter-text');
    const runlistPanel = document.getElementById('runlist-panel');
    const resizer = document.getElementById('resizer');
    const teleprompterView = document.getElementById('teleprompter-view');
    const saveButton = document.getElementById('btn-save');
    const fontSettingsButton = document.getElementById('btn-font-settings');
    const fontPanel = document.getElementById('font-settings-panel');
    const fontFamilySelect = document.getElementById('font-family-select');
    const fontSizeSelect = document.getElementById('font-size-select');
    const fontSizeInc = document.getElementById('font-size-inc');
    const fontSizeDec = document.getElementById('font-size-dec');
    const bgColorButton = document.getElementById('btn-bg-color');
    const bgColorPanel = document.getElementById('bg-color-panel');
    const fontColorButton = document.getElementById('btn-font-color');
    const fontColorPanel = document.getElementById('font-color-panel');
    const boldButton = document.getElementById('btn-bold');
    const italicButton = document.getElementById('btn-italic');
    const underlineButton = document.getElementById('btn-underline');
    const extendMonitorButton = document.getElementById('btn-extend-monitor');
    const startButton = document.getElementById('btn-start');
    const stopButton = document.getElementById('btn-stop');
    const settingsGearButton = document.getElementById('btn-settings-gear');
    const settingsPanel = document.getElementById('settings-panel');
    const f2Prompt = document.getElementById('f2-prompt');
    const newBookmarkButton = document.getElementById('btn-new-bookmark');
    const prevBookmarkButton = document.getElementById('btn-prev-bookmark');
    const nextBookmarkButton = document.getElementById('btn-next-bookmark');
    const invertScrollCheckbox = document.getElementById('invert-scroll-checkbox');

    // Modal Elements
    const columnModal = document.getElementById('column-modal');
    const columnList = document.getElementById('column-list');
    const btnConfirmImport = document.getElementById('btn-confirm-import');
    const btnCancelImport = document.getElementById('btn-cancel-import');
    
    // =========================================
    // 2. IMMEDIATE UI SETUP (CRITICAL: Must be AFTER selectors)
    // =========================================
    // This ensures the main screen is black/white immediately on load
    teleprompterView.style.backgroundColor = "#000000";
    teleprompterText.style.color = "#ffffff";
    teleprompterText.style.fontSize = "48px";

    // =========================================
    // 2. APP STATE
    // =========================================
    let fileStore = [];         // Stores File objects
    let contentStore = [];      // Stores HTML strings
    let currentFileIndex = -1;
    let isRunlistVisible = false;
    let savedSelection = null;
    let mirrorWindow = null;
    let isTeleprompting = false;
    let isMouseControlActive = true;
    let scrollSpeed = 0;
    let scrollInterval = null;
    let lastActiveSpeed = 1;
    let isInvertScroll = false;
    let isMirrorActive = false;
    let mirrorScrollOffset = -150;
    let animationFrameId = null;
    
    // Excel Import Temp State
    let pendingWorkbook = null;
    let pendingFileIndex = -1;
    
    startButton.onclick = () => {
				scrollSpeed = 2; // Or your preferred speed
				startScrolling();
		};
		stopButton.onclick = () => {
				isTeleprompting = false;
				if (animationFrameId) cancelAnimationFrame(animationFrameId);
		};

    console.log("🚀 Teleprompter Engine Initialized");
    
		// =========================================
    // 3. CORE HELPERS
    // =========================================
    function saveSelection() {
        const sel = window.getSelection();
        savedSelection = (sel.rangeCount > 0) ? sel.getRangeAt(0) : null;
    }

    function restoreSelection() {
        if (savedSelection) {
            const s = window.getSelection();
            s.removeAllRanges();
            s.addRange(savedSelection);
        }
    }

    function togglePanel(button, panel) {
        [fontPanel, bgColorPanel, fontColorPanel].forEach(p => {
            if (p !== panel) p.style.display = 'none';
        });
        if (panel.style.display === 'block') {
            panel.style.display = 'none';
        } else {
            const rect = button.getBoundingClientRect();
            panel.style.top = (rect.bottom + 5) + 'px';
            panel.style.left = rect.left + 'px';
            panel.style.display = 'block';
        }
    }

function indexToColumnLetter(colIndex) {
        let letter = '';
        while (colIndex > 0) {
            let temp = (colIndex - 1) % 26;
            letter = String.fromCharCode(temp + 65) + letter;
            colIndex = (colIndex - temp - 1) / 26;
        }
        return letter;
    }

		function syncMirrorByPixels() {
        if (mirrorWindow && !mirrorWindow.closed) {
            mirrorWindow.postMessage({ 
                type: 'pixelSync', 
                scrollTop: teleprompterView.scrollTop 
            }, '*');
        }
    }

    function refreshMirrorData() {
        if (!mirrorWindow || mirrorWindow.closed) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        const cleanTextArray = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            const scriptCol = cols[cols.length - 1]; 
            return scriptCol ? scriptCol.innerText.trim() || "&nbsp;" : "&nbsp;";
        });
        mirrorWindow.postMessage({ type: 'loadContent', rows: cleanTextArray }, '*');
    }

    
    // =========================================
// 4. FILE & EXCEL MANAGEMENT (WITH LOGGING)
// =========================================

openFileButton.onclick = () => {
    console.log("📂 Open File button clicked");
    fileOpener.click();
};

fileOpener.onchange = (e) => {
    const files = Array.from(e.target.files);
    console.log(`Files selected: ${files.length}`, files);
    files.forEach(file => addFileToRunlist(file));
    fileOpener.value = ""; 
};

function addFileToRunlist(file) {
    console.log(`Adding to runlist: ${file.name}`);
    const index = fileStore.length;
    fileStore.push(file);
    contentStore.push(""); 

    const row = document.createElement('div');
    row.className = 'runlist-row';
    row.dataset.index = index;
    row.innerHTML = `
        <span class="file-name">${file.name}</span>
        <button class="btn-remove-file">×</button>
    `;

    row.onclick = (e) => {
        if (e.target.classList.contains('btn-remove-file')) return;
        console.log(`Row clicked: loading index ${index}`);
        loadScriptToEditor(index);
    };

    runlistContainer.appendChild(row);
    processFileContent(file, index);
}

async function processFileContent(file, index) {
    const extension = file.name.split('.').pop().toLowerCase();
    console.log(`Starting process for .${extension} at index ${index}`);

    try {
        if (extension === 'txt' || extension === 'html') {
            const text = await file.text();
            contentStore[index] = text;
            console.log("Text/HTML content loaded successfully");
        } 
        else if (extension === 'xlsx' || extension === 'xls') {
            console.log("Excel detected. Initializing FileReader...");
            const reader = new FileReader();
            reader.onload = (e) => {
                try {
                    console.log("FileReader load complete. Parsing with XLSX...");
                    const data = new Uint8Array(e.target.result);
                    const workbook = XLSX.read(data, { type: 'array' });
                    
                    pendingWorkbook = workbook;
                    pendingFileIndex = index;
                    
                    console.log("Workbook parsed. Opening modal.");
                    showColumnModal(workbook);
                } catch (innerErr) {
                    console.error("❌ XLSX Parsing Error:", innerErr);
                }
            };
            reader.onerror = (err) => console.error("❌ FileReader Error:", err);
            reader.readAsArrayBuffer(file);
        } else {
            console.warn(`Unsupported file type: .${extension}`);
        }
    } catch (err) {
        console.error("❌ Global processFileContent Error:", err);
    }
}

function showColumnModal(workbook) {
    if (!columnModal || !columnList) {
        console.error("❌ Modal elements not found in DOM!");
        return;
    }

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(sheet['!ref']);
    
    console.log(`Sheet: ${sheetName}, Columns: ${range.e.c + 1}`);
    columnList.innerHTML = '';

    for (let c = range.s.c; c <= range.e.c; c++) {
				const colLetter = indexToColumnLetter(c + 1);
				const label = document.createElement('label');
				label.className = "modal-checkbox-label";
				// Added 'checked' below so they start turned on
				label.innerHTML = `<input type="checkbox" value="${c}" checked> Column ${colLetter}`;
				columnList.appendChild(label);
		}
    
    columnModal.classList.remove('hidden');
    columnModal.style.display = 'flex';
}

// Excel Column Confirmation
    btnConfirmImport.onclick = () => {
        const selectedCols = Array.from(columnList.querySelectorAll('input:checked')).map(i => parseInt(i.value));
        if (selectedCols.length === 0) return;

        const sheet = pendingWorkbook.Sheets[pendingWorkbook.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
        
        let html = '<div class="script-container">'; 
        json.forEach((row, rIdx) => {
            html += `<div class="script-row-wrapper">`; 
            selectedCols.forEach(colIdx => {
                html += `<div class="script-column">${row[colIdx] || ""}</div>`;
            });
            html += '</div>';
        });
        html += '</div>';

        contentStore[pendingFileIndex] = html;
        columnModal.style.display = 'none';
        loadScriptToEditor(pendingFileIndex);

    console.log("🚀 Teleprompter Engine Ready");

};

// 2. THESE ARE THE INDEPENDENT HELPERS AT THE BOTTOM
function getOS() {
    const platform = window.navigator.userAgent.toLowerCase();
    if (platform.includes('mac')) return 'mac';
    if (platform.includes('win')) return 'win';
    return 'other';
}

async function matchResolutions() {
    const os = getOS();
    let targetW = window.screen.availWidth;
    let targetH = window.screen.availHeight;

    window.resizeTo(targetW, targetH);
    window.moveTo(0, 0);

    return { width: targetW, height: targetH, os: os };
}

function loadFileContent(file, index) {
    const reader = new FileReader();
    const ext = file.name.split('.').pop().toLowerCase();

    reader.onload = function(e) {
        if (ext === 'docx' || ext === 'doc') {
            mammoth.convertToHtml({arrayBuffer: e.target.result})
                .then(result => {
                    contentStore[index] = result.value;
                    if (currentFileIndex === index) {
                        teleprompterText.innerHTML = result.value;
                        processTableColumns();
                        updateBookmarkSidebar();
                    }
                });
        } else if (ext === 'xlsx' || ext === 'xls') {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const html = XLSX.utils.sheet_to_html(firstSheet);
            contentStore[index] = html;
            if (currentFileIndex === index) {
                teleprompterText.innerHTML = html;
                processTableColumns();
                updateBookmarkSidebar();
            }
        } else {
            const text = new TextDecoder().decode(e.target.result);
            contentStore[index] = text;
            if (currentFileIndex === index) {
                teleprompterText.innerHTML = text;
                updateBookmarkSidebar();
            }
        }
    };

    if (ext === 'txt') {
        reader.readAsArrayBuffer(file);
    } else {
        reader.readAsArrayBuffer(file);
    }
}
    
    // =========================================
    // 5. SCRIPT LOADING & EDITOR SYNC
    // =========================================
function loadScriptToEditor(index) {
    console.log(`Attempting to load index: ${index}`);
    if (fileStore[index] === null) return;

    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const row = document.querySelector(`.runlist-row[data-index="${index}"]`);
    if (row) row.classList.add('active');

    currentFileIndex = index;
    teleprompterText.innerHTML = contentStore[index];

    if (mirrorWindow && !mirrorWindow.closed) {
        // ❌ REMOVE THIS LINE:
        // mirrorWindow.postMessage({ type: 'loadContent', content: teleprompterText.innerHTML }, '*');
        
        // ✅ KEEP THESE:
        refreshMirrorData(); // This handles the clean text extraction
        syncMirrorStyles();
    }
    console.log("Editor and Mirror updated with Content and Styles.");
}
    
// =========================================
// 6. MIRROR WINDOW LOGIC (OS AWARE)
// =========================================

// Helper to handle the actual Window opening & HTML injection
function setupMirrorWindow(width, height) {
    const mirrorHTML = `
    <html>
    <head>
<style>
    body { background: #000; margin: 0; padding: 0; overflow: hidden; color: #fff; }
    #scroll-container { height: 100vh; width: 100vw; overflow-y: scroll; scrollbar-width: none; }
    #scroll-container::-webkit-scrollbar { display: none; }
    #mirror-content { width: 100%; }
    .mirror-row { 
        width: 100%; 
        padding: 10px 40px; /* Matches your .cell-locker padding */
        box-sizing: border-box;
        border-bottom: 1px solid #222; 
        line-height: 1.4 !important; /* Forces the same height as main screen */
        white-space: pre-wrap; 
        display: block;
    }
</style>
    </head>
    <body>
        <div id="scroll-container"><div id="mirror-content"></div></div>
        <script>
            const container = document.getElementById('scroll-container');
            const display = document.getElementById('mirror-content');
            window.onmessage = (e) => {
                const { type, rows, style, scrollTop } = e.data;
                if (type === 'loadContent' && rows) {
                    display.innerHTML = rows.map(text => '<div class="mirror-row">' + text + '</div>').join('');
                }
                if (type === 'syncStyleLite') {
                    display.style.fontSize = style.fontSize;
                    display.style.fontFamily = style.fontFamily;
                    display.style.lineHeight = style.lineHeight;
                }
                if (type === 'pixelSync') {
                    container.scrollTo({ top: scrollTop, behavior: 'instant' });
                }
            };
        <\/script>
    </body>
    </html>`;
    mirrorWindow.document.write(mirrorHTML);
    mirrorWindow.document.close();
}

extendMonitorButton.onclick = async () => {
		document.body.classList.add('broadcasting');
    try {
        if ('getScreenDetails' in window) {
            const screenDetails = await window.getScreenDetails();
            const secondary = screenDetails.screens.find(s => s !== screenDetails.currentScreen);
            const left = secondary ? secondary.availLeft : window.screen.width;
            const specs = `left=${left},top=0,width=800,height=600`;
            mirrorWindow = window.open('', 'TeleprompterMirror', specs);
        } else {
            mirrorWindow = window.open('', 'TeleprompterMirror', 'width=800,height=600,left=2000');
        }
        
        if (mirrorWindow) {
            setupMirrorWindow(); 
            setTimeout(() => {
                refreshMirrorData();
            }, 500);
        }
    } catch (err) {
        console.error("Monitor Extension Failed:", err);
    }
};
    
// =========================================
// 7. SCROLLING ENGINE (PIXEL-PERFECT)
// =========================================

function syncMirrorStyles() {
    if (!mirrorWindow || mirrorWindow.closed) return;
    const mainStyle = window.getComputedStyle(teleprompterText);
    mirrorWindow.postMessage({ 
        type: 'syncStyleLite', 
        style: {
            fontSize: mainStyle.fontSize,
            fontFamily: mainStyle.fontFamily,
            lineHeight: mainStyle.lineHeight
        } 
    }, '*');
}

function startScrolling() {
    if (isTeleprompting) return;
    isTeleprompting = true;
    const move = () => {
        if (!isTeleprompting) return;
        const direction = isInvertScroll ? -1 : 1;
        teleprompterView.scrollTop += (scrollSpeed * direction);
        syncMirrorByPixels(); 
        animationFrameId = requestAnimationFrame(move);
    };
    animationFrameId = requestAnimationFrame(move);
}
};
