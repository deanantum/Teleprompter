window.onload = function() {
    // =========================================
    // 1. ELEMENT SELECTORS
    // =========================================
    const openFileButton = document.getElementById('btn-open-file');
    const toggleRunlistButton = document.getElementById('btn-toggle-runlist');
    const fileOpener = document.getElementById('file-opener');
    const runlistContainer = document.querySelector('.runlist-files');
    const teleprompterText = document.getElementById('teleprompter-text');
    const runlistPanel = document.getElementById('runlist-panel');
    const resizer = document.getElementById('resizer');
    const teleprompterView = document.getElementById('teleprompter-view');
    const teleprompterStage = document.querySelector('.teleprompter-stage');
    const indicatorWrapper = document.getElementById('indicator-wrapper');
    const indicatorTriangle = document.getElementById('indicator-triangle');
    const saveButton = document.getElementById('btn-save');
    const fontFamilySelect = document.getElementById('font-family-select');
    const fontSizeSelect = document.getElementById('font-size-select');
    const selectFontTarget = document.getElementById('select-font-target');
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
    const btnMoveFileUp = document.getElementById('btn-move-file-up');
    const btnMoveFileDown = document.getElementById('btn-move-file-down');

    // =========================================
    // 2. IMMEDIATE UI SETUP (CRITICAL: Must be AFTER selectors)
    // =========================================
    // This ensures the main screen is black/white immediately on load
    teleprompterView.style.backgroundColor = "#000000";
    teleprompterText.style.color = "#ffffff";
    teleprompterText.style.fontFamily = fontFamilySelect?.value || "Arial";
    teleprompterText.style.fontSize = (fontSizeSelect?.value || "80") + "px";

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
    let isPaused = false;
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

    let lastColumnWidthPx = null;
    let col2WidthPx = null;
    let col3WidthPx = null;
    let charCountsPerRow = [];
    let rowColorsCache = []; /* Stores row color classes from unextended view so colors don't change when extended */
    let measuredRowHeights = []; /* Actual measured heights from DOM to prevent text overlap when extended */
    let selectedFontTarget = null; /* { fontFamily, fontSize } when user selects from Select dropdown */
    let savedFontSelection = null; /* Cloned range saved when user focuses font dropdowns */
    let lastFontChangeSource = null; /* 'family' | 'size' - which dropdown triggered the change */

    function saveFontSelectionFromEditor() {
        const sel = window.getSelection();
        if (!sel.rangeCount) {
            savedFontSelection = null;
            return;
        }
        const r = sel.getRangeAt(0);
        if (r.collapsed) {
            savedFontSelection = null;
            return;
        }
        const root = r.commonAncestorContainer;
        const el = root.nodeType === Node.TEXT_NODE ? root.parentElement : root;
        if (!el || !teleprompterText.contains(el)) {
            savedFontSelection = null;
            return;
        }
        try {
            savedFontSelection = r.cloneRange();
        } catch (_) {
            savedFontSelection = null;
        }
    }

    function getComputedFontFromSelection(range) {
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            if (!range.intersectsNode(node)) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            return {
                fontFamily: (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, ''),
                fontSize: style.fontSize || ''
            };
        }
        return null;
    }

    function wrapSelectionInSpan(range, fontVal, sizeVal, preserveSize, preserveFamily) {
        const getCell = (node) => {
            let n = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
            while (n && n !== teleprompterText) {
                if (n.classList?.contains('script-column')) return n;
                n = n.parentElement;
            }
            return null;
        };
        const startCell = getCell(range.startContainer);
        const endCell = getCell(range.endContainer);
        if (startCell != null && endCell != null && startCell !== endCell) return;

        const computed = getComputedFontFromSelection(range);
        const useFont = preserveFamily && computed ? computed.fontFamily : (fontVal != null && fontVal !== '' ? fontVal : null);
        const useSize = preserveSize && computed ? computed.fontSize : (sizeVal != null && sizeVal !== '' ? sizeVal + 'px' : null);

        const toWrap = [];
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            if (!range.intersectsNode(node)) continue;
            let startOff = 0, endOff = node.length;
            if (node === range.startContainer) startOff = range.startOffset;
            if (node === range.endContainer) endOff = range.endOffset;
            if (startOff >= endOff) continue;
            toWrap.push({ node, startOff, endOff });
        }
        const newSpans = [];
        toWrap.forEach(({ node, startOff, endOff }) => {
            let target = node;
            if (startOff > 0) target = node.splitText(startOff);
            const len = endOff - startOff;
            if (target.length > len && len > 0) target.splitText(len);
            const span = document.createElement('span');
            span.style.display = 'inline';
            if (useFont) span.style.fontFamily = useFont;
            if (useSize) span.style.fontSize = useSize;
            target.parentNode.insertBefore(span, target);
            span.appendChild(target);
            newSpans.push(span);
        });
        if (newSpans.length > 0) {
            const r = document.createRange();
            r.setStartBefore(newSpans[0]);
            r.setEndAfter(newSpans[newSpans.length - 1]);
            const sel = window.getSelection();
            sel.removeAllRanges();
            sel.addRange(r);
        }
    }

    function applyFontToMatchingTarget() {
        if (!selectedFontTarget || (!fontFamilySelect && !fontSizeSelect)) return;
        const { fontFamily: targetFont, fontSize: targetSize } = selectedFontTarget;
        const newFont = fontFamilySelect?.value;
        const newSize = fontSizeSelect?.value;
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        const toWrap = [];
        let node;
        while ((node = walker.nextNode())) {
            if (!node.textContent.trim()) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            const pFont = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
            const pSize = style.fontSize || '';
            if (pFont === targetFont && pSize === targetSize) {
                toWrap.push(node);
            }
        }
        toWrap.forEach(textNode => {
            const span = document.createElement('span');
            if (newFont) span.style.fontFamily = newFont;
            if (newSize) span.style.fontSize = newSize + 'px';
            textNode.parentNode.insertBefore(span, textNode);
            span.appendChild(textNode);
        });
    }

    function collectUniqueFontSizes() {
        const pairs = new Map();
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            if (!node.textContent.trim()) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            const fontFamily = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
            const fontSize = style.fontSize || '';
            const key = fontFamily + '|' + fontSize;
            if (!pairs.has(key)) pairs.set(key, { fontFamily, fontSize });
        }
        return Array.from(pairs.values()).sort((a, b) => {
            const c = (a.fontFamily || '').localeCompare(b.fontFamily || '');
            return c !== 0 ? c : (parseFloat(a.fontSize) || 0) - (parseFloat(b.fontSize) || 0);
        });
    }

    function refreshSelectFontTarget() {
        if (!selectFontTarget) return;
        const current = selectFontTarget.value;
        const opts = collectUniqueFontSizes();
        selectFontTarget.innerHTML = '<option value="">Select</option>';
        opts.forEach(({ fontFamily, fontSize }) => {
            const sizeNum = parseInt(fontSize, 10) || fontSize;
            const label = `${fontFamily} ${sizeNum}`;
            const val = JSON.stringify({ fontFamily, fontSize });
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = label;
            if (current === val) opt.selected = true;
            selectFontTarget.appendChild(opt);
        });
        try {
            selectedFontTarget = selectFontTarget.value ? JSON.parse(selectFontTarget.value) : null;
        } catch (_) {
            selectedFontTarget = null;
        }
    }

    function applyFontColorToTarget(color) {
        if (selectedFontTarget) {
            const { fontFamily, fontSize } = selectedFontTarget;
            const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
            const toWrap = [];
            let node;
            while ((node = walker.nextNode())) {
                if (!node.textContent.trim()) continue;
                const parent = node.parentElement;
                if (!parent) continue;
                const style = window.getComputedStyle(parent);
                const pFont = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
                const pSize = style.fontSize || '';
                if (pFont === fontFamily && pSize === fontSize) {
                    toWrap.push(node);
                }
            }
            toWrap.forEach(textNode => {
                const span = document.createElement('span');
                span.style.color = color;
                textNode.parentNode.insertBefore(span, textNode);
                span.appendChild(textNode);
            });
        } else {
            teleprompterText.style.color = color;
        }
        syncEditorState();
    }

    function updatePlayPauseButton() {
        const icon = startButton.querySelector('i');
        if (!icon) return;
        if (isTeleprompting && !isPaused) {
            icon.className = 'fa-solid fa-pause';
            startButton.title = 'Pause';
            if (indicatorWrapper) indicatorWrapper.classList.remove('stopped', 'paused'), indicatorWrapper.classList.add('running');
        } else {
            icon.className = 'fa-solid fa-play';
            startButton.title = 'Play';
            if (indicatorWrapper) {
                indicatorWrapper.classList.remove('running');
                indicatorWrapper.classList.add(isPaused ? 'paused' : 'stopped');
            }
        }
    }

    startButton.onclick = () => {
        if (!isTeleprompting) {
            scrollSpeed = 2;
            isPaused = false;
            startScrolling();
        } else if (isPaused) {
            isPaused = false;
            startScrolling();
        } else {
            isPaused = true;
            isTeleprompting = false;
            if (animationFrameId) cancelAnimationFrame(animationFrameId);
            animationFrameId = null;
        }
        updatePlayPauseButton();
    };
    stopButton.onclick = () => {
        isTeleprompting = false;
        isPaused = false;
        if (animationFrameId) cancelAnimationFrame(animationFrameId);
        animationFrameId = null;
        teleprompterView.scrollTop = 0;
        syncMirrorByPixels();
        updatePlayPauseButton();
    };

    console.log("🚀 Teleprompter Engine Initialized");

    requestAnimationFrame(() => refreshSelectFontTarget());

    function updateFontSelectsFromCaret() {
        const sel = window.getSelection();
        if (!sel.rangeCount || sel.rangeCount === 0) return;
        const node = sel.anchorNode;
        if (!node || !teleprompterText.contains(node)) return;
        const el = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
        if (!el || !teleprompterText.contains(el)) return;

        const style = window.getComputedStyle(el);
        const fontFamily = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
        const fontSizePx = parseFloat(style.fontSize) || 80;

        if (fontFamilySelect) {
            const opts = Array.from(fontFamilySelect.options).map(o => o.value);
            if (opts.includes(fontFamily)) {
                fontFamilySelect.value = fontFamily;
            }
        }
        if (fontSizeSelect) {
            const opts = Array.from(fontSizeSelect.options).map(o => parseInt(o.value, 10));
            const nearest = opts.reduce((a, b) => Math.abs(a - fontSizePx) < Math.abs(b - fontSizePx) ? a : b);
            fontSizeSelect.value = String(nearest);
        }
    }

    function applyFontSettings() {
        const storedRange = savedFontSelection;
        savedFontSelection = null;
        const fontVal = fontFamilySelect?.value;
        const sizeVal = fontSizeSelect?.value;

        const doApply = () => {
            let range = null;
            let hasSelection = false;

            if (storedRange && !storedRange.collapsed) {
                try {
                    if (document.contains(storedRange.startContainer)) {
                        teleprompterText.focus();
                        const sel = window.getSelection();
                        sel.removeAllRanges();
                        sel.addRange(storedRange);
                        range = storedRange;
                        hasSelection = true;
                    }
                } catch (_) {}
            }
            if (!hasSelection) {
                teleprompterText.focus();
                const sel = window.getSelection();
                range = sel.rangeCount ? sel.getRangeAt(0) : null;
                hasSelection = range && !range.collapsed && teleprompterText.contains(range.commonAncestorContainer);
            }

            if (hasSelection && (fontFamilySelect || fontSizeSelect)) {
                const preserveSize = lastFontChangeSource === 'family';
                const preserveFamily = lastFontChangeSource === 'size';
                wrapSelectionInSpan(range, fontVal, sizeVal, preserveSize, preserveFamily);
                teleprompterText.focus();
            } else if (selectedFontTarget) {
                applyFontToMatchingTarget();
            } else {
                if (fontFamilySelect) teleprompterText.style.fontFamily = fontVal || fontFamilySelect.value;
                if (fontSizeSelect) teleprompterText.style.fontSize = (sizeVal || fontSizeSelect.value) + 'px';
            }
            if (mirrorWindow && !mirrorWindow.closed) syncMirrorStyles();
            syncEditorState();
        };

        teleprompterText.focus();
        requestAnimationFrame(() => {
            requestAnimationFrame(doApply);
        });
    }
    const saveSelectionWhenFocusingFontControl = (e) => {
        if (fontFamilySelect?.contains(e.target) || fontSizeSelect?.contains(e.target)) {
            saveFontSelectionFromEditor();
        }
    };
    document.addEventListener('mousedown', saveSelectionWhenFocusingFontControl, true);
    document.addEventListener('pointerdown', saveSelectionWhenFocusingFontControl, true);

    if (fontFamilySelect) {
        fontFamilySelect.addEventListener('change', () => {
            lastFontChangeSource = 'family';
            applyFontSettings();
        });
    }
    if (fontSizeSelect) {
        fontSizeSelect.addEventListener('change', () => {
            lastFontChangeSource = 'size';
            applyFontSettings();
        });
    }

    document.addEventListener('selectionchange', () => {
        const sel = window.getSelection();
        if (sel.rangeCount) {
            const r = sel.getRangeAt(0);
            if (!r.collapsed && teleprompterText.contains(r.commonAncestorContainer)) {
                try {
                    savedFontSelection = r.cloneRange();
                    selectedFontTarget = null;
                } catch (_) {}
            }
        }
        if (document.activeElement === teleprompterText) {
            updateFontSelectsFromCaret();
        }
    });
    teleprompterText.addEventListener('click', updateFontSelectsFromCaret);
    teleprompterText.addEventListener('keyup', updateFontSelectsFromCaret);
    teleprompterText.addEventListener('focus', updateFontSelectsFromCaret);

    document.addEventListener('keydown', (e) => {
        if (!teleprompterText.contains(document.activeElement) && document.activeElement !== teleprompterText) return;
        const mod = e.metaKey || e.ctrlKey;
        if (mod && e.key === 'z') {
            e.preventDefault();
            document.execCommand(e.shiftKey ? 'redo' : 'undo');
        } else if (mod && e.key === 'y') {
            e.preventDefault();
            document.execCommand('redo');
        }
    });

    if (selectFontTarget) {
        selectFontTarget.addEventListener('change', () => {
            try {
                selectedFontTarget = selectFontTarget.value ? JSON.parse(selectFontTarget.value) : null;
            } catch (_) {
                selectedFontTarget = null;
            }
        });
    }

    if (bgColorButton && bgColorPanel) bgColorButton.onclick = () => togglePanel(bgColorButton, bgColorPanel);
    if (fontColorButton && fontColorPanel) fontColorButton.onclick = () => togglePanel(fontColorButton, fontColorPanel);

    document.addEventListener('click', (e) => {
        const fc = e.target.closest('.color-options');
        if (fc && fc.closest('#font-color-panel')) {
            const box = e.target.closest('.color-box');
            if (box) {
                const color = box.getAttribute('data-color');
                if (color) applyFontColorToTarget(color);
                fontColorPanel.style.display = 'none';
            }
        } else if (fc && fc.closest('#bg-color-panel')) {
            const box = e.target.closest('.color-box');
            if (box) {
                const color = box.getAttribute('data-color');
                if (color) teleprompterView.style.backgroundColor = color;
                bgColorPanel.style.display = 'none';
            }
        }
    });

    teleprompterView.addEventListener('scroll', syncMirrorByPixels);

    (function setupIndicatorDrag() {
        if (!indicatorWrapper || !indicatorTriangle || !teleprompterStage) return;
        let isDragging = false;
        let startY = 0;
        let startTop = 0;

        indicatorTriangle.addEventListener('mousedown', (e) => {
            e.preventDefault();
            isDragging = true;
            const rect = teleprompterStage.getBoundingClientRect();
            startY = e.clientY;
            const wrapperRect = indicatorWrapper.getBoundingClientRect();
            startTop = wrapperRect.top - rect.top + (wrapperRect.height / 2);
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            const rect = teleprompterStage.getBoundingClientRect();
            const deltaY = e.clientY - startY;
            let newTop = startTop + deltaY;
            newTop = Math.max(0, Math.min(rect.height, newTop));
            indicatorWrapper.style.top = newTop + 'px';
            indicatorWrapper.style.transform = 'translateY(-50%)';
            startY = e.clientY;
            startTop = newTop;
            syncMirrorByPixels();
        }, { passive: true });

        document.addEventListener('mouseup', () => {
            if (isDragging) {
                isDragging = false;
                syncMirrorByPixels();
            }
        });
    })();

    let isUpdatingFromMirrorScroll = false;

    function handleMirrorScroll(data) {
        if (!mirrorWindow || mirrorWindow.closed) return;
        if (!data || typeof data.scrollTop !== 'number') return;
        isUpdatingFromMirrorScroll = true;
        const maxScroll = data.scrollHeight - data.clientHeight;
        const ratio = (maxScroll > 0 && typeof data.scrollHeight === 'number')
            ? data.scrollTop / maxScroll
            : null;
        if (ratio !== null && !isNaN(ratio) && ratio >= 0 && ratio <= 1) {
            const mainMax = teleprompterView.scrollHeight - teleprompterView.clientHeight;
            teleprompterView.scrollTop = Math.round(ratio * Math.max(0, mainMax));
        } else {
            teleprompterView.scrollTop = Math.max(0, Math.min(data.scrollTop, teleprompterView.scrollHeight));
        }
        requestAnimationFrame(() => { isUpdatingFromMirrorScroll = false; });
    }

    function handleMirrorReady() {
        syncColumnWidths();
        refreshMirrorData();
        syncMirrorStyles();
    }

    let overflowReportCount = 0;
    function handleRowOverflowReport(data) {
        if (!mirrorWindow || mirrorWindow.closed || !Array.isArray(data.rowHeights)) return;
        if (overflowReportCount >= 3) return; /* prevent infinite loop */
        const changed = measuredRowHeights.length !== data.rowHeights.length ||
            data.rowHeights.some((h, i) => Math.abs((measuredRowHeights[i] || 0) - h) > 2);
        if (!changed) return;
        overflowReportCount++;
        measuredRowHeights = data.rowHeights;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        rows.forEach((row, i) => {
            const h = measuredRowHeights[i];
            if (h > 0) {
                row.style.minHeight = h + 'px';
                row.style.height = h + 'px';
            }
        });
        mirrorWindow.postMessage({ type: 'updateRowHeights', rowHeights: measuredRowHeights }, '*');
        requestAnimationFrame(() => {
            syncMirrorByPixels();
            syncMirrorStyles();
        });
    }

    try {
        const mirrorChannel = new BroadcastChannel('teleprompter-mirror-sync');
        mirrorChannel.onmessage = (e) => {
            try {
                const data = e.data || {};
                if (data.type === 'mirrorReady') handleMirrorReady();
                if (data.type === 'mirrorScroll') handleMirrorScroll(data);
                if (data.type === 'rowOverflowReport') handleRowOverflowReport(data);
            } catch (err) { console.warn('Mirror channel message error:', err); }
        };
    } catch (e) { /* BroadcastChannel not supported */ }

    window.addEventListener('message', (e) => {
        try {
            const data = e.data || {};
            if (data.type === 'mirrorReady') handleMirrorReady();
            if (data.type === 'mirrorScroll') handleMirrorScroll(data);
            if (data.type === 'rowOverflowReport') handleRowOverflowReport(data);
        } catch (err) { console.warn('Mirror message error:', err); }
    });
    
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
        [bgColorPanel, fontColorPanel].forEach(p => {
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

		function getIndicatorPositionPercent() {
        if (!indicatorWrapper || !teleprompterStage) return 50;
        const stageRect = teleprompterStage.getBoundingClientRect();
        const wrapperRect = indicatorWrapper.getBoundingClientRect();
        const centerY = wrapperRect.top - stageRect.top + (wrapperRect.height / 2);
        return Math.max(0, Math.min(100, (centerY / stageRect.height) * 100));
    }

		function syncMirrorByPixels() {
        if (isUpdatingFromMirrorScroll || !mirrorWindow || mirrorWindow.closed) return;
        mirrorWindow.postMessage({ 
                type: 'pixelSync', 
                scrollTop: teleprompterView.scrollTop,
                indicatorPosition: getIndicatorPositionPercent()
            }, '*');
    }

    /** Measures actual content height of each row's script column to prevent overlap. Call before hiding last column. */
    function measureRowHeightsFromContent() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        measuredRowHeights = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            const scriptCol = cols[cols.length - 1];
            const cellLocker = scriptCol?.querySelector('.cell-locker');
            if (!cellLocker) return 0;
            const pad = 8; /* vertical padding buffer */
            const scrollH = cellLocker.scrollHeight;
            const clientH = cellLocker.clientHeight;
            return Math.max(scrollH, clientH, 20) + pad;
        });
    }

    function buildCharCountsPerRow() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        charCountsPerRow = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            let maxChars = 0;
            cols.forEach(col => {
                const n = (col.innerText || '').trim().length;
                if (n > maxChars) maxChars = n;
            });
            return Math.max(maxChars, 1);
        });
    }

    function getCharsPerLine(widthPx) {
        const mainStyle = window.getComputedStyle(teleprompterText);
        const sample = 'The quick brown fox jumps over the lazy dog.';
        const span = document.createElement('span');
        span.style.cssText = 'position:absolute;visibility:hidden;white-space:nowrap;font:' + mainStyle.font + ';';
        span.textContent = sample;
        document.body.appendChild(span);
        const sampleWidth = span.getBoundingClientRect().width;
        span.remove();
        const avgCharWidth = sampleWidth / sample.length;
        return Math.max(1, Math.floor(widthPx / avgCharWidth));
    }

    const ROW_COLOR_CLASSES = ['row-lines-same', 'row-col2-more', 'row-col3-more', 'row-col3-much-more'];

    function applyRowColors() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;

        const maxCols = Math.max(...rows.map(r => r.querySelectorAll('.script-column').length));
        if (maxCols !== 3) {
            rows.forEach(row => ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c)));
            rowColorsCache = [];
            return;
        }

        const isBroadcasting = document.body.classList.contains('broadcasting');

        if (isBroadcasting && rowColorsCache.length === rows.length) {
            /* Use cached colors from unextended view so colors don't change when extended */
            rows.forEach((row, i) => {
                ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c));
                const cached = rowColorsCache[i] || 'row-lines-same';
                row.classList.add(cached);
            });
            return;
        }

        /* Compute colors when unextended and store in cache */
        const firstRow = rows[0];
        const r0cols = firstRow.querySelectorAll('.script-column');
        const w2px = r0cols[1]?.getBoundingClientRect().width || col2WidthPx || lastColumnWidthPx || 400;
        const w3px = r0cols[2]?.getBoundingClientRect().width || col3WidthPx || col2WidthPx || lastColumnWidthPx || 400;
        const colPadding = 40;
        const w2 = Math.max(50, w2px - colPadding);
        const w3 = Math.max(50, w3px - colPadding);
        const cpl2 = getCharsPerLine(w2);
        const cpl3 = getCharsPerLine(w3);

        rowColorsCache = [];
        rows.forEach(row => {
            ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c));
            const cols = row.querySelectorAll('.script-column');
            const col2 = cols[1];
            const col3 = cols[2];
            const text2 = (col2?.innerText || '').trim();
            const text3 = (col3?.innerText || '').trim();

            const lines2 = Math.max(1, Math.ceil(text2.length / cpl2));
            const lines3 = Math.max(1, Math.ceil(text3.length / cpl3));
            let cls;
            if (lines2 === lines3) {
                cls = 'row-lines-same';
            } else if (lines2 > lines3) {
                cls = 'row-col2-more';
            } else if (lines3 >= lines2 + 2) {
                cls = 'row-col3-much-more';
            } else {
                cls = 'row-col3-more';
            }
            row.classList.add(cls);
            rowColorsCache.push(cls);
        });
    }

    function refreshMirrorData() {
        if (!mirrorWindow || mirrorWindow.closed) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        let cleanTextArray;
        let rowHeights = [];
        if (rows.length > 0) {
            buildCharCountsPerRow();

            const mainStyle = window.getComputedStyle(teleprompterText);
            const fontSize = parseFloat(mainStyle.fontSize) || 48;
            const lh = parseFloat(mainStyle.lineHeight) || 1.4;
            const oneLineHeight = Math.ceil(lh < 3 ? fontSize * lh : lh);
            const colPadding = 40;
            const contentWidthPx = Math.max(100, (lastColumnWidthPx || teleprompterText.clientWidth) - colPadding);
            const rowPad = 4;

            if (document.body.classList.contains('broadcasting') && charCountsPerRow.length > 0) {
                const charsPerLine = getCharsPerLine(contentWidthPx);
                const estimatedHeights = charCountsPerRow.map(maxChars => {
                    const numLines = Math.max(1, Math.ceil(maxChars / charsPerLine));
                    return numLines * oneLineHeight + rowPad;
                });
                rowHeights = (measuredRowHeights.length === rows.length)
                    ? measuredRowHeights.map((mh, i) => Math.max(mh, estimatedHeights[i] || 0))
                    : estimatedHeights;
                rows.forEach((row, i) => {
                    const h = rowHeights[i];
                    if (h > 0) {
                        row.style.minHeight = h + 'px';
                        row.style.height = h + 'px';
                    }
                });
            }

            cleanTextArray = rows.map(row => {
                const cols = row.querySelectorAll('.script-column');
                const scriptCol = cols[cols.length - 1];
                const scriptText = scriptCol ? scriptCol.innerText.trim() : '';
                return scriptText || "\u00A0";
            });
        } else {
            const fullText = teleprompterText.innerText || '';
            cleanTextArray = fullText ? fullText.split(/\r?\n/).map(line => line.trim() || "\u00A0") : ["\u00A0"];
        }
        const contentWidth = (lastColumnWidthPx && lastColumnWidthPx > 50) ? `${lastColumnWidthPx}px` : null;
        const rowColors = rows.length > 0
            ? (document.body.classList.contains('broadcasting') && rowColorsCache.length === rows.length
                ? rowColorsCache
                : rows.map(row => ROW_COLOR_CLASSES.find(c => row.classList.contains(c)) || ''))
            : [];
        mirrorWindow.postMessage({ type: 'loadContent', rows: cleanTextArray, rowHeights: rowHeights, contentWidth: contentWidth, rowColors: rowColors }, '*');
    }

    function applyBroadcastingVisibility() {
        const isBroadcasting = document.body.classList.contains('broadcasting');
        const lastCols = teleprompterText.querySelectorAll('.script-row-wrapper .script-column:last-child');
        lastCols.forEach(el => {
            if (isBroadcasting) {
                el.classList.add('broadcast-hidden');
            } else {
                el.classList.remove('broadcast-hidden');
            }
        });
        teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
            row.style.minHeight = '';
            row.style.height = '';
        });
    }

    function syncColumnWidths() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) {
            applyBroadcastingVisibility();
            return;
        }

        let maxCols = 0;
        rows.forEach(row => {
            const cols = row.querySelectorAll('.script-column');
            maxCols = Math.max(maxCols, cols.length);
        });
        if (maxCols === 0) return;

        const widths = new Array(maxCols).fill(0);
        let cellStyle = null;

        const sample = document.createElement('span');
        sample.textContent = '0000';
        sample.style.position = 'absolute';
        sample.style.visibility = 'hidden';
        sample.style.whiteSpace = 'pre';
        sample.style.font = window.getComputedStyle(teleprompterText).font;
        document.body.appendChild(sample);
        const digitsWidth = sample.getBoundingClientRect().width;
        sample.remove();

        const firstRow = rows[0];
        const firstCell = firstRow?.querySelector('.script-column .cell-locker') || firstRow?.querySelector('.script-column');
        if (firstCell) cellStyle = window.getComputedStyle(firstCell);

        let paddingWidth = 0;
        if (cellStyle) {
            paddingWidth = (parseFloat(cellStyle.paddingLeft) || 0) + (parseFloat(cellStyle.paddingRight) || 0);
        }
        widths[0] = Math.ceil(digitsWidth + paddingWidth);

        const isBroadcasting = document.body.classList.contains('broadcasting');
        const visibleColCount = isBroadcasting && maxCols > 1 ? maxCols - 1 : maxCols;
        const otherColCount = visibleColCount - 1;

        const availableWidth = teleprompterText.clientWidth;
        const remaining = Math.max(availableWidth - widths[0], 0);
        const equalWidth = otherColCount > 0 ? Math.floor(remaining / otherColCount) : 0;

        for (let i = 1; i < widths.length; i += 1) {
            widths[i] = (i < visibleColCount) ? equalWidth : 0;
        }
        if (otherColCount > 0 && equalWidth > 0 && visibleColCount > 1) {
            widths[visibleColCount - 1] += remaining - (equalWidth * otherColCount);
        }

        lastColumnWidthPx = (maxCols > 1 && rows[0]) ? rows[0].querySelectorAll('.script-column')[1]?.getBoundingClientRect().width || remaining / (maxCols - 1) : null;

        if (maxCols >= 3 && rows[0] && !isBroadcasting) {
            const r0cols = rows[0].querySelectorAll('.script-column');
            col2WidthPx = r0cols[1]?.getBoundingClientRect().width || null;
            col3WidthPx = r0cols[2]?.getBoundingClientRect().width || col2WidthPx;
        }

        rows.forEach(row => {
            const cols = row.querySelectorAll('.script-column');
            cols.forEach((col, idx) => {
                const isLast = idx === cols.length - 1;
                if (isBroadcasting && isLast) {
                    col.style.flex = '0 0 0';
                    col.style.width = '0';
                    col.style.minWidth = '0';
                } else {
                    const width = widths[idx] || 0;
                    if (isBroadcasting && idx === visibleColCount - 1) {
                        col.style.flex = '1 1 auto';
                        col.style.width = '';
                        col.style.minWidth = '0';
                    } else {
                        col.style.flex = width > 0 ? `0 0 ${width}px` : '';
                        col.style.width = width > 0 ? `${width}px` : '';
                    }
                }
            });
        });
        applyBroadcastingVisibility();
        if (rows.length > 0) {
            buildCharCountsPerRow();
            applyRowColors();
        }
    }

    /** Sync editor state to arrays (contentStore, charCountsPerRow, rowColorsCache) and mirror. Call after user actions (blur, save, etc.). */
    function syncEditorState() {
        if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
            const html = teleprompterText.innerHTML.trim();
            contentStore[currentFileIndex] = (html === '<br>' || html === '') ? '' : html;
        }
        refreshSelectFontTarget();
        syncColumnWidths();
        if (mirrorWindow && !mirrorWindow.closed) {
            refreshMirrorData();
            syncMirrorStyles();
        }
    }

    teleprompterText.addEventListener('blur', syncEditorState);
    document.addEventListener('click', (e) => {
        if (!teleprompterText.contains(e.target)) syncEditorState();
    });

    let inputColorDebounce;
    teleprompterText.addEventListener('input', () => {
        clearTimeout(inputColorDebounce);
        inputColorDebounce = setTimeout(() => {
            refreshSelectFontTarget();
            const rows = teleprompterText.querySelectorAll('.script-row-wrapper');
            const maxCols = rows.length ? Math.max(...Array.from(rows).map(r => r.querySelectorAll('.script-column').length)) : 0;
            if (maxCols >= 3) {
                rowColorsCache = [];
                syncColumnWidths();
                if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
            } else {
                syncColumnWidths();
            }
        }, 120);
    });

    // =========================================
// 4. FILE & EXCEL MANAGEMENT (WITH LOGGING)
// =========================================

openFileButton.onclick = () => {
    console.log("📂 Open File button clicked");
    fileOpener.click();
};

async function saveCurrentFile() {
    if (currentFileIndex < 0 || currentFileIndex >= fileStore.length) return;
    syncEditorState();
    const file = fileStore[currentFileIndex];
    const ext = file.name.split('.').pop().toLowerCase();

    const html = teleprompterText.innerHTML.trim();
    contentStore[currentFileIndex] = (html === '<br>' || html === '') ? '' : html;

    let blob;
    let suggestedName = file.name;

    if (ext === 'txt') {
        const text = teleprompterText.innerText || '';
        blob = new Blob([text], { type: 'text/plain' });
    } else if (ext === 'html' || ext === 'htm') {
        blob = new Blob([contentStore[currentFileIndex]], { type: 'text/html' });
    } else if (ext === 'xlsx' || ext === 'xls') {
        const container = teleprompterText.querySelector('.script-container');
        const rows = [];
        if (container) {
            container.querySelectorAll('.script-row-wrapper').forEach(rowEl => {
                const row = [];
                rowEl.querySelectorAll('.script-column').forEach(col => {
                    row.push((col.innerText || '').trim());
                });
                rows.push(row);
            });
        }
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet(rows.length > 0 ? rows : [['']]);
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        blob = new Blob([wbout], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        suggestedName = suggestedName.replace(/\.xls$/i, '.xlsx');
    } else if (ext === 'docx' || ext === 'doc') {
        blob = new Blob([contentStore[currentFileIndex]], { type: 'text/html' });
        suggestedName = suggestedName.replace(/\.doc$/i, '.html');
    } else {
        blob = new Blob([contentStore[currentFileIndex]], { type: 'text/plain' });
    }

    try {
        if (typeof saveAs === 'function') {
            saveAs(blob, suggestedName);
        } else if ('showSaveFilePicker' in window) {
            let types;
            if (ext === 'xlsx' || ext === 'xls') {
                types = [{ description: 'Excel', accept: { 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'] } }];
            } else if (ext === 'html' || ext === 'htm') {
                types = [{ description: 'HTML', accept: { 'text/html': ['.html'] } }];
            } else {
                types = [{ description: 'Text', accept: { 'text/plain': ['.txt'] } }];
            }
            const handle = await window.showSaveFilePicker({ suggestedName, types });
            const writable = await handle.createWritable();
            await writable.write(blob);
            await writable.close();
        } else {
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = suggestedName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(a.href), 100);
        }
    } catch (err) {
        if (err.name !== 'AbortError') console.error('Save failed:', err);
    }
}

saveButton.onclick = () => saveCurrentFile();

fileOpener.onchange = (e) => {
    const files = Array.from(e.target.files);
    console.log(`Files selected: ${files.length}`, files);
    files.forEach(file => addFileToRunlist(file));
    fileOpener.value = "";
};

function moveFileInRunlist(direction) {
    if (currentFileIndex < 0 || fileStore.length < 2) return;
    const newIndex = direction === 'up' ? currentFileIndex - 1 : currentFileIndex + 1;
    if (newIndex < 0 || newIndex >= fileStore.length) return;
    [fileStore[currentFileIndex], fileStore[newIndex]] = [fileStore[newIndex], fileStore[currentFileIndex]];
    [contentStore[currentFileIndex], contentStore[newIndex]] = [contentStore[newIndex], contentStore[currentFileIndex]];
    const rows = Array.from(runlistContainer.querySelectorAll('.runlist-row'));
    [rows[currentFileIndex], rows[newIndex]] = [rows[newIndex], rows[currentFileIndex]];
    runlistContainer.innerHTML = '';
    rows.forEach((row, i) => {
        row.dataset.index = i;
        row.querySelector('.file-name').textContent = fileStore[i].name;
        row.onclick = () => loadScriptToEditor(i);
        runlistContainer.appendChild(row);
    });
    currentFileIndex = newIndex;
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const activeRow = runlistContainer.querySelector(`.runlist-row[data-index="${newIndex}"]`);
    if (activeRow) activeRow.classList.add('active');
}

if (btnMoveFileUp) btnMoveFileUp.onclick = () => moveFileInRunlist('up');
if (btnMoveFileDown) btnMoveFileDown.onclick = () => moveFileInRunlist('down');

function addFileToRunlist(file) {
    console.log(`Adding to runlist: ${file.name}`);
    const index = fileStore.length;
    fileStore.push(file);
    contentStore.push(""); 

    const row = document.createElement('div');
    row.className = 'runlist-row';
    row.dataset.index = index;
    row.innerHTML = `<span class="file-name">${file.name}</span>`;

    row.onclick = () => {
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
function updateBookmarkSidebar() {
    const list = document.querySelector('.bookmark-list');
    if (!list) return;
    const anchors = teleprompterText.querySelectorAll('[id]');
    list.innerHTML = '';
    anchors.forEach((el, i) => {
        if (!el.id || el.id === 'teleprompter-text') return;
        const item = document.createElement('div');
        item.className = 'bookmark-item';
        item.textContent = el.id || `Bookmark ${i + 1}`;
        item.onclick = () => {
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
        };
        list.appendChild(item);
    });
}

function loadScriptToEditor(index) {
    console.log(`Attempting to load index: ${index}`);
    if (fileStore[index] === null) return;

    if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
        const html = teleprompterText.innerHTML.trim();
        contentStore[currentFileIndex] = (html === '<br>' || html === '') ? '' : html;
    }

    rowColorsCache = [];
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const row = document.querySelector(`.runlist-row[data-index="${index}"]`);
    if (row) row.classList.add('active');

    currentFileIndex = index;
    const content = contentStore[index];
    teleprompterText.innerHTML = (content === '' || !content) ? '<br>' : content;
    teleprompterText.focus();
    requestAnimationFrame(() => {
        refreshSelectFontTarget();
        syncColumnWidths();
        updateBookmarkSidebar();
    });

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

extendMonitorButton.onclick = async () => {
    const isExtended = document.body.classList.contains('broadcasting') && mirrorWindow && !mirrorWindow.closed;

    if (isExtended) {
        const maxScroll = teleprompterView.scrollHeight - teleprompterView.clientHeight;
        const scrollRatio = maxScroll > 0 ? teleprompterView.scrollTop / maxScroll : 0;

        mirrorWindow.close();
        mirrorWindow = null;
        measuredRowHeights = [];
        document.body.classList.remove('broadcasting');
        applyBroadcastingVisibility();
        syncColumnWidths();

        const restoreUnextend = () => {
            const newMax = teleprompterView.scrollHeight - teleprompterView.clientHeight;
            teleprompterView.scrollTop = Math.round(scrollRatio * Math.max(0, newMax));
        };
        restoreUnextend();
        requestAnimationFrame(restoreUnextend);
        setTimeout(restoreUnextend, 100);
        return;
    }

    const maxScroll = teleprompterView.scrollHeight - teleprompterView.clientHeight;
    const scrollRatio = maxScroll > 0 ? teleprompterView.scrollTop / maxScroll : 0;

    function restoreScrollPosition() {
        const newMax = teleprompterView.scrollHeight - teleprompterView.clientHeight;
        teleprompterView.scrollTop = Math.round(scrollRatio * Math.max(0, newMax));
    }

    overflowReportCount = 0;
    measureRowHeightsFromContent();
    document.body.classList.add('broadcasting');
    applyBroadcastingVisibility();
    syncColumnWidths();
    restoreScrollPosition();

    try {
        let currentScreen = null;
        let secondaryScreen = null;

        if ('getScreenDetails' in window) {
            const screenDetails = await window.getScreenDetails();
            currentScreen = screenDetails.currentScreen;
            secondaryScreen = screenDetails.screens.find(s => s !== currentScreen) || null;
        }

        if (currentScreen) {
            window.moveTo(currentScreen.availLeft, currentScreen.availTop);
            window.resizeTo(currentScreen.availWidth, currentScreen.availHeight);
        } else {
            window.moveTo(0, 0);
            window.resizeTo(window.screen.availWidth, window.screen.availHeight);
        }

        restoreScrollPosition();

        const mirrorLeft = secondaryScreen ? secondaryScreen.availLeft : (window.screen.availWidth || 0);
        const mirrorTop = secondaryScreen ? secondaryScreen.availTop : 0;
        const mirrorWidth = secondaryScreen ? secondaryScreen.availWidth : 800;
        const mirrorHeight = secondaryScreen ? secondaryScreen.availHeight : 600;
        const specs = `left=${mirrorLeft},top=${mirrorTop},width=${mirrorWidth},height=${mirrorHeight}`;

        mirrorWindow = window.open('mirror.html', 'TeleprompterMirror', specs);
        if (mirrorWindow) {
            const sendWhenReady = () => {
                syncColumnWidths();
                refreshMirrorData();
                syncMirrorStyles();
                const attemptRestore = () => {
                    restoreScrollPosition();
                    syncMirrorByPixels();
                };
                requestAnimationFrame(attemptRestore);
                requestAnimationFrame(() => requestAnimationFrame(attemptRestore));
                setTimeout(attemptRestore, 100);
            };
            mirrorWindow.addEventListener('load', sendWhenReady);
            setTimeout(sendWhenReady, 800);
        }
    } catch (err) {
        console.error("Monitor Extension Failed:", err);
    }
};
    
// =========================================
// 7. SCROLLING ENGINE (PIXEL-PERFECT)
// =========================================

function getMirrorStylePayload() {
    const mainStyle = window.getComputedStyle(teleprompterText);
    const cellSample = teleprompterText.querySelector('.cell-locker');
    const colSample = teleprompterText.querySelector('.script-column');
    const rowSample = teleprompterText.querySelector('.script-row-wrapper');
    const cellStyle = cellSample ? window.getComputedStyle(cellSample) : (colSample ? window.getComputedStyle(colSample) : null);
    const rowHeightPx = rowSample ? rowSample.getBoundingClientRect().height : 0;
    const rowHeight = rowHeightPx > 0 ? rowHeightPx : null;

    let contentWidthPx = lastColumnWidthPx;
    if (contentWidthPx == null || contentWidthPx <= 0) {
        const rows = teleprompterText.querySelectorAll('.script-row-wrapper');
        if (rows.length > 0) {
            const firstRow = rows[0];
            const cols = firstRow.querySelectorAll('.script-column');
            const secondCol = cols[1];
            if (secondCol) {
                const rect = secondCol.getBoundingClientRect();
                if (rect.width > 0) contentWidthPx = rect.width;
            }
            if ((contentWidthPx == null || contentWidthPx <= 0) && cols.length > 1) {
                const availableWidth = teleprompterText.clientWidth;
                const firstCol = cols[0];
                const firstColWidth = firstCol ? firstCol.getBoundingClientRect().width : 0;
                contentWidthPx = (availableWidth - firstColWidth) / (cols.length - 1);
            }
        }
    }

    return {
        fontSize: mainStyle.fontSize,
        fontFamily: mainStyle.fontFamily,
        lineHeight: mainStyle.lineHeight,
        paddingTop: cellStyle ? cellStyle.paddingTop : null,
        paddingRight: cellStyle ? cellStyle.paddingRight : null,
        paddingBottom: cellStyle ? cellStyle.paddingBottom : null,
        paddingLeft: cellStyle ? cellStyle.paddingLeft : null,
        rowHeight: rowHeight ? `${rowHeight}px` : null,
        contentWidth: (contentWidthPx && contentWidthPx > 50) ? `${contentWidthPx}px` : null
    };
}

function syncMirrorStyles() {
    if (!mirrorWindow || mirrorWindow.closed) return;
    mirrorWindow.postMessage({ 
        type: 'syncStyleLite', 
        style: getMirrorStylePayload()
    }, '*');
}

window.addEventListener('resize', () => {
    syncColumnWidths();
    syncMirrorStyles();
});

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
