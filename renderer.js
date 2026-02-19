document.addEventListener('DOMContentLoaded', function() {
    // =========================================
    // 1. ELEMENT SELECTORS
    // =========================================
    const openFileButton = document.getElementById('btn-open-file');
    const toggleRunlistButton = document.getElementById('btn-toggle-runlist');
    const fileOpener = document.getElementById('file-opener');
    const runlistContainer = document.querySelector('.runlist-files');
    let runlistDropIndicator = null;
    function getRunlistDropIndicator() {
        if (!runlistDropIndicator && runlistContainer) {
            runlistDropIndicator = document.createElement('div');
            runlistDropIndicator.className = 'runlist-drop-indicator';
            runlistDropIndicator.setAttribute('aria-hidden', 'true');
        }
        return runlistDropIndicator;
    }
    function hideRunlistDropIndicator() {
        if (runlistDropIndicator && runlistDropIndicator.parentNode) runlistDropIndicator.remove();
        runlistContainer.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('runlist-drop-target'));
    }
    let runlistDraggingIndex = null;
    let runlistDraggingRow = null;
    let runlistJustDragged = false;
    function onRunlistMouseMove(e) {
        if (runlistDraggingIndex == null || !runlistContainer) return;
        const el = document.elementFromPoint(e.clientX, e.clientY);
        const row = el && el.closest ? el.closest('.runlist-row') : null;
        if (!row || !runlistContainer.contains(row)) {
            hideRunlistDropIndicator();
            return;
        }
        const rowIndex = parseInt(row.dataset.index, 10);
        const rect = row.getBoundingClientRect();
        const topHalf = rect.height > 0 && (e.clientY - rect.top) < rect.height / 2;
        const insertIndex = topHalf ? rowIndex : rowIndex + 1;
        const indicator = getRunlistDropIndicator();
        indicator.dataset.insertIndex = String(insertIndex);
        const rows = runlistContainer.querySelectorAll('.runlist-row');
        if (insertIndex >= rows.length) runlistContainer.appendChild(indicator);
        else runlistContainer.insertBefore(indicator, rows[insertIndex]);
    }
    function onRunlistMouseUp(e) {
        if (runlistDraggingIndex == null) return;
        const indicator = getRunlistDropIndicator();
        const toIndex = parseInt(indicator.dataset.insertIndex, 10);
        hideRunlistDropIndicator();
        if (!isNaN(toIndex) && toIndex !== runlistDraggingIndex) {
            moveRunlistRow(runlistDraggingIndex, toIndex);
            runlistJustDragged = true;
            setTimeout(() => { runlistJustDragged = false; }, 100);
        }
        if (runlistDraggingRow) runlistDraggingRow.classList.remove('runlist-dragging');
        runlistDraggingIndex = null;
        runlistDraggingRow = null;
        document.removeEventListener('mousemove', onRunlistMouseMove, true);
        document.removeEventListener('mouseup', onRunlistMouseUp, true);
    }
    if (runlistContainer) {
        runlistContainer.addEventListener('dragleave', (e) => {
            if (!runlistContainer.contains(e.relatedTarget)) hideRunlistDropIndicator();
        });
    }
    const teleprompterText = document.getElementById('teleprompter-text');
    const runlistPanel = document.getElementById('runlist-panel');
    const resizer = document.getElementById('resizer');
    const teleprompterView = document.getElementById('teleprompter-view');
    const teleprompterStage = document.querySelector('.teleprompter-stage');
    const indicatorWrapper = document.getElementById('indicator-wrapper');
    const indicatorTriangle = document.getElementById('indicator-triangle');
    const saveButton = document.getElementById('btn-save');
    const undoButton = document.getElementById('btn-undo');
    const redoButton = document.getElementById('btn-redo');
    const fontFamilySelect = document.getElementById('font-family-select');
    const fontSizeSelect = document.getElementById('font-size-select');
    const selectFontTarget = document.getElementById('select-font-target');
    const btnRemoveFontTarget = document.getElementById('btn-remove-font-target');
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
    const settingsOverlay = document.getElementById('settings-overlay');
    const f2Prompt = document.getElementById('f2-prompt');
    const newBookmarkButton = document.getElementById('btn-new-bookmark');
    const prevBookmarkButton = document.getElementById('btn-prev-bookmark');
    const nextBookmarkButton = document.getElementById('btn-next-bookmark');
    const invertScrollCheckbox = document.getElementById('invert-scroll-checkbox');
    const secondMonitorPositionRadios = document.querySelectorAll('input[name="secondMonitorPosition"]');

    // Modal Elements
    const btnMoveFileUp = document.getElementById('btn-move-file-up');
    const btnMoveFileDown = document.getElementById('btn-move-file-down');
    const speedSlider = document.getElementById('speed-slider');
    const speedValueEl = document.getElementById('speed-value');

    const FONT_FAMILY_STORAGE_KEY = 'teleprompter_fontFamily';
    const FONT_SIZE_STORAGE_KEY = 'teleprompter_fontSize';
    try {
        const savedFamily = localStorage.getItem(FONT_FAMILY_STORAGE_KEY);
        if (savedFamily && fontFamilySelect && Array.from(fontFamilySelect.options).some(o => o.value === savedFamily)) {
            fontFamilySelect.value = savedFamily;
        }
        const savedSize = localStorage.getItem(FONT_SIZE_STORAGE_KEY);
        if (savedSize && fontSizeSelect && Array.from(fontSizeSelect.options).some(o => o.value === savedSize)) {
            fontSizeSelect.value = savedSize;
        }
    } catch (_) {}

    // =========================================
    // 2. IMMEDIATE UI SETUP (CRITICAL: Must be AFTER selectors)
    // =========================================
    // This ensures the main screen is black/white immediately on load
    teleprompterView.style.backgroundColor = "#000000";
    teleprompterText.style.color = "#ffffff";
    teleprompterText.style.fontFamily = fontFamilySelect?.value || "Arial";
    teleprompterText.style.fontSize = (fontSizeSelect?.value || "80") + "px";
    teleprompterView.focus();

    function clearPlaceholderIfActive() {
        if (teleprompterText.dataset.placeholder !== 'true') return;
        teleprompterText.innerHTML = '<br>';
        delete teleprompterText.dataset.placeholder;
        teleprompterText.focus();
        const sel = window.getSelection();
        const range = document.createRange();
        range.setStart(teleprompterText, 0);
        range.collapse(true);
        sel.removeAllRanges();
        sel.addRange(range);
    }
    teleprompterText.addEventListener('focus', clearPlaceholderIfActive);
    teleprompterView.addEventListener('click', () => {
        if (teleprompterText.dataset.placeholder === 'true') {
            clearPlaceholderIfActive();
        }
    });

    // =========================================
    // 2. APP STATE
    // =========================================
    const MAX_COLUMNS = 3;
    let fileStore = [];         // Stores File objects
    let contentStore = [];      // Stores HTML strings
    let currentFileIndex = -1;
    let isRunlistVisible = false;
    let savedSelection = null;
    let mirrorWindow = null;
    let mirrorCloseCheckInterval = null;
    /** When set, opener sends this position to mirror via postMessage so mirror can move itself (works when opener moveTo is blocked on PC). */
    let pendingMirrorPosition = null;
    /** When true, we are extended (mirror open) but main window shows all columns so user can edit without closing the mirror. */
    let broadcastEditMode = false;
    let isTeleprompting = false;
    let isPaused = false;
    let spacebarControlsPlay = true; /* false after any mouse click so only Option+A starts; Option+A resets so spacebar works again */
    let isMouseControlActive = true;
    let scrollSpeed = 0;
    let scrollAccum = 0; /* Fractional accumulator so small speeds (0.1, 0.2) actually scroll */
    let scrollInterval = null;
    let lastActiveSpeed = 1;
    let isInvertScroll = false;
    let scrollSensitivity = 1;
    let isMirrorActive = false;
    let mirrorScrollOffset = -150;
    let animationFrameId = null;
    
    // Excel Import Temp State
    let pendingWorkbook = null;
    let pendingFileIndex = -1;
    /** Per-file: number of script columns (0 = non-table or not yet set). */
    let fileColumnCount = [];
    /** Per-file: which columns are visible; fileColumnVisibility[i][j] = true to show column j. */
    let fileColumnVisibility = [];
    /** Per-file: true if first column is numeric (show as "ID" in runlist). */
    let fileFirstColIsId = [];

    let lastColumnWidthPx = null;
    let col2WidthPx = null;
    let col3WidthPx = null;
    let extendedFixedWidth = null;  /* When broadcasting: fixed table width so main/mirror stay aligned */
    let extendedWindowWidth = null; /* 5:4 outer window size for main and mirror */
    let extendedWindowHeight = null;
    let charCountsPerRow = [];
    let rowColorsCache = []; /* Stores row color classes from unextended view so colors don't change when extended */
    let rowFont12Cache = []; /* Rows containing |v use font 12, cache for extended view */
    let measuredRowHeights = []; /* Actual measured heights from DOM to prevent text overlap when extended */
    let mirrorReportedRowHeights = []; /* Mirror's natural row heights (col 3 only) – used when reassessing so we don't force mirror too tall */
    let selectedFontTarget = null; /* { fontFamily, fontSize, color?, backgroundColor? } when user selects from Select dropdown */
    let savedFontSelections = []; /* Array of cloned ranges for multi-select */
    let lastFontChangeSource = null; /* 'family' | 'size' - which dropdown triggered the change */
    let modifierHeldForAdd = false; /* Cmd/Ctrl held during mousedown = add to multi-selection */
    let isOverviewMode = false;
    let savedFontSizeBeforeOverview = null;
    /** When true, refreshMirrorData should not overwrite measuredRowHeights (used when we re-enter Aa after extending from overview). */
    let preserveRowHeightsAfterExtend = false;
    const undoStack = [];
    const redoStack = [];
    const MAX_UNDO = 50;

    function saveFontSelectionFromEditor(addToMulti) {
        const sel = window.getSelection();
        if (!sel.rangeCount) {
            if (!addToMulti) savedFontSelections = [];
            updateMultiSelectState();
            return;
        }
        const r = sel.getRangeAt(0);
        if (r.collapsed) {
            if (!addToMulti) savedFontSelections = [];
            updateMultiSelectState();
            return;
        }
        const root = r.commonAncestorContainer;
        const el = root.nodeType === Node.TEXT_NODE ? root.parentElement : root;
        if (!el || !teleprompterText.contains(el)) {
            if (!addToMulti) savedFontSelections = [];
            updateMultiSelectState();
            return;
        }
        try {
            const cloned = r.cloneRange();
            if (addToMulti) {
                savedFontSelections.push(cloned);
            } else {
                savedFontSelections = [cloned];
            }
            updateMultiSelectState();
        } catch (_) {
            if (!addToMulti) savedFontSelections = [];
            updateMultiSelectState();
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
            flattenRedundantSpans();
        }
    }

    function colorsMatch(target, computedColor, computedBg) {
        const tColor = normalizeColorForKey(target.color);
        const tBg = normalizeColorForKey(target.backgroundColor);
        const c = normalizeColorForKey(computedColor);
        const b = normalizeColorForKey(computedBg);
        if (tColor && c !== tColor) return false;
        if (tBg && b !== tBg) return false;
        return true;
    }

    function unwrapFontFromMatchingTarget(target) {
        if (!target) return;
        const { fontFamily: targetFont, fontSize: targetSize } = target;
        const normalizeFont = (s) => (s || '').split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
        const sizeNum = (s) => parseFloat(String(s || '').trim()) || 0;
        const isStructural = (el) => !el || el === teleprompterText || el?.classList?.contains('script-row-wrapper') || el?.classList?.contains('script-column') || el?.classList?.contains('cell-locker') || el?.classList?.contains('cell-content') || el?.classList?.contains('script-container');
        const tFont = normalizeFont(targetFont);
        const tSizeNum = sizeNum(targetSize);
        if (!tFont || !tSizeNum) return;
        const matchingTextNodes = [];
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            if (!node.textContent.trim()) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            const pFont = normalizeFont(style.fontFamily);
            const pSizeNum = sizeNum(style.fontSize);
            const { color: effColor, backgroundColor: effBg } = getEffectiveColorsFromNode(node);
            if (tFont && tSizeNum && pFont === tFont && Math.abs(pSizeNum - tSizeNum) < 0.5 && colorsMatch(target, effColor, effBg)) {
                matchingTextNodes.push(node);
            }
        }
        const spansToUnwrap = new Set();
        matchingTextNodes.forEach(textNode => {
            let el = textNode.parentElement;
            while (el && !isStructural(el)) {
                if (el.tagName === 'SPAN' || el.tagName === 'B' || el.tagName === 'I' || el.tagName === 'U') {
                    spansToUnwrap.add(el);
                }
                el = el.parentElement;
            }
        });
        const sortedByDepth = Array.from(spansToUnwrap).sort((a, b) => {
            const depth = (el) => { let d = 0; let n = el; while (n && n !== teleprompterText) { d++; n = n.parentElement; } return d; };
            return depth(b) - depth(a);
        });
        sortedByDepth.forEach(span => {
            const parent = span.parentElement;
            if (!parent) return;
            if (isStructural(span)) return;
            while (span.firstChild) parent.insertBefore(span.firstChild, span);
            span.remove();
        });
        flattenRedundantSpans();
        refreshSelectFontTarget();
    }

    function applyFontToMatchingTarget(target, newFontVal, newSizeVal, resetLineHeight) {
        if (!target || (!newFontVal && !newSizeVal)) return;
        const { fontFamily: targetFont, fontSize: targetSize } = target;
        const preserveSize = lastFontChangeSource === 'family';
        const preserveFamily = lastFontChangeSource === 'size';
        const newFont = preserveFamily ? targetFont : (newFontVal || targetFont);
        const newSize = preserveSize ? targetSize : ((newSizeVal || '') + (newSizeVal ? 'px' : ''));
        const rootLineHeight = resetLineHeight ? (window.getComputedStyle(teleprompterText).lineHeight || '1.4') : null;
        const normalizeFont = (s) => (s || '').split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
        const normalizeSize = (s) => String(s || '').trim();
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        const toWrap = [];
        let node;
        while ((node = walker.nextNode())) {
            if (!node.textContent.trim()) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            const pFont = normalizeFont(style.fontFamily);
            const pSize = normalizeSize(style.fontSize);
            const tFont = normalizeFont(targetFont);
            const tSize = normalizeSize(targetSize);
            const { color: effColor, backgroundColor: effBg } = getEffectiveColorsFromNode(node);
            if (tFont && tSize && pFont === tFont && pSize === tSize && colorsMatch(target, effColor, effBg)) {
                toWrap.push(node);
            }
        }
        const isStructuralNode = (el) => el === teleprompterText || el?.classList?.contains('script-row-wrapper') || el?.classList?.contains('script-column') || el?.classList?.contains('cell-locker') || el?.classList?.contains('cell-content');
        const BLOCK_TAGS = ['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'];
        const unwrapBlockAncestors = (child) => {
            if (!child || !child.parentElement) return;
            let parent = child.parentElement;
            while (parent && !isStructuralNode(parent)) {
                const next = parent.parentElement;
                if (parent.childNodes.length === 1 && (parent.tagName === 'SPAN' || BLOCK_TAGS.includes(parent.tagName))) {
                    next?.insertBefore(child, parent);
                    parent.remove();
                    parent = next;
                } else {
                    break;
                }
            }
        };
        toWrap.forEach(textNode => {
            const span = document.createElement('span');
            span.style.display = 'inline';
            if (newFont) span.style.fontFamily = newFont;
            if (newSize) span.style.fontSize = newSize.endsWith('px') ? newSize : newSize + 'px';
            if (rootLineHeight) span.style.lineHeight = rootLineHeight;
            textNode.parentNode.insertBefore(span, textNode);
            span.appendChild(textNode);
            if (resetLineHeight) {
                unwrapBlockAncestors(span);
            }
        });
        flattenRedundantSpans();
    }

    function flattenRedundantSpans() {
        const isStructural = (el) => el === teleprompterText || el?.classList?.contains('script-row-wrapper') || el?.classList?.contains('script-column') || el?.classList?.contains('cell-locker') || el?.classList?.contains('cell-content') || el?.classList?.contains('bookmark-cursor-dot') || el?.classList?.contains('bookmark-cursor-dot-nobreak') || el?.classList?.contains('bookmark-cursor-dot-wrap') || el?.classList?.contains('color-span-inline') || el?.hasAttribute?.('data-bookmark-inline');
        const BLOCK_TAGS = ['DIV', 'P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'];
        let changed = true;
        while (changed) {
            changed = false;
            const all = Array.from(teleprompterText.querySelectorAll('span, div, p'));
            all.forEach(el => {
                if (el?.classList?.contains('bookmark-cursor-dot') || el?.classList?.contains('bookmark-cursor-dot-wrap') || el?.classList?.contains('color-span-inline') || isStructural(el.parentElement)) return;
                if (el.childNodes.length === 1) {
                    const child = el.firstChild;
                    if (child.nodeType === Node.TEXT_NODE) {
                        const parent = el.parentElement;
                        if (parent?.tagName === 'SPAN' && !isStructural(parent) && parent.childNodes.length === 1) {
                            if (el.tagName === 'SPAN') Object.assign(parent.style, el.style);
                            parent.insertBefore(child, el);
                            el.remove();
                            changed = true;
                        }
                    } else if (child.nodeType === Node.ELEMENT_NODE && (child.tagName === 'SPAN' || BLOCK_TAGS.includes(child.tagName))) {
                        if (el.tagName === 'SPAN' && child.tagName === 'SPAN') Object.assign(el.style, child.style);
                        while (child.firstChild) el.insertBefore(child.firstChild, child);
                        child.remove();
                        changed = true;
                    }
                } else if (BLOCK_TAGS.includes(el.tagName) && el.childNodes.length === 1) {
                    const inner = el.firstChild;
                    const grandparent = el.parentElement;
                    if (grandparent && !isStructural(grandparent)) {
                        if (inner.nodeType === Node.ELEMENT_NODE && inner.tagName === 'SPAN') {
                            grandparent.insertBefore(inner, el);
                        } else {
                            const span = document.createElement('span');
                            span.style.display = 'inline';
                            while (el.firstChild) span.appendChild(el.firstChild);
                            grandparent.insertBefore(span, el);
                        }
                        el.remove();
                        changed = true;
                    }
                }
            });
        }
        teleprompterText.querySelectorAll('span').forEach(span => {
            if (!span.textContent.trim() && span.parentElement && !isStructural(span.parentElement)) span.remove();
        });
    }

    function normalizeColorForKey(c) {
        if (!c) return '';
        const t = c.trim().replace(/\s+/g, ' ');
        if (t === 'transparent' || t === 'rgba(0, 0, 0, 0)' || t === 'rgba(0,0,0,0)') return '';
        const m = t.match(/rgba\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)/);
        if (m) return 'rgba(' + m[1] + ',' + m[2] + ',' + m[3] + ',' + m[4] + ')';
        const m2 = t.match(/rgb\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
        if (m2) return 'rgb(' + m2[1] + ',' + m2[2] + ',' + m2[3] + ')';
        const hex = t.match(/^#([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$/);
        if (hex) {
            const h = hex[1];
            const r = h.length === 3 ? parseInt(h[0] + h[0], 16) : parseInt(h.slice(0, 2), 16);
            const g = h.length === 3 ? parseInt(h[1] + h[1], 16) : parseInt(h.slice(2, 4), 16);
            const b = h.length === 3 ? parseInt(h[2] + h[2], 16) : parseInt(h.slice(4, 6), 16);
            return 'rgb(' + r + ',' + g + ',' + b + ')';
        }
        return t;
    }

    function getEffectiveColorsFromNode(textNode) {
        let el = textNode.parentElement;
        let color = '';
        let backgroundColor = '';
        const isStructural = (n) => n === teleprompterText || n?.classList?.contains('script-row-wrapper') || n?.classList?.contains('script-column') || n?.classList?.contains('cell-locker') || n?.classList?.contains('cell-content');
        while (el && !isStructural(el)) {
            const style = window.getComputedStyle(el);
            const c = (style.color || '').trim();
            const b = (style.backgroundColor || '').trim();
            if (c && !color) color = c;
            if (isOpaqueColor(b) && !backgroundColor) backgroundColor = b;
            el = el.parentElement;
        }
        if (!color && el) color = (window.getComputedStyle(teleprompterText).color || '').trim();
        return { color, backgroundColor };
    }

    function getEffectiveColorsFromElement(el) {
        const textNode = el.nodeType === Node.TEXT_NODE ? el : Array.from(el.childNodes).find(n => n.nodeType === Node.TEXT_NODE);
        if (textNode) return getEffectiveColorsFromNode(textNode);
        const style = window.getComputedStyle(el);
        return { color: (style.color || '').trim(), backgroundColor: (style.backgroundColor || '').trim() };
    }

    function collectUniqueFontSizes() {
        const rootColor = (window.getComputedStyle(teleprompterText).color || '').trim();
        const rootColorNorm = normalizeColorForKey(rootColor);
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
            const { color, backgroundColor } = getEffectiveColorsFromNode(node);
            const colorNorm = normalizeColorForKey(color);
            const bgNorm = normalizeColorForKey(backgroundColor);
            const key = fontFamily + '|' + fontSize + '|' + colorNorm + '|' + bgNorm;
            if (!pairs.has(key)) pairs.set(key, { fontFamily, fontSize, color: color || (style.color || '').trim(), backgroundColor: backgroundColor || (style.backgroundColor || '').trim() });
        }
        return Array.from(pairs.values()).sort((a, b) => {
            const isUntouched = (x) => {
                const cNorm = normalizeColorForKey(x.color);
                const bgNorm = normalizeColorForKey(x.backgroundColor);
                return !isOpaqueColor(x.backgroundColor) && (cNorm === rootColorNorm || !isOpaqueColor(x.color));
            };
            const aU = isUntouched(a) ? 0 : 1;
            const bU = isUntouched(b) ? 0 : 1;
            if (aU !== bU) return aU - bU;
            const c = (a.fontFamily || '').localeCompare(b.fontFamily || '');
            if (c !== 0) return c;
            const d = (parseFloat(a.fontSize) || 0) - (parseFloat(b.fontSize) || 0);
            if (d !== 0) return d;
            return (a.color || '').localeCompare(b.color || '') || (a.backgroundColor || '').localeCompare(b.backgroundColor || '');
        });
    }

    function getMostCommonFont() {
        const counts = new Map();
        const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            const len = (node.textContent || '').trim().length;
            if (len === 0) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const style = window.getComputedStyle(parent);
            const fontFamily = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
            const fontSize = style.fontSize || '';
            const key = fontFamily + '|' + fontSize;
            counts.set(key, (counts.get(key) || 0) + len);
        }
        let best = null;
        let bestCount = 0;
        counts.forEach((count, key) => {
            if (count > bestCount) {
                bestCount = count;
                const [fontFamily, fontSize] = key.split('|');
                best = { fontFamily, fontSize };
            }
        });
        return best;
    }

    function isOpaqueColor(c) {
        if (!c || c === 'transparent') return false;
        const rgbaMatch = c.match(/rgba\s*\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*([\d.]+)\s*\)/);
        if (rgbaMatch && parseFloat(rgbaMatch[1]) === 0) return false;
        return true;
    }

    /** Get format (fontFamily, fontSize, color, backgroundColor) at current selection or caret for Select menu sync. */
    function getFormatAtCaretOrSelection() {
        const sel = window.getSelection();
        if (!sel || sel.rangeCount === 0) return null;
        const range = sel.getRangeAt(0);
        if (!teleprompterText.contains(range.startContainer)) return null;
        let node = range.startContainer;
        if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
        if (!node) return null;
        const style = window.getComputedStyle(node);
        const fontFamily = (style.fontFamily || '').split(',')[0].trim().replace(/^["']|["']$/g, '');
        const fontSize = style.fontSize || '';
        const { color, backgroundColor } = getEffectiveColorsFromElement(node);
        return { fontFamily, fontSize, color: (color || '').trim(), backgroundColor: (backgroundColor || '').trim() };
    }

    function refreshSelectFontTarget() {
        if (!selectFontTarget) return;
        const trigger = document.getElementById('font-target-trigger');
        const listEl = document.getElementById('font-target-list');
        if (!trigger || !listEl) return;
        const current = selectFontTarget.value;
        const opts = collectUniqueFontSizes();
        selectFontTarget.innerHTML = '<option value="">Select</option>';
        listEl.innerHTML = '';
        trigger.textContent = 'Select';
        const clearItem = document.createElement('div');
        clearItem.className = 'font-target-item';
        clearItem.textContent = 'Select';
        clearItem.addEventListener('click', () => {
            selectFontTarget.value = '';
            trigger.textContent = 'Select';
            listEl.classList.remove('open');
            selectedFontTarget = null;
            selectFontTarget.dispatchEvent(new Event('change'));
        });
        listEl.appendChild(clearItem);
        opts.forEach(({ fontFamily, fontSize, color, backgroundColor }) => {
            const sizeNum = parseInt(fontSize, 10) || fontSize;
            const label = `${fontFamily} ${sizeNum}`;
            const hasFontColor = isOpaqueColor(color);
            const hasBgColor = isOpaqueColor(backgroundColor);
            const val = JSON.stringify({ fontFamily, fontSize, color: hasFontColor ? color : null, backgroundColor: hasBgColor ? backgroundColor : null });
            const opt = document.createElement('option');
            opt.value = val;
            opt.textContent = label;
            selectFontTarget.appendChild(opt);
            const item = document.createElement('div');
            item.className = 'font-target-item';
            item.dataset.value = val;
            const labelSpan = document.createElement('span');
            labelSpan.className = 'font-target-label';
            labelSpan.textContent = label;
            item.appendChild(labelSpan);
            if (hasFontColor || hasBgColor) {
                const swatches = document.createElement('span');
                swatches.className = 'font-target-swatches';
                if (hasBgColor) {
                    const sq = document.createElement('span');
                    sq.className = 'font-target-swatch font-target-swatch-bg';
                    sq.style.backgroundColor = backgroundColor;
                    sq.title = 'Background color';
                    swatches.appendChild(sq);
                }
                if (hasFontColor) {
                    const sq = document.createElement('span');
                    sq.className = 'font-target-swatch font-target-swatch-text';
                    sq.style.backgroundColor = color;
                    sq.title = 'Text color';
                    swatches.appendChild(sq);
                }
                item.appendChild(swatches);
            }
            if (current === val) {
                opt.selected = true;
                trigger.innerHTML = '';
                const tLabel = document.createElement('span');
                tLabel.textContent = label;
                trigger.appendChild(tLabel);
                if (hasFontColor || hasBgColor) {
                    const sw = document.createElement('span');
                    sw.className = 'font-target-swatches';
                    if (hasBgColor) {
                        const s = document.createElement('span');
                        s.className = 'font-target-swatch font-target-swatch-bg';
                        s.style.backgroundColor = backgroundColor;
                        sw.appendChild(s);
                    }
                    if (hasFontColor) {
                        const s = document.createElement('span');
                        s.className = 'font-target-swatch font-target-swatch-text';
                        s.style.backgroundColor = color;
                        sw.appendChild(s);
                    }
                    trigger.appendChild(sw);
                }
            }
            item.addEventListener('click', () => {
                selectFontTarget.value = val;
                trigger.innerHTML = '';
                const tLabel = document.createElement('span');
                tLabel.textContent = label;
                trigger.appendChild(tLabel);
                if (hasFontColor || hasBgColor) {
                    const sw = document.createElement('span');
                    sw.className = 'font-target-swatches';
                    if (hasBgColor) {
                        const s = document.createElement('span');
                        s.className = 'font-target-swatch font-target-swatch-bg';
                        s.style.backgroundColor = backgroundColor;
                        sw.appendChild(s);
                    }
                    if (hasFontColor) {
                        const s = document.createElement('span');
                        s.className = 'font-target-swatch font-target-swatch-text';
                        s.style.backgroundColor = color;
                        sw.appendChild(s);
                    }
                    trigger.appendChild(sw);
                }
                listEl.classList.remove('open');
                selectFontTarget.dispatchEvent(new Event('change'));
            });
            listEl.appendChild(item);
        });
        if (!selectFontTarget.value) {
            const atCaret = getFormatAtCaretOrSelection();
            if (atCaret && opts.length > 0) {
                const hasFontColor = isOpaqueColor(atCaret.color);
                const hasBgColor = isOpaqueColor(atCaret.backgroundColor);
                const wantVal = JSON.stringify({ fontFamily: atCaret.fontFamily, fontSize: atCaret.fontSize, color: hasFontColor ? atCaret.color : null, backgroundColor: hasBgColor ? atCaret.backgroundColor : null });
                let match = opts.find(({ fontFamily, fontSize, color, backgroundColor }) => {
                    const hasC = isOpaqueColor(color);
                    const hasB = isOpaqueColor(backgroundColor);
                    const v = JSON.stringify({ fontFamily, fontSize, color: hasC ? color : null, backgroundColor: hasB ? backgroundColor : null });
                    return v === wantVal;
                });
                if (!match && normalizeColorForKey) {
                    match = opts.find(({ fontFamily, fontSize, color, backgroundColor }) => {
                        return fontFamily === atCaret.fontFamily && fontSize === atCaret.fontSize &&
                            normalizeColorForKey(color || '') === normalizeColorForKey(atCaret.color || '') &&
                            normalizeColorForKey(backgroundColor || '') === normalizeColorForKey(atCaret.backgroundColor || '');
                    });
                }
                if (match) {
                    const { fontFamily, fontSize, color, backgroundColor } = match;
                    const sizeNum = parseInt(fontSize, 10) || fontSize;
                    const label = `${fontFamily} ${sizeNum}`;
                    const val = JSON.stringify({ fontFamily, fontSize, color: isOpaqueColor(color) ? color : null, backgroundColor: isOpaqueColor(backgroundColor) ? backgroundColor : null });
                    selectFontTarget.value = val;
                    trigger.innerHTML = '';
                    const tLabel = document.createElement('span');
                    tLabel.textContent = label;
                    trigger.appendChild(tLabel);
                    if (isOpaqueColor(color) || isOpaqueColor(backgroundColor)) {
                        const sw = document.createElement('span');
                        sw.className = 'font-target-swatches';
                        if (isOpaqueColor(backgroundColor)) {
                            const s = document.createElement('span');
                            s.className = 'font-target-swatch font-target-swatch-bg';
                            s.style.backgroundColor = backgroundColor;
                            sw.appendChild(s);
                        }
                        if (isOpaqueColor(color)) {
                            const s = document.createElement('span');
                            s.className = 'font-target-swatch font-target-swatch-text';
                            s.style.backgroundColor = color;
                            sw.appendChild(s);
                        }
                        trigger.appendChild(sw);
                    }
                } else {
                    trigger.textContent = 'Select';
                }
            } else {
                trigger.textContent = 'Select';
            }
        }
        try {
            selectedFontTarget = selectFontTarget.value ? JSON.parse(selectFontTarget.value) : null;
        } catch (_) {
            selectedFontTarget = null;
        }
    }

    function removeRedundantBrInContainer(container) {
        if (!container || !teleprompterText.contains(container)) return;
        const toRemove = [];
        const scan = (node) => {
            if (!node) return;
            if (node.nodeType === Node.ELEMENT_NODE && node.tagName === 'SPAN' && (node.style.color || node.style.backgroundColor)) {
                const next = node.nextSibling;
                if (next && next.nodeName === 'BR') toRemove.push(next);
                const prev = node.previousSibling;
                if (prev && prev.nodeName === 'BR' && prev.previousSibling && !prev.previousSibling.textContent?.trim()) toRemove.push(prev);
            }
            if (node.childNodes) {
                for (let i = 0; i < node.childNodes.length; i++) scan(node.childNodes[i]);
            }
        };
        scan(container);
        toRemove.forEach(n => n.remove());
    }

    function applyColorToRanges(ranges, color) {
        const newRanges = [];
        const sel = window.getSelection();
        teleprompterText.focus();
        /* Use execCommand for single selection - browser handles DOM, avoids line breaks (per GitHub/Stack Overflow) */
        if (ranges.length === 1) {
            try {
                sel.removeAllRanges();
                sel.addRange(ranges[0]);
                document.execCommand('styleWithCSS', false, true);
                if (document.execCommand('foreColor', false, color)) {
                    const r = sel.rangeCount ? sel.getRangeAt(0) : null;
                    if (r && !r.collapsed) newRanges.push(r.cloneRange());
                }
            } catch (_) { /* fall through to manual */ }
        }
        if (newRanges.length === 0) {
            /* Fallback: manual span creation (multi-select or execCommand failed) */
            let container = null;
            if (ranges.length > 0) {
                try { container = ranges[0].commonAncestorContainer; } catch (_) {}
                if (container && container.nodeType === Node.TEXT_NODE) container = container.parentElement;
            }
            ranges.sort((a, b) => {
                try { return b.compareBoundaryPoints(Range.START_TO_START, a); } catch (_) { return 0; }
            });
            ranges.forEach(range => {
                const atStart = range.startOffset === 0;
                const cursorNode = range.startContainer;
                let prev = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.previousSibling : cursorNode.previousSibling;
                let el = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.parentElement : cursorNode;
                while (el && !prev && el !== teleprompterText) {
                    prev = el.previousSibling;
                    el = el.parentElement;
                }
                if (atStart && prev) {
                    if (prev.nodeName === 'BR') {
                        prev.remove();
                        range.insertNode(document.createTextNode(' '));
                    } else if (prev.nodeType === Node.ELEMENT_NODE && (!prev.textContent?.trim() || (prev.childNodes.length === 1 && prev.firstChild?.nodeName === 'BR'))) {
                        prev.remove();
                        range.insertNode(document.createTextNode(' '));
                    }
                }
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
                toWrap.forEach(({ node, startOff, endOff }) => {
                    let target = node;
                    if (startOff > 0) target = node.splitText(startOff);
                    const len = endOff - startOff;
                    if (target.length > len && len > 0) target.splitText(len);
                    const span = document.createElement('span');
                    span.className = 'color-span-inline';
                    span.style.color = color;
                    target.parentNode.insertBefore(span, target);
                    span.appendChild(target);
                    const r = document.createRange();
                    r.selectNodeContents(span);
                    newRanges.push(r);
                });
            });
            if (container) {
                removeRedundantBrInContainer(container);
                requestAnimationFrame(() => removeRedundantBrInContainer(container));
            }
        }
        return newRanges;
    }

    function applyFontColorToTarget(color) {
        const storedRanges = [...savedFontSelections];
        updateMultiSelectState();
        let rangesToApply = [];
        if (!selectedFontTarget) {
            const validStored = storedRanges.filter(r => {
                try {
                    return r && !r.collapsed && document.contains(r.startContainer) && teleprompterText.contains(r.commonAncestorContainer);
                } catch (_) { return false; }
            });
            if (validStored.length > 0) {
                rangesToApply = validStored;
            } else {
                const sel = window.getSelection();
                if (sel.rangeCount) {
                    const r = sel.getRangeAt(0);
                    if (!r.collapsed && teleprompterText.contains(r.commonAncestorContainer)) {
                        rangesToApply = [r];
                    }
                }
            }
        }
        if (rangesToApply.length > 0) {
            const updated = applyColorToRanges(rangesToApply, color);
            if (updated.length > 0) {
                savedFontSelections = updated.map(r => r.cloneRange());
                updateMultiSelectState();
            }
        } else if (selectedFontTarget) {
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
                if (pFont !== fontFamily || pSize !== fontSize) continue;
                const { color: effColor, backgroundColor: effBg } = getEffectiveColorsFromNode(node);
                if (!colorsMatch(selectedFontTarget, effColor, effBg)) continue;
                toWrap.push(node);
            }
            toWrap.forEach(textNode => {
                const span = document.createElement('span');
                span.className = 'color-span-inline';
                span.style.color = color;
                textNode.parentNode.insertBefore(span, textNode);
                span.appendChild(textNode);
            });
        } else {
            teleprompterText.style.color = color;
        }
        syncEditorState();
    }

    function applyBackgroundColorToRanges(ranges, color) {
        const newRanges = [];
        const sel = window.getSelection();
        teleprompterText.focus();
        /* Use execCommand for single selection - browser handles DOM, avoids line breaks */
        if (ranges.length === 1) {
            try {
                sel.removeAllRanges();
                sel.addRange(ranges[0]);
                document.execCommand('styleWithCSS', false, true);
                const ok = document.execCommand('hiliteColor', false, color)
                    || document.execCommand('backColor', false, color);
                if (ok) {
                    const r = sel.rangeCount ? sel.getRangeAt(0) : null;
                    if (r && !r.collapsed) newRanges.push(r.cloneRange());
                }
            } catch (_) { /* fall through to manual */ }
        }
        if (newRanges.length === 0) {
            let container = null;
            if (ranges.length > 0) {
                try { container = ranges[0].commonAncestorContainer; } catch (_) {}
                if (container && container.nodeType === Node.TEXT_NODE) container = container.parentElement;
            }
            ranges.sort((a, b) => {
                try { return b.compareBoundaryPoints(Range.START_TO_START, a); } catch (_) { return 0; }
            });
            ranges.forEach(range => {
                const atStart = range.startOffset === 0;
                const cursorNode = range.startContainer;
                let prev = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.previousSibling : cursorNode.previousSibling;
                let el = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.parentElement : cursorNode;
                while (el && !prev && el !== teleprompterText) {
                    prev = el.previousSibling;
                    el = el.parentElement;
                }
                if (atStart && prev) {
                    if (prev.nodeName === 'BR') {
                        prev.remove();
                        range.insertNode(document.createTextNode(' '));
                    } else if (prev.nodeType === Node.ELEMENT_NODE && (!prev.textContent?.trim() || (prev.childNodes.length === 1 && prev.firstChild?.nodeName === 'BR'))) {
                        prev.remove();
                        range.insertNode(document.createTextNode(' '));
                    }
                }
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
                toWrap.forEach(({ node, startOff, endOff }) => {
                    let target = node;
                    if (startOff > 0) target = node.splitText(startOff);
                    const len = endOff - startOff;
                    if (target.length > len && len > 0) target.splitText(len);
                    const span = document.createElement('span');
                    span.className = 'color-span-inline';
                    span.style.backgroundColor = color;
                    target.parentNode.insertBefore(span, target);
                    span.appendChild(target);
                    const r = document.createRange();
                    r.selectNodeContents(span);
                    newRanges.push(r);
                });
            });
            if (container) {
                removeRedundantBrInContainer(container);
                requestAnimationFrame(() => removeRedundantBrInContainer(container));
            }
        }
        return newRanges;
    }

    function applyBackgroundColorToTarget(color) {
        const storedRanges = [...savedFontSelections];
        updateMultiSelectState();
        let rangesToApply = [];
        if (!selectedFontTarget) {
            const validStored = storedRanges.filter(r => {
                try {
                    return r && !r.collapsed && document.contains(r.startContainer) && teleprompterText.contains(r.commonAncestorContainer);
                } catch (_) { return false; }
            });
            if (validStored.length > 0) {
                rangesToApply = validStored;
            } else {
                const sel = window.getSelection();
                if (sel.rangeCount) {
                    const r = sel.getRangeAt(0);
                    if (!r.collapsed && teleprompterText.contains(r.commonAncestorContainer)) {
                        rangesToApply = [r];
                    }
                }
            }
        }
        if (rangesToApply.length > 0) {
            const updated = applyBackgroundColorToRanges(rangesToApply, color);
            flattenRedundantSpans();
            if (updated.length > 0) {
                savedFontSelections = updated.map(r => r.cloneRange());
                updateMultiSelectState();
            }
        } else if (selectedFontTarget) {
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
                if (pFont !== fontFamily || pSize !== fontSize) continue;
                const { color: effColor, backgroundColor: effBg } = getEffectiveColorsFromNode(node);
                if (!colorsMatch(selectedFontTarget, effColor, effBg)) continue;
                toWrap.push(node);
            }
            toWrap.forEach(textNode => {
                const span = document.createElement('span');
                span.className = 'color-span-inline';
                span.style.backgroundColor = color;
                textNode.parentNode.insertBefore(span, textNode);
                span.appendChild(textNode);
            });
            flattenRedundantSpans();
        } else {
            teleprompterView.style.backgroundColor = color;
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

    const speedSliderFill = document.getElementById('speed-slider-fill');
    function sliderValueToScrollSpeed(v) {
        if (v === 0 || isNaN(v)) return 0;
        const abs = Math.abs(v);
        const sign = Math.sign(v);
        if (abs <= 5) return sign * abs;
        return sign * (5 + Math.pow(abs - 5, 2) * 4);
    }
    if (speedSlider && speedValueEl) {
        const updateSpeedFromSlider = () => {
            const v = parseFloat(speedSlider.value);
            speedValueEl.textContent = isNaN(v) ? '' : v;
            scrollSpeed = sliderValueToScrollSpeed(v);
            if (speedSliderFill) {
                const pct = Math.abs(v) / 10 * 50;
                speedSliderFill.classList.remove('fill-left', 'fill-right');
                if (v > 0) {
                    speedSliderFill.classList.add('fill-right');
                    speedSliderFill.style.left = '50%';
                    speedSliderFill.style.right = 'auto';
                    speedSliderFill.style.width = pct + '%';
                } else if (v < 0) {
                    speedSliderFill.classList.add('fill-left');
                    speedSliderFill.style.left = (50 - pct) + '%';
                    speedSliderFill.style.right = 'auto';
                    speedSliderFill.style.width = pct + '%';
                } else {
                    speedSliderFill.style.width = '0';
                }
            }
        };
        speedSlider.addEventListener('input', updateSpeedFromSlider);
        speedSlider._updateSpeedFromSlider = updateSpeedFromSlider;
        updateSpeedFromSlider();
        speedSlider.addEventListener('wheel', (e) => {
            e.preventDefault();
            e.stopPropagation();
            let v = parseFloat(speedSlider.value) || 0;
            const step = 0.1 * scrollSensitivity;
            v += e.deltaY < 0 ? step : -step;
            v = Math.max(0, Math.min(10, v));
            speedSlider.value = v;
            speedSlider.dispatchEvent(new Event('input'));
        }, { passive: false, capture: true });
    }

    function startScrolling() {
        if (isTeleprompting) return;
        isTeleprompting = true;
        const move = () => {
            if (!isTeleprompting) return;
            scrollAccum += scrollSpeed;
            const delta = Math.trunc(scrollAccum);
            if (delta !== 0) {
                teleprompterView.scrollTop += delta;
                scrollAccum -= delta;
            }
            syncMirrorByPixels();
            if (typeof checkTopPillAndGoToPreviousFile === 'function') checkTopPillAndGoToPreviousFile();
            if (typeof checkBottomPillAndAdvanceToNextFile === 'function') checkBottomPillAndAdvanceToNextFile();
            animationFrameId = requestAnimationFrame(move);
        };
        animationFrameId = requestAnimationFrame(move);
    }

    startButton.onclick = () => {
        if (!isTeleprompting) {
            const v = speedSlider ? parseFloat(speedSlider.value) : 0;
            scrollSpeed = sliderValueToScrollSpeed(v);
            isPaused = false;
            spacebarControlsPlay = true;
            startScrolling();
        } else if (isPaused) {
            isPaused = false;
            spacebarControlsPlay = true;
            startScrolling();
        } else {
            isPaused = true;
            isTeleprompting = false;
            if (animationFrameId) cancelAnimationFrame(animationFrameId);
            animationFrameId = null;
        }
        updatePlayPauseButton();
        startButton.blur();
        if (speedSlider) speedSlider.focus();
    };
    const btnKeyboardShortcuts = document.getElementById('btn-keyboard-shortcuts');
    const shortcutsOverlay = document.getElementById('shortcuts-overlay');
    const btnCloseShortcuts = document.getElementById('btn-close-shortcuts');
    if (btnKeyboardShortcuts && shortcutsOverlay) {
        btnKeyboardShortcuts.onclick = () => shortcutsOverlay.classList.remove('hidden');
    }
    if (shortcutsOverlay && btnCloseShortcuts) {
        btnCloseShortcuts.onclick = () => shortcutsOverlay.classList.add('hidden');
    }
    if (shortcutsOverlay) {
        shortcutsOverlay.onclick = (e) => {
            if (e.target === shortcutsOverlay) shortcutsOverlay.classList.add('hidden');
        };
        document.addEventListener('keydown', (e) => {
            if (e.key === 'Escape' && !shortcutsOverlay.classList.contains('hidden')) {
                shortcutsOverlay.classList.add('hidden');
            }
        });
    }

    stopButton.onclick = () => {
        isTeleprompting = false;
        isPaused = false;
        scrollAccum = 0;
        if (animationFrameId) cancelAnimationFrame(animationFrameId);
        animationFrameId = null;
        teleprompterView.scrollTop = 0;
        if (speedSlider) {
            speedSlider.value = 0;
            speedSlider.dispatchEvent(new Event('input'));
        }
        syncMirrorByPixels();
        updatePlayPauseButton();
    };

    document.addEventListener('keydown', (e) => {
        if (e.key === 'Escape' && (document.activeElement === teleprompterText || teleprompterText.contains(document.activeElement))) {
            teleprompterText.blur();
            teleprompterView.focus();
        }
        if (e.altKey && e.code === 'KeyB') {
            e.preventDefault();
            e.stopPropagation();
            addBookmarkAtCursor();
            return;
        }
        const isPrevSegment = e.altKey && (e.code === 'ArrowUp' || e.key === 'ArrowUp');
        const isNextSegment = e.altKey && (e.code === 'ArrowDown' || e.key === 'ArrowDown');
        if (isPrevSegment || isNextSegment) {
            e.preventDefault();
            e.stopPropagation();
            const inEditor = document.activeElement === teleprompterText || teleprompterText.contains(document.activeElement);
            if (inEditor) {
                teleprompterText.setAttribute('contenteditable', 'false');
                teleprompterView.focus();
                requestAnimationFrame(() => {
                    teleprompterText.setAttribute('contenteditable', 'true');
                    teleprompterView.focus();
                    requestAnimationFrame(() => {
                        if (isPrevSegment) goToPrevBookmark();
                        else goToNextBookmark();
                    });
                });
            } else {
                if (isPrevSegment) goToPrevBookmark();
                else goToNextBookmark();
            }
            return;
        }
        const opt = e.altKey;
        const isA = e.code === 'KeyA' || (e.key || '').toLowerCase() === 'a';
        const isS = e.code === 'KeyS' || (e.key || '').toLowerCase() === 's';
        if (isA && opt) {
            e.preventDefault();
            e.stopPropagation();
            spacebarControlsPlay = true;
            startButton.click();
        } else if (isS && opt && isTeleprompting) {
            e.preventDefault();
            e.stopPropagation();
            stopButton.click();
        } else if (e.code === 'Space' || e.key === ' ') {
            const canToggle = isTeleprompting || isPaused;
            const canStart = spacebarControlsPlay && !canToggle;
            if (canToggle || canStart) {
                e.preventDefault();
                e.stopPropagation();
                startButton.click();
            }
        }
    }, true);

    document.addEventListener('mousedown', () => { spacebarControlsPlay = false; }, true);

    let heldKeyInterval = null;
    /** isRepeat: true when called from the hold interval. When !isInvertScroll, Down/Left stops at 0. When isInvertScroll, Up/Right stops at 0 (opposite direction). */
    const applyArrowStep = (key, step, isRepeat = false) => {
        let v = parseFloat(speedSlider.value) || 0;
        const isDown = key === 'ArrowLeft' || key === 'ArrowDown';
        const isUp = key === 'ArrowRight' || key === 'ArrowUp';
        const stopAtZeroKey = isInvertScroll ? isUp : isDown;
        if (stopAtZeroKey) {
            if (isDown) {
                if (v > 0) {
                    v = Math.max(0, v - step);  /* slowing down: stop at 0 */
                } else if (v === 0 && isRepeat) {
                    v = 0;  /* holding at 0: don't go reverse until key pressed again */
                } else {
                    v = Math.max(-10, v - step);  /* new press at 0, or already negative: allow reverse / keep going */
                }
            } else {
                if (v < 0) {
                    v = Math.min(0, v + step);  /* slowing down toward 0: stop at 0 */
                } else if (v === 0 && isRepeat) {
                    v = 0;  /* holding at 0: don't go forward until key pressed again */
                } else {
                    v = Math.min(10, v + step);  /* new press at 0, or already positive: allow forward / keep going */
                }
            }
        } else {
            if (isDown) {
                v = Math.max(-10, v - step);  /* no stop at 0 */
            } else {
                v = Math.min(10, v + step);  /* no stop at 0 */
            }
        }
        v = Math.round(v * 100) / 100;
        speedSlider.value = v;
        speedSlider.dispatchEvent(new Event('input'));
    };
    document.addEventListener('keydown', (e) => {
        if (!isTeleprompting || isPaused || !speedSlider) return;
        const key = e.key;
        if (['ArrowLeft', 'ArrowDown', 'ArrowRight', 'ArrowUp'].includes(key)) {
            e.preventDefault();
            e.stopPropagation();
            if (e.repeat && heldKeyInterval) return;

            if (heldKeyInterval) clearInterval(heldKeyInterval);
            applyArrowStep(key, 0.1, false);  /* initial press: can cross zero */
            const heldKeyStart = Date.now();
            heldKeyInterval = setInterval(() => {
                const elapsedSec = (Date.now() - heldKeyStart) / 1000;
                const step = 0.1 * (1 + elapsedSec * 2);
                applyArrowStep(key, step, true);  /* repeat: stop at 0, don't cross */
            }, 80);
        }
    }, true);
    document.addEventListener('keyup', (e) => {
        if (['ArrowLeft', 'ArrowDown', 'ArrowRight', 'ArrowUp'].includes(e.key)) {
            if (heldKeyInterval) {
                clearInterval(heldKeyInterval);
                heldKeyInterval = null;
            }
        }
    }, true);

    /* When Mouse mode is selected and we're playing: wheel over teleprompter changes speed. Otherwise wheel scrolls the content. Listen on document (capture) so we always receive wheel when cursor is over the page. */
    if (teleprompterView && speedSlider) {
        const WHEEL_DEBUG = typeof window !== 'undefined' && window.DEBUG_WHEEL_SPEED;
        let wheelLogCount = 0;
        function handleWheelForSpeed(e) {
            const overTeleprompter = teleprompterView && teleprompterView.contains(e.target);
            const mouseMode = document.querySelector('input[name="controlMode"]:checked')?.value === 'mouse';
            const useWheelForSpeed = overTeleprompter && mouseMode && isTeleprompting && !isPaused;
            const shouldLog = WHEEL_DEBUG || wheelLogCount < 5;
            if (shouldLog) {
                wheelLogCount += 1;
                console.log('[WheelSpeed]' + (wheelLogCount <= 5 ? ' (first 5 always)' : ''), {
                    overTeleprompter,
                    mouseMode,
                    isTeleprompting,
                    isPaused,
                    useWheelForSpeed,
                    deltaY: e.deltaY,
                    deltaMode: e.deltaMode,
                    targetId: e.target?.id,
                    targetTag: e.target?.tagName
                });
            }
            if (!useWheelForSpeed) return;
            e.preventDefault();
            e.stopPropagation();
            const isTrackpad = e.deltaMode === 0;
            const mag = Math.abs(e.deltaY);
            let step = isTrackpad
                ? Math.max(0.02, Math.min(0.22, mag * 0.0005))
                : Math.max(0.5, Math.min(2.5, mag * 0.01));
            step *= scrollSensitivity;
            const goUp = e.deltaY < 0;
            const key = (goUp && !isInvertScroll) || (!goUp && isInvertScroll) ? 'ArrowUp' : 'ArrowDown';
            const prevVal = parseFloat(speedSlider.value) || 0;
            applyArrowStep(key, step, false);
            let newVal = parseFloat(speedSlider.value) || 0;
            if ((prevVal > 0 && newVal < 0) || (prevVal < 0 && newVal > 0)) {
                speedSlider.value = 0;
                newVal = 0;
            }
            if (newVal < 0) {
                speedSlider.value = 0;
            }
            if (typeof speedSlider._updateSpeedFromSlider === 'function') speedSlider._updateSpeedFromSlider();
            if (shouldLog) console.log('[WheelSpeed] applied', { step, key, newValue: speedSlider?.value });
        }
        document.addEventListener('wheel', handleWheelForSpeed, { passive: false, capture: true });
        console.log('[WheelSpeed] Listener attached on document (capture). Set DEBUG_WHEEL_SPEED = true for every wheel, or scroll wheel to see first 5.');
    }

    // -----------------------------------------
    // WebHID: ShuttleXPress (Contour Design)
    // Button 1 = play/pause, 2 = none, 3 = stop/top, 4 = prev file, 5 = next file
    // Shuttle ring = scroll; holding = faster
    // -----------------------------------------
    const SHUTTLE_VENDOR_ID = 0x0b33;
    const SHUTTLE_PRODUCT_ID = 0x0020;
    let shuttleDevice = null;
    let shuttlePrevButtons = 0;
    let shuttleHoldStart = 0;
    const SHUTTLE_ACCEL_BASE = 8;
    const SHUTTLE_ACCEL_MAX = 80;

    function parseShuttleReport(data) {
        if (!data || data.byteLength < 3) return null;
        const buf = new Uint8Array(data);
        const off = data.byteLength >= 4 ? 1 : 0;
        const buttons = buf[off] & 0x1f;
        const jog = buf[off + 1];
        let shuttle = buf[off + 2];
        if (shuttle > 7) shuttle -= 16;
        shuttle = Math.max(-7, Math.min(7, shuttle));
        return { buttons, jog, shuttle };
    }

    function onShuttleInput(e) {
        const report = parseShuttleReport(e.data.buffer);
        if (!report) return;
        const { buttons, shuttle } = report;
        const now = Date.now();
        for (let b = 0; b < 5; b++) {
            const mask = 1 << b;
            if ((buttons & mask) && !(shuttlePrevButtons & mask)) {
                if (b === 0) startButton?.click();
                else if (b === 2) stopButton?.click();
                else if (b === 3) {
                    if (typeof loadScriptToEditor === 'function' && currentFileIndex > 0)
                        loadScriptToEditor(currentFileIndex - 1);
                } else if (b === 4) {
                    if (typeof loadScriptToEditor === 'function' && fileStore?.length && currentFileIndex < fileStore.length - 1)
                        loadScriptToEditor(currentFileIndex + 1);
                }
            }
        }
        shuttlePrevButtons = buttons;

        if (shuttle !== 0) {
            if (shuttleHoldStart === 0) shuttleHoldStart = now;
            const holdSec = (now - shuttleHoldStart) / 1000;
            const accel = Math.min(SHUTTLE_ACCEL_MAX, SHUTTLE_ACCEL_BASE + holdSec * 40);
            const delta = Math.round((shuttle > 0 ? 1 : -1) * accel);
            if (teleprompterView) {
                const maxScroll = teleprompterView.scrollHeight - teleprompterView.clientHeight;
                teleprompterView.scrollTop = Math.max(0, Math.min(maxScroll, teleprompterView.scrollTop + delta));
            }
            if (typeof syncMirrorByPixels === 'function') syncMirrorByPixels();
        } else {
            shuttleHoldStart = 0;
        }
    }

    function setShuttleConnected(connected) {
        const btn = document.getElementById('btn-connect-shuttle');
        if (!btn) return;
        if (connected) {
            btn.title = 'Disconnect Shuttle (click to release device)';
            btn.classList.add('shuttle-connected');
        } else {
            btn.title = 'Connect ShuttleXPress (Chrome). Option+click: show all HID devices.';
            btn.classList.remove('shuttle-connected');
        }
    }

    async function disconnectShuttle() {
        if (!shuttleDevice) return;
        try {
            shuttleDevice.removeEventListener('inputreport', onShuttleInput);
            await shuttleDevice.close();
        } catch (_) {}
        shuttleDevice = null;
        shuttlePrevButtons = 0;
        shuttleHoldStart = 0;
        setShuttleConnected(false);
    }

    async function connectShuttle(e) {
        if (shuttleDevice) {
            await disconnectShuttle();
            return;
        }
        if (!navigator.hid) {
            alert('WebHID is not supported in this browser. Use Chrome.');
            return;
        }
        try {
            /* Option/Alt: show all HID devices (to find Shuttle if it doesn't match Contour vendor, and log its ids) */
            const showAll = navigator.platform?.toLowerCase().includes('mac') ? e?.altKey : false;
            const devices = await navigator.hid.requestDevice({
                filters: showAll ? [] : [{ vendorId: SHUTTLE_VENDOR_ID }]
            });
            if (!devices.length) return;
            if (showAll && devices[0]) {
                console.log('Selected HID device:', devices[0].productName, 'vendorId:', '0x' + devices[0].vendorId.toString(16), 'productId:', '0x' + devices[0].productId.toString(16));
            }
            /* Use only the first selected device to avoid holding multiple */
            shuttleDevice = devices[0];
            await shuttleDevice.open();
            shuttleDevice.addEventListener('inputreport', onShuttleInput);
            setShuttleConnected(true);
            if (navigator.hid) navigator.hid.addEventListener('disconnect', function onDisconnect(ev) {
                if (ev.device === shuttleDevice) {
                    shuttleDevice = null;
                    shuttlePrevButtons = 0;
                    shuttleHoldStart = 0;
                    setShuttleConnected(false);
                    navigator.hid.removeEventListener('disconnect', onDisconnect);
                }
            });
        } catch (err) {
            console.error('Shuttle connect failed:', err);
            alert('Could not connect to Shuttle: ' + (err.message || err));
        }
    }

    const btnConnectShuttle = document.getElementById('btn-connect-shuttle');
    if (btnConnectShuttle) {
        btnConnectShuttle.onclick = (e) => connectShuttle(e);
    }

    console.log("🚀 Teleprompter Engine Initialized");

    requestAnimationFrame(() => refreshSelectFontTarget());

    let skipCaretSyncForFontSelects = false;
    function updateFontSelectsFromCaret() {
        if (skipCaretSyncForFontSelects) return;
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

    function applyFontToAllContent(fontVal, sizeVal) {
        if (fontVal) teleprompterText.style.fontFamily = fontVal;
        if (sizeVal) teleprompterText.style.fontSize = sizeVal + 'px';
        teleprompterText.querySelectorAll('*').forEach(el => {
            if (fontVal) el.style.fontFamily = fontVal;
            if (sizeVal) el.style.fontSize = sizeVal + 'px';
        });
    }

    /** Apply only font size so font family (and other) changes from Aa/un-Aa stay. */
    function applyFontSizeOnlyToAllContent(sizeVal) {
        if (!sizeVal) return;
        teleprompterText.style.fontSize = sizeVal + 'px';
        teleprompterText.querySelectorAll('*').forEach(el => {
            el.style.fontSize = sizeVal + 'px';
        });
    }

    function applyFontSettings() {
        pushUndoState();
        const storedRanges = [...savedFontSelections];
        updateMultiSelectState();
        const fontVal = fontFamilySelect?.value;
        const sizeVal = fontSizeSelect?.value;
        const targetToApply = selectedFontTarget ? { ...selectedFontTarget } : null;

        const doApply = () => {
            let rangesToApply = [];
            if (!targetToApply) {
                const validStored = storedRanges.filter(r => {
                    try {
                        return r && !r.collapsed && document.contains(r.startContainer) && teleprompterText.contains(r.commonAncestorContainer);
                    } catch (_) { return false; }
                });
                if (validStored.length > 0) {
                    rangesToApply = validStored;
                } else {
                    teleprompterText.focus();
                    const sel = window.getSelection();
                    const r = sel.rangeCount ? sel.getRangeAt(0) : null;
                    if (r && !r.collapsed && teleprompterText.contains(r.commonAncestorContainer)) {
                        rangesToApply = [r];
                    }
                }
            }

            if (rangesToApply.length > 0 && (fontFamilySelect || fontSizeSelect)) {
                const preserveSize = lastFontChangeSource === 'family';
                const preserveFamily = lastFontChangeSource === 'size';
                rangesToApply.sort((a, b) => {
                    try { return b.compareBoundaryPoints(Range.START_TO_START, a); } catch (_) { return 0; }
                });
                rangesToApply.forEach(range => {
                    try {
                        if (document.contains(range.startContainer)) {
                            wrapSelectionInSpan(range, fontVal, sizeVal, preserveSize, preserveFamily);
                        }
                    } catch (_) {}
                });
                teleprompterText.focus();
            } else if (targetToApply && (fontFamilySelect || fontSizeSelect)) {
                applyFontToMatchingTarget(targetToApply, fontVal, sizeVal);
            } else if (!targetToApply && (fontFamilySelect || fontSizeSelect)) {
                applyFontToAllContent(fontVal || fontFamilySelect?.value, sizeVal || fontSizeSelect?.value);
            } else {
                applyFontToAllContent(fontVal || fontFamilySelect?.value, sizeVal || fontSizeSelect?.value);
            }
            flattenRedundantSpans();
            teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
                row.style.minHeight = '';
                row.style.height = '';
            });
            measuredRowHeights = [];
            requestAnimationFrame(() => {
                requestAnimationFrame(() => {
                    void teleprompterText.offsetHeight;
                    measureRowHeightsFromContent();
                    if (mirrorWindow && !mirrorWindow.closed) syncMirrorStyles();
                    syncEditorState();
                    skipCaretSyncForFontSelects = false;
                });
            });
        };

        skipCaretSyncForFontSelects = true;
        teleprompterText.focus();
        requestAnimationFrame(() => {
            requestAnimationFrame(doApply);
        });
    }
    const isFontControl = (el) => el && (fontFamilySelect?.contains(el) || fontSizeSelect?.contains(el) || fontColorButton?.contains(el) || fontColorPanel?.contains(el) || bgColorButton?.contains(el) || bgColorPanel?.contains(el) || boldButton?.contains(el) || italicButton?.contains(el) || underlineButton?.contains(el) || document.getElementById('font-target-dropdown')?.contains(el) || selectFontTarget?.contains(el) || document.getElementById('btn-overview-toggle')?.contains(el));
    const saveSelectionWhenFocusingFontControl = (e) => {
        if (!isFontControl(e.target)) return;
        const sel = window.getSelection();
        if (sel.rangeCount && sel.getRangeAt(0) && !sel.getRangeAt(0).collapsed && teleprompterText.contains(sel.anchorNode))
            saveFontSelectionFromEditor(false);
        else if (savedFontSelections.length === 0) saveFontSelectionFromEditor(false);
        updateMultiSelectHighlight();
    };
    const saveSelectionOnEditorBlur = (e) => {
        const next = e.relatedTarget;
        const ribbon = document.querySelector('.top-ribbon');
        if (next && (ribbon?.contains(next) || bgColorPanel?.contains(next) || fontColorPanel?.contains(next) || runlistPanel?.contains(next))) return;
        if (isFontControl(next) && savedFontSelections.length === 0) saveFontSelectionFromEditor();
        updateMultiSelectState();
    };
    function updateMultiSelectState() {
        updateMultiSelectHighlight();
    }
    function pushUndoState() {
        if (undoStack.length >= MAX_UNDO) undoStack.shift();
        undoStack.push({
            html: teleprompterText.innerHTML,
            bgColor: teleprompterView.style.backgroundColor,
            fontColor: teleprompterText.style.color,
            fontFamily: teleprompterText.style.fontFamily,
            fontSize: teleprompterText.style.fontSize
        });
        redoStack.length = 0;
    }
    function undo() {
        if (undoStack.length === 0) return false;
        redoStack.push({
            html: teleprompterText.innerHTML,
            bgColor: teleprompterView.style.backgroundColor,
            fontColor: teleprompterText.style.color,
            fontFamily: teleprompterText.style.fontFamily,
            fontSize: teleprompterText.style.fontSize
        });
        const state = undoStack.pop();
        teleprompterText.innerHTML = state.html;
        teleprompterView.style.backgroundColor = state.bgColor;
        teleprompterText.style.color = state.fontColor;
        if (state.fontFamily) teleprompterText.style.fontFamily = state.fontFamily;
        if (state.fontSize) teleprompterText.style.fontSize = state.fontSize;
        if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
            contentStore[currentFileIndex] = state.html;
        }
        if (mirrorWindow && !mirrorWindow.closed) {
            refreshMirrorData();
            syncMirrorStyles();
        }
        syncEditorState();
        updateBookmarkSidebar();
        return true;
    }
    function redo() {
        if (redoStack.length === 0) return false;
        undoStack.push({
            html: teleprompterText.innerHTML,
            bgColor: teleprompterView.style.backgroundColor,
            fontColor: teleprompterText.style.color,
            fontFamily: teleprompterText.style.fontFamily,
            fontSize: teleprompterText.style.fontSize
        });
        const state = redoStack.pop();
        teleprompterText.innerHTML = state.html;
        teleprompterView.style.backgroundColor = state.bgColor;
        teleprompterText.style.color = state.fontColor;
        if (state.fontFamily) teleprompterText.style.fontFamily = state.fontFamily;
        if (state.fontSize) teleprompterText.style.fontSize = state.fontSize;
        if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
            contentStore[currentFileIndex] = state.html;
        }
        if (mirrorWindow && !mirrorWindow.closed) {
            refreshMirrorData();
            syncMirrorStyles();
        }
        syncEditorState();
        updateBookmarkSidebar();
        return true;
    }
    function updateMultiSelectHighlight() {
        if (typeof CSS === 'undefined' || !CSS.highlights || typeof Highlight === 'undefined') return;
        try {
            if (savedFontSelections.length > 0) {
                const valid = savedFontSelections.filter(r => {
                    try { return r && document.contains(r.startContainer); } catch (_) { return false; }
                });
                if (valid.length > 0) {
                    const highlight = new Highlight(...valid);
                    CSS.highlights.set('multi-select', highlight);
                    return;
                }
            }
            CSS.highlights.delete('multi-select');
        } catch (_) {}
    }
    document.addEventListener('mouseover', saveSelectionWhenFocusingFontControl, true);
    document.addEventListener('mousedown', saveSelectionWhenFocusingFontControl, true);
    document.addEventListener('pointerdown', saveSelectionWhenFocusingFontControl, true);
    teleprompterText.addEventListener('focusout', saveSelectionOnEditorBlur);

    /* On paste, insert plain text only so pasted content does not bring span colors/font rules from other rows. */
    teleprompterText.addEventListener('paste', (e) => {
        e.preventDefault();
        const text = (e.clipboardData || window.clipboardData)?.getData('text/plain') ?? '';
        if (text !== '') {
            document.execCommand('insertText', false, text);
        }
    });

    if (fontFamilySelect) {
        fontFamilySelect.addEventListener('change', () => {
            lastFontChangeSource = 'family';
            try { localStorage.setItem(FONT_FAMILY_STORAGE_KEY, fontFamilySelect.value); } catch (_) {}
            skipCaretSyncForFontSelects = true;
            applyFontSettings();
            updateFilenamePills();
        });
    }
    if (fontSizeSelect) {
        fontSizeSelect.addEventListener('change', () => {
            lastFontChangeSource = 'size';
            try { localStorage.setItem(FONT_SIZE_STORAGE_KEY, fontSizeSelect.value); } catch (_) {}
            skipCaretSyncForFontSelects = true;
            if (isOverviewMode) {
                savedFontSizeBeforeOverview = fontSizeSelect.value;
                applyFontSettings();
                requestAnimationFrame(() => refreshSelectFontTarget());
                updateFilenamePills();
                return;
            }
            applyFontSettings();
            updateFilenamePills();
        });
    }

    const btnOverviewToggle = document.getElementById('btn-overview-toggle');
    if (btnOverviewToggle && fontSizeSelect) {
        btnOverviewToggle.onclick = () => {
            const isExtended = document.body.classList.contains('broadcasting') && mirrorWindow && !mirrorWindow.closed;
            if (isExtended) {
                broadcastEditMode = !broadcastEditMode;
                if (broadcastEditMode) {
                    savedFontSizeBeforeOverview = fontSizeSelect.value;
                    fontSizeSelect.value = '12';
                    lastFontChangeSource = 'size';
                    applyFontSizeOnlyToAllContent('12');
                    isOverviewMode = true;
                    document.body.classList.add('overview-mode');
                    btnOverviewToggle.classList.add('overview-active');
                } else {
                    const rawRestore = savedFontSizeBeforeOverview || '80';
                    const restore = (rawRestore && rawRestore !== '12') ? rawRestore : '80';
                    fontSizeSelect.value = restore;
                    lastFontChangeSource = 'size';
                    applyFontSizeOnlyToAllContent(restore);
                    isOverviewMode = false;
                    document.body.classList.remove('overview-mode');
                    btnOverviewToggle.classList.remove('overview-active');
                }
                flattenRedundantSpans();
                /* When un-Aa'ing: compute row colors while col3 is still visible, then hide col3. Otherwise hidden col3 yields empty text and we get wrong (green) colors. */
                if (!broadcastEditMode) {
                    syncColumnWidths();
                }
                applyBroadcastingVisibility();
                syncColumnWidths();
                requestAnimationFrame(() => {
                    measureRowHeightsFromContent();
                    refreshSelectFontTarget();
                    updateFilenamePills();
                    if (mirrorWindow && !mirrorWindow.closed) {
                        syncMirrorStyles();
                        refreshMirrorData();
                    }
                    syncEditorState();
                    if (!broadcastEditMode) {
                        processTableColumns();
                        wrapCellContentInBlock();
                        syncColumnWidths();
                        measureRowHeightsWithProbeForBroadcasting();
                        if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
                    }
                });
                if (broadcastEditMode) {
                    btnOverviewToggle.classList.add('broadcast-edit-active');
                    btnOverviewToggle.title = 'Hide all columns (back to control view)';
                } else {
                    btnOverviewToggle.classList.remove('broadcast-edit-active');
                    btnOverviewToggle.title = 'Show all columns to edit (mirror stays open)';
                }
                return;
            }
            let sizeToReapply = null;
            if (isOverviewMode) {
                const rawRestore = savedFontSizeBeforeOverview || '80';
                const restore = (rawRestore && rawRestore !== '12') ? rawRestore : '80';
                sizeToReapply = restore;
                fontSizeSelect.value = restore;
                lastFontChangeSource = 'size';
                applyFontSizeOnlyToAllContent(restore);
                isOverviewMode = false;
                btnOverviewToggle.classList.remove('overview-active');
                document.body.classList.remove('overview-mode');
            } else {
                savedFontSizeBeforeOverview = fontSizeSelect.value;
                fontSizeSelect.value = '12';
                lastFontChangeSource = 'size';
                applyFontSizeOnlyToAllContent('12');
                isOverviewMode = true;
                btnOverviewToggle.classList.add('overview-active');
                document.body.classList.add('overview-mode');
            }
            flattenRedundantSpans();
            requestAnimationFrame(() => {
                syncColumnWidths();
                measureRowHeightsFromContent();
                if (sizeToReapply) applyFontSizeOnlyToAllContent(sizeToReapply);
                refreshSelectFontTarget();
                updateFilenamePills();
                if (mirrorWindow && !mirrorWindow.closed) syncMirrorStyles();
                syncEditorState();
            });
        };
    }

    let modifierHeldAtMouseDown = false;
    let modifierKeyCurrentlyPressed = false;
    let mouseDownInEditor = false;
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Meta' || e.key === 'Control') modifierKeyCurrentlyPressed = true;
    }, true);
    document.addEventListener('keyup', (e) => {
        if (e.key === 'Meta' || e.key === 'Control') modifierKeyCurrentlyPressed = false;
    }, true);
    const onPtrDown = (e) => {
        if (teleprompterText.contains(e.target)) {
            modifierHeldAtMouseDown = e.metaKey || e.ctrlKey;
            mouseDownInEditor = true;
            if (!e.metaKey && !e.ctrlKey && savedFontSelections.length > 0) {
                savedFontSelections = [];
                updateMultiSelectState();
            }
        }
    };
    document.addEventListener('mousedown', onPtrDown, true);
    document.addEventListener('pointerdown', onPtrDown, true);
    const onPtrUp = (e) => {
        if (!mouseDownInEditor) return;
        mouseDownInEditor = false;
        const sel = window.getSelection();
        if (!sel.rangeCount) return;
        const r = sel.getRangeAt(0);
        if (r.collapsed || !teleprompterText.contains(r.commonAncestorContainer)) return;
        const addToMulti = modifierHeldAtMouseDown || (e && (e.metaKey || e.ctrlKey)) || modifierKeyCurrentlyPressed;
        modifierHeldAtMouseDown = false;
        try {
            if (addToMulti) {
                savedFontSelections.push(r.cloneRange());
            } else {
                savedFontSelections = [r.cloneRange()];
            }
            updateMultiSelectState();
            if (document.activeElement === teleprompterText || teleprompterText.contains(document.activeElement)) {
                selectedFontTarget = null;
            }
        } catch (_) {}
    };
    document.addEventListener('mouseup', onPtrUp, true);
    document.addEventListener('pointerup', onPtrUp, true);
    document.addEventListener('pointercancel', () => { mouseDownInEditor = false; modifierHeldAtMouseDown = false; }, true);
    document.addEventListener('selectionchange', () => {
        if (document.activeElement === teleprompterText) {
            updateFontSelectsFromCaret();
        }
    });
    teleprompterText.addEventListener('click', updateFontSelectsFromCaret);
    teleprompterText.addEventListener('keyup', updateFontSelectsFromCaret);
    teleprompterText.addEventListener('focus', updateFontSelectsFromCaret);

    document.addEventListener('keydown', (e) => {
        const mod = e.metaKey || e.ctrlKey;
        const key = (e.key || '').toLowerCase();
        const inEditor = teleprompterText.contains(document.activeElement) || document.activeElement === teleprompterText;

        if (mod && key === 'z') {
            e.preventDefault();
            if (e.shiftKey) {
                if (!redo()) document.execCommand('redo');
            } else {
                if (!undo()) document.execCommand('undo');
            }
        } else if (mod && key === 'y') {
            e.preventDefault();
            if (!redo()) document.execCommand('redo');
        } else if (key === 'enter') {
            if (inEditor) {
                e.preventDefault();
                document.execCommand('insertLineBreak');
            }
        } else if (key === 'escape' && savedFontSelections.length > 0) {
            savedFontSelections = [];
            updateMultiSelectState();
        }
    }, true);

    if (selectFontTarget) {
        selectFontTarget.addEventListener('change', () => {
            try {
                selectedFontTarget = selectFontTarget.value ? JSON.parse(selectFontTarget.value) : null;
            } catch (_) {
                selectedFontTarget = null;
            }
        });
    }
    const fontTargetDropdown = document.getElementById('font-target-dropdown');
    const fontTargetTrigger = document.getElementById('font-target-trigger');
    const fontTargetList = document.getElementById('font-target-list');
    if (fontTargetTrigger && fontTargetList) {
        fontTargetTrigger.addEventListener('click', (e) => {
            e.stopPropagation();
            fontTargetList.classList.toggle('open');
        });
        document.addEventListener('click', () => fontTargetList.classList.remove('open'));
        fontTargetList.addEventListener('click', (e) => e.stopPropagation());
    }
    const eventPlanTrigger = document.getElementById('btn-event-plan-trigger');
    const eventPlanList = document.getElementById('event-plan-list');
    if (eventPlanTrigger && eventPlanList) {
        eventPlanTrigger.addEventListener('click', (e) => {
            e.stopPropagation();
            eventPlanList.classList.toggle('open');
        });
        document.addEventListener('click', () => eventPlanList.classList.remove('open'));
        eventPlanList.addEventListener('click', (e) => {
            e.stopPropagation();
            const item = e.target.closest('.event-plan-item');
            if (!item) return;
            eventPlanList.classList.remove('open');
            const action = item.dataset.action;
            if (action === 'event-plan') {
                openEventPlanOverlay();
            } else if (action === 'event-script') {
                if (!fileStore || fileStore.length === 0) {
                    alert('No files in the Files list. Add files first.');
                    return;
                }
                if (typeof syncEditorState === 'function') syncEditorState();
                const html = buildMasterScriptHtml();
                const w = window.open('', '_blank');
                if (!w) {
                    alert('Popup blocked. Please allow popups to print.');
                    return;
                }
                w.document.write(html.replace('</body></html>', '<script>window.onload=function(){window.onafterprint=function(){window.close()};window.print()}<\/script></body></html>'));
                w.document.close();
                w.focus();
            }
        });
    }
    if (btnRemoveFontTarget) {
        btnRemoveFontTarget.onclick = () => {
            if (!selectedFontTarget) return;
            pushUndoState();
            unwrapFontFromMatchingTarget(selectedFontTarget);
            selectFontTarget.value = '';
            selectedFontTarget = null;
            const trigger = document.getElementById('font-target-trigger');
            if (trigger) trigger.textContent = 'Select';
            teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
                row.style.minHeight = '';
                row.style.height = '';
            });
            measuredRowHeights = [];
            const scheduleRemeasure = () => {
                requestAnimationFrame(() => {
                    requestAnimationFrame(() => {
                        void teleprompterText.offsetHeight;
                        measureRowHeightsFromContent();
                        syncEditorState();
                        if (mirrorWindow && !mirrorWindow.closed) {
                            refreshMirrorData();
                            syncMirrorStyles();
                        }
                    });
                });
            };
            scheduleRemeasure();
        };
    }

    if (bgColorButton && bgColorPanel) bgColorButton.onclick = () => togglePanel(bgColorButton, bgColorPanel);
    if (fontColorButton && fontColorPanel) fontColorButton.onclick = () => togglePanel(fontColorButton, fontColorPanel);

    function applyFormattingToRanges(ranges, tagName) {
        const getCell = (node) => {
            let n = node.nodeType === Node.TEXT_NODE ? node.parentElement : node;
            while (n && n !== teleprompterText) {
                if (n.classList?.contains('script-column')) return n;
                n = n.parentElement;
            }
            return null;
        };
        ranges.sort((a, b) => {
            try { return b.compareBoundaryPoints(Range.START_TO_START, a); } catch (_) { return 0; }
        });
        ranges.forEach(range => {
            try {
                const startCell = getCell(range.startContainer);
                const endCell = getCell(range.endContainer);
                if (startCell != null && endCell != null && startCell !== endCell) return;
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
                toWrap.forEach(({ node, startOff, endOff }) => {
                    let target = node;
                    if (startOff > 0) target = node.splitText(startOff);
                    const len = endOff - startOff;
                    if (target.length > len && len > 0) target.splitText(len);
                    const parent = target.parentNode;
                    if (parent && parent.tagName === tagName.toUpperCase() && parent.childNodes.length === 1) {
                        parent.parentNode.insertBefore(target, parent);
                        parent.remove();
                    } else {
                        const el = document.createElement(tagName);
                        target.parentNode.insertBefore(el, target);
                        el.appendChild(target);
                    }
                });
            } catch (_) {}
        });
    }

    function applyFormatting(cmd) {
        pushUndoState();
        const tagMap = { bold: 'b', italic: 'i', underline: 'u' };
        const tagName = tagMap[cmd];
        let ranges = [];
        if (selectedFontTarget) {
            const { fontFamily, fontSize } = selectedFontTarget;
            const normalizeFont = (s) => (s || '').split(',')[0].trim().replace(/^["']|["']$/g, '').toLowerCase();
            const normalizeSize = (s) => String(s || '').trim();
            const walker = document.createTreeWalker(teleprompterText, NodeFilter.SHOW_TEXT, null, false);
            let node;
            while ((node = walker.nextNode())) {
                if (!node.textContent.trim()) continue;
                const parent = node.parentElement;
                if (!parent) continue;
                const style = window.getComputedStyle(parent);
                const pFont = normalizeFont(style.fontFamily);
                const pSize = normalizeSize(style.fontSize);
                const tFont = normalizeFont(fontFamily);
                const tSize = normalizeSize(fontSize);
                const { color: effColor, backgroundColor: effBg } = getEffectiveColorsFromNode(node);
                if (tFont && tSize && pFont === tFont && pSize === tSize && colorsMatch(selectedFontTarget, effColor, effBg)) {
                    const r = document.createRange();
                    r.setStart(node, 0);
                    r.setEnd(node, node.length);
                    ranges.push(r);
                }
            }
        }
        if (ranges.length === 0 && savedFontSelections.length > 0) {
            ranges = savedFontSelections.filter(r => {
                try { return r && document.contains(r.startContainer); } catch (_) { return false; }
            });
        }
        teleprompterText.focus();
        if (ranges.length > 0 && tagName) {
            applyFormattingToRanges(ranges, tagName);
        } else {
            const sel = window.getSelection();
            const toRestore = savedFontSelections.find(r => { try { return r && document.contains(r.startContainer); } catch (_) { return false; } });
            if (toRestore) {
                sel.removeAllRanges();
                sel.addRange(toRestore.cloneRange());
            }
            if (sel.rangeCount) document.execCommand(cmd, false, null);
        }
        flattenRedundantSpans();
        updateMultiSelectState();
        syncEditorState();
    }
    if (boldButton) boldButton.onclick = () => applyFormatting('bold');
    if (italicButton) italicButton.onclick = () => applyFormatting('italic');
    if (underlineButton) underlineButton.onclick = () => applyFormatting('underline');

    function handleColorBoxSelect(e) {
        const fc = e.target.closest('.color-options');
        if (fc && fc.closest('#font-color-panel')) {
            const box = e.target.closest('.color-box');
            if (box) {
                const color = box.getAttribute('data-color');
                if (color) {
                    e.preventDefault();
                    pushUndoState();
                    applyFontColorToTarget(color);
                }
                fontColorPanel.style.display = 'none';
            }
        } else if (fc && fc.closest('#bg-color-panel')) {
            const box = e.target.closest('.color-box');
            if (box) {
                const color = box.getAttribute('data-color');
                if (color) {
                    e.preventDefault();
                    pushUndoState();
                    applyBackgroundColorToTarget(color);
                }
                bgColorPanel.style.display = 'none';
            }
        }
    }
    document.addEventListener('mousedown', handleColorBoxSelect, true);

    teleprompterView.addEventListener('scroll', syncMirrorByPixels);
    let bookmarkHighlightRaf = null;
    teleprompterView.addEventListener('scroll', () => {
        if (typeof checkTopPillAndGoToPreviousFile === 'function') checkTopPillAndGoToPreviousFile();
        if (typeof checkBottomPillAndAdvanceToNextFile === 'function') checkBottomPillAndAdvanceToNextFile();
        if (bookmarkHighlightRaf) return;
        bookmarkHighlightRaf = requestAnimationFrame(() => {
            bookmarkHighlightRaf = null;
            if (typeof updateBookmarkHighlightFromScroll === 'function') updateBookmarkHighlightFromScroll();
        });
    }, { passive: true });
    teleprompterView.addEventListener('click', (e) => {
        if (typeof getBookmarkIndexAtY === 'function' && typeof setActiveBookmarkByIndex === 'function') {
            const idx = getBookmarkIndexAtY(e.clientY);
            if (idx >= 0) setActiveBookmarkByIndex(idx);
        }
    });

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

    function sendPendingMirrorPosition() {
        if (!mirrorWindow || mirrorWindow.closed || !pendingMirrorPosition) return;
        try {
            mirrorWindow.postMessage({
                type: 'setPosition',
                left: pendingMirrorPosition.left,
                top: pendingMirrorPosition.top,
                width: pendingMirrorPosition.width,
                height: pendingMirrorPosition.height
            }, '*');
        } catch (_) {}
    }
    function handleMirrorReady() {
        if (pendingMirrorPosition) {
            sendPendingMirrorPosition();
            setTimeout(sendPendingMirrorPosition, 100);
            setTimeout(sendPendingMirrorPosition, 300);
            setTimeout(sendPendingMirrorPosition, 600);
        }
        syncColumnWidths();
        refreshMirrorData();
        syncMirrorStyles();
    }

    let overflowReportCount = 0;
    function handleRowOverflowReport(data) {
        if (!mirrorWindow || mirrorWindow.closed || !Array.isArray(data.rowHeights)) return;
        if (overflowReportCount >= 3) return; /* prevent infinite loop */
        const mainHeights = measuredRowHeights;
        const merged = data.rowHeights.map((h, i) => Math.max(mainHeights[i] || 0, h || 0, 1));
        const changed = measuredRowHeights.length !== merged.length ||
            merged.some((h, i) => Math.abs((measuredRowHeights[i] || 0) - h) > 2);
        if (!changed) return;
        overflowReportCount++;
        measuredRowHeights = merged;
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

    /** No-op: mirror now uses same table + hidden columns, so no row height sync. */
    function handleRowHeightsReport() {}

    try {
        const mirrorChannel = new BroadcastChannel('teleprompter-mirror-sync');
        mirrorChannel.onmessage = (e) => {
            try {
                const data = e.data || {};
                if (data.type === 'mirrorReady') handleMirrorReady();
                if (data.type === 'mirrorScroll') handleMirrorScroll(data);
                if (data.type === 'rowOverflowReport') handleRowOverflowReport(data);
                if (data.type === 'rowHeightsReport') handleRowHeightsReport(data);
            } catch (err) { console.warn('Mirror channel message error:', err); }
        };
    } catch (e) { /* BroadcastChannel not supported */ }

    window.addEventListener('message', (e) => {
        try {
            const data = e.data || {};
            if (data.type === 'mirrorReady') handleMirrorReady();
            if (data.type === 'mirrorScroll') handleMirrorScroll(data);
            if (data.type === 'rowOverflowReport') handleRowOverflowReport(data);
            if (data.type === 'rowHeightsReport') handleRowHeightsReport(data);
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
    function setupPanelCloseOnLeave(panel, button) {
        if (!panel) return;
        panel.addEventListener('mouseleave', (e) => {
            const to = e.relatedTarget;
            if (to && button && button.contains(to)) return;
            if (to && panel.contains(to)) return;
            panel.style.display = 'none';
        });
    }
    setupPanelCloseOnLeave(bgColorPanel, bgColorButton);
    setupPanelCloseOnLeave(fontColorPanel, fontColorButton);
    document.addEventListener('click', (e) => {
        if (bgColorPanel?.style.display === 'block' && !bgColorPanel.contains(e.target) && !bgColorButton?.contains(e.target)) {
            bgColorPanel.style.display = 'none';
        }
        if (fontColorPanel?.style.display === 'block' && !fontColorPanel.contains(e.target) && !fontColorButton?.contains(e.target)) {
            fontColorPanel.style.display = 'none';
        }
    });

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

    /** Measures row height from line count × line-height (formula-based). Uses tight line-height and full width to avoid over-measurement. */
    function measureRowHeightsFromContent() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) {
            if (typeof updateBookmarkPositions === 'function') updateBookmarkPositions();
            return;
        }
        if (typeof applyRowFont12 === 'function') applyRowFont12();
        const mainStyle = window.getComputedStyle(teleprompterText);
        const fontSize = parseFloat(mainStyle.fontSize) || 48;
        /* Use tight multiplier when broadcasting to match actual rendered line-height and avoid gap */
        const lh = document.body.classList.contains('broadcasting') ? 1.0 : (parseFloat(mainStyle.lineHeight) || 1.4);
        const oneLinePx = Math.floor(lh < 3 ? fontSize * lh : lh);
        const font12LinePx = Math.floor(12 * 1.15);

        measuredRowHeights = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            let maxH = 0;
            cols.forEach(col => {
                const w = col.getBoundingClientRect().width;
                if (w <= 0) return;
                const chars = (col.innerText || '').trim().length || 1;
                const cpl = getCharsPerLine(Math.max(20, w));
                const numLines = Math.max(1, Math.ceil(chars / cpl));
                const isFont12 = row.classList.contains('row-font-12');
                const linePx = isFont12 ? font12LinePx : oneLinePx;
                maxH = Math.max(maxH, numLines * linePx);
            });
            return Math.max(1, Math.floor(maxH));
        });
        updateBookmarkPositions();
    }

    /**
     * Build the single row-height array for extended view: for each row, use max height of visible content columns (col 2 and col 3).
     * Col 1 (ID) does not affect row height.
     */
    function buildRowHeightArrayForExtendedView() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        if (typeof applyRowFont12 === 'function') applyRowFont12();
        const mainStyle = window.getComputedStyle(teleprompterText);
        const fontSize = parseFloat(mainStyle.fontSize) || 48;
        const lh = 1.0;
        const oneLinePx = Math.floor(fontSize * lh);
        const font12LinePx = Math.floor(12 * 1.15);
        const viewportMax = Math.min(2000, (window.innerHeight || 800) * 2);
        const firstRow = rows[0];
        const cols = firstRow.querySelectorAll('.script-column');
        const numCols = cols.length;
        if (numCols < 1) return;
        const contentIndices = getVisibleContentColumnIndices(numCols);
        const cplPerCol = contentIndices.map(idx => {
            const col = cols[idx];
            const w = col ? col.getBoundingClientRect().width : 100;
            return getCharsPerLine(Math.max(20, w));
        });
        measuredRowHeights = rows.map(row => {
            const rowCols = row.querySelectorAll('.script-column');
            let maxH = 1;
            contentIndices.forEach((idx, k) => {
                const col = rowCols[idx];
                const chars = col ? (col.innerText || '').trim().length || 1 : 1;
                const cpl = cplPerCol[k] || 20;
                const numLines = Math.max(1, Math.ceil(chars / cpl));
                const linePx = row.classList.contains('row-font-12') ? font12LinePx : oneLinePx;
                const h = numLines * linePx;
                const floorH = Math.min(viewportMax, Math.max(1, Math.floor(h)));
                if (floorH > maxH) maxH = floorH;
            });
            return maxH;
        });
        if (typeof updateBookmarkPositions === 'function') updateBookmarkPositions();
    }

    /** Same approach as when file is opened then extended: measure row heights with probe; use max of visible content columns (col 2 and col 3). */
    function measureRowHeightsWithProbeForBroadcasting() {
        if (!document.body.classList.contains('broadcasting')) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        const scriptColWidth = lastColumnWidthPx != null && lastColumnWidthPx > 0 ? lastColumnWidthPx : 400;
        const mainStyle = window.getComputedStyle(teleprompterText);
        const probe = document.createElement('div');
        probe.style.position = 'absolute';
        probe.style.left = '-9999px';
        probe.style.top = '0';
        probe.style.width = scriptColWidth + 'px';
        probe.style.fontSize = mainStyle.fontSize || '';
        probe.style.fontFamily = mainStyle.fontFamily || '';
        probe.style.lineHeight = '1.1';
        probe.style.whiteSpace = 'pre-wrap';
        probe.style.wordWrap = 'break-word';
        probe.style.padding = '0';
        probe.style.margin = '0';
        probe.style.boxSizing = 'border-box';
        document.body.appendChild(probe);
        const numCols = rows[0] ? rows[0].querySelectorAll('.script-column').length : 0;
        const contentIndices = getVisibleContentColumnIndices(numCols);
        measuredRowHeights = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            let maxH = 1;
            contentIndices.forEach(idx => {
                const col = cols[idx];
                if (col && scriptColWidth > 0) {
                    const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content') || col;
                    const text = cell ? (cell.innerText || '').trim() || '\u00A0' : '\u00A0';
                    probe.textContent = text;
                    probe.style.fontSize = row.classList.contains('row-font-12') ? '12px' : (mainStyle.fontSize || '');
                    probe.style.lineHeight = row.classList.contains('row-font-12') ? '1.2' : '1.1';
                    const h = Math.max(1, probe.scrollHeight || 0);
                    if (h > maxH) maxH = h;
                }
            });
            return maxH;
        });
        probe.remove();
        rows.forEach((row, i) => {
            const h = measuredRowHeights[i];
            if (h > 0) {
                row.style.minHeight = h + 'px';
                row.style.height = h + 'px';
            }
        });
        if (mirrorWindow && !mirrorWindow.closed) {
            mirrorWindow.postMessage({ type: 'updateRowHeights', rowHeights: measuredRowHeights }, '*');
        }
        if (typeof updateBookmarkPositions === 'function') updateBookmarkPositions();
    }

    /** After extend: temporarily let rows size to content, measure actual height (max of visible content cols), then apply. Reduces gap under text. */
    function reassessRowHeightsFromActualLayout() {
        if (!document.body.classList.contains('broadcasting')) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        const numCols = rows[0] ? rows[0].querySelectorAll('.script-column').length : 0;
        const contentIndices = getVisibleContentColumnIndices(numCols);
        rows.forEach(row => {
            const cols = row.querySelectorAll('.script-column');
            contentIndices.forEach(idx => { if (cols[idx]) cols[idx].classList.remove('broadcast-hidden'); });
        });
        void teleprompterText.offsetHeight;
        rows.forEach(row => {
            row.style.alignItems = 'flex-start';
            row.style.height = 'auto';
            row.style.minHeight = '0';
            row.querySelectorAll('.script-column').forEach(col => {
                col.style.flex = '0 0 auto';
                col.style.minHeight = '0';
                const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content');
                if (cell) {
                    cell.style.flex = '0 0 auto';
                    cell.style.height = 'auto';
                }
            });
        });
        void teleprompterText.offsetHeight;
        const viewportMax = Math.min(2000, (window.innerHeight || 800) * 2);
        const mainActualHeights = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            let maxSh = 1;
            contentIndices.forEach(idx => {
                const col = cols[idx];
                const el = col ? (col.querySelector('.cell-locker') || col.querySelector('.cell-content') || col) : null;
                const sh = el ? (el.scrollHeight || 0) : 0;
                if (sh > maxSh) maxSh = sh;
            });
            return Math.max(1, Math.min(Math.floor(maxSh || 1), viewportMax));
        });
        /* Merge with mirror reported so both views get same height */
        if (mirrorReportedRowHeights.length === rows.length) {
            measuredRowHeights = mainActualHeights.map((mh, i) => {
                const mirrorH = mirrorReportedRowHeights[i];
                const v = Math.max(mh, (typeof mirrorH === 'number' && mirrorH > 0 ? mirrorH : 0), 1);
                return Math.min(v, viewportMax);
            });
        } else {
            measuredRowHeights = mainActualHeights.slice();
        }
        rows.forEach((row, i) => {
            const h = measuredRowHeights[i];
            row.style.minHeight = h + 'px';
            row.style.height = h + 'px';
            row.style.alignItems = '';
            row.querySelectorAll('.script-column').forEach(col => {
                col.style.flex = '';
                col.style.minHeight = '';
                const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content');
                if (cell) {
                    cell.style.flex = '';
                    cell.style.height = '';
                }
            });
        });
        applyBroadcastingVisibility();
        syncColumnWidths();
        if (mirrorWindow && !mirrorWindow.closed) {
            mirrorWindow.postMessage({ type: 'updateRowHeights', rowHeights: measuredRowHeights }, '*');
        }
        if (typeof updateBookmarkPositions === 'function') updateBookmarkPositions();
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

    const ROW_COLOR_CLASSES = ['row-lines-same', 'row-col2-more', 'row-col3-more', 'row-col3-much-more', 'row-col2-one-more', 'row-col3-one-more'];

    function applyRowColors() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;

        const maxCols = Math.max(...rows.map(r => r.querySelectorAll('.script-column').length));
        if (maxCols !== 3) {
            rows.forEach(row => ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c)));
            rowColorsCache = [];
            return;
        }

        const firstRow = rows[0];
        const r0cols = firstRow.querySelectorAll('.script-column');
        /* When col3 is unchecked by user (runlist), we are not comparing — clear all red/green. */
        if (r0cols[2]?.classList?.contains('user-col-hidden')) {
            rows.forEach(row => ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c)));
            rowColorsCache = [];
            return;
        }
        /* In extend mode one content column has broadcast-hidden on main; temporarily show it so we can read text and recompute colors. */
        let broadcastHiddenColIndex = -1;
        for (let i = 1; i < r0cols.length; i++) {
            if (r0cols[i]?.classList?.contains('broadcast-hidden')) {
                broadcastHiddenColIndex = i;
                break;
            }
        }
        const broadcastHiddenElements = broadcastHiddenColIndex >= 0 ? rows.map(row => row.querySelectorAll('.script-column')[broadcastHiddenColIndex]).filter(Boolean) : [];
        if (broadcastHiddenElements.length > 0) {
            broadcastHiddenElements.forEach(el => el.classList.remove('broadcast-hidden'));
            void teleprompterText.offsetHeight;
        }

        /* Left text column = col2 (green when more), right text column = col3 (red when more). DOM order: index 0 = numbers, 1 = left, 2 = right. */
        const WORD_DIFF_GLOW = 5;   /* word count difference >= this → glow */
        const WORD_DIFF_FILL = 10;  /* word count difference >= this → fill */
        const idxCol2 = Math.max(0, maxCols - 2);  /* left = col2 → green when more */
        const idxCol3 = maxCols - 1;               /* right = col3 → red when more */

        const debugRowColors = typeof window !== 'undefined' && window.DEBUG_ROW_COLORS;
        if (debugRowColors && rows.length > 0) {
            console.log('[RowColors] DEBUG on. We compare idxCol2=' + idxCol2 + ' vs idxCol3=' + idxCol3 + '. Col3 more => red. Check "byIndex" to see which DOM index has your column-3 text.');
        }

        rowColorsCache = [];
        rows.forEach((row, rowIndex) => {
            ROW_COLOR_CLASSES.forEach(c => row.classList.remove(c));
            const cols = Array.from(row.children).filter(c => c.classList?.contains('script-column'));
            const n = cols.length;
            if (n < 2) {
                row.classList.add('row-lines-same');
                rowColorsCache.push('row-lines-same');
                return;
            }
            const col2 = cols[Math.min(idxCol2, n - 1)];
            const col3 = cols[Math.min(idxCol3, n - 1)];
            const text2 = (col2?.textContent ?? '').toString().trim();
            const text3 = (col3?.textContent ?? '').toString().trim();

            const words2 = (text2.split(/\s+/).filter(Boolean)).length;
            const words3 = (text3.split(/\s+/).filter(Boolean)).length;
            const wordDiff = Math.abs(words2 - words3);
            let cls;
            if (words2 === words3 || wordDiff < WORD_DIFF_GLOW) {
                cls = 'row-lines-same';
            } else if (words2 > words3) {
                cls = wordDiff >= WORD_DIFF_FILL ? 'row-col2-more' : 'row-col2-one-more';
            } else {
                cls = wordDiff >= WORD_DIFF_FILL ? 'row-col3-much-more' : 'row-col3-one-more';
            }
            row.classList.add(cls);
            rowColorsCache.push(cls);

            /* Troubleshooting: when DEBUG_ROW_COLORS is true, log one full report for row 0 and whenever we assign green. */
            if (debugRowColors && (rowIndex === 0 || cls.startsWith('row-col2'))) {
                const byIndex = cols.map((c, i) => {
                    const t = (c?.textContent ?? '').toString().trim();
                    return { i, chars: t.length, preview: t.slice(0, 30) + (t.length > 30 ? '…' : '') };
                });
                const triggerLabels = {
                    'row-lines-same': 'same (diff < 5 words)',
                    'row-col2-one-more': 'green glow (col2 has 5–9 more words)',
                    'row-col2-more': 'green fill (col2 has 10+ more words)',
                    'row-col3-one-more': 'red glow (col3 has 5–9 more words)',
                    'row-col3-more': 'red fill (col3 has 10+ more words)',
                    'row-col3-much-more': 'red fill (col3 has 10+ more words)'
                };
                const cssTrigger = triggerLabels[cls] || cls;
                console.log('[RowColors]', rowIndex === 0 ? 'Sample row 0 (all columns)' : 'GREEN row ' + rowIndex, {
                    numCols: n,
                    idxCol2,
                    idxCol3,
                    byIndex,
                    words2,
                    words3,
                    wordDiff,
                    cssClass: cls,
                    cssTrigger
                });
            }
        });
        if (broadcastHiddenElements.length > 0) {
            broadcastHiddenElements.forEach(el => el.classList.add('broadcast-hidden'));
        }
    }

    /** When applying 12px or clearing, skip elements that have a user-set font-size (e.g. from font size dropdown in Aa mode). */
    function setRowFontSizeIfNotUserSet(el, fontSizeVal, lineHeightVal) {
        const cur = el.style.fontSize;
        if (cur && cur !== '12px' && cur !== '') return;
        el.style.fontSize = fontSizeVal;
        el.style.lineHeight = lineHeightVal || '';
    }

    function applyRowFont12() {
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        if (isCurrentFileXlsx()) return;
        const isBroadcasting = document.body.classList.contains('broadcasting');
        if (isBroadcasting && rowFont12Cache.length === rows.length) {
            rows.forEach((row, i) => {
                if (rowFont12Cache[i]) {
                    row.classList.add('row-font-12');
                    setRowFontSizeIfNotUserSet(row, '12px', '1.2');
                    row.querySelectorAll('.script-column, .cell-content, .cell-locker').forEach(el => {
                        setRowFontSizeIfNotUserSet(el, '12px', '1.2');
                    });
                    row.querySelectorAll('.cell-content span, .cell-content div, .cell-content p').forEach(el => {
                        setRowFontSizeIfNotUserSet(el, '12px', '1.2');
                    });
                } else {
                    row.classList.remove('row-font-12');
                    setRowFontSizeIfNotUserSet(row, '', '');
                    row.querySelectorAll('.script-column, .cell-content, .cell-locker').forEach(el => {
                        setRowFontSizeIfNotUserSet(el, '', '');
                    });
                    row.querySelectorAll('.cell-content span, .cell-content div, .cell-content p').forEach(el => {
                        setRowFontSizeIfNotUserSet(el, '', '');
                    });
                }
            });
            return;
        }
        rowFont12Cache = [];
        rows.forEach(row => {
            const text = (row.innerText || '').trim();
            const hasPipeV = text.includes('|v');
            if (hasPipeV) {
                row.classList.add('row-font-12');
                setRowFontSizeIfNotUserSet(row, '12px', '1.2');
                row.querySelectorAll('.script-column, .cell-content, .cell-locker').forEach(el => {
                    setRowFontSizeIfNotUserSet(el, '12px', '1.2');
                });
                row.querySelectorAll('.cell-content span, .cell-content div, .cell-content p').forEach(el => {
                    setRowFontSizeIfNotUserSet(el, '12px', '1.2');
                });
                rowFont12Cache.push(true);
            } else {
                row.classList.remove('row-font-12');
                setRowFontSizeIfNotUserSet(row, '', '');
                row.querySelectorAll('.script-column, .cell-content, .cell-locker').forEach(el => {
                    setRowFontSizeIfNotUserSet(el, '', '');
                });
                row.querySelectorAll('.cell-content span, .cell-content div, .cell-content p').forEach(el => {
                    setRowFontSizeIfNotUserSet(el, '', '');
                });
                rowFont12Cache.push(false);
            }
        });
    }

    /** When in Aa (broadcast edit) mode: let main rows size to content, measure heights as max of visible content columns (col 2 and col 3), then apply. */
    function syncRowHeightsFromMainInBroadcastEditMode() {
        if (!document.body.classList.contains('broadcasting') || !broadcastEditMode) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length === 0) return;
        rows.forEach(row => {
            row.style.height = 'auto';
            row.style.minHeight = '0';
            row.querySelectorAll('.script-column').forEach(col => {
                const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content');
                if (cell) {
                    cell.style.height = 'auto';
                }
            });
        });
        void teleprompterText.offsetHeight;
        const viewportMax = Math.min(2000, (window.innerHeight || 800) * 2);
        const numCols = rows[0] ? rows[0].querySelectorAll('.script-column').length : 0;
        const contentIndices = getVisibleContentColumnIndices(numCols);
        measuredRowHeights = rows.map(row => {
            const cols = row.querySelectorAll('.script-column');
            let maxH = 1;
            contentIndices.forEach(idx => {
                const col = cols[idx];
                const cell = col ? (col.querySelector('.cell-locker') || col.querySelector('.cell-content')) : null;
                const h = cell ? cell.getBoundingClientRect().height : 0;
                if (h > maxH) maxH = h;
            });
            return Math.max(1, Math.min(Math.floor(maxH), viewportMax));
        });
        rows.forEach((row, i) => {
            const h = measuredRowHeights[i];
            if (h > 0) {
                row.style.minHeight = h + 'px';
                row.style.height = h + 'px';
            }
            row.querySelectorAll('.script-column').forEach(col => {
                const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content');
                if (cell) cell.style.height = '';
            });
        });
        if (mirrorWindow && !mirrorWindow.closed) {
            mirrorWindow.postMessage({ type: 'updateRowHeights', rowHeights: measuredRowHeights }, '*');
        }
    }

    /** Index of the last visible column (per file checkboxes). Mirror always shows this column. */
    function getLastVisibleColumnIndex(maxCols) {
        const vis = currentFileIndex >= 0 && fileColumnVisibility[currentFileIndex] ? fileColumnVisibility[currentFileIndex] : null;
        if (!vis || vis.length === 0) return maxCols > 0 ? maxCols - 1 : 0;
        let last = -1;
        for (let i = 0; i < vis.length; i++) if (vis[i] !== false) last = i;
        return last >= 0 ? last : (maxCols > 0 ? maxCols - 1 : 0);
    }

    /** Number of visible columns (per file checkboxes). When 1, main and mirror show the same column. */
    function getVisibleColumnCount(maxCols) {
        const vis = currentFileIndex >= 0 && fileColumnVisibility[currentFileIndex] ? fileColumnVisibility[currentFileIndex] : null;
        if (!vis || vis.length === 0) return maxCols;
        return vis.filter(v => v !== false).length;
    }

    /** Visible content column indices (>= 1). Row height in broadcast = max of these columns' heights. */
    function getVisibleContentColumnIndices(maxCols) {
        const vis = currentFileIndex >= 0 && fileColumnVisibility[currentFileIndex] ? fileColumnVisibility[currentFileIndex] : null;
        if (!vis || vis.length === 0) return Array.from({ length: Math.max(0, maxCols - 1) }, (_, i) => i + 1);
        return [...vis].map((v, i) => (v !== false && i >= 1 ? i : -1)).filter(i => i >= 0);
    }

    function refreshMirrorData() {
        if (!mirrorWindow || mirrorWindow.closed) return;
        const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        if (rows.length > 0) {
            applyRowFont12();
        }
        const contentWidth = (lastColumnWidthPx && lastColumnWidthPx > 50) ? `${lastColumnWidthPx}px` : null;
        const isBroadcasting = document.body.classList.contains('broadcasting');
        if (isBroadcasting && broadcastEditMode && rows.length > 0) {
            if (!preserveRowHeightsAfterExtend) {
                syncRowHeightsFromMainInBroadcastEditMode();
            } else {
                preserveRowHeightsAfterExtend = false;
            }
            /* Recompute glow/fill for current column widths so main and mirror stay in sync */
            buildCharCountsPerRow();
            applyRowColors();
        }
        /* Copy full table to mirror (all columns); mirror shows the last visible column (per file checkboxes). */
        const topPill = (currentFileIndex >= 0 && fileStore[currentFileIndex]) ? stripFileExtension(fileStore[currentFileIndex].name) : '';
        const bottomPill = (currentFileIndex >= 0 && currentFileIndex + 1 < fileStore.length && fileStore[currentFileIndex + 1]) ? stripFileExtension(fileStore[currentFileIndex + 1].name) : '';
        if (rows.length > 0 && isBroadcasting) {
            const maxCols = Math.max(...rows.map(r => r.querySelectorAll('.script-column').length));
            const visibleColumnIndex = getLastVisibleColumnIndex(maxCols);
            const rowData = rows.map(row => {
                const cols = row.querySelectorAll('.script-column');
                const cells = Array.from(cols).map(col => {
                    const c = col.querySelector('.cell-content') || col;
                    return (c.innerText || '').trim() || "\u00A0";
                });
                const colorClass = ROW_COLOR_CLASSES.find(c => row.classList.contains(c)) || '';
                const keywordPill = ['keyword-pill-red', 'keyword-pill-yellow', 'keyword-pill-green', 'keyword-pill-blue', 'keyword-pill-white'].find(c => row.classList.contains(c)) || '';
                const rowClass = [colorClass, keywordPill].filter(Boolean).join(' ');
                const font12 = row.classList.contains('row-font-12');
                return { cells, rowClass, font12 };
            });
            const rowColors = (rowColorsCache.length === rows.length) ? rowColorsCache : rows.map(row => ROW_COLOR_CLASSES.find(c => row.classList.contains(c)) || '');
            const rowFont12 = (rowFont12Cache.length === rows.length) ? rowFont12Cache : rows.map(row => row.classList.contains('row-font-12'));
            const rowHeights = (measuredRowHeights.length === rows.length) ? measuredRowHeights : null;
            mirrorWindow.postMessage({
                type: 'loadContent',
                table: rowData,
                visibleColumnIndex,
                contentWidth,
                rowColors,
                rowFont12,
                rowHeights,
                topPill,
                bottomPill
            }, '*');
        } else if (rows.length > 0) {
            const rowData = rows.map(row => {
                const cols = row.querySelectorAll('.script-column');
                const cells = Array.from(cols).map(col => {
                    const c = col.querySelector('.cell-content') || col;
                    return (c.innerText || '').trim() || "\u00A0";
                });
                const colorClass = ROW_COLOR_CLASSES.find(c => row.classList.contains(c)) || '';
                const keywordPill = ['keyword-pill-red', 'keyword-pill-yellow', 'keyword-pill-green', 'keyword-pill-blue', 'keyword-pill-white'].find(c => row.classList.contains(c)) || '';
                const rowClass = [colorClass, keywordPill].filter(Boolean).join(' ');
                const font12 = row.classList.contains('row-font-12');
                return { cells, rowClass, font12 };
            });
            const maxCols = Math.max(...rows.map(r => r.querySelectorAll('.script-column').length));
            const visibleColumnIndex = getLastVisibleColumnIndex(maxCols);
            const rowColors = rows.map(row => ROW_COLOR_CLASSES.find(c => row.classList.contains(c)) || '');
            const rowFont12 = rows.map(row => row.classList.contains('row-font-12'));
            mirrorWindow.postMessage({
                type: 'loadContent',
                table: rowData,
                visibleColumnIndex,
                contentWidth,
                rowColors,
                rowFont12,
                topPill,
                bottomPill
            }, '*');
        } else {
            const fullText = teleprompterText.innerText || '';
            const fallback = fullText ? fullText.split(/\r?\n/).map(line => line.trim() || "\u00A0") : ["\u00A0"];
            mirrorWindow.postMessage({
                type: 'loadContent',
                table: fallback.map(t => ({ cells: [t], rowClass: '', font12: false })),
                visibleColumnIndex: 0,
                contentWidth: null,
                rowColors: [],
                rowFont12: [],
                topPill,
                bottomPill
            }, '*');
        }
    }

    function applyBroadcastingVisibility() {
        const isBroadcasting = document.body.classList.contains('broadcasting');
        const rows = teleprompterText.querySelectorAll('.script-row-wrapper');
        const maxCols = rows.length > 0 ? Math.max(...Array.from(rows).map(r => r.querySelectorAll('.script-column').length)) : 0;
        const vis = currentFileIndex >= 0 && fileColumnVisibility[currentFileIndex] ? fileColumnVisibility[currentFileIndex] : null;
        const visibleIndices = vis ? [...vis].map((v, i) => (v !== false ? i : -1)).filter(i => i >= 0) : null;
        const contentVisibleCount = visibleIndices ? visibleIndices.filter(i => i >= 1).length : 0;
        const lastVisible = isBroadcasting && maxCols > 0 ? getLastVisibleColumnIndex(maxCols) : -1;
        const hideLastOnMain = isBroadcasting && contentVisibleCount >= 2 && lastVisible >= 0;
        rows.forEach(row => {
            const cols = row.querySelectorAll('.script-column');
            cols.forEach(col => col.classList.remove('broadcast-hidden'));
            /* When 2+ content columns visible: main shows all except the last; mirror shows the last. When only 1 content column: main shows same as mirror. */
            if (hideLastOnMain && cols[lastVisible]) {
                cols[lastVisible].classList.add('broadcast-hidden');
            }
        });
        if (isBroadcasting) {
            /* Keep locked heights (set by extend flow); only clear when leaving broadcasting */
        } else {
            broadcastEditMode = false;
            teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
                row.style.minHeight = '';
                row.style.height = '';
            });
        }
    }

    /** Restore main window to show all columns when mirror is closed (e.g. user hits X on extended window). */
    function restoreMainWindowFromMirrorClose() {
        if (mirrorCloseCheckInterval) {
            clearInterval(mirrorCloseCheckInterval);
            mirrorCloseCheckInterval = null;
        }
        mirrorWindow = null;
        pendingMirrorPosition = null;
        broadcastEditMode = false;
        measuredRowHeights = [];
        extendedFixedWidth = null;
        extendedWindowWidth = null;
        extendedWindowHeight = null;
        teleprompterText.style.width = '';
        teleprompterText.style.maxWidth = '';
        document.body.classList.remove('broadcasting');
        applyBroadcastingVisibility();
        syncColumnWidths();
        const btnOverview = document.getElementById('btn-overview-toggle');
        if (btnOverview) {
            btnOverview.classList.remove('broadcast-edit-active');
            btnOverview.title = 'Toggle overview (size 12) / restore size';
        }
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
        const tf = window.getComputedStyle(teleprompterText);
        sample.style.fontFamily = tf.fontFamily;
        sample.style.fontSize = isOverviewMode ? '12px' : tf.fontSize;
        sample.style.lineHeight = isOverviewMode ? '1.2' : tf.lineHeight;
        document.body.appendChild(sample);
        const digitsWidth = sample.getBoundingClientRect().width;
        sample.remove();

        const firstRow = rows[0];
        const firstCell = firstRow?.querySelector('.script-column .cell-locker') || firstRow?.querySelector('.script-column .cell-content') || firstRow?.querySelector('.script-column');
        if (firstCell) cellStyle = window.getComputedStyle(firstCell);

        let paddingWidth = 0;
        if (cellStyle) {
            paddingWidth = (parseFloat(cellStyle.paddingLeft) || 0) + (parseFloat(cellStyle.paddingRight) || 0);
        }
        widths[0] = Math.ceil(digitsWidth + paddingWidth);

        const isBroadcasting = document.body.classList.contains('broadcasting');
        const isXlsx = isCurrentFileXlsx();

        /* XLSX: col1 shrink-to-fit, cols 2 and 3 equal width */
        if (isXlsx && maxCols >= 3 && !isBroadcasting) {
            const probe = document.createElement('span');
            probe.style.position = 'absolute';
            probe.style.visibility = 'hidden';
            probe.style.whiteSpace = 'pre';
            const tf = window.getComputedStyle(teleprompterText);
            probe.style.fontFamily = tf.fontFamily;
            probe.style.fontSize = isOverviewMode ? '12px' : tf.fontSize;
            probe.style.lineHeight = isOverviewMode ? '1.2' : tf.lineHeight;
            document.body.appendChild(probe);
            let col0Max = 0;
            rows.forEach(row => {
                const cols = row.querySelectorAll('.script-column');
                const cell = cols[0]?.querySelector('.cell-content') || cols[0]?.querySelector('.cell-locker') || cols[0];
                const txt = (cell?.textContent || '').trim() || '0';
                probe.textContent = txt;
                col0Max = Math.max(col0Max, probe.getBoundingClientRect().width);
            });
            probe.remove();
            /* Add buffer to prevent numbers wrapping from subpixel rounding */
            widths[0] = Math.ceil(Math.max(col0Max, digitsWidth) + paddingWidth) + 8;
            const xlsxRemaining = Math.max(0, (isBroadcasting && extendedFixedWidth != null && extendedFixedWidth > 0 ? extendedFixedWidth : teleprompterText.clientWidth) - widths[0]);
            const half = Math.floor(xlsxRemaining / 2);
            widths[1] = half;
            widths[2] = xlsxRemaining - half;
        }
        const nVisible = getVisibleColumnCount(maxCols);
        const vis = currentFileIndex >= 0 && fileColumnVisibility[currentFileIndex] ? fileColumnVisibility[currentFileIndex] : null;
        const visibleIndices = vis ? [...vis].map((v, i) => (v !== false ? i : -1)).filter(i => i >= 0) : null;
        const useVisibleColumns = (isBroadcasting && nVisible > 0) || (!isBroadcasting && visibleIndices && visibleIndices.length > 0);
        const visibleColCount = useVisibleColumns && visibleIndices ? visibleIndices.length : maxCols;
        const otherColCount = Math.max(0, visibleColCount - 1);

        /* When extended use fixed width so table doesn't change on window resize. */
        const availableWidth = (isBroadcasting && extendedFixedWidth != null && extendedFixedWidth > 0)
            ? extendedFixedWidth
            : teleprompterText.clientWidth;
        const remaining = Math.max(availableWidth - widths[0], 0);
        let equalWidth = otherColCount > 0 ? Math.floor(remaining / (visibleColCount > 0 ? visibleColCount : 1)) : 0;

        const useXlsxWidths = isXlsx && maxCols >= 3 && !isBroadcasting;
        if (!useXlsxWidths && useVisibleColumns && visibleIndices) {
            const contentVisibleCount = visibleIndices.filter(i => i >= 1).length;
            /* In extend mode with 2+ content columns: main shows all visible except the last (mirror shows the last). */
            const mainVisibleIndices = (isBroadcasting && contentVisibleCount >= 2)
                ? visibleIndices.slice(0, -1)
                : visibleIndices;
            const contentIndices = mainVisibleIndices.filter(j => j >= 1);
            const hasCol0 = mainVisibleIndices.indexOf(0) >= 0;
            if (hasCol0) {
                widths[0] = Math.ceil(digitsWidth + paddingWidth);
            } else {
                widths[0] = 0;
            }
            for (let i = 1; i < widths.length; i++) widths[i] = 0;
            const contentRemaining = Math.max(0, availableWidth - widths[0]);
            const nContent = contentIndices.length || 1;
            const contentEqual = Math.floor(contentRemaining / nContent);
            contentIndices.forEach((j, k) => {
                widths[j] = contentEqual;
                if (k === contentIndices.length - 1) widths[j] += contentRemaining - (contentEqual * contentIndices.length);
            });
        } else if (!useXlsxWidths) {
            for (let i = 1; i < widths.length; i += 1) {
                widths[i] = (i < visibleColCount) ? equalWidth : 0;
            }
            if (otherColCount > 0 && equalWidth > 0 && visibleColCount > 1) {
                widths[visibleColCount - 1] += remaining - (equalWidth * otherColCount);
            }
        }

        /* When broadcasting, mirror shows one column; set its width for mirror layout. */
        if (isBroadcasting && maxCols > 0 && nVisible > 0) {
            const lastIdx = getLastVisibleColumnIndex(maxCols);
            const contentRemainingForMirror = Math.max(0, availableWidth - (widths[0] || 0));
            lastColumnWidthPx = (lastIdx < widths.length && widths[lastIdx] > 0)
                ? widths[lastIdx]
                : (useVisibleColumns && visibleIndices && visibleIndices.length > 1 ? contentRemainingForMirror : Math.floor(remaining / nVisible));
        } else if (!isBroadcasting) {
            lastColumnWidthPx = (maxCols > 1 && rows[0]) ? rows[0].querySelectorAll('.script-column')[1]?.getBoundingClientRect().width || remaining / (maxCols - 1) : null;
        }

        if (maxCols >= 3 && rows[0] && !isBroadcasting) {
            const r0cols = rows[0].querySelectorAll('.script-column');
            col2WidthPx = r0cols[1]?.getBoundingClientRect().width || null;
            col3WidthPx = r0cols[2]?.getBoundingClientRect().width || col2WidthPx;
        }

        const lastVisibleIdx = useVisibleColumns && visibleIndices && visibleIndices.length > 0
            ? visibleIndices[visibleIndices.length - 1]
            : (maxCols > 0 ? maxCols - 1 : 0);
        rows.forEach(row => {
            const cols = row.querySelectorAll('.script-column');
            cols.forEach((col, idx) => {
                const isLastVisible = idx === lastVisibleIdx;
                const width = widths[idx] || 0;
                if (isBroadcasting) {
                    col.style.flex = width > 0 ? `0 0 ${width}px` : '';
                    col.style.width = width > 0 ? `${width}px` : '';
                    col.style.minWidth = '0';
                } else if (!isBroadcasting && isLastVisible) {
                    /* Last visible column grows to fill so col 2 expands when col 3 is hidden */
                    col.style.flex = `1 1 0%`;
                    col.style.minWidth = '0';
                    col.style.width = 'auto';
                } else {
                    col.style.flex = width > 0 ? `0 0 ${width}px` : '';
                    col.style.width = width > 0 ? `${width}px` : '';
                }
            });
        });
        applyBroadcastingVisibility();
        if (rows.length > 0) {
            buildCharCountsPerRow();
            applyRowColors();
            applyRowFont12();
        }
    }

    /** Sync editor state to arrays (contentStore, charCountsPerRow, rowColorsCache) and mirror. Call after user actions (blur, save, etc.). */
    function syncEditorState() {
        if (typeof stripPipeVFromFirstColumn === 'function') stripPipeVFromFirstColumn();
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

    const isToolbarOrPanel = (el) => {
        if (!el) return false;
        const header = document.querySelector('.top-ribbon');
        return header?.contains(el) || bgColorPanel?.contains(el) || fontColorPanel?.contains(el) || runlistPanel?.contains(el);
    };
    teleprompterText.addEventListener('focusout', (e) => {
        const next = e.relatedTarget;
        if (isToolbarOrPanel(next)) return;
        syncEditorState();
    });
    document.addEventListener('click', (e) => {
        if (isToolbarOrPanel(e.target)) return;
        if (!teleprompterText.contains(e.target)) syncEditorState();
    });

    let inputColorDebounce;
    teleprompterText.addEventListener('input', () => {
        clearTimeout(inputColorDebounce);
        inputColorDebounce = setTimeout(() => {
            if (typeof stripPipeVFromFirstColumn === 'function') stripPipeVFromFirstColumn();
            refreshSelectFontTarget();
            applyKeywordPills();
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

if (undoButton) undoButton.onclick = () => { if (!undo()) document.execCommand('undo'); };
if (redoButton) redoButton.onclick = () => { if (!redo()) document.execCommand('redo'); };

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
    const countBefore = fileStore.length;
    files.forEach(file => addFileToRunlist(file));
    if (countBefore > 0 && fileStore.length > countBefore) {
        sortRunlistIfNumericPrefix(countBefore);
    }
    fileOpener.value = "";
};

function moveFileInRunlist(direction) {
    if (currentFileIndex < 0 || fileStore.length < 2) return;
    const newIndex = direction === 'up' ? currentFileIndex - 1 : currentFileIndex + 1;
    if (newIndex < 0 || newIndex >= fileStore.length) return;
    [fileStore[currentFileIndex], fileStore[newIndex]] = [fileStore[newIndex], fileStore[currentFileIndex]];
    [contentStore[currentFileIndex], contentStore[newIndex]] = [contentStore[newIndex], contentStore[currentFileIndex]];
    [fileColumnCount[currentFileIndex], fileColumnCount[newIndex]] = [fileColumnCount[newIndex], fileColumnCount[currentFileIndex]];
    [fileColumnVisibility[currentFileIndex], fileColumnVisibility[newIndex]] = [fileColumnVisibility[newIndex], fileColumnVisibility[currentFileIndex]];
    [fileFirstColIsId[currentFileIndex], fileFirstColIsId[newIndex]] = [fileFirstColIsId[newIndex], fileFirstColIsId[currentFileIndex]];
    const rows = Array.from(runlistContainer.querySelectorAll('.runlist-row'));
    [rows[currentFileIndex], rows[newIndex]] = [rows[newIndex], rows[currentFileIndex]];
    runlistContainer.innerHTML = '';
    rows.forEach((row, i) => {
        row.dataset.index = i;
        const nameEl = row.querySelector('.file-name');
        if (nameEl) nameEl.textContent = fileStore[i].name;
        row.onclick = (e) => {
            if (runlistJustDragged) { runlistJustDragged = false; return; }
            if (e.target.closest('.runlist-column-toggles') || e.target.closest('.runlist-close') || e.target.closest('.runlist-drag-handle')) return;
            loadScriptToEditor(i);
        };
        const closeBtn = row.querySelector('.runlist-close');
        if (closeBtn) closeBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); removeFileFromRunlist(i); };
        attachRunlistRowDrag(row, i);
        runlistContainer.appendChild(row);
        updateRunlistRowColumnToggles(i);
    });
    currentFileIndex = newIndex;
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const activeRow = runlistContainer.querySelector(`.runlist-row[data-index="${newIndex}"]`);
    if (activeRow) activeRow.classList.add('active');
    resetPillTriggerState();
    updateFilenamePills();
}

if (btnMoveFileUp) btnMoveFileUp.onclick = () => moveFileInRunlist('up');
if (btnMoveFileDown) btnMoveFileDown.onclick = () => moveFileInRunlist('down');

if (toggleRunlistButton && runlistPanel && resizer) {
    toggleRunlistButton.onclick = () => {
        const isHidden = runlistPanel.classList.toggle('hidden');
        resizer.classList.toggle('hidden', isHidden);
        toggleRunlistButton.title = isHidden ? 'Show Runlist' : 'Hide Runlist';
    };
}

if (settingsGearButton && settingsOverlay) {
    settingsGearButton.onclick = () => {
        settingsOverlay.classList.remove('hidden');
    };
}
const btnCloseSettings = document.getElementById('btn-close-settings');
if (btnCloseSettings && settingsOverlay) {
    btnCloseSettings.onclick = () => settingsOverlay.classList.add('hidden');
}
if (settingsOverlay) {
    settingsOverlay.onclick = (e) => {
        if (e.target === settingsOverlay) settingsOverlay.classList.add('hidden');
    };
    const settingsPanelEl = settingsOverlay.querySelector('.settings-overlay-panel');
    if (settingsPanelEl) settingsPanelEl.onclick = (e) => e.stopPropagation();
}

const eventPlanOverlay = document.getElementById('event-plan-overlay');
const btnEventPlanPrint = document.getElementById('btn-event-plan-print');
const btnEventPlanClose = document.getElementById('btn-event-plan-close');
function openEventPlanOverlay() {
    if (!eventPlanOverlay) return;
    eventPlanOverlay.classList.remove('hidden');
}
function closeEventPlanOverlay() {
    if (eventPlanOverlay) eventPlanOverlay.classList.add('hidden');
}
if (btnEventPlanClose && eventPlanOverlay) {
    btnEventPlanClose.onclick = closeEventPlanOverlay;
}
if (eventPlanOverlay) {
    eventPlanOverlay.onclick = (e) => {
        if (e.target === eventPlanOverlay) closeEventPlanOverlay();
    };
    const eventPlanPanelEl = eventPlanOverlay.querySelector('.event-plan-panel');
    if (eventPlanPanelEl) eventPlanPanelEl.onclick = (e) => e.stopPropagation();
}

function parseEventPlanFilename(name) {
    if (!name || typeof name !== 'string') return null;
    const base = stripFileExtension(name);
    const parts = base.split('_');
    if (parts.length !== 7) return null;
    const [orderNum, interpreter, item, speaker, minutes, location, type] = parts;
    const order = orderNum.trim();
    const hasNumericOrder = /^\d+$/.test(order);
    return {
        file: hasNumericOrder ? order : '-',
        item: (item || '').trim(),
        assignment: (speaker || '').trim(),
        location: (location || '').trim(),
        interpreter: interpreter.trim(),
        tel: (type || '').trim(),
        hasNumericOrder
    };
}

function buildEventPlanHtml(files, title, callTime, runTime) {
    const rows = [];
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const parsed = parseEventPlanFilename(file.name);
        if (parsed) rows.push(parsed);
        else rows.push({
            file: '-',
            item: '-',
            assignment: stripFileExtension(file.name) || file.name,
            location: '-',
            interpreter: '-',
            tel: '-',
            hasNumericOrder: false
        });
    }
    const locationImgMap = { L: 'graphics/Left_Side.jpg', R: 'graphics/Right_side.jpg', LR: 'graphics/Middle.jpg' };
    const locationCell = (loc) => {
        if (!loc) return '<span>-</span>';
        const imgFile = locationImgMap[loc];
        const escaped = String(loc).replace(/</g, '&lt;').replace(/>/g, '&gt;');
        if (imgFile) {
            return '<img src="' + imgFile + '" alt="' + escaped + '" class="event-plan-loc-img" title="' + escaped + '">';
        }
        return '<span>' + escaped + '</span>';
    };
    const esc = (s) => (s == null || s === '') ? '' : String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const tableRows = rows.map(r => {
        const rowClass = r.hasNumericOrder ? 'event-plan-row-highlight' : '';
        return `<tr class="${rowClass}"><td>${esc(r.file)}</td><td>${esc(r.item)}</td><td>${esc(r.assignment)}</td><td>${locationCell(r.location)}</td><td>${esc(r.interpreter)}</td><td>${esc(r.tel)}</td></tr>`;
    }).join('');
    return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
.event-plan-print{font-family:Arial,sans-serif;color:#000;background:#fff;padding:0;margin:0;}
.event-plan-header{background:#1a365d;color:#fff;text-align:center;padding:20px;margin:0;}
.event-plan-header h1{font-size:24px;font-weight:bold;margin:0 0 8px 0;}
.event-plan-header .sub{font-size:16px;font-weight:bold;margin:4px 0;}
.event-plan-table{width:100%;border-collapse:collapse;margin:0;font-size:14px;}
.event-plan-table th{background:#fff;color:#000;border:1px solid #333;padding:8px 10px;text-align:left;font-weight:bold;}
.event-plan-table td{border:1px solid #333;padding:8px 10px;}
.event-plan-row-highlight{background:#d4e8f7;}
.event-plan-loc-img{max-width:24px;max-height:24px;vertical-align:middle;}
</style></head><body class="event-plan-print">
<div class="event-plan-header">
<h1>${(title || 'Event Plan').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</h1>
<div class="sub">${(callTime || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>
<div class="sub">${(runTime || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div>
</div>
<table class="event-plan-table">
<thead><tr><th>File</th><th>Item</th><th>Assignment</th><th>Location</th><th>Interpreter</th><th>TEL</th></tr></thead>
<tbody>${tableRows}</tbody>
</table>
</body></html>`;
}

/** Compute pill class for a row's raw text (for master script / print). switchLabel is 'STAY' or 'SWITCH'. cells=optional array of cell texts for per-cell matching (xlsx). */
function getPillClassForRow(raw, switchLabel, cells) {
    const lower = (raw || '').trim().toLowerCase();
    const r = (raw || '').trim();
    const check = (t) => {
        const L = (t || '').trim().toLowerCase();
        const T = (t || '').trim();
        if (KEYWORD_PILL_RED.includes(L)) return 'ms-pill-red';
        if (['switch', 'stay'].includes(L)) return 'ms-pill-green';
        if (KEYWORD_PILL_YELLOW_EXACT.includes(L) || KEYWORD_PILL_YELLOW_VERSE.test(T) || KEYWORD_PILL_YELLOW_NAME.test(T) || KEYWORD_PILL_YELLOW_BRACKETS.test(T)) return 'ms-pill-yellow';
        if (KEYWORD_PILL_BLUE_NAME_TIME.test(T)) return 'ms-pill-blue';
        if (T.includes('|v')) return 'ms-pill-white';
        return '';
    };
    let pill = check(r);
    if (!pill && cells && cells.length > 0) {
        for (let i = 0; i < cells.length; i++) {
            pill = check(cells[i]);
            if (pill) break;
        }
    }
    return pill;
}

/** Build printable master script: all files combined with pill/color format (Aa-mode style). */
function buildMasterScriptHtml() {
    const esc = (s) => (s == null || s === '') ? '' : String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    const sections = [];
    const temp = document.createElement('div');

    for (let fi = 0; fi < fileStore.length; fi++) {
        const file = fileStore[fi];
        const content = contentStore[fi];
        const name = stripFileExtension(file?.name || '');
        const fileExt = (file?.name || '').split('.').pop().toLowerCase();
        const isDocx = fileExt === 'docx' || fileExt === 'doc';
        const nextFile = fi + 1 < fileStore.length ? fileStore[fi + 1] : null;
        const currentInterpreter = file ? getInterpreterFromFilename(file.name) : '';
        const nextInterpreter = nextFile ? getInterpreterFromFilename(nextFile.name) : '';
        const switchLabel = (nextInterpreter && nextInterpreter !== currentInterpreter) ? 'SWITCH' : 'STAY';

        const sectionRows = [];

        if (!content || (typeof content === 'string' && !content.trim())) {
            sectionRows.push({ type: 'row', cells: ['\u00A0'], pill: '', font12: false, singleCol: true });
            sections.push({ header: name || 'Untitled', rows: sectionRows });
            continue;
        }

        temp.innerHTML = typeof content === 'string' ? content : '';
        let rowEls = temp.querySelectorAll('.script-row-wrapper');
        if (rowEls.length === 0) {
            /* Fallback: parse tables (e.g. from docx with multiple or nested tables) */
            const allTables = temp.querySelectorAll('table');
            const topLevelTables = Array.from(allTables).filter(t => !t.closest('table'));
            if (topLevelTables.length > 0) {
                const collected = [];
                topLevelTables.forEach(tbl => {
                    tbl.querySelectorAll('tr').forEach(tr => {
                        let cells = Array.from(tr.querySelectorAll('td, th')).map(td => (td.innerText || '').trim() || '\u00A0');
                        if (cells.length > 3) {
                            const extra = cells.slice(3).join(' ').trim();
                            cells = [cells[0] || '', cells[1] || '', (cells[2] || '') + (extra ? ' ' + extra : '')];
                        }
                        const div = document.createElement('div');
                        div.className = 'script-row-wrapper';
                        cells.forEach(c => {
                            const col = document.createElement('div');
                            col.className = 'script-column';
                            const cell = document.createElement('div');
                            cell.className = 'cell-content';
                            cell.textContent = c;
                            col.appendChild(cell);
                            div.appendChild(col);
                        });
                        collected.push(div);
                    });
                });
                rowEls = collected;
            }
        }
        if (rowEls.length === 0) {
            const text = (temp.innerText || '').trim() || '\u00A0';
            sectionRows.push({ type: 'row', cells: [text], pill: getPillClassForRow(text, switchLabel, [text]), font12: text.includes('|v'), singleCol: true });
            sections.push({ header: name || 'Untitled', rows: sectionRows });
            continue;
        }

        rowEls.forEach(rowEl => {
            const cols = rowEl.querySelectorAll('.script-column');
            let cells = Array.from(cols).map(col => {
                const c = col.querySelector('.cell-content') || col.querySelector('.cell-locker') || col;
                let t = (c?.innerText ?? c?.textContent ?? '').toString().trim();
                if (!isDocx) t = t.replace(/^\|v(.*?)\|?$/, '$1').trim();
                return t || '\u00A0';
            });
            const bkmkDot = rowEl.querySelector('.bookmark-dot');
            const bkmkNum = (bkmkDot?.dataset?.bookmarkNum || '').toString();
            if (isDocx && cells.length > 3) {
                const extra = cells.slice(3).join(' ').trim();
                cells = [cells[0] || '', cells[1] || '', (cells[2] || '') + (extra ? ' ' + extra : '')];
            }
            const raw = (rowEl.textContent || '').trim();
            let pill = getPillClassForRow(raw, switchLabel, cells);
            for (let ci = 0; ci < cells.length; ci++) {
                if (pill === 'ms-pill-green' && ['switch', 'stay'].includes((cells[ci] || '').trim().toLowerCase())) {
                    cells[ci] = switchLabel;
                    break;
                }
            }
            const font12 = raw.includes('|v') || cells.some(c => (c || '').includes('|v'));
            const singleCol = cells.length <= 1;
            sectionRows.push({ type: 'row', cells, pill, font12, singleCol, bookmarkNum: bkmkNum });
        });
        sections.push({ header: name || 'Untitled', rows: sectionRows });
    }

    const tocHtml = '<div class="ms-toc-page"><h2>Master Script – File List</h2><div class="ms-toc-list">' +
        sections.map((s, i) => '<div class="ms-toc-item"><span class="ms-toc-num">' + (i + 1) + '</span>' + esc(s.header) + '</div>').join('') + '</div></div>';
    const sectionHtml = sections.map((sec, secIdx) => {
        const secNum = secIdx + 1;
        const maxCols = sec.rows.length ? Math.max(...sec.rows.map(r => r.cells.length)) : 1;
        let colgroup = '<colgroup>';
        if (maxCols <= 1) colgroup += '<col style="width:100%">';
        else {
            colgroup += '<col style="width:55px">';
            for (let i = 1; i < maxCols; i++) colgroup += '<col>';
        }
        colgroup += '</colgroup>';
        const rowHtml = sec.rows.map(r => {
            const pillCl = r.pill ? ' ' + r.pill : '';
            const fontCl = r.font12 ? ' ms-font-12' : '';
            const singleCl = r.singleCol ? ' ms-row-single-col' : '';
            const cells = r.cells.slice();
            while (cells.length < maxCols) cells.push('\u00A0');
            const cellsHtml = cells.map((c, ci) => {
                let inner = esc(c).replace(/\n/g, '<br>');
                if (ci === 0 && r.bookmarkNum) inner = '<span class="ms-bookmark-dot"></span> ' + inner;
                if (r.bookmarkNum && /^[\u2022\u00B7•·]/.test(c)) inner = inner.replace(/^([\u2022\u00B7•·])\s*/, '<span class="ms-bullet-red">$1</span> ');
                return '<td class="ms-cell">' + inner + '</td>';
            }).join('');
            return '<tr class="ms-row' + pillCl + fontCl + singleCl + '">' + cellsHtml + '</tr>';
        }).join('');
        return '<div class="ms-section"><div class="ms-file-header"><span class="ms-file-num">' + secNum + '</span>' + esc(sec.header) + '</div><table class="ms-table">' + colgroup + '<tbody>' + rowHtml + '</tbody></table></div>';
    }).join('');

    return `<!DOCTYPE html><html><head><meta charset="utf-8"><title>Master Script</title><style>
*{-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;}
.ms-print{font-family:Arial,sans-serif;color:#000;background:#fff;padding:16px;margin:0;}
.ms-toc-page{page-break-after:always;min-height:80vh;padding:24px 0;}
.ms-toc-page h2{font-size:16px;margin:0 0 12px 0;}
.ms-toc-list{font-size:14px;margin:0;}
.ms-toc-item{padding:12px 14px;border-bottom:1px solid #333;line-height:1.5;font-size:18px;font-weight:600;}
.ms-toc-item:last-child{border-bottom:1px solid #333;}
.ms-toc-num{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;border-radius:50%;background:#c62828;color:#fff;font-size:14px;font-weight:700;margin-right:10px;vertical-align:middle;}
.ms-section{margin-bottom:24px;}
.ms-file-header{font-size:18px;font-weight:bold;margin:0 0 8px 0;padding:8px 0;border-bottom:2px solid #333;}
.ms-file-num{display:inline-flex;align-items:center;justify-content:center;min-width:24px;height:24px;border-radius:50%;background:#c62828;color:#fff;font-size:14px;font-weight:700;margin-right:10px;vertical-align:middle;}
.ms-table{width:100%;max-width:100%;table-layout:fixed;border-collapse:separate;border-spacing:0;font-size:12px;overflow-wrap:break-word;}
.ms-table td{vertical-align:top;padding:6px 10px;border-bottom:1px solid #333;border-right:1px solid #ccc;overflow-wrap:break-word;word-wrap:break-word;word-break:break-all;min-width:0;max-width:100%;box-sizing:border-box;white-space:normal;}
.ms-table td:last-child{border-right:none;}
.ms-table td:first-child{width:55px;min-width:55px;white-space:nowrap;padding-right:12px;}
.ms-bookmark-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#c62828;margin-right:6px;vertical-align:middle;}
.ms-bullet-red{color:#c62828 !important;}
.ms-row.ms-row-single-col td{width:100% !important;max-width:100% !important;}
.ms-row.ms-font-12 td{font-size:12px;line-height:1.2;}
.ms-pill-red,.ms-pill-red td{background-color:#dc2626 !important;color:#000 !important;}
.ms-pill-yellow,.ms-pill-yellow td{background-color:#ffda03 !important;color:#000 !important;}
.ms-pill-green,.ms-pill-green td{background-color:#22c55e !important;color:#000 !important;}
.ms-pill-blue,.ms-pill-blue td{background-color:#2563eb !important;color:#fff !important;}
.ms-pill-white,.ms-pill-white td{background-color:#000 !important;color:#fff !important;padding:2px 8px !important;line-height:1.1 !important;font-size:11px !important;}
@page{size:letter;margin:0.5in;}
@media print{html{width:100%;height:100%;} body{width:100% !important;max-width:7.5in !important;margin:0 auto !important;padding:0 !important;background:#fff !important;-webkit-print-color-adjust:exact !important;print-color-adjust:exact !important;box-sizing:border-box !important;} *{box-sizing:border-box !important;} .ms-print{width:100% !important;max-width:7.5in !important;padding:16px !important;margin:0 auto !important;} .ms-section{width:100% !important;max-width:100% !important;} .ms-table{width:100% !important;max-width:100% !important;min-width:0 !important;table-layout:fixed !important;} .ms-table td{overflow-wrap:break-word !important;word-wrap:break-word !important;word-break:break-all !important;white-space:normal !important;} .ms-row.ms-row-single-col td{width:100% !important;max-width:100% !important;}}
</style></head><body class="ms-print">
<h1>Master Script</h1>
${tocHtml}
${sectionHtml}
</body></html>`;
}

if (btnEventPlanPrint && eventPlanOverlay) {
    btnEventPlanPrint.onclick = () => {
        if (!fileStore || fileStore.length === 0) {
            alert('No files in the Files list. Add files first.');
            return;
        }
        const titleEl = document.getElementById('event-plan-title');
        const callTimeEl = document.getElementById('event-plan-call-time');
        const runTimeEl = document.getElementById('event-plan-run-time');
        const title = titleEl ? titleEl.value.trim() : '';
        const callTime = callTimeEl ? callTimeEl.value.trim() : '';
        const runTime = runTimeEl ? runTimeEl.value.trim() : '';
        const html = buildEventPlanHtml(fileStore, title, callTime, runTime);
        const w = window.open('', '_blank');
        if (!w) {
            alert('Popup blocked. Please allow popups to print.');
            return;
        }
        w.document.write(html.replace('</body></html>', '<script>window.onload=function(){window.onafterprint=function(){window.close()};window.print()}<\/script></body></html>'));
        w.document.close();
        w.focus();
        closeEventPlanOverlay();
    };
}
const CONTROL_MODE_STORAGE_KEY = 'teleprompter_controlMode';
const INVERT_SCROLL_STORAGE_KEY = 'teleprompter_invertScroll';
const SCROLL_SENSITIVITY_STORAGE_KEY = 'teleprompter_scrollSensitivity';
const controlModeRadios = document.querySelectorAll('input[name="controlMode"]');
if (controlModeRadios && controlModeRadios.length) {
    try {
        const saved = localStorage.getItem(CONTROL_MODE_STORAGE_KEY);
        if (saved === 'off' || saved === 'mouse') {
            controlModeRadios.forEach((r) => { r.checked = r.value === saved; });
        }
    } catch (_) {}
    controlModeRadios.forEach((r) => {
        r.addEventListener('change', () => {
            try { localStorage.setItem(CONTROL_MODE_STORAGE_KEY, r.value); } catch (_) {}
        });
    });
}
if (invertScrollCheckbox) {
    try {
        const saved = localStorage.getItem(INVERT_SCROLL_STORAGE_KEY);
        if (saved === 'true') invertScrollCheckbox.checked = true;
        if (saved === 'false') invertScrollCheckbox.checked = false;
    } catch (_) {}
    isInvertScroll = invertScrollCheckbox.checked;
    invertScrollCheckbox.addEventListener('change', function() {
        isInvertScroll = invertScrollCheckbox.checked;
        try { localStorage.setItem(INVERT_SCROLL_STORAGE_KEY, invertScrollCheckbox.checked ? 'true' : 'false'); } catch (_) {}
        if (speedSlider) speedSlider.dispatchEvent(new Event('input'));
    });
}
const scrollSensitivityRange = document.getElementById('scroll-sensitivity');
const scrollSensitivityValueEl = document.getElementById('scroll-sensitivity-value');
if (scrollSensitivityRange) {
    try {
        const saved = localStorage.getItem(SCROLL_SENSITIVITY_STORAGE_KEY);
        if (saved != null) {
            const v = parseFloat(saved);
            if (!isNaN(v) && v >= 0.25 && v <= 2) {
                scrollSensitivityRange.value = String(v);
                scrollSensitivity = v;
            }
        }
    } catch (_) {}
    if (scrollSensitivityValueEl) scrollSensitivityValueEl.textContent = Math.round((parseFloat(scrollSensitivityRange.value) || 1) * 100) + '%';
    scrollSensitivity = parseFloat(scrollSensitivityRange.value) || 1;
    scrollSensitivityRange.addEventListener('input', function() {
        scrollSensitivity = parseFloat(scrollSensitivityRange.value) || 1;
        try { localStorage.setItem(SCROLL_SENSITIVITY_STORAGE_KEY, String(scrollSensitivity)); } catch (_) {}
        if (scrollSensitivityValueEl) scrollSensitivityValueEl.textContent = Math.round(scrollSensitivity * 100) + '%';
    });
}

if (resizer && runlistPanel) {
    let resizing = false;
    let startX = 0;
    let startWidth = 0;
    resizer.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        resizing = true;
        startX = e.clientX;
        startWidth = runlistPanel.offsetWidth;
        document.body.style.cursor = 'col-resize';
        document.body.style.userSelect = 'none';
    });
    document.addEventListener('mousemove', (e) => {
        if (!resizing) return;
        const deltaX = e.clientX - startX;
        let newWidth = startWidth - deltaX;
        newWidth = Math.max(200, Math.min(600, newWidth));
        runlistPanel.style.width = newWidth + 'px';
    });
    document.addEventListener('mouseup', () => {
        if (resizing) {
            resizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

// Files/Bookmarks vertical split resizer
const runlistBookmarkResizer = document.getElementById('runlist-bookmark-resizer');
const runlistSection = document.querySelector('.runlist-section');
const bookmarkSection = document.querySelector('.bookmark-section');
if (runlistBookmarkResizer && runlistSection && bookmarkSection && runlistPanel) {
    let bookmarkResizing = false;
    let startY = 0;
    let startFilesHeight = 0;
    runlistBookmarkResizer.addEventListener('mousedown', (e) => {
        if (e.button !== 0) return;
        bookmarkResizing = true;
        startY = e.clientY;
        startFilesHeight = runlistSection.offsetHeight;
        document.body.style.cursor = 'row-resize';
        document.body.style.userSelect = 'none';
    });
    document.addEventListener('mousemove', (e) => {
        if (!bookmarkResizing) return;
        const deltaY = e.clientY - startY;
        const panelHeight = runlistPanel.offsetHeight;
        const resizerHeight = 14; /* 10px height + 4px margin */
        const minFiles = 100;
        const minBookmarks = 80;
        let newFilesHeight = startFilesHeight + deltaY;
        newFilesHeight = Math.max(minFiles, Math.min(panelHeight - minBookmarks - resizerHeight, newFilesHeight));
        runlistSection.style.flex = `0 0 ${newFilesHeight}px`;
        bookmarkSection.style.flex = '1 1 0';
    });
    document.addEventListener('mouseup', () => {
        if (bookmarkResizing) {
            bookmarkResizing = false;
            document.body.style.cursor = '';
            document.body.style.userSelect = '';
        }
    });
}

function addFileToRunlist(file) {
    console.log(`Adding to runlist: ${file.name}`);
    const index = fileStore.length;
    fileStore.push(file);
    contentStore.push("");
    fileColumnCount.push(0);
    fileColumnVisibility.push([]);
    fileFirstColIsId.push(false);

    const row = document.createElement('div');
    row.className = 'runlist-row';
    row.dataset.index = index;
    row.innerHTML = `<span class="runlist-drag-handle" title="Drag to reorder" aria-label="Drag to reorder">⋮⋮</span><div class="runlist-row-left"><span class="file-name">${file.name}</span><div class="runlist-column-toggles"></div></div><button type="button" class="runlist-close" title="Close file" aria-label="Close file">×</button>`;

    row.onclick = (e) => {
        if (runlistJustDragged) { runlistJustDragged = false; return; }
        if (e.target.closest('.runlist-column-toggles') || e.target.closest('.runlist-close') || e.target.closest('.runlist-drag-handle')) return;
        console.log(`Row clicked: loading index ${index}`);
        loadScriptToEditor(index);
    };
    const closeBtn = row.querySelector('.runlist-close');
    if (closeBtn) {
        closeBtn.onclick = (e) => {
            e.preventDefault();
            e.stopPropagation();
            removeFileFromRunlist(index);
        };
    }
    attachRunlistRowDrag(row, index);

    runlistContainer.appendChild(row);
    processFileContent(file, index);
    resetPillTriggerState();
    updateFilenamePills();
}

function attachRunlistRowDrag(row, index) {
    if (row._runlistDragAttached) return;
    row._runlistDragAttached = true;
    const handle = row.querySelector('.runlist-drag-handle');
    if (!handle) return;
    handle.addEventListener('mousedown', (e) => {
        e.preventDefault();
        e.stopPropagation();
        runlistDraggingIndex = parseInt(row.dataset.index, 10);
        runlistDraggingRow = row;
        row.classList.add('runlist-dragging');
        document.addEventListener('mousemove', onRunlistMouseMove, true);
        document.addEventListener('mouseup', onRunlistMouseUp, true);
    });
}

function moveRunlistRow(fromIndex, toIndex) {
    if (fromIndex < 0 || fromIndex >= fileStore.length || toIndex < 0 || toIndex > fileStore.length || fromIndex === toIndex) return;
    const move = (arr) => {
        const item = arr.splice(fromIndex, 1)[0];
        arr.splice(toIndex, 0, item);
    };
    move(fileStore);
    move(contentStore);
    move(fileColumnCount);
    move(fileColumnVisibility);
    move(fileFirstColIsId);
    const rows = Array.from(runlistContainer.querySelectorAll('.runlist-row'));
    const draggedRow = rows[fromIndex];
    if (draggedRow) {
        if (toIndex < fromIndex) runlistContainer.insertBefore(draggedRow, rows[toIndex]);
        else if (toIndex >= rows.length) runlistContainer.appendChild(draggedRow);
        else runlistContainer.insertBefore(draggedRow, rows[toIndex].nextSibling);
    }
    const reindexed = Array.from(runlistContainer.querySelectorAll('.runlist-row'));
    reindexed.forEach((r, i) => {
        r.dataset.index = String(i);
        const nameEl = r.querySelector('.file-name');
        if (nameEl) nameEl.textContent = fileStore[i].name;
        r.onclick = (e) => {
            if (runlistJustDragged) { runlistJustDragged = false; return; }
            if (e.target.closest('.runlist-column-toggles') || e.target.closest('.runlist-close') || e.target.closest('.runlist-drag-handle')) return;
            loadScriptToEditor(i);
        };
        const closeBtn = r.querySelector('.runlist-close');
        if (closeBtn) closeBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); removeFileFromRunlist(i); };
        updateRunlistRowColumnToggles(i);
    });
    if (currentFileIndex === fromIndex) currentFileIndex = toIndex;
    else if (fromIndex < currentFileIndex && toIndex >= currentFileIndex) currentFileIndex--;
    else if (fromIndex > currentFileIndex && toIndex <= currentFileIndex) currentFileIndex++;
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const activeRow = runlistContainer.querySelector(`.runlist-row[data-index="${currentFileIndex}"]`);
    if (activeRow) activeRow.classList.add('active');
    resetPillTriggerState();
    updateFilenamePills();
}

function removeFileFromRunlist(index) {
    if (index < 0 || index >= fileStore.length) return;
    const wasActive = currentFileIndex === index;
    fileStore.splice(index, 1);
    contentStore.splice(index, 1);
    fileColumnCount.splice(index, 1);
    fileColumnVisibility.splice(index, 1);
    fileFirstColIsId.splice(index, 1);
    const row = runlistContainer.querySelector(`.runlist-row[data-index="${index}"]`);
    if (row) row.remove();
    const rows = Array.from(runlistContainer.querySelectorAll('.runlist-row'));
    rows.forEach((r, i) => {
        r.dataset.index = String(i);
        const nameEl = r.querySelector('.file-name');
        if (nameEl) nameEl.textContent = fileStore[i].name;
        r.onclick = (e) => {
            if (runlistJustDragged) { runlistJustDragged = false; return; }
            if (e.target.closest('.runlist-column-toggles') || e.target.closest('.runlist-close') || e.target.closest('.runlist-drag-handle')) return;
            loadScriptToEditor(i);
        };
        const closeBtn = r.querySelector('.runlist-close');
        if (closeBtn) closeBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); removeFileFromRunlist(i); };
        attachRunlistRowDrag(r, i);
        updateRunlistRowColumnToggles(i);
    });
    if (fileStore.length === 0) {
        currentFileIndex = -1;
        teleprompterText.innerHTML = '<br>';
        delete teleprompterText.dataset.placeholder;
        document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
        resetPillTriggerState();
        updateFilenamePills();
        return;
    }
    if (wasActive) {
        const newIndex = index >= fileStore.length ? fileStore.length - 1 : index;
        currentFileIndex = newIndex;
        loadScriptToEditor(newIndex);
    } else if (currentFileIndex > index) {
        currentFileIndex--;
    }
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const activeRow = runlistContainer.querySelector(`.runlist-row[data-index="${currentFileIndex}"]`);
    if (activeRow) activeRow.classList.add('active');
    resetPillTriggerState();
    updateFilenamePills();
}

/** If existing files have numeric prefix (e.g. "02_"), sort full runlist by filename. */
function sortRunlistIfNumericPrefix(countBefore) {
    const firstSet = fileStore.slice(0, countBefore);
    const allStartWithNumber = firstSet.length > 0 && firstSet.every(f => /^\d+_/.test(f.name));
    if (!allStartWithNumber) return;
    /* Save current file's content before reorder so we don't overwrite with wrong index */
    if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
        const html = teleprompterText.innerHTML.trim();
        contentStore[currentFileIndex] = (html === '<br>' || html === '') ? '' : html;
    }
    const currentName = currentFileIndex >= 0 && fileStore[currentFileIndex] ? fileStore[currentFileIndex].name : null;
    const indices = fileStore.map((_, i) => i);
    indices.sort((a, b) => (fileStore[a].name || '').localeCompare(fileStore[b].name || '', undefined, { numeric: true }));
    const newFileStore = indices.map(i => fileStore[i]);
    const newContentStore = indices.map(i => contentStore[i]);
    const newFileColumnCount = indices.map(i => fileColumnCount[i]);
    const newFileColumnVisibility = indices.map(i => fileColumnVisibility[i]);
    const newFileFirstColIsId = indices.map(i => fileFirstColIsId[i]);
    fileStore.length = 0;
    fileStore.push(...newFileStore);
    contentStore.length = 0;
    contentStore.push(...newContentStore);
    fileColumnCount.length = 0;
    fileColumnCount.push(...newFileColumnCount);
    fileColumnVisibility.length = 0;
    fileColumnVisibility.push(...newFileColumnVisibility);
    fileFirstColIsId.length = 0;
    fileFirstColIsId.push(...newFileFirstColIsId);
    const newCurrentIndex = currentName ? newFileStore.findIndex(f => f.name === currentName) : -1;
    currentFileIndex = newCurrentIndex >= 0 ? newCurrentIndex : 0;
    runlistContainer.innerHTML = '';
    fileStore.forEach((file, i) => {
        const row = document.createElement('div');
        row.className = 'runlist-row';
        row.dataset.index = String(i);
        row.innerHTML = `<span class="runlist-drag-handle" title="Drag to reorder" aria-label="Drag to reorder">⋮⋮</span><div class="runlist-row-left"><span class="file-name">${file.name}</span><div class="runlist-column-toggles"></div></div><button type="button" class="runlist-close" title="Close file" aria-label="Close file">×</button>`;
        row.onclick = (e) => {
            if (runlistJustDragged) { runlistJustDragged = false; return; }
            if (e.target.closest('.runlist-column-toggles') || e.target.closest('.runlist-close') || e.target.closest('.runlist-drag-handle')) return;
            loadScriptToEditor(i);
        };
        const closeBtn = row.querySelector('.runlist-close');
        if (closeBtn) closeBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); removeFileFromRunlist(i); };
        attachRunlistRowDrag(row, i);
        runlistContainer.appendChild(row);
        updateRunlistRowColumnToggles(i);
    });
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    /* If target file has no content yet (async load pending), load first file that does to avoid empty editor */
    let idxToLoad = currentFileIndex;
    const targetContent = contentStore[idxToLoad];
    const hasContent = (c) => c && typeof c === 'string' && c.trim() && c.trim() !== '<br>';
    if (!hasContent(targetContent)) {
        const firstWithContent = contentStore.findIndex(hasContent);
        if (firstWithContent >= 0) {
            idxToLoad = firstWithContent;
            currentFileIndex = firstWithContent;
        }
    }
    const activeRow = runlistContainer.querySelector(`.runlist-row[data-index="${currentFileIndex}"]`);
    if (activeRow) activeRow.classList.add('active');
    if (hasContent(contentStore[idxToLoad])) {
        loadScriptToEditor(idxToLoad);
    } else {
        updateFilenamePills();
    }
}

/** Reset pill trigger state so top/bottom filename pills can fire again when file list changes. */
function resetPillTriggerState() {
    topPillTriggerFired = false;
    bottomPillTriggerFired = false;
    const view = document.getElementById('teleprompter-view');
    if (view) lastScrollTopForPillTrigger = view.scrollTop;
}

/** Normalize HTML to max 3 columns. Merges extra cells into the 3rd. Returns { html, wasTrimmed }. */
function normalizeContentToMax3Columns(html) {
    if (!html || typeof html !== 'string' || !html.trim()) return { html: html || '', wasTrimmed: false };
    const temp = document.createElement('div');
    temp.innerHTML = html;
    let wasTrimmed = false;
    temp.querySelectorAll('.script-row-wrapper').forEach(row => {
        const cols = row.querySelectorAll('.script-column');
        if (cols.length > MAX_COLUMNS) {
            wasTrimmed = true;
            const extras = Array.from(cols).slice(MAX_COLUMNS);
            const extraText = extras.map(c => (c.innerText || c.textContent || '').trim()).join(' ').trim();
            extras.forEach(c => c.remove());
            const third = row.querySelector('.script-column:nth-child(3)');
            if (third) {
                const cell = third.querySelector('.cell-content') || third.querySelector('.cell-locker') || third;
                if (cell && extraText) cell.textContent = (cell.textContent || '').trim() + (cell.textContent?.trim() ? ' ' : '') + extraText;
            }
        }
    });
    temp.querySelectorAll('table tr').forEach(tr => {
        const cells = Array.from(tr.querySelectorAll('td, th'));
        if (cells.length > MAX_COLUMNS) {
            wasTrimmed = true;
            const third = cells[MAX_COLUMNS - 1];
            const extras = cells.slice(MAX_COLUMNS);
            const extraText = extras.map(c => (c.innerText || c.textContent || '').trim()).join(' ').trim();
            extras.forEach(c => c.remove());
            if (third && extraText) third.textContent = (third.textContent || '').trim() + (third.textContent?.trim() ? ' ' : '') + extraText;
        }
    });
    return { html: temp.innerHTML, wasTrimmed };
}

async function processFileContent(file, index) {
    const extension = file.name.split('.').pop().toLowerCase();
    console.log(`Starting process for .${extension} at index ${index}`);

    try {
        if (extension === 'txt' || extension === 'html') {
            let text = await file.text();
            const currentIdx = fileStore.findIndex(f => f === file);
            const slot = currentIdx >= 0 ? currentIdx : index;
            if (extension === 'html') {
                const { html: normalized, wasTrimmed } = normalizeContentToMax3Columns(text);
                text = normalized;
                if (wasTrimmed) alert(`"${file.name}" had more than ${MAX_COLUMNS} columns. Extra columns were merged into column ${MAX_COLUMNS}.`);
            }
            contentStore[slot] = text;
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
                    const sheet = workbook.Sheets[workbook.SheetNames[0]];
                    const range = sheet['!ref'] ? XLSX.utils.decode_range(sheet['!ref']) : { s: { c: 0, r: 0 }, e: { c: 0, r: 0 } };
                    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
                    const allCols = [];
                    for (let c = range.s.c; c <= range.e.c; c++) allCols.push(c);
                    const maxCols = MAX_COLUMNS;
                    const selectedCols = allCols.slice(0, maxCols);
                    let html = '<div class="script-container">';
                    json.forEach((row) => {
                        const cells = selectedCols.map(colIdx => row[colIdx] != null ? String(row[colIdx]).trim() : '');
                        if (allCols.length > maxCols) {
                            const extra = allCols.slice(maxCols).map(colIdx => row[colIdx] != null ? String(row[colIdx]).trim() : '').join(' ').trim();
                            if (extra && cells.length >= maxCols) cells[maxCols - 1] = (cells[maxCols - 1] || '') + (cells[maxCols - 1] ? ' ' : '') + extra;
                        }
                        html += '<div class="script-row-wrapper">';
                        cells.forEach(val => {
                            html += `<div class="script-column"><div class="cell-content">${(val || '').replace(/</g, '&lt;').replace(/>/g, '&gt;')}</div></div>`;
                        });
                        html += '</div>';
                    });
                    html += '</div>';
                    const currentIdx = fileStore.findIndex(f => f === file);
                    const slot = currentIdx >= 0 ? currentIdx : index;
                    contentStore[slot] = html;
                    fileColumnCount[slot] = Math.min(allCols.length, maxCols);
                    fileColumnVisibility[slot] = selectedCols.map(() => true);
                    fileFirstColIsId[slot] = selectedCols.length > 0 && isFirstColumnNumeric(sheet);
                    updateRunlistRowColumnToggles(slot);
                    loadScriptToEditor(slot);
                    if (allCols.length > maxCols) {
                        alert(`"${file.name}" has ${allCols.length} columns. Teleprompter uses a maximum of ${MAX_COLUMNS} columns. Extra columns were merged into column ${MAX_COLUMNS}.`);
                    }
                    console.log("Excel imported with " + selectedCols.length + " column(s). Use file checkboxes to show/hide columns.");
                } catch (innerErr) {
                    console.error("❌ XLSX Parsing Error:", innerErr);
                }
            };
            reader.onerror = (err) => console.error("❌ FileReader Error:", err);
            reader.readAsArrayBuffer(file);
        } else if (extension === 'docx' || extension === 'doc') {
            console.log("Word document detected. Converting with Mammoth...");
            const reader = new FileReader();
            reader.onload = (e) => {
                    mammoth.convertToHtml({ arrayBuffer: e.target.result })
                    .then(result => {
                        const currentIdx = fileStore.findIndex(f => f === file);
                        const slot = currentIdx >= 0 ? currentIdx : index;
                        const { html, wasTrimmed } = normalizeContentToMax3Columns(result.value);
                        contentStore[slot] = html;
                        if (wasTrimmed) alert(`"${file.name}" had more than ${MAX_COLUMNS} columns. Extra columns were merged into column ${MAX_COLUMNS}.`);
                        console.log("Docx converted successfully");
                        loadScriptToEditor(slot);
                    })
                    .catch(err => console.error("❌ Mammoth conversion error:", err));
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

const UUID_REGEX = /[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g;
const FIRST_COL_JUNK_REGEX = /[\s\u00A0\-_,;:]+|[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}/g;

/** Returns true if the first column of the sheet is mostly numbers (after stripping junk). */
function isFirstColumnNumeric(sheet) {
    if (!sheet || !sheet['!ref']) return false;
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    let numericCount = 0;
    let total = 0;
    for (let r = 0; r < json.length; r++) {
        const cell = json[r][0];
        if (cell == null && cell !== 0) continue;
        const raw = String(cell).trim();
        if (!raw) continue;
        total++;
        const cleaned = raw.replace(FIRST_COL_JUNK_REGEX, '').trim();
        const looksNumeric = /^\d+$/.test(cleaned) || (cleaned !== '' && !isNaN(Number(cleaned)));
        if (looksNumeric) numericCount++;
    }
    return total > 0 && numericCount / total >= 0.5;
}

/** Strip |v prefix and trailing | from first column (e.g. |v13 or |v13| -> 13). Only for xlsx with numeric first column. */
function stripPipeVFromFirstColumn() {
    if (!teleprompterText) return;
    if (!isCurrentFileXlsx()) return;
    if (!(currentFileIndex >= 0 && fileFirstColIsId[currentFileIndex])) return;
    teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
        const firstCol = row.querySelector('.script-column:first-child .cell-content') || row.querySelector('.script-column:first-child .cell-locker') || row.querySelector('.script-column:first-child');
        if (!firstCol) return;
        const raw = (firstCol.textContent || '').trim();
        const cleaned = raw.replace(/^\|v(.*?)\|?$/, '$1').trim();
        if (cleaned !== raw) firstCol.textContent = cleaned;
    });
}

function removeUuidFromFirstColumn() {
    teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
        const firstCol = row.querySelector('.script-column:first-child .cell-content') || row.querySelector('.script-column:first-child');
        if (!firstCol) return;
        const walk = (node) => {
            if (node.nodeType === Node.TEXT_NODE) {
                const cleaned = node.textContent.replace(UUID_REGEX, '');
                if (cleaned !== node.textContent) {
                    node.textContent = cleaned;
                }
            } else if (node.childNodes) {
                for (let i = 0; i < node.childNodes.length; i++) walk(node.childNodes[i]);
            }
        };
        walk(firstCol);
    });
}

function processTableColumns() {
    const BLOCK_TAGS = ['P', 'DIV', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6'];
    /* Convert tables to script-container/script-row-wrapper/script-column structure */
    const tables = teleprompterText.querySelectorAll('table');
    tables.forEach(table => {
        const container = document.createElement('div');
        container.className = 'script-container';
        const rows = table.querySelectorAll('tr');
        rows.forEach(tr => {
            const rowDiv = document.createElement('div');
            rowDiv.className = 'script-row-wrapper';
            tr.querySelectorAll('td, th').forEach(cell => {
                const colDiv = document.createElement('div');
                colDiv.className = 'script-column';
                const contentDiv = document.createElement('div');
                contentDiv.className = 'cell-content';
                contentDiv.innerHTML = cell.innerHTML;
                colDiv.appendChild(contentDiv);
                rowDiv.appendChild(colDiv);
            });
            container.appendChild(rowDiv);
        });
        table.parentNode?.replaceChild(container, table);
    });
    /* Convert standalone blocks (paragraphs, headings) to single-column rows for docx mixed content */
    const directChildren = Array.from(teleprompterText.children);
    for (const child of directChildren) {
        if (child.nodeType !== Node.ELEMENT_NODE) continue;
        const tag = child.tagName?.toUpperCase();
        if (!BLOCK_TAGS.includes(tag)) continue;
        if (child.classList?.contains('script-container')) continue;
        const rowDiv = document.createElement('div');
        rowDiv.className = 'script-row-wrapper';
        const colDiv = document.createElement('div');
        colDiv.className = 'script-column';
        const contentDiv = document.createElement('div');
        contentDiv.className = 'cell-content';
        contentDiv.innerHTML = child.innerHTML;
        colDiv.appendChild(contentDiv);
        rowDiv.appendChild(colDiv);
        const wrapper = document.createElement('div');
        wrapper.className = 'script-container';
        wrapper.appendChild(rowDiv);
        child.parentNode?.replaceChild(wrapper, child);
    }
    /* Merge all script-containers into one so every row shares the same table layout */
    ensureSingleScriptContainer();
    /* Split any cell content that contains hard returns (<br> or newline) into separate rows */
    splitMultiLineCellsIntoRows();
    /* Ensure blank/empty rows have nbsp so they keep height and stay separated */
    ensureEmptyRowsHaveNbsp();
}

/** Set empty cell-content to nbsp so blank rows preserve height and separation. */
function ensureEmptyRowsHaveNbsp() {
    teleprompterText.querySelectorAll('.cell-content').forEach(cell => {
        const text = (cell.textContent || '').trim();
        if (!text) cell.innerHTML = '\u00A0';
    });
}

/** Split cell innerHTML by <br> and newline into an array of HTML fragments (one per line). */
function splitHtmlByLineBreaks(html) {
    if (!html || typeof html !== 'string') return [html || ''];
    const SENTINEL = '\u0000';
    const normalized = html.replace(/\r\n|\r|\n/g, '<br>').replace(/<br\s*\/?>/gi, SENTINEL);
    return normalized.split(SENTINEL).map(s => s.trim());
}

/** Expand rows whose cell-content contains <br> or newlines into multiple rows, one per line. */
function splitMultiLineCellsIntoRows() {
    const container = teleprompterText.querySelector('.script-container');
    if (!container) return;
    const rows = Array.from(container.querySelectorAll(':scope > .script-row-wrapper'));
    for (const row of rows) {
        const cols = row.querySelectorAll('.script-column');
        const partsPerCol = Array.from(cols).map(col => {
            const cell = col.querySelector('.cell-content') || col;
            const html = cell.innerHTML || '';
            return splitHtmlByLineBreaks(html);
        });
        const maxParts = Math.max(1, ...partsPerCol.map(p => p.length));
        if (maxParts <= 1) continue;
        const numCols = cols.length;
        let insertBefore = row;
        for (let i = 0; i < maxParts; i++) {
            const newRow = document.createElement('div');
            newRow.className = row.className;
            for (let j = 0; j < numCols; j++) {
                const origCol = cols[j];
                const origCell = origCol.querySelector('.cell-content') || origCol;
                const colDiv = document.createElement('div');
                colDiv.className = origCol.className;
                const contentDiv = document.createElement('div');
                contentDiv.className = 'cell-content';
                if (origCell.style && origCell.style.cssText) contentDiv.style.cssText = origCell.style.cssText;
                const parts = partsPerCol[j];
                const raw = (parts[i] !== undefined ? parts[i] : '').trim();
                contentDiv.innerHTML = raw || '\u00A0'; /* empty lines: use nbsp to preserve row height */
                colDiv.appendChild(contentDiv);
                newRow.appendChild(colDiv);
            }
            row.parentNode.insertBefore(newRow, insertBefore);
            insertBefore = newRow.nextSibling;
        }
        row.remove();
    }
}

/** Merge all direct-child script-containers into a single one. Prevents pill/row layout differences when docx/xlsx produce multiple containers. */
function ensureSingleScriptContainer() {
    const containers = Array.from(teleprompterText.children).filter(el =>
        el.nodeType === Node.ELEMENT_NODE && el.classList?.contains('script-container')
    );
    if (containers.length <= 1) return;
    const first = containers[0];
    for (let i = 1; i < containers.length; i++) {
        const c = containers[i];
        while (c.firstChild) first.appendChild(c.firstChild);
        c.remove();
    }
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
                        delete teleprompterText.dataset.placeholder;
                        processTableColumns();
                        wrapCellContentInBlock();
                        requestAnimationFrame(() => {
                            convertBkmkPlaceholdersToBookmarks();
                            updateBookmarkSidebar();
                            if (currentFileIndex === index && index < contentStore.length)
                                contentStore[index] = teleprompterText.innerHTML.trim();
                        });
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
                delete teleprompterText.dataset.placeholder;
                processTableColumns();
                wrapCellContentInBlock();
                requestAnimationFrame(() => {
                    convertBkmkPlaceholdersToBookmarks();
                    updateBookmarkSidebar();
                    if (currentFileIndex === index && index < contentStore.length)
                        contentStore[index] = teleprompterText.innerHTML.trim();
                });
            }
        } else {
            const text = new TextDecoder().decode(e.target.result);
            contentStore[index] = text;
            if (currentFileIndex === index) {
                teleprompterText.innerHTML = text;
                delete teleprompterText.dataset.placeholder;
                requestAnimationFrame(() => {
                    convertBkmkPlaceholdersToBookmarks();
                    updateBookmarkSidebar();
                    if (currentFileIndex === index && index < contentStore.length)
                        contentStore[index] = teleprompterText.innerHTML.trim();
                });
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
function removeBookmark(bookmarkId) {
    if (!bookmarkId) return;
    const numCircle = teleprompterText?.querySelector(`.bookmark-dot[data-bookmark-id="${bookmarkId}"]`);
    const cursorDot = teleprompterText?.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${bookmarkId}"]`);
    if (numCircle) numCircle.remove();
    if (cursorDot) {
        const text = (cursorDot.innerText || cursorDot.textContent || '').replace(/^\s*•\s*/, '').trim();
        const textNode = document.createTextNode(text ? ' ' + text : ' ');
        cursorDot.parentNode?.replaceChild(textNode, cursorDot);
    }
    renumberBookmarks();
    updateBookmarkSidebar();
    if (typeof syncEditorState === 'function') syncEditorState();
    if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
}

function updateBookmarkSidebar() {
    const list = document.querySelector('.bookmark-list');
    if (!list) return;
    const bars = getSortedBookmarkDots();
    list.innerHTML = '';
    const esc = (s) => (s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
    bars.forEach((el, i) => {
        const num = el.dataset.bookmarkNum || String(i + 1);
        const label = el.dataset.bookmarkLabel || '';
        const bid = el.dataset.bookmarkId || '';
        const item = document.createElement('div');
        item.className = 'bookmark-item';
        item.dataset.bookmarkIndex = String(i);
        item.dataset.bookmarkId = bid;
        item.innerHTML = `<span class="bookmark-num">${num}</span><span class="bookmark-label">${esc(label)}</span><button type="button" class="bookmark-close" title="Delete bookmark" aria-label="Delete bookmark">×</button>`;
        const closeBtn = item.querySelector('.bookmark-close');
        if (closeBtn) closeBtn.onclick = (e) => { e.preventDefault(); e.stopPropagation(); removeBookmark(bid); };
        item.onclick = (e) => {
            if (e.target.closest('.bookmark-close')) return;
            if (pendingBookmarkNavTimeoutId) {
                clearTimeout(pendingBookmarkNavTimeoutId);
                pendingBookmarkNavTimeoutId = null;
            }
            bookmarkNavigationInProgress = true;
            pendingBookmarkTargetIndex = i;
            setActiveBookmarkByIndex(i);
            el.scrollIntoView({ behavior: 'smooth', block: 'center' });
            pendingBookmarkNavTimeoutId = setTimeout(() => {
                bookmarkNavigationInProgress = false;
                pendingBookmarkTargetIndex = -1;
                pendingBookmarkNavTimeoutId = null;
            }, BOOKMARK_NAV_ARRIVAL_TIMEOUT_MS);
        };
        list.appendChild(item);
    });
    /* Sync highlight with current scroll position */
    updateBookmarkHighlightFromScroll();
}

function setActiveBookmarkByIndex(idx) {
    const list = document.querySelector('.bookmark-list');
    if (!list) return;
    const items = list.querySelectorAll('.bookmark-item');
    items.forEach((item, i) => {
        item.classList.toggle('active', idx >= 0 && i === idx);
    });
}

function updateBookmarkHighlightFromScroll() {
    const view = document.getElementById('teleprompter-view');
    checkTopPillAndGoToPreviousFile();
    checkBottomPillAndAdvanceToNextFile();
    const bars = getSortedBookmarkDots();
    if (!view || bars.length === 0) return;
    if (bookmarkNavigationInProgress && pendingBookmarkTargetIndex >= 0) {
        const barsArr = bars;
        if (pendingBookmarkTargetIndex < barsArr.length) {
            const wrapper = document.getElementById('indicator-wrapper');
            const indicatorY = wrapper ? (wrapper.getBoundingClientRect().top + wrapper.getBoundingClientRect().height / 2) : (view.getBoundingClientRect().top + view.getBoundingClientRect().height / 2);
            const targetCursor = teleprompterText?.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${barsArr[pendingBookmarkTargetIndex].dataset.bookmarkId}"]`);
            if (targetCursor) {
                const rect = targetCursor.getBoundingClientRect();
                const centerY = rect.top + rect.height / 2;
                if (centerY <= indicatorY) {
                    bookmarkNavigationInProgress = false;
                    pendingBookmarkTargetIndex = -1;
                    if (pendingBookmarkNavTimeoutId) {
                        clearTimeout(pendingBookmarkNavTimeoutId);
                        pendingBookmarkNavTimeoutId = null;
                    }
                }
            }
        }
        return;
    }
    if (bookmarkNavigationInProgress) return;
    const idx = getCurrentBookmarkIndexFromCursorDot();
    if (idx >= 0 && idx < bars.length) setActiveBookmarkByIndex(idx);
}

function checkTopPillAndGoToPreviousFile() {
    const topPill = document.getElementById('filename-pill-top');
    const wrapper = document.getElementById('indicator-wrapper');
    const view = document.getElementById('teleprompter-view');
    if (!topPill || !wrapper || !view) return;
    if (topPill.classList.contains('hidden')) {
        topPillTriggerFired = false;
        return;
    }
    const hasPrev = typeof currentFileIndex !== 'undefined' && currentFileIndex > 0 && typeof fileStore !== 'undefined';
    if (!hasPrev) return;
    const currentScrollTop = view.scrollTop;
    const scrollDelta = lastScrollTopForPillTrigger != null ? currentScrollTop - lastScrollTopForPillTrigger : 0;
    const scrollingUp = (typeof scrollSpeed !== 'undefined' && scrollSpeed < 0) || (scrollDelta < 0);
    if (!scrollingUp) return;
    const indicatorY = wrapper.getBoundingClientRect().top + wrapper.getBoundingClientRect().height / 2;
    const rect = topPill.getBoundingClientRect();
    const indicatorTouchingPill = indicatorY >= rect.top && indicatorY <= rect.bottom;
    if (!indicatorTouchingPill) {
        topPillTriggerFired = false;
        return;
    }
    if (topPillTriggerFired) return;
    topPillTriggerFired = true;
    if (typeof playPageTurnAndGoToPreviousFile === 'function') playPageTurnAndGoToPreviousFile(currentFileIndex - 1);
}

function checkBottomPillAndAdvanceToNextFile() {
    const bottomPill = document.getElementById('filename-pill-bottom');
    const wrapper = document.getElementById('indicator-wrapper');
    const view = document.getElementById('teleprompter-view');
    if (!bottomPill || !wrapper || !view) return;
    if (bottomPill.classList.contains('hidden')) {
        bottomPillTriggerFired = false;
        return;
    }
    const hasNext = typeof currentFileIndex !== 'undefined' && currentFileIndex >= 0 && typeof fileStore !== 'undefined' && currentFileIndex + 1 < fileStore.length;
    if (!hasNext) return;
    const currentScrollTop = view.scrollTop;
    const scrollDelta = lastScrollTopForPillTrigger != null ? currentScrollTop - lastScrollTopForPillTrigger : 0;
    lastScrollTopForPillTrigger = currentScrollTop;
    const scrollingDown = (typeof scrollSpeed !== 'undefined' && scrollSpeed > 0) || (scrollDelta > 0);
    if (!scrollingDown) return;
    const indicatorY = wrapper.getBoundingClientRect().top + wrapper.getBoundingClientRect().height / 2;
    const rect = bottomPill.getBoundingClientRect();
    const indicatorTouchingPill = indicatorY >= rect.top && indicatorY <= rect.bottom;
    if (!indicatorTouchingPill) {
        bottomPillTriggerFired = false;
        return;
    }
    if (bottomPillTriggerFired) return;
    bottomPillTriggerFired = true;
    if (typeof playPageTurnAndAdvanceToNextFile === 'function') playPageTurnAndAdvanceToNextFile(currentFileIndex + 1);
}

const BOOKMARK_CLICK_MATCH_THRESHOLD_PX = 80;
const BOOKMARK_NAV_ARRIVAL_TIMEOUT_MS = 3000;
let bookmarkNavigationInProgress = false;
let pendingBookmarkTargetIndex = -1;
let pendingBookmarkNavTimeoutId = null;
let bottomPillTriggerFired = false;
let topPillTriggerFired = false;
let lastScrollTopForPillTrigger = null;

function getBookmarkIndexAtY(clientY) {
    const cursorDots = teleprompterText ? Array.from(teleprompterText.querySelectorAll('.bookmark-cursor-dot')) : [];
    if (cursorDots.length === 0) return -1;
    const bars = getSortedBookmarkDots();
    let bestIdx = -1;
    let bestDist = Infinity;
    cursorDots.forEach((cursorDot) => {
        const bid = cursorDot.dataset.bookmarkId;
        if (!bid) return;
        const rect = cursorDot.getBoundingClientRect();
        const centerY = rect.top + rect.height / 2;
        const dist = Math.abs(centerY - clientY);
        if (dist < bestDist && dist <= BOOKMARK_CLICK_MATCH_THRESHOLD_PX) {
            const dot = bars.find((b) => b.dataset.bookmarkId === bid);
            if (dot) {
                const idx = bars.indexOf(dot);
                if (idx >= 0) {
                    bestDist = dist;
                    bestIdx = idx;
                }
            }
        }
    });
    return bestIdx;
}

function getCurrentBookmarkIndexFromCursorDot() {
    const view = document.getElementById('teleprompter-view');
    const bars = getSortedBookmarkDots();
    if (!view || bars.length === 0) return -1;
    const wrapper = document.getElementById('indicator-wrapper');
    const viewRect = view.getBoundingClientRect();
    const indicatorY = wrapper ? (wrapper.getBoundingClientRect().top + wrapper.getBoundingClientRect().height / 2) : (viewRect.top + viewRect.height / 2);
    let lastPassedIdx = -1;
    for (let i = 0; i < bars.length; i++) {
        const cursorDot = teleprompterText?.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${bars[i].dataset.bookmarkId}"]`);
        if (!cursorDot) continue;
        const rect = cursorDot.getBoundingClientRect();
        const centerY = rect.top + rect.height / 2;
        if (centerY <= indicatorY) {
            lastPassedIdx = i;
        }
    }
    if (lastPassedIdx >= 0) return lastPassedIdx;
    const firstCursor = teleprompterText?.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${bars[0].dataset.bookmarkId}"]`);
    if (firstCursor) {
        const firstRect = firstCursor.getBoundingClientRect();
        const firstCenterY = firstRect.top + firstRect.height / 2;
        if (firstCenterY > indicatorY) {
            return 0;
        }
    }
    return -1;
}

function addBookmarkAtCursor() {
    pushUndoState();
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    let node = range.startContainer;
    if (node.nodeType === Node.TEXT_NODE) node = node.parentElement;
    if (!node || !teleprompterText.contains(node)) return;

    /* Remove preceding <br> or empty blocks to avoid double line break when bookmark is added after Enter */
    const atStart = range.startOffset === 0;
    const cursorNode = range.startContainer;
    let prev = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.previousSibling : cursorNode.previousSibling;
    let el = cursorNode.nodeType === Node.TEXT_NODE ? cursorNode.parentElement : cursorNode;
    while (el && !prev && el !== teleprompterText) {
        prev = el.previousSibling;
        el = el.parentElement;
    }
    if (atStart && prev) {
        if (prev.nodeName === 'BR') {
            prev.remove();
            range.insertNode(document.createTextNode(' '));
        } else if (prev.nodeType === Node.ELEMENT_NODE && (!prev.textContent?.trim() || (prev.childNodes.length === 1 && prev.firstChild?.nodeName === 'BR'))) {
            prev.remove();
            range.insertNode(document.createTextNode(' '));
        }
    }

    let row = node.closest('.script-row-wrapper');
    let cursorCol = node.closest('.script-column');
    let col1 = null;
    if (row) {
        const cols = row.querySelectorAll('.script-column');
        col1 = cols[0];
    }
    if (!row) {
        row = node.closest('div, p') || teleprompterText;
        col1 = row;
    }
    if (!row) return;

    const existingDots = teleprompterText.querySelectorAll('.bookmark-dot');
    const nextNum = existingDots.length + 1;
    const bookmarkId = `bookmark-${nextNum}`;
    const stableId = `bk-${Date.now()}-${Math.random().toString(36).slice(2)}`;

    const label = (function () {
        const cell = cursorCol || row;
        if (!cell) return '';
        const endRange = document.createRange();
        endRange.selectNodeContents(cell);
        endRange.setStart(range.startContainer, range.startOffset);
        let textAfter = (endRange.toString() || '').trim().replace(/\s+/g, ' ');
        if (!textAfter) {
            const full = (cell.innerText || '').trim().replace(/\s+/g, ' ');
            textAfter = full.slice(-25).trim() || full.slice(0, 30);
        }
        if (!textAfter) return '';
        const maxLen = 30;
        return textAfter.slice(0, maxLen).trim() + (textAfter.length > maxLen ? '…' : '');
    })();

    const cursorRect = range.getBoundingClientRect();

    /* 1. Numbered circle in first column */
    const numCircle = document.createElement('div');
    numCircle.className = 'bookmark-dot';
    numCircle.id = bookmarkId;
    numCircle.dataset.bookmarkNum = String(nextNum);
    numCircle.dataset.bookmarkLabel = label;

    /* Insert bookmark-dot in row (not col1) so it stays visible when first column is unchecked */
    const rowStyle = window.getComputedStyle(row);
    if (rowStyle.position === 'static') row.style.position = 'relative';

    const rowRect = row.getBoundingClientRect();
    const numCircleHeight = 26;
    const topOffsetNum = cursorRect.top + cursorRect.height / 2 - rowRect.top - numCircleHeight / 2;
    const rowHeight = rowRect.height;
    const topPercentNum = rowHeight > 0 ? Math.max(0, Math.min(100, (topOffsetNum / rowHeight) * 100)) : 0;

    numCircle.style.position = 'absolute';
    numCircle.style.top = topPercentNum + '%';
    numCircle.style.left = '';
    numCircle.dataset.bookmarkId = stableId;

    row.insertBefore(numCircle, row.firstChild);

    /* 2. Small red dot inline - dot + first word in nowrap (no break after dot), rest wraps normally */
    const cursorDot = document.createElement('span');
    cursorDot.className = 'bookmark-cursor-dot';
    cursorDot.dataset.bookmarkId = stableId;

    const bulletSpan = document.createElement('span');
    bulletSpan.className = 'bookmark-cursor-dot-bullet';
    bulletSpan.textContent = '\u2022';

    range.insertNode(cursorDot);
    cursorDot.appendChild(bulletSpan);

    const noBreakSpan = document.createElement('span');
    noBreakSpan.className = 'bookmark-cursor-dot-nobreak';
    cursorDot.appendChild(noBreakSpan);

    let n = cursorDot.nextSibling;
    if (n && n.nodeType === Node.TEXT_NODE && n.textContent.length > 0) {
        const rest = n.textContent;
        const wordMatch = rest.match(/^(\S+(?:\s|$)?)/);
        const firstWord = wordMatch ? wordMatch[1] : rest.charAt(0);
        n.textContent = rest.slice(firstWord.length);
        noBreakSpan.textContent = firstWord;
    }

    /* Move following text into cursorDot (no extra wrap span - avoids break) */
    n = cursorDot.nextSibling;
    while (n) {
        const sib = n;
        n = n.nextSibling;
        if (sib.nodeType === Node.ELEMENT_NODE) {
            const tag = sib.tagName?.toUpperCase();
            if (['DIV', 'P', 'BR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'UL', 'OL', 'LI'].includes(tag)) break;
        }
        cursorDot.appendChild(sib);
    }

    range.setStart(cursorDot, cursorDot.childNodes.length);
    range.collapse(true);
    sel.removeAllRanges();
    sel.addRange(range);
    renumberBookmarks();
    requestAnimationFrame(() => {
        updateBookmarkPositions();
        const parent = cursorDot.parentNode;
        if (parent) {
            if (cursorDot.nextSibling?.nodeName === 'BR') cursorDot.nextSibling.remove();
            const prev = cursorDot.previousSibling;
            if (prev?.nodeName === 'BR' && prev.previousSibling && !prev.previousSibling.textContent?.trim()) prev.remove();
        }
    });
    syncEditorState();
    updateBookmarkSidebar();
    if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
}

function renumberBookmarks() {
    const bars = getSortedBookmarkDots();
    bars.forEach((bar, i) => {
        const num = String(i + 1);
        bar.id = `bookmark-${num}`;
        bar.dataset.bookmarkNum = num;
    });
}

function getSortedBookmarkDots() {
    const el = document.getElementById('teleprompter-text');
    if (!el) return [];
    const dots = Array.from(el.querySelectorAll('.bookmark-dot'));
    return dots.sort((a, b) => {
        const cursorA = el.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${a.dataset.bookmarkId}"]`);
        const cursorB = el.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${b.dataset.bookmarkId}"]`);
        const posA = cursorA ? (cursorA.getBoundingClientRect?.().top ?? a.getBoundingClientRect?.().top ?? 0) : (a.getBoundingClientRect?.().top ?? 0);
        const posB = cursorB ? (cursorB.getBoundingClientRect?.().top ?? b.getBoundingClientRect?.().top ?? 0) : (b.getBoundingClientRect?.().top ?? 0);
        return posA - posB;
    });
}

function getCurrentBookmarkIndex() {
    const view = document.getElementById('teleprompter-view');
    const bars = getSortedBookmarkDots();
    if (!view || bars.length === 0) return -1;
    const viewRect = view.getBoundingClientRect();
    const viewCenterY = viewRect.top + viewRect.height / 2;
    const getCenterY = (bar) => {
        const cursor = teleprompterText?.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${bar.dataset.bookmarkId}"]`);
        const el = cursor || bar;
        const r = el.getBoundingClientRect();
        return r.top + r.height / 2;
    };
    const firstCenterY = getCenterY(bars[0]);
    const lastCenterY = getCenterY(bars[bars.length - 1]);
    if (viewCenterY < firstCenterY) return -1;
    if (viewCenterY > lastCenterY) return bars.length;
    let bestIdx = 0;
    let bestDist = Infinity;
    bars.forEach((el, i) => {
        const centerY = getCenterY(el);
        const dist = Math.abs(centerY - viewCenterY);
        if (dist < bestDist) {
            bestDist = dist;
            bestIdx = i;
        }
    });
    return bestIdx;
}

const BOOKMARK_SCROLL_DURATION_MS = 280;

function scrollViewToCenterElement(view, targetEl, onComplete) {
    if (!view || !targetEl) return;
    const viewRect = view.getBoundingClientRect();
    const targetRect = targetEl.getBoundingClientRect();
    const targetCenterY = targetRect.top + targetRect.height / 2;
    const viewCenterY = viewRect.top + viewRect.height / 2;
    const scrollDelta = targetCenterY - viewCenterY;
    const startTop = view.scrollTop;
    const endTop = Math.max(0, Math.min(view.scrollHeight - view.clientHeight, startTop + scrollDelta));
    if (endTop === startTop) {
        if (onComplete) onComplete();
        return;
    }
    const startTime = performance.now();
    function tick(now) {
        const elapsed = now - startTime;
        const t = Math.min(1, elapsed / BOOKMARK_SCROLL_DURATION_MS);
        const ease = 1 - Math.pow(1 - t, 2);
        view.scrollTop = startTop + (endTop - startTop) * ease;
        if (t < 1) {
            requestAnimationFrame(tick);
        } else if (onComplete) {
            onComplete();
        }
    }
    requestAnimationFrame(tick);
}

function goToPrevBookmark() {
    const view = document.getElementById('teleprompter-view');
    const bars = getSortedBookmarkDots();
    if (!view || bars.length === 0) return;
    const currentIdx = getCurrentBookmarkIndex();
    const prevIdx = currentIdx <= 0 ? 0 : (currentIdx >= bars.length - 1 ? Math.max(0, bars.length - 2) : currentIdx - 1);
    const target = bars[prevIdx];
    bookmarkNavigationInProgress = true;
    setActiveBookmarkByIndex(prevIdx);
    scrollViewToCenterElement(view, target, () => { bookmarkNavigationInProgress = false; });
    view.focus();
}

function goToNextBookmark() {
    const view = document.getElementById('teleprompter-view');
    const bars = getSortedBookmarkDots();
    if (!view || bars.length === 0) return;
    const currentIdx = getCurrentBookmarkIndex();
    const nextIdx = currentIdx >= bars.length - 1 ? bars.length - 1 : (currentIdx < 0 ? 0 : currentIdx + 1);
    const target = bars[nextIdx];
    bookmarkNavigationInProgress = true;
    setActiveBookmarkByIndex(nextIdx);
    scrollViewToCenterElement(view, target, () => { bookmarkNavigationInProgress = false; });
    view.focus();
}

function updateBookmarkPositions() {
    const dots = teleprompterText.querySelectorAll('.bookmark-dot');
    dots.forEach((numCircle) => {
        const bid = numCircle.dataset.bookmarkId;
        if (!bid) return;
        const cursorDot = teleprompterText.querySelector(`.bookmark-cursor-dot[data-bookmark-id="${bid}"]`);
        if (!cursorDot) return;
        const bullet = cursorDot.querySelector('.bookmark-cursor-dot-bullet');
        const alignEl = bullet || cursorDot;
        const row = numCircle.closest('.script-row-wrapper') || numCircle.parentElement;
        const textCol = cursorDot.closest('.script-column');
        const container = row || numCircle.parentElement;
        if (!container) return;
        const style = window.getComputedStyle(container);
        if (style.position === 'static') container.style.position = 'relative';
        const cursorFontSize = parseFloat(window.getComputedStyle(alignEl).fontSize) || 12;
        const circleSize = Math.max(26, Math.min(52, cursorFontSize * 0.65));
        numCircle.style.width = circleSize + 'px';
        numCircle.style.height = circleSize + 'px';
        numCircle.style.fontSize = Math.max(12, Math.min(22, Math.round(circleSize * 0.5))) + 'px';
        const numCircleRect = numCircle.getBoundingClientRect();
        const numCircleHeight = numCircleRect.height;
        const alignRect = alignEl.getBoundingClientRect();
        const alignCenterY = alignRect.top + alignRect.height / 2;
        const refContainer = textCol || container;
        const refRect = refContainer.getBoundingClientRect();
        const topOffset = alignCenterY - refRect.top - numCircleHeight / 2;
        const containerHeight = container.getBoundingClientRect().height;
        const refHeight = refRect.height;
        const topPercent = refHeight > 0 ? Math.max(0, Math.min(100, (topOffset / refHeight) * 100)) : 0;
        numCircle.style.top = topPercent + '%';
    });
}

/** Wrap flex cell content in a block so spans stay inline (prevents color-spans from creating line breaks) */
function wrapCellContentInBlock() {
    const wrap = (container) => {
        if (!container || !container.firstChild) return;
        const first = container.firstChild;
        if (first.nodeType === Node.ELEMENT_NODE && first.classList?.contains('cell-content')) return;
        const wrapDiv = document.createElement('div');
        wrapDiv.className = 'cell-content';
        while (container.firstChild) wrapDiv.appendChild(container.firstChild);
        container.appendChild(wrapDiv);
    };
    teleprompterText.querySelectorAll('.script-column').forEach(col => wrap(col));
    teleprompterText.querySelectorAll('.cell-locker').forEach(cell => wrap(cell));
}

const BKMK_PLACEHOLDER = '{BKMK}';

/** Collect all text nodes under root in document order (recursive, no TreeWalker). */
function getTextNodesUnder(root) {
    const out = [];
    function walk(n) {
        if (!n || !root.contains(n)) return;
        if (n.nodeType === Node.TEXT_NODE) {
            out.push(n);
            return;
        }
        for (let i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
    }
    for (let i = 0; i < root.childNodes.length; i++) walk(root.childNodes[i]);
    return out;
}

/** Find next "{BKMK}" in document order; returns { textNode, offset } or null. */
function findNextBkmkPlaceholder() {
    const textNodes = getTextNodesUnder(teleprompterText);
    for (let i = 0; i < textNodes.length; i++) {
        const text = textNodes[i].textContent || '';
        const idx = text.indexOf(BKMK_PLACEHOLDER);
        if (idx !== -1) return { textNode: textNodes[i], offset: idx };
    }
    return null;
}

/** Find all "{BKMK}" in the editor and replace each with a real bookmark (number circle + cursor dot). Call after layout is ready. */
function convertBkmkPlaceholdersToBookmarks() {
    const initialDotCount = teleprompterText.querySelectorAll('.bookmark-dot').length;
    let nextNum = initialDotCount;
    let hit;
    const rawHtml = teleprompterText.innerHTML || '';
    const hasPlaceholderInHtml = rawHtml.indexOf(BKMK_PLACEHOLDER) !== -1;
    console.log('[BKMK] convertBkmkPlaceholdersToBookmarks called. initialDotCount=', initialDotCount, 'innerHTML contains "{BKMK}":', hasPlaceholderInHtml, 'innerHTML length:', rawHtml.length);
    const textNodes = getTextNodesUnder(teleprompterText);
    console.log('[BKMK] getTextNodesUnder returned', textNodes.length, 'text nodes');
    textNodes.forEach((tn, i) => {
        const t = (tn.textContent || '').slice(0, 80);
        const hasBkmk = (tn.textContent || '').indexOf(BKMK_PLACEHOLDER) !== -1;
        if (hasBkmk || t.trim()) console.log('[BKMK] textNode[' + i + '] length=', (tn.textContent || '').length, 'hasBKMK=', hasBkmk, 'preview:', JSON.stringify(t));
    });
    while ((hit = findNextBkmkPlaceholder())) {
        const { textNode, offset } = hit;
        console.log('[BKMK] found placeholder at offset', offset, 'in textNode, text length=', (textNode.textContent || '').length);
        if (!teleprompterText.contains(textNode)) {
            console.log('[BKMK] skip: textNode not contained in teleprompterText');
            continue;
        }
        const text = textNode.textContent || '';
        if (text.indexOf(BKMK_PLACEHOLDER, offset) !== offset) {
            console.log('[BKMK] skip: placeholder not at expected offset');
            continue;
        }
        const parentEl = textNode.parentElement;
        if (!parentEl) {
            console.log('[BKMK] skip: textNode has no parentElement');
            continue;
        }
        let row = parentEl.closest('.script-row-wrapper');
        const cursorCol = parentEl.closest('.script-column');
        if (!row) row = parentEl.closest('div, p') || teleprompterText;
        const col1 = row ? row.querySelector('.script-column') : null;
        console.log('[BKMK] row=', !!row, row?.className || row?.nodeName, 'col1=', !!col1, 'cursorCol=', !!cursorCol);
        if (!row || !teleprompterText.contains(row)) {
            console.log('[BKMK] skip: no row or row not in editor');
            continue;
        }
        nextNum += 1;
        const stableId = `bk-${Date.now()}-${nextNum}-${Math.random().toString(36).slice(2)}`;
        const range = document.createRange();
        try {
            range.setStart(textNode, offset);
            range.setEnd(textNode, offset + BKMK_PLACEHOLDER.length);
        } catch (err) {
            console.log('[BKMK] skip: range set failed', err);
            continue;
        }
        range.deleteContents();
        range.collapse(true);
        const cursorDot = document.createElement('span');
        cursorDot.className = 'bookmark-cursor-dot';
        cursorDot.dataset.bookmarkId = stableId;
        const bulletSpan = document.createElement('span');
        bulletSpan.className = 'bookmark-cursor-dot-bullet';
        bulletSpan.textContent = '\u2022';
        cursorDot.appendChild(bulletSpan);
        const noBreakSpan = document.createElement('span');
        noBreakSpan.className = 'bookmark-cursor-dot-nobreak';
        cursorDot.appendChild(noBreakSpan);
        range.insertNode(cursorDot);
        const cell = cursorCol || row;
        const fullCellText = (cell.innerText || '').trim().replace(/\s+/g, ' ');
        const afterBullet = fullCellText.split(/\s*\u2022\s*/).slice(1).join(' ').trim() || fullCellText;
        const label = afterBullet ? (afterBullet.slice(0, 30).trim() + (afterBullet.length > 30 ? '…' : '')) : '';
        /* Insert bookmark-dot in row (not col1) so it stays visible when first column is unchecked */
        if (row.style.position === 'static' || !row.style.position) row.style.position = 'relative';
        const numCircle = document.createElement('div');
        numCircle.className = 'bookmark-dot';
        numCircle.dataset.bookmarkNum = String(nextNum);
        numCircle.dataset.bookmarkLabel = label;
        numCircle.dataset.bookmarkId = stableId;
        numCircle.style.position = 'absolute';
        numCircle.style.top = '0%';
        numCircle.style.left = '';
        row.insertBefore(numCircle, row.firstChild);
        console.log('[BKMK] inserted bookmark', nextNum, 'stableId=', stableId);
    }
    if (nextNum <= initialDotCount) {
        console.log('[BKMK] no bookmarks added, returning early (nextNum=', nextNum, 'initialDotCount=', initialDotCount, ')');
        return;
    }
    console.log('[BKMK] done: added', nextNum - initialDotCount, 'bookmarks, calling renumberBookmarks + rAF');
    renumberBookmarks();
    requestAnimationFrame(() => {
        updateBookmarkPositions();
        updateBookmarkSidebar();
    });
}

function applyColumnVisibilityToEditor() {
    if (currentFileIndex < 0 || !fileColumnVisibility[currentFileIndex]) return;
    const vis = fileColumnVisibility[currentFileIndex];
    teleprompterText.querySelectorAll('.script-row-wrapper').forEach(row => {
        row.querySelectorAll('.script-column').forEach((col, i) => {
            col.classList.toggle('user-col-hidden', vis[i] === false);
        });
    });
    syncColumnWidths();
    if (document.body.classList.contains('broadcasting')) {
        if (broadcastEditMode) syncRowHeightsFromMainInBroadcastEditMode();
        else measureRowHeightsWithProbeForBroadcasting();
        if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
    } else if (mirrorWindow && !mirrorWindow.closed) {
        refreshMirrorData();
    }
}

function updateRunlistRowColumnToggles(fileIndex) {
    const row = runlistContainer && fileIndex >= 0 ? runlistContainer.querySelector(`.runlist-row[data-index="${fileIndex}"]`) : null;
    const togglesEl = row ? row.querySelector('.runlist-column-toggles') : null;
    if (!togglesEl) return;
    togglesEl.innerHTML = '';
    const n = fileColumnCount[fileIndex] || 0;
    const vis = fileColumnVisibility[fileIndex];
    if (n <= 0 || !vis) return;
    for (let i = 0; i < n; i++) {
        const label = document.createElement('label');
        label.className = 'runlist-col-toggle';
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.checked = vis[i] !== false;
        cb.dataset.colIndex = String(i);
        cb.addEventListener('change', (e) => {
            e.stopPropagation();
            const colIndex = parseInt(cb.dataset.colIndex, 10);
            if (!fileColumnVisibility[fileIndex]) return;
            fileColumnVisibility[fileIndex][colIndex] = cb.checked;
            /* Ensure the file whose toggles we changed is shown in the teleprompter */
            if (currentFileIndex !== fileIndex) {
                loadScriptToEditor(fileIndex);
                return;
            }
            teleprompterText.querySelectorAll('.script-row-wrapper').forEach(r => {
                const cols = r.querySelectorAll('.script-column');
                if (cols[colIndex]) cols[colIndex].classList.toggle('user-col-hidden', !cb.checked);
            });
            syncColumnWidths();
            if (document.body.classList.contains('broadcasting')) {
                if (broadcastEditMode) syncRowHeightsFromMainInBroadcastEditMode();
                else measureRowHeightsWithProbeForBroadcasting();
                if (mirrorWindow && !mirrorWindow.closed) refreshMirrorData();
            } else if (mirrorWindow && !mirrorWindow.closed) {
                refreshMirrorData();
            }
        });
        label.appendChild(cb);
        const labelText = (i === 0 && fileFirstColIsId[fileIndex]) ? 'ID' : String(i + 1);
        label.appendChild(document.createTextNode(' ' + labelText));
        togglesEl.appendChild(label);
    }
}

const KEYWORD_PILL_RED = ['end', 'full screen', 'stop', 'out'];
const KEYWORD_PILL_YELLOW_EXACT = ['panel', 'chorus'];
const KEYWORD_PILL_YELLOW_VERSE = /^verse\s*\d*$/i;
const KEYWORD_PILL_YELLOW_NAME = /^(elder|president|sister|brother)\s+.+$/i;
/** Text that starts with [ and ends with ] -> yellow pill */
const KEYWORD_PILL_YELLOW_BRACKETS = /^\[.*\]$/;
/** Matches "Name (00:00 – 18:41)" or "Brianna (1:23 - 5:00)" - speaker/time rows -> blue pill */
const KEYWORD_PILL_BLUE_NAME_TIME = /^[A-Za-z][A-Za-z0-9\s\-'.]*\s*\([^)]+\)\s*$/;

function getInterpreterFromFilename(name) {
    if (!name || typeof name !== 'string') return '';
    const base = stripFileExtension(name);
    const parts = base.split('_');
    return parts.length >= 2 ? (parts[1] || '').trim() : '';
}

function isCurrentFileXlsx() {
    if (currentFileIndex < 0 || currentFileIndex >= fileStore.length) return false;
    const name = fileStore[currentFileIndex].name;
    if (!name || typeof name !== 'string') return false;
    const ext = name.split('.').pop().toLowerCase();
    return ext === 'xlsx' || ext === 'xls';
}

function applyKeywordPills() {
    if (!teleprompterText) return;
    if (isCurrentFileXlsx()) return;
    const rows = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
    const currentInterpreter = (currentFileIndex >= 0 && fileStore[currentFileIndex]) ? getInterpreterFromFilename(fileStore[currentFileIndex].name) : '';
    const nextInterpreter = (currentFileIndex >= 0 && currentFileIndex + 1 < fileStore.length && fileStore[currentFileIndex + 1]) ? getInterpreterFromFilename(fileStore[currentFileIndex + 1].name) : '';
    const switchLabel = (nextInterpreter && nextInterpreter !== currentInterpreter) ? 'SWITCH' : 'STAY';

    rows.forEach(row => {
        const raw = (row.textContent || '').trim();
        const lower = raw.toLowerCase();
        row.classList.remove('keyword-pill-red', 'keyword-pill-yellow', 'keyword-pill-green', 'keyword-pill-blue', 'keyword-pill-white');

        if (KEYWORD_PILL_RED.includes(lower)) {
            row.classList.add('keyword-pill-red');
        } else if (['switch', 'stay'].includes(lower)) {
            row.classList.add('keyword-pill-green');
            setFirstVisibleCellText(row, switchLabel);
        } else if (KEYWORD_PILL_YELLOW_EXACT.includes(lower) || KEYWORD_PILL_YELLOW_VERSE.test(raw) || KEYWORD_PILL_YELLOW_NAME.test(raw) || KEYWORD_PILL_YELLOW_BRACKETS.test(raw)) {
            row.classList.add('keyword-pill-yellow');
        } else if (KEYWORD_PILL_BLUE_NAME_TIME.test(raw)) {
            row.classList.add('keyword-pill-blue');
        } else if (raw.includes('|v')) {
            row.classList.add('keyword-pill-white');
        }
        if (row.classList.contains('keyword-pill-red') || row.classList.contains('keyword-pill-yellow') || row.classList.contains('keyword-pill-green') || row.classList.contains('keyword-pill-blue') || row.classList.contains('keyword-pill-white')) {
            row.querySelectorAll('.script-column, .cell-content, .cell-locker, .script-column *').forEach(el => {
                if (el && el.style && el.style.fontSize) el.style.removeProperty('font-size');
            });
        }
    });
}

function setFirstVisibleCellText(row, text) {
    const cols = row.querySelectorAll('.script-column');
    for (let i = 0; i < cols.length; i++) {
        const col = cols[i];
        if (col.classList.contains('user-col-hidden') || col.classList.contains('broadcast-hidden')) continue;
        const cell = col.querySelector('.cell-content') || col.querySelector('.cell-locker') || col;
        if (cell) {
            cell.textContent = text;
            break;
        }
    }
}

function stripFileExtension(name) {
    if (!name || typeof name !== 'string') return '';
    const lastDot = name.lastIndexOf('.');
    return lastDot > 0 ? name.slice(0, lastDot) : name;
}

function updateFilenamePills() {
    const topPill = document.getElementById('filename-pill-top');
    const bottomPill = document.getElementById('filename-pill-bottom');
    const staySwitchPill = document.getElementById('stay-switch-pill');
    const view = document.getElementById('teleprompter-view');
    if (!topPill || !bottomPill) return;
    const hasCurrent = currentFileIndex >= 0 && fileStore.length > 0 && fileStore[currentFileIndex];
    const hasNext = hasCurrent && currentFileIndex + 1 < fileStore.length;
    if (view) view.classList.toggle('has-file', !!hasCurrent);
    if (hasCurrent) {
        topPill.textContent = stripFileExtension(fileStore[currentFileIndex].name);
        topPill.classList.remove('hidden');
        topPill.setAttribute('aria-hidden', 'false');
    } else {
        topPill.textContent = '';
        topPill.classList.add('hidden');
        topPill.setAttribute('aria-hidden', 'true');
    }
    if (hasNext) {
        bottomPill.textContent = stripFileExtension(fileStore[currentFileIndex + 1].name);
        bottomPill.classList.remove('hidden');
        bottomPill.setAttribute('aria-hidden', 'false');
    } else {
        bottomPill.textContent = '';
        bottomPill.classList.add('hidden');
        bottomPill.setAttribute('aria-hidden', 'true');
    }
    /* Stay/Switch pill: show when next file exists (green STAY/SWITCH); when last file, show red END */
    if (staySwitchPill) {
        if (hasNext) {
            const currentInterpreter = getInterpreterFromFilename(fileStore[currentFileIndex].name);
            const nextInterpreter = getInterpreterFromFilename(fileStore[currentFileIndex + 1].name);
            const label = nextInterpreter && nextInterpreter !== currentInterpreter ? 'SWITCH' : 'STAY';
            const nextName = nextInterpreter || stripFileExtension(fileStore[currentFileIndex + 1].name) || '';
            staySwitchPill.textContent = nextName ? label + ' ' + nextName : label;
            staySwitchPill.classList.remove('hidden', 'end-pill');
            staySwitchPill.setAttribute('aria-hidden', 'false');
        } else if (hasCurrent) {
            staySwitchPill.textContent = 'END';
            staySwitchPill.classList.add('end-pill');
            staySwitchPill.classList.remove('hidden');
            staySwitchPill.setAttribute('aria-hidden', 'false');
        } else {
            staySwitchPill.textContent = '';
            staySwitchPill.classList.remove('end-pill');
            staySwitchPill.classList.add('hidden');
            staySwitchPill.setAttribute('aria-hidden', 'true');
        }
    }
    /* Sync font size and family from teleprompter text to pills */
    if (teleprompterText) {
        const style = window.getComputedStyle(teleprompterText);
        const fontSize = style.fontSize;
        const fontFamily = style.fontFamily;
        const pillsToSync = [topPill, bottomPill];
        if (staySwitchPill && !staySwitchPill.classList.contains('hidden')) pillsToSync.push(staySwitchPill);
        pillsToSync.forEach(p => {
            if (fontSize) p.style.fontSize = fontSize;
            if (fontFamily) p.style.fontFamily = fontFamily;
        });
    }
    shrinkAllPillsToFit();
}

const PILL_SHRINK_MIN_PX = 10;

/** Shrink one pill's font-size until its content fits without overflow (no wrap). */
function shrinkPillTextToFit(pillEl, minPx) {
    if (!pillEl || pillEl.offsetParent === null) return;
    minPx = minPx ?? PILL_SHRINK_MIN_PX;
    const computed = window.getComputedStyle(pillEl);
    let size = parseFloat(computed.fontSize) || 16;
    const setSize = (px) => {
        pillEl.style.setProperty('font-size', px + 'px', 'important');
    };
    setSize(size);
    const col = pillEl.closest('.script-column');
    const maxW = col ? (col.clientWidth || pillEl.clientWidth) : pillEl.clientWidth;
    if (maxW <= 0) return;
    while (pillEl.scrollWidth > maxW && size > minPx) {
        size = Math.max(minPx, size - 2);
        setSize(size);
    }
}

/** Run shrink-to-fit on all visible pills (filename, stay/switch, keyword rows). */
function shrinkAllPillsToFit() {
    const view = document.getElementById('teleprompter-view');
    if (!view) return;
    const pills = [
        document.getElementById('filename-pill-top'),
        document.getElementById('filename-pill-bottom'),
        document.getElementById('stay-switch-pill')
    ].filter(Boolean);
    pills.forEach(p => {
        if (!p.classList.contains('hidden')) shrinkPillTextToFit(p);
    });
    const keywordRows = teleprompterText ? teleprompterText.querySelectorAll('.script-row-wrapper.keyword-pill-red, .script-row-wrapper.keyword-pill-yellow, .script-row-wrapper.keyword-pill-green, .script-row-wrapper.keyword-pill-blue, .script-row-wrapper.keyword-pill-white') : [];
    keywordRows.forEach(row => {
        row.querySelectorAll('.cell-content').forEach(cell => {
            cell.style.removeProperty('font-size');
            shrinkPillTextToFit(cell);
        });
    });
}

function playPageTurnAndAdvanceToNextFile(nextIndex) {
    const overlay = document.getElementById('page-turn-overlay');
    if (overlay) overlay.classList.add('page-turn-active');
    const DURATION_MS = 500;
    const LOAD_AT_MS = 150;
    setTimeout(() => {
        if (typeof loadScriptToEditor === 'function') loadScriptToEditor(nextIndex, { scrollToTop: true, triggerNewTalkPill: true });
    }, LOAD_AT_MS);
    setTimeout(() => {
        if (overlay) overlay.classList.remove('page-turn-active');
    }, DURATION_MS);
}

function playPageTurnAndGoToPreviousFile(prevIndex) {
    const overlay = document.getElementById('page-turn-overlay');
    if (overlay) overlay.classList.add('page-turn-active');
    const DURATION_MS = 500;
    const LOAD_AT_MS = 150;
    setTimeout(() => {
        if (typeof loadScriptToEditor === 'function') loadScriptToEditor(prevIndex, { scrollToBottom: true });
    }, LOAD_AT_MS);
    setTimeout(() => {
        if (overlay) overlay.classList.remove('page-turn-active');
    }, DURATION_MS);
}

function loadScriptToEditor(index, options) {
    console.log(`Attempting to load index: ${index}`);
    bottomPillTriggerFired = false;
    topPillTriggerFired = false;
    lastScrollTopForPillTrigger = null;
    if (fileStore[index] === null) return;

    /* Save current file only when switching to a different file (avoid overwriting with placeholder/empty when reloading same file after sort) */
    if (currentFileIndex >= 0 && currentFileIndex < contentStore.length && currentFileIndex !== index) {
        const html = teleprompterText.innerHTML.trim();
        contentStore[currentFileIndex] = (html === '<br>' || html === '') ? '' : html;
    }

    rowColorsCache = [];
    rowFont12Cache = [];
    document.querySelectorAll('.runlist-row').forEach(r => r.classList.remove('active'));
    const row = document.querySelector(`.runlist-row[data-index="${index}"]`);
    if (row) row.classList.add('active');

    currentFileIndex = index;
    const content = contentStore[index];
    const contentHasBkmk = (content && typeof content === 'string') ? content.indexOf('{BKMK}') !== -1 : false;
    console.log('[BKMK] loadScriptToEditor index=', index, 'content length=', (content && content.length) || 0, 'content contains "{BKMK}":', contentHasBkmk);
    teleprompterText.innerHTML = (content === '' || !content) ? '<br>' : content;
    delete teleprompterText.dataset.placeholder;
    processTableColumns();
    wrapCellContentInBlock();
    removeUuidFromFirstColumn();
    stripPipeVFromFirstColumn();
    applyColumnVisibilityToEditor();
    if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
        contentStore[currentFileIndex] = teleprompterText.innerHTML.trim();
    }

    /* Aa is global: apply current overview state to the loaded file so all files respect the same mode */
    const btnOverviewToggle = document.getElementById('btn-overview-toggle');
    if (btnOverviewToggle && fontSizeSelect) {
        if (isOverviewMode) {
            btnOverviewToggle.classList.add('overview-active');
            document.body.classList.add('overview-mode');
            fontSizeSelect.value = '12';
            lastFontChangeSource = 'size';
            applyFontSizeOnlyToAllContent('12');
        } else {
            btnOverviewToggle.classList.remove('overview-active');
            document.body.classList.remove('overview-mode');
            fontSizeSelect.value = savedFontSizeBeforeOverview || '80';
            lastFontChangeSource = 'size';
            applyFontSizeOnlyToAllContent(fontSizeSelect.value);
        }
    }

    teleprompterText.focus();
    requestAnimationFrame(() => {
        refreshSelectFontTarget();
        syncColumnWidths();
        /* Replace {BKMK} with bookmark (red dot + number in list) after layout so formatting is unchanged */
        console.log('[BKMK] rAF: about to call convertBkmkPlaceholdersToBookmarks for file index', index);
        convertBkmkPlaceholdersToBookmarks();
        applyKeywordPills();
        shrinkAllPillsToFit();
        updateBookmarkSidebar();
        if (currentFileIndex >= 0 && currentFileIndex < contentStore.length) {
            contentStore[currentFileIndex] = teleprompterText.innerHTML.trim();
        }
    });

    if (mirrorWindow && !mirrorWindow.closed) {
        // ❌ REMOVE THIS LINE:
        // mirrorWindow.postMessage({ type: 'loadContent', content: teleprompterText.innerHTML }, '*');
        
        // ✅ KEEP THESE:
        refreshMirrorData(); // This handles the clean text extraction
        syncMirrorStyles();
    }
    updateFilenamePills();
    if (options && options.scrollToTop) {
        const view = document.getElementById('teleprompter-view');
        if (view) view.scrollTop = 0;
    }
    if (options && options.scrollToBottom) {
        const view = document.getElementById('teleprompter-view');
        if (view) view.scrollTop = Math.max(0, view.scrollHeight - view.clientHeight);
    }
    requestAnimationFrame(() => {
        const view = document.getElementById('teleprompter-view');
        if (view) lastScrollTopForPillTrigger = view.scrollTop;
    });
    if (options && options.triggerNewTalkPill) {
        const topPill = document.getElementById('filename-pill-top');
        if (topPill) {
            topPill.classList.remove('new-talk');
            void topPill.offsetWidth;
            topPill.classList.add('new-talk');
            setTimeout(() => topPill.classList.remove('new-talk'), 600);
        }
    }
    console.log("Editor and Mirror updated with Content and Styles.");
}
    
// =========================================
// Second monitor position (when getScreenDetails not available)
// =========================================
const SECOND_MONITOR_STORAGE_KEY = 'teleprompter_secondMonitorPosition';
function getSecondMonitorPosition() {
    const el = document.querySelector('input[name="secondMonitorPosition"]:checked');
    if (el) return el.value;
    try {
        return localStorage.getItem(SECOND_MONITOR_STORAGE_KEY) || 'right';
    } catch (_) {
        return 'right';
    }
}

if (secondMonitorPositionRadios && secondMonitorPositionRadios.length) {
    try {
        const saved = localStorage.getItem(SECOND_MONITOR_STORAGE_KEY);
        if (saved) {
            secondMonitorPositionRadios.forEach((r) => {
                r.checked = r.value === saved;
            });
        }
    } catch (_) {}
    secondMonitorPositionRadios.forEach((r) => {
        r.addEventListener('change', () => {
            try {
                localStorage.setItem(SECOND_MONITOR_STORAGE_KEY, r.value);
            } catch (_) {}
        });
    });
}

const btnRequestMultiscreen = document.getElementById('btn-request-multiscreen');
const multiscreenStatusEl = document.getElementById('multiscreen-status');
if (btnRequestMultiscreen && multiscreenStatusEl) {
    btnRequestMultiscreen.addEventListener('click', async () => {
        multiscreenStatusEl.textContent = '';
        const isFileProtocol = window.location.protocol === 'file:';
        if (isFileProtocol) {
            multiscreenStatusEl.textContent = 'Open the app from http://localhost (not the file directly). Run "npx serve" in this folder, then open http://localhost:3000 in Chrome.';
            return;
        }
        if (!('getScreenDetails' in window)) {
            multiscreenStatusEl.textContent = 'Multi-screen API not available (Chrome on Mac or file://). Use Left/Right; if the mirror doesn\'t move, drag it to your other monitor.';
            return;
        }
        try {
            const screenDetails = await window.getScreenDetails();
            const n = screenDetails.screens ? screenDetails.screens.length : 0;
            multiscreenStatusEl.textContent = n > 1 ? 'Access allowed (' + n + ' screen(s)).' : 'Allowed (1 screen).';
        } catch (err) {
            multiscreenStatusEl.textContent = 'Denied or failed. Using Left/Right for mirror.';
        }
    });
}

// =========================================
// 6. MIRROR WINDOW LOGIC (OS AWARE)
// =========================================

if (newBookmarkButton) newBookmarkButton.onclick = () => addBookmarkAtCursor();
if (prevBookmarkButton) prevBookmarkButton.onclick = () => goToPrevBookmark();
if (nextBookmarkButton) nextBookmarkButton.onclick = () => goToNextBookmark();

extendMonitorButton.onclick = async () => {
    const isExtended = document.body.classList.contains('broadcasting') && mirrorWindow && !mirrorWindow.closed;

    if (isExtended) {
        const maxScroll = teleprompterView.scrollHeight - teleprompterView.clientHeight;
        const scrollRatio = maxScroll > 0 ? teleprompterView.scrollTop / maxScroll : 0;

        if (mirrorCloseCheckInterval) {
            clearInterval(mirrorCloseCheckInterval);
            mirrorCloseCheckInterval = null;
        }
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

    /* If user had Aa (overview) on, temporarily exit so we measure row heights with normal font (same as Extend then Aa). */
    const wasInOverview = document.body.classList.contains('overview-mode') || isOverviewMode;
    if (wasInOverview) {
        const rawRestore = savedFontSizeBeforeOverview || '80';
        const restore = (rawRestore && rawRestore !== '12') ? rawRestore : '80';
        if (fontSizeSelect) {
            fontSizeSelect.value = restore;
            lastFontChangeSource = 'size';
        }
        applyFontSizeOnlyToAllContent(restore);
        isOverviewMode = false;
        document.body.classList.remove('overview-mode');
        const btnO = document.getElementById('btn-overview-toggle');
        if (btnO) btnO.classList.remove('overview-active');
        flattenRedundantSpans();
        void teleprompterText.offsetHeight;
        syncColumnWidths();
    }

    document.body.classList.add('broadcasting');
    void teleprompterText.offsetHeight;
    applyBroadcastingVisibility();
    syncColumnWidths();
    restoreScrollPosition();
    const btnOverview = document.getElementById('btn-overview-toggle');
    if (btnOverview) {
        btnOverview.classList.remove('broadcast-edit-active');
        btnOverview.title = 'Show all columns to edit (mirror stays open)';
    }

    try {
        let currentScreen = null;
        let secondaryScreen = null;

        if ('getScreenDetails' in window) {
            const screenDetails = await window.getScreenDetails();
            currentScreen = screenDetails.currentScreen;
            secondaryScreen = screenDetails.screens.find(s => s !== currentScreen) || null;
        }

        /* 5:4 aspect ratio, fixed for both views */
        const availW = currentScreen ? currentScreen.availWidth : (window.screen.availWidth || 800);
        const availH = currentScreen ? currentScreen.availHeight : (window.screen.availHeight || 600);
        const ratio54 = 5 / 4;
        let extendW = Math.min(availW, Math.floor(availH * ratio54));
        let extendH = Math.floor(extendW / ratio54);
        if (extendH > availH) {
            extendH = availH;
            extendW = Math.floor(extendH * ratio54);
        }
        extendedWindowWidth = extendW;
        extendedWindowHeight = extendH;

        if (currentScreen) {
            window.moveTo(currentScreen.availLeft, currentScreen.availTop);
        } else {
            window.moveTo(0, 0);
        }
        window.resizeTo(extendedWindowWidth, extendedWindowHeight);

        /* Fix table width so it doesn't change on resize; subtract left gap so content fits and 5:4 is preserved */
        const extendLeftGap = 32;
        extendedFixedWidth = Math.max(100, (window.innerWidth || document.documentElement.clientWidth || extendW) - extendLeftGap);
        teleprompterText.style.width = extendedFixedWidth + 'px';
        teleprompterText.style.maxWidth = extendedFixedWidth + 'px';

        /* Table is now at final width; sync column widths again so layout is correct, then measure row heights */
        syncColumnWidths();
        void teleprompterText.offsetHeight;
        const rowsForMeasure = Array.from(teleprompterText.querySelectorAll('.script-row-wrapper'));
        const scriptColWidth = lastColumnWidthPx;
        const mainStyle = window.getComputedStyle(teleprompterText);
        const probe = document.createElement('div');
        probe.style.position = 'absolute';
        probe.style.left = '-9999px';
        probe.style.top = '0';
        probe.style.width = (scriptColWidth != null && scriptColWidth > 0 ? scriptColWidth : 400) + 'px';
        probe.style.fontSize = mainStyle.fontSize || '';
        probe.style.fontFamily = mainStyle.fontFamily || '';
        probe.style.lineHeight = '1.1';
        probe.style.whiteSpace = 'pre-wrap';
        probe.style.wordWrap = 'break-word';
        probe.style.padding = '0';
        probe.style.margin = '0';
        probe.style.boxSizing = 'border-box';
        document.body.appendChild(probe);
        measuredRowHeights = rowsForMeasure.map(row => {
            const cols = row.querySelectorAll('.script-column');
            const visibleCols = Array.from(cols).slice(0, -1);
            let visibleMax = 1;
            if (visibleCols.length > 0) {
                const visibleHeights = visibleCols.map(col => {
                    const cell = col.querySelector('.cell-locker') || col.querySelector('.cell-content') || col;
                    return cell ? (cell.scrollHeight || 0) : 0;
                });
                visibleMax = Math.max(1, ...visibleHeights);
            }
            const lastCol = cols[cols.length - 1];
            let mirrorH = 0;
            if (lastCol && scriptColWidth != null && scriptColWidth > 0) {
                const cell = lastCol.querySelector('.cell-locker') || lastCol.querySelector('.cell-content') || lastCol;
                const text = cell ? (cell.innerText || '').trim() || '\u00A0' : '\u00A0';
                probe.textContent = text;
                probe.style.fontSize = row.classList.contains('row-font-12') ? '12px' : (mainStyle.fontSize || '');
                probe.style.lineHeight = row.classList.contains('row-font-12') ? '1.2' : '1.1';
                mirrorH = probe.scrollHeight || 0;
            }
            return Math.max(1, visibleMax, mirrorH);
        });
        probe.remove();
        rowsForMeasure.forEach((row, i) => {
            const h = measuredRowHeights[i];
            if (h > 0) {
                row.style.minHeight = h + 'px';
                row.style.height = h + 'px';
            }
        });
        restoreScrollPosition();

        /* Position mirror on other monitor: use getScreenDetails when available, else user's "second monitor position" setting */
        let mirrorLeft, mirrorTop;
        if (secondaryScreen) {
            mirrorLeft = secondaryScreen.availLeft;
            mirrorTop = secondaryScreen.availTop;
        } else {
            const pos = getSecondMonitorPosition();
            const pw = window.screen.availWidth || 1920;
            const ph = window.screen.availHeight || 1080;
            const gap = 50;
            const assumedOtherW = 1920;
            const assumedOtherH = 1080;
            switch (pos) {
                case 'left':
                    mirrorLeft = -assumedOtherW - gap;
                    mirrorTop = 0;
                    break;
                case 'above':
                    mirrorLeft = 0;
                    mirrorTop = -assumedOtherH - gap;
                    break;
                case 'below':
                    mirrorLeft = 0;
                    mirrorTop = ph + gap;
                    break;
                case 'right':
                default:
                    mirrorLeft = Math.max(pw + gap, 2560);
                    mirrorTop = 0;
                    break;
            }
        }
        /* Same 5:4 ratio and fixed width as main: both views use extendedWindowWidth × extendedWindowHeight so teleprompter layout matches. */
        const mirrorW = extendedWindowWidth;
        const mirrorH = extendedWindowHeight;
        const specs = `left=${mirrorLeft},top=${mirrorTop},width=${mirrorW},height=${mirrorH},toolbar=no,menubar=no,location=no,status=no,scrollbars=no`;

        if (!secondaryScreen) {
            pendingMirrorPosition = { left: mirrorLeft, top: mirrorTop, width: mirrorW, height: mirrorH };
            const posLabel = getSecondMonitorPosition();
            console.log('Mirror manual position: ' + posLabel + ' → x=' + mirrorLeft + ', y=' + mirrorTop + ', size ' + mirrorW + '×' + mirrorH);
            const toast = document.getElementById('mirror-placement-toast');
            if (toast) {
                const isFile = window.location.protocol === 'file:';
                toast.textContent = isFile
                    ? 'Mirror: ' + posLabel + ' (x=' + mirrorLeft + '). Window move is blocked when opening from file. Run from http://localhost — see Settings.'
                    : 'Mirror: ' + posLabel + ' (x=' + mirrorLeft + '). If it didn\'t move, drag the mirror window to your other monitor.';
                toast.classList.remove('hidden');
                setTimeout(function() { toast.classList.add('hidden'); }, 8000);
            }
        } else {
            pendingMirrorPosition = null;
        }

        const mirrorUrl = 'mirror.html';
        mirrorWindow = window.open(mirrorUrl, 'TeleprompterMirror', specs);
        if (mirrorWindow) {
            if (mirrorCloseCheckInterval) clearInterval(mirrorCloseCheckInterval);
            mirrorCloseCheckInterval = setInterval(() => {
                if (mirrorWindow && mirrorWindow.closed) {
                    restoreMainWindowFromMirrorClose();
                }
            }, 300);

            const sendWhenReady = () => {
                if (!secondaryScreen && pendingMirrorPosition) {
                    try {
                        mirrorWindow.moveTo(mirrorLeft, mirrorTop);
                        mirrorWindow.resizeTo(mirrorW, mirrorH);
                    } catch (_) {}
                }
                /* Re-enter Aa (overview) if we temporarily exited so result matches Extend then Aa */
                if (wasInOverview) {
                    broadcastEditMode = true;
                    savedFontSizeBeforeOverview = fontSizeSelect ? fontSizeSelect.value : '80';
                    if (fontSizeSelect) {
                        fontSizeSelect.value = '12';
                        lastFontChangeSource = 'size';
                    }
                    applyFontSizeOnlyToAllContent('12');
                    isOverviewMode = true;
                    document.body.classList.add('overview-mode');
                    if (btnOverview) {
                        btnOverview.classList.add('overview-active', 'broadcast-edit-active');
                        btnOverview.title = 'Hide all columns (back to control view)';
                    }
                    applyBroadcastingVisibility();
                    syncColumnWidths();
                    /* Let row heights be re-measured from 12px content so we don't keep tall rows (no huge gaps) */
                }
                syncColumnWidths();
                syncMirrorStyles(); /* Font etc. before content so mirror measures with correct styles */
                refreshMirrorData();
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

    let fontSize = mainStyle.fontSize;
    let fontFamily = mainStyle.fontFamily;
    let lineHeight = mainStyle.lineHeight;
    /* When broadcasting the last column is hidden; sample visible script column so mirror font matches what's on main. */
    const scriptCols = document.body.classList.contains('broadcasting')
        ? teleprompterText.querySelectorAll('.script-row-wrapper .script-column:not(.broadcast-hidden)')
        : teleprompterText.querySelectorAll('.script-row-wrapper .script-column:last-child');
    const visibleCols = scriptCols.length ? Array.from(scriptCols).filter(col => col.getBoundingClientRect().width > 0 && col.getBoundingClientRect().height > 0) : [];
    const colsToSample = visibleCols.length ? visibleCols : Array.from(scriptCols);
    const fontCounts = new Map();
    colsToSample.forEach(col => {
        const walker = document.createTreeWalker(col, NodeFilter.SHOW_TEXT, null, false);
        let node;
        while ((node = walker.nextNode())) {
            const len = (node.textContent || '').trim().length;
            if (len === 0) continue;
            const parent = node.parentElement;
            if (!parent) continue;
            const s = window.getComputedStyle(parent);
            const key = (s.fontFamily || '') + '|' + (s.fontSize || '');
            if (!fontCounts.has(key)) fontCounts.set(key, { count: 0, lineHeight: s.lineHeight });
            fontCounts.get(key).count += len;
        }
    });
    if (fontCounts.size > 0) {
        let bestKey = '';
        let bestCount = 0;
        fontCounts.forEach((data, key) => {
            if (data.count > bestCount) {
                bestCount = data.count;
                bestKey = key;
            }
        });
        if (bestKey) {
            const data = fontCounts.get(bestKey);
            const [fam, size] = bestKey.split('|');
            fontSize = size;
            fontFamily = fam;
            if (data.lineHeight) lineHeight = data.lineHeight;
        }
    }

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

    if (document.body.classList.contains('broadcasting')) {
        lineHeight = '1.1';
    }
    return {
        fontSize: fontSize,
        fontFamily: fontFamily,
        lineHeight: lineHeight,
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
    /* When extended keep main window at fixed 5:4 */
    if (document.body.classList.contains('broadcasting') && extendedWindowWidth != null && extendedWindowHeight != null) {
        if (window.outerWidth !== extendedWindowWidth || window.outerHeight !== extendedWindowHeight) {
            window.resizeTo(extendedWindowWidth, extendedWindowHeight);
        }
    }
    syncColumnWidths();
    syncMirrorStyles();
    if (typeof updateBookmarkPositions === 'function') updateBookmarkPositions();
    requestAnimationFrame(() => shrinkAllPillsToFit());
});
});
