// ===============================================
//               ELEMENT SELECTORS
// ===============================================
const newFileButton = document.getElementById('btn-new-script');
const openFileButton = document.getElementById('btn-open-file');
const toggleRunlistButton = document.getElementById('btn-toggle-runlist');
const fileOpener = document.getElementById('file-opener');
const runlistContainer = document.querySelector('.runlist-files');
const teleprompterText = document.getElementById('teleprompter-text');
const runlistPanel = document.getElementById('runlist-panel');
const resizer = document.getElementById('resizer');
const teleprompterView = document.getElementById('teleprompter-view');
// ===============================================
//                APP STATE
// ===============================================
let fileStore = [];
let contentStore = [];
let currentFileIndex = -1;
// A variable to track the visibility state
let isRunlistVisible = false;
// ===============================================
//              UI STATE FUNCTIONS
// ===============================================
function showRunlist() {
    runlistPanel.style.display = 'flex';
    resizer.style.display = 'block';
    isRunlistVisible = true;
}
function hideRunlist() {
    runlistPanel.style.display = 'none';
    resizer.style.display = 'none';
    isRunlistVisible = false;
}
// ===============================================
// BUTTON & EVENT LISTENERS
// ===============================================
newFileButton.addEventListener('click', () => {
const newFileIndex = fileStore.push(null) - 1;
contentStore.push('');
addFileToRunlist("new_file", newFileIndex);
if (currentFileIndex !== -1) {
contentStore[currentFileIndex] = teleprompterText.innerHTML;
}
currentFileIndex = newFileIndex;
teleprompterText.innerHTML = '';
showRunlist();
});
openFileButton.addEventListener('click', () => {
fileOpener.accept = ".rtf,.doc,.docx,.xls,.xlsx";
fileOpener.click();
});
fileOpener.addEventListener('change', (event) => {
const file = event.target.files[0];
if (!file) { return; }
const newFileIndex = fileStore.push(file) - 1;
contentStore.push(null);
addFileToRunlist(file.name, newFileIndex);
if (currentFileIndex !== -1) {
contentStore[currentFileIndex] = teleprompterText.innerHTML;
}
currentFileIndex = newFileIndex;
loadFileContent(file, newFileIndex);
fileOpener.value = '';
});
runlistContainer.addEventListener('click', (event) => {
const clickedItem = event.target.closest('.runlist-item');
if (clickedItem) {
const fileIndex = parseInt(clickedItem.dataset.index);
if (currentFileIndex !== -1 && currentFileIndex !== fileIndex) {
contentStore[currentFileIndex] = teleprompterText.innerHTML;
}
currentFileIndex = fileIndex;
if (contentStore[fileIndex] !== null && contentStore[fileIndex] !== undefined) {
teleprompterText.innerHTML = contentStore[fileIndex];
} else {
const file = fileStore[fileIndex];
if (file) {
loadFileContent(file, fileIndex);
} else {
teleprompterText.innerHTML = '';
contentStore[fileIndex] = '';
}
}
}
});
toggleRunlistButton.addEventListener('click', () => {
    console.log('Toggle clicked');
    if (isRunlistVisible) {
        hideRunlist();
    } else {
        showRunlist();
    }
});
// ===============================================
// HELPER FUNCTIONS
// ===============================================
function loadFileContent(file, index) {
const reader = new FileReader();
reader.onload = function(e) {
const fileContent = e.target.result;
const uuidRegex = /[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}/gi;
if (file.name.endsWith('.docx')) {
mammoth.convertToHtml({ arrayBuffer: fileContent })
.then(result => {
teleprompterText.innerHTML = result.value.replace(uuidRegex, '');
contentStore[index] = teleprompterText.innerHTML;
})
.catch(error => {
console.error("Error processing DOCX file:", error);
teleprompterText.innerHTML = '<p style="color:red;">Could not read the DOCX file.</p>';
contentStore[index] = teleprompterText.innerHTML;
});
} else if (file.name.endsWith('.xlsx') || file.name.endsWith('.xls')) {
try {
const workbook = XLSX.read(fileContent, { type: 'buffer' });
const firstSheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[firstSheetName];
const html = XLSX.utils.sheet_to_html(worksheet);
teleprompterText.innerHTML = html.replace(uuidRegex, '');
contentStore[index] = teleprompterText.innerHTML;
} catch (error) {
console.error("Error processing Excel file:", error);
teleprompterText.innerHTML = '<p style="color:red;">Could not read the DOCX file.</p>';
contentStore[index] = teleprompterText.innerHTML;
}
}
};
reader.readAsArrayBuffer(file);
}
function addFileToRunlist(fileName, index) {
const newItem = document.createElement('div');
newItem.classList.add('runlist-item');
newItem.textContent = fileName;
newItem.dataset.index = index;
runlistContainer.appendChild(newItem);
showRunlist();
}
// ===============================================
//           COLUMN RESIZING LOGIC
// ===============================================
const onMouseDown = (e) => {
    e.preventDefault();
    let x = e.clientX;
    let panelWidth = runlistPanel.getBoundingClientRect().width;
    const onMouseMove = (e) => {
        const dx = e.clientX - x;
        runlistPanel.style.width = `${panelWidth - dx}px`;
    };
    const onMouseUp = () => {
        document.removeEventListener('mousemove', onMouseMove);
        document.removeEventListener('mouseup', onMouseUp);
    };
    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
};
resizer.addEventListener('mousedown', onMouseDown);
// ===============================================
//           INITIALIZE APP STATE
// ===============================================
// Set the initial state of the UI when the script loads
hideRunlist();