document.addEventListener('DOMContentLoaded', () => {
    // --- URL del backend ---
    // 1) Si config.js define window.API_BASE_URL, se usa esa (backend en otro servidor).
    // 2) Si abres index.html desde tu PC (file://) o localhost:5500 etc., usa http://localhost:8001.
    // 3) Si la web está publicada y sirve el backend (mismo dominio), usa rutas relativas.
    const API_BASE = (typeof window.API_BASE_URL === 'string' && window.API_BASE_URL.trim() !== '')
        ? window.API_BASE_URL.trim().replace(/\/+$/, '')
        : ((location.protocol === 'file:' || location.hostname === 'localhost' || location.hostname === '127.0.0.1') && location.port !== '8001'
            ? 'http://localhost:8001'
            : '');

    // --- Navigation Logic ---
    const navHome = document.getElementById('nav-home');
    const navUnifier = document.getElementById('nav-unifier');
    const navEditor = document.getElementById('nav-editor');
    const navConverter = document.getElementById('nav-converter');
    
    const homeView = document.getElementById('home-view');
    const unifierView = document.getElementById('unifier-view');
    const editorView = document.getElementById('editor-view');
    const converterView = document.getElementById('converter-view');

    // Cards
    const cardUnifier = document.getElementById('card-unifier');
    const cardEditor = document.getElementById('card-editor');
    const cardConverter = document.getElementById('card-converter');

    function hideAllViews() {
        homeView.style.display = 'none';
        unifierView.style.display = 'none';
        editorView.style.display = 'none';
        converterView.style.display = 'none';
        
        if(navHome) navHome.classList.remove('active');
        navUnifier.classList.remove('active');
        navEditor.classList.remove('active');
        navConverter.classList.remove('active');
    }

    if(navHome) {
        navHome.addEventListener('click', () => {
            hideAllViews();
            navHome.classList.add('active');
            homeView.style.display = 'block';
        });
    }

    function showUnifier() {
        hideAllViews();
        navUnifier.classList.add('active');
        unifierView.style.display = 'block';
    }
    
    function showEditor() {
        hideAllViews();
        navEditor.classList.add('active');
        editorView.style.display = 'block';
    }

    function showConverter() {
        hideAllViews();
        navConverter.classList.add('active');
        converterView.style.display = 'block';
    }

    navUnifier.addEventListener('click', showUnifier);
    navEditor.addEventListener('click', showEditor);
    navConverter.addEventListener('click', showConverter);

    if(cardUnifier) cardUnifier.addEventListener('click', showUnifier);
    if(cardEditor) cardEditor.addEventListener('click', showEditor);
    if(cardConverter) cardConverter.addEventListener('click', showConverter);

    // --- Unificador Logic ---
    const folderInput = document.getElementById('folder-input');
    const uploadSection = document.querySelector('#unifier-view .upload-section');
    const filesSection = document.getElementById('files-section');
    const loadingSection = document.getElementById('loading-section');
    const successSection = document.getElementById('success-section');
    const fileList = document.getElementById('file-list');
    const fileCount = document.getElementById('file-count');
    const mergeBtn = document.getElementById('merge-btn');
    const resetBtn = document.getElementById('reset-btn');

    let pdfFiles = [];

    folderInput.addEventListener('change', (event) => {
        const files = Array.from(event.target.files);
        pdfFiles = files.filter(file => file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf'));

        if (pdfFiles.length === 0) {
            alert('No se encontraron archivos PDF en la carpeta seleccionada.');
            return;
        }

        pdfFiles.sort((a, b) => a.lastModified - b.lastModified);
        renderFileList();
        
        uploadSection.style.display = 'none';
        filesSection.style.display = 'block';
    });

    function renderFileList() {
        fileCount.textContent = pdfFiles.length;
        fileList.innerHTML = '';
        pdfFiles.forEach((file, index) => {
            const date = new Date(file.lastModified).toLocaleString();
            const li = document.createElement('li');
            li.className = 'file-item';
            li.innerHTML = `
                <div class="file-info">
                    <span class="file-name" title="${file.name}">${file.name}</span>
                    <span class="file-date">${date}</span>
                </div>
                <div class="file-index">${index + 1}</div>
            `;
            fileList.appendChild(li);
        });
    }

    mergeBtn.addEventListener('click', async () => {
        filesSection.style.display = 'none';
        loadingSection.style.display = 'block';

        try {
            await mergePdfs(pdfFiles);
            loadingSection.style.display = 'none';
            successSection.style.display = 'block';
        } catch (error) {
            console.error(error);
            alert('Ocurrió un error al unir los PDFs. Por favor, revisa la consola para más detalles.');
            loadingSection.style.display = 'none';
            filesSection.style.display = 'block';
        }
    });

    async function mergePdfs(files) {
        const mergedPdf = await PDFLib.PDFDocument.create();
        for (const file of files) {
            const arrayBuffer = await file.arrayBuffer();
            const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
            const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
            copiedPages.forEach((page) => {
                mergedPdf.addPage(page);
            });
        }
        const mergedPdfBytes = await mergedPdf.save();
        downloadBlob(mergedPdfBytes, 'PdfGenerado.pdf', 'application/pdf');
    }

    function downloadBlob(bytes, filename, mimeType) {
        const blob = new Blob([bytes], { type: mimeType });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);
    }

    resetBtn.addEventListener('click', () => {
        pdfFiles = [];
        folderInput.value = ''; 
        successSection.style.display = 'none';
        uploadSection.style.display = 'block';
    });

    // --- Editor de PDF Avanzado Logic ---
    const editorFileInput = document.getElementById('editor-file-input');
    const editorUploadSection = document.getElementById('editor-upload-section');
    const pdfViewerContainer = document.getElementById('pdf-viewer-container');
    const editorTools = document.getElementById('editor-tools');
    const pdfRenderCanvas = document.getElementById('pdf-render');
    const interactionLayer = document.getElementById('interaction-layer');
    const ctx = pdfRenderCanvas.getContext('2d');
    const appContainer = document.getElementById('app-container');
    
    // Tools
    const toolText = document.getElementById('tool-text');
    const toolImage = document.getElementById('tool-image');
    const toolWatermark = document.getElementById('tool-watermark');
    const toolFullscreen = document.getElementById('tool-fullscreen');
    const savePdfBtn = document.getElementById('save-pdf-btn');
    const pageNumDisplay = document.getElementById('page-num');
    const pageCountDisplay = document.getElementById('page-count');
    const prevPageBtn = document.getElementById('prev-page');
    const nextPageBtn = document.getElementById('next-page');

    // Floating Toolbar
    const floatingToolbar = document.getElementById('floating-toolbar');
    const ftFont = document.getElementById('ft-font');
    const ftBold = document.getElementById('ft-bold');
    const ftSizeDown = document.getElementById('ft-size-down');
    const ftSizeUp = document.getElementById('ft-size-up');
    const ftColor = document.getElementById('ft-color');
    const ftDelete = document.getElementById('ft-delete');

    let pdfDoc = null;
    let pageNum = 1;
    let originalPdfFile = null;
    let activeTool = null;
    let pendingActions = []; // Para cosas como marcas de agua
    let activeTextBox = null;

    editorFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) {
            originalPdfFile = file;
            editorUploadSection.style.display = 'none';
            pdfViewerContainer.style.display = 'flex';
            editorTools.style.display = 'flex';
            
            // Limpiar overlays anteriores si los hay
            interactionLayer.querySelectorAll('.pdf-overlay-box, .pdf-overlay-image').forEach(el => el.remove());
            
            const fileReader = new FileReader();
            fileReader.onload = function(e) {
                const typedarray = new Uint8Array(e.target.result);
                pdfjsLib.getDocument(typedarray).promise.then(pdf => {
                    pdfDoc = pdf;
                    pageCountDisplay.textContent = pdf.numPages;
                    renderPage(pageNum);
                }).catch(err => {
                    console.error("Error al cargar el PDF: ", err);
                    alert("No se pudo cargar el PDF.");
                });
            };
            fileReader.readAsArrayBuffer(file);
        }
    });

    function renderPage(num) {
        pdfDoc.getPage(num).then(page => {
            const scale = appContainer.classList.contains('fullscreen-active') ? 2.0 : 1.5;
            const viewport = page.getViewport({ scale: scale });
            pdfRenderCanvas.height = viewport.height;
            pdfRenderCanvas.width = viewport.width;
            
            interactionLayer.style.width = viewport.width + 'px';
            interactionLayer.style.height = viewport.height + 'px';

            const renderContext = {
                canvasContext: ctx,
                viewport: viewport
            };
            
            page.render(renderContext).promise.then(() => {
                pageNumDisplay.textContent = num;
                updateFloatingToolbarPosition();
            });
        });
    }

    prevPageBtn.addEventListener('click', () => {
        if (pageNum <= 1) return;
        pageNum--;
        interactionLayer.querySelectorAll('.pdf-overlay-box, .pdf-overlay-image').forEach(el => {
            el.style.display = parseInt(el.dataset.page) === pageNum ? 'block' : 'none';
        });
        selectTextBox(null);
        renderPage(pageNum);
    });

    nextPageBtn.addEventListener('click', () => {
        if (pageNum >= pdfDoc.numPages) return;
        pageNum++;
        interactionLayer.querySelectorAll('.pdf-overlay-box, .pdf-overlay-image').forEach(el => {
            el.style.display = parseInt(el.dataset.page) === pageNum ? 'block' : 'none';
        });
        selectTextBox(null);
        renderPage(pageNum);
    });

    toolFullscreen.addEventListener('click', () => {
        appContainer.classList.toggle('fullscreen-active');
        if(appContainer.classList.contains('fullscreen-active')) {
            toolFullscreen.textContent = '❌ Salir Pantalla Completa';
            document.body.style.overflow = 'hidden';
        } else {
            toolFullscreen.textContent = '🔲 Pantalla Completa';
            document.body.style.overflow = 'auto';
        }
        if(pdfDoc) renderPage(pageNum);
    });

    const tools = [toolText, toolImage, toolWatermark];
    
    function selectTool(toolElement, toolName) {
        tools.forEach(t => t.classList.remove('active'));
        if (activeTool === toolName) {
            activeTool = null;
            interactionLayer.classList.remove('active');
        } else {
            toolElement.classList.add('active');
            activeTool = toolName;
            interactionLayer.classList.add('active');
        }
    }

    toolText.addEventListener('click', () => selectTool(toolText, 'text'));
    
    // Image tool logic
    const imageUploadInput = document.getElementById('image-upload-input');
    toolImage.addEventListener('click', () => {
        imageUploadInput.click();
    });

    imageUploadInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if(!file) return;
        
        const reader = new FileReader();
        reader.onload = function(event) {
            const base64Src = event.target.result;
            // Comprimir imagen antes de agregarla (reduce tamaño de base64)
            const tempImg = new Image();
            tempImg.onload = function() {
                const MAX_W = 1200;
                let w = tempImg.width, h = tempImg.height;
                if (w > MAX_W) { h = Math.round(h * MAX_W / w); w = MAX_W; }
                const compCanvas = document.createElement('canvas');
                compCanvas.width = w;
                compCanvas.height = h;
                compCanvas.getContext('2d').drawImage(tempImg, 0, 0, w, h);
                const compressed = compCanvas.toDataURL('image/jpeg', 0.80);
                addImageOverlay(compressed);
            };
            tempImg.src = base64Src;
        };
        reader.readAsDataURL(file);
        
        imageUploadInput.value = ''; // reset
    });

    // Watermark tool logic
    toolWatermark.addEventListener('click', () => {
        const rect = interactionLayer.getBoundingClientRect();
        const centerX = (rect.width / 2) - 150;
        const centerY = rect.height / 2;
        
        const box = addTextOverlay(centerX, centerY);
        const content = box.querySelector('div[contenteditable]');
        content.innerText = 'CONFIDENCIAL';
        
        box.style.fontSize = '40px';
        box.style.color = '#9ca3af'; // Gris claro
        box.style.fontWeight = 'bold';
    });

    interactionLayer.addEventListener('mousedown', (e) => {
        if (e.target === interactionLayer) {
            selectTextBox(null);
        }

        if (activeTool === 'text' && e.target === interactionLayer) {
            const rect = interactionLayer.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;
            addTextOverlay(x, y);
            selectTool(toolText, 'text'); 
        }
    });

    function rgbToHex(rgb) {
        if(!rgb) return '#000000';
        if(rgb.startsWith('#')) return rgb;
        const match = rgb.match(/^rgb\((\d+),\s*(\d+),\s*(\d+)\)$/);
        if (!match) return '#000000';
        function hex(x) {
            return ("0" + parseInt(x).toString(16)).slice(-2);
        }
        return "#" + hex(match[1]) + hex(match[2]) + hex(match[3]);
    }

    function selectTextBox(box) {
        if (activeTextBox) activeTextBox.classList.remove('selected');
        activeTextBox = box;
        
        if (box) {
            box.classList.add('selected');
            floatingToolbar.style.display = 'flex';
            updateFloatingToolbarPosition();
            
            ftColor.value = rgbToHex(box.style.color || 'rgb(0,0,0)');
            if(box.style.fontWeight === 'bold') ftBold.style.background = 'rgba(255,255,255,0.3)';
            else ftBold.style.background = 'transparent';
            
            ftFont.value = box.dataset.font || 'helv';
        } else {
            floatingToolbar.style.display = 'none';
        }
    }

    function updateFloatingToolbarPosition() {
        if (!activeTextBox) return;
        const rect = activeTextBox.getBoundingClientRect();
        const containerRect = interactionLayer.getBoundingClientRect();
        
        floatingToolbar.style.left = (rect.left - containerRect.left) + 'px';
        floatingToolbar.style.top = Math.max(0, (rect.top - containerRect.top - 45)) + 'px';
    }

    function addTextOverlay(x, y) {
        const boxWrapper = document.createElement('div');
        boxWrapper.className = 'pdf-overlay-box';
        boxWrapper.style.left = x + 'px';
        boxWrapper.style.top = y + 'px';
        boxWrapper.style.fontSize = '14px';
        boxWrapper.style.color = '#000000';
        boxWrapper.style.fontWeight = 'normal';
        boxWrapper.style.fontFamily = 'Arial, Helvetica, sans-serif';
        boxWrapper.dataset.font = 'helv';
        boxWrapper.dataset.page = pageNum;

        const content = document.createElement('div');
        content.contentEditable = true;
        content.style.outline = 'none';
        content.style.minWidth = '50px';
        content.style.minHeight = '20px';
        content.innerText = 'Nuevo Texto';
        boxWrapper.appendChild(content);

        interactionLayer.appendChild(boxWrapper);
        selectTextBox(boxWrapper);

        setTimeout(() => {
            content.focus();
            document.execCommand('selectAll', false, null);
        }, 10);

        // Drag logic for the entire box
        let isDragging = false;
        let startX, startY, initialLeft, initialTop;

        boxWrapper.addEventListener('mousedown', (e) => {
            // Drag only if clicking on the padding/border, not the text directly
            // or if it's not focused yet
            if (e.target === boxWrapper || document.activeElement !== content) {
                isDragging = true;
                startX = e.clientX;
                startY = e.clientY;
                initialLeft = parseFloat(boxWrapper.style.left || 0);
                initialTop = parseFloat(boxWrapper.style.top || 0);
                selectTextBox(boxWrapper);
                e.preventDefault(); // prevent losing focus unexpectedly
            }
        });

        document.addEventListener('mousemove', (e) => {
            if (!isDragging) return;
            const dx = e.clientX - startX;
            const dy = e.clientY - startY;
            boxWrapper.style.left = (initialLeft + dx) + 'px';
            boxWrapper.style.top = (initialTop + dy) + 'px';
            updateFloatingToolbarPosition();
        });

        document.addEventListener('mouseup', () => {
            isDragging = false;
        });

        content.addEventListener('mousedown', (e) => {
            selectTextBox(boxWrapper);
            e.stopPropagation(); // don't trigger wrapper drag if clicking exactly on text
        });
        
        content.addEventListener('input', () => {
            updateFloatingToolbarPosition();
        });
        
        return boxWrapper;
    }

    function addImageOverlay(src) {
        const boxWrapper = document.createElement('div');
        boxWrapper.className = 'pdf-overlay-image';
        boxWrapper.style.left = '50px';
        boxWrapper.style.top = '50px';
        boxWrapper.style.width = '200px';
        boxWrapper.dataset.page = pageNum;

        const img = document.createElement('img');
        img.src = src;
        img.draggable = false; // Prevent native browser drag
        
        img.onload = () => {
            const ratio = img.naturalWidth / img.naturalHeight;
            boxWrapper.style.height = (200 / ratio) + 'px';
        };

        boxWrapper.appendChild(img);
        
        // Toolbar for image
        const toolbar = document.createElement('div');
        toolbar.className = 'image-toolbar';
        
        const delBtn = document.createElement('button');
        delBtn.innerText = '🗑️ Eliminar';
        delBtn.style.color = '#ef4444';
        delBtn.onclick = () => boxWrapper.remove();
        
        const bgBtn = document.createElement('button');
        bgBtn.innerText = '🪄 Quitar Fondo';
        bgBtn.onclick = () => removeWhiteBackground(img);
        
        toolbar.appendChild(bgBtn);
        toolbar.appendChild(delBtn);
        boxWrapper.appendChild(toolbar);

        // Custom resize handle
        const resizeHandle = document.createElement('div');
        resizeHandle.className = 'image-resize-handle';
        boxWrapper.appendChild(resizeHandle);

        interactionLayer.appendChild(boxWrapper);
        
        let isDragging = false;
        let isResizing = false;
        let startX, startY, initialLeft, initialTop, initialWidth, initialHeight;

        boxWrapper.addEventListener('mousedown', (e) => {
            if (e.target === toolbar || toolbar.contains(e.target)) return;
            
            if (e.target === resizeHandle) {
                isResizing = true;
            } else {
                isDragging = true;
            }
            
            startX = e.clientX;
            startY = e.clientY;
            initialLeft = parseFloat(boxWrapper.style.left || 0);
            initialTop = parseFloat(boxWrapper.style.top || 0);
            initialWidth = boxWrapper.offsetWidth;
            initialHeight = boxWrapper.offsetHeight;
            e.preventDefault(); 
        });

        document.addEventListener('mousemove', (e) => {
            if (isResizing) {
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                boxWrapper.style.width = Math.max(50, initialWidth + dx) + 'px';
                boxWrapper.style.height = Math.max(50, initialHeight + dy) + 'px';
            } else if (isDragging) {
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                boxWrapper.style.left = (initialLeft + dx) + 'px';
                boxWrapper.style.top = (initialTop + dy) + 'px';
            }
        });

        document.addEventListener('mouseup', () => {
            isDragging = false;
            isResizing = false;
        });
    }

    function removeWhiteBackground(imgElement) {
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d');
        canvas.width = imgElement.naturalWidth;
        canvas.height = imgElement.naturalHeight;
        
        ctx.drawImage(imgElement, 0, 0);
        const imageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        const data = imageData.data;
        
        // Loop through pixels
        for (let i = 0; i < data.length; i += 4) {
            const r = data[i];
            const g = data[i + 1];
            const b = data[i + 2];
            
            // If the pixel is very light (close to white), make it transparent
            if (r > 200 && g > 200 && b > 200) {
                data[i + 3] = 0; // Alpha to 0
            }
        }
        
        ctx.putImageData(imageData, 0, 0);
        imgElement.src = canvas.toDataURL('image/png');
    }

    // Floating Toolbar actions
    ftBold.addEventListener('click', () => {
        if (!activeTextBox) return;
        if (activeTextBox.style.fontWeight === 'bold') {
            activeTextBox.style.fontWeight = 'normal';
            ftBold.style.background = 'transparent';
        } else {
            activeTextBox.style.fontWeight = 'bold';
            ftBold.style.background = 'rgba(255,255,255,0.3)';
        }
    });

    ftSizeUp.addEventListener('click', () => {
        if (!activeTextBox) return;
        let size = parseInt(activeTextBox.style.fontSize) || 14;
        activeTextBox.style.fontSize = (size + 2) + 'px';
        updateFloatingToolbarPosition();
    });

    ftSizeDown.addEventListener('click', () => {
        if (!activeTextBox) return;
        let size = parseInt(activeTextBox.style.fontSize) || 14;
        activeTextBox.style.fontSize = Math.max(8, size - 2) + 'px';
        updateFloatingToolbarPosition();
    });

    ftColor.addEventListener('input', (e) => {
        if (!activeTextBox) return;
        activeTextBox.style.color = e.target.value;
    });

    ftFont.addEventListener('change', (e) => {
        if (!activeTextBox) return;
        const val = e.target.value;
        activeTextBox.dataset.font = val;
        if(val === 'helv') activeTextBox.style.fontFamily = 'Arial, Helvetica, sans-serif';
        if(val === 'ti-ro') activeTextBox.style.fontFamily = '"Times New Roman", Times, serif';
        if(val === 'cour') activeTextBox.style.fontFamily = '"Courier New", Courier, monospace';
    });

    ftDelete.addEventListener('click', () => {
        if (!activeTextBox) return;
        activeTextBox.remove();
        selectTextBox(null);
    });

    // Save to Backend
    savePdfBtn.addEventListener('click', async () => {
        if (!originalPdfFile) return;
        
        const scale = appContainer.classList.contains('fullscreen-active') ? 2.0 : 1.5;
        const boxes = interactionLayer.querySelectorAll('.pdf-overlay-box');
        
        const finalActions = [...pendingActions];
        
        boxes.forEach(box => {
            const contentDiv = box.querySelector('div[contenteditable]');
            const text = contentDiv.innerText;
            if (text.trim() === '') return;
            
            const x = parseFloat(box.style.left) / scale;
            const size = (parseInt(box.style.fontSize) || 14) / scale;
            const y = (parseFloat(box.style.top) / scale) + size; 
            
            const colorHex = rgbToHex(box.style.color);
            const isBold = box.style.fontWeight === 'bold';
            const font = box.dataset.font || 'helv';
            const elementPage = parseInt(box.dataset.page) || pageNum;

            finalActions.push({
                type: 'text',
                page: elementPage,
                text: text,
                x: x,
                y: y,
                size: size,
                color: colorHex,
                bold: isBold,
                font: font
            });
        });

        const images = interactionLayer.querySelectorAll('.pdf-overlay-image');
        images.forEach(box => {
            const img = box.querySelector('img');
            const x = parseFloat(box.style.left) / scale;
            const y = parseFloat(box.style.top) / scale;
            const w = box.offsetWidth / scale;
            const h = box.offsetHeight / scale;
            const elementPage = parseInt(box.dataset.page) || pageNum;

            finalActions.push({
                type: 'image',
                page: elementPage,
                src: img.src,
                x: x,
                y: y,
                width: w,
                height: h
            });
        });

        const formData = new FormData();
        formData.append('file', originalPdfFile);
        formData.append('actions', JSON.stringify(finalActions));

        try {
            savePdfBtn.textContent = 'Guardando...';
            savePdfBtn.disabled = true;

            const response = await fetch(API_BASE + '/api/pdf/edit', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let errMsg = `Error del servidor (${response.status})`;
                try {
                    const errJson = await response.json();
                    errMsg = errJson.detail || errMsg;
                } catch(_) {}
                throw new Error(errMsg);
            }

            const arrayBuffer = await response.arrayBuffer();
            const blob = new Blob([arrayBuffer], { type: 'application/pdf' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'PdfEditado.pdf';
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 100);
            
            alert('PDF guardado exitosamente.');
        } catch (error) {
            console.error(error);
            if (error.message && error.message.startsWith('Error del servidor')) {
                alert('Error del servidor: ' + error.message);
            } else {
                alert('Hubo un problema al procesar el PDF. Asegúrate de que el backend esté corriendo.\n\nDetalle: ' + error.message);
            }
        } finally {
            savePdfBtn.textContent = 'Guardar PDF';
            savePdfBtn.disabled = false;
        }
    });

    // --- Convertidor Logic ---
    const converterFileInput = document.getElementById('converter-file-input');
    const converterUploadSection = document.getElementById('converter-upload-section');
    const converterOptions = document.getElementById('converter-options');
    const converterFilename = document.getElementById('converter-filename');
    const convertTarget = document.getElementById('convert-target');
    const convertModeContainer = document.getElementById('convert-mode-container');
    const labelModeAll = document.getElementById('label-mode-all');
    const labelModeTables = document.getElementById('label-mode-tables');
    const changeConvertFileBtn = document.getElementById('change-convert-file-btn');
    const startConvertBtn = document.getElementById('start-convert-btn');
    const converterLoading = document.getElementById('converter-loading');
    
    let converterFile = null;
    const supportedImages = ['png', 'jpg', 'jpeg', 'webp', 'bmp', 'tiff', 'tif'];

    function updateModeUI() {
        const selectedRadio = document.querySelector('input[name="conversion-mode"]:checked');
        const modeVal = selectedRadio ? selectedRadio.value : 'all';
        if (labelModeAll && labelModeTables) {
            labelModeAll.classList.toggle('active', modeVal === 'all');
            labelModeTables.classList.toggle('active', modeVal === 'tables');
        }
        // Mostrar selector de modo cuando la salida sea Excel, Word o PowerPoint
        if (convertModeContainer) {
            const target = convertTarget.value;
            if (target === 'xlsx' || target === 'docx' || target === 'pptx') {
                convertModeContainer.style.display = 'flex';
            } else {
                convertModeContainer.style.display = 'none';
            }
        }
    }

    document.querySelectorAll('input[name="conversion-mode"]').forEach(radio => {
        radio.addEventListener('change', updateModeUI);
    });

    if (convertTarget) {
        convertTarget.addEventListener('change', updateModeUI);
    }

    if (changeConvertFileBtn) {
        changeConvertFileBtn.addEventListener('click', () => {
            converterFile = null;
            converterFileInput.value = '';
            converterOptions.style.display = 'none';
            converterUploadSection.style.display = 'block';
        });
    }

    converterFileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) {
            converterFile = file;
            converterFilename.textContent = file.name;
            
            convertTarget.innerHTML = '';
            const ext = file.name.split('.').pop().toLowerCase();
            
            if (ext === 'pdf') {
                convertTarget.innerHTML = `
                    <option value="xlsx">Excel (.xlsx)</option>
                    <option value="docx">Word (.docx)</option>
                    <option value="pptx">PowerPoint (.pptx)</option>
                `;
            } else if (supportedImages.includes(ext)) {
                convertTarget.innerHTML = `
                    <option value="xlsx">Excel (.xlsx)</option>
                    <option value="docx">Word (.docx)</option>
                    <option value="pptx">PowerPoint (.pptx)</option>
                    <option value="pdf">PDF (.pdf)</option>
                `;
            } else if (ext === 'docx' || ext === 'pptx') {
                convertTarget.innerHTML = `
                    <option value="pdf">PDF (.pdf)</option>
                `;
            } else {
                alert('Formato no soportado. Selecciona un archivo PDF, Imagen (PNG, JPG, WEBP, BMP, TIFF), DOCX o PPTX.');
                converterFileInput.value = '';
                return;
            }
            
            updateModeUI();
            converterUploadSection.style.display = 'none';
            converterOptions.style.display = 'block';
        }
    });

    startConvertBtn.addEventListener('click', async () => {
        if (!converterFile) return;
        const targetFormat = convertTarget.value;
        const selectedRadio = document.querySelector('input[name="conversion-mode"]:checked');
        const conversionMode = selectedRadio ? selectedRadio.value : 'all';
        
        converterOptions.style.display = 'none';
        converterLoading.style.display = 'block';
        
        const formData = new FormData();
        formData.append('file', converterFile);
        formData.append('target_format', targetFormat);
        formData.append('conversion_mode', conversionMode);
        
        try {
            const response = await fetch(API_BASE + '/api/convert', {
                method: 'POST',
                body: formData
            });

            if (!response.ok) {
                let detail = 'Error en la conversión';
                try {
                    const err = await response.json();
                    detail = err.detail || detail;
                } catch (_) {}
                throw new Error(detail);
            }

            const arrayBuffer = await response.arrayBuffer();
            const mimeType = targetFormat === 'pdf' ? 'application/pdf' : 
                             targetFormat === 'docx' ? 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' :
                             targetFormat === 'pptx' ? 'application/vnd.openxmlformats-officedocument.presentationml.presentation' :
                             targetFormat === 'xlsx' ? 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' : '';
            const blob = new Blob([arrayBuffer], { type: mimeType });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            
            const dotIdx = converterFile.name.lastIndexOf('.');
            const originalName = dotIdx !== -1 ? converterFile.name.substring(0, dotIdx) : converterFile.name;
            const suffix = (conversionMode === 'tables' && targetFormat !== 'pdf') ? 'tablas' : 'convertido';
            a.download = `${originalName}_${suffix}.${targetFormat}`;
            
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
            }, 100);
            
            alert('¡Conversión exitosa!');
            converterLoading.style.display = 'none';
            converterUploadSection.style.display = 'block';
            converterFile = null;
            converterFileInput.value = '';
        } catch (error) {
            console.error(error);
            alert('Hubo un problema: ' + error.message);
            converterLoading.style.display = 'none';
            converterOptions.style.display = 'block';
        }
    });

    // --- Drag and Drop Global Helper ---
    function setupDragAndDrop(dropZoneId, inputId) {
        const dropZone = document.getElementById(dropZoneId);
        const input = document.getElementById(inputId);
        if(!dropZone || !input) return;

        ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, preventDefaults, false);
        });

        function preventDefaults(e) {
            e.preventDefault();
            e.stopPropagation();
        }

        ['dragenter', 'dragover'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.style.background = 'rgba(139, 92, 246, 0.2)', false);
        });

        ['dragleave', 'drop'].forEach(eventName => {
            dropZone.addEventListener(eventName, () => dropZone.style.background = '', false);
        });

        dropZone.addEventListener('drop', (e) => {
            const dt = e.dataTransfer;
            const files = dt.files;
            
            if(files && files.length > 0) {
                input.files = files;
                input.dispatchEvent(new Event('change'));
            }
        }, false);
        
        // Evitar que el navegador abra el archivo si se suelta fuera de la zona
        window.addEventListener('dragover', preventDefaults, false);
        window.addEventListener('drop', preventDefaults, false);
    }

    setupDragAndDrop('drop-zone', 'folder-input');
    setupDragAndDrop('editor-drop-zone', 'editor-file-input');
    setupDragAndDrop('converter-drop-zone', 'converter-file-input');

});
