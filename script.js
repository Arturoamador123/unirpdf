document.addEventListener('DOMContentLoaded', () => {
    const folderInput = document.getElementById('folder-input');
    const uploadSection = document.querySelector('.upload-section');
    const filesSection = document.getElementById('files-section');
    const loadingSection = document.getElementById('loading-section');
    const successSection = document.getElementById('success-section');
    const fileList = document.getElementById('file-list');
    const fileCount = document.getElementById('file-count');
    const mergeBtn = document.getElementById('merge-btn');
    const resetBtn = document.getElementById('reset-btn');

    let pdfFiles = [];

    // Manejar selección de carpeta
    folderInput.addEventListener('change', (event) => {
        const files = Array.from(event.target.files);
        
        // Filtrar solo archivos PDF
        pdfFiles = files.filter(file => file.type === 'application/pdf' || file.name.toLowerCase().endsWith('.pdf'));

        if (pdfFiles.length === 0) {
            alert('No se encontraron archivos PDF en la carpeta seleccionada.');
            return;
        }

        // Ordenar por fecha de última modificación (ascendente: más antiguos primero)
        // Puedes cambiar a b.lastModified - a.lastModified para descendente
        pdfFiles.sort((a, b) => a.lastModified - b.lastModified);

        renderFileList();
        
        uploadSection.style.display = 'none';
        filesSection.style.display = 'block';
    });

    // Renderizar la lista de archivos en la UI
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

    // Iniciar fusión de PDFs
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

    // Función principal para unir los PDFs usando pdf-lib
    async function mergePdfs(files) {
        // Crear un nuevo documento PDF vacío
        const mergedPdf = await PDFLib.PDFDocument.create();

        for (const file of files) {
            // Leer el archivo como ArrayBuffer
            const arrayBuffer = await file.arrayBuffer();
            
            // Cargar el documento PDF
            const pdfDoc = await PDFLib.PDFDocument.load(arrayBuffer);
            
            // Obtener todas las páginas del PDF actual
            const copiedPages = await mergedPdf.copyPages(pdfDoc, pdfDoc.getPageIndices());
            
            // Añadir las páginas copiadas al documento principal
            copiedPages.forEach((page) => {
                mergedPdf.addPage(page);
            });
        }

        // Serializar el PDF final a bytes
        const mergedPdfBytes = await mergedPdf.save();
        
        // Descargar el archivo
        downloadBlob(mergedPdfBytes, 'Merged_Document.pdf', 'application/pdf');
    }

    // Utilidad para descargar el Blob generado
    function downloadBlob(bytes, filename, mimeType) {
        const blob = new Blob([bytes], { type: mimeType });
        const url = URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        
        // Limpiar
        setTimeout(() => {
            document.body.removeChild(a);
            URL.revokeObjectURL(url);
        }, 100);
    }

    // Reiniciar la UI para otra operación
    resetBtn.addEventListener('click', () => {
        pdfFiles = [];
        folderInput.value = ''; // Limpiar input
        successSection.style.display = 'none';
        uploadSection.style.display = 'block';
    });
});
