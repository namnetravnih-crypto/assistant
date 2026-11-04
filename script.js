// --- Основные переменные ---
const fileInput = document.getElementById('fileInput');
const uploadLabel = document.getElementById('uploadLabel');
const dropZone = document.getElementById('dropZone');
const canvasContainer = document.getElementById('canvasContainer');
const canvas = document.getElementById('mainCanvas');
const ctx = canvas.getContext('2d');
const statusBar = document.getElementById('statusBar');
const fileNameSpan = document.getElementById('fileName');
const brushCursor = document.getElementById('brushCursor');

const btnMagicCursor = document.getElementById('btnMagicCursor');
const btnEraser = document.getElementById('btnEraser');
const btnUndo = document.getElementById('btnUndo');
const btnMakeRound = document.getElementById('btnMakeRound');
const btnDownload = document.getElementById('btnDownload');
const btnReset = document.getElementById('btnReset');

const removeTextCheck = document.getElementById('removeTextCheck');
const paramBrightness = document.getElementById('paramBrightness');
const paramChroma = document.getElementById('paramChroma');
const paramShadows = document.getElementById('paramShadows');
const paramSmoothing = document.getElementById('paramSmoothing');
const brushSize = document.getElementById('brushSize');

const val1 = document.getElementById('val1');
const val2 = document.getElementById('val2');
const val3 = document.getElementById('val3');
const val4 = document.getElementById('val4');
const valBrushSize = document.getElementById('valBrushSize');

// --- Состояние приложения ---
let originalImageData = null;
let workingImageData = null;
let historyStack = [];
let isDrawing = false;
let currentTool = 'brush'; // brush или eraser
let brushRadius = 30;

// --- Обновление значений ползунков ---
val1.textContent = paramBrightness.value;
val2.textContent = paramChroma.value;
val3.textContent = paramShadows.value;
val4.textContent = paramSmoothing.value;
valBrushSize.textContent = `${brushSize.value} px`;

paramBrightness.addEventListener('input', (e) => { val1.textContent = e.target.value; applyFilters(); });
paramChroma.addEventListener('input', (e) => { val2.textContent = e.target.value; applyFilters(); });
paramShadows.addEventListener('input', (e) => { val3.textContent = e.target.value; applyFilters(); });
paramSmoothing.addEventListener('input', (e) => { val4.textContent = e.target.value; applyFilters(); });
brushSize.addEventListener('input', (e) => { 
    valBrushSize.textContent = `${e.target.value} px`; 
    brushRadius = e.target.value; 
});

// --- Загрузка файла (кнопка и drag-and-drop) ---
const loadFile = (file) => {
    if (!file.type.startsWith('image/')) {
        alert('Пожалуйста, загрузите изображение.');
        return;
    }
    
    saveState(); // Сохраняем чистый холст в историю
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const img = new Image();
        img.onload = () => {
            // Настраиваем размер холста под изображение
            canvas.width = img.width;
            canvas.height = img.height;
            canvasContainer.style.display = 'block';
            dropZone.style.display = 'none';
            
            ctx.drawImage(img, 0, 0);
            workingImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            originalImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
            
            fileNameSpan.textContent = file.name;
            statusBar.style.color = 'var(--text-muted)';
            applyFilters();
        };
        img.src = e.target.result;
    };
    reader.readAsDataURL(file);
};

fileInput.addEventListener('change', (e) => {
    if (e.target.files[0]) {
        loadFile(e.target.files[0]);
    }
});

uploadLabel.addEventListener('dragover', (e) => e.preventDefault());
uploadLabel.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});

['dragenter', 'dragover'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.add('dragover');
    });
});

['dragleave', 'drop'].forEach(evt => {
    dropZone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropZone.classList.remove('dragover');
    });
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    if (e.dataTransfer.files[0]) loadFile(e.dataTransfer.files[0]);
});

// --- Инструменты (Ластик и Маска) ---
const getBrushCoords = (e) => {
    const rect = canvas.getBoundingClientRect();
    return {
        x: e.clientX - rect.left,
        y: e.clientY - rect.top
    };
};

canvas.addEventListener('mousedown', (e) => {
    isDrawing = true;
    const { x, y } = getBrushCoords(e);
    if (currentTool === 'eraser') {
        eraseAt(x, y);
    }
    // Для маски-круга клик просто ставит точку (полная реализация круга требует выделения области)
});

canvas.addEventListener('mousemove', (e) => {
    const { x, y } = getBrushCoords(e);
    brushCursor.style.left = `${e.clientX}px`;
    brushCursor.style.top = `${e.clientY}px`;
    brushCursor.style.width = `${brushRadius * 2}px`;
    brushCursor.style.height = `${brushRadius * 2}px`;
    
    if (isDrawing && currentTool === 'eraser') {
        eraseAt(x, y);
    }
});

['mouseup', 'mouseleave'].forEach(evt => {
    canvas.addEventListener(evt, () => {
        isDrawing = false;
    });
});

btnMagicCursor.addEventListener('click', () => {
    currentTool = 'brush';
    btnMagicCursor.classList.add('tool-active');
    btnEraser.classList.remove('tool-active');
});

btnEraser.addEventListener('click', () => {
    currentTool = 'eraser';
    btnEraser.classList.add('tool-active');
    btnMagicCursor.classList.remove('tool-active');
});

const eraseAt = (centerX, centerY) => {
    if (!workingImageData) return;
    const data = workingImageData.data;
    const radius = brushRadius;
    const radiusSq = radius * radius;
    
    for (let y = centerY - radius; y <= centerY + radius; y++) {
        for (let x = centerX - radius; x <= centerX + radius; x++) {
            const dx = x - centerX;
            const dy = y - centerY;
            if (dx * dx + dy * dy <= radiusSq && x >= 0 && x < canvas.width && y >= 0 && y < canvas.height) {
                const i = (y * canvas.width + x) * 4;
                // Делаем пиксель прозрачным
                data[i + 3] = 0; 
            }
        }
    }
    ctx.putImageData(workingImageData, 0, 0);
};

// --- Фильтрация (упрощенная симуляция) ---
const applyFilters = () => {
    if (!originalImageData) return;
    
    // Копируем оригинал, чтобы не накладывать фильтры друг на друга
    workingImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    const data = workingImageData.data;
    
    // Простейший фильтр: делаем всё, что светлее порога, белым (прозрачным), а остальное — черным.
    // В реальном проекте здесь должен быть сложный алгоритм (например, K-means или GrabCut).
    const threshold = parseInt(paramBrightness.value);
    
    for (let i = 0; i < data.length; i += 4) {
        const r = data[i];
        const g = data[i+1];
        const b = data[i+2];
        const brightness = (r + g + b) / 3;
        
        if (brightness > threshold) {
            // Белый фон -> делаем прозрачным
            data[i + 3] = 0; 
        } else {
            // Черный оттиск -> делаем чисто черным
            data[i] = 0;
            data[i+1] = 0;
            data[i+2] = 0;
        }
    }
    
    // Сглаживание (упрощенное, через повторный проход или Canvas фильтры)
    // В данном примере сглаживание реализовано визуально через state, 
    // но для реального эффекта нужны Stack Blur или Stackblur.js
    
    ctx.putImageData(workingImageData, 0, 0);
};

// --- Управление историей (Undo) ---
const saveState = () => {
    if (workingImageData) {
        historyStack.push(ctx.getImageData(0, 0, canvas.width, canvas.height));
        if (historyStack.length > 20) historyStack.shift(); // Ограничиваем память
    }
};

btnUndo.addEventListener('click', () => {
    if (historyStack.length > 0) {
        const lastState = historyStack.pop();
        ctx.putImageData(lastState, 0, 0);
        workingImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
    }
});

// --- Геометрия (Круг) ---
btnMakeRound.addEventListener('click', () => {
    if (!workingImageData) return;
    saveState();
    
    const tempCanvas = document.createElement('canvas');
    tempCanvas.width = canvas.width;
    tempCanvas.height = canvas.height;
    const tempCtx = tempCanvas.getContext('2d');
    
    tempCtx.drawImage(canvas, 0, 0);
    
    // Создаем круглый путь
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.save();
    
    // Находим центр и радиус (по меньшей стороне)
    const size = Math.min(canvas.width, canvas.height);
    const centerX = canvas.width / 2;
    const centerY = canvas.height / 2;
    
    ctx.beginPath();
    ctx.arc(centerX, centerY, size / 2, 0, Math.PI * 2);
    ctx.closePath();
    ctx.clip();
    
    // Рисуем изображение внутри круга
    ctx.drawImage(tempCanvas, 0, 0);
    ctx.restore();
    
    workingImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
});

// --- Экспорт ---
btnDownload.addEventListener('click', () => {
    if (!workingImageData) return;
    
    // Создаем временный холст для экспорта 40мм при 300dpi
    // 40 мм = 1.575 дюйма. 1.575 * 300dpi = ~472px
    const exportSize = 472; 
    const exportCanvas = document.createElement('canvas');
    exportCanvas.width = exportSize;
    exportCanvas.height = exportSize;
    const exportCtx = exportCanvas.getContext('2d');
    
    // Сохраняем пропорции
    const scale = exportSize / Math.max(canvas.width, canvas.height);
    const drawW = canvas.width * scale;
    const drawH = canvas.height * scale;
    const offsetX = (exportSize - drawW) / 2;
    const offsetY = (exportSize - drawH) / 2;
    
    exportCtx.drawImage(canvas, offsetX, offsetY, drawW, drawH);
    
    const link = document.createElement('a');
    link.download = 'stamp-cleaned.png';
    link.href = exportCanvas.toDataURL('image/png', 1.0);
    link.click();
});

// --- Сброс ---
btnReset.addEventListener('click', () => {
    if (originalImageData) {
        saveState();
        ctx.putImageData(originalImageData, 0, 0);
        workingImageData = ctx.getImageData(0, 0, canvas.width, canvas.height);
        historyStack = [];
    } else {
        canvasContainer.style.display = 'none';
        dropZone.style.display = 'flex';
        fileNameSpan.textContent = '';
        statusBar.style.color = 'var(--text-muted)';
    }
});

// --- Курсор ---
canvas.addEventListener('mouseenter', () => {
    if (workingImageData) brushCursor.style.display = 'block';
});
canvas.addEventListener('mouseleave', () => {
    brushCursor.style.display = 'none';
    isDrawing = false;
});

// Инициализация инструментов
btnMagicCursor.classList.add('tool-active');