(function () {
    'use strict';

    const $ = (id) => document.getElementById(id);

    // ── Form Elements ───────────────────────────────────────────────────
    const labelStyle = $('labelStyle');
    const signatureType = $('signatureType');
    const trackingNumber = $('trackingNumber');
    const sortCode = $('sortCode');
    const routingCode = $('routingCode');
    const reference = $('reference');
    const parcelType = $('parcelType');
    const weight = $('weight');
    const postage = $('postage');
    const postByDate = $('postByDate');
    const recipientName = $('recipientName');
    const recipientAddr1 = $('recipientAddr1');
    const recipientAddr2 = $('recipientAddr2');
    const recipientCity = $('recipientCity');
    const recipientPostcode = $('recipientPostcode');
    const senderName = $('senderName');
    const senderAddr1 = $('senderAddr1');
    const senderCity = $('senderCity');
    const senderPostcode = $('senderPostcode');
    const sellerType = $('sellerType');
    const printedFrom = $('printedFrom');
    const customPrintedFrom = $('customPrintedFrom');
    const customPrintedFromGroup = $('customPrintedFromGroup');

    const generateBtn = $('generateBtn');
    const clearBtn = $('clearBtn');
    const downloadPdfBtn = $('downloadPdfBtn');
    const downloadPngBtn = $('downloadPngBtn');
    const printBtn = $('printBtn');
    const labelPreview = $('labelPreview');
    const saveTemplateBtn = $('saveTemplateBtn');
    const templateNameInput = $('templateName');
    const templateList = $('templateList');

    const fields = [
        'labelStyle', 'signatureType', 'trackingNumber', 'sortCode', 'routingCode', 'reference',
        'parcelType', 'weight', 'postage', 'postByDate',
        'recipientName', 'recipientAddr1', 'recipientAddr2', 'recipientCity', 'recipientPostcode',
        'senderName', 'senderAddr1', 'senderCity', 'senderPostcode',
        'sellerType', 'printedFrom', 'customPrintedFrom'
    ];

    // ── Royal Mail Crown Logo SVG ───────────────────────────────────────
    const CROWN_SVG = `
        <g transform="translate(12, 3) scale(0.85)">
            <path d="M2 22 L6 10 L13 16 L25 5 L37 16 L44 10 L48 22 Z" fill="#000"/>
            <circle cx="6" cy="7" r="3" fill="#000"/>
            <circle cx="25" cy="2" r="3.5" fill="#000"/>
            <circle cx="44" cy="7" r="3" fill="#000"/>
            <rect x="23" y="0" width="4" height="5" fill="#000"/>
        </g>
        <text x="50" y="42" text-anchor="middle" font-family="Arial, sans-serif" font-size="13" font-weight="bold" fill="#000">Royal Mail</text>
    `;

    // ── Event Handlers ──────────────────────────────────────────────────
    printedFrom.addEventListener('change', () => {
        customPrintedFromGroup.style.display = printedFrom.value === 'Custom' ? 'block' : 'none';
    });

    // ── Helper Functions ────────────────────────────────────────────────
    function esc(str) {
        return (str || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
    }

    function removeSpaces(str) {
        return (str || '').replace(/\s+/g, '');
    }

    function formatTrackingDisplay(tracking) {
        const clean = removeSpaces(tracking).toUpperCase();
        if (clean.length >= 13) {
            return `${clean.slice(0, 2)} ${clean.slice(2, 6)} ${clean.slice(6, 10)} ${clean.slice(10)}`;
        }
        return clean;
    }

    function parsePostage(val) {
        const match = (val || '').replace(/[£,]/g, '').match(/[\d.]+/);
        return match ? parseFloat(match[0]) : 0;
    }

    function formatPostageForDM(val) {
        const pence = Math.round(parsePostage(val) * 100);
        return pence.toString().padStart(5, '0');
    }

    function formatRefForDM(ref) {
        return removeSpaces((ref || '').replace(/-/g, '')).toUpperCase();
    }

    function genDateCode() {
        const now = new Date();
        const dd = now.getDate().toString().padStart(2, '0');
        const mm = (now.getMonth() + 1).toString().padStart(2, '0');
        const yy = now.getFullYear().toString().slice(-2);
        const seq = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        return `${dd}${mm}${yy}${seq}`;
    }

    // ── DataMatrix Content ──────────────────────────────────────────────
    function buildDMContent(data) {
        const accountId = '8215FA';
        const refCode = formatRefForDM(data.reference);
        const serviceCode = data.signatureType === 'signature' ? '00010001' : '00020001';
        const postageFmt = formatPostageForDM(data.postage);
        const dateCode = genDateCode();
        const trackingClean = removeSpaces(data.tracking).toUpperCase();
        const destPC = removeSpaces(data.recipientPostcode).toUpperCase();
        const returnPC = removeSpaces(data.senderPostcode).toUpperCase();
        let routeCode = removeSpaces(data.routingCode).toUpperCase() || destPC.slice(0, 3);

        return `JGB${accountId}${refCode}${serviceCode}${postageFmt}${dateCode}062${trackingClean}${routeCode.slice(0,3)}${destPC}GB${returnPC}`;
    }

    // ── Barcode Generation ──────────────────────────────────────────────
    async function genDataMatrix(content, canvasId) {
        try {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            bwipjs.toCanvas(canvas, {
                bcid: 'datamatrix',
                text: content,
                scale: 4,
                padding: 0,
            });
        } catch (e) {
            console.error('DataMatrix error:', e);
        }
    }

    async function genCode128(content, canvasId) {
        try {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;
            bwipjs.toCanvas(canvas, {
                bcid: 'code128',
                text: content,
                scale: 2,
                height: 20,
                includetext: false,
            });
        } catch (e) {
            console.error('Code128 error:', e);
        }
    }

    // ── Form Data ───────────────────────────────────────────────────────
    function getFormData() {
        return {
            labelStyle: labelStyle.value,
            signatureType: signatureType.value,
            tracking: trackingNumber.value.trim(),
            sortCode: sortCode.value.trim(),
            routingCode: routingCode.value.trim(),
            reference: reference.value.trim(),
            parcelType: parcelType.value,
            weight: weight.value.trim(),
            postage: postage.value.trim(),
            postByDate: postByDate.value.trim(),
            recipientName: recipientName.value.trim(),
            recipientAddr1: recipientAddr1.value.trim(),
            recipientAddr2: recipientAddr2.value.trim(),
            recipientCity: recipientCity.value.trim(),
            recipientPostcode: recipientPostcode.value.trim(),
            senderName: senderName.value.trim(),
            senderAddr1: senderAddr1.value.trim(),
            senderCity: senderCity.value.trim(),
            senderPostcode: senderPostcode.value.trim(),
            sellerType: sellerType.value.trim(),
            printedFrom: printedFrom.value === 'Custom' ? customPrintedFrom.value.trim() : printedFrom.value,
        };
    }

    function setFormData(data) {
        if (data.labelStyle) labelStyle.value = data.labelStyle;
        if (data.signatureType) signatureType.value = data.signatureType;
        if (data.tracking) trackingNumber.value = data.tracking;
        if (data.sortCode) sortCode.value = data.sortCode;
        if (data.routingCode) routingCode.value = data.routingCode;
        if (data.reference) reference.value = data.reference;
        if (data.parcelType) parcelType.value = data.parcelType;
        if (data.weight) weight.value = data.weight;
        if (data.postage) postage.value = data.postage;
        if (data.postByDate) postByDate.value = data.postByDate;
        if (data.recipientName) recipientName.value = data.recipientName;
        if (data.recipientAddr1) recipientAddr1.value = data.recipientAddr1;
        if (data.recipientAddr2) recipientAddr2.value = data.recipientAddr2;
        if (data.recipientCity) recipientCity.value = data.recipientCity;
        if (data.recipientPostcode) recipientPostcode.value = data.recipientPostcode;
        if (data.senderName) senderName.value = data.senderName;
        if (data.senderAddr1) senderAddr1.value = data.senderAddr1;
        if (data.senderCity) senderCity.value = data.senderCity;
        if (data.senderPostcode) senderPostcode.value = data.senderPostcode;
        if (data.sellerType) sellerType.value = data.sellerType;
        if (data.printedFrom) {
            if (['Click & Drop', 'eBay Simple Delivery', 'Parcel2Go', 'ShipStation'].includes(data.printedFrom)) {
                printedFrom.value = data.printedFrom;
                customPrintedFromGroup.style.display = 'none';
            } else {
                printedFrom.value = 'Custom';
                customPrintedFrom.value = data.printedFrom;
                customPrintedFromGroup.style.display = 'block';
            }
        }
    }

    // ── Generate Label ──────────────────────────────────────────────────
    generateBtn.addEventListener('click', async () => {
        const data = getFormData();

        if (!data.tracking) { alert('Enter tracking number.'); return; }
        if (!data.recipientPostcode) { alert('Enter recipient postcode.'); return; }

        const isClickDrop = data.labelStyle === 'clickdrop';
        const sigText = data.signatureType === 'signature' ? '' : 'No Signature';
        const trackingDisplay = formatTrackingDisplay(data.tracking);
        const trackingClean = removeSpaces(data.tracking).toUpperCase();
        const postageNum = parsePostage(data.postage);
        const postageDisplay = postageNum > 0 ? `£${postageNum.toFixed(2)}` : '';

        // Return address lines
        const returnLines = ['Return Address', data.senderName, data.senderAddr1, data.senderCity, data.senderPostcode.toUpperCase()].filter(Boolean);

        // Build SVG label - exact Royal Mail template
        const svgLabel = `
<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 600" class="rm-svg-label">
    <defs>
        <style>
            .rm-text { font-family: Arial, Helvetica, sans-serif; fill: #000; }
            .rm-bold { font-weight: 700; }
            .rm-black { font-weight: 900; }
            .rm-italic { font-style: italic; }
            .rm-small { font-size: 10px; }
            .rm-tiny { font-size: 8px; }
        </style>
    </defs>

    <!-- Background -->
    <rect width="400" height="600" fill="#fff"/>
    <rect x="0.5" y="0.5" width="399" height="599" fill="none" stroke="#000" stroke-width="1"/>

    <!-- Row 1: Header -->
    <line x1="0" y1="85" x2="400" y2="85" stroke="#000" stroke-width="1"/>

    <!-- Tracked text -->
    <text x="15" y="45" class="rm-text rm-black rm-italic" font-size="38">Tracked</text>
    <text x="15" y="70" class="rm-text" font-size="18">${esc(sigText)}</text>

    <!-- 24 -->
    <text x="200" y="65" class="rm-text rm-black" font-size="72" text-anchor="middle">24</text>

    <!-- Right side - Delivered By + Logo -->
    <text x="385" y="18" class="rm-text rm-small" text-anchor="end">Delivered By</text>
    <rect x="295" y="22" width="90" height="48" rx="3" fill="none" stroke="#000" stroke-width="1.5"/>
    <svg x="295" y="22" width="90" height="48" viewBox="0 0 100 55">
        ${CROWN_SVG}
    </svg>
    <text x="385" y="82" class="rm-text rm-tiny" text-anchor="end">Postage Paid GB</text>

    <!-- Row 2: Routing codes -->
    <line x1="0" y1="${isClickDrop ? 140 : 140}" x2="400" y2="${isClickDrop ? 140 : 140}" stroke="#000" stroke-width="1"/>

    <!-- Sort code boxes -->
    <rect x="15" y="95" width="70" height="38" fill="#000"/>
    <text x="50" y="122" class="rm-text rm-black" font-size="28" fill="#fff" text-anchor="middle">${esc(data.sortCode.toUpperCase())}</text>

    <rect x="90" y="95" width="70" height="38" fill="#000"/>
    <text x="125" y="122" class="rm-text rm-black" font-size="28" fill="#fff" text-anchor="middle">${esc(data.routingCode.toUpperCase())}</text>

    <!-- Parcel info box (Click & Drop style) -->
    ${isClickDrop ? `
    <rect x="280" y="95" width="105" height="38" fill="none" stroke="#000" stroke-width="1"/>
    <text x="332" y="113" class="rm-text rm-bold" font-size="12" text-anchor="middle">${esc(data.parcelType)}</text>
    <text x="332" y="130" class="rm-text rm-bold" font-size="16" text-anchor="middle">${esc(data.weight)}</text>
    ` : ''}

    <!-- Row 3: Reference (eBay style) -->
    ${!isClickDrop ? `
    <line x1="0" y1="165" x2="400" y2="165" stroke="#000" stroke-width="1"/>
    <text x="15" y="157" class="rm-text" font-size="13" fill="#444">${esc(data.reference)}</text>
    ` : ''}

    <!-- Row 4: Barcodes -->
    <line x1="0" y1="${isClickDrop ? 265 : 290}" x2="400" y2="${isClickDrop ? 265 : 290}" stroke="#000" stroke-width="1"/>

    <!-- DataMatrix placeholder -->
    <foreignObject x="15" y="${isClickDrop ? 150 : 175}" width="100" height="100">
        <canvas xmlns="http://www.w3.org/1999/xhtml" id="dmCanvas" width="100" height="100" style="width:100px;height:100px;"></canvas>
    </foreignObject>

    <!-- Code128 placeholder -->
    <foreignObject x="130" y="${isClickDrop ? 160 : 185}" width="255" height="70">
        <canvas xmlns="http://www.w3.org/1999/xhtml" id="c128Canvas" width="255" height="50" style="width:255px;height:50px;"></canvas>
    </foreignObject>

    <!-- Tracking number text -->
    <text x="257" y="${isClickDrop ? 250 : 275}" class="rm-text" font-size="14" text-anchor="middle">${esc(trackingDisplay)}</text>

    <!-- Row 5: Address -->
    <line x1="0" y1="${isClickDrop ? 420 : 445}" x2="400" y2="${isClickDrop ? 420 : 445}" stroke="#000" stroke-width="1"/>

    <!-- Recipient address -->
    <text x="15" y="${isClickDrop ? 295 : 320}" class="rm-text rm-bold" font-size="20">${esc(data.recipientName.toUpperCase())}</text>
    <text x="15" y="${isClickDrop ? 320 : 345}" class="rm-text" font-size="17">${esc(data.recipientAddr1)}</text>
    <text x="15" y="${isClickDrop ? 345 : 370}" class="rm-text" font-size="17">${esc(data.recipientAddr2)}</text>
    <text x="15" y="${isClickDrop ? 370 : 395}" class="rm-text" font-size="17">${esc(data.recipientCity)}</text>
    <text x="15" y="${isClickDrop ? 400 : 425}" class="rm-text rm-bold" font-size="24">${esc(data.recipientPostcode.toUpperCase())}</text>

    <!-- Return address (rotated) -->
    <g transform="translate(385, ${isClickDrop ? 410 : 435}) rotate(-90)">
        ${returnLines.map((line, i) => `<text x="${i * 12}" y="0" class="rm-text rm-tiny" fill="#333">${esc(line)}</text>`).join('')}
    </g>
    <line x1="365" y1="${isClickDrop ? 275 : 300}" x2="365" y2="${isClickDrop ? 410 : 435}" stroke="#ccc" stroke-width="1"/>

    <!-- Row 6: Footer -->
    <line x1="0" y1="${isClickDrop ? 530 : 545}" x2="400" y2="${isClickDrop ? 530 : 545}" stroke="#000" stroke-width="1"/>

    <!-- Seller type (eBay style) -->
    ${!isClickDrop && data.sellerType ? `
    <text x="15" y="${445 + 25}" class="rm-text rm-black" font-size="20">${esc(data.sellerType.toUpperCase().split(' ')[0] || '')}</text>
    <text x="15" y="${445 + 50}" class="rm-text rm-black" font-size="20">${esc(data.sellerType.toUpperCase().split(' ').slice(1).join(' ') || '')}</text>
    ` : ''}

    <!-- Payment info -->
    ${postageDisplay ? `
    <text x="385" y="${isClickDrop ? 445 : 470}" class="rm-text rm-small" text-anchor="end">Postage Cost</text>
    <text x="385" y="${isClickDrop ? 465 : 490}" class="rm-text rm-bold" font-size="20" text-anchor="end">${esc(postageDisplay)}</text>
    ` : ''}

    <text x="385" y="${isClickDrop ? 485 : 505}" class="rm-text rm-small" text-anchor="end">Post by the end of</text>
    <text x="385" y="${isClickDrop ? 502 : 520}" class="rm-text rm-bold" font-size="14" text-anchor="end">${esc(data.postByDate)}</text>

    <text x="385" y="${isClickDrop ? 518 : 533}" class="rm-text rm-small" text-anchor="end">Paid and printed from</text>
    <text x="385" y="${isClickDrop ? 532 : 545}" class="rm-text rm-bold" font-size="14" text-anchor="end">${esc(data.printedFrom)}</text>

    <!-- Row 7: Carbon footer -->
    <text x="15" y="${isClickDrop ? 555 : 575}" class="rm-text" font-size="9" fill="#666">Royal Mail: UK's lowest average parcel carbon footprint 200g CO2e</text>
</svg>`;

        labelPreview.innerHTML = `<div class="svg-label-container">${svgLabel}</div>`;

        // Generate barcodes after SVG is in DOM
        setTimeout(async () => {
            const dmContent = buildDMContent(data);
            console.log('DM:', dmContent);
            await genDataMatrix(dmContent, 'dmCanvas');
            await genCode128(trackingClean, 'c128Canvas');
        }, 100);

        downloadPdfBtn.disabled = false;
        downloadPngBtn.disabled = false;
        printBtn.disabled = false;
    });

    // ── Export PNG ──────────────────────────────────────────────────────
    downloadPngBtn.addEventListener('click', async () => {
        const container = labelPreview.querySelector('.svg-label-container');
        if (!container) return;
        const canvas = await html2canvas(container, { scale: 4, backgroundColor: '#ffffff' });
        const link = document.createElement('a');
        link.download = 'royal-mail-tracked24.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });

    // ── Export PDF ──────────────────────────────────────────────────────
    downloadPdfBtn.addEventListener('click', async () => {
        const container = labelPreview.querySelector('.svg-label-container');
        if (!container) return;
        const canvas = await html2canvas(container, { scale: 4, backgroundColor: '#ffffff' });
        const imgData = canvas.toDataURL('image/png');
        const { jsPDF } = window.jspdf;
        const pdf = new jsPDF({ unit: 'mm', format: [101.6, 152.4] });
        const pdfW = pdf.internal.pageSize.getWidth();
        const pdfH = pdf.internal.pageSize.getHeight();
        const ratio = Math.min(pdfW / canvas.width, pdfH / canvas.height);
        const w = canvas.width * ratio;
        const h = canvas.height * ratio;
        pdf.addImage(imgData, 'PNG', (pdfW - w) / 2, (pdfH - h) / 2, w, h);
        pdf.save('royal-mail-tracked24.pdf');
    });

    // ── Print ───────────────────────────────────────────────────────────
    printBtn.addEventListener('click', () => window.print());

    // ── Clear ───────────────────────────────────────────────────────────
    clearBtn.addEventListener('click', () => {
        fields.forEach(f => {
            const el = $(f);
            if (el) {
                if (el.tagName === 'SELECT') el.selectedIndex = 0;
                else el.value = '';
            }
        });
        customPrintedFromGroup.style.display = 'none';
        labelPreview.innerHTML = '<div class="placeholder-message"><p>Fill in details and click <strong>Generate Label</strong></p></div>';
        downloadPdfBtn.disabled = true;
        downloadPngBtn.disabled = true;
        printBtn.disabled = true;
    });

    // ── Templates ───────────────────────────────────────────────────────
    const STORAGE_KEY = 'rmTracked24Templates';

    function getTemplates() {
        try { return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); }
        catch { return {}; }
    }

    function saveTemplates(t) { localStorage.setItem(STORAGE_KEY, JSON.stringify(t)); }

    function renderTemplates() {
        const templates = getTemplates();
        const names = Object.keys(templates);
        if (!names.length) {
            templateList.innerHTML = '<div style="font-size:0.8rem;color:#555;">No saved templates.</div>';
            return;
        }
        templateList.innerHTML = names.map(n => `
            <div class="template-item">
                <button onclick="window._loadTpl('${esc(n)}')">${esc(n)}</button>
                <button class="btn btn-danger" onclick="window._delTpl('${esc(n)}')">Delete</button>
            </div>
        `).join('');
    }

    saveTemplateBtn.addEventListener('click', () => {
        const name = templateNameInput.value.trim();
        if (!name) { alert('Enter template name.'); return; }
        const t = getTemplates();
        t[name] = getFormData();
        saveTemplates(t);
        templateNameInput.value = '';
        renderTemplates();
    });

    window._loadTpl = n => { const t = getTemplates(); if (t[n]) setFormData(t[n]); };
    window._delTpl = n => { const t = getTemplates(); delete t[n]; saveTemplates(t); renderTemplates(); };

    renderTemplates();
})();
