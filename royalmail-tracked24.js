(function () {
    'use strict';

    // ── DOM References ──────────────────────────────────────────────────
    const $ = (id) => document.getElementById(id);

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

    // All form fields for templates
    const fields = [
        'signatureType', 'trackingNumber', 'sortCode', 'routingCode', 'reference',
        'parcelType', 'weight', 'postage', 'postByDate',
        'recipientName', 'recipientAddr1', 'recipientAddr2', 'recipientCity', 'recipientPostcode',
        'senderName', 'senderAddr1', 'senderCity', 'senderPostcode',
        'sellerType', 'printedFrom', 'customPrintedFrom'
    ];

    // Royal Mail logo SVG with crown
    const ROYAL_MAIL_LOGO = `<svg viewBox="0 0 85 45" xmlns="http://www.w3.org/2000/svg" class="rm-logo-svg">
        <!-- Crown -->
        <g transform="translate(22, 2)">
            <!-- Crown base -->
            <path d="M0 20 L4 8 L10 14 L20 4 L30 14 L36 8 L40 20 Z" fill="#000" stroke="#000" stroke-width="1"/>
            <!-- Crown jewels/dots -->
            <circle cx="4" cy="6" r="2" fill="#000"/>
            <circle cx="20" cy="2" r="2.5" fill="#000"/>
            <circle cx="36" cy="6" r="2" fill="#000"/>
            <!-- Cross on top -->
            <rect x="18" y="0" width="4" height="4" fill="#000"/>
        </g>
        <!-- Royal Mail text -->
        <text x="42.5" y="38" text-anchor="middle" font-family="Arial, sans-serif" font-size="11" font-weight="bold" fill="#000">Royal Mail</text>
    </svg>`;

    // ── Custom printed from toggle ──────────────────────────────────────
    printedFrom.addEventListener('change', () => {
        customPrintedFromGroup.style.display = printedFrom.value === 'Custom' ? 'block' : 'none';
    });

    // ── Helper Functions ────────────────────────────────────────────────
    function esc(str) {
        const d = document.createElement('div');
        d.textContent = str || '';
        return d.innerHTML;
    }

    function removeSpaces(str) {
        return (str || '').replace(/\s+/g, '');
    }

    function formatTrackingForDisplay(tracking) {
        // Format: MZ 3170 8295 1GB
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

    function formatPostageForDataMatrix(val) {
        const pence = Math.round(parsePostage(val) * 100);
        return pence.toString().padStart(5, '0');
    }

    function formatReferenceForDataMatrix(ref) {
        return removeSpaces((ref || '').replace(/-/g, '')).toUpperCase();
    }

    function generateDateCode() {
        const now = new Date();
        const dd = now.getDate().toString().padStart(2, '0');
        const mm = (now.getMonth() + 1).toString().padStart(2, '0');
        const yy = now.getFullYear().toString().slice(-2);
        const seq = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        return `${dd}${mm}${yy}${seq}`;
    }

    // ── DataMatrix Encoding (Mailmark format) ───────────────────────────
    function buildDataMatrixContent(data) {
        const accountId = '8215FA';
        const refCode = formatReferenceForDataMatrix(data.reference);
        const serviceCode = data.signatureType === 'signature' ? '00010001' : '00020001';
        const postageFmt = formatPostageForDataMatrix(data.postage);
        const dateCode = generateDateCode();
        const trackingClean = removeSpaces(data.tracking).toUpperCase();
        const destPostcode = removeSpaces(data.recipientPostcode).toUpperCase();
        const returnPostcode = removeSpaces(data.senderPostcode).toUpperCase();

        let routeCode = removeSpaces(data.routingCode).toUpperCase();
        if (!routeCode) {
            routeCode = destPostcode.slice(0, 3);
        }

        return `JGB${accountId}${refCode}${serviceCode}${postageFmt}${dateCode}062${trackingClean}${routeCode.slice(0,3)}${destPostcode}GB${returnPostcode}`;
    }

    // ── Generate Barcodes ───────────────────────────────────────────────
    async function generateDataMatrix(content, canvasId) {
        try {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;

            // DataMatrix matching Royal Mail label size (~72x72px display)
            bwipjs.toCanvas(canvas, {
                bcid: 'datamatrix',
                text: content,
                scale: 3,
                height: 24,
                width: 24,
                padding: 0,
                includetext: false,
            });
        } catch (e) {
            console.error('DataMatrix generation failed:', e);
        }
    }

    async function generateCode128(content, canvasId) {
        try {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;

            // Code 128 matching Royal Mail label size
            bwipjs.toCanvas(canvas, {
                bcid: 'code128',
                text: content,
                scale: 2,
                height: 18,
                width: 200,
                includetext: false,
            });
        } catch (e) {
            console.error('Code128 generation failed:', e);
        }
    }

    // ── Collect Form Data ───────────────────────────────────────────────
    function getFormData() {
        return {
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

        if (!data.tracking) {
            alert('Please enter a tracking number.');
            return;
        }
        if (!data.recipientPostcode) {
            alert('Please enter the recipient postcode.');
            return;
        }

        const isSignature = data.signatureType === 'signature';
        const signatureText = isSignature ? '' : 'No Signature';
        const trackingDisplay = formatTrackingForDisplay(data.tracking);
        const trackingClean = removeSpaces(data.tracking).toUpperCase();

        // Build return address lines
        const returnLines = [
            'Return Address',
            data.senderName,
            data.senderAddr1,
            data.senderCity,
            data.senderPostcode.toUpperCase()
        ].filter(Boolean);

        // Postage display
        const postageNum = parsePostage(data.postage);
        const postageDisplay = postageNum > 0 ? `£${postageNum.toFixed(2)}` : '';

        const labelHtml = `
            <div class="rml">
                <!-- Row 1: Header -->
                <div class="rml-header">
                    <div class="rml-header-left">
                        <div class="rml-tracked">Tracked</div>
                        <div class="rml-nosig">${signatureText}</div>
                    </div>
                    <div class="rml-header-mid">
                        <div class="rml-24">24</div>
                    </div>
                    <div class="rml-header-right">
                        <div class="rml-delivered">Delivered By</div>
                        <div class="rml-logo-box">
                            ${ROYAL_MAIL_LOGO}
                        </div>
                        <div class="rml-postage-paid">Postage Paid GB</div>
                    </div>
                </div>

                <!-- Row 2: Routing -->
                <div class="rml-routing">
                    <div class="rml-codes">
                        <span class="rml-code">${esc(data.sortCode.toUpperCase())}</span>
                        <span class="rml-code rml-code-inv">${esc(data.routingCode.toUpperCase())}</span>
                    </div>
                    <div class="rml-parcel-box">
                        <div class="rml-parcel-type">${esc(data.parcelType)}</div>
                        <div class="rml-parcel-weight">${esc(data.weight)}</div>
                    </div>
                </div>

                <!-- Row 3: Reference -->
                <div class="rml-ref">${esc(data.reference)}</div>

                <!-- Row 4: Barcodes -->
                <div class="rml-barcodes">
                    <div class="rml-dm">
                        <canvas id="dmCanvas"></canvas>
                    </div>
                    <div class="rml-c128">
                        <canvas id="c128Canvas"></canvas>
                        <div class="rml-tracking">${trackingDisplay}</div>
                    </div>
                </div>

                <!-- Row 5: Address -->
                <div class="rml-address">
                    <div class="rml-recipient">
                        <div class="rml-addr-line rml-addr-name">${esc(data.recipientName.toUpperCase())}</div>
                        ${data.recipientAddr1 ? `<div class="rml-addr-line">${esc(data.recipientAddr1)}</div>` : ''}
                        ${data.recipientAddr2 ? `<div class="rml-addr-line">${esc(data.recipientAddr2)}</div>` : ''}
                        ${data.recipientCity ? `<div class="rml-addr-line">${esc(data.recipientCity)}</div>` : ''}
                        <div class="rml-addr-line rml-addr-pc">${esc(data.recipientPostcode.toUpperCase())}</div>
                    </div>
                    <div class="rml-return">
                        ${returnLines.map(l => `<span>${esc(l)}</span>`).join('')}
                    </div>
                </div>

                <!-- Row 6: Footer info -->
                <div class="rml-footer">
                    <div class="rml-seller">${data.sellerType ? esc(data.sellerType.toUpperCase()) : ''}</div>
                    <div class="rml-payment">
                        ${postageDisplay ? `<div class="rml-pay-row"><span>Postage Cost</span><strong>${postageDisplay}</strong></div>` : ''}
                        ${data.postByDate ? `<div class="rml-pay-row"><span>Post by the end of</span><strong>${esc(data.postByDate)}</strong></div>` : ''}
                        <div class="rml-pay-row"><span>Paid and printed from</span><strong>${esc(data.printedFrom)}</strong></div>
                    </div>
                </div>

                <!-- Row 7: Carbon footer -->
                <div class="rml-carbon">Royal Mail: UK's lowest average parcel carbon footprint 200g CO2e</div>
            </div>
        `;

        labelPreview.innerHTML = labelHtml;

        // Generate barcodes
        setTimeout(async () => {
            const dmContent = buildDataMatrixContent(data);
            console.log('DataMatrix:', dmContent);
            await generateDataMatrix(dmContent, 'dmCanvas');
            await generateCode128(trackingClean, 'c128Canvas');
        }, 50);

        downloadPdfBtn.disabled = false;
        downloadPngBtn.disabled = false;
        printBtn.disabled = false;
    });

    // ── Export: PNG ─────────────────────────────────────────────────────
    downloadPngBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.rml');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 4, backgroundColor: '#ffffff' });
        const link = document.createElement('a');
        link.download = 'royal-mail-tracked24-label.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });

    // ── Export: PDF ─────────────────────────────────────────────────────
    downloadPdfBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.rml');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 4, backgroundColor: '#ffffff' });
        const imgData = canvas.toDataURL('image/png');
        const { jsPDF } = window.jspdf;
        // 4x6 inch label
        const pdf = new jsPDF({ unit: 'mm', format: [101.6, 152.4] });
        const pdfW = pdf.internal.pageSize.getWidth();
        const pdfH = pdf.internal.pageSize.getHeight();
        const ratio = Math.min(pdfW / canvas.width, pdfH / canvas.height);
        const w = canvas.width * ratio;
        const h = canvas.height * ratio;
        pdf.addImage(imgData, 'PNG', (pdfW - w) / 2, (pdfH - h) / 2, w, h);
        pdf.save('royal-mail-tracked24-label.pdf');
    });

    // ── Print ───────────────────────────────────────────────────────────
    printBtn.addEventListener('click', () => window.print());

    // ── Clear ───────────────────────────────────────────────────────────
    clearBtn.addEventListener('click', () => {
        fields.forEach((f) => {
            const el = $(f);
            if (el) {
                if (el.tagName === 'SELECT') el.selectedIndex = 0;
                else el.value = '';
            }
        });
        customPrintedFromGroup.style.display = 'none';
        labelPreview.innerHTML = '<div class="placeholder-message"><p>Fill in the details and click <strong>Generate Label</strong> to preview.</p></div>';
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

    function saveTemplates(templates) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
    }

    function renderTemplates() {
        const templates = getTemplates();
        const names = Object.keys(templates);
        if (names.length === 0) {
            templateList.innerHTML = '<div style="font-size:0.8rem;color:#555;">No saved templates.</div>';
            return;
        }
        templateList.innerHTML = names.map(name => `
            <div class="template-item">
                <button onclick="window._loadTemplate('${esc(name)}')">${esc(name)}</button>
                <button class="btn btn-danger" onclick="window._deleteTemplate('${esc(name)}')">Delete</button>
            </div>
        `).join('');
    }

    saveTemplateBtn.addEventListener('click', () => {
        const name = templateNameInput.value.trim();
        if (!name) { alert('Enter a template name.'); return; }
        const templates = getTemplates();
        templates[name] = getFormData();
        saveTemplates(templates);
        templateNameInput.value = '';
        renderTemplates();
    });

    window._loadTemplate = (name) => {
        const templates = getTemplates();
        if (templates[name]) setFormData(templates[name]);
    };

    window._deleteTemplate = (name) => {
        const templates = getTemplates();
        delete templates[name];
        saveTemplates(templates);
        renderTemplates();
    };

    renderTemplates();
})();
