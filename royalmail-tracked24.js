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
        // Extract numeric value from "£4.29" or "4.29"
        const match = (val || '').replace(/[£,]/g, '').match(/[\d.]+/);
        return match ? parseFloat(match[0]) : 0;
    }

    function formatPostageForDataMatrix(val) {
        // Convert £4.29 to "00429" (5 digits, pence)
        const pence = Math.round(parsePostage(val) * 100);
        return pence.toString().padStart(5, '0');
    }

    function formatReferenceForDataMatrix(ref) {
        // Remove dashes and spaces: "11-02D 71C AC8" -> "1102D71CAC8"
        return removeSpaces((ref || '').replace(/-/g, '')).toUpperCase();
    }

    function generateDateCode() {
        // Generate a date-based code (format appears to be DDMMYY + sequence)
        const now = new Date();
        const dd = now.getDate().toString().padStart(2, '0');
        const mm = (now.getMonth() + 1).toString().padStart(2, '0');
        const yy = now.getFullYear().toString().slice(-2);
        // Add a random 3-digit sequence
        const seq = Math.floor(Math.random() * 1000).toString().padStart(3, '0');
        return `${dd}${mm}${yy}${seq}`;
    }

    // ── DataMatrix Encoding (Mailmark format) ───────────────────────────
    function buildDataMatrixContent(data) {
        // Format: JGB [AccountID] [Reference] [ServiceCode][Postage] [DateCode]062 [Tracking] [RouteCode] [DestPostcode] GB [ReturnPostcode]
        // Example: JGB 8215FA1102D71CAC80002000100429070425062 MZ317082951GB817 N49PT GB M610YU

        const accountId = '8215FA'; // Fixed account ID from examples
        const refCode = formatReferenceForDataMatrix(data.reference);
        const serviceCode = data.signatureType === 'signature' ? '00010001' : '00020001';
        const postageFmt = formatPostageForDataMatrix(data.postage);
        const dateCode = generateDateCode();
        const trackingClean = removeSpaces(data.tracking).toUpperCase();
        const destPostcode = removeSpaces(data.recipientPostcode).toUpperCase();
        const returnPostcode = removeSpaces(data.senderPostcode).toUpperCase();

        // Route code - use provided or derive from postcode
        let routeCode = removeSpaces(data.routingCode).toUpperCase();
        if (!routeCode) {
            // Simple derivation from postcode
            routeCode = destPostcode.slice(0, 3);
        }

        // Build the DataMatrix string
        // Note: The exact format has the tracking number followed by a shorter route indicator
        const dataMatrixContent = `JGB${accountId}${refCode}${serviceCode}${postageFmt}${dateCode}062${trackingClean}${routeCode.slice(0,3)}${destPostcode}GB${returnPostcode}`;

        return dataMatrixContent;
    }

    // ── Generate Barcodes ───────────────────────────────────────────────
    async function generateDataMatrix(content, canvasId) {
        try {
            const canvas = document.getElementById(canvasId);
            if (!canvas) return;

            bwipjs.toCanvas(canvas, {
                bcid: 'datamatrix',
                text: content,
                scale: 3,
                height: 25,
                width: 25,
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

            bwipjs.toCanvas(canvas, {
                bcid: 'code128',
                text: content,
                scale: 2,
                height: 15,
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

        // Validation
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

        // Build return address for rotated display
        const returnAddressLines = [
            'Return Address',
            data.senderName,
            data.senderAddr1,
            data.senderCity,
            data.senderPostcode.toUpperCase()
        ].filter(Boolean);

        // Postage display
        const postageDisplay = data.postage.startsWith('£') ? data.postage : `£${data.postage}`;

        const labelHtml = `
            <div class="rm-label">
                <!-- Header Row -->
                <div class="rm-header">
                    <div class="rm-header-left">
                        <div class="rm-service-name">Tracked</div>
                        <div class="rm-signature-text">${signatureText}</div>
                    </div>
                    <div class="rm-header-center">
                        <div class="rm-24">24</div>
                    </div>
                    <div class="rm-header-right">
                        <div class="rm-delivered-by">Delivered By</div>
                        <div class="rm-logo">
                            <svg viewBox="0 0 100 60" class="rm-crown-logo">
                                <text x="50" y="35" text-anchor="middle" font-size="12" font-weight="bold" fill="#000">Royal Mail</text>
                            </svg>
                        </div>
                        <div class="rm-postage-paid">Postage Paid GB</div>
                    </div>
                </div>

                <!-- Routing Row -->
                <div class="rm-routing-row">
                    <div class="rm-routing-codes">
                        <span class="rm-sort-code">${esc(data.sortCode.toUpperCase())}</span>
                        <span class="rm-route-code">${esc(data.routingCode.toUpperCase())}</span>
                    </div>
                    <div class="rm-parcel-info">
                        <div class="rm-parcel-type">${esc(data.parcelType)}</div>
                        <div class="rm-weight">${esc(data.weight)}</div>
                    </div>
                </div>

                <!-- Reference Row -->
                <div class="rm-reference-row">
                    ${esc(data.reference)}
                </div>

                <!-- Barcode Row -->
                <div class="rm-barcode-row">
                    <div class="rm-datamatrix">
                        <canvas id="datamatrixCanvas"></canvas>
                    </div>
                    <div class="rm-code128-container">
                        <canvas id="code128Canvas"></canvas>
                        <div class="rm-tracking-text">${trackingDisplay}</div>
                    </div>
                </div>

                <!-- Address Section -->
                <div class="rm-address-section">
                    <div class="rm-recipient-address">
                        <div class="rm-address-line rm-name">${esc(data.recipientName.toUpperCase())}</div>
                        ${data.recipientAddr1 ? `<div class="rm-address-line">${esc(data.recipientAddr1.toUpperCase())}</div>` : ''}
                        ${data.recipientAddr2 ? `<div class="rm-address-line">${esc(data.recipientAddr2.toUpperCase())}</div>` : ''}
                        ${data.recipientCity ? `<div class="rm-address-line">${esc(data.recipientCity.toUpperCase())}</div>` : ''}
                        <div class="rm-address-line rm-postcode">${esc(data.recipientPostcode.toUpperCase())}</div>
                    </div>
                    <div class="rm-return-address">
                        ${returnAddressLines.map(line => `<span>${esc(line)}</span>`).join('')}
                    </div>
                </div>

                <!-- Footer Section -->
                <div class="rm-footer-section">
                    ${data.sellerType ? `<div class="rm-seller-type">${esc(data.sellerType.toUpperCase())}</div>` : '<div class="rm-seller-type"></div>'}
                    <div class="rm-postage-info">
                        ${data.postage ? `<div class="rm-postage-cost"><span>Postage Cost</span><strong>${postageDisplay}</strong></div>` : ''}
                        ${data.postByDate ? `<div class="rm-post-by"><span>Post by the end of</span><strong>${esc(data.postByDate)}</strong></div>` : ''}
                        <div class="rm-printed-from"><span>Paid and printed from</span><strong>${esc(data.printedFrom)}</strong></div>
                    </div>
                </div>

                <!-- Bottom Footer -->
                <div class="rm-bottom-footer">
                    Royal Mail: UK's lowest average parcel carbon footprint 200g CO2e
                </div>
            </div>
        `;

        labelPreview.innerHTML = labelHtml;

        // Generate barcodes after DOM is ready
        setTimeout(async () => {
            // Generate DataMatrix with Mailmark format
            const dataMatrixContent = buildDataMatrixContent(data);
            console.log('DataMatrix content:', dataMatrixContent);
            await generateDataMatrix(dataMatrixContent, 'datamatrixCanvas');

            // Generate Code 128 with tracking number (no spaces)
            await generateCode128(trackingClean, 'code128Canvas');
        }, 50);

        // Enable export buttons
        downloadPdfBtn.disabled = false;
        downloadPngBtn.disabled = false;
        printBtn.disabled = false;
    });

    // ── Export: PNG ─────────────────────────────────────────────────────
    downloadPngBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.rm-label');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 3, backgroundColor: '#ffffff' });
        const link = document.createElement('a');
        link.download = 'royal-mail-tracked24-label.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });

    // ── Export: PDF ─────────────────────────────────────────────────────
    downloadPdfBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.rm-label');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 3, backgroundColor: '#ffffff' });
        const imgData = canvas.toDataURL('image/png');
        const { jsPDF } = window.jspdf;
        // 4x6 inch label = 101.6mm x 152.4mm
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
    printBtn.addEventListener('click', () => {
        window.print();
    });

    // ── Clear ───────────────────────────────────────────────────────────
    clearBtn.addEventListener('click', () => {
        fields.forEach((f) => {
            const el = $(f);
            if (el) {
                if (el.tagName === 'SELECT') {
                    el.selectedIndex = 0;
                } else {
                    el.value = '';
                }
            }
        });
        customPrintedFromGroup.style.display = 'none';
        labelPreview.innerHTML = '<div class="placeholder-message"><p>Fill in the details and click <strong>Generate Label</strong> to preview your Royal Mail Tracked 24 label here.</p></div>';
        downloadPdfBtn.disabled = true;
        downloadPngBtn.disabled = true;
        printBtn.disabled = true;
    });

    // ── Templates (localStorage) ────────────────────────────────────────
    const STORAGE_KEY = 'rmTracked24Templates';

    function getTemplates() {
        try {
            return JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}');
        } catch {
            return {};
        }
    }

    function saveTemplates(templates) {
        localStorage.setItem(STORAGE_KEY, JSON.stringify(templates));
    }

    function renderTemplates() {
        const templates = getTemplates();
        const names = Object.keys(templates);
        if (names.length === 0) {
            templateList.innerHTML = '<div style="font-size:0.8rem;color:#555;">No saved templates yet.</div>';
            return;
        }
        templateList.innerHTML = names
            .map(
                (name) => `
            <div class="template-item">
                <button onclick="window._loadTemplate('${esc(name)}')">${esc(name)}</button>
                <button class="btn btn-danger" onclick="window._deleteTemplate('${esc(name)}')">Delete</button>
            </div>`
            )
            .join('');
    }

    saveTemplateBtn.addEventListener('click', () => {
        const name = templateNameInput.value.trim();
        if (!name) {
            alert('Enter a template name.');
            return;
        }
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

    // ── Init ────────────────────────────────────────────────────────────
    renderTemplates();
})();
