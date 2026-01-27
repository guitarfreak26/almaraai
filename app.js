(function () {
    'use strict';

    // ── Courier configurations ──────────────────────────────────────────
    const COURIERS = {
        'royal-mail': {
            name: 'Royal Mail',
            services: [
                'First Class',
                'Second Class',
                'Special Delivery Guaranteed by 1pm',
                'Special Delivery Guaranteed by 9am',
                'Tracked 24',
                'Tracked 48',
                'Royal Mail 24',
                'Royal Mail 48',
            ],
            trackingPrefix: 'RM',
        },
        ups: {
            name: 'UPS',
            services: [
                'UPS Express',
                'UPS Express Saver',
                'UPS Standard',
                'UPS Expedited',
                'UPS Access Point',
            ],
            trackingPrefix: '1Z',
        },
        evri: {
            name: 'Evri',
            services: [
                'Standard Delivery',
                'Next Day',
                'ParcelShop Drop-off',
                'Economy',
            ],
            trackingPrefix: 'EV',
        },
        dpd: {
            name: 'DPD',
            services: [
                'DPD Next Day',
                'DPD Express',
                'DPD Local',
                'DPD Saturday',
                'DPD Two Day',
            ],
            trackingPrefix: 'DPD',
        },
        dhl: {
            name: 'DHL',
            services: [
                'DHL Express',
                'DHL Economy',
                'DHL Parcel UK',
                'DHL Same Day',
            ],
            trackingPrefix: 'JD',
        },
        yodel: {
            name: 'Yodel',
            services: [
                'Yodel Direct',
                'Yodel Xpress 24',
                'Yodel Xpress 48',
                'Yodel Economy',
            ],
            trackingPrefix: 'YOL',
        },
    };

    // ── State ───────────────────────────────────────────────────────────
    let selectedCourier = 'royal-mail';

    // ── DOM refs ────────────────────────────────────────────────────────
    const $ = (id) => document.getElementById(id);
    const courierGrid = $('courierGrid');
    const serviceType = $('serviceType');
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
        'senderName', 'senderAddr1', 'senderAddr2', 'senderCity', 'senderPostcode', 'senderPhone',
        'recipientName', 'recipientAddr1', 'recipientAddr2', 'recipientCity', 'recipientPostcode', 'recipientPhone',
        'weight', 'reference', 'instructions',
    ];

    // ── Courier selection ───────────────────────────────────────────────
    courierGrid.addEventListener('click', (e) => {
        const btn = e.target.closest('.courier-btn');
        if (!btn) return;
        courierGrid.querySelectorAll('.courier-btn').forEach((b) => b.classList.remove('selected'));
        btn.classList.add('selected');
        selectedCourier = btn.dataset.courier;
        populateServices();
    });

    function populateServices() {
        const courier = COURIERS[selectedCourier];
        serviceType.innerHTML = courier.services
            .map((s) => `<option value="${s}">${s}</option>`)
            .join('');
    }

    // ── Generate tracking number ────────────────────────────────────────
    function generateTrackingNumber() {
        const prefix = COURIERS[selectedCourier].trackingPrefix;
        const chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789';
        let code = prefix;
        for (let i = 0; i < 14; i++) {
            code += chars[Math.floor(Math.random() * chars.length)];
        }
        return code;
    }

    // ── Collect form data ───────────────────────────────────────────────
    function getFormData() {
        const data = {};
        fields.forEach((f) => {
            data[f] = $(f).value.trim();
        });
        data.service = serviceType.value;
        data.courier = selectedCourier;
        return data;
    }

    function setFormData(data) {
        fields.forEach((f) => {
            if (data[f] !== undefined) $(f).value = data[f];
        });
        if (data.courier) {
            selectedCourier = data.courier;
            courierGrid.querySelectorAll('.courier-btn').forEach((b) => {
                b.classList.toggle('selected', b.dataset.courier === data.courier);
            });
            populateServices();
        }
        if (data.service) {
            serviceType.value = data.service;
        }
    }

    // ── Generate label ──────────────────────────────────────────────────
    generateBtn.addEventListener('click', () => {
        const data = getFormData();

        if (!data.recipientName || !data.recipientAddr1 || !data.recipientPostcode) {
            alert('Please fill in at least the recipient name, address line 1, and postcode.');
            return;
        }

        const tracking = generateTrackingNumber();
        const courierConfig = COURIERS[selectedCourier];
        const date = new Date().toLocaleDateString('en-GB');

        const labelHtml = `
            <div class="shipping-label ${selectedCourier}">
                <div class="label-header">
                    <span class="label-courier-name">${courierConfig.name}</span>
                    <span class="label-service-badge">${data.service}</span>
                </div>
                <div class="label-body">
                    <div class="label-qr-row">
                        <div id="qrCode"></div>
                        <div class="label-tracking-info">
                            <div class="label-tracking-label">Tracking Number</div>
                            <div class="label-tracking-number">${tracking}</div>
                            <div class="label-barcode-text">${tracking}</div>
                        </div>
                    </div>

                    <div class="label-section">
                        <div class="label-section-title">Deliver To</div>
                        <div class="label-address-name">${esc(data.recipientName)}</div>
                        <div class="label-address-line">${esc(data.recipientAddr1)}</div>
                        ${data.recipientAddr2 ? `<div class="label-address-line">${esc(data.recipientAddr2)}</div>` : ''}
                        <div class="label-address-line">${esc(data.recipientCity)}</div>
                        <div class="label-postcode">${esc(data.recipientPostcode.toUpperCase())}</div>
                        ${data.recipientPhone ? `<div class="label-address-line">Tel: ${esc(data.recipientPhone)}</div>` : ''}
                    </div>

                    <div class="label-section">
                        <div class="label-section-title">From</div>
                        <div class="label-address-name">${esc(data.senderName)}</div>
                        <div class="label-address-line">${esc(data.senderAddr1)}</div>
                        ${data.senderAddr2 ? `<div class="label-address-line">${esc(data.senderAddr2)}</div>` : ''}
                        <div class="label-address-line">${esc(data.senderCity)} ${esc(data.senderPostcode.toUpperCase())}</div>
                        ${data.senderPhone ? `<div class="label-address-line">Tel: ${esc(data.senderPhone)}</div>` : ''}
                    </div>

                    ${data.instructions ? `
                    <div class="label-section">
                        <div class="label-section-title">Delivery Instructions</div>
                        <div class="label-address-line">${esc(data.instructions)}</div>
                    </div>` : ''}

                    <div class="label-meta-row">
                        <div class="label-meta-item">
                            <strong>Weight</strong>
                            ${data.weight ? data.weight + ' kg' : 'N/A'}
                        </div>
                        <div class="label-meta-item">
                            <strong>Reference</strong>
                            ${esc(data.reference) || 'N/A'}
                        </div>
                        <div class="label-meta-item">
                            <strong>Date</strong>
                            ${date}
                        </div>
                    </div>
                </div>
                <div class="label-footer">
                    This label was generated for demonstration purposes.
                </div>
            </div>
        `;

        labelPreview.innerHTML = labelHtml;

        // Generate QR code encoding tracking + postcode
        const qrData = JSON.stringify({
            tracking,
            courier: courierConfig.name,
            service: data.service,
            postcode: data.recipientPostcode.toUpperCase(),
            recipient: data.recipientName,
        });

        new QRCode(document.getElementById('qrCode'), {
            text: qrData,
            width: 100,
            height: 100,
            correctLevel: QRCode.CorrectLevel.M,
        });

        downloadPdfBtn.disabled = false;
        downloadPngBtn.disabled = false;
        printBtn.disabled = false;
    });

    function esc(str) {
        const d = document.createElement('div');
        d.textContent = str || '';
        return d.innerHTML;
    }

    // ── Export: PNG ──────────────────────────────────────────────────────
    downloadPngBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.shipping-label');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 2, backgroundColor: '#ffffff' });
        const link = document.createElement('a');
        link.download = 'shipping-label.png';
        link.href = canvas.toDataURL('image/png');
        link.click();
    });

    // ── Export: PDF ──────────────────────────────────────────────────────
    downloadPdfBtn.addEventListener('click', async () => {
        const label = labelPreview.querySelector('.shipping-label');
        if (!label) return;
        const canvas = await html2canvas(label, { scale: 2, backgroundColor: '#ffffff' });
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
        pdf.save('shipping-label.pdf');
    });

    // ── Print ───────────────────────────────────────────────────────────
    printBtn.addEventListener('click', () => {
        window.print();
    });

    // ── Clear ───────────────────────────────────────────────────────────
    clearBtn.addEventListener('click', () => {
        fields.forEach((f) => ($(f).value = ''));
        labelPreview.innerHTML = '<div class="placeholder-message"><p>Fill in the details and click <strong>Generate Label</strong> to preview your shipping label here.</p></div>';
        downloadPdfBtn.disabled = true;
        downloadPngBtn.disabled = true;
        printBtn.disabled = true;
    });

    // ── Templates (localStorage) ────────────────────────────────────────
    function getTemplates() {
        try {
            return JSON.parse(localStorage.getItem('shippingTemplates') || '{}');
        } catch {
            return {};
        }
    }

    function saveTemplates(templates) {
        localStorage.setItem('shippingTemplates', JSON.stringify(templates));
    }

    function renderTemplates() {
        const templates = getTemplates();
        const names = Object.keys(templates);
        if (names.length === 0) {
            templateList.innerHTML = '<div style="font-size:0.8rem;color:#9ca3af;">No saved templates yet.</div>';
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
    populateServices();
    renderTemplates();
})();
