<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>قارئ باركود GS1 مع تصدير إكسيل</title>
    <!-- استخدام Tailwind CSS لتصميم سريع وجميل -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- مكتبة قراءة الباركود Html5-Qrcode -->
    <script src="https://unpkg.com/html5-qrcode"></script>
    <!-- مكتبة SheetJS لتصدير ملفات Excel -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
</head>
<body class="bg-gray-100 font-sans min-h-screen p-4">

    <div class="max-w-md mx-auto bg-white rounded-xl shadow-md overflow-hidden md:max-w-2xl p-6">
        <h2 class="text-2xl font-bold text-center text-gray-800 mb-4">ماسح باركود GS1</h2>
        
        <!-- أزرار التحكم بالكاميرا -->
        <div class="text-center mb-6 space-y-2">
            <button id="start-btn" onclick="startScanner()" class="bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg w-full transition duration-200">
                📷 ابدأ مسح الباركود
            </button>
            <button id="stop-btn" onclick="stopScanner()" class="bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-6 rounded-lg shadow-lg w-full hidden transition duration-200">
                🛑 إيقاف الكاميرا
            </button>
        </div>

        <!-- مكان ظهور الكاميرا -->
        <div id="reader" class="w-full mb-4 rounded-lg overflow-hidden"></div>

        <!-- منطقة عرض النتائج وتصدير الإكسيل -->
        <div id="result-container" class="hidden bg-gray-50 border border-gray-200 rounded-lg p-4 shadow-inner space-y-4">
            <h3 class="text-lg font-semibold text-gray-700 border-b pb-2">📦 تفاصيل البيانات المستخرجة:</h3>
            
            <div class="space-y-2 text-sm text-gray-600">
                <p><strong>الباركود الخام:</strong> <span id="raw-data" class="text-black font-mono break-all"></span></p>
                <p><strong>كود المنتج (GTIN - 01):</strong> <span id="gtin-val" class="text-blue-600 font-mono font-bold"></span></p>
                <p><strong>تاريخ الصلاحية (17):</strong> <span id="expiry-val" class="text-green-600 font-mono font-bold"></span></p>
                <p><strong>رقم التشغيلة / الباتش (10):</strong> <span id="batch-val" class="text-purple-600 font-mono font-bold"></span></p>
            </div>

            <!-- زر تصدير الإكسيل -->
            <button onclick="exportToExcel()" class="bg-green-600 hover:bg-green-700 text-white font-bold py-2.5 px-4 rounded-lg shadow w-full transition duration-200 flex items-center justify-center gap-2">
                📊 تصدير إلى Excel
            </button>
        </div>
    </div>

    <script>
        let html5QrCode;
        let lastScannedData = {
            raw: "",
            gtin: "",
            expiry: "",
            batch: "",
            timestamp: ""
        };

        function startScanner() {
            document.getElementById('start-btn').classList.add('hidden');
            document.getElementById('stop-btn').classList.remove('hidden');
            document.getElementById('result-container').classList.add('hidden');

            html5QrCode = new Html5Qrcode("reader");
            
            // محاولة تشغيل الكاميرا الخلفية أولاً، ولو فشلت يتم تجربة الكاميرا البديلة لتجنب خطأ التعذر
            html5QrCode.start(
                { facingMode: "environment" },
                {
                    fps: 10,
                    qrbox: { width: 250, height: 250 }
                },
                (decodedText, decodedResult) => {
                    processGS1Data(decodedText);
                    stopScanner();
                },
                (errorMessage) => {}
            ).catch((err) => {
                // خطوة بديلة في حال واجه الهاتف مشكلة مع الكاميرا الخلفية
                html5QrCode.start(
                    { facingMode: "user" },
                    { fps: 10, qrbox: { width: 250, height: 250 } },
                    (decodedText, decodedResult) => {
                        processGS1Data(decodedText);
                        stopScanner();
                    }
                ).catch(innerErr => {
                    alert("عذراً، لا يمكن الوصول للكاميرا. تأكد أن المتصفح يمتلك الصلاحية وأن الصفحة تعمل عبر رابط آمن HTTPS.");
                    stopScanner();
                });
            });
        }

        function stopScanner() {
            if (html5QrCode && html5QrCode.isScanning) {
                html5QrCode.stop().then(() => {
                    document.getElementById('start-btn').classList.remove('hidden');
                    document.getElementById('stop-btn').classList.add('hidden');
                }).catch(err => {
                    console.error("فشل إيقاف الكاميرا", err);
                });
            } else {
                document.getElementById('start-btn').classList.remove('hidden');
                document.getElementById('stop-btn').classList.add('hidden');
            }
        }

        // دالة تحليل باركود GS1 واستخراج المعطيات
        function processGS1Data(raw) {
            document.getElementById('result-container').classList.remove('hidden');
            document.getElementById('raw-data').innerText = raw;

            let gtin = "غير متوفر";
            let expiry = "غير متوفر";
            let batch = "غير متوفر";
            let now = new Date().toLocaleString('ar-EG');

            // استخراج GTIN (يبدأ بـ 01 وطوله 14 رقم)
            let gtinIndex = raw.indexOf("01");
            if (gtinIndex !== -1 && raw.length >= gtinIndex + 16) {
                gtin = raw.substring(gtinIndex + 2, gtinIndex + 16);
            }

            // استخراج تاريخ الصلاحية (يبدأ بـ 17 وطوله 6 أرقام YYMMDD)
            let expiryIndex = raw.indexOf("17");
            if (expiryIndex !== -1 && raw.length >= expiryIndex + 8) {
                let yy = raw.substring(expiryIndex + 2, expiryIndex + 4);
                let mm = raw.substring(expiryIndex + 4, expiryIndex + 6);
                let dd = raw.substring(expiryIndex + 6, expiryIndex + 8);
                expiry = `20${yy}-${mm}-${dd}`;
            }

            // استخراج رقم الباتش (يبدأ بـ 10)
            let batchIndex = raw.indexOf("10");
            if (batchIndex !== -1) {
                let sub = raw.substring(batchIndex + 2);
                batch = sub.split(/[\x1d]/)[0]; 
            }

            // حفظ البيانات مؤقتاً لتصديرها
            lastScannedData = {
                raw: raw,
                gtin: gtin,
                expiry: expiry,
                batch: batch,
                timestamp: now
            };

            // عرض النتائج في الواجهة
            document.getElementById('gtin-val').innerText = gtin;
            document.getElementById('expiry-val').innerText = expiry;
            document.getElementById('batch-val').innerText = batch;
        }

        // دالة إنشاء وتنزيل ملف الإكسيل
        function exportToExcel() {
            let excelData = [
                {
                    "وقت المسح": lastScannedData.timestamp,
                    "كود المنتج (GTIN)": lastScannedData.gtin,
                    "تاريخ الصلاحية": lastScannedData.expiry,
                    "رقم التشغيلة (Batch)": lastScannedData.batch,
                    "الباركود الخام": lastScannedData.raw
                }
            ];

            let worksheet = XLSX.utils.json_to_sheet(excelData);
            let workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, worksheet, "GS1 Scan Data");

            let fileName = `GS1_Scan_${Date.now()}.xlsx`;
            XLSX.writeFile(workbook, fileName);
        }
    </script>
</body>
</html>
