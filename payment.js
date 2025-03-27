(function() {
    'use strict';
    
    // デバッグ用のログ出力関数
    const debugLog = (message, data) => {
        console.log(`[Payment] [${new Date().toISOString()}] ${message}`, data);
    };
    
    // ExcelJSの初期化チェック
    const getXLSX = () => {
        return new Promise((resolve, reject) => {
            if (typeof ExcelJS !== 'undefined') {
                resolve(ExcelJS);
            } else {
                reject(new Error('ExcelJSライブラリが見つかりません'));
            }
        });
    };

    // シート1用のヘッダー定義
    const getSheet1Headers = () => [
        "レコード番号",
        "作業日",
        "現場名",
        "顧客名",
        '職種_実績_支払',
        "郵便番号",
        "住所",
        "tel",
        "勤怠",
        "単価_実績_支払",
        "遅刻時間_支払い",
        "早退時間_支払い",
        "残業時間_支払い",
        "早出時間_支払い",
        "人工数_支払",
        "取引種別",
        "日勤_夜勤",
        "単価調整_支払_",
        "作業員検索",
        "計算式結果"
    ];

    // シート1へのデータ書き込み関数
    async function writeSheet1(workbook, records) {
        try {
            let worksheet = workbook.worksheets.find(sheet => sheet.name === 'Sheet1');
            if (!worksheet) {
                worksheet = workbook.addWorksheet('Sheet1');
            }
            
            // ヘッダー行の作成 - 適切な順序に修正
            const headers = [
                "レコード番号",
                "作業日",
                "現場名",
                "顧客名",
                '職種_実績_支払',
                "郵便番号",
                "住所",
                "tel",
                "勤怠",
                "単価_実績_支払",
                "遅刻時間_支払い",
                "早退時間_支払い",
                "残業時間_支払い",
                "早出時間_支払い",
                "人工数_支払",
                "取引種別",
                "経費種類_支払い",
                "単価_実績_支払",
                "日勤_夜勤",
                "単価調整_支払_",  
                "作業員検索",
                "計算式結果"
            ];
            
            worksheet.addRow(headers);
            
            let rowIndex = 2; // ヘッダー行が1行目なので、データは2行目から始まる
            
            records.forEach(record => {
                const expensesTable = record["経費_支払い"].value || [];
                
                if (expensesTable.length > 0) {
                    expensesTable.forEach(expense => {
                        const dataRow = [
                            record["レコード番号"].value,
                            new Date(record["作業日"].value),
                            record["現場名"].value,
                            record["顧客名"].value,
                            record["職種_実績_支払"].value,
                            record["郵便番号"].value,
                            record["住所"].value,
                            record["tel"].value,
                            record["勤怠"].value,
                            Number(record["単価_実績_支払"].value),
                            Number(record["遅刻時間_支払い"].value),
                            Number(record["早退時間_支払い"].value),
                            Number(record["残業時間_支払い"].value),
                            Number(record["早出時間_支払い"].value),
                            Number(record["人工数_支払"].value),
                            record["取引種別"].value,
                            expense.value["経費種類_支払い"]?.value || '',
                            Number(expense.value["単価_実績_支払"]?.value || 0),
                            record["日勤_夜勤"].value,
                            Number(record["単価調整_支払_"]?.value || 0),
                            record["作業員検索"].value
                        ];
                        
                        worksheet.addRow(dataRow);
                        
                        // W列にフォーミュラを設定
                        const formulaCell = worksheet.getCell(`W${rowIndex}`);
                        formulaCell.value = {
                            formula: `=J${rowIndex}/8*(M${rowIndex}+N${rowIndex})*1.25`,
                            date1904: false
                        };
                        
                        rowIndex++;
                    });
                } else {
                    const dataRow = [
                        record["レコード番号"].value,
                        new Date(record["作業日"].value),
                        record["現場名"].value,
                        record["顧客名"].value,
                        record["職種_実績_支払"].value,
                        record["郵便番号"].value,
                        record["住所"].value,
                        record["tel"].value,
                        record["勤怠"].value,
                        Number(record["単価_実績_支払"].value),
                        Number(record["遅刻時間_支払い"].value),
                        Number(record["早退時間_支払い"].value),
                        Number(record["残業時間_支払い"].value),
                        Number(record["早出時間_支払い"].value),
                        Number(record["人工数_支払"].value),
                        record["取引種別"].value,
                        '',  // 経費種類
                        0,   // 単価を数値0として設定
                        record["日勤_夜勤"].value,
                        Number(record["単価調整_支払_"]?.value || 0),
                        record["作業員検索"].value
                    ];
                    
                    worksheet.addRow(dataRow);
                    
                    // W列にフォーミュラを設定
                    const formulaCell = worksheet.getCell(`W${rowIndex}`);
                    formulaCell.value = {
                        formula: `=J${rowIndex}/8*(M${rowIndex}+N${rowIndex})*1.25`,
                        date1904: false
                    };
                    
                    rowIndex++;
                }
            });
            
            return workbook;
        } catch (error) {
            console.error('Sheet1への書き込みでエラーが発生しました:', error);
            throw error;
        }
    }

    

    // 支払い通知書シートへのデータ書き込み関数
    async function writePaymentSheet(workbook, record) {
        try {
            let paymentSheet = workbook.worksheets.find(sheet => sheet.name === '支払い通知書');
            if (!paymentSheet) {
                paymentSheet = workbook.addWorksheet('支払い通知書');
            }
            
            const Name = record["person_name"].value;
            const company = record["company_name"].value;
            const bank_code = record["銀行コード"].value;
            const bank_name = record["銀行名"].value;
            const account_number = record["口座番号"].value;
            const account_name = record["口座名義"].value;
            const branch_name = record["支店名"].value;
            const branch_code = record["支店番号"].value;
            const account_type = record["口座種別"].value;
            const tel = record["固定番号_作業員"].value;
            const postalCode = record["郵便番号_作業員"].value;
            const address_work = record["住所_作業員"].value;
            const buildingName = record["建物名_作業員"].value;
            
            //お支払い通知書
            paymentSheet.getCell('C2').value = Name;
            paymentSheet.getCell('C5').value = postalCode;
            paymentSheet.getCell('C7').value = buildingName;
            paymentSheet.getCell('C6').value = address_work;
            paymentSheet.getCell('C8').value = tel;  
            paymentSheet.getCell('C40').value = bank_name;
            paymentSheet.getCell('C41').value = branch_name;
            paymentSheet.getCell('C42').value = account_number;
            paymentSheet.getCell('C43').value = account_name;
            paymentSheet.getCell('E41').value = branch_code;
            paymentSheet.getCell('E42').value = account_type;
            //請求書
            paymentSheet.getCell('O2').value = Name;        // C2 → O2 (+12)
            paymentSheet.getCell('O5').value = postalCode;  // C5 → O5 (+12)
            paymentSheet.getCell('O7').value = buildingName; // C7 → O7 (+12)
            paymentSheet.getCell('O6').value = address_work; // C6 → O6 (+12)
            paymentSheet.getCell('O8').value = tel;         // C8 → O8 (+12)
            paymentSheet.getCell('O40').value = bank_name;  // C40 → O40 (+12)
            paymentSheet.getCell('O41').value = branch_name; // C41 → O41 (+12)
            paymentSheet.getCell('O42').value = account_number; // C42 → O42 (+12)
            paymentSheet.getCell('O43').value = account_name; // C43 → O43 (+12)
            paymentSheet.getCell('Q41').value = branch_code; // E41 → Q41 (+12)
            paymentSheet.getCell('Q42').value = account_type; // E42 → Q42 (+12)
            //相殺明細書
            paymentSheet.getCell('AA2').value = Name;        // C2 → AA2 (+24)
            paymentSheet.getCell('AA5').value = postalCode;  // C5 → AA5 (+24)
            paymentSheet.getCell('AA7').value = buildingName; // C7 → AA7 (+24)
            paymentSheet.getCell('AA6').value = address_work; // C6 → AA6 (+24)
            paymentSheet.getCell('AA8').value = tel;         // C8 → AA8 (+24)
            paymentSheet.getCell('AA40').value = bank_name;  // C40 → AA40 (+24)
            paymentSheet.getCell('AA41').value = branch_name; // C41 → AA41 (+24)
            paymentSheet.getCell('AA42').value = account_number; // C42 → AA42 (+24)
            paymentSheet.getCell('AA43').value = account_name; // C43 → AA43 (+24)
            paymentSheet.getCell('AC41').value = branch_code; // E41 → AC41 (+24)
            paymentSheet.getCell('AC42').value = account_type; // E42 → AC42 (+24)
            
            return workbook;
        } catch (error) {
            console.error('支払い通知書シートの書き込みでエラーが発生しました:', error);
            throw error;
        }
    }

    async function generatePaymentReport(eventRecord, templateFile) {
        try {
            const ExcelJS = await getXLSX();
            const workbook = new ExcelJS.Workbook();
            await workbook.xlsx.load(templateFile);
            
            const Name = eventRecord["person"].value;
            const startDate = eventRecord["開始日"].value;
            const endDate = eventRecord["終了日"].value;
            const queryCategory = eventRecord["query_category"].value;
            
            // appid24からレコードを取得
            const client = new KintoneRestAPIClient();
            
            // query_categoryに応じて条件を分岐
            let condition;
            switch (queryCategory) {
                case "常用":
                    condition = `作業員検索 = "${Name}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_支払 != 0 and 取引種別 in("常用")`;
                    break;
                case "請負(自)":
                    condition = `作業員検索 = "${Name}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_支払 != 0 and 取引種別 in("請負(自)")`;
                    break;
                case "請負(他)":
                    condition = `作業員検索 = "${Name}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_支払 != 0 and 取引種別 in("請負(他)")`;
                    break;
                case "すべて":
                default:
                    condition = `作業員検索 = "${Name}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_支払 != 0`;
                    break;
            }
            
            debugLog('レコード取得を試みます', { condition });
            const records = await client.record.getAllRecords({
                app: 24,
                condition: condition,
                fields: [
                    "レコード番号",
                    "作業日",
                    "現場名",
                    "顧客名",
                    "職種_実績_支払",
                    "郵便番号",
                    "住所",
                    "tel",
                    "勤怠",
                    "単価_実績_支払",
                    "遅刻時間_支払い",
                    "早退時間_支払い",
                    "残業時間_支払い",
                    "早出時間_支払い",
                    "人工数_支払",
                    "取引種別",
                    "経費_支払い",
                    "日勤_夜勤",
                    "単価調整_支払_",
                    "作業員検索"
                ]
            });
            
            if (records.length === 0) {
                alert("指定された期間内に該当するデータがありません。\n以下の条件を確認してください：\n・顧客名\n・作業日（開始日～終了日）\n・取引種別\n・人工数_支払が0以外");
                throw new Error("該当するデータがありません");
            }
            
            console.log('取得されたレコード数:', records.length);
            
            // 各シートへの書き込み
            await writeSheet1(workbook, records);
            await writePaymentSheet(workbook, eventRecord);
            
            return await workbook.xlsx.writeBuffer();
        } catch (error) {
            console.error('Excelファイルの生成でエラーが発生しました:', error);
            throw error;
        }
    }

    // テンプレートファイル取得関数
    async function getTemplateFile(app, recordNumber) {
        try {
            const body = {
                app: app,
                query: `レコード番号 = "${recordNumber}"`
            };
            
            const response = await kintone.api(kintone.api.url("/k/v1/records", true), "GET", body);
            
            if (!response.records || response.records.length === 0) {
                debugLog('テンプレートレコードが見つかりませんでした', { app, recordNumber });
                throw new Error("テンプレートレコードが見つかりません");
            }
            
            const record = response.records[0];
            const fileKey = record.添付ファイル.value?.[0]?.fileKey;
            
            if (!fileKey) {
                debugLog('添付ファイルが見つかりませんでした', { record });
                throw new Error("添付ファイルが見つかりません");
            }
            
            const fileUrl = kintone.api.urlForGet("/k/v1/file", { fileKey }, true);
            const responseArrayBuffer = await fetch(fileUrl, {
                headers: {
                    'X-Requested-With': 'XMLHttpRequest'
                }
            });
            
            if (!responseArrayBuffer.ok) {
                debugLog('テンプレートファイルの取得に失敗しました', { status: responseArrayBuffer.status });
                throw new Error(`テンプレートファイルの取得に失敗しました: ${responseArrayBuffer.status}`);
            }
            
            const arrayBuffer = await responseArrayBuffer.arrayBuffer();
            const view = new Uint8Array(arrayBuffer);
            debugLog('テンプレートファイルが正常に取得されました', { fileSize: arrayBuffer.byteLength });
            
            return view.buffer;
        } catch (error) {
            debugLog('エラーが発生しました:', error);
            throw error;
        }
    }

    // ファイルダウンロード関数
    async function downloadExcelFile(data, fileName) {
        try {
            debugLog('Excelファイルのダウンロードを開始します', { fileName, dataSize: data.byteLength });
            const blob = new Blob([data], {
                type: 'application/vnd.openxmlformats-offreedocument.spreadsheetml.sheet'
            });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = fileName;
            const clickEvent = new MouseEvent('click', {
                bubbles: true,
                cancelable: true,
                view: window
            });
            a.dispatchEvent(clickEvent);
            
            setTimeout(() => {
                window.URL.revokeObjectURL(url);
                debugLog('URLが正常に解放されました');
            }, 1000);
        } catch (error) {
            debugLog('ファイルダウンロードでエラーが発生しました:', error);
            throw new Error('Excelファイルのダウンロードに失敗しました');
        }
    }

    // 初期化関数
    const initializePaymentHandler = () => {
        try {
            debugLog('支払い通知書ハンドラーの初期化を開始します');
            
            // 既存のイベントリスナーを削除
            kintone.events.off('app.record.detail.show', handlePaymentDetailShow);
            
            // 新しいイベントリスナーを登録
            kintone.events.on('app.record.detail.show', handlePaymentDetailShow);
            
            debugLog('支払い通知書ハンドラーの初期化が完了しました');
        } catch (error) {
            debugLog('支払い通知書ハンドラーの初期化でエラーが発生しました:', error);
        }
    };

    // キントーンのレコード詳細表示イベントに応答
    const handlePaymentDetailShow = function(event) {
        try {
            debugLog('支払い通知書のレコード詳細表示イベントが発火しました', {
                recordId: event.recordId,
                record: event.record
            });
            
            // out_categoryが「支払い通知書」でない場合は処理を終了
            if (!event.record || !event.record["out_category"]) {
                debugLog('out_categoryが存在しません', event.record);
                return;
            }

            if (event.record["out_category"].value !== "支払い通知書") {
                debugLog('支払い通知書以外のレコードです', {
                    out_category: event.record["out_category"].value
                });
                return;
            }

            // 既にボタンが存在する場合は処理を終了
            const existingButton = document.getElementById('export_payment_button');
            if (existingButton) {
                debugLog('支払い通知書出力ボタンは既に存在します');
                return;
            }
            
            const button = document.createElement('button');
            button.id = 'export_payment_button';
            button.innerText = '支払い通知書出力';
            button.style.margin = '10px';
            button.style.backgroundColor = '#4CAF50';
            button.style.color = 'white';
            button.style.border = 'none';
            button.style.padding = '8px 16px';
            button.style.borderRadius = '4px';
            button.style.cursor = 'pointer';
            button.style.fontSize = '14px';
            button.style.fontWeight = 'bold';
            button.style.transition = 'all 0.3s ease';
            button.style.boxShadow = '0 2px 4px rgba(0,0,0,0.2)';
            
            // ホバー効果
            button.onmouseover = function() {
                this.style.backgroundColor = '#45a049';
                this.style.transform = 'translateY(-1px)';
                this.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
            };
            
            button.onmouseout = function() {
                this.style.backgroundColor = '#4CAF50';
                this.style.transform = 'translateY(0)';
                this.style.boxShadow = '0 2px 4px rgba(0,0,0,0.2)';
            };
            
            button.onclick = async function() {
                try {
                    const record = event.record;
                    const fileName = `（支払い通知書）${record["顧客名"].value}_${record["開始日"].value}.xlsx`;
                    
                    debugLog('テンプレートファイルの取得を試みます');
                    const templateFile = await getTemplateFile(31, 7);
                    
                    debugLog('Excelデータの書き込みを試みます');
                    const excelBuffer = await generatePaymentReport(record, templateFile);
                    
                    debugLog('ファイルダウンロードを試みます', { fileName });
                    await downloadExcelFile(excelBuffer, fileName);
                    
                    debugLog('処理が正常に完了しました');
                    alert("支払い通知書の出力が完了しました。");
                } catch (error) {
                    debugLog('エラーが発生しました:', error);
                    alert("処理中にエラーが発生しました。\n管理者にご連絡ください。");
                }
            };
            
            const headerMenuSpace = kintone.app.record.getHeaderMenuSpaceElement();
            if (headerMenuSpace) {
                headerMenuSpace.appendChild(button);
                debugLog('支払い通知書出力ボタンを追加しました');
            } else {
                debugLog('ヘッダーメニュー要素が見つかりません');
            }
        } catch (error) {
            debugLog('イベントハンドラでエラーが発生しました:', error);
        }
    };

    // 初期化を実行
    debugLog('支払い通知書スクリプトの読み込みを開始します');
    initializePaymentHandler();
})(); 