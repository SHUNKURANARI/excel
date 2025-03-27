(function() {
    'use strict';
    
    // デバッグ用のログ出力関数
  const debugLog = (message, data) => {
    console.log(`[${new Date().toISOString()}] ${message}`, data);
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
    '職種_実績_請求',
    "郵便番号",
    "住所",
    "tel",
    "勤怠",
    "単価",
    "遅刻時間_請求",
    "早退時間_請求",
    "残業時間_請求",
    "早出時間_請求",
    "人工数_請求",
    "取引種別",
    "日勤_夜勤",
    "単価調整_請求"
  ];
  
  // シート2用のヘッダー定義
  // const getSheet2Headers = () => [
  //     "レコード番号",
  //     "作業日",
  //     "顧客名",
  //     "経費種類",
  //     "請求経費",
  //     "金額_経費"
  // ];
  
  // シート1へのデータ書き込み関数（更新版）
  async function writeSheet1(workbook, records) {
    try {
        let worksheet = workbook.worksheets.find(sheet => sheet.name === 'Sheet1');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('Sheet1');
        }
        
        // ヘッダー行の作成（経費列を追加）
        const headers = [
            "レコード番号",
            "作業日",
            "現場名",
            '職種_実績_請求種',
            "顧客名",
            "郵便番号",
            "住所",
            "tel",
            "勤怠",
            "単価",
            "遅刻時間_請求",
            "早退時間_請求",
            "残業時間_請求",
            "早出時間_請求",
            "人工数_請求",
            "取引種別",
            // 経費関連列を追加
            "経費種類",
            "請求経費",
            "金額_経費",
            "日勤_夜勤",
            "単価調整_請求"
        ];
        
        worksheet.addRow(headers);
        
        records.forEach(record => {
            const expensesTable = record["経費_請求"].value || [];
            
            // サブテーブルのデータが存在する場合のみ処理
            if (expensesTable.length > 0) {
                expensesTable.forEach(expense => {
                    worksheet.addRow([
                        record["レコード番号"].value,
                        new Date(record["作業日"].value),
                        record["現場名"].value,
                        record["顧客名"].value,
                        record["職種_実績_請求"].value,
                        record["郵便番号"].value,
                        record["住所"].value,
                        record["tel"].value,
                        record["勤怠"].value,
                        Number(record["単価"].value),
                        Number(record["遅刻時間_請求"].value),
                        Number(record["早退時間_請求"].value),
                        Number(record["残業時間_請求"].value),
                        Number(record["早出時間_請求"].value),
                        Number(record["人工数_請求"].value),
                        record["取引種別"].value,
                        expense.value["経費種類_支払い"]?.value || '',
                        Number(expense.value["単価_実績_支払"]?.value || 0),
                        Number(expense.value["金額_経費"]?.value || 0),
                        record["日勤_夜勤"].value,
                        Number(record["単価調整_請求"]?.value || 0)
                    ]);
                });
            } else {
                // サブテーブルデータがない場合、基本情報のみを追加
                worksheet.addRow([
                    record["レコード番号"].value,
                    new Date(record["作業日"].value),
                    record["現場名"].value,
                    record["顧客名"].value,
                    record["職種_実績_請求"].value,
                    record["郵便番号"].value,
                    record["住所"].value,
                    record["tel"].value,
                    record["勤怠"].value,
                    Number(record["単価"].value),
                    Number(record["遅刻時間_請求"].value),
                    Number(record["早退時間_請求"].value),
                    Number(record["残業時間_請求"].value),
                    Number(record["早出時間_請求"].value),
                    Number(record["人工数_請求"].value),
                    record["取引種別"].value,
                    record["日勤_夜勤"].value,
                    '', '', '',
                    Number(record["単価調整_請求"]?.value || 0)
                ]);
            }
        });
        
        return workbook;
    } catch (error) {
        console.error('Sheet1への書き込みでエラーが発生しました:', error);
        throw error;
    }
  }
  
  async function writeSheet2(workbook, records) {
    try {
        let worksheet = workbook.worksheets.find(sheet => sheet.name === 'Sheet2');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('Sheet2');
        }
        worksheet.addRow(getSheet1Headers());
        
        records.forEach((record, index) => {
            const rowNumber = index + 2; // ヘッダー行を考慮
            
            worksheet.addRow([
                record["レコード番号"].value,
                new Date(record["作業日"].value),
                record["現場名"].value,
                record["顧客名"].value,
                record["職種_実績_請求"].value,
                record["郵便番号"].value,
                record["住所"].value,
                record["tel"].value,
                record["勤怠"].value,
                Number(record["単価"].value),
                Number(record["遅刻時間_請求"].value),
                Number(record["早退時間_請求"].value),
                Number(record["残業時間_請求"].value),
                Number(record["早出時間_請求"].value),
                Number(record["人工数_請求"].value),
                record["取引種別"].value,
                record["日勤_夜勤"].value,
                Number(record["単価調整_請求"]?.value || 0),
                { formula: `=J${rowNumber}/8*M${rowNumber}*1.25`, value: null },
                { formula: `=J${rowNumber}/8*N${rowNumber}*1.25`, value: null }
            ]);
        });
        
        return workbook;
    } catch (error) {
        console.error('Sheet2への書き込みでエラーが発生しました:', error);
        throw error;
    }
  }
  // データ書き込み関数
  // async function writeSheet2(workbook, records, customerName, startDate, endDate) {
  //     try {
  //         let expenseSheet = workbook.worksheets.find(sheet => sheet.name === 'Sheet2');
  //         if (!expenseSheet) {
  //             expenseSheet = workbook.addWorksheet('Sheet2');
  //         }
        
  //         // ヘッダー行の追加
  //         expenseSheet.addRow(getSheet2Headers());
        
  //         // データの書き込み処理を改善
  //         records.forEach(record => {
  //             const expensesTable = record["経費_請求"].value || [];
            
  //             // 顧客名チェックを削除（メインレコードで既にフィルタリング済み）
  //             expensesTable.forEach(expense => {
  //                 const row = [
  //                     record["$id"].value,
  //                     record["作業日"].value,
  //                     record["顧客名"].value,
  //                     expense.value["経費種類"]?.value || '',
  //                     expense.value["請求経費"]?.value || false,
  //                     expense.value["金額_経費"]?.value || 0
  //                 ];
                
  //                 // デバッグ用ログ出力
  //                 debugLog('Sheet2に書き込むデータ:', row);
  //                 expenseSheet.addRow(row);
  //             });
  //         });
        
  //         return workbook;
  //     } catch (error) {
  //         console.error('Sheet2への書き込みでエラーが発生しました:', error);
  //         throw error;
  //     }
  // }
  // 請求書シートへのデータ書き込み関数
  async function writeInvoiceSheet(workbook, record) {
    try {
        let invoiceSheet = workbook.worksheets.find(sheet => sheet.name === '請求書');
        if (!invoiceSheet) {
            invoiceSheet = workbook.addWorksheet('請求書');
        }
        
        const customerName = record["顧客名"].value;
        const postalCode = record["郵便番号"].value;
        const address = record["住所"].value;
        const buildingName = record["ビル名"]?.value || '';
        const tel = record["tel"].value;
        const fax = record["fax"]?.value || '';
        const claim_date = record["請求日"].value;
        const serial_number = record["請求管理番号"].value;
        const payment_cicle = record["支払いサイクル"].value;
        const payment_date = record["締め日"].value;
        const payment_deadline = record["支払い期限"].value;
        const end_date = record["終了日"].value;
  
        
        invoiceSheet.getCell('C2').value = customerName;
        invoiceSheet.getCell('C5').value = postalCode;
        invoiceSheet.getCell('C6').value = address;
        invoiceSheet.getCell('C7').value = buildingName;
        invoiceSheet.getCell('C8').value = tel;
        invoiceSheet.getCell('C9').value = fax;
        invoiceSheet.getCell('I3').value = claim_date;
        invoiceSheet.getCell('I2').value = serial_number;
        invoiceSheet.getCell('J40').value = Number(payment_cicle);
        invoiceSheet.getCell('I40').value = payment_date;
        invoiceSheet.getCell('H40').value = end_date;
        const dateValue = new Date(payment_deadline); 
        invoiceSheet.getCell('C41').value = dateValue;
        invoiceSheet.getCell("C41").numFmt = '[$-ja-JP]ggge年m月d日';
        return workbook;
    } catch (error) {
        console.error('請求書シートの書き込みでエラーが発生しました:', error);
        throw error;
    }
  }
  
  
  async function writeWorkSiteSheet(workbook, records) {
    try {
        console.log('writeWorkSiteSheetが呼び出されました。レコード数:', records.length);
        // シートの取得または作成
        let worksheet = workbook.worksheets.find(sheet =>
            sheet.name === '【請求明細】現場職種別');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('【請求明細】現場職種別');
            
            // 列定義を更新（列番号で指定）
            worksheet.columns = [
                { header: '', width: 10 },           // A列（項番用）
                { header: '作業日', width: 12 },     // B列
                { header: '現場名', width: 20 },     // C列
                { header: '日勤_夜勤', width: 12 },  // D列
                { header: '職種_実績_請求', width: 15 },   // E列
                { header: '単価', width: 10 },       // F列
                { header: '人工数_請求', width: 15 },// G列
                { header: '早出残業時間', width: 10 },           // H列（早出残業合計）
                { header: '単価', width: 10 },       // I列
                { header: '早出時間_請求 - 残業時間_請求', width: 25 },// J列
                { header: '', width: 10 },           // K列（空）
                { header: '', width: 10 },           // L列（新規追加）
                { header: '電車・バス', width: 12 },  // M列
                { header: '飛行機', width: 12 },     // N列
                { header: 'タクシー', width: 12 },    // O列
                { header: '車両', width: 12 },       // P列
                { header: '', width: 12 },           // Q列（交通費合計）
                { header: '消耗品', width: 12 },     // R列
                { header: '道具持込', width: 12 },   // S列
                { header: 'リース代', width: 12 },    // T列
                { header: 'その他', width: 12 },      // U列
                { header: '', width: 15 },           // V列（経費合計）
                { header: '', width: 10 }            // W列（最終合計）
            ];
        }
  
        // 新しい集計方法 - 現場職種別データを先に集計
        console.log('現場職種別に時間とコストを集計します');
        const groupedData = {};
        
        // 最初のパスで現場職種別のデータを収集
        records.forEach((record, recordIndex) => {
            // グループ化キーを生成: 作業日|現場名|日勤_夜勤|職種|単価
            const key = `${record["作業日"].value}|${record["現場名"].value}|${record["日勤_夜勤"].value}|${record["職種_実績_請求"].value}|${record["単価"].value}`;
            const expenses = record["経費_請求"].value || [];
            
            console.log(`レコード${recordIndex+1}を処理: キー=${key}`);
            console.log(`  早出時間=${record["早出時間_請求"]?.value || 0}, 残業時間=${record["残業時間_請求"]?.value || 0}`);
            
            if (!groupedData[key]) {
                groupedData[key] = {
                    作業日: record["作業日"].value,
                    現場名: record["現場名"].value,
                    日勤_夜勤: record["日勤_夜勤"].value,
                    職種_実績_請求: record["職種_実績_請求"].value,
                    単価: Number(record["単価"].value),
                    人工数_請求: 0,
                    早出時間: 0,
                    残業時間: 0,
                    単価調整: Number(record["単価調整_請求"]?.value || 0),
                    電車バス: 0,
                    飛行機: 0,
                    タクシー: 0,
                    車両: 0,
                    消耗品: 0,
                    道具持込: 0,
                    リース代: 0,
                    その他: 0
                };
                console.log('  新しいグループを作成:', key);
            }
            
            // 人工数と時間の集計
            groupedData[key].人工数_請求 += 1;
            groupedData[key].早出時間 += Number(record["早出時間_請求"]?.value || 0);
            groupedData[key].残業時間 += Number(record["残業時間_請求"]?.value || 0);
            
            console.log(`  現在の集計: 早出=${groupedData[key].早出時間}, 残業=${groupedData[key].残業時間}`);
            
            // 経費の集計
            expenses.forEach(expense => {
                const 経費種類 = expense.value["経費種類"]?.value || '';
                const 金額 = Number(expense.value["金額_経費"]?.value || 0);
                
                switch (経費種類) {
                    case "電車・バス": groupedData[key].電車バス += 金額; break;
                    case "飛行機": groupedData[key].飛行機 += 金額; break;
                    case "タクシー": groupedData[key].タクシー += 金額; break;
                    case "車両": groupedData[key].車両 += 金額; break;
                    case "消耗品": groupedData[key].消耗品 += 金額; break;
                    case "道具持込": groupedData[key].道具持込 += 金額; break;
                    case "リース代": groupedData[key].リース代 += 金額; break;
                    case "その他": groupedData[key].その他 += 金額; break;
                }
            });
        });
        
        // 集計結果をログ出力
        console.log('現場職種別集計結果:', groupedData);
        
        // 集計データをシートに書き込み
        const groupedEntries = Object.values(groupedData);
        console.log('グループ化されたエントリ数:', groupedEntries.length);
        const startingRow = 4; // ヘッダー行の後に開始
        
        groupedEntries.forEach((entry, index) => {
            const rowNumber = startingRow + index;
            console.log(`行${rowNumber}に書き込み:`, entry);

            // A列に項番を設定（1から始まる連番）
            worksheet.getCell(`A${rowNumber}`).value = index + 1;  // 項番を1から開始
            worksheet.getCell(`A${rowNumber}`).alignment = { horizontal: 'center' }; // 中央揃え

            // 他の列の設定
            worksheet.getCell(`B${rowNumber}`).value = new Date(entry.作業日);
            worksheet.getCell(`B${rowNumber}`).numFmt = 'm"月"d"日"';
            worksheet.getCell(`C${rowNumber}`).value = entry.現場名;
            worksheet.getCell(`D${rowNumber}`).value = entry.日勤_夜勤;
            worksheet.getCell(`E${rowNumber}`).value = entry.職種_実績_請求;
            worksheet.getCell(`F${rowNumber}`).value = entry.単価;
            worksheet.getCell(`G${rowNumber}`).value = entry.人工数_請求;
            worksheet.getCell(`G${rowNumber}`).numFmt = '#,##0';
            
            // 早出時間と残業時間の合計を設定
            const extraTimeTotal = entry.早出時間 + entry.残業時間;
            console.log(`${entry.現場名}/${entry.職種_実績_請求}の早出残業合計=${extraTimeTotal}`);
            worksheet.getCell(`H${rowNumber}`).value = extraTimeTotal;
            worksheet.getCell(`H${rowNumber}`).numFmt = '#,##0.0';
            
            worksheet.getCell(`I${rowNumber}`).value = {
                formula: `=F${rowNumber}*G${rowNumber}`,
                date1904: false
            };
            worksheet.getCell(`I${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`J${rowNumber}`).value = {
                formula: `=F${rowNumber}*H${rowNumber}/8*1.25`,
                date1904: false
            };
            worksheet.getCell(`J${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`K${rowNumber}`).value = entry.単価調整;
            worksheet.getCell(`K${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`L${rowNumber}`).value = ''; // 新規追加列
            worksheet.getCell(`M${rowNumber}`).value = entry.電車バス;
            worksheet.getCell(`M${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`N${rowNumber}`).value = entry.飛行機;
            worksheet.getCell(`N${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`O${rowNumber}`).value = entry.タクシー;
            worksheet.getCell(`O${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`P${rowNumber}`).value = entry.車両;
            worksheet.getCell(`P${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`Q${rowNumber}`).value = ''; // 交通費合計列
            worksheet.getCell(`Q${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`R${rowNumber}`).value = entry.消耗品;
            worksheet.getCell(`R${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`S${rowNumber}`).value = entry.道具持込;
            worksheet.getCell(`S${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`T${rowNumber}`).value = entry.リース代;
            worksheet.getCell(`T${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`U${rowNumber}`).value = entry.その他;
            worksheet.getCell(`U${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`V${rowNumber}`).value = ''; // 経費合計列
            worksheet.getCell(`V${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`W${rowNumber}`).value = ''; // 最終合計列
            worksheet.getCell(`W${rowNumber}`).numFmt = '#,##0';

            // 既存の計算式の設定
            worksheet.getCell(`Q${rowNumber}`).value = {
                formula: `=SUM(M${rowNumber}:P${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`Q${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`V${rowNumber}`).value = {
                formula: `=SUM(R${rowNumber}:U${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`V${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`L${rowNumber}`).value = {
                formula: `=SUM(I${rowNumber}:K${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`L${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`W${rowNumber}`).value = {
                formula: `=(L${rowNumber}+Q${rowNumber}+V${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`W${rowNumber}`).numFmt = '#,##0';
        });

        // 最終行の処理
        const totalRow = startingRow + groupedEntries.length;
        
        // A列からF列までのセルを結合して「合計」を入力
        worksheet.mergeCells(`A${totalRow}:F${totalRow}`);
        worksheet.getCell(`A${totalRow}`).value = '合計';
        worksheet.getCell(`A${totalRow}`).alignment = { horizontal: 'center' };

        // 最終行の合計式設定（G～W列）
        ['G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W'].forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            cell.value = {
                formula: `=SUM(${column}${startingRow}:${column}${totalRow-1})`,
                date1904: false
            };
            cell.numFmt = '#,##0';
        });

        // 単価調整_請求の合計を計算
        const totalPriceAdjustment = records.reduce((sum, record) => {
            return sum + Number(record["単価調整_請求"]?.value || 0);
        }, 0);

        // 単価調整_請求の合計を最終行のK列に設定
        worksheet.getCell(`K${totalRow}`).value = totalPriceAdjustment;
        worksheet.getCell(`K${totalRow}`).numFmt = '#,##0';

        // 罫線スタイルの定義
        const borderStyle = {
            style: 'thin',
            color: { argb: '000000' }
        };

        const doubleBorderStyle = {
            style: 'double',
            color: { argb: '000000' }
        };

        // データ行の罫線設定（最終行を除く）
        for (let row = startingRow; row < totalRow; row++) {
            ['A','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W'].forEach(column => {
                const cell = worksheet.getCell(`${column}${row}`);
                cell.border = {
                    top: borderStyle,
                    bottom: borderStyle,
                    left: borderStyle,
                    right: borderStyle
                };
            });
        }

        // 最終行の特別な罫線設定と着色（W列まで）
        ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W'].forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            // 罫線設定
            cell.border = {
                top: doubleBorderStyle,
                bottom: borderStyle,
                left: borderStyle,
                right: borderStyle
            };
            // 着色設定
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFA500' }
            };
            // 書式設定を保持
            if (cell.value && typeof cell.value === 'object' && cell.value.formula) {
                const formula = cell.value.formula;
                cell.value = { formula: formula, date1904: false };
            }
        });

        return workbook;
    } catch (error) {
        console.error('【請求明細】現場職種別シートの処理でエラーが発生しました:', error);
        throw error;
    }
  }
  
  async function writeWorkSiteSheet2(workbook, records) {
    try {
        console.log('writeWorkSiteSheet2が呼び出されました。レコード数:', records.length);
        // シートの取得または作成
        let worksheet = workbook.worksheets.find(sheet =>
            sheet.name === '【請求明細】現場別');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('【請求明細】現場別');
            
            // 列定義を更新（列番号で指定）
            worksheet.columns = [
                { header: '', width: 10 },           // A列（項番用）
                // { header: '作業日', width: 12 },     // B列
                { header: '現場名', width: 20 },     // C列
                // { header: '日勤_夜勤', width: 12 },  // D列
                // { header: '職種_実績_請求', width: 15 },   // E列
                { header: '単価', width: 10 },       // F列
                { header: '人工数_請求', width: 15 },// G列
                { header: '早出残業時間', width: 12 },           // H列（早出残業時間）
                { header: '単価', width: 10 },       // I列
                { header: '早出時間_請求 - 残業時間_請求', width: 25 },// J列
                { header: '', width: 10 },           // K列（空）
                { header: '', width: 10 },           // L列（新規追加）
                { header: '電車・バス', width: 12 },  // M列
                { header: '飛行機', width: 12 },     // N列
                { header: 'タクシー', width: 12 },    // O列
                { header: '車両', width: 12 },       // P列
                { header: '', width: 12 },           // Q列（交通費合計）
                { header: '消耗品', width: 12 },     // R列
                { header: '道具持込', width: 12 },   // S列
                { header: 'リース代', width: 12 },    // T列
                { header: 'その他', width: 12 },      // U列
                { header: '', width: 15 },           // V列（経費合計）
                { header: '', width: 10 }            // W列（最終合計）
            ];
        }
  
        // 新しい初期集計方法 - 現場別データを先に集計
        console.log('現場別に時間とコストを集計します');
        const siteData = {};
        
        // 最初のパスで現場ごとのデータを収集
        records.forEach((record, recordIndex) => {
            const siteName = record["現場名"].value;
            console.log(`レコード${recordIndex+1}を処理: 現場=${siteName}`);
            
            if (!siteData[siteName]) {
                // 現場データの初期化
                siteData[siteName] = {
                    現場名: siteName,
                    単価: 0,
                    人工数: 0,
                    早出時間: 0,
                    残業時間: 0,
                    単価調整: 0,
                    電車バス: 0,
                    飛行機: 0,
                    タクシー: 0,
                    車両: 0,
                    消耗品: 0,
                    道具持込: 0,
                    リース代: 0,
                    その他: 0
                };
            }
            
            // 基本データの集計
            siteData[siteName].単価 += Number(record["単価"].value || 0);
            siteData[siteName].人工数 += 1;
            siteData[siteName].早出時間 += Number(record["早出時間_請求"]?.value || 0);
            siteData[siteName].残業時間 += Number(record["残業時間_請求"]?.value || 0);
            siteData[siteName].単価調整 += Number(record["単価調整_請求"]?.value || 0);
            
            console.log(`  早出時間=${record["早出時間_請求"]?.value || 0}, 残業時間=${record["残業時間_請求"]?.value || 0}`);
            console.log(`  現在の集計: 早出=${siteData[siteName].早出時間}, 残業=${siteData[siteName].残業時間}`);
            
            // 経費の集計
            const expenses = record["経費_請求"]?.value || [];
            expenses.forEach(expense => {
                const 経費種類 = expense.value["経費種類"]?.value || '';
                const 金額 = Number(expense.value["金額_経費"]?.value || 0);
                
                switch (経費種類) {
                    case "電車・バス": siteData[siteName].電車バス += 金額; break;
                    case "飛行機": siteData[siteName].飛行機 += 金額; break;
                    case "タクシー": siteData[siteName].タクシー += 金額; break;
                    case "車両": siteData[siteName].車両 += 金額; break;
                    case "消耗品": siteData[siteName].消耗品 += 金額; break;
                    case "道具持込": siteData[siteName].道具持込 += 金額; break;
                    case "リース代": siteData[siteName].リース代 += 金額; break;
                    case "その他": siteData[siteName].その他 += 金額; break;
                }
            });
        });
        
        // 集計結果をログ出力
        console.log('現場別集計結果:', siteData);
        
        // 集計データからシートにデータを書き込み
        const siteSummaries = Object.values(siteData);
        const startingRow = 4; // ヘッダー行の後に開始
        
        siteSummaries.forEach((site, index) => {
            const rowNumber = startingRow + index;
            console.log(`行${rowNumber}に書き込み:`, site);
            
            // A列に項番を設定
            worksheet.getCell(`A${rowNumber}`).value = index + 1;
            worksheet.getCell(`A${rowNumber}`).alignment = { horizontal: 'center' };
            
            // データ行の書き込み
            worksheet.getCell(`B${rowNumber}`).value = site.現場名;
            worksheet.getCell(`C${rowNumber}`).value = site.人工数;
            
            // 早出時間と残業時間の合計をD列に設定（ただしここでは単一の値）
            const totalExtraTime = site.早出時間 + site.残業時間;
            console.log(`${site.現場名}の早出残業合計=${totalExtraTime}`);
            worksheet.getCell(`D${rowNumber}`).value = totalExtraTime;
            worksheet.getCell(`D${rowNumber}`).numFmt = '#,##0.0';
            
            worksheet.getCell(`E${rowNumber}`).value = site.単価;
            worksheet.getCell(`E${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`F${rowNumber}`).value = { formula: `=(E${rowNumber}/8*D${rowNumber}*1.25)`, date1904: false };
            worksheet.getCell(`F${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`G${rowNumber}`).value = site.単価調整;
            worksheet.getCell(`G${rowNumber}`).numFmt = '#,##0';
            
            // 経費データの設定
            worksheet.getCell(`I${rowNumber}`).value = site.電車バス;
            worksheet.getCell(`I${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`J${rowNumber}`).value = site.飛行機;
            worksheet.getCell(`J${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`K${rowNumber}`).value = site.タクシー;
            worksheet.getCell(`K${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`L${rowNumber}`).value = site.車両;
            worksheet.getCell(`L${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`N${rowNumber}`).value = site.消耗品;
            worksheet.getCell(`N${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`O${rowNumber}`).value = site.道具持込;
            worksheet.getCell(`O${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`P${rowNumber}`).value = site.リース代;
            worksheet.getCell(`P${rowNumber}`).numFmt = '#,##0';
            worksheet.getCell(`Q${rowNumber}`).value = site.その他;
            worksheet.getCell(`Q${rowNumber}`).numFmt = '#,##0';
            
            // 合計計算式の設定
            worksheet.getCell(`H${rowNumber}`).value = {
                formula: `=SUM(E${rowNumber}:G${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`H${rowNumber}`).numFmt = '#,##0';
            
            worksheet.getCell(`M${rowNumber}`).value = {
                formula: `=SUM(I${rowNumber}:L${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`M${rowNumber}`).numFmt = '#,##0';
            
            worksheet.getCell(`R${rowNumber}`).value = {
                formula: `=SUM(N${rowNumber}:Q${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`R${rowNumber}`).numFmt = '#,##0';
            
            worksheet.getCell(`S${rowNumber}`).value = {
                formula: `=(H${rowNumber}+M${rowNumber}+R${rowNumber})`,
                date1904: false
            };
            worksheet.getCell(`S${rowNumber}`).numFmt = '#,##0';
        });

        // 既存の最終行処理コードはそのまま残す
        const totalRow = startingRow + siteSummaries.length;
        
        // A列からB列までのセルを結合して「合計」を入力
        worksheet.mergeCells(`A${totalRow}:B${totalRow}`);
        const totalCell = worksheet.getCell(`A${totalRow}`);
        totalCell.value = '合計';
        totalCell.alignment = { horizontal: 'center' };

        // 最終行の合計式設定（C～S列）
        ['C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S'].forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            cell.value = {
                formula: `=SUM(${column}${startingRow}:${column}${totalRow-1})`,
                date1904: false
            };
            cell.numFmt = '#,##0';
        });

        // 単価調整_請求の合計を計算
        const totalPriceAdjustment = records.reduce((sum, record) => {
            return sum + Number(record["単価調整_請求"]?.value || 0);
        }, 0);

        // 単価調整_請求の合計を最終行のR列に設定
        worksheet.getCell(`R${totalRow}`).value = totalPriceAdjustment;
        worksheet.getCell(`R${totalRow}`).numFmt = '#,##0';

        // 罫線スタイルの定義
        const borderStyle = {
            style: 'thin',
            color: { argb: '000000' }
        };

        const doubleBorderStyle = {
            style: 'double',
            color: { argb: '000000' }
        };

        // データ行の罫線設定（最終行を除く）
        for (let row = startingRow; row < totalRow; row++) {
            ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S'].forEach(column => {
                const cell = worksheet.getCell(`${column}${row}`);
                cell.border = {
                    top: borderStyle,
                    bottom: borderStyle,
                    left: borderStyle,
                    right: borderStyle
                };
            });
        }

        // 最終行の特別な罫線設定と着色（S列まで）
        ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S'].forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            cell.border = {
                top: doubleBorderStyle,
                bottom: borderStyle,
                left: borderStyle,
                right: borderStyle
            };
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFA500' }
            };
        });

        return workbook;
    } catch (error) {
        console.error('【請求明細】現場別シートの処理でエラーが発生しました:', error);
        throw error;
    }
  }
  
  async function writeBillingSheet(workbook, records) {
    try {
        // シートの取得または作成
        let worksheet = workbook.worksheets.find(sheet => 
            sheet.name === '【請求】現場別日別出面表');
        if (!worksheet) {
            worksheet = workbook.addWorksheet('【請求】現場別日別出面表');
        }
  
        // 列定義
        worksheet.columns = [
            { header: '現場名', key: 'site', width: 20 },    // A列
            { header: '職種', key: 'job', width: 15 },       // B列
            { header: '単価', key: 'rate', width: 10 },      // C列
            { header: '', width: 10 },                       // D列（空）
            { header: '', width: 10 },                       // E列（空）
            { header: '', width: 10 },                       // F列（空）
            { header: '', width: 10 },                       // G列（空）
            { header: '', width: 10 },                       // H列（空）
            { header: '', width: 10 }                        // I列（空）
        ];
  
        // データのグループ化
        const groupedData = records.reduce((acc, record) => {
            const key = `${record["現場名"].value}|${record["職種_実績_請求"].value}`;
            if (!acc[key]) {
                acc[key] = {
                    現場名: record["現場名"].value,
                    職種: record["職種_実績_請求"].value,
                    単価: Number(record["単価"].value)
                };
            }
            return acc;
        }, {});
  
        // データをシートに書き込み（4行目から開始）
        const groupedEntries = Object.values(groupedData);
        const startingRow = 4;
  
        // データ行の書き込み
        groupedEntries.forEach((entry, index) => {
            const rowNumber = startingRow + index;
            worksheet.getCell(`A${rowNumber}`).value = entry.現場名;
            worksheet.getCell(`B${rowNumber}`).value = entry.職種;
            worksheet.getCell(`C${rowNumber}`).value = entry.単価;
        });
  
        // SUMIFS式の設定（D列からAI列）
        const columns = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI','AJ'];
        groupedEntries.forEach((entry, index) => {
            const rowNumber = startingRow + index;
            columns.forEach((column, colIndex) => {
                worksheet.getCell(`${column}${rowNumber}`).value = {
                    formula: `=COUNTIFS(Sheet2!$B$2:$B$4000, ${column}$2, Sheet2!$C$2:$C$4000, $A${rowNumber}, Sheet2!$E$2:$E$4000, $B${rowNumber})`,
                    date1904: false
                };
            });
        });
  
            // AI列の合計式の設定
            groupedEntries.forEach((entry, index) => {
                const rowNumber = startingRow + index;
                worksheet.getCell(`AI${rowNumber}`).value = {
                    formula: `=SUM(D${rowNumber}:AH${rowNumber})`,
                    date1904: false
                };
            });
            
            // AJ列の計算式の設定
            groupedEntries.forEach((entry, index) => {
                const rowNumber = startingRow + index;
                worksheet.getCell(`AJ${rowNumber}`).value = {
                    formula: `=(C${rowNumber}*AI${rowNumber})`,
                    date1904: false
                };
            });
  
  
        // 最終行の処理
        const totalRow = startingRow + groupedEntries.length;
  
        // A列からC列までのセルを結合して「合計」を入力
        worksheet.mergeCells(`A${totalRow}:C${totalRow}`);
        const totalCell = worksheet.getCell(`A${totalRow}`);
        totalCell.value = '合計';
        totalCell.alignment = { horizontal: 'center' };
  
        // 最終行の合計式設定（D～AJ列）
        const totalColumns = ['D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z',
                             'AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ'];
        totalColumns.forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            cell.value = {
                formula: `=SUM(${column}${startingRow}:${column}${totalRow-1})`,
                date1904: false
            };
            cell.numFmt = '#,##0';
        });
  
        // 単価調整_請求の合計を計算
        const totalPriceAdjustment = records.reduce((sum, record) => {
            return sum + Number(record["単価調整_請求"]?.value || 0);
        }, 0);

        // 単価調整_請求の合計を最終行のR列に設定
        worksheet.getCell(`R${totalRow}`).value = totalPriceAdjustment;
        worksheet.getCell(`R${totalRow}`).numFmt = '#,##0';
  
        // 罫線スタイルの定義
        const borderStyle = {
            style: 'thin',
            color: { argb: '000000' }
        };
  
        const doubleBorderStyle = {
            style: 'double',
            color: { argb: '000000' }
        };
  
        // データ行の罫線設定（最終行を除く）
        for (let row = startingRow; row < totalRow; row++) {
            // A,B,C列にも罫線を適用
            ['A', 'B', 'C'].concat(totalColumns).forEach(column => {
                const cell = worksheet.getCell(`${column}${row}`);
                cell.border = {
                    top: borderStyle,
                    bottom: borderStyle,
                    left: borderStyle,
                    right: borderStyle
                };
            });
        }
  
        // 最終行の特別な罫線設定と着色（AJ列まで）
        // A,B,C列も含めて処理
        ['A', 'B', 'C'].concat(totalColumns).forEach(column => {
            const cell = worksheet.getCell(`${column}${totalRow}`);
            // 罫線設定
            cell.border = {
                top: doubleBorderStyle,
                bottom: borderStyle,
                left: borderStyle,
                right: borderStyle
            };
            // 着色設定
            cell.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: { argb: 'FFA500' }
            };
        });
  
        return workbook;
    } catch (error) {
        console.error('【請求】現場別日別出面表の処理でエラーが発生しました:', error);
        throw error;
    }
  }
  async function generateExcelReport(eventRecord, templateFile) {
    try {
        const ExcelJS = await getXLSX();
        const workbook = new ExcelJS.Workbook();
        await workbook.xlsx.load(templateFile);
        
        const customerName = eventRecord["顧客名"].value;
        const startDate = eventRecord["開始日"].value;
        const endDate = eventRecord["終了日"].value;
        const queryCategory = eventRecord["query_category"].value;
        
        // appid24からレコードを取得
        const client = new KintoneRestAPIClient();
        
        // query_categoryに応じて条件を分岐
        let condition;
        switch (queryCategory) {
            case "常用":
                condition = `顧客名 = "${customerName}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_請求 != 0 and 取引種別 in("常用")`;
                break;
            case "請負(自)":
                condition = `顧客名 = "${customerName}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_請求 != 0 and 取引種別 in("請負(自)")`;
                break;
            case "請負(他)":
                condition = `顧客名 = "${customerName}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_請求 != 0 and 取引種別 in("請負(他)")`;
                break;
            case "すべて":
            default:
                condition = `顧客名 = "${customerName}" and 作業日 >= "${startDate}" and 作業日 <= "${endDate}" and 人工数_請求 != 0`;
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
                "職種_実績_請求",
                "郵便番号",
                "住所",
                "tel",
                "勤怠",
                "単価",
                "遅刻時間_請求",
                "早退時間_請求",
                "残業時間_請求",
                "早出時間_請求",
                "人工数_請求",
                "取引種別",
                "経費_請求",
                "日勤_夜勤",
                "単価調整_請求"
            ]
        });
        
        if (records.length === 0) {
            alert("指定された期間内に該当するデータがありません。\n以下の条件を確認してください：\n・顧客名\n・作業日（開始日～終了日）\n・取引種別\n・人工数_請求が0以外");
            throw new Error("該当するデータがありません");
        }
        
        console.log('取得されたレコード数:', records.length);
        
        // 各シートへの書き込み
      await writeSheet1(workbook, records);
      await writeSheet2(workbook, records);
      await writeInvoiceSheet(workbook, eventRecord);
      console.log('【請求明細】現場職種別シート処理開始');
      await writeWorkSiteSheet(workbook, records);
      console.log('【請求明細】現場別シート処理開始');
      const recordsCopy = JSON.parse(JSON.stringify(records)); // ディープコピーを作成
      await writeWorkSiteSheet2(workbook, recordsCopy);
      await writeBillingSheet(workbook, records);
        
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
        
        // URLの解放（メモリリーク防止）
        setTimeout(() => {
            window.URL.revokeObjectURL(url);
            debugLog('URLが正常に解放されました');
        }, 1000);
    } catch (error) {
        debugLog('ファイルダウンロードでエラーが発生しました:', error);
        throw new Error('Excelファイルのダウンロードに失敗しました');
    }
  }
  
  // キントーンのレコード詳細表示イベントに応答
  kintone.events.on('app.record.detail.show', function(event) {
    // out_categoryが「請求書」でない場合は処理を終了
    if (event.record["out_category"].value !== "請求書・常用") {
      return;
    }

    // 既にボタンが存在する場合は処理を終了
    if (document.getElementById('export_excel_button')) return;
    
    const button = document.createElement('button');
    button.id = 'export_excel_button';
    button.innerText = '請求書出力';
    button.style.margin = '10px';
    button.style.backgroundColor = '#2196F3'; // 青色に変更
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
        this.style.backgroundColor = '#1976D2';
        this.style.transform = 'translateY(-1px)';
        this.style.boxShadow = '0 4px 8px rgba(0,0,0,0.2)';
    };
    
    button.onmouseout = function() {
        this.style.backgroundColor = '#2196F3';
        this.style.transform = 'translateY(0)';
        this.style.boxShadow = '0 2px 4px rgba(0,0,0,0.2)';
    };

    button.onclick = async function() {
        try {
            const record = event.record;
            const fileName = `（請求書）${record["顧客名"].value}_${record["開始日"].value}.xlsx`;
            
            debugLog('テンプレートファイルの取得を試みます');
            const templateFile = await getTemplateFile(31, 6);
            
            debugLog('Excelデータの書き込みを試みます');
            const excelBuffer = await generateExcelReport(record, templateFile);
            
            debugLog('ファイルダウンロードを試みます', { fileName });
            await downloadExcelFile(excelBuffer, fileName);
            
            debugLog('処理が正常に完了しました');
            alert("Excelファイルの出力が完了しました。");
        } catch (error) {
            debugLog('エラーが発生しました:', error);
            alert("処理中にエラーが発生しました。\\n管理者にご連絡ください。");
        }
    };
    
    kintone.app.record.getHeaderMenuSpaceElement().appendChild(button);
  });
  })();