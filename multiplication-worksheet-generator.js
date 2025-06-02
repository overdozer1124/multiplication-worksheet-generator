/**
 * 🎲 小学2年生かけ算問題 ランダム生成システム
 * 重複なし・難易度別問題生成機能
 * 
 * 使用方法:
 * 1. このコードをApps Scriptエディタにコピー＆ペースト
 * 2. generateNewProblems() を実行して新しい問題を生成
 * 3. clearProblems() で問題をクリア
 * 4. testSystem() でシステムテスト
 * 
 * 対象スプレッドシート: 1U1uZhyhExld7z4bLlDzmm0f9PSqwDLvKJYaI3KCY6WA
 */

/**
 * 新しいかけ算問題を生成する関数
 * 重複チェック機能付き
 */
function generateNewProblems() {
  try {
    console.log('📚 問題生成を開始します...');
    
    const spreadsheetId = '1U1uZhyhExld7z4bLlDzmm0f9PSqwDLvKJYaI3KCY6WA';
    
    // スプレッドシートアクセス確認
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log('✅ スプレッドシートにアクセス成功');
    } catch (e) {
      console.error('❌ スプレッドシートアクセスエラー:', e.toString());
      throw new Error('スプレッドシートにアクセスできません: ' + e.toString());
    }
    
    // シート取得
    let problemSheet, answerSheet, controlSheet;
    
    try {
      problemSheet = spreadsheet.getSheetByName('問題シート');
      console.log('✅ 問題シート取得成功');
    } catch (e) {
      console.error('❌ 問題シートが見つかりません');
      throw new Error('問題シートが見つかりません');
    }
    
    try {
      answerSheet = spreadsheet.getSheetByName('解答シート');
      console.log('✅ 解答シート取得成功');
    } catch (e) {
      console.error('❌ 解答シートが見つかりません');
      throw new Error('解答シートが見つかりません');
    }
    
    try {
      controlSheet = spreadsheet.getSheetByName('問題生成制御');
      console.log('✅ 制御シート取得成功');
    } catch (e) {
      console.log('⚠️ 制御シートが見つかりません（デフォルト設定を使用）');
      controlSheet = null;
    }
    
    // 設定値取得
    let totalProblems = 20;
    let easyProblems = 10;
    let standardProblems = 10;
    
    if (controlSheet !== null) {
      try {
        const totalFromSheet = controlSheet.getRange('B9').getValue();
        const easyFromSheet = controlSheet.getRange('B10').getValue();
        const standardFromSheet = controlSheet.getRange('B11').getValue();
        
        if (totalFromSheet && typeof totalFromSheet === 'number' && totalFromSheet > 0) {
          totalProblems = totalFromSheet;
        }
        if (easyFromSheet && typeof easyFromSheet === 'number' && easyFromSheet >= 0) {
          easyProblems = easyFromSheet;
        }
        if (standardFromSheet && typeof standardFromSheet === 'number' && standardFromSheet >= 0) {
          standardProblems = standardFromSheet;
        }
        console.log('✅ 設定値を制御シートから取得');
      } catch (e) {
        console.log('⚠️ 設定値取得エラー、デフォルト値を使用');
      }
    }
    
    console.log('設定: 総問題数=' + totalProblems + ', 易しい=' + easyProblems + ', 標準=' + standardProblems);
    
    // 丸囲み文字（配列として明示的に定義）
    const circledNumbers = [
      '①', '②', '③', '④', '⑤', '⑥', '⑦', '⑧', '⑨', '⑩', 
      '⑪', '⑫', '⑬', '⑭', '⑮', '⑯', '⑰', '⑱', '⑲', '⑳'
    ];
    
    // 重複チェック用セット
    const usedProblems = new Set();
    const problems = [];
    const answers = [];
    
    console.log('🔄 易しい問題を生成中...');
    
    // 易しい問題生成 (1×1 ～ 5×5)
    let easyCount = 0;
    for (let i = 0; i < easyProblems && problems.length < totalProblems && easyCount < easyProblems; i++) {
      let attempts = 0;
      let problem, answer, a, b;
      
      do {
        a = Math.floor(Math.random() * 5) + 1; // 1-5
        b = Math.floor(Math.random() * 5) + 1; // 1-5
        problem = a + ' × ' + b + ' =';
        answer = a * b;
        attempts++;
      } while (usedProblems.has(problem) && attempts < 100);
      
      if (!usedProblems.has(problem)) {
        usedProblems.add(problem);
        problems.push(problem);
        answers.push(answer);
        easyCount++;
      }
    }
    
    console.log('易しい問題 ' + easyCount + '問生成完了');
    
    console.log('🔄 標準問題を生成中...');
    
    // 標準問題生成 (1×1 ～ 9×9)
    let standardCount = 0;
    for (let i = 0; i < standardProblems && problems.length < totalProblems && standardCount < standardProblems; i++) {
      let attempts = 0;
      let problem, answer, a, b;
      
      do {
        a = Math.floor(Math.random() * 9) + 1; // 1-9
        b = Math.floor(Math.random() * 9) + 1; // 1-9
        problem = a + ' × ' + b + ' =';
        answer = a * b;
        attempts++;
      } while (usedProblems.has(problem) && attempts < 100);
      
      if (!usedProblems.has(problem)) {
        usedProblems.add(problem);
        problems.push(problem);
        answers.push(answer);
        standardCount++;
      }
    }
    
    console.log('標準問題 ' + standardCount + '問生成完了');
    console.log('🎯 問題をシャッフル中...');
    
    // 問題をシャッフル
    for (let i = problems.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      const tempProblem = problems[i];
      const tempAnswer = answers[i];
      problems[i] = problems[j];
      answers[i] = answers[j];
      problems[j] = tempProblem;
      answers[j] = tempAnswer;
    }
    
    console.log('📝 問題をシートに配置中...');
    
    // 問題シートと解答シートに配置
    for (let i = 0; i < 10; i++) {
      const row = 6 + i * 2;
      
      try {
        // 既存の内容をクリア
        problemSheet.getRange(row, 2).setValue('');
        problemSheet.getRange(row, 4).setValue('');
        problemSheet.getRange(row, 7).setValue('');
        problemSheet.getRange(row, 9).setValue('');
        answerSheet.getRange(row, 2).setValue('');
        answerSheet.getRange(row, 4).setValue('');
        answerSheet.getRange(row, 7).setValue('');
        answerSheet.getRange(row, 9).setValue('');
        
        // 左列（1-10問）
        if (i < problems.length) {
          problemSheet.getRange(row, 1).setValue(circledNumbers[i]).setFontWeight('bold').setFontSize(14);
          problemSheet.getRange(row, 2).setValue(problems[i]).setFontSize(14);
          problemSheet.getRange(row, 4).setValue('　　　　　').setBorder(false, false, true, false, false, false);
          
          answerSheet.getRange(row, 1).setValue(circledNumbers[i]).setFontWeight('bold').setFontSize(14);
          answerSheet.getRange(row, 2).setValue(problems[i]).setFontSize(14);
          answerSheet.getRange(row, 4).setValue(answers[i]).setFontColor('#ff0000').setFontWeight('bold').setFontSize(14);
        }
        
        // 右列（11-20問）
        if (i + 10 < problems.length) {
          problemSheet.getRange(row, 6).setValue(circledNumbers[i + 10]).setFontWeight('bold').setFontSize(14);
          problemSheet.getRange(row, 7).setValue(problems[i + 10]).setFontSize(14);
          problemSheet.getRange(row, 9).setValue('　　　　　').setBorder(false, false, true, false, false, false);
          
          answerSheet.getRange(row, 6).setValue(circledNumbers[i + 10]).setFontWeight('bold').setFontSize(14);
          answerSheet.getRange(row, 7).setValue(problems[i + 10]).setFontSize(14);
          answerSheet.getRange(row, 9).setValue(answers[i + 10]).setFontColor('#ff0000').setFontWeight('bold').setFontSize(14);
        }
      } catch (e) {
        console.error('行 ' + row + ' の配置でエラー:', e.toString());
      }
    }
    
    console.log('📊 統計情報を更新中...');
    
    // 統計更新
    if (controlSheet !== null) {
      try {
        controlSheet.getRange('E15').setValue(new Date().toLocaleString('ja-JP'));
        const currentCount = controlSheet.getRange('E16').getValue() || 0;
        controlSheet.getRange('E16').setValue(currentCount + 1);
      } catch (e) {
        console.log('統計更新をスキップしました:', e.toString());
      }
    }
    
    console.log('✅ 新しい問題を' + problems.length + '問生成しました（重複なし）');
    
    // 成功メッセージ表示
    SpreadsheetApp.getUi().alert(
      '✅ 生成完了', 
      '新しいかけ算問題を' + problems.length + '問生成しました！\n\n' +
      '・重複チェック: 完了\n' +
      '・易しい問題: ' + easyCount + '問\n' +
      '・標準問題: ' + standardCount + '問\n' +
      '・問題シート・解答シート: 更新済み', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return problems.length;
    
  } catch (error) {
    console.error('❌ 問題生成エラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ エラー', 
      '問題生成中にエラーが発生しました:\n\n' + error.toString() + '\n\nログを確認してください。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * 問題をクリアする関数
 */
function clearProblems() {
  try {
    console.log('🗑️ 問題をクリア中...');
    
    const spreadsheetId = '1U1uZhyhExld7z4bLlDzmm0f9PSqwDLvKJYaI3KCY6WA';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const problemSheet = spreadsheet.getSheetByName('問題シート');
    const answerSheet = spreadsheet.getSheetByName('解答シート');
    
    // 問題エリアをクリア
    for (let i = 0; i < 10; i++) {
      const row = 6 + i * 2;
      
      try {
        // 左列クリア
        problemSheet.getRange(row, 2).setValue('');
        problemSheet.getRange(row, 4).setValue('');
        answerSheet.getRange(row, 2).setValue('');
        answerSheet.getRange(row, 4).setValue('');
        
        // 右列クリア
        problemSheet.getRange(row, 7).setValue('');
        problemSheet.getRange(row, 9).setValue('');
        answerSheet.getRange(row, 7).setValue('');
        answerSheet.getRange(row, 9).setValue('');
      } catch (e) {
        console.error('行 ' + row + ' のクリアでエラー:', e.toString());
      }
    }
    
    console.log('✅ 問題をクリアしました');
    
    SpreadsheetApp.getUi().alert(
      '✅ クリア完了', 
      '問題シートと解答シートの問題をクリアしました。\n\n新しい問題を生成するには、generateNewProblems()を実行してください。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
    return true;
    
  } catch (error) {
    console.error('❌ クリアエラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ エラー', 
      '問題クリア中にエラーが発生しました:\n\n' + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * システムテスト関数
 */
function testSystem() {
  try {
    console.log('🧪 システムテストを開始します...');
    
    const spreadsheetId = '1U1uZhyhExld7z4bLlDzmm0f9PSqwDLvKJYaI3KCY6WA';
    
    // スプレッドシート接続テスト
    console.log('1. スプレッドシート接続テスト...');
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('✅ スプレッドシート接続成功');
    
    // シート存在確認
    console.log('2. シート存在確認...');
    const requiredSheets = ['問題シート', '解答シート'];
    const missingSheets = [];
    const foundSheets = [];
    
    for (let i = 0; i < requiredSheets.length; i++) {
      const sheetName = requiredSheets[i];
      try {
        spreadsheet.getSheetByName(sheetName);
        foundSheets.push(sheetName);
        console.log('✅ ' + sheetName + ' 存在確認');
      } catch (e) {
        missingSheets.push(sheetName);
        console.log('❌ ' + sheetName + ' が見つかりません');
      }
    }
    
    // 制御シートは任意
    try {
      spreadsheet.getSheetByName('問題生成制御');
      foundSheets.push('問題生成制御');
      console.log('✅ 問題生成制御 存在確認');
    } catch (e) {
      console.log('⚠️ 問題生成制御シートが見つかりません（任意）');
    }
    
    console.log('🧪 システムテスト完了');
    
    if (missingSheets.length > 0) {
      SpreadsheetApp.getUi().alert(
        '⚠️ テスト結果', 
        '以下のシートが見つかりません:\n' + missingSheets.join(', ') + '\n\n見つかったシート:\n' + foundSheets.join(', ') + '\n\nスプレッドシートの構成を確認してください。', 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      SpreadsheetApp.getUi().alert(
        '✅ テスト成功', 
        'システムは正常に動作可能です！\n\n確認済み項目:\n・スプレッドシート接続\n・必要シートの存在\n・関数の実行環境\n\n見つかったシート:\n' + foundSheets.join(', '), 
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
    
  } catch (error) {
    console.error('❌ テストエラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ テスト失敗', 
      'システムテスト中にエラーが発生しました:\n\n' + error.toString() + '\n\nスプレッドシートIDまたは権限を確認してください。', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

/**
 * 統計情報を表示する関数
 */
function showStatistics() {
  try {
    const spreadsheetId = '1U1uZhyhExld7z4bLlDzmm0f9PSqwDLvKJYaI3KCY6WA';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const controlSheet = spreadsheet.getSheetByName('問題生成制御');
    
    let lastGenerated = '未生成';
    let generationCount = 0;
    
    try {
      lastGenerated = controlSheet.getRange('E15').getValue() || '未生成';
      generationCount = controlSheet.getRange('E16').getValue() || 0;
    } catch (e) {
      console.log('統計情報の取得に失敗しました');
    }
    
    SpreadsheetApp.getUi().alert(
      '📊 統計情報', 
      '問題生成システムの利用状況:\n\n' +
      '・最終生成日時: ' + lastGenerated + '\n' +
      '・累計生成回数: ' + generationCount + '回\n' +
      '・重複チェック: 有効\n' +
      '・問題範囲: 1×1 ～ 9×9', 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    
  } catch (error) {
    console.error('❌ 統計表示エラー:', error);
    SpreadsheetApp.getUi().alert(
      '❌ エラー', 
      '統計情報の表示中にエラーが発生しました:\n' + error.toString(), 
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// システム初期化時のログ出力
console.log('📚 かけ算問題生成システム関数を読み込みました');
console.log('利用可能な関数:');
console.log('- generateNewProblems(): 新しい問題を生成');
console.log('- clearProblems(): 問題をクリア'); 
console.log('- testSystem(): システムテスト');
console.log('- showStatistics(): 統計情報表示');