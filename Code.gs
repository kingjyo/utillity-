/**
 * 구글 스프레드시트 유틸리티 도구모음
 * 다양한 데이터 처리 및 포맷팅 기능 제공
 */

// 스프레드시트가 열릴 때 메뉴 생성
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 유틸리티 도구')
    .addSubMenu(ui.createMenu('🔄 데이터 취합 및 정리')
      .addItem('1. 여러 시트에서 데이터 취합', 'consolidateData')
      .addItem('2. 헤더 자동 정렬 및 데이터 정리', 'autoAlignHeaders')
      .addItem('12. 다른 스프레드시트에서 데이터 가져오기', 'importFromOtherSpreadsheet')
      .addItem('13. 여러 스프레드시트 데이터 통합', 'consolidateMultipleSpreadsheets')
      .addItem('14. 모든 시트 데이터 삭제', 'clearAllSheets')
      .addItem('15. Index-Match 기능', 'indexMatchFunction'))
    .addSubMenu(ui.createMenu('⏱️ 시간 변환')
      .addItem('3. 시간을 분으로 변환', 'convertTimeToMinutes')
      .addItem('6. 분을 시간으로 변환', 'convertMinutesToTime'))
    .addSubMenu(ui.createMenu('🔢 숫자 포맷팅')
      .addItem('5. 숫자를 #,##0 형식으로 변환', 'formatNumbersWithCommas')
      .addItem('8-1. 천원 단위로 표시', 'convertToThousandWon')
      .addItem('8-2. 천원을 원래 숫자로 변환', 'convertThousandWonToNumber'))
    .addSubMenu(ui.createMenu('✏️ 텍스트 변환')
      .addItem('7-1. 하이픈(-)을 언더스코어(_)로 변환', 'convertDashToUnderscore')
      .addItem('7-2. 언더스코어(_)를 하이픈(-)으로 변환', 'convertUnderscoreToDash')
      .addItem('19. 문법 오류 교정', 'correctGrammar'))
    .addSubMenu(ui.createMenu('📅 날짜 변환')
      .addItem('17. 날짜 양식 통일', 'unifyDateFormats')
      .addItem('18. 날짜의 년도/월 변경', 'changeDateYearMonth'))
    .addSubMenu(ui.createMenu('🗑️ 삭제 및 정리')
      .addItem('4. 특정 열/행 삭제', 'deleteSpecificColumnsRows')
      .addItem('16. 데이터 표 자동 포맷팅', 'autoFormatTable'))
    .addSubMenu(ui.createMenu('🖼️ 이미지')
      .addItem('20. 여러 이미지 일괄 업로드', 'uploadMultipleImages')
      .addItem('21. 이미지를 셀에 맞게 삽입', 'insertImagesInCells'))
    .addSubMenu(ui.createMenu('🖨️ 인쇄 및 내보내기')
      .addItem('9. 데이터를 Google Docs로 이동', 'exportToGoogleDocs')
      .addItem('10. 특정 시트 인쇄', 'printSpecificSheet')
      .addItem('11. 여러 시트 인쇄', 'printMultipleSheets'))
    .addToUi();
}

// ==================== 1. 여러 시트에서 데이터 취합 ====================
function consolidateData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // 시트 목록 가져오기
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(sheet) { return sheet.getName(); }).join(', ');
  
  var sheetSelection = ui.prompt(
    '데이터 취합',
    '취합할 시트 이름들을 쉼표로 구분하여 입력하세요:\n(사용 가능한 시트: ' + sheetNames + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetSelection.getSelectedButton() != ui.Button.OK) return;
  
  var selectedSheets = sheetSelection.getResponseText().split(',').map(function(s) { return s.trim(); });
  
  var columnInput = ui.prompt(
    '열 선택',
    '취합할 열 번호를 쉼표로 구분하여 입력하세요 (예: 1,3,5):\n빈칸으로 두면 전체 열을 취합합니다.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (columnInput.getSelectedButton() != ui.Button.OK) return;
  
  var columns = columnInput.getResponseText().trim() === '' ? 
    null : columnInput.getResponseText().split(',').map(function(c) { return parseInt(c.trim()); });
  
  var rowInput = ui.prompt(
    '행 선택',
    '취합할 행 범위를 입력하세요 (예: 2-100, 전체는 빈칸):\n빈칸으로 두면 전체 행을 취합합니다.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (rowInput.getSelectedButton() != ui.Button.OK) return;
  
  var rowRange = rowInput.getResponseText().trim();
  
  // 새 시트 생성
  var newSheet = ss.insertSheet('취합_데이터_' + new Date().getTime());
  var consolidatedData = [];
  
  selectedSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      ui.alert('경고', '시트 "' + sheetName + '"을(를) 찾을 수 없습니다.', ui.ButtonSet.OK);
      return;
    }
    
    var data = sheet.getDataRange().getValues();
    
    // 행 범위 적용
    var startRow = 0;
    var endRow = data.length;
    if (rowRange !== '') {
      var rangeParts = rowRange.split('-');
      startRow = parseInt(rangeParts[0]) - 1;
      endRow = rangeParts.length > 1 ? parseInt(rangeParts[1]) : data.length;
    }
    
    for (var i = startRow; i < endRow && i < data.length; i++) {
      var row = data[i];
      var newRow = [];
      
      if (columns === null) {
        newRow = row;
      } else {
        columns.forEach(function(col) {
          if (col - 1 < row.length) {
            newRow.push(row[col - 1]);
          }
        });
      }
      
      // 시트 이름 추가
      newRow.unshift(sheetName);
      consolidatedData.push(newRow);
    }
  });
  
  if (consolidatedData.length > 0) {
    newSheet.getRange(1, 1, consolidatedData.length, consolidatedData[0].length).setValues(consolidatedData);
    ui.alert('완료', '데이터 취합이 완료되었습니다!\n새 시트: ' + newSheet.getName(), ui.ButtonSet.OK);
  } else {
    ui.alert('오류', '취합할 데이터가 없습니다.', ui.ButtonSet.OK);
  }
}

// ==================== 2. 헤더 자동 정렬 및 데이터 정리 ====================
function autoAlignHeaders() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var data = sheet.getDataRange().getValues();
  
  if (data.length === 0) {
    ui.alert('오류', '데이터가 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  // 헤더 찾기 (첫 번째 비어있지 않은 행)
  var headerRowIndex = -1;
  for (var i = 0; i < data.length; i++) {
    var hasData = data[i].some(function(cell) { return cell !== ''; });
    if (hasData) {
      headerRowIndex = i;
      break;
    }
  }
  
  if (headerRowIndex === -1) {
    ui.alert('오류', '헤더를 찾을 수 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  // 헤더와 데이터 추출
  var headers = data[headerRowIndex];
  var cleanedData = [headers];
  
  for (var i = headerRowIndex + 1; i < data.length; i++) {
    var hasData = data[i].some(function(cell) { return cell !== ''; });
    if (hasData) {
      cleanedData.push(data[i]);
    }
  }
  
  // 시트 클리어 및 데이터 입력
  sheet.clear();
  sheet.getRange(1, 1, cleanedData.length, cleanedData[0].length).setValues(cleanedData);
  
  // 헤더 포맷팅
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff');
  
  ui.alert('완료', '헤더가 A1에 맞춰졌고 데이터가 정리되었습니다!', ui.ButtonSet.OK);
}

// ==================== 3. 시간을 분으로 변환 ====================
function convertTimeToMinutes() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'string' && cell.match(/^\d{1,3}:\d{2}$/)) {
        var parts = cell.split(':');
        var hours = parseInt(parts[0]);
        var minutes = parseInt(parts[1]);
        return hours * 60 + minutes;
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '시간이 분으로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 4. 특정 열/행 삭제 ====================
function deleteSpecificColumnsRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var typeResponse = ui.alert(
    '삭제 유형',
    '열을 삭제하시겠습니까? (아니오를 선택하면 행 삭제)',
    ui.ButtonSet.YES_NO
  );
  
  var isColumn = (typeResponse == ui.Button.YES);
  
  var input = ui.prompt(
    isColumn ? '열 삭제' : '행 삭제',
    '삭제할 ' + (isColumn ? '열 번호' : '행 번호') + '를 쉼표로 구분하여 입력하세요 (예: 2,4,5):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (input.getSelectedButton() != ui.Button.OK) return;
  
  var indices = input.getResponseText().split(',').map(function(s) { 
    return parseInt(s.trim()); 
  }).sort(function(a, b) { return b - a; }); // 역순 정렬
  
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    indices.forEach(function(index) {
      if (isColumn) {
        if (index <= sheet.getMaxColumns()) {
          sheet.deleteColumn(index);
        }
      } else {
        if (index <= sheet.getMaxRows()) {
          sheet.deleteRow(index);
        }
      }
    });
  });
  
  ui.alert('완료', '모든 시트에서 ' + (isColumn ? '열' : '행') + ' 삭제가 완료되었습니다!', ui.ButtonSet.OK);
}

// ==================== 5. 숫자를 #,##0 형식으로 변환 ====================
function formatNumbersWithCommas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    var range = sheet.getDataRange();
    var values = range.getValues();
    
    for (var i = 0; i < values.length; i++) {
      for (var j = 0; j < values[i].length; j++) {
        if (typeof values[i][j] === 'number') {
          sheet.getRange(i + 1, j + 1).setNumberFormat('#,##0');
        }
      }
    }
  });
  
  ui.alert('완료', '모든 시트의 숫자가 #,##0 형식으로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 6. 분을 시간으로 변환 ====================
function convertMinutesToTime() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'number') {
        var hours = Math.floor(cell / 60);
        var minutes = cell % 60;
        return hours + ':' + (minutes < 10 ? '0' : '') + minutes;
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '분이 시간으로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 7-1. 하이픈을 언더스코어로 변환 ====================
function convertDashToUnderscore() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'string') {
        return cell.replace(/-/g, '_');
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '하이픈(-)이 언더스코어(_)로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 7-2. 언더스코어를 하이픈으로 변환 ====================
function convertUnderscoreToDash() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'string') {
        return cell.replace(/_/g, '-');
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '언더스코어(_)가 하이픈(-)으로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 8-1. 천원 단위로 표시 ====================
function convertToThousandWon() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'number') {
        var thousands = cell / 1000;
        return thousands.toLocaleString('ko-KR') + '천원';
      } else if (typeof cell === 'string') {
        var num = parseFloat(cell.replace(/,/g, ''));
        if (!isNaN(num)) {
          var thousands = num / 1000;
          return thousands.toLocaleString('ko-KR') + '천원';
        }
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '숫자가 천원 단위로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 8-2. 천원을 원래 숫자로 변환 ====================
function convertThousandWonToNumber() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      if (typeof cell === 'string' && cell.includes('천원')) {
        var numStr = cell.replace('천원', '').replace(/,/g, '').trim();
        var num = parseFloat(numStr);
        if (!isNaN(num)) {
          return num * 1000;
        }
      }
      return cell;
    });
  });
  
  selection.setValues(newValues);
  selection.setNumberFormat('#,##0');
  ui.alert('완료', '천원이 원래 숫자로 변환되었습니다!', ui.ButtonSet.OK);
}

// ==================== 9. 데이터를 Google Docs로 이동 ====================
function exportToGoogleDocs() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  // 새 문서 생성
  var doc = DocumentApp.create(sheet.getName() + '_데이터_' + new Date().getTime());
  var body = doc.getBody();
  
  // 제목 추가
  body.appendParagraph(sheet.getName() + ' 데이터').setHeading(DocumentApp.ParagraphHeading.HEADING1);
  
  // 테이블 생성
  var table = body.appendTable();
  
  values.forEach(function(row) {
    var tableRow = table.appendTableRow();
    row.forEach(function(cell) {
      tableRow.appendTableCell(cell.toString());
    });
  });
  
  // 스타일 적용
  var headerRow = table.getRow(0);
  for (var i = 0; i < headerRow.getNumCells(); i++) {
    headerRow.getCell(i).setBackgroundColor('#4a86e8').getChild(0).asParagraph().setForegroundColor('#ffffff').setBold(true);
  }
  
  ui.alert('완료', 'Google Docs로 내보내기 완료!\n문서 URL: ' + doc.getUrl(), ui.ButtonSet.OK);
}

// ==================== 10. 특정 시트 인쇄 ====================
function printSpecificSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(sheet) { return sheet.getName(); }).join(', ');
  
  var input = ui.prompt(
    '시트 인쇄',
    '인쇄할 시트 이름을 입력하세요:\n(사용 가능한 시트: ' + sheetNames + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (input.getSelectedButton() != ui.Button.OK) return;
  
  var sheetName = input.getResponseText().trim();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    ui.alert('오류', '시트 "' + sheetName + '"을(를) 찾을 수 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var url = ss.getUrl();
  var printUrl = url.replace(/edit.*/, 'export?format=pdf&gid=' + sheet.getSheetId());
  
  ui.alert('인쇄', '아래 링크를 클릭하여 시트를 인쇄하세요:\n' + printUrl, ui.ButtonSet.OK);
}

// ==================== 11. 여러 시트 인쇄 ====================
function printMultipleSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var sheets = ss.getSheets();
  var sheetNames = sheets.map(function(sheet) { return sheet.getName(); }).join(', ');
  
  var input = ui.prompt(
    '여러 시트 인쇄',
    '인쇄할 시트 이름들을 쉼표로 구분하여 입력하세요:\n(사용 가능한 시트: ' + sheetNames + ')',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (input.getSelectedButton() != ui.Button.OK) return;
  
  var selectedSheets = input.getResponseText().split(',').map(function(s) { return s.trim(); });
  var gids = [];
  
  selectedSheets.forEach(function(sheetName) {
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      gids.push(sheet.getSheetId());
    } else {
      ui.alert('경고', '시트 "' + sheetName + '"을(를) 찾을 수 없습니다.', ui.ButtonSet.OK);
    }
  });
  
  if (gids.length === 0) {
    ui.alert('오류', '유효한 시트를 찾을 수 없습니다.', ui.ButtonSet.OK);
    return;
  }
  
  var url = ss.getUrl();
  var printUrl = url.replace(/edit.*/, 'export?format=pdf&gid=' + gids.join(','));
  
  ui.alert('인쇄', '아래 링크를 클릭하여 시트들을 인쇄하세요:\n' + printUrl, ui.ButtonSet.OK);
}

// ==================== 12. 다른 스프레드시트에서 데이터 가져오기 ====================
function importFromOtherSpreadsheet() {
  var ui = SpreadsheetApp.getUi();
  
  var urlInput = ui.prompt(
    '스프레드시트 가져오기',
    '가져올 스프레드시트의 URL을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (urlInput.getSelectedButton() != ui.Button.OK) return;
  
  try {
    var sourceSpreadsheet = SpreadsheetApp.openByUrl(urlInput.getResponseText());
    var sheets = sourceSpreadsheet.getSheets();
    var sheetNames = sheets.map(function(sheet) { return sheet.getName(); }).join(', ');
    
    var sheetInput = ui.prompt(
      '시트 선택',
      '가져올 시트 이름을 입력하세요:\n(사용 가능한 시트: ' + sheetNames + ')',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (sheetInput.getSelectedButton() != ui.Button.OK) return;
    
    var sourceSheet = sourceSpreadsheet.getSheetByName(sheetInput.getResponseText().trim());
    
    if (!sourceSheet) {
      ui.alert('오류', '시트를 찾을 수 없습니다.', ui.ButtonSet.OK);
      return;
    }
    
    var columnInput = ui.prompt(
      '열 선택',
      '가져올 열 번호를 쉼표로 구분하여 입력하세요 (예: 1,3,5):\n빈칸으로 두면 전체를 가져옵니다.',
      ui.ButtonSet.OK_CANCEL
    );
    
    if (columnInput.getSelectedButton() != ui.Button.OK) return;
    
    var columns = columnInput.getResponseText().trim() === '' ? 
      null : columnInput.getResponseText().split(',').map(function(c) { return parseInt(c.trim()); });
    
    var data = sourceSheet.getDataRange().getValues();
    var newData = [];
    
    data.forEach(function(row) {
      if (columns === null) {
        newData.push(row);
      } else {
        var newRow = [];
        columns.forEach(function(col) {
          if (col - 1 < row.length) {
            newRow.push(row[col - 1]);
          }
        });
        newData.push(newRow);
      }
    });
    
    var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var newSheet = currentSpreadsheet.insertSheet('가져온_데이터_' + new Date().getTime());
    
    if (newData.length > 0) {
      newSheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
      ui.alert('완료', '데이터 가져오기가 완료되었습니다!', ui.ButtonSet.OK);
    }
    
  } catch (e) {
    ui.alert('오류', '스프레드시트를 열 수 없습니다. URL을 확인해주세요.\n' + e.toString(), ui.ButtonSet.OK);
  }
}

// ==================== 13. 여러 스프레드시트 데이터 통합 ====================
function consolidateMultipleSpreadsheets() {
  var ui = SpreadsheetApp.getUi();
  var currentSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var input = ui.prompt(
    '여러 스프레드시트 통합',
    '통합할 스프레드시트 URL들을 쉼표로 구분하여 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (input.getSelectedButton() != ui.Button.OK) return;
  
  var urls = input.getResponseText().split(',').map(function(s) { return s.trim(); });
  
  urls.forEach(function(url) {
    try {
      var sourceSpreadsheet = SpreadsheetApp.openByUrl(url);
      var sourceSheet = sourceSpreadsheet.getSheets()[0]; // 첫 번째 시트만 가져옴
      var data = sourceSheet.getDataRange().getValues();
      
      var newSheet = currentSpreadsheet.insertSheet(sourceSpreadsheet.getName() + '_' + new Date().getTime());
      
      if (data.length > 0) {
        newSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      
    } catch (e) {
      ui.alert('경고', 'URL을 처리할 수 없습니다: ' + url + '\n' + e.toString(), ui.ButtonSet.OK);
    }
  });
  
  ui.alert('완료', '스프레드시트 통합이 완료되었습니다!', ui.ButtonSet.OK);
}

// ==================== 14. 모든 시트 데이터 삭제 ====================
function clearAllSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var response = ui.alert(
    '경고',
    '정말로 모든 시트의 데이터를 삭제하시겠습니까?\n이 작업은 되돌릴 수 없습니다!',
    ui.ButtonSet.YES_NO
  );
  
  if (response != ui.Button.YES) return;
  
  var sheets = ss.getSheets();
  
  sheets.forEach(function(sheet) {
    sheet.clear();
  });
  
  ui.alert('완료', '모든 시트의 데이터가 삭제되었습니다!', ui.ButtonSet.OK);
}

// ==================== 15. Index-Match 기능 ====================
function indexMatchFunction() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var lookupInput = ui.prompt(
    'Index-Match',
    '찾을 값을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (lookupInput.getSelectedButton() != ui.Button.OK) return;
  
  var lookupValue = lookupInput.getResponseText();
  
  var lookupColumnInput = ui.prompt(
    '검색 열',
    '검색할 열 번호를 입력하세요 (예: 1):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (lookupColumnInput.getSelectedButton() != ui.Button.OK) return;
  
  var lookupColumn = parseInt(lookupColumnInput.getResponseText());
  
  var returnColumnInput = ui.prompt(
    '반환 열',
    '반환할 열 번호를 입력하세요 (예: 3):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (returnColumnInput.getSelectedButton() != ui.Button.OK) return;
  
  var returnColumn = parseInt(returnColumnInput.getResponseText());
  
  var data = sheet.getDataRange().getValues();
  var result = null;
  
  for (var i = 0; i < data.length; i++) {
    if (data[i][lookupColumn - 1] == lookupValue) {
      result = data[i][returnColumn - 1];
      break;
    }
  }
  
  if (result !== null) {
    ui.alert('결과', '찾은 값: ' + result, ui.ButtonSet.OK);
  } else {
    ui.alert('결과', '값을 찾을 수 없습니다.', ui.ButtonSet.OK);
  }
}

// ==================== 16. 데이터 표 자동 포맷팅 ====================
function autoFormatTable() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var range = sheet.getDataRange();
  
  // 헤더 포맷팅
  var headerRange = sheet.getRange(1, 1, 1, range.getLastColumn());
  headerRange
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  // 테두리 추가
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  
  // 교대 색상
  for (var i = 2; i <= range.getLastRow(); i++) {
    var rowRange = sheet.getRange(i, 1, 1, range.getLastColumn());
    if (i % 2 === 0) {
      rowRange.setBackground('#f3f3f3');
    } else {
      rowRange.setBackground('#ffffff');
    }
  }
  
  // 열 자동 크기 조정
  for (var i = 1; i <= range.getLastColumn(); i++) {
    sheet.autoResizeColumn(i);
  }
  
  // 고정 헤더
  sheet.setFrozenRows(1);
  
  ui.alert('완료', '데이터 표가 포맷팅되었습니다!', ui.ButtonSet.OK);
}

// ==================== 17. 날짜 양식 통일 ====================
function unifyDateFormats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var formatInput = ui.prompt(
    '날짜 양식',
    '변환할 날짜 양식을 선택하세요:\n1: YYYY-MM-DD\n2: YYYY년 MM월 DD일\n3: MM/DD/YYYY\n4: DD/MM/YYYY',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (formatInput.getSelectedButton() != ui.Button.OK) return;
  
  var formatType = parseInt(formatInput.getResponseText());
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      var date = null;
      
      // 날짜 파싱
      if (cell instanceof Date) {
        date = cell;
      } else if (typeof cell === 'string') {
        // 다양한 형식 파싱
        if (cell.match(/\d{4}년\s?\d{1,2}월\s?\d{1,2}일/)) {
          var parts = cell.match(/(\d{4})년\s?(\d{1,2})월\s?(\d{1,2})일/);
          date = new Date(parts[1], parts[2] - 1, parts[3]);
        } else if (cell.match(/\d{4}-\d{2}-\d{2}/)) {
          date = new Date(cell);
        } else if (cell.match(/\d{1,2}\/\d{1,2}\/\d{4}/)) {
          date = new Date(cell);
        }
      }
      
      if (date && !isNaN(date.getTime())) {
        var year = date.getFullYear();
        var month = ('0' + (date.getMonth() + 1)).slice(-2);
        var day = ('0' + date.getDate()).slice(-2);
        
        switch(formatType) {
          case 1:
            return year + '-' + month + '-' + day;
          case 2:
            return year + '년 ' + month + '월 ' + day + '일';
          case 3:
            return month + '/' + day + '/' + year;
          case 4:
            return day + '/' + month + '/' + year;
          default:
            return cell;
        }
      }
      
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '날짜 양식이 통일되었습니다!', ui.ButtonSet.OK);
}

// ==================== 18. 날짜의 년도/월 변경 ====================
function changeDateYearMonth() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var typeInput = ui.alert(
    '변경 유형',
    '년도를 변경하시겠습니까? (아니오를 선택하면 월 변경)',
    ui.ButtonSet.YES_NO
  );
  
  var isYear = (typeInput == ui.Button.YES);
  
  var valueInput = ui.prompt(
    isYear ? '년도 변경' : '월 변경',
    '새로운 ' + (isYear ? '년도' : '월') + '을 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (valueInput.getSelectedButton() != ui.Button.OK) return;
  
  var newValue = parseInt(valueInput.getResponseText());
  var selection = sheet.getActiveRange();
  var values = selection.getValues();
  
  var newValues = values.map(function(row) {
    return row.map(function(cell) {
      var date = null;
      
      if (cell instanceof Date) {
        date = new Date(cell);
      } else if (typeof cell === 'string') {
        date = new Date(cell);
      }
      
      if (date && !isNaN(date.getTime())) {
        if (isYear) {
          date.setFullYear(newValue);
        } else {
          date.setMonth(newValue - 1);
        }
        return date;
      }
      
      return cell;
    });
  });
  
  selection.setValues(newValues);
  ui.alert('완료', '날짜의 ' + (isYear ? '년도' : '월') + '가 변경되었습니다!', ui.ButtonSet.OK);
}

// ==================== 19. 문법 오류 교정 ====================
function correctGrammar() {
  var ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '문법 교정',
    '죄송합니다. 이 기능은 외부 API(예: OpenAI, Google Cloud Natural Language)가 필요합니다.\n' +
    '현재 버전에서는 구현되지 않았습니다.\n\n' +
    '대신 다음을 사용해보세요:\n' +
    '1. Google Docs의 맞춤법 검사\n' +
    '2. LanguageTool 애드온\n' +
    '3. Grammarly',
    ui.ButtonSet.OK
  );
}

// ==================== 20. 여러 이미지 일괄 업로드 ====================
function uploadMultipleImages() {
  var ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '이미지 업로드',
    '죄송합니다. Google Apps Script는 로컬 파일 시스템에 직접 접근할 수 없습니다.\n\n' +
    '대신 다음 방법을 사용해주세요:\n' +
    '1. 이미지를 Google Drive에 업로드\n' +
    '2. 스프레드시트에서 삽입 > 이미지 > 드라이브에서 이미지 선택\n' +
    '3. 또는 IMAGE() 함수를 사용하여 URL로 이미지 삽입',
    ui.ButtonSet.OK
  );
}

// ==================== 21. 이미지를 셀에 맞게 삽입 ====================
function insertImagesInCells() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  
  var columnInput = ui.prompt(
    '이미지 삽입',
    '이미지 URL이 있는 열 번호를 입력하세요:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (columnInput.getSelectedButton() != ui.Button.OK) return;
  
  var column = parseInt(columnInput.getResponseText());
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) { // 헤더 제외
    var imageUrl = data[i][column - 1];
    
    if (typeof imageUrl === 'string' && (imageUrl.startsWith('http://') || imageUrl.startsWith('https://'))) {
      var formula = '=IMAGE("' + imageUrl + '", 1)';
      sheet.getRange(i + 1, column).setFormula(formula);
      
      // 행 높이 조정
      sheet.setRowHeight(i + 1, 100);
    }
  }
  
  ui.alert('완료', '이미지가 셀에 삽입되었습니다!', ui.ButtonSet.OK);
}
