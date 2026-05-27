/**
 * 상품 매칭 검수툴 - 백엔드 (자유 모드 + 표준 모드 지원)
 *
 * 메인 시트 (이 파일):
 *   Projects : 프로젝트 목록 (id, name, createdAt, spreadsheetId, mode, schemaJson)
 *   Workers  : 전역 작업자 풀
 *
 * 프로젝트별 스프레드시트 (자동 생성):
 *   Data         : 데이터 (자유 컬럼)
 *   Results      : 검수 결과
 *   Assignments  : 행 범위 할당
 *   Reasons      : 오매핑 사유
 *   Logs         : 작업 시간 로그
 *
 * 모드:
 *   standard : 네이버↔쿠팡 고정 컬럼 (기존 프로젝트 호환용)
 *   free     : 자유 컬럼 - 1행=그룹헤더(병합), 2행=실제 컬럼명
 */

const ADMIN_PASSWORD = 'admin1234';
const SHEET_NAMES = {
  PROJECTS: 'Projects',
  WORKERS: 'Workers'
};

const DEFAULT_REASONS = [
  { name: '다른상품_브랜드', shortcut: 'Q' },
  { name: '다른상품_속성상이 (색상, 수량, 사이즈)', shortcut: 'W' },
  { name: '다른상품_구성상이 (세트 등)', shortcut: 'E' },
  { name: '다른상품_모델상이 (제목 내 모델명 기준)', shortcut: 'R' },
  { name: '다른상품_원산지/생산지상이 (식품 등)', shortcut: 'A' },
  { name: '다른상품_전혀다름', shortcut: 'S' },
  { name: '기타_애매한케이스', shortcut: 'D' },
  { name: '확인불가(대상/매칭 상품 정보 알 수 없음)', shortcut: 'F' }
];

// 표준 모드(네이버/쿠팡)의 고정 컬럼
const STANDARD_HEADERS = [
  '상품번호', '상품명', '네이버_URL', '브랜드명', '네이버_옵션', '로켓_여부',
  'C_상품번호', 'C_상품명', 'C_옵션명', 'C_URL'
];
const STANDARD_SCHEMA = {
  mode: 'standard',
  leftGroup: { name: '네이버', columns: ['상품번호', '상품명', '브랜드명', '네이버_옵션', '네이버_URL'] },
  rightGroup: { name: '쿠팡', columns: ['C_상품번호', 'C_상품명', 'C_옵션명', 'C_URL'] }
};

const PROJECT_FOLDER_NAME = '검수툴_프로젝트시트';

// ============ 진입점 ============
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'download' && e.parameter.projectId && e.parameter.token) {
    return handleDownloadRequest_(e.parameter);
  }
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('상품 매칭 검수툴')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ============ 헬퍼 ============
function getOrCreateProjectFolder_() {
  const folders = DriveApp.getFoldersByName(PROJECT_FOLDER_NAME);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(PROJECT_FOLDER_NAME);
}

function getOrCreateSheet_(name, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function initSheets() {
  const projectsSheet = getOrCreateSheet_(SHEET_NAMES.PROJECTS, ['프로젝트ID', '프로젝트명', '생성일', '스프레드시트ID', '모드', '스키마JSON']);
  // 컬럼 부족하면 추가
  const lastCol = projectsSheet.getLastColumn();
  if (lastCol < 6) {
    const need = ['스프레드시트ID', '모드', '스키마JSON'];
    for (let i = lastCol + 1; i <= 6; i++) {
      const idx = i - 4;
      if (idx >= 0 && idx < need.length) {
        projectsSheet.getRange(1, i).setValue(need[idx]);
      }
    }
  }
  getOrCreateSheet_(SHEET_NAMES.WORKERS, ['작업자명']);
  return '메인 시트 초기화 완료';
}

// 프로젝트 스프레드시트 생성 (모드별)
function createProjectSpreadsheet_(projectName, projectId, schema) {
  const folder = getOrCreateProjectFolder_();
  const fileName = `[검수] ${projectName} (${projectId})`;
  const ss = SpreadsheetApp.create(fileName);

  const file = DriveApp.getFileById(ss.getId());
  folder.addFile(file);
  DriveApp.getRootFolder().removeFile(file);

  const sheets = ss.getSheets();
  const firstSheet = sheets[0];
  firstSheet.setName('Data');

  // Data 헤더는 모드에 따라 다름
  if (schema.mode === 'standard') {
    firstSheet.getRange(1, 1, 1, STANDARD_HEADERS.length).setValues([STANDARD_HEADERS]);
  } else {
    // 자유 모드는 일단 빈 헤더로 시작 - 업로드 시 채워짐
    firstSheet.getRange(1, 1).setValue('(자유 모드 - 데이터 업로드 시 컬럼이 자동 설정됩니다)');
  }
  firstSheet.setFrozenRows(1);

  // Results 시트 - 동적 컬럼이라 키만 고정
  const resultsSheet = ss.insertSheet('Results');
  resultsSheet.getRange(1, 1, 1, 7).setValues([['키', '식별값1', '식별값2', '검수결과', '오매핑유형', '메모', '작업자', '작업시간'].slice(0, 7)]);
  // 실제로는 8개 컬럼이 필요하므로 다시 설정
  resultsSheet.getRange(1, 1, 1, 8).setValues([['키', '식별값1', '식별값2', '검수결과', '오매핑유형', '메모', '작업자', '작업시간']]);
  resultsSheet.setFrozenRows(1);
  resultsSheet.setColumnWidth(1, 80);
  resultsSheet.setColumnWidth(2, 250);
  resultsSheet.setColumnWidth(3, 250);
  resultsSheet.setColumnWidth(6, 200);

  const assignSheet = ss.insertSheet('Assignments');
  assignSheet.getRange(1, 1, 1, 3).setValues([['작업자', '시작', '끝']]);
  assignSheet.setFrozenRows(1);

  const reasonsSheet = ss.insertSheet('Reasons');
  reasonsSheet.getRange(1, 1, 1, 2).setValues([['사유명', '단축키']]);
  reasonsSheet.setFrozenRows(1);
  const reasonRows = DEFAULT_REASONS.map(r => [r.name, r.shortcut]);
  reasonsSheet.getRange(2, 1, reasonRows.length, 2).setValues(reasonRows);

  const logsSheet = ss.insertSheet('Logs');
  logsSheet.getRange(1, 1, 1, 4).setValues([['타임스탬프', '작업자', '키', '소요초']]);
  logsSheet.setFrozenRows(1);

  SpreadsheetApp.flush();
  return ss.getId();
}

// 프로젝트의 스프레드시트 ID 가져오기 (없으면 마이그레이션)
function getProjectSpreadsheetId_(projectId) {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === projectId) {
      const ssId = String(values[i][3] || '');
      if (ssId) {
        try {
          DriveApp.getFileById(ssId);
          return ssId;
        } catch (e) {
          // 파일 삭제됨 → 다시 생성
        }
      }
      // 마이그레이션 필요
      const projectName = String(values[i][1] || projectId);
      const newSsId = createProjectSpreadsheet_(projectName, projectId, STANDARD_SCHEMA);
      sheet.getRange(i + 1, 4).setValue(newSsId);
      sheet.getRange(i + 1, 5).setValue('standard');
      sheet.getRange(i + 1, 6).setValue(JSON.stringify(STANDARD_SCHEMA));
      return newSsId;
    }
  }
  throw new Error('프로젝트를 찾을 수 없습니다: ' + projectId);
}

// 프로젝트 정보 (모드/스키마 포함)
function getProjectInfo_(projectId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === projectId) {
      const mode = String(values[i][4] || 'standard');
      let schema;
      try {
        const sj = String(values[i][5] || '');
        schema = sj ? JSON.parse(sj) : STANDARD_SCHEMA;
      } catch(e) {
        schema = STANDARD_SCHEMA;
      }
      return {
        id: projectId,
        name: String(values[i][1] || ''),
        createdAt: String(values[i][2] || ''),
        spreadsheetId: String(values[i][3] || ''),
        mode: mode,
        schema: schema,
        rowIdx: i + 1
      };
    }
  }
  return null;
}

// ============ 인증 ============
function verifyAdminPassword(pw) {
  return pw === ADMIN_PASSWORD;
}

// ============ 프로젝트 목록 ============
function getProjects_() {
  initSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  const values = sheet.getDataRange().getValues();
  const projects = [];
  for (let i = 1; i < values.length; i++) {
    const [id, name, createdAt, spreadsheetId, mode, schemaJson] = values[i];
    if (id) {
      let schema;
      try {
        schema = schemaJson ? JSON.parse(schemaJson) : STANDARD_SCHEMA;
      } catch(e) {
        schema = STANDARD_SCHEMA;
      }
      projects.push({
        id: String(id),
        name: String(name || ''),
        createdAt: String(createdAt || ''),
        spreadsheetId: String(spreadsheetId || ''),
        mode: String(mode || 'standard'),
        schema: schema
      });
    }
  }
  return projects;
}

function getWorkersList_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
  const values = sheet.getDataRange().getValues();
  const workers = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i][0]) workers.push(String(values[i][0]));
  }
  return workers;
}

function getInitialState() {
  initSheets();
  return {
    projects: getProjects_(),
    workers: getWorkersList_()
  };
}

function getProjectState(payload) {
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '존재하지 않는 프로젝트입니다' };

  const projectSsId = getProjectSpreadsheetId_(projectId);
  const projectSs = SpreadsheetApp.openById(projectSsId);

  // Data 로드
  const dataSheet = projectSs.getSheetByName('Data');
  const dataValues = dataSheet.getDataRange().getValues();
  const data = [];
  let dataHeaders = [];

  if (proj.mode === 'free') {
    // 자유 모드: 1행=그룹, 2행=헤더, 3행~=데이터
    if (dataValues.length >= 2) {
      dataHeaders = dataValues[1].map(String);
      for (let i = 2; i < dataValues.length; i++) {
        const row = {};
        let hasContent = false;
        dataHeaders.forEach((h, j) => {
          const v = String(dataValues[i][j] || '');
          row[h] = v;
          if (v) hasContent = true;
        });
        if (hasContent) data.push(row);
      }
    }
  } else {
    // 표준 모드
    if (dataValues.length > 1) {
      dataHeaders = dataValues[0].map(String);
      for (let i = 1; i < dataValues.length; i++) {
        const row = {};
        dataHeaders.forEach((h, j) => { row[h] = String(dataValues[i][j] || ''); });
        if (row['상품번호'] || row['C_상품번호']) data.push(row);
      }
    } else {
      dataHeaders = STANDARD_HEADERS.slice();
    }
  }

  // Results 로드 (헤더 기반)
  const resultsSheet = projectSs.getSheetByName('Results');
  const resultsValues = resultsSheet.getDataRange().getValues();
  const results = {};
  if (resultsValues.length > 0) {
    const rh = resultsValues[0].map(String);
    const cKey = findHeader_(rh, ['키']);
    const cDecision = findHeader_(rh, ['검수결과']);
    const cReason = findHeader_(rh, ['오매핑유형']);
    const cMemo = findHeader_(rh, ['메모']);
    const cWorker = findHeader_(rh, ['작업자']);
    const cTimestamp = findHeader_(rh, ['작업시간']);
    for (let i = 1; i < resultsValues.length; i++) {
      const row = resultsValues[i];
      const key = cKey >= 0 ? String(row[cKey] || '') : '';
      if (!key) continue;
      results[key] = {
        decision: cDecision >= 0 ? String(row[cDecision] || '') : '',
        reason: cReason >= 0 ? String(row[cReason] || '') : '',
        memo: cMemo >= 0 ? String(row[cMemo] || '') : '',
        worker: cWorker >= 0 ? String(row[cWorker] || '') : '',
        timestamp: cTimestamp >= 0 ? String(row[cTimestamp] || '') : ''
      };
    }
  }

  // Assignments
  const assignSheet = projectSs.getSheetByName('Assignments');
  const assignValues = assignSheet.getDataRange().getValues();
  const assignments = [];
  for (let i = 1; i < assignValues.length; i++) {
    const [worker, start, end] = assignValues[i];
    if (worker) {
      assignments.push({
        worker: String(worker),
        start: parseInt(start) || 1,
        end: parseInt(end) || 1
      });
    }
  }

  // Reasons
  const reasonsSheet = projectSs.getSheetByName('Reasons');
  const reasonsValues = reasonsSheet.getDataRange().getValues();
  const reasons = [];
  for (let i = 1; i < reasonsValues.length; i++) {
    const [name, shortcut] = reasonsValues[i];
    if (name) {
      reasons.push({
        name: String(name),
        shortcut: String(shortcut || '')
      });
    }
  }

  return {
    ok: true,
    project: proj,
    schema: proj.schema,
    dataHeaders: dataHeaders,
    data,
    results,
    assignments,
    reasons,
    workers: getWorkersList_()
  };
}

function findHeader_(headers, candidates) {
  for (let i = 0; i < headers.length; i++) {
    if (candidates.indexOf(headers[i]) >= 0) return i;
  }
  return -1;
}

// ============ 프로젝트 관리 ============
function createProject(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const name = String(payload.name || '').trim();
  if (!name) return { ok: false, error: '프로젝트명을 입력해주세요' };
  const mode = (payload.mode === 'free') ? 'free' : 'standard';
  initSheets();
  const projects = getProjects_();
  if (projects.some(p => p.name === name)) {
    return { ok: false, error: '동일한 이름의 프로젝트가 이미 존재합니다' };
  }
  const id = 'p' + Date.now();
  const createdAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  const initialSchema = (mode === 'free')
    ? { mode: 'free', leftGroup: { name: '', columns: [] }, rightGroup: { name: '', columns: [] } }
    : STANDARD_SCHEMA;

  const newSsId = createProjectSpreadsheet_(name, id, initialSchema);

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  sheet.appendRow([id, name, createdAt, newSsId, mode, JSON.stringify(initialSchema)]);
  return { ok: true, projectId: id };
}

function renameProject(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  const newName = String(payload.newName || '').trim();
  if (!projectId || !newName) return { ok: false, error: '잘못된 입력' };
  const projects = getProjects_();
  if (projects.some(p => p.id !== projectId && p.name === newName)) {
    return { ok: false, error: '동일한 이름의 프로젝트가 이미 존재합니다' };
  }
  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '프로젝트를 찾을 수 없습니다' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  sheet.getRange(proj.rowIdx, 2).setValue(newName);
  try {
    if (proj.spreadsheetId) {
      const file = DriveApp.getFileById(proj.spreadsheetId);
      file.setName(`[검수] ${newName} (${projectId})`);
    }
  } catch(e) {}
  return { ok: true };
}

function deleteProject(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '프로젝트를 찾을 수 없습니다' };

  if (proj.spreadsheetId) {
    try {
      DriveApp.getFileById(proj.spreadsheetId).setTrashed(true);
    } catch(e) {}
  }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  sheet.deleteRow(proj.rowIdx);
  return { ok: true };
}

// ============ 데이터 업로드 ============
function uploadData(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '존재하지 않는 프로젝트' };

  const rows = payload.rows;
  if (!rows || rows.length === 0) return { ok: false, error: '데이터가 비어있습니다' };

  const appendMode = payload.appendMode === true;

  // 덮어쓰기 + Results 초기화 옵션
  if (!appendMode && payload.resetResults === true) {
    const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
    const sheet = projectSs.getSheetByName('Results');
    sheet.clear();
    sheet.getRange(1, 1, 1, 8).setValues([['키', '식별값1', '식별값2', '검수결과', '오매핑유형', '메모', '작업자', '작업시간']]);
    sheet.setFrozenRows(1);
    // Logs는 유지
  }

  if (proj.mode === 'free') {
    return uploadDataFree_(proj, rows, payload.merges || [], appendMode, payload);
  } else {
    return uploadDataStandard_(proj, rows, appendMode);
  }
}

function uploadDataStandard_(proj, rows, appendMode) {
  const headers = rows[0].map(String);
  const required = ['상품번호', '상품명', '네이버_URL', 'C_상품번호', 'C_상품명', 'C_URL'];
  const missing = required.filter(c => headers.indexOf(c) === -1);
  if (missing.length > 0) return { ok: false, error: '필수 컬럼이 없습니다: ' + missing.join(', ') };

  const colIdx = STANDARD_HEADERS.map(h => headers.indexOf(h));
  const newDataRows = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    const isEmpty = r.every(c => c === '' || c === null || c === undefined);
    if (isEmpty) continue;
    newDataRows.push(colIdx.map(idx => idx >= 0 ? String(r[idx] || '') : ''));
  }

  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(proj.id));
  const sheet = projectSs.getSheetByName('Data');

  if (appendMode) {
    const lastRow = sheet.getLastRow();
    if (lastRow < 1) {
      const allRows = [STANDARD_HEADERS.slice(), ...newDataRows];
      sheet.getRange(1, 1, allRows.length, STANDARD_HEADERS.length).setValues(allRows);
      sheet.setFrozenRows(1);
    } else if (newDataRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, newDataRows.length, STANDARD_HEADERS.length).setValues(newDataRows);
    }
  } else {
    const normalizedRows = [STANDARD_HEADERS.slice(), ...newDataRows];
    sheet.clear();
    sheet.getRange(1, 1, normalizedRows.length, STANDARD_HEADERS.length).setValues(normalizedRows);
    sheet.setFrozenRows(1);
  }

  return { ok: true, count: newDataRows.length };
}

/**
 * 자유 모드 업로드:
 * - rows[0]: 그룹 헤더 (병합된 셀의 1행)
 * - rows[1]: 실제 컬럼 헤더 (2행)
 * - rows[2]~: 데이터
 * - merges: [{startCol, endCol, value}, ...] - 1행의 병합 정보
 *
 * 그룹 분리: merges 정보로 좌/우 그룹 자동 추출
 */
function uploadDataFree_(proj, rows, merges, appendMode, payload) {
  if (rows.length < 2) return { ok: false, error: '자유 모드는 최소 2행(그룹헤더 + 컬럼헤더)이 필요합니다' };

  const groupRow = rows[0].map(v => String(v == null ? '' : v));
  const headerRow = rows[1].map(v => String(v == null ? '' : v));

  if (headerRow.every(h => !h)) {
    return { ok: false, error: '2행에 컬럼 이름이 없습니다' };
  }

  // 헤더 행의 마지막 비어있지 않은 컬럼 위치 찾기 (전체 컬럼 수 결정)
  let totalCols = headerRow.length;
  for (let i = headerRow.length - 1; i >= 0; i--) {
    if (headerRow[i] && headerRow[i].trim()) { totalCols = i + 1; break; }
  }

  Logger.log('uploadDataFree_ 디버그');
  Logger.log('groupRow: ' + JSON.stringify(groupRow));
  Logger.log('headerRow: ' + JSON.stringify(headerRow));
  Logger.log('merges: ' + JSON.stringify(merges));
  Logger.log('totalCols: ' + totalCols);

  // === 그룹 경계 결정 ===
  let groups = [];

  // 1) 병합 정보 우선 사용 (1행의 병합만)
  if (merges && merges.length > 0) {
    const rowOneMerges = merges.filter(m => m.row === 0).sort((a, b) => a.startCol - b.startCol);
    if (rowOneMerges.length >= 2) {
      rowOneMerges.forEach(m => {
        const v = String(m.value || groupRow[m.startCol] || '').trim();
        groups.push({
          name: v || `그룹${groups.length + 1}`,
          startCol: m.startCol,
          endCol: m.endCol
        });
      });
    } else if (rowOneMerges.length === 1) {
      const m = rowOneMerges[0];
      groups.push({
        name: String(m.value || groupRow[m.startCol] || '그룹1').trim(),
        startCol: m.startCol,
        endCol: m.endCol
      });
      for (let i = 0; i < totalCols; i++) {
        if (i >= m.startCol && i <= m.endCol) continue;
        const v = groupRow[i] && groupRow[i].trim();
        if (v) {
          let endC = i;
          for (let j = i + 1; j < totalCols; j++) {
            if (j >= m.startCol && j <= m.endCol) break;
            endC = j;
          }
          groups.push({ name: v, startCol: i, endCol: endC });
          break;
        }
      }
      groups.sort((a, b) => a.startCol - b.startCol);
    }
  }

  // 2) 병합 정보가 부족하면 1행의 비어있지 않은 셀 위치로 추론
  if (groups.length < 2) {
    groups = [];
    const namedCells = [];
    for (let i = 0; i < totalCols; i++) {
      const v = groupRow[i] && groupRow[i].trim();
      if (v) namedCells.push({ col: i, name: v });
    }

    if (namedCells.length >= 2) {
      for (let i = 0; i < namedCells.length; i++) {
        const startCol = namedCells[i].col;
        const endCol = (i < namedCells.length - 1) ? namedCells[i + 1].col - 1 : totalCols - 1;
        groups.push({ name: namedCells[i].name, startCol, endCol });
      }
    }
  }

  // 3) 그래도 부족하면 2행 컬럼을 단순히 절반으로 나누기 (최후 수단)
  if (groups.length < 2) {
    const half = Math.floor(totalCols / 2);
    if (half >= 1 && totalCols - half >= 1) {
      groups = [
        { name: '왼쪽', startCol: 0, endCol: half - 1 },
        { name: '오른쪽', startCol: half, endCol: totalCols - 1 }
      ];
    } else {
      return { ok: false, error: '그룹이 2개 이상 필요합니다. 1행에 그룹명을 병합 셀로 만들어주세요.' };
    }
  }

  Logger.log('인식된 그룹: ' + JSON.stringify(groups));

  const leftGroup = groups[0];
  const rightGroup = groups[1];

  const leftCols = [];
  const leftOrigIdx = [];
  for (let c = leftGroup.startCol; c <= leftGroup.endCol; c++) {
    const h = headerRow[c];
    if (h && h.trim()) {
      leftCols.push(h.trim());
      leftOrigIdx.push(c);
    }
  }
  const rightCols = [];
  const rightOrigIdx = [];
  for (let c = rightGroup.startCol; c <= rightGroup.endCol; c++) {
    const h = headerRow[c];
    if (h && h.trim()) {
      rightCols.push(h.trim());
      rightOrigIdx.push(c);
    }
  }

  Logger.log('leftCols: ' + JSON.stringify(leftCols));
  Logger.log('rightCols: ' + JSON.stringify(rightCols));

  if (leftCols.length === 0 || rightCols.length === 0) {
    return { ok: false, error: `각 그룹에 최소 1개의 컬럼이 필요합니다 (좌: ${leftCols.length}개, 우: ${rightCols.length}개). 2행에 컬럼명을 모두 입력했는지 확인해주세요.` };
  }

  // 중복 컬럼명 자동 처리
  const seenCols = new Set();
  function uniqueName(name) {
    if (!seenCols.has(name)) { seenCols.add(name); return name; }
    let i = 2;
    while (seenCols.has(name + '_' + i)) i++;
    const out = name + '_' + i;
    seenCols.add(out);
    return out;
  }
  const leftColsUnique = leftCols.map(uniqueName);
  const rightColsUnique = rightCols.map(uniqueName);
  const allCols = [...leftColsUnique, ...rightColsUnique];

  const schema = {
    mode: 'free',
    leftGroup: { name: leftGroup.name || '타겟', columns: leftColsUnique },
    rightGroup: { name: rightGroup.name || '매칭', columns: rightColsUnique }
  };

  // 데이터 행 파싱 (시트 접근 전)
  const dataRows = [];
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i] || [];
    const isEmpty = r.every(c => c === '' || c === null || c === undefined);
    if (isEmpty) continue;
    const out = [];
    leftOrigIdx.forEach(idx => out.push(String(r[idx] == null ? '' : r[idx])));
    rightOrigIdx.forEach(idx => out.push(String(r[idx] == null ? '' : r[idx])));
    dataRows.push(out);
  }

  // 추가 모드: 기존 스키마와 컬럼 비교 후 데이터만 append
  if (appendMode) {
    const existingSchema = proj.schema;
    if (existingSchema && existingSchema.leftGroup && existingSchema.leftGroup.columns.length > 0) {
      const existingCols = [...existingSchema.leftGroup.columns, ...existingSchema.rightGroup.columns];
      // fileHeaderRowIdx: 0 = 단일 헤더 포맷(1행=컬럼명, 2행~=데이터)
      //                   1 = 이중 헤더 포맷(1행=그룹헤더, 2행=컬럼명, 3행~=데이터)
      const fileHeaderRowIdx = (payload && payload.fileHeaderRowIdx != null) ? Number(payload.fileHeaderRowIdx) : 1;
      const fileHdr = (rows[fileHeaderRowIdx] || []).map(v => String(v == null ? '' : v));

      // ── 사용자가 확인한 컬럼 매핑이 있는 경우 ──────────────
      const colMapping = payload && payload.columnMapping;
      if (colMapping && Array.isArray(colMapping) && colMapping.length === existingCols.length) {
        const mappedRows = [];
        for (let i = fileHeaderRowIdx + 1; i < rows.length; i++) {
          const r = rows[i] || [];
          if (r.every(c => c === '' || c === null || c === undefined)) continue;
          const out = colMapping.map(fileColName => {
            if (!fileColName || fileColName === '__skip__') return '';
            const idx = fileHdr.indexOf(fileColName);
            return idx >= 0 ? String(r[idx] == null ? '' : r[idx]) : '';
          });
          mappedRows.push(out);
        }
        const projectSsM = SpreadsheetApp.openById(getProjectSpreadsheetId_(proj.id));
        const dataSheetM = projectSsM.getSheetByName('Data');
        const lastRowM = dataSheetM.getLastRow();
        if (mappedRows.length > 0) {
          dataSheetM.getRange(lastRowM + 1, 1, mappedRows.length, existingCols.length).setValues(mappedRows);
        }
        return { ok: true, count: mappedRows.length, schema: existingSchema };
      }

      // ── 매핑 없음: 파일 헤더가 기존 스키마와 완전히 일치해야 함 ──────────────
      const fileCols = fileHdr.filter(h => h && h.trim());
      if (existingCols.join(',') !== fileCols.join(',')) {
        return { ok: false, error: `컬럼 구조가 기존과 다릅니다.\n기존: [${existingCols.join(', ')}]\n신규: [${fileCols.join(', ')}]` };
      }
      const projectSsA = SpreadsheetApp.openById(getProjectSpreadsheetId_(proj.id));
      const dataSheetA = projectSsA.getSheetByName('Data');
      const lastRow = dataSheetA.getLastRow();
      const appendedRows = [];
      for (let i = fileHeaderRowIdx + 1; i < rows.length; i++) {
        const r = rows[i] || [];
        if (r.every(c => c === '' || c === null || c === undefined)) continue;
        const out = existingCols.map(col => {
          const idx = fileHdr.indexOf(col);
          return idx >= 0 ? String(r[idx] == null ? '' : r[idx]) : '';
        });
        appendedRows.push(out);
      }
      if (appendedRows.length > 0) {
        dataSheetA.getRange(lastRow + 1, 1, appendedRows.length, existingCols.length).setValues(appendedRows);
      }
      return { ok: true, count: appendedRows.length, schema: existingSchema };
    }
    // 기존 스키마 없으면 덮어쓰기로 처리
  }

  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(proj.id));
  const dataSheet = projectSs.getSheetByName('Data');
  dataSheet.clear();
  dataSheet.clearFormats();

  // 1행: 그룹 헤더 (병합)
  const groupHeaderRow = [];
  for (let i = 0; i < leftColsUnique.length; i++) {
    groupHeaderRow.push(i === 0 ? schema.leftGroup.name : '');
  }
  for (let i = 0; i < rightColsUnique.length; i++) {
    groupHeaderRow.push(i === 0 ? schema.rightGroup.name : '');
  }
  dataSheet.getRange(1, 1, 1, allCols.length).setValues([groupHeaderRow]);
  if (leftColsUnique.length > 1) {
    dataSheet.getRange(1, 1, 1, leftColsUnique.length).merge();
  }
  if (rightColsUnique.length > 1) {
    dataSheet.getRange(1, leftColsUnique.length + 1, 1, rightColsUnique.length).merge();
  }
  dataSheet.getRange(1, 1, 1, allCols.length)
    .setHorizontalAlignment('center').setBackground('#E6F1FB').setFontWeight('bold');

  // 2행: 컬럼 헤더
  dataSheet.getRange(2, 1, 1, allCols.length).setValues([allCols]);
  dataSheet.getRange(2, 1, 1, allCols.length).setFontWeight('bold').setBackground('#F1EFE8');

  // 3행~: 데이터
  if (dataRows.length > 0) {
    dataSheet.getRange(3, 1, dataRows.length, allCols.length).setValues(dataRows);
  }
  dataSheet.setFrozenRows(2);

  // 스키마 저장
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const projSheet = ss.getSheetByName(SHEET_NAMES.PROJECTS);
  projSheet.getRange(proj.rowIdx, 6).setValue(JSON.stringify(schema));

  return { ok: true, count: dataRows.length, schema: schema };
}

// ============ 검수 결과 저장 ============
function saveResult(payload) {
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const sheet = projectSs.getSheetByName('Results');
  const values = sheet.getDataRange().getValues();
  const key = String(payload.key);
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  const rh = (values[0] || []).map(String);
  const cKey = findHeader_(rh, ['키']);
  const cIdent1 = findHeader_(rh, ['식별값1']);
  const cIdent2 = findHeader_(rh, ['식별값2']);
  const cDecision = findHeader_(rh, ['검수결과']);
  const cReason = findHeader_(rh, ['오매핑유형']);
  const cMemo = findHeader_(rh, ['메모']);
  const cWorker = findHeader_(rh, ['작업자']);
  const cTime = findHeader_(rh, ['작업시간']);

  let rowIdx = -1;
  if (cKey >= 0) {
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][cKey]) === key) { rowIdx = i + 1; break; }
    }
  }

  const isClearing = !payload.decision || String(payload.decision).trim() === '';
  if (isClearing) {
    if (rowIdx > 0) sheet.deleteRow(rowIdx);
    const logsSheet = projectSs.getSheetByName('Logs');
    const lv = logsSheet.getDataRange().getValues();
    const lh = (lv[0] || []).map(String);
    const lcKey = findHeader_(lh, ['키']);
    if (lcKey >= 0) {
      for (let i = lv.length - 1; i >= 1; i--) {
        if (String(lv[i][lcKey]) === key) logsSheet.deleteRow(i + 1);
      }
    }
    return { ok: true, cleared: true };
  }

  const ident1 = String(payload.ident1 || '');
  const ident2 = String(payload.ident2 || '');

  const newRow = new Array(8);
  if (cKey >= 0) newRow[cKey] = key;
  if (cIdent1 >= 0) newRow[cIdent1] = ident1;
  if (cIdent2 >= 0) newRow[cIdent2] = ident2;
  if (cDecision >= 0) newRow[cDecision] = payload.decision || '';
  if (cReason >= 0) newRow[cReason] = payload.reason || '';
  if (cMemo >= 0) newRow[cMemo] = payload.memo || '';
  if (cWorker >= 0) newRow[cWorker] = payload.worker || '';
  if (cTime >= 0) newRow[cTime] = timestamp;
  for (let i = 0; i < newRow.length; i++) if (newRow[i] === undefined) newRow[i] = '';

  let isFirstDecision = true;
  if (rowIdx > 0 && cDecision >= 0) {
    const existing = String(values[rowIdx - 1][cDecision] || '');
    if (existing) isFirstDecision = false;
  }

  if (rowIdx > 0) {
    sheet.getRange(rowIdx, 1, 1, newRow.length).setValues([newRow]);
  } else {
    sheet.appendRow(newRow);
  }

  if (isFirstDecision && payload.decision && payload.elapsedSec && payload.elapsedSec > 0 && payload.elapsedSec < 3600) {
    const logsSheet = projectSs.getSheetByName('Logs');
    logsSheet.appendRow([timestamp, payload.worker || '', key, payload.elapsedSec]);
  }
  return { ok: true, timestamp };
}

// ============ 작업자 풀 ============
function saveWorkers(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.WORKERS);
  sheet.clear();
  sheet.getRange(1, 1).setValue('작업자명');
  if (payload.workers && payload.workers.length > 0) {
    const rows = payload.workers.map(w => [String(w)]);
    sheet.getRange(2, 1, rows.length, 1).setValues(rows);
  }
  sheet.setFrozenRows(1);
  return { ok: true };
}

// ============ 할당 ============
function saveAssignments(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const sheet = projectSs.getSheetByName('Assignments');
  sheet.clear();
  sheet.getRange(1, 1, 1, 3).setValues([['작업자', '시작', '끝']]);
  if (payload.assignments && payload.assignments.length > 0) {
    const rows = payload.assignments.map(a => [String(a.worker), parseInt(a.start), parseInt(a.end)]);
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
  sheet.setFrozenRows(1);
  return { ok: true };
}

// ============ 사유 ============
function saveReasons(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const reasons = payload.reasons || [];
  const names = reasons.map(r => r.name);
  if (names.length !== new Set(names).size) return { ok: false, error: '중복된 사유명이 있습니다' };
  if (reasons.some(r => !r.name || !r.name.trim())) return { ok: false, error: '빈 사유명이 있습니다' };

  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const sheet = projectSs.getSheetByName('Reasons');
  sheet.clear();
  sheet.getRange(1, 1, 1, 2).setValues([['사유명', '단축키']]);
  if (reasons.length > 0) {
    const rows = reasons.map(r => [String(r.name).trim(), String(r.shortcut || '').trim()]);
    sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  }
  sheet.setFrozenRows(1);
  return { ok: true };
}

function renameReason(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const oldName = String(payload.oldName || '');
  const newName = String(payload.newName || '');
  if (!oldName || !newName || oldName === newName) return { ok: true, updated: 0 };

  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const sheet = projectSs.getSheetByName('Results');
  const values = sheet.getDataRange().getValues();
  if (values.length < 1) return { ok: true, updated: 0 };
  const rh = values[0].map(String);
  const cReason = findHeader_(rh, ['오매핑유형']);
  if (cReason < 0) return { ok: true, updated: 0 };
  let updated = 0;
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][cReason]) === oldName) {
      sheet.getRange(i + 1, cReason + 1).setValue(newName);
      updated++;
    }
  }
  return { ok: true, updated };
}

// ============ 결과 초기화 ============
function resetResults(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));

  const sheet = projectSs.getSheetByName('Results');
  sheet.clear();
  sheet.getRange(1, 1, 1, 8).setValues([['키', '식별값1', '식별값2', '검수결과', '오매핑유형', '메모', '작업자', '작업시간']]);
  sheet.setFrozenRows(1);

  const logsSheet = projectSs.getSheetByName('Logs');
  logsSheet.clear();
  logsSheet.getRange(1, 1, 1, 4).setValues([['타임스탬프', '작업자', '키', '소요초']]);
  logsSheet.setFrozenRows(1);
  return { ok: true };
}

// ============ 통계 ============
function getWorkerStats(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const startDate = payload.startDate || null;
  const endDate = payload.endDate || null;
  return { ok: true, stats: computeWorkerStats_(projectId, startDate, endDate) };
}

function computeWorkerStats_(projectId, startDate, endDate) {
  const proj = getProjectInfo_(projectId);
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const logsSheet    = projectSs.getSheetByName('Logs');
  const resultsSheet = projectSs.getSheetByName('Results');
  const dataSheet    = projectSs.getSheetByName('Data');

  // ── Data 시트에서 key → Collected date 맵 생성 ──────────
  // "Collected date" 컬럼을 대소문자/공백 무시하고 탐색
  const keyToCollectedDate = {};
  const hasDateFilter = !!(startDate || endDate);
  if (hasDateFilter) {
    const dv = dataSheet.getDataRange().getValues();
    const isFree = proj && proj.mode === 'free';
    const headerCount = isFree ? 2 : 1;
    if (dv.length > headerCount) {
      const dataHeaders = dv[headerCount - 1].map(String);
      const cdIdx = dataHeaders.findIndex(
        h => h.trim().toLowerCase().replace(/\s+/g, '') === 'collecteddate'
      );
      if (cdIdx >= 0) {
        for (let i = headerCount; i < dv.length; i++) {
          const key = String(i - headerCount + 1); // 1-based
          const val = dv[i][cdIdx];
          keyToCollectedDate[key] = val ? String(val).slice(0, 10) : '';
        }
      }
    }
  }

  // key 기준으로 Collected date 범위 체크
  function inRange(key) {
    if (!startDate && !endDate) return true;
    const d = keyToCollectedDate[String(key)] || '';
    if (!d) return false;
    if (startDate && d < startDate) return false;
    if (endDate && d > endDate) return false;
    return true;
  }

  // ── Logs: 소요시간 집계 ──────────────────────────────────
  const lv = logsSheet.getDataRange().getValues();
  const workerLogs = {};
  if (lv.length > 0) {
    const lh = lv[0].map(String);
    const lcW   = findHeader_(lh, ['작업자']);
    const lcS   = findHeader_(lh, ['소요초']);
    const lcKey = findHeader_(lh, ['키']);
    if (lcW >= 0 && lcS >= 0) {
      for (let i = 1; i < lv.length; i++) {
        const w   = String(lv[i][lcW] || '');
        const s   = parseFloat(lv[i][lcS]) || 0;
        const key = lcKey >= 0 ? String(lv[i][lcKey] || '') : '';
        if (!w || s <= 0) continue;
        if (!inRange(key)) continue;
        if (!workerLogs[w]) workerLogs[w] = [];
        workerLogs[w].push(s);
      }
    }
  }

  // ── Results: 일치/불일치 집계 ────────────────────────────
  const rv = resultsSheet.getDataRange().getValues();
  const workerCounts = {};
  if (rv.length > 0) {
    const rh   = rv[0].map(String);
    const cD   = findHeader_(rh, ['검수결과']);
    const cW   = findHeader_(rh, ['작업자']);
    const cKey = findHeader_(rh, ['키']);
    if (cD >= 0 && cW >= 0) {
      for (let i = 1; i < rv.length; i++) {
        const decision = String(rv[i][cD] || '');
        const w        = String(rv[i][cW] || '');
        const key      = cKey >= 0 ? String(rv[i][cKey] || '') : '';
        if (!w || !decision) continue;
        if (!inRange(key)) continue;
        if (!workerCounts[w]) workerCounts[w] = { match: 0, mismatch: 0, total: 0 };
        workerCounts[w].total++;
        if (decision === '일치') workerCounts[w].match++;
        else if (decision === '불일치') workerCounts[w].mismatch++;
      }
    }
  }

  const allWorkers = new Set([...Object.keys(workerLogs), ...Object.keys(workerCounts)]);
  const stats = [];
  allWorkers.forEach(w => {
    const logs = workerLogs[w] || [];
    const counts = workerCounts[w] || { match: 0, mismatch: 0, total: 0 };
    let medianSec = 0;
    if (logs.length > 0) {
      const sorted = logs.slice().sort((a, b) => a - b);
      const mid = Math.floor(sorted.length / 2);
      medianSec = sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
    }
    const perHour = medianSec > 0 ? Math.round(3600 / medianSec) : 0;
    const matchRate = counts.total > 0 ? Math.round(counts.match / counts.total * 100) : 0;
    stats.push({
      worker: w,
      total: counts.total,
      match: counts.match,
      mismatch: counts.mismatch,
      matchRate: matchRate,
      avgSec: Math.round(medianSec * 10) / 10,
      perHour: perHour,
      sampleCount: logs.length
    });
  });
  stats.sort((a, b) => b.total - a.total);
  return stats;
}

// ============ 통계 시트 업데이트 ============
/**
 * 프로젝트 Google Sheet 내 '통계' 탭을 현재 데이터로 업데이트한다.
 * 내용: 검수 결과 통계 / 처리주체별 정확도 / 전체 진행 현황 / 작업자별 성과
 */
function updateStatsSheet(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };

  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '프로젝트를 찾을 수 없습니다' };

  const projectSs   = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const dataSheet    = projectSs.getSheetByName('Data');
  const resultsSheet = projectSs.getSheetByName('Results');
  const assignSheet  = projectSs.getSheetByName('Assignments');

  // ── 데이터 로드 ────────────────────────────────────────
  const dv = dataSheet.getDataRange().getValues();
  let dataHeaders = [], dataRows = [];
  if (proj.mode === 'free' && dv.length >= 2) {
    dataHeaders = dv[1].map(String);
    dataRows = dv.slice(2).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
  } else if (dv.length > 0) {
    dataHeaders = dv[0].map(String);
    dataRows = dv.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
  }

  // 처리주체 컬럼 인덱스 (대소문자·공백 무시)
  const procIdx = dataHeaders.findIndex(
    h => String(h || '').trim().toLowerCase().replace(/\s+/g, '') === '처리주체'
  );

  // ── 결과 로드 ──────────────────────────────────────────
  const rv = resultsSheet.getDataRange().getValues();
  const results = {};
  if (rv.length > 0) {
    const rh = rv[0].map(String);
    const cKey = findHeader_(rh, ['키']);
    const cD   = findHeader_(rh, ['검수결과']);
    const cR   = findHeader_(rh, ['오매핑유형']);
    for (let i = 1; i < rv.length; i++) {
      const k = cKey >= 0 ? String(rv[i][cKey] || '') : '';
      if (!k) continue;
      results[k] = {
        decision: cD >= 0 ? String(rv[i][cD] || '') : '',
        reason:   cR >= 0 ? String(rv[i][cR] || '') : ''
      };
    }
  }

  // ── 할당 로드 ──────────────────────────────────────────
  const av = assignSheet.getDataRange().getValues();
  const assignments = [];
  for (let i = 1; i < av.length; i++) {
    if (av[i][0]) assignments.push({ worker: String(av[i][0]), start: parseInt(av[i][1]) || 1, end: parseInt(av[i][2]) || 1 });
  }

  // ── 통계 계산 ──────────────────────────────────────────
  const total = dataRows.length;
  let doneCount = 0, matchCount = 0, mismatchCount = 0;
  const reasonCounts = {};
  const procStats = {
    human:  { total: 0, done: 0, match: 0, mismatch: 0 },
    system: { total: 0, done: 0, match: 0, mismatch: 0 }
  };

  for (let i = 0; i < dataRows.length; i++) {
    const key  = String(i + 1);
    const res  = results[key];
    const proc = procIdx >= 0 ? String(dataRows[i][procIdx] || '').toLowerCase().trim() : '';

    if (proc === 'human' || proc === 'system') procStats[proc].total++;

    if (res && res.decision) {
      doneCount++;
      const isMatch = res.decision === '일치';
      if (isMatch) {
        matchCount++;
      } else if (res.decision === '불일치') {
        mismatchCount++;
        if (res.reason) reasonCounts[res.reason] = (reasonCounts[res.reason] || 0) + 1;
      }
      if (proc === 'human' || proc === 'system') {
        procStats[proc].done++;
        if (isMatch) procStats[proc].match++;
        else procStats[proc].mismatch++;
      }
    }
  }

  const decisionTotal = matchCount + mismatchCount;
  const ws = computeWorkerStats_(projectId);
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');

  // ── 통계 시트 초기화 ───────────────────────────────────
  let sh = projectSs.getSheetByName('통계');
  if (sh) { sh.clear(); sh.clearFormats(); }
  else     { sh = projectSs.insertSheet('통계'); }

  // ── 데이터 작성 ────────────────────────────────────────
  const COLS = 6;
  const rows = [];
  const sectionIdxs    = [];   // 섹션 헤더 행 (0-based)
  const tableHdrIdxs   = [];   // 테이블 헤더 행 (0-based)
  const labelColEndIdx = [];   // 섹션1 레이블 행 범위

  function pad(arr) {
    while (arr.length < COLS) arr.push('');
    return arr;
  }
  function push_(...args) { rows.push(pad([...args])); }
  function pushSection(label) { sectionIdxs.push(rows.length); push_(label); }
  function pushTHead(...args)  { tableHdrIdxs.push(rows.length); push_(...args); }

  // ── 섹션 1: 검수 결과 통계 ─────────────────────────────
  pushSection('📊 검수 결과 통계');
  const sec1Start = rows.length;
  push_('업데이트 시각', updatedAt);
  push_('프로젝트',      proj.name);
  push_('전체 데이터',   total);
  push_('검수 완료',     doneCount);
  push_('미완료',        total - doneCount);
  push_('일치',          matchCount);
  push_('불일치',        mismatchCount);
  push_('진행률 (%)',    total > 0 ? Math.round(doneCount / total * 100) : 0);
  push_('정확도 (일치율, %)', decisionTotal > 0 ? Math.round(matchCount / decisionTotal * 100) : 0);
  const sec1End = rows.length - 1;
  push_('');

  // ── 섹션 2: 처리주체별 정확도 (NEW) ────────────────────
  pushSection('🤖 처리주체별 정확도');
  pushTHead('처리주체', '전체 데이터', '검수 완료', '일치', '불일치', '정확도 (%)');
  ['human', 'system'].forEach(p => {
    const s   = procStats[p];
    const acc = s.done > 0 ? Math.round(s.match / s.done * 100) : 0;
    push_(p, s.total, s.done, s.match, s.mismatch, acc);
  });
  push_('');

  // ── 섹션 3: 불일치 사유 분포 ────────────────────────────
  pushSection('❌ 불일치 사유 분포 (건수 많은 순)');
  pushTHead('사유명', '건수', '비율 (%)');
  const sortedReasons = Object.entries(reasonCounts).sort((a, b) => b[1] - a[1]);
  if (sortedReasons.length === 0) {
    push_('(불일치 항목 없음)');
  } else {
    sortedReasons.forEach(([n, c]) => {
      push_(n, c, mismatchCount > 0 ? Math.round(c / mismatchCount * 100) : 0);
    });
  }
  push_('');

  // ── 섹션 4: 전체 진행 현황 ─────────────────────────────
  pushSection('👥 전체 진행 현황');
  pushTHead('작업자', '범위', '완료', '할당', '진행률 (%)');
  if (assignments.length === 0) {
    push_('(할당된 작업자가 없습니다)');
  } else {
    const byWorker = {};
    assignments.forEach(a => {
      if (!byWorker[a.worker]) byWorker[a.worker] = [];
      byWorker[a.worker].push(a);
    });
    Object.keys(byWorker).forEach(w => {
      const ranges = byWorker[w];
      const indices = new Set();
      ranges.forEach(r => { for (let i = Math.max(1, r.start); i <= Math.min(dataRows.length, r.end); i++) indices.add(i); });
      const t = indices.size;
      let d = 0;
      indices.forEach(i => { if (results[String(i)] && results[String(i)].decision) d++; });
      push_(w, ranges.map(r => `${r.start}-${r.end}`).join(', '), d, t, t > 0 ? Math.round(d / t * 100) : 0);
    });
  }
  push_('');

  // ── 섹션 5: 작업자별 성과 ─────────────────────────────
  pushSection('⏱ 작업자별 성과');
  pushTHead('작업자', '처리 건수', '일치율 (%)', '중간 작업시간 (초)', '시간당 처리량 (건)');
  if (ws.length === 0) {
    push_('(데이터 없음)');
  } else {
    ws.forEach(s => push_(s.worker, s.total, s.matchRate, s.avgSec > 0 ? s.avgSec : '-', s.perHour > 0 ? s.perHour : '-'));
  }

  sh.getRange(1, 1, rows.length, COLS).setValues(rows);

  // ── 서식 적용 ──────────────────────────────────────────
  sectionIdxs.forEach(idx => {
    sh.getRange(idx + 1, 1, 1, COLS)
      .setFontWeight('bold').setFontSize(12)
      .setBackground('#E6F1FB').setFontColor('#042C53');
  });
  tableHdrIdxs.forEach(idx => {
    sh.getRange(idx + 1, 1, 1, COLS)
      .setFontWeight('bold').setBackground('#F0F0F0');
  });
  // 섹션1 레이블 열 (key 값)
  sh.getRange(sec1Start + 1, 1, sec1End - sec1Start + 1, 1).setFontWeight('bold');

  sh.setColumnWidth(1, 230);
  sh.setColumnWidth(2, 140);
  sh.setColumnWidth(3, 110);
  sh.setColumnWidth(4, 90);
  sh.setColumnWidth(5, 90);
  sh.setColumnWidth(6, 110);

  SpreadsheetApp.flush();
  return { ok: true, sheetUrl: 'https://docs.google.com/spreadsheets/d/' + projectSs.getId() };
}

// ============ 다운로드 ============
function getDownloadToken(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };
  const projects = getProjects_();
  if (!projects.some(p => p.id === projectId)) return { ok: false, error: '존재하지 않는 프로젝트' };
  const token = Utilities.getUuid();
  const cache = CacheService.getScriptCache();
  cache.put('dl_token_' + token, projectId, 300);
  const url = ScriptApp.getService().getUrl();
  return { ok: true, url: url + '?action=download&projectId=' + encodeURIComponent(projectId) + '&token=' + token };
}

function handleDownloadRequest_(params) {
  const projectId = String(params.projectId || '');
  const token = String(params.token || '');
  if (!projectId || !token) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;">잘못된 요청</p>');
  }
  const cache = CacheService.getScriptCache();
  if (cache.get('dl_token_' + token) !== projectId) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;">인증 실패 또는 토큰 만료(5분)</p>');
  }
  cache.remove('dl_token_' + token);
  try {
    return doDownloadRedirect_(projectId);
  } catch (err) {
    return HtmlService.createHtmlOutput('<p style="padding:40px;color:red;">오류: ' + String(err.message || err).replace(/</g, '&lt;') + '</p>');
  }
}

function doDownloadRedirect_(projectId) {
  const tempSsId = createTempExportSpreadsheet_(projectId);
  const url = 'https://docs.google.com/spreadsheets/d/' + tempSsId + '/export?format=xlsx';
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>다운로드</title>
<style>body{font-family:-apple-system,sans-serif;padding:40px;text-align:center;background:#fafaf7}
.box{max-width:480px;margin:0 auto;padding:24px;background:#fff;border-radius:12px;border:1px solid rgba(0,0,0,0.1)}
h2{margin:0 0 12px;font-size:16px}p{font-size:13px;color:#5f5e5a;margin:8px 0}
a.btn{display:inline-block;margin-top:16px;padding:10px 20px;background:#185FA5;color:#fff;text-decoration:none;border-radius:6px}</style></head>
<body><div class="box"><h2>📥 엑셀 파일 준비 완료</h2><p>아래 버튼을 클릭하여 다운로드하세요</p>
<a class="btn" href="${url}" target="_top" download>엑셀 파일 다운로드</a><p style="font-size:11px;color:#888780;margin-top:16px">자동 시작이 안 되면 위 버튼을 눌러주세요</p></div>
<script>setTimeout(function(){try{window.top.location.href='${url}'}catch(e){window.location.href='${url}'}},800)</script>
</body></html>`;
  return HtmlService.createHtmlOutput(html).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function createTempExportSpreadsheet_(projectId) {
  const proj = getProjectInfo_(projectId);
  const projName = proj ? proj.name : projectId;
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));

  const dataSheet = projectSs.getSheetByName('Data');
  const resultsSheet = projectSs.getSheetByName('Results');
  const assignSheet = projectSs.getSheetByName('Assignments');
  const reasonsSheet = projectSs.getSheetByName('Reasons');

  const dv = dataSheet.getDataRange().getValues();
  let dataHeaders = [];
  let dataRows = [];
  if (proj.mode === 'free' && dv.length >= 2) {
    dataHeaders = dv[1].map(String);
    dataRows = dv.slice(2).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
  } else if (dv.length > 0) {
    dataHeaders = dv[0].map(String);
    dataRows = dv.slice(1).filter(r => r.some(c => c !== '' && c !== null && c !== undefined));
  }

  const rv = resultsSheet.getDataRange().getValues();
  const results = {};
  if (rv.length > 0) {
    const rh = rv[0].map(String);
    const cKey = findHeader_(rh, ['키']);
    const cD = findHeader_(rh, ['검수결과']);
    const cR = findHeader_(rh, ['오매핑유형']);
    const cM = findHeader_(rh, ['메모']);
    const cW = findHeader_(rh, ['작업자']);
    const cT = findHeader_(rh, ['작업시간']);
    for (let i = 1; i < rv.length; i++) {
      const k = cKey >= 0 ? String(rv[i][cKey] || '') : '';
      if (!k) continue;
      results[k] = {
        decision: cD >= 0 ? String(rv[i][cD] || '') : '',
        reason: cR >= 0 ? String(rv[i][cR] || '') : '',
        memo: cM >= 0 ? String(rv[i][cM] || '') : '',
        worker: cW >= 0 ? String(rv[i][cW] || '') : '',
        timestamp: cT >= 0 ? String(rv[i][cT] || '') : ''
      };
    }
  }

  const av = assignSheet.getDataRange().getValues();
  const assignments = [];
  for (let i = 1; i < av.length; i++) {
    if (av[i][0]) {
      assignments.push({
        worker: String(av[i][0]),
        start: parseInt(av[i][1]) || 1,
        end: parseInt(av[i][2]) || 1
      });
    }
  }

  const ts = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  const tempSs = SpreadsheetApp.create(`[검수결과] ${projName} ${ts}`);

  const sheet1 = tempSs.getSheets()[0];
  sheet1.setName('검수결과');
  const mergedHeaders = ['행번호'].concat(dataHeaders).concat(['검수결과', '오매핑유형', '메모', '작업자', '작업시간']);
  const mergedRows = [mergedHeaders];
  dataRows.forEach((r, idx) => {
    const rowNum = idx + 1;
    const key = String(rowNum);
    const res = results[key] || {};
    const out = [rowNum].concat(r.map(v => v == null ? '' : v));
    out.push(res.decision || '', res.reason || '', res.memo || '', res.worker || '', res.timestamp || '');
    mergedRows.push(out);
  });
  if (mergedRows.length > 0) {
    sheet1.getRange(1, 1, mergedRows.length, mergedHeaders.length).setValues(mergedRows);
    sheet1.setFrozenRows(1);
    sheet1.getRange(1, 1, 1, mergedHeaders.length).setFontWeight('bold').setBackground('#E6F1FB');
  }

  const sheet2 = tempSs.insertSheet('전체 진행 현황');
  const summaryRows = [['작업자', '범위', '완료', '할당', '진행률(%)']];
  if (assignments.length === 0) {
    summaryRows.push(['(할당된 작업자가 없습니다)', '', '', '', '']);
  } else {
    const byWorker = {};
    assignments.forEach(a => {
      if (!byWorker[a.worker]) byWorker[a.worker] = [];
      byWorker[a.worker].push(a);
    });
    Object.keys(byWorker).forEach(w => {
      const ranges = byWorker[w];
      const indices = new Set();
      ranges.forEach(r => {
        const s = Math.max(1, r.start);
        const e = Math.min(dataRows.length, r.end);
        for (let i = s; i <= e; i++) indices.add(i);
      });
      const total = indices.size;
      let done = 0;
      indices.forEach(i => { if (results[String(i)] && results[String(i)].decision) done++; });
      const pct = total > 0 ? Math.round(done / total * 100) : 0;
      const rangeStr = ranges.map(r => `${r.start}-${r.end}`).join(', ');
      summaryRows.push([w, rangeStr, done, total, pct]);
    });
  }
  sheet2.getRange(1, 1, summaryRows.length, 5).setValues(summaryRows);
  sheet2.setFrozenRows(1);
  sheet2.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#E6F1FB');

  const sheet3 = tempSs.insertSheet('검수 결과 통계');
  let total = dataRows.length;
  let doneCount = 0, matchCount = 0, mismatchCount = 0;
  const reasonCounts = {};
  Object.values(results).forEach(r => {
    if (r.decision) {
      doneCount++;
      if (r.decision === '일치') matchCount++;
      else if (r.decision === '불일치') {
        mismatchCount++;
        if (r.reason) reasonCounts[r.reason] = (reasonCounts[r.reason] || 0) + 1;
      }
    }
  });
  const decisionTotal = matchCount + mismatchCount;
  const statsRows = [
    ['프로젝트', projName],
    ['모드', proj.mode === 'free' ? '자유 모드' : '표준 모드'],
    ['전체 데이터', total],
    ['검수 완료', doneCount],
    ['일치', matchCount],
    ['불일치', mismatchCount],
    ['진행률(%)', total > 0 ? Math.round(doneCount / total * 100) : 0],
    ['일치율(%)', decisionTotal > 0 ? Math.round(matchCount / decisionTotal * 100) : 0],
    ['', ''],
    ['[불일치 사유 분포 (건수 많은 순)]', ''],
    ['사유명', '건수']
  ];
  const sortedReasons = Object.entries(reasonCounts).sort((a, b) => b[1] - a[1]);
  if (sortedReasons.length === 0) statsRows.push(['(불일치 항목 없음)', '']);
  else sortedReasons.forEach(([n, c]) => {
    const pct = mismatchCount > 0 ? Math.round(c / mismatchCount * 100) : 0;
    statsRows.push([n, `${c} (${pct}%)`]);
  });
  sheet3.getRange(1, 1, statsRows.length, 2).setValues(statsRows);
  sheet3.getRange(1, 1, 8, 1).setFontWeight('bold');
  sheet3.getRange(10, 1).setFontWeight('bold').setBackground('#E6F1FB');
  sheet3.getRange(11, 1, 1, 2).setFontWeight('bold').setBackground('#F0F0F0');
  sheet3.setColumnWidth(1, 280);
  sheet3.setColumnWidth(2, 200);

  const sheet4 = tempSs.insertSheet('작업자별 성과');
  const ws = computeWorkerStats_(projectId);
  const wsRows = [['작업자', '처리 건수', '일치', '불일치', '일치율(%)', '중간 작업시간(초)', '시간당 처리량(건)', '시간 측정 표본수']];
  if (ws.length === 0) wsRows.push(['(데이터 없음)', '', '', '', '', '', '', '']);
  else ws.forEach(s => {
    wsRows.push([s.worker, s.total, s.match, s.mismatch, s.matchRate, s.avgSec, s.perHour, s.sampleCount || 0]);
  });
  sheet4.getRange(1, 1, wsRows.length, 8).setValues(wsRows);
  sheet4.setFrozenRows(1);
  sheet4.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#E6F1FB');

  const sheet5 = tempSs.insertSheet('오매핑 사유 목록');
  const reasonsValues = reasonsSheet.getDataRange().getValues();
  if (reasonsValues.length > 0) {
    sheet5.getRange(1, 1, reasonsValues.length, reasonsValues[0].length).setValues(reasonsValues);
    sheet5.setFrozenRows(1);
    sheet5.getRange(1, 1, 1, reasonsValues[0].length).setFontWeight('bold').setBackground('#E6F1FB');
  }

  SpreadsheetApp.flush();
  return tempSs.getId();
}

// ============ 미작업 행 삭제 ============
/**
 * 검수 결과(decision)가 없는 미작업 행을 Data 시트에서 제거하고,
 * Results / Logs / Assignments 의 키(행번호)를 새 순서로 재매핑한다.
 *
 * 핵심 알고리즘:
 *   1. 완료된 행의 oldKey 집합 수집
 *   2. 완료 행만 남긴 newDataRows 구성 + oldKey→newKey 매핑 테이블 생성
 *   3. Data 시트 재작성 (자유 모드면 그룹헤더 병합 복원)
 *   4. Results / Logs 의 '키' 컬럼을 매핑 테이블로 일괄 치환
 *   5. Assignments 의 start/end 를 해당 범위 내 살아남은 newKey 의 min/max 로 갱신
 *      (범위 내 완료 행이 없으면 그 할당은 제거)
 */
function deleteUnworkedRows(payload) {
  if (!verifyAdminPassword(payload.password)) return { ok: false, error: '관리자 인증 실패' };
  const projectId = String(payload.projectId || '');
  if (!projectId) return { ok: false, error: '프로젝트가 선택되지 않았습니다' };

  const proj = getProjectInfo_(projectId);
  if (!proj) return { ok: false, error: '프로젝트를 찾을 수 없습니다' };

  const projectSs   = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const dataSheet    = projectSs.getSheetByName('Data');
  const resultsSheet = projectSs.getSheetByName('Results');
  const logsSheet    = projectSs.getSheetByName('Logs');
  const assignSheet  = projectSs.getSheetByName('Assignments');

  // ── 1. Data 읽기 ──────────────────────────────────────────
  const dv = dataSheet.getDataRange().getValues();
  const isFree      = (proj.mode === 'free');
  const headerCount = isFree ? 2 : 1;
  if (dv.length <= headerCount) return { ok: true, deleted: 0 };

  const headerRows = dv.slice(0, headerCount);
  const dataRows   = dv.slice(headerCount);
  const totalCols  = (headerRows[headerCount - 1] || []).length || 1;

  // ── 2. 완료 키 수집 ──────────────────────────────────────
  const rv = resultsSheet.getDataRange().getValues();
  const rh = rv.length > 0 ? rv[0].map(String) : [];
  const cKey      = findHeader_(rh, ['키']);
  const cDecision = findHeader_(rh, ['검수결과']);

  const doneKeys = new Set();
  for (let i = 1; i < rv.length; i++) {
    const k = cKey >= 0 ? String(rv[i][cKey] || '') : '';
    const d = cDecision >= 0 ? String(rv[i][cDecision] || '') : '';
    if (k && d) doneKeys.add(k);
  }

  // ── 3. oldKey→newKey 매핑 + 살아남은 데이터 ─────────────
  const oldToNew   = {};   // { "3": "2", "5": "3", ... }
  const newDataRows = [];
  let newKeyCounter = 1;

  for (let i = 0; i < dataRows.length; i++) {
    const oldKey = String(i + 1);
    if (doneKeys.has(oldKey)) {
      oldToNew[oldKey] = String(newKeyCounter++);
      newDataRows.push(dataRows[i]);
    }
    // 미작업 행은 그냥 건너뜀 (삭제)
  }

  const deletedCount = dataRows.length - newDataRows.length;
  if (deletedCount === 0) return { ok: true, deleted: 0 };

  // ── 4. Data 시트 재작성 ──────────────────────────────────
  dataSheet.clear();
  dataSheet.clearFormats();

  if (isFree) {
    const schema   = proj.schema || {};
    const leftLen  = (schema.leftGroup  && schema.leftGroup.columns)  ? schema.leftGroup.columns.length  : 0;
    const rightLen = (schema.rightGroup && schema.rightGroup.columns) ? schema.rightGroup.columns.length : 0;

    // 1행: 그룹 헤더
    dataSheet.getRange(1, 1, 1, totalCols).setValues([headerRows[0]]);
    if (leftLen > 1)  dataSheet.getRange(1, 1,           1, leftLen).merge();
    if (rightLen > 1) dataSheet.getRange(1, leftLen + 1, 1, rightLen).merge();
    dataSheet.getRange(1, 1, 1, totalCols)
      .setHorizontalAlignment('center').setBackground('#E6F1FB').setFontWeight('bold');

    // 2행: 컬럼 헤더
    dataSheet.getRange(2, 1, 1, totalCols).setValues([headerRows[1]]);
    dataSheet.getRange(2, 1, 1, totalCols).setFontWeight('bold').setBackground('#F1EFE8');

    // 3행~: 데이터
    if (newDataRows.length > 0) {
      dataSheet.getRange(3, 1, newDataRows.length, totalCols).setValues(newDataRows);
    }
    dataSheet.setFrozenRows(2);
  } else {
    // 표준 모드
    dataSheet.getRange(1, 1, 1, totalCols).setValues([headerRows[0]]);
    dataSheet.setFrozenRows(1);
    if (newDataRows.length > 0) {
      dataSheet.getRange(2, 1, newDataRows.length, totalCols).setValues(newDataRows);
    }
  }

  // ── 5. Results 키 재매핑 ─────────────────────────────────
  if (rv.length > 1 && cKey >= 0) {
    const newRv = [rv[0]];
    for (let i = 1; i < rv.length; i++) {
      const oldKey = String(rv[i][cKey] || '');
      const newKey = oldToNew[oldKey];
      if (newKey !== undefined) {
        const row = rv[i].slice();
        row[cKey] = newKey;
        newRv.push(row);
      }
      // 매핑에 없는 결과(완료 아닌 잔여) → 제거
    }
    resultsSheet.clearContents();
    resultsSheet.getRange(1, 1, newRv.length, rh.length).setValues(newRv);
    resultsSheet.setFrozenRows(1);
  }

  // ── 6. Logs 키 재매핑 ────────────────────────────────────
  const lv = logsSheet.getDataRange().getValues();
  if (lv.length > 1) {
    const lh    = lv[0].map(String);
    const lcKey = findHeader_(lh, ['키']);
    if (lcKey >= 0) {
      const newLv = [lv[0]];
      for (let i = 1; i < lv.length; i++) {
        const oldKey = String(lv[i][lcKey] || '');
        const newKey = oldToNew[oldKey];
        if (newKey !== undefined) {
          const row = lv[i].slice();
          row[lcKey] = newKey;
          newLv.push(row);
        }
      }
      logsSheet.clearContents();
      logsSheet.getRange(1, 1, newLv.length, lh.length).setValues(newLv);
      logsSheet.setFrozenRows(1);
    }
  }

  // ── 7. Assignments 범위 재계산 ───────────────────────────
  const av = assignSheet.getDataRange().getValues();
  if (av.length > 1) {
    const newAv = [av[0]];
    for (let i = 1; i < av.length; i++) {
      const worker   = String(av[i][0] || '');
      const oldStart = parseInt(av[i][1]) || 1;
      const oldEnd   = parseInt(av[i][2]) || 1;
      if (!worker) continue;

      // 이 할당 범위 안에서 살아남은 행들의 newKey 를 수집
      const survivingNewKeys = [];
      for (let k = oldStart; k <= oldEnd; k++) {
        const nk = oldToNew[String(k)];
        if (nk !== undefined) survivingNewKeys.push(parseInt(nk));
      }

      if (survivingNewKeys.length > 0) {
        // min~max 범위로 재설정 (새 번호 기준)
        newAv.push([worker, Math.min(...survivingNewKeys), Math.max(...survivingNewKeys)]);
      }
      // 범위 내 완료 행이 하나도 없으면 이 할당 제거
    }
    assignSheet.clearContents();
    assignSheet.getRange(1, 1, newAv.length, 3).setValues(newAv);
    assignSheet.setFrozenRows(1);
  }

  SpreadsheetApp.flush();
  return { ok: true, deleted: deletedCount, remaining: newDataRows.length };
}

// ============ 진단/유틸 함수 ============
function repairLogsSheet(projectId) {
  if (!projectId) {
    const projects = getProjects_();
    if (projects.length === 0) return '프로젝트 없음';
    projectId = projects[projects.length - 1].id;
  }
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  const resultsSheet = projectSs.getSheetByName('Results');
  const logsSheet = projectSs.getSheetByName('Logs');

  const rv = resultsSheet.getDataRange().getValues();
  const keyToWorker = {};
  if (rv.length > 0) {
    const rh = rv[0].map(String);
    const cK = findHeader_(rh, ['키']);
    const cW = findHeader_(rh, ['작업자']);
    if (cK >= 0 && cW >= 0) {
      for (let i = 1; i < rv.length; i++) {
        const k = String(rv[i][cK] || '');
        const w = String(rv[i][cW] || '');
        if (k && w) keyToWorker[k] = w;
      }
    }
  }

  const lv = logsSheet.getDataRange().getValues();
  if (lv.length <= 1) return '로그 없음';
  const lh = lv[0].map(String);
  const lcK = findHeader_(lh, ['키']);
  const lcW = findHeader_(lh, ['작업자']);
  if (lcK < 0 || lcW < 0) return 'Logs 헤더 잘못됨';

  let fixed = 0, dropped = 0;
  const newRows = [lv[0]];
  for (let i = 1; i < lv.length; i++) {
    const row = lv[i].slice();
    const k = String(row[lcK] || '');
    if (k && keyToWorker[k]) {
      if (row[lcW] !== keyToWorker[k]) { row[lcW] = keyToWorker[k]; fixed++; }
      newRows.push(row);
    } else dropped++;
  }
  logsSheet.clear();
  logsSheet.getRange(1, 1, newRows.length, newRows[0].length).setValues(newRows);
  logsSheet.setFrozenRows(1);
  return `교정 ${fixed}, 제거 ${dropped}, 남은 ${newRows.length - 1}`;
}

function debugProjectSheets(projectId) {
  if (!projectId) {
    const projects = getProjects_();
    if (projects.length === 0) return '프로젝트 없음';
    projectId = projects[projects.length - 1].id;
  }
  const proj = getProjectInfo_(projectId);
  const projectSs = SpreadsheetApp.openById(getProjectSpreadsheetId_(projectId));
  Logger.log('=== ' + proj.name + ' (' + proj.mode + ') ===');
  Logger.log('스키마: ' + JSON.stringify(proj.schema));
  ['Data', 'Results', 'Logs'].forEach(name => {
    const s = projectSs.getSheetByName(name);
    const v = s.getDataRange().getValues();
    Logger.log('--- ' + name + ' (' + v.length + '행) ---');
    if (v[0]) Logger.log('헤더: ' + JSON.stringify(v[0]));
    if (v[1]) Logger.log('샘플: ' + JSON.stringify(v[1]));
  });
  return 'debugProjectSheets 실행 완료';
}
