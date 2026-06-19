// =====================================================
// タイムカード用 Google Apps Script コード
// =====================================================
//
// 【セットアップ手順】
// 1. Googleドライブで「新しいスプレッドシート」を作成する
// 2. メニュー「拡張機能」→「Apps Script」を開く
// 3. 表示されたコードを全部消して、このファイルの内容を貼り付ける
// 4. 「保存」（Ctrl+S）する
// 5. 「デプロイ」→「新しいデプロイ」をクリック
// 6. 種類：「ウェブアプリ」を選択
// 7. 次のアクセスを許可するユーザー：「全員」を選択
// 8. 「デプロイ」→「アクセスを承認」→自分のGoogleアカウントで承認
// 9. 表示されたウェブアプリのURLをコピーして
//    タイムカード_GAS版.html の GAS_URL に貼り付ける
// =====================================================

const SS = SpreadsheetApp.getActiveSpreadsheet()

// ★ここに好きなパスワードを設定してください（HTMLと同じにすること）
const SECRET_TOKEN = 'linochino'

// シート名の定義
const SHEET = {
  teams:      'チーム',
  employees:  'スタッフ',
  attendance: '打刻記録',
  requests:   '申請',
  shifts:     'シフト'
}

// 各シートのヘッダー行
const HEADERS = {
  teams:      ['ID', '名前', '順序'],
  employees:  ['ID', '名前', 'チームID', '有効', '順序', '休日曜日(JSON)', 'デフォルトシフト'],
  attendance: ['日付', 'スタッフ名', '出勤', '休憩イン', '休憩アウト', '退勤', 'スタッフID'],
  requests:   ['ID', 'スタッフID', '日付', '申請種類', '理由', '時間数', 'ステータス', '申請日時'],
  shifts:     ['ID', 'スタッフID', '日付', 'シフト種類']
}

// シートを取得（なければ作成してヘッダーを追加）
function getSheet(key) {
  let sh = SS.getSheetByName(SHEET[key])
  if (!sh) {
    sh = SS.insertSheet(SHEET[key])
    sh.appendRow(HEADERS[key])
    sh.getRange(1, 1, 1, HEADERS[key].length).setFontWeight('bold').setBackground('#fce8ef')
    // 打刻記録シートは見やすく設定
    if (key === 'attendance') {
      sh.getRange('A2:A').setNumberFormat('yyyy/m/d')   // 日付 → 2026/6/18
      sh.getRange('C2:F').setNumberFormat('H:mm')         // 出勤〜退勤 → 9:00
      sh.hideColumns(7)  // G列(スタッフID)を非表示にする
    }
  }
  return sh
}

// シートのデータをオブジェクト配列に変換
function sheetToObjects(key) {
  const sh = getSheet(key)
  const data = sh.getDataRange().getValues()
  if (data.length <= 1) return []
  const headers = data[0]
  return data.slice(1).map(row => {
    const obj = {}
    headers.forEach((h, i) => { obj[h] = row[i] === '' ? null : row[i] })
    return obj
  })
}

// 行を追加
function appendRow(key, values) {
  getSheet(key).appendRow(values)
}

// UUID生成
function uuid() {
  return Utilities.getUuid()
}

// 現在の日本時間を返す
function nowJP() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'")
}

// 今日の日付（日本時間）
function todayJP() {
  return Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd')
}

// セルの値を ISO日付(yyyy-MM-dd) に正規化（Date型・文字列どちらでも対応）
function normDate(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd')
  const s = String(v || '')
  // "2026/6/18" → "2026-06-18"
  if (s.indexOf('/') >= 0) {
    const p = s.split('/')
    return `${p[0]}-${String(p[1]).padStart(2,'0')}-${String(p[2]).padStart(2,'0')}`
  }
  return s.slice(0, 10)
}

// セルの値を ISO日時(+09:00付き) に正規化（Date型・文字列どちらでも対応）
function toIso(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Tokyo', "yyyy-MM-dd'T'HH:mm:ss'+09:00'")
  return String(v || '')
}

// 打刻記録（横並び：1人1日1行）を個別レコードの配列に展開する
// 列： 日付 / スタッフ名 / 出勤 / 休憩イン / 休憩アウト / 退勤 / スタッフID
// 例： 1行 → 出勤・休憩イン・休憩アウト・退勤 の最大4レコードに分解
function getAttendanceRecords() {
  const COL_TO_TYPE = {
    '出勤': 'clock_in', '休憩イン': 'break_start',
    '休憩アウト': 'break_end', '退勤': 'clock_out'
  }
  const out = []
  for (const r of sheetToObjects('attendance')) {
    const date  = normDate(r['日付'])
    const empId = String(r['スタッフID'] || '')
    const name  = r['スタッフ名'] || ''
    for (const col in COL_TO_TYPE) {
      const v = r[col]
      if (v) {
        out.push({
          employee_id:   empId,
          employee_name: name,
          type:          COL_TO_TYPE[col],
          date:          date,
          timestamp:     toIso(v)
        })
      }
    }
  }
  return out
}

// =====================================================
// GETリクエスト処理
// =====================================================
function doGet(e) {
  if ((e.parameter && e.parameter.token) !== SECRET_TOKEN) {
    return toJson({ error: '認証エラー' })
  }
  const action = (e.parameter && e.parameter.action) || 'init'
  let result
  try {
    switch (action) {
      case 'init':       result = handleInit(e.parameter); break
      case 'attendance': result = handleGetAttendance(e.parameter); break
      case 'requests':   result = handleGetRequests(e.parameter); break
      case 'shifts':     result = handleGetShifts(e.parameter); break
      case 'attendance-records': result = handleGetAttendanceRecords(e.parameter); break
      default:           result = { error: '不明なアクション: ' + action }
    }
  } catch (err) {
    result = { error: err.message }
  }
  return toJson(result)
}

// =====================================================
// POSTリクエスト処理
// =====================================================
function doPost(e) {
  let data
  try {
    data = JSON.parse(e.postData.contents)
  } catch (err) {
    return toJson({ error: 'JSONパースエラー' })
  }
  if (data.token !== SECRET_TOKEN) {
    return toJson({ error: '認証エラー' })
  }
  let result
  try {
    switch (data.action) {
      case 'record':         result = handleRecord(data); break
      case 'request':        result = handleRequest(data); break
      case 'addTeam':        result = handleAddTeam(data); break
      case 'addEmployee':    result = handleAddEmployee(data); break
      case 'deleteTeam':     result = handleDeleteTeam(data); break
      case 'deleteEmployee': result = handleDeleteEmployee(data); break
      default:               result = { error: '不明なアクション: ' + data.action }
    }
  } catch (err) {
    result = { error: err.message }
  }
  return toJson(result)
}

function toJson(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON)
}

// =====================================================
// データ取得処理
// =====================================================

// 初期データ（チーム・スタッフ・今日の打刻状況）
function handleInit(params) {
  const today = params.today || todayJP()

  const teams = sheetToObjects('teams').map(r => ({
    id:         String(r['ID']),
    name:       r['名前'],
    sort_order: Number(r['順序']) || 0
  }))

  const employees = sheetToObjects('employees')
    .filter(r => r['有効'] === true || r['有効'] === 'TRUE' || r['有効'] === 'true')
    .map(r => ({
      id:             String(r['ID']),
      name:           r['名前'],
      team_id:        String(r['チームID']),
      active:         true,
      sort_order:     Number(r['順序']) || 0,
      fixed_off_days: parseJsonSafe(r['休日曜日(JSON)'], [0, 6]),
      default_shift:  r['デフォルトシフト'] || '日勤'
    }))

  // 今日の打刻（スタッフIDごとに最新の種類を取得）
  const attRaw = getAttendanceRecords().filter(r => r.date === today)
  const attMap = {}
  attRaw.forEach(r => {
    const empId = r.employee_id
    const cur = attMap[empId]
    if (!cur || r.timestamp > cur.timestamp) {
      attMap[empId] = { employee_id: empId, type: r.type, timestamp: r.timestamp }
    }
  })

  return { teams, employees, attendance: Object.values(attMap) }
}

// 打刻記録取得（月次集計用）
function handleGetAttendance(params) {
  const { empId, from, to } = params
  const data = getAttendanceRecords()
    .filter(r => r.employee_id === empId && r.date >= from && r.date <= to)
    .map(r => ({
      employee_id: r.employee_id,
      type:        r.type,
      date:        r.date,
      timestamp:   r.timestamp
    }))
  return { data }
}

// 申請データ取得（月次集計用）
function handleGetRequests(params) {
  const { empId, from, to, status } = params
  const types = ['有給', 'AM有休', 'PM有休', '時間有給', '研修', '中抜け', '自車使用', '打刻補正']
  const data = sheetToObjects('requests')
    .filter(r => {
      if (String(r['スタッフID']) !== empId) return false
      if (r['日付'] < from || r['日付'] > to) return false
      // status指定がない場合は、approvedとpendingの両方を返す
      if (status && r['ステータス'] !== status) return false
      if (!types.includes(r['申請種類'])) return false
      return true
    })
    .map(r => ({
      employee_id:     String(r['スタッフID']),
      date:            r['日付'],
      requested_type:  r['申請種類'],
      reason:          r['理由'],
      requested_hours: Number(r['時間数']) || 0,
      status:          r['ステータス']
    }))
  return { data }
}

// シフトデータ取得
function handleGetShifts(params) {
  const { from, to } = params
  const data = sheetToObjects('shifts')
    .filter(r => r['日付'] >= from && r['日付'] <= to)
    .map(r => ({
      employee_id: String(r['スタッフID']),
      date:        r['日付'],
      shift_type:  r['シフト種類']
    }))
  return { data }
}

// 打刻記録取得（管理画面用・スタッフ名付き）
function handleGetAttendanceRecords(params) {
  const { from, to } = params
  // スタッフ名が空の旧データ用に、IDから名前を引けるようにしておく
  const empMap = {}
  for (const e of sheetToObjects('employees')) {
    empMap[String(e['ID'])] = e['名前']
  }

  const data = getAttendanceRecords()
    .filter(r => r.date >= from && r.date <= to)
    .map(r => ({
      date:          r.date,
      employee_id:   r.employee_id,
      employee_name: r.employee_name || empMap[r.employee_id] || '(削除済み)',
      type:          r.type,
      timestamp:     r.timestamp
    }))
    .sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date)
      return a.employee_name.localeCompare(b.employee_name)
    })

  return { data }
}

// =====================================================
// データ書き込み処理
// =====================================================

// 打刻記録（1人1日1行・横並び）
// 列： 日付(1) / スタッフ名(2) / 出勤(3) / 休憩イン(4) / 休憩アウト(5) / 退勤(6) / スタッフID(7)
const ATT_COL = { clock_in: 3, break_start: 4, break_end: 5, clock_out: 6 }

function handleRecord(data) {
  const now = new Date()  // 実際の打刻時刻
  const targetCol = ATT_COL[data.type]
  if (!targetCol) return { error: '不明な打刻種類: ' + data.type }

  // スタッフ名を取得
  const emp = sheetToObjects('employees').find(e => String(e['ID']) === String(data.employee_id))
  const empName = emp ? emp['名前'] : ''
  const dateValue = new Date(data.date + 'T00:00:00+09:00')  // 「2026/6/18」表示用の日付値

  const sh = getSheet('attendance')

  // 既存の「その日・その人」の行を探す
  const values = sh.getDataRange().getValues()  // ヘッダー含む
  let targetRow = -1
  for (let i = 1; i < values.length; i++) {
    const rowDate  = normDate(values[i][0])              // A列:日付
    const rowEmpId = String(values[i][6] || '')           // G列:スタッフID
    if (rowDate === data.date && rowEmpId === String(data.employee_id)) {
      targetRow = i + 1  // 1始まりの行番号
      break
    }
  }

  if (targetRow === -1) {
    // その日初めての打刻 → 新しい行を作る
    const rowData = [dateValue, empName, '', '', '', '', data.employee_id]
    rowData[targetCol - 1] = now
    sh.appendRow(rowData)
    targetRow = sh.getLastRow()
    sh.getRange(targetRow, 1).setNumberFormat('yyyy/m/d')   // 日付
    sh.getRange(targetRow, 3, 1, 4).setNumberFormat('H:mm')  // 出勤〜退勤
  } else {
    // 既存の行の該当列だけ更新
    sh.getRange(targetRow, targetCol).setValue(now).setNumberFormat('H:mm')
  }

  // 日付 → スタッフ名 の順に並べ替え
  sortAttendance(sh)

  return { success: true, timestamp: toIso(now) }
}

// 打刻記録を 日付→スタッフ名 の順に並べ替え
function sortAttendance(sh) {
  const lastRow = sh.getLastRow()
  if (lastRow <= 2) return  // ヘッダー+1行以下なら並べ替え不要
  sh.getRange(2, 1, lastRow - 1, 7).sort([
    { column: 1, ascending: true },  // 日付
    { column: 2, ascending: true }   // スタッフ名
  ])
}

// 各種申請
function handleRequest(data) {
  const id  = uuid()
  const now = nowJP()
  appendRow('requests', [
    id, data.employee_id, data.date,
    data.requested_type, data.reason || '', data.requested_hours || 0,
    data.status || 'pending', now
  ])
  return { success: true, id }
}

// チーム追加
function handleAddTeam(data) {
  const id    = uuid()
  const count = sheetToObjects('teams').length + 1
  appendRow('teams', [id, data.name, count])
  return { success: true, id }
}

// スタッフ追加
function handleAddEmployee(data) {
  const id    = uuid()
  const count = sheetToObjects('employees').length + 1
  appendRow('employees', [
    id, data.name, data.team_id, true, count,
    JSON.stringify(data.fixed_off_days || [0, 6]),
    data.default_shift || '日勤'
  ])
  return { success: true, id }
}

// チーム削除
function handleDeleteTeam(data) {
  deleteRowById('teams', 'ID', data.id)
  // 所属スタッフを無効化
  const sh   = getSheet('employees')
  const vals = sh.getDataRange().getValues()
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][2]) === String(data.id)) {
      sh.getRange(i + 1, 4).setValue(false)
    }
  }
  return { success: true }
}

// スタッフ削除（論理削除）
function handleDeleteEmployee(data) {
  const sh   = getSheet('employees')
  const vals = sh.getDataRange().getValues()
  for (let i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(data.id)) {
      sh.getRange(i + 1, 4).setValue(false)
      break
    }
  }
  return { success: true }
}

// =====================================================
// ユーティリティ
// =====================================================

function deleteRowById(sheetKey, colName, id) {
  const sh      = getSheet(sheetKey)
  const vals    = sh.getDataRange().getValues()
  const headers = vals[0]
  const colIdx  = headers.indexOf(colName)
  if (colIdx < 0) return
  for (let i = vals.length - 1; i >= 1; i--) {
    if (String(vals[i][colIdx]) === String(id)) {
      sh.deleteRow(i + 1)
      break
    }
  }
}

function parseJsonSafe(str, fallback) {
  try { return JSON.parse(str) } catch (e) { return fallback }
}
