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

// 打刻種類の変換マップ
const TYPE_JP = { clock_in:'出勤', break_start:'休憩イン', break_end:'休憩アウト', clock_out:'退勤' }
const TYPE_EN = { '出勤':'clock_in', '休憩イン':'break_start', '休憩アウト':'break_end', '退勤':'clock_out' }

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
  attendance: ['ID', 'スタッフID', 'スタッフ名', '打刻', '日付', '時刻'],
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
    // 打刻記録シートは「日付」「時刻」列の表示形式を見やすく設定
    if (key === 'attendance') {
      sh.getRange('E2:E').setNumberFormat('yyyy/m/d')  // 日付 → 2026/6/18
      sh.getRange('F2:F').setNumberFormat('H:mm')        // 時刻 → 19:00
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

// 打刻記録の1行を共通フォーマットに変換（新旧の列名どちらにも対応）
function mapAttendanceRow(r) {
  const rawType = r['打刻'] || r['打刻種類']   // 新:打刻 / 旧:打刻種類
  const rawTs   = r['時刻'] || r['日時']        // 新:時刻 / 旧:日時
  return {
    employee_id:   String(r['スタッフID']),
    employee_name: r['スタッフ名'] || '',
    type:          TYPE_EN[rawType] || rawType,
    date:          normDate(r['日付']),
    timestamp:     toIso(rawTs)
  }
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
  const attRaw = sheetToObjects('attendance')
    .map(mapAttendanceRow)
    .filter(r => r.date === today)
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
  const data = sheetToObjects('attendance')
    .map(mapAttendanceRow)
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

  const data = sheetToObjects('attendance')
    .map(mapAttendanceRow)
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

// 打刻記録
function handleRecord(data) {
  const id  = uuid()
  const now = new Date()  // 実際の日時（時刻列に入る）
  // スタッフ名を取得
  const emp = sheetToObjects('employees').find(e => String(e['ID']) === String(data.employee_id))
  const empName = emp ? emp['名前'] : ''
  const typeJP = TYPE_JP[data.type] || data.type
  // 日付は本物のDate値にして「2026/6/18」表示、時刻も本物のDate値にして「19:00」表示
  const dateValue = new Date(data.date + 'T00:00:00+09:00')

  const sh = getSheet('attendance')
  sh.appendRow([id, data.employee_id, empName, typeJP, dateValue, now])
  // 念のため、この行の表示形式を設定（既存シートでも見やすく）
  const row = sh.getLastRow()
  sh.getRange(row, 5).setNumberFormat('yyyy/m/d')  // 日付
  sh.getRange(row, 6).setNumberFormat('H:mm')        // 時刻

  return { success: true, id, timestamp: toIso(now) }
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
