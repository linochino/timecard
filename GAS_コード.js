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

// ★事業所の設定（勤務形態一覧表の見出しに使う。必要なら書き換えてください）
const CONFIG = {
  事業所名:       'リーノちの',
  支援の種類:     '児童発達支援、放課後等デイサービス',
  常勤週標準時間: 40   // 常勤職員が勤務すべき時間数（時間／週）
}

// シフト種類 → 勤務時間区分・勤務時間・サービス提供時間 の対応表
// （申請書フォーマット用。初期値。必要に応じて調整してください）
//   区分 ①9:00-18:00 / ②9:00-13:00 / ③14:00-18:00 / ④休日
//   サービス a 9:20-12:00 / b 14:00-17:00 / c 休日
const SHIFT_MAP = {
  '日勤':     { code: '①', hours: 8, service: 'ab' },
  '午前':     { code: '②', hours: 4, service: 'a'  },
  '午後':     { code: '③', hours: 4, service: 'b'  },
  '変則':     { code: '①', hours: 8, service: 'ab' },
  '春夏':     { code: '①', hours: 8, service: 'ab' },
  '休':       { code: '④', hours: 0, service: 'c'  },
  '有給':     { code: '④', hours: 0, service: 'c'  },
  'AM有休':   { code: '④', hours: 0, service: 'c'  },
  'PM有休':   { code: '④', hours: 0, service: 'c'  },
  '時間有給': { code: '④', hours: 0, service: 'c'  },
  '研修':     { code: '①', hours: 8, service: 'ab' }
}

// 祝日（HTML側と同じ。シフト未設定日を「休」にする判定に使う）
const HOLIDAYS = {
  '2025-01-01':1,'2025-01-13':1,'2025-02-11':1,'2025-02-23':1,'2025-02-24':1,'2025-03-20':1,
  '2025-04-29':1,'2025-05-03':1,'2025-05-04':1,'2025-05-05':1,'2025-07-21':1,'2025-08-11':1,
  '2025-09-15':1,'2025-09-23':1,'2025-10-13':1,'2025-11-03':1,'2025-11-23':1,'2025-11-24':1,
  '2026-01-01':1,'2026-01-12':1,'2026-02-11':1,'2026-02-23':1,'2026-03-20':1,'2026-04-29':1,
  '2026-05-03':1,'2026-05-04':1,'2026-05-05':1,'2026-05-06':1,'2026-07-20':1,'2026-08-11':1,
  '2026-09-21':1,'2026-09-22':1,'2026-09-23':1,'2026-10-12':1,'2026-11-03':1,'2026-11-23':1
}

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
  employees:  ['ID', '名前', 'チームID', '有効', '順序', '休日曜日(JSON)', 'デフォルトシフト', '職種', '勤務形態', '暗証番号', '雇用区分'],
  attendance: ['日付', 'スタッフ名', '雇用区分', '出勤', '退勤', '実働', '残業', '休憩イン', '休憩アウト', 'スタッフID'],
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
    // 列： 日付(A) スタッフ名(B) 雇用区分(C) 出勤(D) 退勤(E) 実働(F) 残業(G) 休憩イン(H) 休憩アウト(I) スタッフID(J)
    if (key === 'attendance') {
      sh.getRange('A2:A').setNumberFormat('yyyy/m/d')   // 日付 → 2026/6/18
      sh.getRange('D2:E').setNumberFormat('H:mm')         // 出勤・退勤 → 9:00
      sh.getRange('F2:G').setNumberFormat('0.00')         // 実働・残業 → 8.00（時間・小数）
      sh.getRange('H2:I').setNumberFormat('H:mm')         // 休憩イン・アウト → 12:00
      sh.hideColumns(8, 3)  // H〜J列（休憩イン・休憩アウト・スタッフID）を非表示
    }
  }
  return sh
}

// シートのヘッダーに不足している列があれば追記する（既存シートの移行用）
// 例：既存「スタッフ」シートに「職種」「勤務形態」列を後から足す
function ensureColumns(key) {
  const sh = getSheet(key)
  const want = HEADERS[key]
  const lastCol = sh.getLastColumn()
  const have = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0] : []
  let added = false
  want.forEach((h, i) => {
    if (have[i] !== h) {
      sh.getRange(1, i + 1).setValue(h).setFontWeight('bold').setBackground('#fce8ef')
      added = true
    }
  })
  return added
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
      case 'generateSchedule': result = generateScheduleSheet(data.month); break
      case 'verifyPin':      result = handleVerifyPin(data); break
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
  ensureColumns('employees')  // 職種・勤務形態・暗証番号 列を自動で用意（なければ追加）

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
      default_shift:  r['デフォルトシフト'] || '日勤',
      job_type:        r['職種'] || '',
      employment_type: r['勤務形態'] || ''
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
// 列： 日付(1) スタッフ名(2) 雇用区分(3) 出勤(4) 退勤(5) 実働(6) 残業(7) 休憩イン(8) 休憩アウト(9) スタッフID(10)
const ATT_COL = { clock_in: 4, clock_out: 5, break_start: 8, break_end: 9 }
const ATT_EMPID_COL = 10
const FIXED_BREAK = { start: 13, end: 14 }  // 休憩を押さない場合の固定昼休み（13:00〜14:00）
const STD_WORK_MIN = 8 * 60                 // 正社の所定労働（8時間）。これを超えた分を残業に

// スタッフの雇用区分（正社/パート）。明示があればそれ、なければ勤務形態A/B→正社・C/D→パート
function empKbn(emp) {
  if (!emp) return ''
  const explicit = String(emp['雇用区分'] == null ? '' : emp['雇用区分']).trim()
  if (explicit) return explicit
  const k = String(emp['勤務形態'] == null ? '' : emp['勤務形態']).trim().toUpperCase()
  if (k === 'A' || k === 'B') return '正社'
  if (k === 'C' || k === 'D') return 'パート'
  return ''
}

// 2つの時刻区間の重なり（分）。固定昼休みの控除に使う
function overlapMinutes(ci, co, startHour, endHour) {
  const bs = new Date(ci); bs.setHours(startHour, 0, 0, 0)
  const be = new Date(ci); be.setHours(endHour, 0, 0, 0)
  const s = Math.max(ci.getTime(), bs.getTime())
  const e = Math.min(co.getTime(), be.getTime())
  return e > s ? (e - s) / 60000 : 0
}

function handleRecord(data) {
  const now = new Date()  // 実際の打刻時刻
  const targetCol = ATT_COL[data.type]
  if (!targetCol) return { error: '不明な打刻種類: ' + data.type }

  const emp = sheetToObjects('employees').find(e => String(e['ID']) === String(data.employee_id))
  const empName = emp ? emp['名前'] : ''
  const kbn = empKbn(emp)
  const dateValue = new Date(data.date + 'T00:00:00+09:00')  // 「2026/6/18」表示用

  const sh = getSheet('attendance')

  // 既存の「その日・その人」の行を探す
  const values = sh.getDataRange().getValues()  // ヘッダー含む
  let targetRow = -1
  for (let i = 1; i < values.length; i++) {
    const rowDate  = normDate(values[i][0])                        // A列:日付
    const rowEmpId = String(values[i][ATT_EMPID_COL - 1] || '')     // J列:スタッフID
    if (rowDate === data.date && rowEmpId === String(data.employee_id)) {
      targetRow = i + 1
      break
    }
  }

  if (targetRow === -1) {
    // その日初めての打刻 → 新しい行を作る（10列）
    const rowData = [dateValue, empName, kbn, '', '', '', '', '', '', data.employee_id]
    rowData[targetCol - 1] = now
    sh.appendRow(rowData)
    targetRow = sh.getLastRow()
    sh.getRange(targetRow, 1).setNumberFormat('yyyy/m/d')   // 日付
    sh.getRange(targetRow, 4, 1, 2).setNumberFormat('H:mm')  // 出勤・退勤
    sh.getRange(targetRow, 6, 1, 2).setNumberFormat('0.00')  // 実働・残業
    sh.getRange(targetRow, 8, 1, 2).setNumberFormat('H:mm')  // 休憩イン・アウト
  } else {
    sh.getRange(targetRow, targetCol).setValue(now).setNumberFormat('H:mm')
    if (kbn && !values[targetRow - 1][2]) sh.getRange(targetRow, 3).setValue(kbn)  // 雇用区分が空なら補う
  }

  // 実働・残業を再計算
  recalcWork(sh, targetRow, kbn)

  // 日付 → スタッフ名 の順に並べ替え
  sortAttendance(sh)

  return { success: true, timestamp: toIso(now) }
}

// その行の実働・残業を計算して書き込む（出勤・退勤がそろっていれば）
function recalcWork(sh, row, kbn) {
  const v  = sh.getRange(row, 1, 1, ATT_EMPID_COL).getValues()[0]
  const ci = v[3], co = v[4], bs = v[7], be = v[8]  // 出勤・退勤・休憩イン・休憩アウト
  if (!(ci instanceof Date) || !(co instanceof Date)) {
    sh.getRange(row, 6, 1, 2).clearContent()  // まだ計算できない
    return
  }
  let mins = (co.getTime() - ci.getTime()) / 60000
  // 休憩：打刻があればその実時間、なければ固定昼休み(13-14)の重なりを控除
  let breakMins = (bs instanceof Date && be instanceof Date)
    ? (be.getTime() - bs.getTime()) / 60000
    : overlapMinutes(ci, co, FIXED_BREAK.start, FIXED_BREAK.end)
  mins = Math.max(0, mins - breakMins)
  sh.getRange(row, 6).setValue(Math.round(mins / 60 * 100) / 100)  // 実働(時間)
  // 残業：正社のみ。所定(8h)を超えた分。パートは空欄
  if (kbn === '正社') {
    sh.getRange(row, 7).setValue(Math.round(Math.max(0, mins - STD_WORK_MIN) / 60 * 100) / 100)
  } else {
    sh.getRange(row, 7).clearContent()
  }
}

// 打刻記録を 日付→スタッフ名 の順に並べ替え
function sortAttendance(sh) {
  const lastRow = sh.getLastRow()
  if (lastRow <= 2) return  // ヘッダー+1行以下なら並べ替え不要
  sh.getRange(2, 1, lastRow - 1, ATT_EMPID_COL).sort([
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
  ensureColumns('employees')
  const id    = uuid()
  const count = sheetToObjects('employees').length + 1
  appendRow('employees', [
    id, data.name, data.team_id, true, count,
    JSON.stringify(data.fixed_off_days || [0, 6]),
    data.default_shift || '日勤',
    data.job_type || '',         // 職種（後でシートに入力）
    data.employment_type || '',  // 勤務形態 A〜D（後でシートに入力）
    data.pin || '',              // 暗証番号（後でシートに入力）
    data.employment_kbn || ''    // 雇用区分 正社/パート（空なら勤務形態から自動判定）
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

// ★とりあえずの共通PIN（個別の暗証番号が未設定のスタッフはこれで開く。あとで各自設定推奨）
const DEFAULT_PIN = '0000'

// 暗証番号(PIN)の照合（勤務記録・シフト閲覧の本人確認）
// PINは画面に送らず、サーバー側でだけ照合する
function handleVerifyPin(data) {
  const emp = sheetToObjects('employees').find(e => String(e['ID']) === String(data.empId))
  if (!emp) return { ok: false, error: 'スタッフが見つかりません' }
  // 個別PINがあればそれ、なければ共通の DEFAULT_PIN を使う（未設定でも保護が効く）
  const pin = String(emp['暗証番号'] == null ? '' : emp['暗証番号']).trim() || DEFAULT_PIN
  return { ok: String(data.pin == null ? '' : data.pin).trim() === pin }
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
// 勤務形態一覧表（出勤表）の生成
// =====================================================
// 指定申請用の「従業者の勤務の体制及び勤務形態一覧表」を
// スプレッドシートに自動生成する。シフト予定ベース。
function generateScheduleSheet(month) {
  // month 例: "2026-07"。未指定なら今月
  if (!month) month = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM')
  const [y, m] = month.split('-').map(Number)
  const N = new Date(y, m, 0).getDate()   // その月の日数
  const reiwa = y - 2018                    // 令和の年

  ensureColumns('employees')

  // 有効なスタッフ（順序順）
  const emps = sheetToObjects('employees')
    .filter(r => r['有効'] === true || r['有効'] === 'TRUE' || r['有効'] === 'true')
    .map(r => ({
      id:           String(r['ID']),
      name:         r['名前'] || '',
      job:          r['職種'] || '',
      empType:      r['勤務形態'] || '',
      fixedOff:     parseJsonSafe(r['休日曜日(JSON)'], [0, 6]),
      defaultShift: r['デフォルトシフト'] || '日勤',
      sort:         Number(r['順序']) || 0
    }))
    .sort((a, b) => a.sort - b.sort)

  // 当月のシフト変更（上書き）を取得
  const from = `${month}-01`, to = `${month}-${String(N).padStart(2, '0')}`
  const shiftMap = {}
  for (const s of sheetToObjects('shifts')) {
    const d = normDate(s['日付'])
    if (d >= from && d <= to) shiftMap[String(s['スタッフID']) + '_' + d] = s['シフト種類']
  }

  // 日付リスト
  const DOW = ['日', '月', '火', '水', '木', '金', '土']
  const days = []
  for (let d = 1; d <= N; d++) {
    const dow = new Date(y, m - 1, d).getDay()
    const ds  = `${month}-${String(d).padStart(2, '0')}`
    days.push({ d, dow, ds, holiday: !!HOLIDAYS[ds] })
  }

  // 列レイアウト： 職種(1) 勤務形態(2) 氏名(3) | 日付... | 合計 週平均 常勤換算
  const FIRST_DAY_COL = 4
  const TOTAL_COL = FIRST_DAY_COL + N
  const AVG_COL   = TOTAL_COL + 1
  const FTE_COL   = AVG_COL + 1
  const LAST_COL  = FTE_COL
  const HEADER_ROWS = 4

  // シフト判定（上書き優先、なければ既定。週末/固定休/祝日は休）
  function resolveShift(emp, day) {
    let st = shiftMap[emp.id + '_' + day.ds]
    if (st === undefined || st === null || st === '') {
      const isWknd   = day.dow === 0 || day.dow === 6
      const fixedOff = Array.isArray(emp.fixedOff) && emp.fixedOff.indexOf(day.dow) >= 0
      st = (isWknd || fixedOff || day.holiday) ? '休' : emp.defaultShift
    }
    return SHIFT_MAP[st] || SHIFT_MAP['休']
  }

  // ---- 値の2次元配列を組み立て ----
  const rows = []
  const blankRow = () => { const a = []; for (let i = 0; i < LAST_COL; i++) a.push(''); return a }

  // 1行目：タイトル
  const titleRow = blankRow()
  titleRow[0] = `従業者の勤務の体制及び勤務形態一覧表　（R${reiwa}年${m}月分）`
  rows.push(titleRow)
  // 2行目：サブ情報
  const subRow = blankRow()
  subRow[0] = `支援の種類：${CONFIG.支援の種類}　　事業所名：${CONFIG.事業所名}　　常勤職員が勤務すべき時間数：${CONFIG.常勤週標準時間}時間／週`
  rows.push(subRow)
  // 3・4行目：ヘッダー（日番号 / 曜日）
  const hdr1 = blankRow(), hdr2 = blankRow()
  hdr1[0] = '職種'; hdr1[1] = '勤務形態'; hdr1[2] = '氏名'
  for (let i = 0; i < N; i++) {
    hdr1[FIRST_DAY_COL - 1 + i] = days[i].d
    hdr2[FIRST_DAY_COL - 1 + i] = DOW[days[i].dow]
  }
  hdr1[TOTAL_COL - 1] = '合計時間'; hdr1[AVG_COL - 1] = '週平均'; hdr1[FTE_COL - 1] = '常勤換算'
  rows.push(hdr1); rows.push(hdr2)

  // 各スタッフ＝2行（上段=勤務時間区分①〜④ / 下段=サービス提供 a/b/c）
  const empRowStart = []
  for (const emp of emps) {
    const up = blankRow(), low = blankRow()
    up[0] = emp.job; up[1] = emp.empType; up[2] = emp.name
    let total = 0
    for (let i = 0; i < N; i++) {
      const info = resolveShift(emp, days[i])
      up[FIRST_DAY_COL - 1 + i]  = info.code
      low[FIRST_DAY_COL - 1 + i] = info.service
      total += info.hours
    }
    const avg = total / N * 7  // 暦月対応：月合計 ÷ 暦日数 × 7
    const fte = CONFIG.常勤週標準時間 ? avg / CONFIG.常勤週標準時間 : 0
    up[TOTAL_COL - 1] = total
    up[AVG_COL - 1]   = Math.round(avg * 10) / 10
    up[FTE_COL - 1]   = Math.round(fte * 100) / 100
    empRowStart.push(rows.length + 1)  // 上段の行番号（1始まり）
    rows.push(up); rows.push(low)
  }

  // 空行＋凡例
  rows.push(blankRow())
  const legend = [
    '【勤務時間区分】　①9:00〜18:00　②9:00〜13:00　③14:00〜18:00　④休日',
    '【サービス提供時間】　a 9:20〜12:00　　b 14:00〜17:00　　c 休日',
    '【勤務形態区分】　A：常勤で専従　B：常勤で兼務　C：常勤以外で専従　D：常勤以外で兼務'
  ]
  const legendStartRow = rows.length + 1
  for (const t of legend) { const r = blankRow(); r[0] = t; rows.push(r) }

  // ---- シート初期化 ----
  const sheetName = '勤務形態一覧表'
  let sh = SS.getSheetByName(sheetName)
  if (sh) sh.clear(); else sh = SS.insertSheet(sheetName)
  // 固定(フリーズ)を解除しておく（全列結合とフリーズが競合してエラーになるのを防ぐ）
  sh.setFrozenRows(0); sh.setFrozenColumns(0)
  // 必要な行数・列数を確保
  if (sh.getMaxColumns() < LAST_COL)   sh.insertColumnsAfter(sh.getMaxColumns(), LAST_COL - sh.getMaxColumns())
  if (sh.getMaxRows()    < rows.length) sh.insertRowsAfter(sh.getMaxRows(), rows.length - sh.getMaxRows())
  // 既存の結合を解除
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart()

  // ---- 一括書き込み ----
  sh.getRange(1, 1, rows.length, LAST_COL).setValues(rows)

  // ---- 結合 ----
  sh.getRange(1, 1, 1, LAST_COL).merge()  // タイトル
  sh.getRange(2, 1, 1, LAST_COL).merge()  // サブ情報
  sh.getRange(3, 1, 2, 1).merge()         // 職種
  sh.getRange(3, 2, 2, 1).merge()         // 勤務形態
  sh.getRange(3, 3, 2, 1).merge()         // 氏名
  sh.getRange(3, TOTAL_COL, 2, 1).merge()
  sh.getRange(3, AVG_COL, 2, 1).merge()
  sh.getRange(3, FTE_COL, 2, 1).merge()
  for (const rStart of empRowStart) {
    sh.getRange(rStart, 1, 2, 1).merge()
    sh.getRange(rStart, 2, 2, 1).merge()
    sh.getRange(rStart, 3, 2, 1).merge()
    sh.getRange(rStart, TOTAL_COL, 2, 1).merge()
    sh.getRange(rStart, AVG_COL, 2, 1).merge()
    sh.getRange(rStart, FTE_COL, 2, 1).merge()
  }
  for (let i = 0; i < legend.length; i++) sh.getRange(legendStartRow + i, 1, 1, LAST_COL).merge()

  // ---- 体裁 ----
  sh.getRange(1, 1).setFontSize(13).setFontWeight('bold').setHorizontalAlignment('center')
  sh.getRange(2, 1).setFontSize(10).setHorizontalAlignment('center')
  sh.getRange(3, 1, 2, LAST_COL)
    .setFontWeight('bold').setHorizontalAlignment('center')
    .setVerticalAlignment('middle').setBackground('#f0f0f0')

  const gridRows = HEADER_ROWS + emps.length * 2
  sh.getRange(3, 1, gridRows - 2, LAST_COL)
    .setBorder(true, true, true, true, true, true)
    .setFontSize(9).setVerticalAlignment('middle')
  if (emps.length > 0) {
    sh.getRange(HEADER_ROWS + 1, FIRST_DAY_COL, emps.length * 2, N).setHorizontalAlignment('center')
    sh.getRange(HEADER_ROWS + 1, 1, emps.length * 2, 3).setHorizontalAlignment('left')
    sh.getRange(HEADER_ROWS + 1, TOTAL_COL, emps.length * 2, 3).setHorizontalAlignment('center')
  }

  // 週末・祝日の色をヘッダー日付列に
  for (let i = 0; i < N; i++) {
    const col = FIRST_DAY_COL + i
    if (days[i].dow === 0 || days[i].holiday) sh.getRange(3, col, 2, 1).setBackground('#fde2e2')
    else if (days[i].dow === 6)               sh.getRange(3, col, 2, 1).setBackground('#e2ecfd')
  }

  // 列幅
  sh.setColumnWidth(1, 160)  // 職種
  sh.setColumnWidth(2, 64)   // 勤務形態
  sh.setColumnWidth(3, 92)   // 氏名
  for (let i = 0; i < N; i++) sh.setColumnWidth(FIRST_DAY_COL + i, 26)
  sh.setColumnWidth(TOTAL_COL, 64)
  sh.setColumnWidth(AVG_COL, 56)
  sh.setColumnWidth(FTE_COL, 60)

  // 凡例の体裁
  sh.getRange(legendStartRow, 1, legend.length, 1).setFontSize(9)

  // 見出し行を固定（列の固定はタイトルの全列結合と競合するため行わない）
  sh.setFrozenRows(HEADER_ROWS)

  return { success: true, sheet: sheetName, month, staff: emps.length }
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
  // 空・null は fallback（JSON.parse(null) は null を返してしまうため明示的に弾く）
  if (str === null || str === undefined || str === '') return fallback
  try {
    const v = JSON.parse(str)
    return (v === null || v === undefined) ? fallback : v
  } catch (e) { return fallback }
}
