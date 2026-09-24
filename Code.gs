/**
 * AP Exam Only Registration — 核准通知寄信
 *
 * 規則：
 *   - 只寄給 M 欄 Approval = "Approved" 的學生
 *   - To:  學生信箱（B 欄 Email）
 *   - Cc:  家長信箱（標題為 PARENT_EMAIL_HEADER 的欄位）
 *   - Bcc: 主管（SUPERVISOR_EMAIL）
 *   - 信件內文的課程清單來自 D 欄、Join Code 清單來自 E 欄（以逗號分隔，一科對一個代碼）
 *   - 寄出後會在「Email Status」欄寫上寄送時間，避免重複寄信
 *
 * 使用方式：
 *   1. 在試算表中：擴充功能 → Apps Script，貼上此程式並儲存
 *   2. 修改下方 CONFIG（主管信箱、家長信箱欄位標題、工作表名稱）
 *   3. 重新整理試算表，上方會出現「📧 AP Exam Only」選單
 *   4. 先按「預覽寄送名單」確認，再按「寄送測試信給自己」看信件樣式，最後按「寄出核准通知」
 */

const CONFIG = {
  SHEET_NAME: '',                        // 工作表名稱；留空 = 使用目前開啟的工作表
  HEADER_ROW: 1,                         // 標題列
  SUPERVISOR_EMAIL: 'supervisor@kcislk.ntpc.edu.tw', // ← 請改成主管信箱（BCC）
  PARENT_EMAIL_HEADER: 'Parent Email',   // ← 家長信箱欄位的標題（可用逗號分隔多個信箱）
  STATUS_HEADER: 'Email Status',         // 寄送紀錄欄；若不存在會自動新增在最右邊

  // 欄位標題（依截圖）
  EMAIL_HEADER: 'Email',
  COURSE_HEADER: 'AP Course(s) for Exam Only',
  JOIN_CODE_HEADER: 'Join Code',
  APPROVAL_HEADER: 'Approval',
  APPROVED_VALUE: 'approved',            // 不分大小寫

  EXAM_YEAR: '2027',                     // SY26-27 的 AP 考試為 2027 年 5 月
  SENDER_NAME: 'Ms. Sia',                // 寄件者顯示名稱
  SUBJECT: '【IPD 中學國際】AP Exam Only Application Confirmation | AP 自學報考申請確認',
  MY_AP_URL: 'https://myap.collegeboard.org/login',
};

/* ---------------- 選單 ---------------- */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📧 AP Exam Only')
    .addItem('1. 預覽寄送名單', 'previewRecipients')
    .addItem('2. 寄送測試信給自己（第一位核准學生）', 'sendTestEmailToMe')
    .addSeparator()
    .addItem('3. 寄出核准通知（Approved 且尚未寄過）', 'sendApprovalEmails')
    .addToUi();
}

/* ---------------- 主要功能 ---------------- */

/** 列出將會寄出的學生，不寄信。 */
function previewRecipients() {
  const { pending, problems } = collectRows_();
  const lines = pending.map(r =>
    `第 ${r.row} 列｜${r.studentEmail}｜cc: ${r.parentEmails.join(', ') || '（無家長信箱）'}\n` +
    r.courses.map((c, i) => `    ${c} → ${r.codes[i]}`).join('\n'));
  let msg = `將寄出 ${pending.length} 封信：\n\n` + (lines.join('\n\n') || '（無）');
  if (problems.length) msg += `\n\n⚠️ 以下列有問題，不會寄出：\n` + problems.join('\n');
  SpreadsheetApp.getUi().alert('預覽寄送名單', msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

/** 用第一位待寄學生的資料寄一封信到自己信箱（不 cc 家長、不 bcc 主管、不寫入狀態）。 */
function sendTestEmailToMe() {
  const { pending } = collectRows_();
  if (!pending.length) {
    SpreadsheetApp.getUi().alert('目前沒有待寄的核准學生。');
    return;
  }
  const me = Session.getActiveUser().getEmail();
  const r = pending[0];
  GmailApp.sendEmail(me, '[TEST] ' + CONFIG.SUBJECT, buildPlainBody_(r.courses, r.codes), {
    htmlBody: buildHtmlBody_(r.courses, r.codes),
    name: CONFIG.SENDER_NAME,
  });
  SpreadsheetApp.getUi().alert(`測試信已寄到 ${me}（使用第 ${r.row} 列 ${r.studentEmail} 的資料）。`);
}

/** 寄出所有 Approved 且尚未寄過的信。 */
function sendApprovalEmails() {
  const ui = SpreadsheetApp.getUi();
  const { sheet, statusCol, pending, problems } = collectRows_();

  if (!pending.length) {
    ui.alert('沒有需要寄出的信（Approved 的學生都已寄過，或沒有 Approved）。' +
      (problems.length ? '\n\n⚠️ 問題列：\n' + problems.join('\n') : ''));
    return;
  }

  const confirm = ui.alert('確認寄信',
    `即將寄出 ${pending.length} 封核准通知信。\n` +
    `To: 學生｜Cc: 家長｜Bcc: ${CONFIG.SUPERVISOR_EMAIL}\n\n確定要寄出嗎？`,
    ui.ButtonSet.YES_NO);
  if (confirm !== ui.Button.YES) return;

  if (MailApp.getRemainingDailyQuota() < pending.length) {
    ui.alert(`今日寄信額度不足（剩 ${MailApp.getRemainingDailyQuota()} 封），請明天再寄。`);
    return;
  }

  let sent = 0;
  const errors = [];
  pending.forEach(r => {
    try {
      const options = {
        htmlBody: buildHtmlBody_(r.courses, r.codes),
        name: CONFIG.SENDER_NAME,
        bcc: CONFIG.SUPERVISOR_EMAIL,
      };
      if (r.parentEmails.length) options.cc = r.parentEmails.join(',');
      GmailApp.sendEmail(r.studentEmail, CONFIG.SUBJECT, buildPlainBody_(r.courses, r.codes), options);
      sheet.getRange(r.row, statusCol).setValue('Sent ' +
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm'));
      sent++;
    } catch (e) {
      errors.push(`第 ${r.row} 列 ${r.studentEmail}：${e.message}`);
      sheet.getRange(r.row, statusCol).setValue('Error: ' + e.message);
    }
  });

  let msg = `已寄出 ${sent} 封信。`;
  if (errors.length) msg += '\n\n❌ 寄送失敗：\n' + errors.join('\n');
  if (problems.length) msg += '\n\n⚠️ 未寄出（資料有問題）：\n' + problems.join('\n');
  ui.alert(msg);
}

/* ---------------- 讀取資料 ---------------- */

function collectRows_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = CONFIG.SHEET_NAME ? ss.getSheetByName(CONFIG.SHEET_NAME) : ss.getActiveSheet();
  if (!sheet) throw new Error('找不到工作表：' + CONFIG.SHEET_NAME);

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(CONFIG.HEADER_ROW, 1, 1, lastCol).getValues()[0]
    .map(h => String(h).trim());
  const col = name => headers.findIndex(h => h.toLowerCase() === name.toLowerCase());

  const idx = {
    email: col(CONFIG.EMAIL_HEADER),
    course: col(CONFIG.COURSE_HEADER),
    code: col(CONFIG.JOIN_CODE_HEADER),
    approval: col(CONFIG.APPROVAL_HEADER),
    parent: col(CONFIG.PARENT_EMAIL_HEADER),
    status: col(CONFIG.STATUS_HEADER),
  };
  const missing = ['email', 'course', 'code', 'approval'].filter(k => idx[k] < 0);
  if (missing.length) throw new Error('找不到必要欄位標題：' + missing.join(', '));

  // 沒有 Email Status 欄就自動新增
  let statusCol = idx.status + 1;
  if (idx.status < 0) {
    statusCol = lastCol + 1;
    sheet.getRange(CONFIG.HEADER_ROW, statusCol).setValue(CONFIG.STATUS_HEADER);
  }

  const numRows = sheet.getLastRow() - CONFIG.HEADER_ROW;
  const pending = [];
  const problems = [];
  if (numRows <= 0) return { sheet, statusCol, pending, problems };

  const data = sheet.getRange(CONFIG.HEADER_ROW + 1, 1, numRows, Math.max(lastCol, statusCol)).getValues();
  data.forEach((rowValues, i) => {
    const row = CONFIG.HEADER_ROW + 1 + i;
    const approval = String(rowValues[idx.approval]).trim().toLowerCase();
    if (approval !== CONFIG.APPROVED_VALUE) return;

    const status = String(rowValues[statusCol - 1] || '').trim();
    if (status.startsWith('Sent')) return; // 已寄過

    const studentEmail = String(rowValues[idx.email]).trim();
    const courses = splitList_(rowValues[idx.course]);
    const codes = splitList_(rowValues[idx.code]);
    const parentEmails = idx.parent >= 0 ? splitList_(rowValues[idx.parent]) : [];

    if (!studentEmail) return problems.push(`第 ${row} 列：沒有學生信箱`);
    if (!courses.length) return problems.push(`第 ${row} 列：沒有 AP 課程`);
    if (courses.length !== codes.length) {
      return problems.push(`第 ${row} 列 ${studentEmail}：課程數 (${courses.length}) 與 Join Code 數 (${codes.length}) 不一致`);
    }
    pending.push({ row, studentEmail, parentEmails, courses, codes });
  });

  if (idx.parent < 0) problems.push(`⚠️ 找不到「${CONFIG.PARENT_EMAIL_HEADER}」欄位，信件將不會 cc 家長`);
  return { sheet, statusCol, pending, problems };
}

/** 以逗號（半形/全形）、頓號、分號、換行分隔。 */
function splitList_(value) {
  return String(value || '')
    .split(/[,，、;；\n]+/)
    .map(s => s.trim())
    .filter(Boolean);
}

/* ---------------- 信件內容 ---------------- */

function buildHtmlBody_(courses, codes) {
  const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ul = items => '<ul>' + items.map(x => `<li>${esc(x)}</li>`).join('') + '</ul>';
  const codeUl = items => '<ul>' + items.map(x => `<li><b>${esc(x)}</b></li>`).join('') + '</ul>';
  const myAp = `<a href="${CONFIG.MY_AP_URL}">My AP</a>`;
  const y = CONFIG.EXAM_YEAR;

  return `
<div style="font-family: Arial, 'Microsoft JhengHei', sans-serif; font-size: 14px; line-height: 1.6; color: #222;">
<p>Dear Student,</p>
<p>This is to inform you that your AP Exam Only application for the following course(s) has been approved:</p>
${ul(courses)}
<p>Please be sure to complete the enrollment process on ${myAp} for the AP Course(s) listed above using the following Join Code(s):</p>
${codeUl(codes)}
<p><b>⚠️ Important Notices</b></p>
<ol>
  <li>You are required to join each AP Exam Only course on ${myAp} within five days from today to secure your eligibility to take the AP Exam(s) in May ${y}.</li>
  <li>According to the school’s AP registration policy, AP self-study students may take the AP Exam Only for any AP course, either on or off campus. However, students who earn a score of 3 or higher may not enroll in the same AP course in the following school year.</li>
  <li>Exam fee(s) will apply for each AP Course you enroll on ${myAp}, and you will be billed later this year.</li>
  <li>If you wish to cancel your AP Exam Only enrollment, please visit Ms. Sia in the 4F IPD Office as early as possible. Please note that a cancellation fee will be charged after the official ordering deadline.</li>
  <li>By completing your enrollment on ${myAp}, you acknowledge and agree to all of the above.</li>
</ol>
<p>For any further questions or inquiries, you may reply to this email or find Ms. Sia in the 4F IPD office.</p>
<p>Warm regards,<br>Ms. Sia</p>

<hr style="border: none; border-top: 1px solid #ccc; margin: 24px 0;">

<p>親愛的學生：</p>
<p>您申請的下列 AP 自學報考課程 (Exam Only) 已核准通過：</p>
${ul(courses)}
<p>請務必使用下方 Join Code 至 ${myAp} 完成上述 AP 課程之報名程序：</p>
${codeUl(codes)}
<p><b>⚠️ 重要注意事項</b></p>
<ol>
  <li>須於今日起五日內完成 ${myAp} 上各 AP 自學報考課程之加入程序，以確保於 ${y} 年 5 月 參加 AP 考試的資格。</li>
  <li>依據本校 AP 課程研修管理辦法，AP 自學學生可選擇於校內或校外參加任一 AP 課程之考試。然而，若學生於自學報考中獲得 3 分或以上成績，則不得於下一學年度修習相同的 AP 課程。</li>
  <li>每門於 ${myAp} 加入的 AP 課程皆需繳交相應考試費用，學校將於本學年稍後統一收費。</li>
  <li>若欲取消 AP 自學報考課程，請儘早至四樓國際處找 Ms. Sia 洽詢。請注意，逾期取消將酌收 AP 考試取消手續費。</li>
  <li>完成 ${myAp} 報名程序即視同已了解並同意以上所有事項。</li>
</ol>
<p>若有任何進一步疑問，請直接回覆此信件，或至四樓國際處找 Ms. Sia 諮詢。</p>
<p>祝 學安<br>Ms. Sia</p>
</div>`;
}

/** 純文字版本（給不支援 HTML 的信箱）。 */
function buildPlainBody_(courses, codes) {
  const list = items => items.map(x => '* ' + x).join('\n');
  const y = CONFIG.EXAM_YEAR;
  const url = CONFIG.MY_AP_URL;
  return `Dear Student,
This is to inform you that your AP Exam Only application for the following course(s) has been approved:

${list(courses)}

Please be sure to complete the enrollment process on My AP (${url}) for the AP Course(s) listed above using the following Join Code(s):

${list(codes)}

⚠️ Important Notices

1. You are required to join each AP Exam Only course on My AP within five days from today to secure your eligibility to take the AP Exam(s) in May ${y}.
2. According to the school’s AP registration policy, AP self-study students may take the AP Exam Only for any AP course, either on or off campus. However, students who earn a score of 3 or higher may not enroll in the same AP course in the following school year.
3. Exam fee(s) will apply for each AP Course you enroll on My AP, and you will be billed later this year.
4. If you wish to cancel your AP Exam Only enrollment, please visit Ms. Sia in the 4F IPD Office as early as possible. Please note that a cancellation fee will be charged after the official ordering deadline.
5. By completing your enrollment on My AP, you acknowledge and agree to all of the above.

For any further questions or inquiries, you may reply to this email or find Ms. Sia in the 4F IPD office.
Warm regards,
Ms. Sia

親愛的學生：
您申請的下列 AP 自學報考課程 (Exam Only) 已核准通過：

${list(courses)}

請務必使用下方 Join Code 至 My AP (${url}) 完成上述 AP 課程之報名程序：

${list(codes)}

⚠️ 重要注意事項

1. 須於今日起五日內完成 My AP 上各 AP 自學報考課程之加入程序，以確保於 ${y} 年 5 月 參加 AP 考試的資格。
2. 依據本校 AP 課程研修管理辦法，AP 自學學生可選擇於校內或校外參加任一 AP 課程之考試。然而，若學生於自學報考中獲得 3 分或以上成績，則不得於下一學年度修習相同的 AP 課程。
3. 每門於 My AP 加入的 AP 課程皆需繳交相應考試費用，學校將於本學年稍後統一收費。
4. 若欲取消 AP 自學報考課程，請儘早至四樓國際處找 Ms. Sia 洽詢。請注意，逾期取消將酌收 AP 考試取消手續費。
5. 完成 My AP 報名程序即視同已了解並同意以上所有事項。

若有任何進一步疑問，請直接回覆此信件，或至四樓國際處找 Ms. Sia 諮詢。
祝 學安
Ms. Sia`;
}
