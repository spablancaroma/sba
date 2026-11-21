const GAS_URL = 'https://script.google.com/macros/s/AKfycbzINJ6cRRGnjW4n1n1zW42LDRqziMAkOKwDZrwMhMWnNdmO74K815ld_gbcnB94eh8/exec';
const SITE_URL = 'https://spa-blanc-aroma.com/approve.html';
const ADMIN_PANEL_URL = 'https://spa-blanc-aroma.com/admin.html';
const ADMIN_EMAIL = 'spa.blanc.aroma@gmail.com';

function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const action = e && e.parameter && e.parameter.action;

  if (action) {
    let result = {};
    try {
      if (action === 'approve') {
        result = processApprovalApi(e.parameter.id);
      } else if (action === 'reject') {
        result = processRejectionApi(e.parameter.id);
      } else if (action === 'getAppInitData') {
        result = getAppInitData(e.parameter.therapist);
      } else if (action === 'getInitialDataForIndex') {
        result = getInitialDataForIndex();
      } else if (action === 'getLatestNews') {
        result = getLatestNews();
      } else if (action === 'getTherapistScheduleData') {
        result = getTherapistScheduleData(e.parameter.therapist);
      } else if (action === 'getTherapistList') {
        result = getTherapistList();
      } else if (action === 'getAllSchedules') {
        result = getAllSchedules(e.parameter.therapist);
      } else if (action === 'getNewsList') {
        result = getNewsList();
      } else if (action === 'getTherapistNewsList') {
        result = getTherapistNewsList();
      } else if (action === 'verifyPassword') {
        result = verifyPassword(e.parameter.pass);
      } else if (action === 'getReservationList') {
        result = getReservationList();
      } else if (action === 'getTodaySchedule') {
        result = getTodaySchedule();
      } else if (action === 'getDefaultShifts') {
        result = getDefaultShifts(e.parameter.therapist);
      } else if (action === 'loginTherapist') {
        result = loginTherapist(e.parameter.id, e.parameter.pass);
      } else if (action === 'getTherapistAccounts') {
        result = getTherapistAccounts(e.parameter.requesterId);
      } else if (action === 'changeOwnPassword') {
        result = changeOwnPassword(e.parameter.id, e.parameter.oldPass, e.parameter.newPass);
      } else if (action === 'getXPostDraft') {
        result = buildTodayShiftTweetData(e.parameter.comment || '');
      }
    } catch (err) {
      result = { success: false, message: err.message };
    }

    if (callback) {
      return ContentService.createTextOutput(`${callback}(${JSON.stringify(result)})`)
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return HtmlService.createHtmlOutput('API Active');
}

function doPost(e) {
  let result = {};
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;

    if (action === 'getAppInitData') {
      result = getAppInitData(postData.therapist);
    } else if (action === 'saveReservation') {
      result = saveReservation(postData.course, postData.datetime, postData.name, postData.email, postData.tel, postData.therapist, postData.outcall, postData.remarks);
    } else if (action === 'updateSchedule') {
      result = updateSchedule(postData.therapist, postData.date, postData.start, postData.end, postData.isHoliday, postData.notes);
    } else if (action === 'postNews') {
      result = postNews(postData.subject, postData.body);
    } else if (action === 'deleteNews') {
      result = deleteNews(postData.index);
    } else if (action === 'updateNews') {
      result = updateNews(postData.index, postData.subject, postData.body);
    } else if (action === 'postTherapistNews') {
      result = postTherapistNews(postData.subject, postData.body);
    } else if (action === 'deleteTherapistNews') {
      result = deleteTherapistNews(postData.index);
    } else if (action === 'updateTherapistNews') {
      result = updateTherapistNews(postData.index, postData.subject, postData.body);
    } else if (action === 'getTherapistNewsList') {
      result = getTherapistNewsList();
    } else if (action === 'getInitialDataForIndex') {
      result = getInitialDataForIndex();
    } else if (action === 'getLatestNews') {
      result = getLatestNews();
    } else if (action === 'getTherapistScheduleData') {
      result = getTherapistScheduleData(postData.therapist);
    } else if (action === 'getTherapistList') {
      result = getTherapistList();
    } else if (action === 'getAllSchedules') {
      result = getAllSchedules(postData.therapist);
    } else if (action === 'getNewsList') {
      result = getNewsList();
    } else if (action === 'verifyPassword') {
      result = verifyPassword(postData.pass);
    } else if (action === 'getReservationList') {
      result = getReservationList();
    } else if (action === 'sendTimeAdjustment') {
      result = sendTimeAdjustment(postData.id, postData.newDatetime);
    } else if (action === 'sendThankYouEmail') {
      result = sendThankYouEmail(postData.id);
    } else if (action === 'approve') {
      result = processApprovalApi(postData.id, postData.newDatetime);
    } else if (action === 'reject') {
      result = processRejectionApi(postData.id);
    } else if (action === 'getTodaySchedule') {
      result = getTodaySchedule();
    } else if (action === 'updateDefaultShifts') {
      result = updateDefaultShifts(postData.therapist, postData.shifts);
    } else if (action === 'loginTherapist') {
      result = loginTherapist(postData.id, postData.pass);
    } else if (action === 'saveTherapistAccount') {
      result = saveTherapistAccount(postData.name, postData.id, postData.pass, postData.role, postData.requesterId);
    } else if (action === 'getTherapistAccounts') {
      result = getTherapistAccounts(postData.requesterId);
    } else if (action === 'changeOwnPassword') {
      result = changeOwnPassword(postData.id, postData.oldPass, postData.newPass);
    } else if (action === 'getXPostDraft') {
      result = buildTodayShiftTweetData(postData.comment || '');
    } else if (action === 'postToX') {
      result = executePostToX(postData.mainText, postData.replyText, postData.requesterId);
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------
// アカウント認証・権限管理
// ----------------------------------------------------
function getAccountSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('セラピストアカウント');
  if (!sheet) {
    sheet = ss.insertSheet('セラピストアカウント');
    sheet.appendRow(['セラピスト名', 'ID', 'PW', '権限']);
    sheet.appendRow(['店舗アカウント', 'admin', 'blanc2025', 'admin']);
    sheet.appendRow(['そうた', 'souta', 'souta', 'therapist']);
    sheet.appendRow(['りょうすけ', 'ryosuke', 'ryosuke', 'therapist']);
  }
  return sheet;
}

function checkAdminPermission(requesterId) {
  if (!requesterId) return false;
  const sheet = getAccountSheet();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][1]).trim() === String(requesterId).trim()) {
      return (values[i][3] || 'therapist') === 'admin';
    }
  }
  return false;
}

function loginTherapist(id, pass) {
  try {
    const sheet = getAccountSheet();
    const values = sheet.getDataRange().getValues();
    const targetId = String(id || '').trim();
    const targetPass = String(pass || '').trim();

    for (let i = 1; i < values.length; i++) {
      const name = String(values[i][0]).trim();
      const accId = String(values[i][1]).trim();
      const accPw = String(values[i][2]).trim();
      const role = String(values[i][3] || 'therapist').trim();

      if (accId === targetId && accPw === targetPass) {
        return {
          success: true,
          therapistName: name,
          id: accId,
          role: role
        };
      }
    }
    return { success: false, message: 'IDまたはパスワードが正しくありません。' };
  } catch(e) {
    return { success: false, message: 'エラーが発生しました: ' + e.message };
  }
}

function getTherapistAccounts(requesterId) {
  try {
    if (requesterId && !checkAdminPermission(requesterId)) {
      return [];
    }
    const sheet = getAccountSheet();
    const values = sheet.getDataRange().getValues();
    const list = [];
    for (let i = 1; i < values.length; i++) {
      if (values[i][1]) {
        list.push({
          name: values[i][0],
          id: values[i][1],
          pass: values[i][2],
          role: values[i][3] || 'therapist'
        });
      }
    }
    return list;
  } catch(e) {
    return [];
  }
}

function saveTherapistAccount(name, id, pass, role, requesterId) {
  try {
    if (!checkAdminPermission(requesterId)) {
      return { success: false, message: 'アカウント登録・更新は管理者（admin）のみ実行可能です。' };
    }

    const sheet = getAccountSheet();
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][1]).trim() === String(id).trim()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1).setValue(name);
      sheet.getRange(targetRow, 3).setValue(pass);
      sheet.getRange(targetRow, 4).setValue(role || 'therapist');
    } else {
      sheet.appendRow([name, id, pass, role || 'therapist']);
    }

    if (name !== '店舗アカウント' && !SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name)) {
      const newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
      newSheet.appendRow(['日付', '開始時間', '終了時間', '備考']);
    }

    clearAllAppInitCache();
    return { success: true, message: 'アカウントを登録・更新しました。' };
  } catch(e) {
    return { success: false, message: '保存エラー: ' + e.message };
  }
}

function changeOwnPassword(id, oldPass, newPass) {
  try {
    const sheet = getAccountSheet();
    const values = sheet.getDataRange().getValues();
    const targetId = String(id || '').trim();
    const targetOldPass = String(oldPass || '').trim();
    const targetNewPass = String(newPass || '').trim();

    if (!targetNewPass) {
      return { success: false, message: '新しいパスワードを入力してください。' };
    }

    let targetRow = -1;
    for (let i = 1; i < values.length; i++) {
      const accId = String(values[i][1]).trim();
      const accPw = String(values[i][2]).trim();

      if (accId === targetId) {
        if (accPw !== targetOldPass) {
          return { success: false, message: '現在のパスワードが正しくありません。' };
        }
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 3).setValue(targetNewPass);
      return { success: true, message: 'パスワードを正常に変更しました！次回から新しいパスワードでログインしてください。' };
    } else {
      return { success: false, message: '該当するアカウントが見つかりません。' };
    }
  } catch(e) {
    return { success: false, message: '変更エラー: ' + e.message };
  }
}

// ----------------------------------------------------
// セラピスト向けお知らせ管理（「Th.お知らせ」シート）
// ----------------------------------------------------
function getTherapistNewsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Th.お知らせ');
  if (!sheet) {
    sheet = ss.insertSheet('Th.お知らせ');
    sheet.appendRow(['投稿日時', '件名', '本文']);
  }
  return sheet;
}

function getTherapistNewsList() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('therapist_news_list');
    if (cached) return JSON.parse(cached);

    const sheet = getTherapistNewsSheet();
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const news = [];
    for (let i = 1; i < values.length; i++) {
      news.push({
        index: i + 1,
        date: normalizeDateStr(values[i][0]).replace(/-/g, '/'),
        subject: values[i][1],
        body: values[i][2]
      });
    }
    const result = news.reverse();
    cache.put('therapist_news_list', JSON.stringify(result), 300);
    return result;
  } catch(e) {
    return [];
  }
}

function postTherapistNews(subject, body) {
  const sheet = getTherapistNewsSheet();
  sheet.appendRow([new Date(), subject, body]);
  try { CacheService.getScriptCache().remove('therapist_news_list'); } catch(e){}
  return { success: true };
}

function deleteTherapistNews(index) {
  getTherapistNewsSheet().deleteRow(index);
  try { CacheService.getScriptCache().remove('therapist_news_list'); } catch(e){}
  return { success: true };
}

function updateTherapistNews(index, subject, body) {
  const sheet = getTherapistNewsSheet();
  sheet.getRange(index, 2).setValue(subject);
  sheet.getRange(index, 3).setValue(body);
  try { CacheService.getScriptCache().remove('therapist_news_list'); } catch(e){}
  return { success: true };
}

// ----------------------------------------------------
// 既存コア機能
// ----------------------------------------------------
function getAppInitData(therapist) {
  const t = therapist || '指名なし';
  return {
    news: getLatestNews(),
    therapists: getTherapistList(),
    schedule: getTherapistScheduleData(t)
  };
}

function getTodaySchedule() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const therapists = getTherapistList(ss);
    const todayStr = normalizeDateStr(new Date());
    const todaySchedules = [];

    therapists.forEach(tName => {
      const sched = getScheduleForDate(tName, todayStr);
      let status = "未定";
      let start = "未定";
      let end = "";
      
      if (sched.exists && sched.start) {
        if (sched.start === "受付終了") {
          status = "受付終了"; start = "受付終了"; end = "";
        } else if (sched.start === "休み") {
          status = "休み"; start = "休み"; end = "";
        } else if (sched.start === "未定") {
          status = "未定"; start = "未定"; end = "";
        } else {
          status = "出勤"; start = sched.start; end = sched.end;
        }
      }
      
      todaySchedules.push({
        name: tName, status: status, start: start, end: end, notes: sched.notes || ""
      });
    });

    return todaySchedules;
  } catch(e) {
    return [];
  }
}

function normalizeDateStr(rawDate) {
  if (!rawDate) return "";
  if (rawDate instanceof Date) {
    const y = rawDate.getFullYear();
    const m = String(rawDate.getMonth() + 1).padStart(2, '0');
    const d = String(rawDate.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  let str = String(rawDate).trim().replace(/\//g, '-');
  if (str.includes('T')) str = str.split('T')[0];
  const parts = str.split('-');
  if (parts.length === 3 && parts[0].length === 4) {
    return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
  }
  return str;
}

function clearAllAppInitCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['news_list', 'therapist_news_list', 'therapist_list', 'reservation_list']);
  } catch(e){}
}

function formatDateTimeForEmail(dt) {
  try {
    const d = new Date(dt);
    if (isNaN(d.getTime())) return dt;
    const days = ["日", "月", "火", "水", "木", "金", "土"];
    const datePart = `${d.getMonth() + 1}月${d.getDate()}日`;
    const timePart = `${String(d.getHours()).padStart(2, '0')}:${String(d.getMinutes()).padStart(2, '0')}`;
    return `${datePart}(${days[d.getDay()]}) ${timePart}～`;
  } catch(e) {
    return dt;
  }
}

function getCourseDetails(course) {
  const isOutcall = String(course).startsWith('出張');
  const min = parseInt(String(course).replace(/[^0-9]/g, ''), 10);
  let text = min + "分";
  let price = "";
  if (min === 60) {
    text = isOutcall ? "出張60分" : "60分"; price = isOutcall ? "10,000円〜 + 交通費" : "9,000円";
  } else if (min === 90) {
    text = isOutcall ? "出張90分" : "90分"; price = isOutcall ? "13,000円〜 + 交通費" : "12,000円";
  } else if (min === 120) {
    text = isOutcall ? "出張120分" : "120分"; price = isOutcall ? "15,000円〜 + 交通費" : "14,000円";
  }
  return { text, price, isOutcall };
}

function checkBlacklist(email, tel) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('ブラックリスト');
    if (!sheet) return false;
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return false;

    const targetEmail = String(email).trim().toLowerCase();
    const targetTel = String(tel).trim().replace(/[-\s]/g, '');

    for (let i = 1; i < values.length; i++) {
      const blEmail = values[i][0] ? String(values[i][0]).trim().toLowerCase() : "";
      const blTel = values[i][1] ? String(values[i][1]).trim().replace(/[-\s]/g, '') : "";
      if ((blEmail && targetEmail && blEmail === targetEmail) || (blTel && targetTel && blTel === targetTel)) {
        return true;
      }
    }
    return false;
  } catch(e) {
    return false;
  }
}

function getTherapistEmail(therapistName) {
  if (!therapistName || therapistName === '指名なし' || therapistName === 'そうた') {
    return ADMIN_EMAIL;
  }
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(therapistName);
    if (!sheet) return ADMIN_EMAIL;
    const b1 = sheet.getRange("B1").getValue();
    if (b1 && String(b1).includes("@")) return String(b1).trim();
  } catch(e) {}
  return ADMIN_EMAIL;
}

function getEmailFooterHtml() {
  return `
    <div style="background-color: #222222; padding: 20px; text-align: center; font-size: 11px; color: #aaaaaa; font-family: 'Montserrat', sans-serif; letter-spacing: 0.05em; margin-top: 30px; border-radius: 0 0 8px 8px;">
      &copy; Spa Blanc Aroma All Rights Reserved.
    </div>
  `;
}

function getEmailFooterText() {
  return `\n\n------------------------------------\n© Spa Blanc Aroma All Rights Reserved.\nhttps://spa-blanc-aroma.com\n------------------------------------`;
}

function sendAdminNotification(id, dt, course, name, email, tel, therapist, outcall, remarks) {
  const approveUrl = SITE_URL + '?action=approve&id=' + id;
  const rejectUrl = SITE_URL + '?action=reject&id=' + id;
  const displayDt = formatDateTimeForEmail(dt);
  const courseInfo = getCourseDetails(course);
  const therapistText = therapist || "指名なし";
  const targetEmail = getTherapistEmail(therapist);

  let outcallMapHtml = "";
  if (outcall) {
    const mapSearchUrl = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(outcall)}`;
    outcallMapHtml = `
      <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; color: #e74c3c; white-space: nowrap;">出張先</th>
          <td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">
            ${outcall}<br>
            <a href="${mapSearchUrl}" target="_blank" style="font-size: 12px; color: #b39b7d; text-decoration: underline;">Googleマップで確認する</a>
          </td>
      </tr>
    `;
  }

  let remarksHtml = "";
  if (remarks) {
    remarksHtml = `<tr><th style="text-align: left; padding: 8px 15px 8px 0; white-space: nowrap;">備考</th><td style="padding: 8px 0;">${remarks}</td></tr>`;
  }

  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">New Reservation</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">【要確認】新規予約リクエスト</h2>
      </div>
      <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0; margin-bottom: 15px; font-size: 13px; color: #666;">
          ※詳細・承認作業は<a href="${ADMIN_PANEL_URL}" target="_blank" style="color: #b39b7d; font-weight: bold; text-decoration: underline;">管理画面</a>からも行えます。
        </p>
        <table style="width: 100%; border-collapse: collapse; background: #faf8f5; border-radius: 6px; padding: 10px 15px; margin-bottom: 20px;">
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; width: 110px; white-space: nowrap;">希望セラピスト</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold; color: #b39b7d;">${therapistText}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">日時</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">${displayDt}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">コース</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${courseInfo.text} (${courseInfo.price})</td></tr>
          ${outcallMapHtml}
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">お名前</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${name} 様</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">電話番号</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${tel}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: ${remarks ? '1px solid #f0ede8' : 'none'}; white-space: nowrap;">メール</th><td style="padding: 8px 0; border-bottom: ${remarks ? '1px solid #f0ede8' : 'none'};">${email}</td></tr>
          ${remarksHtml}
        </table>
        <div style="text-align: center; margin-top: 25px;">
          <a href="${approveUrl}" style="background-color: #b39b7d; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 30px; margin-right: 8px; display: inline-block;">承認する</a>
          <a href="${rejectUrl}" style="background-color: #888888; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 30px; display: inline-block;">却下する</a>
        </div>
      </div>
      ${getEmailFooterHtml()}
    </div>
  `;

  const textBody = `【要確認】新規予約リクエスト\n\n※管理画面: ${ADMIN_PANEL_URL}\n\n希望セラピスト: ${therapistText}\n内容: ${displayDt} ${courseInfo.text}(${courseInfo.price})\n${outcall ? '出張先: ' + outcall + '\n' : ''}お名前: ${name} 様 (電話: ${tel})\n\n承認: ${approveUrl}\n却下: ${rejectUrl}` + getEmailFooterText();
  
  const mailOptions = { htmlBody: htmlBody, name: "Spa Blanc Aroma" };
  if (targetEmail !== ADMIN_EMAIL) {
    mailOptions.bcc = ADMIN_EMAIL;
  }
  GmailApp.sendEmail(targetEmail, '【要確認】新規予約リクエスト', textBody, mailOptions);
}

function sendCustomerRequestReceived(email, name, dt, course, therapist, outcall) {
  const displayDt = formatDateTimeForEmail(dt);
  const courseInfo = getCourseDetails(course);
  const therapistText = therapist || "指名なし";

  let outcallHtml = "";
  if (outcall) {
    outcallHtml = `<tr><th style="text-align: left; padding: 8px 15px 8px 0; white-space: nowrap;">出張先</th><td style="padding: 8px 0;">${outcall}</td></tr>`;
  }

  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Received</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご予約リクエストを承りました</h2>
      </div>
      <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0;"><strong>${name} 様</strong></p>
        <p>ご予約リクエストを承りました。現在スケジュールを確認しております。確定まで今しばらくお待ちください。</p>
        
        <table style="width: 100%; border-collapse: collapse; background: #faf8f5; border-radius: 6px; padding: 10px 15px; margin: 20px 0;">
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; width: 110px; white-space: nowrap;">担当セラピスト</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold; color: #b39b7d;">${therapistText}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">日時</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">${displayDt}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'}; white-space: nowrap;">コース</th><td style="padding: 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'};">${courseInfo.text} (${courseInfo.price})</td></tr>
          ${outcallHtml}
        </table>
      </div>
      ${getEmailFooterHtml()}
    </div>
  `;
  const textBody = `${name} 様\n\nご予約リクエストを承りました。\n担当: ${therapistText}\n日時: ${displayDt}\nコース: ${courseInfo.text} (${courseInfo.price})\n\nSpa Blanc Aroma` + getEmailFooterText();
  GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご予約リクエストを承りました', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma"});
}

function sendCustomerConfirmation(email, name, dt, course, therapist, tel, outcall) {
  const displayDt = formatDateTimeForEmail(dt);
  const courseInfo = getCourseDetails(course);
  const therapistText = therapist || "指名なし";
  const targetStaffEmail = getTherapistEmail(therapist);

  let outcallHtml = "";
  if (outcall) {
    outcallHtml = `<tr><th style="text-align: left; padding: 8px 15px 8px 0; white-space: nowrap;">出張先</th><td style="padding: 8px 0;">${outcall}</td></tr>`;
  }

  const accessMapHtml = !outcall ? `
    <div style="margin-top: 20px; padding: 20px 15px; background: #faf8f5; border-radius: 6px; font-size: 14px; line-height: 1.8; color: #444; text-align: center;">
      <p style="margin: 0 0 8px 0; color: #555;">お時間ちょうどに下記個室住所までお越しください。</p>
      <p style="margin: 0 0 15px 0; font-weight: bold; color: #3b3530; line-height: 1.6;">
        〒533-0023<br>
        大阪市東淀川区東淡路4-24-35<br>
        JPアパートメント東淀川Ⅶ102号室
      </p>
      <div>
        <a href="https://maps.google.com/maps?q=${encodeURIComponent('大阪府大阪市東淀川区東淡路4-24-35 JPアパートメント東淀川Ⅶ 102号室')}" target="_blank" style="background-color: #b39b7d; color: #ffffff; padding: 10px 24px; text-decoration: none; border-radius: 25px; display: inline-block; font-size: 13px; font-weight: 500;">GoogleMapを開く</a>
      </div>
    </div>
  ` : '';

  const customerHtmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Confirmed</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご予約が確定いたしました</h2>
      </div>
      <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0;"><strong>${name} 様</strong></p>
        <p>ご予約が<strong style="color: #b39b7d;">【確定】</strong>いたしました。ご来店を心よりお待ち申し上げております。</p>
        
        <table style="width: 100%; border-collapse: collapse; background: #faf8f5; border-radius: 6px; padding: 10px 15px; margin: 20px 0;">
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; width: 110px; white-space: nowrap;">担当セラピスト</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold; color: #b39b7d;">${therapistText}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">日時</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">${displayDt}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'}; white-space: nowrap;">コース</th><td style="padding: 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'};">${courseInfo.text} (${courseInfo.price})</td></tr>
          ${outcallHtml}
        </table>

        ${accessMapHtml}
      </div>
      ${getEmailFooterHtml()}
    </div>
  `;
  
  const mapText = !outcall ? `\n\nお時間ちょうどに下記個室住所までお越しください。\n〒533-0023\n大阪市東淀川区東淡路4-24-35\nJPアパートメント東淀川Ⅶ102号室\n\n▼ GoogleMapを開く\nhttps://maps.google.com/maps?q=${encodeURIComponent('大阪府大阪市東淀川区東淡路4-24-35 JPアパートメント東淀川Ⅶ 102号室')}` : '';

  const customerTextBody = `${name}様\n\nご予約が【確定】いたしました。\n担当: ${therapistText}\n日時: ${displayDt}\nコース: ${courseInfo.text}${mapText}\n\nSpa Blanc Aroma` + getEmailFooterText();
  GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご予約が確定いたしました', customerTextBody, { htmlBody: customerHtmlBody, name: "Spa Blanc Aroma"});

  const staffHtmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Notification</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">【確定】予約が確定しました</h2>
      </div>
      <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0;">予約の確定処理を行いました。</p>
        <table style="width: 100%; border-collapse: collapse; background: #faf8f5; border-radius: 6px; padding: 10px 15px; margin: 20px 0;">
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; width: 110px; white-space: nowrap;">担当</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold; color: #b39b7d;">${therapistText}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">お名前</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${name} 様</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: 1px solid #f0ede8; white-space: nowrap;">日時</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">${displayDt}</td></tr>
          <tr><th style="text-align: left; padding: 8px 15px 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'}; white-space: nowrap;">コース</th><td style="padding: 8px 0; border-bottom: ${outcall ? '1px solid #f0ede8' : 'none'};">${courseInfo.text} (${courseInfo.price})</td></tr>
          ${outcallHtml}
        </table>
      </div>
      ${getEmailFooterHtml()}
    </div>
  `;
  const staffTextBody = `下記予約を確定しました。\n担当: ${therapistText}\nお名前: ${name}様\n日時: ${displayDt}\nコース: ${courseInfo.text}` + getEmailFooterText();
  
  const staffMailOptions = { htmlBody: staffHtmlBody, name: "Spa Blanc Aroma" };
  if (targetStaffEmail !== ADMIN_EMAIL) {
    staffMailOptions.bcc = ADMIN_EMAIL;
  }
  GmailApp.sendEmail(targetStaffEmail, `【確定】予約確定のお知らせ (${name}様)`, staffTextBody, staffMailOptions);
}

function sendCustomerRejection(email, name) {
  const now = new Date().getTime();
  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Notification</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご予約リクエストに関するお知らせ</h2>
      </div>
      <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0;"><strong>${name} 様</strong></p>
        <p>せっかくリクエスト頂いたのに誠に申し訳ございませんが、ご希望の日時ではご予約を承ることができませんでした。</p>
        <p style="margin-bottom: 0;">またの機会のご利用を心よりお待ちしております。</p>
        <span style="opacity: 0; font-size: 0px; display: none;">${now}</span>
      </div>
      ${getEmailFooterHtml()}
    </div>
  `;
  const textBody = `${name}様\n\nせっかくリクエスト頂いたのに誠に申し訳ございませんが、ご希望の日時ではご予約を承ることができませんでした。\n\nまたの機会のご利用を心よりお待ちしております。\n\nSpa Blanc Aroma` + getEmailFooterText();
  GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご予約リクエストに関するお知らせ', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma" });
}

function sendThankYouEmail(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return { success: false, message: "予約データがありません。" };

    let resData = null;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        resData = values[i];
        break;
      }
    }
    if (!resData) return { success: false, message: "該当する予約が見つかりません。" };

    const email = resData[4];
    const name = resData[5];
    const surveyUrl = "https://forms.gle/93xGi7wGHS5NxVCv6";
    const lineUrl = "https://lin.ee/N9lQ6Ec";

    const htmlBody = `
      <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
        <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
          <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Thank You</span>
          <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご来店ありがとうございました</h2>
        </div>
        <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
          <p style="margin-top: 0;"><strong>${name} 様</strong></p>
          <p style="margin-bottom: 25px;">本日は数あるサロンの中から Spa Blanc Aroma にお越しいただき、誠にありがとうございました。<br>お忙しい日常を忘れ、ゆったりとしたお時間をお過ごしいただけておりましたら幸いです。</p>

          <!-- アンケート枠（独立カード） -->
          <div style="background: #faf6f0; border: 1px solid #e8dfd5; border-radius: 10px; padding: 20px 15px; margin-bottom: 20px; text-align: center;">
            <p style="margin: 0 0 8px 0; font-weight: bold; color: #3b3530; font-size: 14px;">■ お客様の声（アンケート）ご協力のお願い</p>
            <p style="margin: 0 0 15px 0; color: #555; font-size: 13px; line-height: 1.6;">本日の施術のご感想やご意見をぜひお聞かせいただけますと幸いです。</p>
            <div>
              <a href="${surveyUrl}" style="background-color: #ad9579; color: #ffffff; padding: 10px 24px; text-decoration: none; border-radius: 25px; display: inline-block; font-size: 13px; font-weight: 500;">感想を送る（アンケートに回答）</a>
            </div>
          </div>

          <!-- 公式LINE枠（独立カード） -->
          <div style="background: #f4f9f6; border: 1px solid #d2e7dd; border-radius: 10px; padding: 20px 15px; margin-bottom: 10px; text-align: center;">
            <p style="margin: 0 0 8px 0; font-weight: bold; color: #3b3530; font-size: 14px;">■ 公式LINEのご案内</p>
            <p style="margin: 0 0 15px 0; color: #555; font-size: 13px; line-height: 1.6;">次回のご予約やお問合せは、公式LINEからも便利に行っていただけます。</p>
            <div>
              <a href="${lineUrl}" style="background-color: #06C755; color: #ffffff; padding: 10px 24px; text-decoration: none; border-radius: 25px; display: inline-block; font-size: 13px; font-weight: 500;">公式LINEで友だち追加する</a>
            </div>
          </div>

        </div>
        ${getEmailFooterHtml()}
      </div>
    `;

    const textBody = `${name}様\n\n本日は数あるサロンの中から Spa Blanc Aroma にお越しいただき、誠にありがとうございました。\nお忙しい日常を忘れ、ゆったりとしたお時間をお過ごしいただけておりましたら幸いです。\n\n■ お客様の声（アンケート）ご協力のお願い\n本日の施術のご感想やご意見をぜひお聞かせいただけますと幸いです。\n▼ 感想を送る（アンケートに回答）\n${surveyUrl}\n\n■ 公式LINEのご案内\n次回のご予約やお問合せは、公式LINEからも便利に行っていただけます。\n▼ 公式LINEで友だち追加する\n${lineUrl}\n\nSpa Blanc Aroma` + getEmailFooterText();
    GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご来店ありがとうございました', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma" });
    return { success: true, message: "ご来店お礼メールを送信しました。" };
  } catch(e) {
    return { success: false, message: "エラーが発生しました: " + e.message };
  }
}

function sendTimeAdjustment(id, newDatetime) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    let resData = null;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        targetRow = i + 1;
        resData = values[i];
        break;
      }
    }
    if (targetRow === -1) return { success: false, message: "該当する予約が見つかりません。" };

    sheet.getRange(targetRow, 4).setValue('時間調整中');
    if (newDatetime) {
      sheet.getRange(targetRow, 3).setValue(newDatetime);
    }
    clearAllAppInitCache();

    const name = resData[5];
    const course = resData[1];
    const therapist = resData[8] || "指名なし";
    const eventId = resData[7];

    if (eventId && newDatetime) {
      try {
        const event = CalendarApp.getDefaultCalendar().getEventById(eventId);
        if (event) {
          const start = new Date(String(newDatetime).replace(/-/g, '/'));
          const courseMin = parseInt(String(course).replace(/[^0-9]/g, ''), 10) || 60;
          const end = new Date(start.getTime() + courseMin * 60000);
          
          const isOutcall = String(course).startsWith('出張');
          event.setTitle(isOutcall ? `★出張【調整中】${name}様 (${course} / ${therapist})` : `【調整中】${name}様 (${course} / ${therapist})`);
          event.setTime(start, end);
        }
      } catch(err) {}
    }

    const displayNewDt = formatDateTimeForEmail(newDatetime);
    const now = new Date().getTime();

    const htmlBody = `
      <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
        <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
          <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Notification</span>
          <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご予約日時変更のお願い</h2>
        </div>
        <div style="padding: 30px 25px 0; font-size: 14px; line-height: 1.8; color: #444;">
          <p style="margin-top: 0;"><strong>${name} 様</strong></p>
          <p>いつも Spa Blanc Aroma をご利用いただき誠にありがとうございます。ご予約いただきました日時について、スケジュール調整のお願いがございます。</p>
          
          <table style="width: 100%; border-collapse: collapse; background: #faf8f5; border-radius: 6px; padding: 10px 15px; margin: 20px 0;">
            <tr><th style="text-align: left; padding: 8px 15px 8px 0; width: 110px; white-space: nowrap;">ご提案日時</th><td style="padding: 8px 0; font-weight: bold; color: #b39b7d; font-size: 15px;">${displayNewDt}</td></tr>
          </table>
          
          <p>上記のお時間でのご調整が可能でしたら、ご返信または公式LINEにてお知らせいただけますと幸いです。</p>
          <p style="margin-bottom: 0;">ご不便をおかけし誠に恐れ入りますが、何卒よろしくお願い申し上げます。</p>
          <span style="opacity: 0; font-size: 0px; display: none;">${now}</span>
        </div>
        ${getEmailFooterHtml()}
      </div>
    `;

    const textBody = `${name}様\n\nご予約の日時について変更のお願いがございます。\n変更ご提案日時: ${displayNewDt}\n\nSpa Blanc Aroma` + getEmailFooterText();
    GmailApp.sendEmail(resData[4], '【Spa Blanc Aroma】ご予約日時変更のお願い', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma" });
    return { success: true, message: "時間調整依頼メールを送信しました。" };
  } catch(e) {
    return { success: false, message: "エラーが発生しました: " + e.message };
  }
}

function getReservationList() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('reservation_list');
    if (cached) return JSON.parse(cached);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const list = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      list.push({
        index: i + 1,
        id: row[0],
        course: row[1],
        datetime: row[2],
        status: row[3],
        email: row[4],
        name: row[5],
        tel: row[6],
        eventId: row[7],
        therapist: row[8] || '指名なし',
        outcall: row[9] || ''
      });
    }

    const result = list.reverse();
    cache.put('reservation_list', JSON.stringify(result), 30);
    return result;
  } catch(e) {
    return [];
  }
}

function processApprovalApi(id, newDatetime) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    let resData = null;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        targetRow = i + 1;
        resData = values[i];
        break;
      }
    }
    if (targetRow === -1) return { success: false, message: "該当する予約が見つかりません。" };

    if (resData[3] === '予約確定' && !newDatetime) {
      return { success: false, message: "この予約は既に「予約確定」として処理済みです。" };
    }

    let targetDatetime = resData[2];
    if (newDatetime) {
      targetDatetime = newDatetime;
      sheet.getRange(targetRow, 3).setValue(targetDatetime);
    }

    sheet.getRange(targetRow, 4).setValue('予約確定');
    clearAllAppInitCache();

    let eventId = resData[7];
    const therapist = resData[8] || "指名なし";
    const name = resData[5];
    const tel = resData[6];
    const email = resData[4];
    const outcall = resData[9] || "";
    const course = resData[1];

    const isOutcall = String(course).startsWith('出張');
    const eventTitle = isOutcall ? `★出張【確定】${name}様 (${course} / ${therapist})` : `【確定】${name}様 (${course} / ${therapist})`;
    const eventLocation = outcall ? outcall : "〒533-0023 大阪市東淀川区東淡路4丁目24-35 JPアパートメント東淀川Ⅶ 102号室";
    const start = new Date(String(targetDatetime).replace(/-/g, '/'));
    const courseMin = parseInt(String(course).replace(/[^0-9]/g, ''), 10) || 60;
    const end = new Date(start.getTime() + courseMin * 60000);

    let event = null;
    if (eventId) {
      try {
        event = CalendarApp.getDefaultCalendar().getEventById(eventId);
      } catch(err) {}
    }

    if (event) {
      event.setTitle(eventTitle);
      event.setTime(start, end);
      event.setLocation(eventLocation);
    } else {
      try {
        const newEvent = CalendarApp.getDefaultCalendar().createEvent(
          eventTitle,
          start,
          end,
          {
            description: `ID: ${id}\nセラピスト: ${therapist}\nメール: ${email}\n電話: ${tel}\n出張先: ${outcall || 'なし'}`,
            location: eventLocation
          }
        );
        eventId = newEvent.getId();
        sheet.getRange(targetRow, 8).setValue(eventId);
      } catch(err) {}
    }

    sendCustomerConfirmation(resData[4], name, targetDatetime, resData[1], therapist, tel, outcall);
    return { success: true, message: "予約を「承認」し、確定メールを送信しました。" };
  } catch (e) {
    return { success: false, message: "システムエラーが発生しました: " + e.message };
  }
}

function processRejectionApi(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    let resData = null;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        targetRow = i + 1;
        resData = values[i];
        break;
      }
    }
    if (targetRow === -1) return { success: false, message: "該当する予約が見つかりません。" };

    sheet.getRange(targetRow, 4).setValue('却下');
    clearAllAppInitCache();

    const eventId = resData[7];
    if (eventId) {
      try {
        const event = CalendarApp.getDefaultCalendar().getEventById(eventId);
        if (event) event.deleteEvent();
      } catch(err) {}
    }

    sendCustomerRejection(resData[4], resData[5]);
    return { success: true, message: "予約を「却下」しました。" };
  } catch (e) {
    return { success: false, message: "システムエラーが発生しました: " + e.message };
  }
}

function saveReservation(course, datetime, name, email, tel, therapist, outcall, remarks) {
  try {
    if (checkBlacklist(email, tel)) {
      throw new Error("誠に恐れ入りますが、ご指定いただいた情報でのご予約は受け付けておりません。");
    }

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('予約台帳');
    if (!sheet) throw new Error("予約台帳シートが見つかりません");

    const id = 'R' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMddHHmmss');
    const therapistName = therapist || '指名なし';

    let eventId = "";
    try {
      const start = datetime instanceof Date ? datetime : new Date(String(datetime).replace(/-/g, '/'));
      const courseMin = parseInt(String(course).replace(/[^0-9]/g, ''), 10) || 60;
      const end = new Date(start.getTime() + courseMin * 60000);

      const isOutcall = String(course).startsWith('出張');
      const eventTitle = isOutcall ? `★出張【仮】${name}様 (${course} / ${therapistName})` : `【仮】${name}様 (${course} / ${therapistName})`;
      const eventLocation = outcall ? outcall : "〒533-0023 大阪市東淀川区東淡路4丁目24-35 JPアパートメント東淀川Ⅶ 102号室";

      const event = CalendarApp.getDefaultCalendar().createEvent(
        eventTitle,
        start,
        end,
        {
          description: `ID: ${id}\nセラピスト: ${therapistName}\nメール: ${email}\n電話: ${tel}\n出張先: ${outcall || 'なし'}\n備考: ${remarks || 'なし'}`,
          location: eventLocation
        }
      );
      eventId = event.getId();
    } catch(e) {}

    sheet.appendRow([id, course + '分', datetime, 'リクエスト中', email, name, tel, eventId, therapistName, outcall || '']);
    clearAllAppInitCache();

    try {
      sendAdminNotification(id, datetime, course, name, email, tel, therapistName, outcall, remarks);
      sendCustomerRequestReceived(email, name, datetime, course, therapistName, outcall);
    } catch(e) {}

    return { success: true };
  } catch(e) {
    throw new Error(e.message);
  }
}

function formatTimeString(val) {
  if (val == null || val === "") return "未定";
  if (val instanceof Date) {
    const h = String(val.getHours()).padStart(2, '0');
    const m = String(val.getMinutes()).padStart(2, '0');
    return `${h}:${m}`;
  }
  let str = String(val).trim();
  if (str === "休み" || str === "受付終了" || str === "未定") return str;
  let parts = str.split(':');
  if (parts.length >= 2) {
    return parts[0].padStart(2, '0') + ':' + parts[1].padStart(2, '0');
  }
  return str || "未定";
}

function getDefaultShifts(therapistName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(therapistName);
    if (!sheet) return [];
    
    const range = sheet.getRange("H2:I8");
    const values = range.getValues();
    const days = ["日曜日", "月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日"];
    
    const result = [];
    for (let i = 0; i < 7; i++) {
      let startStr = formatTimeString(values[i][0]);
      let endStr = formatTimeString(values[i][1]);
      
      result.push({
        dayIndex: i,
        dayName: days[i],
        start: startStr || "10:00",
        end: endStr || "19:00",
        isHoliday: (startStr === "休み" || endStr === "休み" || (!startStr && !endStr))
      });
    }
    return result;
  } catch(e) {
    return [];
  }
}

function updateDefaultShifts(therapistName, shifts) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(therapistName);
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(therapistName);
      sheet.appendRow(['日付', '開始時間', '終了時間', '備考']);
    }
    
    shifts.forEach(item => {
      let row = item.dayIndex + 2;
      let writeStart = item.isHoliday ? "休み" : item.start;
      let writeEnd = item.isHoliday ? "休み" : item.end;
      
      sheet.getRange("H" + row).setValue(writeStart);
      sheet.getRange("I" + row).setValue(writeEnd);
    });
    
    clearAllAppInitCache();
    return { success: true };
  } catch(e) {
    throw new Error("デフォルトシフト更新エラー: " + e.message);
  }
}

function getBusinessDays(therapistName, ssInstance) {
  try {
    const ss = ssInstance || SpreadsheetApp.getActiveSpreadsheet();
    let targetSheets = [];
    const excludeNames = ['予約台帳', 'お知らせ', 'Th.お知らせ', '設定', 'お問い合わせ内容', 'ブラックリスト', 'アンケート結果', 'セラピストアカウント'];

    if (therapistName && therapistName !== '指名なし') {
      const sheet = ss.getSheetByName(therapistName);
      if (sheet) targetSheets.push(sheet);
    } else {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        if (!excludeNames.includes(s.getName())) {
          targetSheets.push(s);
        }
      }
    }

    if (targetSheets.length === 0) return [];

    const bizDaysMap = {};
    const todayStr = normalizeDateStr(new Date());

    targetSheets.forEach(sheet => {
      const values = sheet.getDataRange().getValues();
      if (values.length <= 1) return;

      for (let i = 1; i < values.length; i++) {
        const row = values[i];
        if (!row[0]) continue;
        let dateStr = normalizeDateStr(row[0]);

        if (dateStr < todayStr) continue;

        let startVal = row[1];
        let endVal = row[2];

        if (startVal === "" || endVal === "" || startVal == null || endVal == null || startVal === "休み" || startVal === "受付終了" || startVal === "未定") continue;

        let startStr = formatTimeString(startVal);
        let endStr = formatTimeString(endVal);

        let startHour = parseInt(startStr.includes(':') ? startStr.split(':')[0] : startStr, 10);
        let endHour = parseInt(endStr.includes(':') ? endStr.split(':')[0] : endStr, 10);

        if (isNaN(startHour) || isNaN(endHour)) continue;

        if (!bizDaysMap[dateStr]) {
          bizDaysMap[dateStr] = { date: dateStr, start: startHour, end: endHour };
        } else {
          bizDaysMap[dateStr].start = Math.min(bizDaysMap[dateStr].start, startHour);
          bizDaysMap[dateStr].end = Math.max(bizDaysMap[dateStr].end, endHour);
        }
      }
    });

    return Object.values(bizDaysMap);
  } catch(e) {
    return [];
  }
}

function getBookedTimes(ssInstance, targetTherapist) {
  try {
    const ss = ssInstance || SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('予約台帳');
    if (!sheet) return [];

    const values = sheet.getDataRange().getValues();
    if (values.length < 2) return [];

    const bookedTimes = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const todayTime = today.getTime();

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (row[2] && row[3] !== '却下' && row[3] !== 'キャンセル') {
        const resTherapist = row[8] || '指名なし';
        const courseStr = String(row[1]);
        const isOutcall = courseStr.startsWith('出張');

        if (targetTherapist && targetTherapist !== '指名なし') {
          if (isOutcall && resTherapist !== targetTherapist) {
            continue;
          }
        }

        let start = (row[2] instanceof Date) ? row[2] : new Date(String(row[2]).replace(/-/g, '/'));
        if (isNaN(start.getTime())) continue;
        if (start.getTime() < todayTime) continue;

        const courseMin = parseInt(courseStr.replace(/[^0-9]/g, ''), 10) || 60;
        const startBlock = -30;
        const endBlock = courseMin + 30;

        for (let m = startBlock; m < endBlock; m += 30) {
          const busyTime = new Date(start.getTime() + m * 60000);
          const y = busyTime.getFullYear();
          const mm = String(busyTime.getMonth() + 1).padStart(2, '0');
          const dd = String(busyTime.getDate()).padStart(2, '0');
          const hh = String(busyTime.getHours()).padStart(2, '0');
          const min = String(busyTime.getMinutes()).padStart(2, '0');
          bookedTimes.push(`${y}-${mm}-${dd} ${hh}:${min}`);
        }
      }
    }
    return bookedTimes;
  } catch(e) {
    return [];
  }
}

function getTherapistScheduleData(therapist, ssInstance) {
  const t = therapist || '指名なし';
  return {
    bizDays: getBusinessDays(t, ssInstance),
    booked: getBookedTimes(ssInstance, t)
  };
}

function getTherapistList(ssInstance) {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('therapist_list');
    if (cached) return JSON.parse(cached);

    const ss = ssInstance || SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    const excludeNames = ['予約台帳', 'お知らせ', 'Th.お知らせ', '設定', 'お問い合わせ内容', 'ブラックリスト', 'アンケート結果', 'セラピストアカウント'];
    const therapists = [];

    sheets.forEach(sheet => {
      const name = sheet.getName();
      if (!excludeNames.includes(name)) {
        therapists.push(name);
      }
    });

    cache.put('therapist_list', JSON.stringify(therapists), 300);
    return therapists;
  } catch(e) {
    return [];
  }
}

function getInitialDataForIndex() {
  return {
    message: getShopMessage(),
    therapists: getTherapistList()
  };
}

function getAllSchedules(therapistName) {
  if (!therapistName) return {};
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(therapistName);
    if (!sheet) return {};
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) return {};

    const schedules = {};
    const defaultMap = {};

    try {
      const range = sheet.getRange("H2:I8").getValues();
      for (let i = 0; i < 7; i++) {
        let st = formatTimeString(range[i][0]);
        let et = formatTimeString(range[i][1]);
        defaultMap[i] = { start: st, end: et };
      }
    } catch(err) {}

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      let d = normalizeDateStr(row[0]);
      if (!d || !d.startsWith('202')) continue;

      let startStr = formatTimeString(row[1]);
      let endStr = formatTimeString(row[2]);
      let notesStr = row[3] ? String(row[3]).trim() : "";

      schedules[d] = {
        exists: true,
        rowIndex: i + 1,
        start: startStr,
        end: endStr,
        notes: notesStr
      };
    }

    const today = new Date();
    for (let i = 0; i < 14; i++) {
      const targetDate = new Date(today.getFullYear(), today.getMonth(), today.getDate() + i);
      const dateKey = normalizeDateStr(targetDate);
      const dayIdx = targetDate.getDay();
      const def = defaultMap[dayIdx] || { start: "10:00", end: "19:00" };

      if (!schedules[dateKey]) {
        if (def && def.start && def.start !== "休み" && def.start !== "未定") {
          schedules[dateKey] = {
            exists: true,
            rowIndex: -1,
            start: def.start,
            end: def.end,
            defaultStart: def.start,
            defaultEnd: def.end,
            notes: "デフォルト時間"
          };
        } else {
          schedules[dateKey] = {
            exists: false,
            rowIndex: -1,
            start: "未定",
            end: "未定",
            defaultStart: def.start || "10:00",
            defaultEnd: def.end || "19:00",
            notes: ""
          };
        }
      } else {
        schedules[dateKey].defaultStart = def.start || "10:00";
        schedules[dateKey].defaultEnd = def.end || "19:00";
      }
    }

    return schedules;
  } catch(e) {
    return {};
  }
}

function getShopMessage() {
  const defaultMsg = `<strong>【ご来店のお客様へ】</strong><br>ご希望の日時を選択してリクエストをお送りください。<br>折り返し確定メールをお送りします。`;
  try {
    const msg = PropertiesService.getScriptProperties().getProperty('SHOP_MESSAGE');
    return msg ? msg : defaultMsg;
  } catch(e) {
    return defaultMsg;
  }
}

function updateSchedule(therapistName, targetDate, start, end, isHoliday, notes) {
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(therapistName);
    if (!sheet) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(therapistName);
      sheet.appendRow(['日付', '開始時間', '終了時間', '備考']);
    }
    
    const current = getScheduleForDate(therapistName, targetDate);

    let writeStart = start;
    let writeEnd = end;
    if (isHoliday && start !== "受付終了" && start !== "未定") {
      writeStart = "休み";
      writeEnd = "休み";
    }

    if (current && current.exists && current.rowIndex && current.rowIndex > 1) {
      sheet.getRange(current.rowIndex, 2).setValue(writeStart);
      sheet.getRange(current.rowIndex, 3).setValue(writeEnd);
      sheet.getRange(current.rowIndex, 4).setValue(notes || "");
    } else {
      sheet.appendRow([targetDate, writeStart, writeEnd, notes || ""]);
    }

    clearAllAppInitCache();
    return { success: true };
  } catch(e) {
    throw new Error("更新エラー: " + e.message);
  }
}

function getScheduleForDate(therapistName, targetDate) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(therapistName);
    if (!sheet) return { exists: false, start: "未定", end: "未定", notes: "" };
    const values = sheet.getDataRange().getValues();

    if (values.length <= 1) return { exists: false, start: "未定", end: "未定", notes: "" };

    const normalizedTarget = normalizeDateStr(targetDate);
    for (let i = 1; i < values.length; i++) {
      let row = values[i];
      if (!row[0]) continue;
      let d = normalizeDateStr(row[0]);
      if (d === normalizedTarget) {
         let startStr = formatTimeString(row[1]);
         let endStr = formatTimeString(row[2]);
         let notesStr = row[3] ? String(row[3].trim()) : "";
         return { exists: true, rowIndex: i + 1, start: startStr, end: endStr, notes: notesStr };
      }
    }

    try {
      const dtObj = new Date(normalizedTarget.replace(/-/g, '/'));
      if (!isNaN(dtObj.getTime())) {
        const dayIdx = dtObj.getDay();
        const range = sheet.getRange("H2:I8").getValues();
        const st = formatTimeString(range[dayIdx][0]);
        const et = formatTimeString(range[dayIdx][1]);
        if (st && st !== "休み" && st !== "未定") {
          return { exists: true, rowIndex: -1, start: st, end: et, notes: "デフォルト" };
        }
      }
    } catch(err) {}

    return { exists: false, start: "未定", end: "未定", notes: "" };
  } catch(e) {
    return { exists: false, start: "未定", end: "未定", notes: "" };
  }
}

function verifyPassword(pass) {
  return pass === "blanc2025";
}

function getNewsList() {
  try {
    const cache = CacheService.getScriptCache();
    const cached = cache.get('news_list');
    if (cached) return JSON.parse(cached);

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('お知らせ');
    if (!sheet) return [];
    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) return [];
    const news = [];
    for (let i = 1; i < values.length; i++) {
      news.push({
        index: i + 1,
        date: normalizeDateStr(values[i][0]).replace(/-/g, '/'),
        subject: values[i][1],
        body: values[i][2]
      });
    }
    const result = news.reverse();
    cache.put('news_list', JSON.stringify(result), 300);
    return result;
  } catch(e) {
    return [];
  }
}

function getLatestNews() {
  const news = getNewsList();
  return news.length > 0 ? news[0] : null;
}

function postNews(subject, body) {
  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('お知らせ');
  if (!sheet) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('お知らせ');
    sheet.appendRow(['投稿日時', '件名', '本文']);
  }
  sheet.appendRow([new Date(), subject, body]);
  clearAllAppInitCache();
  return { success: true };
}

function deleteNews(index) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName('お知らせ').deleteRow(index);
  clearAllAppInitCache();
  return { success: true };
}

function updateNews(index, subject, body) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('お知らせ');
  sheet.getRange(index, 2).setValue(subject);
  sheet.getRange(index, 3).setValue(body);
  clearAllAppInitCache();
  return { success: true };
}

function keepAlive() {
  try {
    getAppInitData('指名なし');
  } catch(e) {}
}

// ====================================================
// X（旧Twitter）自動投稿機能（OAuth 1.0a / API v2）
// ====================================================

/**
 * スプレッドシートの「設定」シートからX投稿用の設定値を取得する
 */
function getXPostSettingsFromSheet() {
  const defaults = {
    SHOP_NAME: 'Spa Blanc Aroma',
    RESERVE_URL: 'https://spa-blanc-aroma.com/reservation.html',
    SCHEDULE_URL: 'https://spa-blanc-aroma.com/schedule.html',
    HASHTAGS: '#大阪 #淡路 #新大阪 #ゲイマッサージ #アロマトリートメント'
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('設定');
    if (!sheet) return defaults;

    const values = sheet.getDataRange().getValues();
    for (let i = 0; i < values.length; i++) {
      const key = String(values[i][0] || '').trim().toUpperCase();
      const val = String(values[i][1] || '').trim();
      if (key && val !== '') {
        defaults[key] = val;
      }
    }
  } catch (e) {
    Logger.log('設定シート読み込みエラー: ' + e.message);
  }

  return defaults;
}

/**
 * 本日の出勤状況と「設定」シートを取得して、X投稿用の親ポスト・リプライ案を自動生成する
 */
function buildTodayShiftTweetData(customComment) {
  const settings = getXPostSettingsFromSheet();
  const today = new Date();
  const days = ["日", "月", "火", "水", "木", "金", "土"];
  const dateLabel = `${today.getMonth() + 1}/${today.getDate()}(${days[today.getDay()]})`;

  const schedules = getTodaySchedule();
  const workingStaff = schedules.filter(s => s.status === "出勤");

  let mainText = "";
  const commentBlock = customComment && customComment.trim() ? `\n${customComment.trim()}\n` : "";

  if (workingStaff.length > 0) {
    const shiftLines = workingStaff.map(s => {
      const notePart = (s.notes && s.notes !== "デフォルト" && s.notes !== "デフォルト時間") ? ` ※${s.notes}` : "";
      return `・${s.name}（${s.start}〜${s.end}）${notePart}`;
    }).join("\n");

    mainText =
      `【本日 ${dateLabel} のご案内】\n\n` +
      `大阪・淡路の完全個室ゲイマッサージ\n` +
      `${settings.SHOP_NAME} です🌿\n\n` +
      `本日の出勤セラピスト：\n` +
      `${shiftLines}\n` +
      `${commentBlock}\n` +
      `新大阪駅からJRで2駅・淡路駅徒歩5分。\n` +
      `5種のアロマオイルと指圧で、心ほどけるひとときをお過ごしください。\n\n` +
      `${settings.HASHTAGS}`;
  } else {
    mainText =
      `【${dateLabel} のご案内】\n\n` +
      `大阪・淡路の完全個室ゲイマッサージ\n` +
      `${settings.SHOP_NAME} です🌿\n` +
      `${commentBlock}\n` +
      `本日のご予約受付は終了しております。\n` +
      `明日以降の出勤スケジュール・空き枠はリプ欄のリンクよりご確認いただけます。\n\n` +
      `${settings.HASHTAGS}`;
  }

  const replyText =
    `▼ 空き状況の確認・ご予約はこちら\n` +
    `${settings.RESERVE_URL}\n\n` +
    `▼ 週間スケジュール\n` +
    `${settings.SCHEDULE_URL}`;

  return {
    success: true,
    hasWorkingStaff: workingStaff.length > 0,
    mainText: mainText,
    replyText: replyText
  };
}

/**
 * 管理画面からのX投稿実行（親ポスト＋任意でリプライぶら下げ）
 */
function executePostToX(mainText, replyText, requesterId) {
  if (requesterId && !checkAdminPermission(requesterId)) {
    return { success: false, message: 'Xへの投稿は店舗管理者（admin）のみ実行可能です。' };
  }

  const cleanMain = String(mainText || '').trim();
  const cleanReply = String(replyText || '').trim();

  if (!cleanMain) {
    return { success: false, message: '投稿本文が空です。' };
  }

  try {
    const parentRes = sendSingleTweet(cleanMain, null);
    const tweetUrl = `https://x.com/spa_blancaroma/status/${parentRes.id}`;

    if (cleanReply) {
      Utilities.sleep(1500);
      try {
        sendSingleTweet(cleanReply, parentRes.id);
      } catch (replyErr) {
        return {
          success: true,
          tweetUrl: tweetUrl,
          message: `親ポストは投稿しましたが、リプ欄の投稿でエラーが発生しました: ${replyErr.message}`
        };
      }
    }

    return {
      success: true,
      tweetUrl: tweetUrl,
      message: cleanReply
        ? 'X（@spa_blancaroma）に親ポスト＋リプ欄リンクを投稿しました！'
        : 'X（@spa_blancaroma）にポストを投稿しました！'
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * 【トリガー用】毎日定時に本日の出勤情報をXへ自動投稿する関数
 */
function autoPostDailyShiftToX() {
  const draft = buildTodayShiftTweetData('');
  if (!draft.hasWorkingStaff) {
    Logger.log('本日出勤のセラピストがいないため、X自動投稿をスキップしました。');
    return;
  }
  const res = executePostToX(draft.mainText, draft.replyText, null);
  Logger.log(JSON.stringify(res));
}

/**
 * X API v2 単発ツイート送信（replyToTweetId指定時はリプライ投稿）
 */
function sendSingleTweet(text, replyToTweetId) {
  const url = 'https://api.twitter.com/2/tweets';
  const payloadObj = { text: text };
  if (replyToTweetId) {
    payloadObj.reply = { in_reply_to_tweet_id: String(replyToTweetId) };
  }

  const authHeader = buildOAuth1Header('POST', url);
  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: authHeader
    },
    payload: JSON.stringify(payloadObj),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  const body = response.getContentText();
  let json = {};
  try {
    json = JSON.parse(body);
  } catch (e) {}

  if (code === 201 && json.data && json.data.id) {
    return { success: true, id: json.data.id };
  } else {
    const detail = json.detail || json.title || body;
    throw new Error(`X APIエラー (HTTP ${code}): ${detail}`);
  }
}

/**
 * OAuth 1.0a 署名ヘッダー生成（外部ライブラリ不要）
 */
function buildOAuth1Header(method, url) {
  const props = PropertiesService.getScriptProperties();
  const apiKey = props.getProperty('X_API_KEY');
  const apiSecret = props.getProperty('X_API_SECRET');
  const accessToken = props.getProperty('X_ACCESS_TOKEN');
  const accessTokenSecret = props.getProperty('X_ACCESS_TOKEN_SECRET');

  if (!apiKey || !apiSecret || !accessToken || !accessTokenSecret) {
    throw new Error('GASのスクリプトプロパティにXのAPIキー（4種）が設定されていません。');
  }

  const nonce = Utilities.getUuid().replace(/-/g, '');
  const timestamp = Math.floor(Date.now() / 1000).toString();

  const oauthParams = {
    oauth_consumer_key: apiKey,
    oauth_nonce: nonce,
    oauth_signature_method: 'HMAC-SHA1',
    oauth_timestamp: timestamp,
    oauth_token: accessToken,
    oauth_version: '1.0'
  };

  const paramString = Object.keys(oauthParams)
    .sort()
    .map(k => `${rfc3986Encode(k)}=${rfc3986Encode(oauthParams[k])}`)
    .join('&');

  const signatureBase = `${method.toUpperCase()}&${rfc3986Encode(url)}&${rfc3986Encode(paramString)}`;
  const signingKey = `${rfc3986Encode(apiSecret)}&${rfc3986Encode(accessTokenSecret)}`;
  const rawSig = Utilities.computeHmacSignature(Utilities.MacAlgorithm.HMAC_SHA_1, signatureBase, signingKey);
  const signature = Utilities.base64Encode(rawSig);

  oauthParams.oauth_signature = signature;

  const headerParts = Object.keys(oauthParams)
    .sort()
    .map(k => `${rfc3986Encode(k)}="${rfc3986Encode(oauthParams[k])}"`)
    .join(', ');

  return `OAuth ${headerParts}`;
}

function rfc3986Encode(str) {
  return encodeURIComponent(String(str)).replace(/[!'()*]/g, function(c) {
    return '%' + c.charCodeAt(0).toString(16).toUpperCase();
  });
}