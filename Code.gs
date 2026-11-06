const GAS_URL = 'https://script.google.com/macros/s/AKfycbzINJ6cRRGnjW4n1n1zW42LDRqziMAkOKwDZrwMhMWnNdmO74K815ld_gbcnB94eh8/exec';
const SITE_URL = 'https://spa-blanc-aroma.com/approve.html'; // 承認画面URL
const ADMIN_EMAIL = 'spa.blanc.aroma@gmail.com'; // 管理者通知の宛先アドレス

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
      } else if (action === 'verifyPassword') {
        result = verifyPassword(e.parameter.pass);
      } else if (action === 'getReservationList') {
        result = getReservationList();
      } else if (action === 'getTodaySchedule') {
        result = getTodaySchedule();
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
    }
  } catch (err) {
    result = { success: false, error: err.message };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

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
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');
    const todaySchedules = [];

    therapists.forEach(tName => {
      const sched = getScheduleForDate(tName, todayStr);
      let status = "休み";
      let start = "休み";
      let end = "";
      
      if (sched.exists && sched.start && sched.start !== "休み") {
        status = "出勤";
        start = sched.start;
        end = sched.end;
      }
      
      todaySchedules.push({
        name: tName,
        status: status,
        start: start,
        end: end,
        notes: sched.notes || ""
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
    return Utilities.formatDate(rawDate, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  let str = String(rawDate).trim().replace(/\//g, '-');
  if (str.includes('T')) str = str.split('T')[0];
  const parts = str.split('-');
  if (parts.length === 3) {
    return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
  }
  return str;
}

function clearAllAppInitCache() {
  try {
    const cache = CacheService.getScriptCache();
    cache.removeAll(['news_list', 'therapist_list', 'reservation_list']);
  } catch(e){}
}

function formatDateTimeForEmail(dt) {
  try {
    const d = new Date(dt);
    if (isNaN(d.getTime())) return dt;
    const days = ["日", "月", "火", "水", "木", "金", "土"];
    const datePart = Utilities.formatDate(d, 'Asia/Tokyo', 'M月d日');
    const timePart = Utilities.formatDate(d, 'Asia/Tokyo', 'HH:mm');
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
    text = isOutcall ? "出張60分" : "60分";
    price = isOutcall ? "10,000円〜 + 交通費" : "9,000円";
  } else if (min === 90) {
    text = isOutcall ? "出張90分" : "90分";
    price = isOutcall ? "13,000円〜 + 交通費" : "12,000円";
  } else if (min === 120) {
    text = isOutcall ? "出張120分" : "120分";
    price = isOutcall ? "15,000円〜 + 交通費" : "14,000円";
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
    if (b1 && String(b1).includes("@")) {
      return String(b1).trim();
    }
  } catch(e) {}
  return ADMIN_EMAIL;
}

function getStartRowIndex(sheetName) {
  return sheetName === 'そうた' ? 1 : 2;
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
      <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8; color: #e74c3c;">出張先</th>
          <td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold;">
            ${outcall}<br>
            <a href="${mapSearchUrl}" target="_blank" style="font-size: 12px; color: #b39b7d; text-decoration: underline;">📍 Googleマップで場所を確認する</a>
          </td>
      </tr>
    `;
  }

  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333;">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530;">【要確認】新規予約リクエスト</h2>
      </div>
      <div style="padding: 30px 25px; font-size: 14px; line-height: 1.8;">
        <table style="width: 100%; border-collapse: collapse;">
          <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8; width: 110px;">希望セラピスト</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8; font-weight: bold; color: #b39b7d;">${therapistText}</td></tr>
          <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8;">日時</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${displayDt}</td></tr>
          <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8;">コース</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${courseInfo.text} (${courseInfo.price})</td></tr>
          ${outcallMapHtml}
          <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8;">お名前</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${name} 様</td></tr>
          <tr><th style="text-align: left; padding: 8px 0; border-bottom: 1px solid #f0ede8;">電話番号</th><td style="padding: 8px 0; border-bottom: 1px solid #f0ede8;">${tel}</td></tr>
          <tr><th style="text-align: left; padding: 8px 0;">メール</th><td style="padding: 8px 0;">${email}</td></tr>
        </table>
        <div style="text-align: center; margin-top: 25px;">
          <a href="${approveUrl}" style="background-color: #b39b7d; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 30px; margin-right: 8px;">承認する</a>
          <a href="${rejectUrl}" style="background-color: #888888; color: #ffffff; padding: 12px 24px; text-decoration: none; border-radius: 30px;">却下する</a>
        </div>
      </div>
    </div>
  `;

  const textBody = `希望セラピスト: ${therapistText}\n内容: ${displayDt} ${courseInfo.text}(${courseInfo.price})\n${outcall ? '出張先: ' + outcall + '\n' : ''}お名前: ${name} 様 (電話: ${tel})\n\n承認: ${approveUrl}\n却下: ${rejectUrl}`;
  GmailApp.sendEmail(targetEmail, '【要確認】新規予約リクエスト', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma" });
}

function sendCustomerRequestReceived(email, name, dt, course, therapist, outcall) {
  const displayDt = formatDateTimeForEmail(dt);
  const courseInfo = getCourseDetails(course);
  const therapistText = therapist || "指名なし";

  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333;">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530;">ご予約リクエストを承りました</h2>
      </div>
      <div style="padding: 30px 25px; font-size: 14px; line-height: 1.8;">
        <p><strong>${name} 様</strong></p>
        <p>ご予約リクエストを承りました。以下の内容で確認中です。</p>
        <p>担当セラピスト： <strong>${therapistText}</strong><br>日時： <strong>${displayDt}</strong><br>コース： <strong>${courseInfo.text} (${courseInfo.price})</strong></p>
      </div>
    </div>
  `;
  const textBody = `${name} 様\n\nご予約リクエストを承りました。\n担当: ${therapistText}\n日時: ${displayDt}\nコース: ${courseInfo.text} (${courseInfo.price})\n\nSpa Blanc Aroma`;
  GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご予約リクエストを承りました', textBody, { htmlBody: htmlBody, name: "Spa Blanc Aroma"});
}

function sendCustomerConfirmation(email, name, dt, course, therapist, tel, outcall) {
  const mapUrl = "https://maps.google.com/maps?q=大阪市東淀川区東淡路4丁目24-35";
  const displayDt = formatDateTimeForEmail(dt);
  const courseInfo = getCourseDetails(course);
  const therapistText = therapist || "指名なし";
  const targetStaffEmail = getTherapistEmail(therapist);

  const customerHtmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333;">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530;">ご予約が確定いたしました</h2>
      </div>
      <div style="padding: 30px 25px; font-size: 14px; line-height: 1.8;">
        <p><strong>${name} 様</strong></p>
        <p>ご予約が<strong style="color: #b39b7d;">【確定】</strong>いたしました。</p>
        <p>担当セラピスト： <strong>${therapistText}</strong><br>日時： <strong>${displayDt}</strong><br>コース： <strong>${courseInfo.text} (${courseInfo.price})</strong></p>
      </div>
    </div>
  `;
  const customerTextBody = `${name}様\n\nご予約が【確定】いたしました。\n担当: ${therapistText}\n日時: ${displayDt}\nコース: ${courseInfo.text}\n\nSpa Blanc Aroma`;
  GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご予約が確定いたしました', customerTextBody, { htmlBody: customerHtmlBody, name: "Spa Blanc Aroma"});

  const staffHtmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333;">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center;">
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530;">【確定】予約が確定しました</h2>
      </div>
      <div style="padding: 30px 25px; font-size: 14px; line-height: 1.8;">
        <p>担当：<strong>${therapistText}</strong> / お名前：<strong>${name}様</strong><br>日時：${displayDt}<br>コース：${courseInfo.text} (${courseInfo.price})</p>
      </div>
    </div>
  `;
  const staffTextBody = `下記予約を確定しました。\n担当: ${therapistText}\nお名前: ${name}様\n日時: ${displayDt}\nコース: ${courseInfo.text}`;
  GmailApp.sendEmail(targetStaffEmail, `【確定】予約確定のお知らせ (${name}様)`, staffTextBody, { htmlBody: staffHtmlBody, name: "Spa Blanc Aroma"});
}

function sendCustomerRejection(email, name) {
  const now = new Date().getTime();
  const htmlBody = `
    <div style="max-width: 520px; margin: 0 auto; background-color: #ffffff; border: 1px solid #f0ede8; border-radius: 8px; font-family: 'Shippori Mincho', serif; color: #333333; overflow: hidden; box-shadow: 0 4px 15px rgba(0,0,0,0.03);">
      <div style="background-color: #f2ece4; padding: 25px 20px; text-align: center; border-bottom: 1px solid #e0dcd5;">
        <span style="display:block; font-family: 'Montserrat', sans-serif; font-size: 11px; letter-spacing: 0.25em; color: #b39b7d; text-transform: uppercase; margin-bottom: 6px;">Reservation Notification</span>
        <h2 style="margin: 0; font-size: 18px; font-weight: 500; color: #3b3530; letter-spacing: 0.05em;">ご予約リクエストに関するお知らせ</h2>
      </div>
      <div style="padding: 30px 25px; font-size: 14px; line-height: 1.8; color: #444;">
        <p style="margin-top: 0;"><strong>${name} 様</strong></p>
        <p>せっかくリクエスト頂いたのに誠に申し訳ございませんが、ご希望の日時ではご予約を承ることができませんでした。</p>
        <p style="margin-bottom: 0;">またの機会のご利用を心よりお待ちしております。</p>
        <span style="opacity: 0; font-size: 0px; display: none;">${now}</span>
      </div>
    </div>
  `;
  const textBody = `${name}様\n\nせっかくリクエスト頂いたのに誠に申し訳ございませんが、ご希望の日時ではご予約を承ることができませんでした。\n\nまたの機会のご利用を心よりお待ちしております。\n\nSpa Blanc Aroma`;
  
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

    const textBody = `${name}様\n\n本日は Spa Blanc Aroma にお越しいただき、誠にありがとうございました。\n\n▼ アンケートご協力\n${surveyUrl}\n\n▼ 公式LINE\n${lineUrl}\n\nSpa Blanc Aroma`;
    GmailApp.sendEmail(email, '【Spa Blanc Aroma】ご来店ありがとうございました', textBody, { name: "Spa Blanc Aroma" });
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

    sheet.getRange(targetRow, 3).setValue(newDatetime);
    sheet.getRange(targetRow, 4).setValue('時間調整中');
    clearAllAppInitCache();

    const textBody = `${resData[5]}様\n\nご予約の日時について変更のお願いがございます。\n変更ご提案日時: ${formatDateTimeForEmail(newDatetime)}\n\nSpa Blanc Aroma`;
    GmailApp.sendEmail(resData[4], '【Spa Blanc Aroma】ご予約日時変更のお願い', textBody, { name: "Spa Blanc Aroma" });
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

    if (resData[3] === '予約確定') {
      return { success: false, message: "この予約は既に「予約確定」として処理済みです。" };
    }

    let targetDatetime = resData[2];
    if (newDatetime) {
      targetDatetime = newDatetime;
      sheet.getRange(targetRow, 3).setValue(targetDatetime);
    }

    sheet.getRange(targetRow, 4).setValue('予約確定');
    clearAllAppInitCache();

    const eventId = resData[7];
    const therapist = resData[8] || "指名なし";
    const name = resData[5];
    const tel = resData[6];
    const outcall = resData[9] || "";
    const course = resData[1];

    if (eventId) {
      try {
        const event = CalendarApp.getDefaultCalendar().getEventById(eventId);
        if (event) {
          const isOutcall = String(course).startsWith('出張');
          event.setTitle(isOutcall ? `★出張【確定】${name}様 (${course} / ${therapist})` : `【確定】${name}様 (${course} / ${therapist})`);
          if (newDatetime) {
            const start = new Date(String(newDatetime).replace(/-/g, '/'));
            const courseMin = parseInt(String(course).replace(/[^0-9]/g, ''), 10) || 60;
            event.setTime(start, new Date(start.getTime() + courseMin * 60000));
          }
        }
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

function getBusinessDays(therapistName, ssInstance) {
  try {
    const ss = ssInstance || SpreadsheetApp.getActiveSpreadsheet();
    let targetSheets = [];
    const excludeNames = ['予約台帳', 'お知らせ', '設定', 'お問い合わせ内容', 'ブラックリスト', 'アンケート結果'];

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
    const todayStr = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd');

    targetSheets.forEach(sheet => {
      const values = sheet.getDataRange().getValues();
      const sheetName = sheet.getName();
      const startRowIdx = getStartRowIndex(sheetName);

      if (values.length <= startRowIdx) return;

      for (let i = startRowIdx; i < values.length; i++) {
        const row = values[i];
        if (!row[0]) continue;
        let dateStr = normalizeDateStr(row[0]);

        if (dateStr < todayStr) continue;
        if (row[1] === "" || row[2] === "" || row[1] == null || row[2] == null || row[1] === "休み") continue;

        let startStr = (row[1] instanceof Date) ? Utilities.formatDate(row[1], 'Asia/Tokyo', 'HH:mm') : String(row[1]);
        let endStr = (row[2] instanceof Date) ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'HH:mm') : String(row[2]);

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
          bookedTimes.push(Utilities.formatDate(busyTime, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm'));
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
    const excludeNames = ['予約台帳', 'お知らせ', '設定', 'お問い合わせ内容', 'ブラックリスト','アンケート結果'];
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
    const startRowIdx = getStartRowIndex(therapistName);

    if (values.length <= startRowIdx) return {};

    const schedules = {};
    for (let i = startRowIdx; i < values.length; i++) {
      const row = values[i];
      if (!row[0]) continue;
      let d = normalizeDateStr(row[0]);
      let startStr = (row[1] instanceof Date) ? Utilities.formatDate(row[1], 'Asia/Tokyo', 'HH:mm') : String(row[1]);
      let endStr = (row[2] instanceof Date) ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'HH:mm') : String(row[2]);
      let notesStr = row[3] ? String(row[3]).trim() : "";

      schedules[d] = {
        exists: true,
        rowIndex: i + 1,
        start: startStr,
        end: endStr,
        notes: notesStr
      };
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

    let writeStart = isHoliday ? "休み" : start;
    let writeEnd = isHoliday ? "休み" : end;
    if (current.exists) {
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
    if (!sheet) return { exists: false, start: "", end: "", notes: "" };
    const values = sheet.getDataRange().getValues();
    const startRowIdx = getStartRowIndex(therapistName);

    if (values.length <= startRowIdx) return { exists: false, start: "", end: "", notes: "" };

    const normalizedTarget = normalizeDateStr(targetDate);
    for (let i = startRowIdx; i < values.length; i++) {
      let row = values[i];
      if (!row[0]) continue;
      let d = normalizeDateStr(row[0]);
      if (d === normalizedTarget) {
         let startStr = (row[1] instanceof Date) ? Utilities.formatDate(row[1], 'Asia/Tokyo', 'HH:mm') : String(row[1]);
         let endStr = (row[2] instanceof Date) ? Utilities.formatDate(row[2], 'Asia/Tokyo', 'HH:mm') : String(row[2]);
         let notesStr = row[3] ? String(row[3].trim()) : "";
         return { exists: true, rowIndex: i + 1, start: startStr, end: endStr, notes: notesStr };
      }
    }
    return { exists: false, start: "", end: "", notes: "" };
  } catch(e) {
    return { exists: false, start: "", end: "", notes: "" };
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
        date: Utilities.formatDate(new Date(values[i][0]), 'Asia/Tokyo', 'yyyy/MM/dd'),
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