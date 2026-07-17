/**
 * SheetLab — معالج نموذج الطلبات (Google Apps Script)
 *
 * الوظيفة:
 *  - يستقبل الطلب من نموذج sheetlab.net
 *  - يولّد رقم طلب يومي بصيغة RRQ-YYYYMMDD-NNN (يعاد من 001 كل يوم)
 *  - يسجّل الطلب في Google Sheet (أرشيف الطلبات)
 *  - يرسل إيميل HTML مصمّم للمالك
 *  - يرسل رداً تلقائياً (HTML) للعميل يحتوي رقم طلبه
 */

const OWNER_EMAIL = 'sheetlabstudio@gmail.com';
const OWNER_CC = 'khaledalnemer@gmail.com'; // نسخة إشعار إضافية بكل طلب
const BRAND = 'SheetLab';
const TIMEZONE = 'Asia/Riyadh';
const SHEET_NAME = 'Orders';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const p = (e && e.parameter) ? e.parameter : {};

    // مصيدة السبام — إذا امتلأ الحقل المخفي نتجاهل الطلب بصمت
    if (p._honey) {
      return json({ ok: true, skipped: true });
    }

    const data = {
      name:     (p['الاسم'] || '').toString().trim(),
      email:    (p['email'] || '').toString().trim(),
      whatsapp: (p['واتساب'] || '').toString().trim(),
      service:  (p['الخدمة'] || '').toString().trim(),
      details:  (p['التفاصيل'] || '').toString().trim(),
      fileLink: (p['رابط_الملف'] || '').toString().trim()
    };

    const now = new Date();
    const dateStr = Utilities.formatDate(now, TIMEZONE, 'yyyyMMdd');
    const requestId = generateRequestId(dateStr);
    const timestamp = Utilities.formatDate(now, TIMEZONE, 'yyyy-MM-dd HH:mm:ss');

    logToSheet([
      timestamp, requestId, data.name, data.email,
      data.whatsapp, data.service, data.details, data.fileLink
    ]);

    // إيميل المالك (مع نسخة إشعار إضافية)
    MailApp.sendEmail({
      to: OWNER_EMAIL,
      cc: OWNER_CC,
      subject: 'طلب جديد ' + requestId + ' — ' + (data.service || 'غير محدد'),
      htmlBody: ownerEmailHtml(requestId, timestamp, data),
      replyTo: data.email || OWNER_EMAIL,
      name: BRAND
    });

    // رد تلقائي للعميل
    if (data.email) {
      MailApp.sendEmail({
        to: data.email,
        subject: 'تأكيد استلام طلبك ' + requestId + ' — ' + BRAND,
        htmlBody: customerEmailHtml(requestId, data),
        replyTo: OWNER_EMAIL,
        name: BRAND
      });
    }

    return json({ ok: true, requestId: requestId });
  } catch (err) {
    return json({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

/** عدّاد يومي يعاد من 001 كل يوم، محفوظ في خصائص السكربت */
function generateRequestId(dateStr) {
  const props = PropertiesService.getScriptProperties();
  const key = 'count_' + dateStr;
  const current = parseInt(props.getProperty(key) || '0', 10) + 1;
  props.setProperty(key, String(current));
  const seq = ('000' + current).slice(-3);
  return 'RRQ-' + dateStr + '-' + seq;
}

/** تسجيل الطلب في ورقة Orders (تُنشأ تلقائياً مع العناوين إن لم تكن موجودة) */
function logToSheet(row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['الوقت', 'رقم الطلب', 'الاسم', 'البريد', 'واتساب', 'الخدمة', 'التفاصيل', 'رابط الملف']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow(row);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function escapeHtml(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** إيميل HTML للمالك */
function ownerEmailHtml(requestId, timestamp, d) {
  const fields = [
    { label: 'الاسم', value: d.name },
    { label: 'البريد الإلكتروني', value: d.email, type: 'email' },
    { label: 'واتساب / جوال', value: d.whatsapp },
    { label: 'الخدمة المطلوبة', value: d.service },
    { label: 'التفاصيل', value: d.details },
    { label: 'رابط الملف', value: d.fileLink, type: 'link' }
  ];

  const rows = fields.map(function (f) {
    let val;
    if (!f.value) {
      val = '<span style="color:#a1a1aa;">—</span>';
    } else if (f.type === 'email') {
      val = '<a href="mailto:' + escapeHtml(f.value) + '" style="color:#0f766e;text-decoration:none;">' + escapeHtml(f.value) + '</a>';
    } else if (f.type === 'link') {
      val = '<a href="' + escapeHtml(f.value) + '" style="color:#0f766e;text-decoration:none;">' + escapeHtml(f.value) + '</a>';
    } else {
      val = escapeHtml(f.value);
    }
    return '' +
      '<tr>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #e4e4e7;font-size:13px;font-weight:600;color:#3f3f46;white-space:nowrap;vertical-align:top;background:#fafafa;">' + f.label + '</td>' +
        '<td style="padding:12px 16px;border-bottom:1px solid #e4e4e7;font-size:14px;color:#18181b;line-height:1.7;">' + val + '</td>' +
      '</tr>';
  }).join('');

  return baseEmail(
    '<tr><td style="padding:0 0 20px;">' +
      '<div style="display:inline-block;background:#f0fdfa;border:1px solid #99f6e4;color:#0f766e;font-size:13px;font-weight:700;letter-spacing:.02em;padding:8px 14px;border-radius:9999px;">' +
        'رقم الطلب: ' + escapeHtml(requestId) +
      '</div>' +
      '<div style="color:#71717a;font-size:12px;margin-top:10px;">' + escapeHtml(timestamp) + '</div>' +
    '</td></tr>' +
    '<tr><td>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #e4e4e7;border-radius:10px;border-collapse:separate;overflow:hidden;">' +
        rows +
      '</table>' +
    '</td></tr>',
    'طلب جديد'
  );
}

/** رد تلقائي HTML للعميل — بلغة خليجية ترحيبية (RTL) */
function customerEmailHtml(requestId, d) {
  const hi = d.name ? ('هلا والله فيك يا ' + escapeHtml(d.name) + '،') : 'هلا والله وغلا،';
  return baseEmail(
    '<tr><td dir="rtl" style="padding:0 0 14px;text-align:right;color:#18181b;font-size:17px;font-weight:700;">' + hi + '</td></tr>' +
    '<tr><td dir="rtl" style="padding:0 0 18px;text-align:right;color:#3f3f46;font-size:14px;line-height:1.9;">' +
      'حيّاك الله وأهلين فيك في ' + BRAND + '. وصلنا طلبك تمام، وإحنا سعداء إنك خصّيتنا بثقتك. ' +
      'بنراجع طلبك عن قرب، ونرجع لك بأقرب وقت بعرض سعر واضح ومدة تنفيذ — بدون لف ولا دوران.' +
    '</td></tr>' +
    '<tr><td style="padding:0 0 18px;">' +
      '<div dir="rtl" style="background:#f0fdfa;border:1px solid #99f6e4;border-radius:10px;padding:16px 18px;text-align:center;">' +
        '<div style="color:#0f766e;font-size:12px;font-weight:600;margin-bottom:6px;">رقم طلبك</div>' +
        '<div style="color:#134e4a;font-size:20px;font-weight:700;letter-spacing:.03em;">' + escapeHtml(requestId) + '</div>' +
      '</div>' +
    '</td></tr>' +
    (d.service ? ('<tr><td dir="rtl" style="padding:0 0 10px;text-align:right;color:#71717a;font-size:13px;">الخدمة اللي طلبتها: <span style="color:#18181b;font-weight:600;">' + escapeHtml(d.service) + '</span></td></tr>') : '') +
    '<tr><td dir="rtl" style="padding:0 0 18px;text-align:right;color:#3f3f46;font-size:14px;line-height:1.9;">' +
      'لا يهمك، إحنا معك خطوة بخطوة. تكفى احتفظ برقم طلبك فوق عشان نخدمك أسرع لو تواصلت معنا.' +
    '</td></tr>' +
    '<tr><td dir="rtl" style="padding:0 0 4px;text-align:right;color:#18181b;font-size:14px;font-weight:600;">' +
      'شاكرين لك ثقتك، وحيّاك الله دايم.' +
    '</td></tr>' +
    '<tr><td dir="rtl" style="padding:0 0 14px;text-align:right;color:#71717a;font-size:13px;">فريق ' + BRAND + '</td></tr>' +
    '<tr><td dir="rtl" style="padding:12px 0 0;text-align:right;color:#a1a1aa;font-size:12px;line-height:1.7;border-top:1px solid #f4f4f5;">' +
      'خصوصيتك تهمّنا: ملفاتك تُخزَّن محلياً وتُحذف خلال 14 يوماً من التسليم.' +
    '</td></tr>',
    'تم استلام طلبك'
  );
}

/** الهيكل العام للإيميل (رأس + محتوى + تذييل) */
function baseEmail(contentRows, heading) {
  return '' +
  '<!DOCTYPE html><html lang="ar" dir="rtl"><head><meta charset="UTF-8">' +
  '<meta name="viewport" content="width=device-width, initial-scale=1.0"></head>' +
  '<body style="margin:0;padding:0;background:#f4f4f5;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f4f4f5;padding:32px 12px;">' +
      '<tr><td align="center">' +
        '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border:1px solid #e4e4e7;border-radius:14px;overflow:hidden;font-family:\'Segoe UI\',Tahoma,Arial,sans-serif;">' +
          // header
          '<tr><td style="background:#18181b;padding:22px 28px;">' +
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
              '<td style="color:#ffffff;font-size:18px;font-weight:700;letter-spacing:-.01em;">' + BRAND + '</td>' +
              '<td align="left" style="color:#a1a1aa;font-size:13px;">' + escapeHtml(heading) + '</td>' +
            '</tr></table>' +
          '</td></tr>' +
          // content
          '<tr><td style="padding:28px;">' +
            '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + contentRows + '</table>' +
          '</td></tr>' +
          // footer
          '<tr><td style="background:#fafafa;border-top:1px solid #e4e4e7;padding:18px 28px;color:#a1a1aa;font-size:12px;text-align:center;">' +
            '© ' + BRAND + ' — sheetlab.net' +
          '</td></tr>' +
        '</table>' +
      '</td></tr>' +
    '</table>' +
  '</body></html>';
}
