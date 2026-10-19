/*
 * 承認依頼アドイン
 * メール作成画面で「承認依頼」ボタンを押すと、
 *   1. To の宛先を承認者として読み取り
 *   2. 本文の先頭に「承認する／却下する」ボタンを差し込み
 *   3. 件名の先頭に【承認依頼】を付ける
 * 承認者がボタンを押すと、依頼者（このアドインを使った人）宛ての回答メールが作成される。
 */

const SUBJECT_PREFIX = "【承認依頼】";
const RECEIPT_LABEL = "受付番号：";

Office.onReady(() => {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("insertApproval", insertApproval);
  }
});

// 古いクライアント向けにグローバルにも公開
if (typeof window !== "undefined") {
  window.insertApproval = insertApproval;
}

async function insertApproval(event) {
  const item = Office.context.mailbox.item;
  try {
    const [toList, subject, bodyText] = await Promise.all([
      getAsyncValue(item.to),
      getAsyncValue(item.subject),
      getBodyText(item),
    ]);

    if (!toList || toList.length === 0) {
      await notify(item, "error", "To に承認者を入れてから「承認依頼」を押してください。");
      return;
    }
    if (bodyText && bodyText.indexOf(RECEIPT_LABEL + "AP-") !== -1) {
      await notify(item, "error", "承認欄はすでに入っています。やり直す場合は、本文の承認欄を削除してから押してください。");
      return;
    }

    const me = Office.context.mailbox.userProfile;
    const receiptId = makeReceiptId();
    const baseSubject = stripPrefix(subject || "");

    const html = buildApprovalBlock({
      approvers: toList,
      requesterName: me.displayName,
      requesterEmail: me.emailAddress,
      subject: baseSubject,
      receiptId: receiptId,
    });

    await prependHtml(item, html);

    if (!(subject || "").startsWith(SUBJECT_PREFIX)) {
      await setAsyncValue(item.subject, SUBJECT_PREFIX + (subject || ""));
    }

    await notify(item, "info", "承認欄を差し込みました（" + receiptId + "）。内容を確認して送信してください。");
  } catch (e) {
    await notify(item, "error", "承認欄を差し込めませんでした：" + (e && e.message ? e.message : e));
  } finally {
    if (event && event.completed) event.completed();
  }
}

/* ---------- 承認欄のHTML ---------- */

function buildApprovalBlock(o) {
  const approverNames = o.approvers
    .map((r) => escapeHtml(r.displayName || r.emailAddress) + " さん")
    .join("、");

  const approveLink = buildMailto(o.requesterEmail,
    "【承認回答】承認｜" + o.subject + "（" + o.receiptId + "）",
    "承認します。\r\n\r\nコメント：\r\n");

  const rejectLink = buildMailto(o.requesterEmail,
    "【承認回答】却下｜" + o.subject + "（" + o.receiptId + "）",
    "却下します。\r\n\r\n理由：\r\n");

  const font = "font-family:'Yu Gothic UI','Meiryo UI',Meiryo,sans-serif;";

  return (
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" ' +
      'style="width:100%;max-width:560px;border:1px solid #9cc7b3;border-left:6px solid #166e4a;' +
      'background:#f3faf6;margin:0 0 16px 0;' + font + '">' +
      '<tr><td style="padding:14px 18px;color:#1f2a24;font-size:14px;line-height:1.6;">' +
        '<div style="font-size:16px;font-weight:bold;margin:0 0 6px 0;">承認のお願い</div>' +
        '<div>承認者：' + approverNames + '</div>' +
        '<div style="margin:0 0 12px 0;">ボタンを押すと ' + escapeHtml(o.requesterName) +
          ' 宛ての回答メールが作成されます。コメントがあれば追記して送信してください。</div>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
          '<td style="padding:0 10px 0 0;">' +
            '<a href="' + approveLink + '" style="display:inline-block;padding:9px 22px;background:#166e4a;' +
            'color:#ffffff;text-decoration:none;font-weight:bold;font-size:14px;border-radius:4px;' + font + '">承認する</a>' +
          '</td>' +
          '<td>' +
            '<a href="' + rejectLink + '" style="display:inline-block;padding:8px 22px;background:#ffffff;' +
            'color:#a1262b;text-decoration:none;font-weight:bold;font-size:14px;border:1px solid #a1262b;' +
            'border-radius:4px;' + font + '">却下する</a>' +
          '</td>' +
        '</tr></table>' +
        '<div style="margin:12px 0 0 0;font-size:12px;color:#55625b;">' +
          'CC の方は回答不要です。　' + RECEIPT_LABEL + escapeHtml(o.receiptId) +
        '</div>' +
      '</td></tr>' +
    '</table>'
  );
}

function buildMailto(to, subject, body) {
  return "mailto:" + encodeURIComponent(to) +
    "?subject=" + encodeURIComponent(subject) +
    "&body=" + encodeURIComponent(body);
}

/* ---------- ユーティリティ ---------- */

function makeReceiptId() {
  const d = new Date();
  const p = (n) => String(n).padStart(2, "0");
  return "AP-" + d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) +
    "-" + p(d.getHours()) + p(d.getMinutes()) + p(d.getSeconds());
}

function stripPrefix(s) {
  return s.startsWith(SUBJECT_PREFIX) ? s.slice(SUBJECT_PREFIX.length) : s;
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function getAsyncValue(prop) {
  return new Promise((resolve, reject) => {
    prop.getAsync((r) => r.status === Office.AsyncResultStatus.Succeeded
      ? resolve(r.value) : reject(r.error));
  });
}

function setAsyncValue(prop, value) {
  return new Promise((resolve, reject) => {
    prop.setAsync(value, (r) => r.status === Office.AsyncResultStatus.Succeeded
      ? resolve() : reject(r.error));
  });
}

function getBodyText(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Text, (r) =>
      resolve(r.status === Office.AsyncResultStatus.Succeeded ? r.value : ""));
  });
}

function prependHtml(item, html) {
  return new Promise((resolve, reject) => {
    item.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve() : reject(r.error));
  });
}

function notify(item, kind, message) {
  return new Promise((resolve) => {
    const details = kind === "error"
      ? { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message }
      : {
          type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
          message: message,
          icon: "Icon.16",
          persistent: false,
        };
    item.notificationMessages.replaceAsync("approvalAddin", details, () => resolve());
  });
}
