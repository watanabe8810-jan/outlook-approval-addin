/*
 * 承認依頼アドイン
 * メール作成画面で「承認依頼」ボタンを押すと、
 *   1. To の宛先を承認者として読み取り
 *   2. 本文の先頭に「承認／却下」ボタンと定型文を差し込み
 *   3. 件名の「Re:【ECN判定】」までを削除して【ECN承認依頼】に置き換える
 * 承認者がボタンを押すと、依頼者（このアドインを使った人）宛ての回答メールが作成される。
 */

const SUBJECT_PREFIX = "【ECN承認依頼】";
const SOURCE_MARKER = "【ECN判定】";
const ANSWER_PREFIX = "【承認回答】";
// 本文HTML内のmailtoリンクに含まれる文字列で、承認欄の重複を判定する
const DUPLICATE_MARKER = encodeURIComponent(ANSWER_PREFIX);

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
    const [toList, subject, bodyHtml] = await Promise.all([
      getAsyncValue(item.to),
      getAsyncValue(item.subject),
      getBodyHtml(item),
    ]);

    if (!toList || toList.length === 0) {
      await notify(item, "error", "To に承認者を入れてから「承認依頼」を押してください。");
      return;
    }
    if (bodyHtml && (bodyHtml.indexOf(DUPLICATE_MARKER) !== -1 || bodyHtml.indexOf(ANSWER_PREFIX) !== -1)) {
      await notify(item, "error", "承認欄はすでに入っています。やり直す場合は、本文の承認欄を削除してから押してください。");
      return;
    }

    const me = Office.context.mailbox.userProfile;
    const receiptId = makeReceiptId();
    const baseSubject = extractBaseSubject(subject || "");

    const html =
      buildApprovalBlock({
        requesterEmail: me.emailAddress,
        subject: baseSubject,
        receiptId: receiptId,
      }) +
      buildMessageText({
        approvers: toList,
        senderDisplayName: me.displayName,
      });

    await prependHtml(item, html);
    await setAsyncValue(item.subject, SUBJECT_PREFIX + baseSubject);

    await notify(item, "info", "承認欄を差し込みました（" + receiptId + "）。内容を確認して送信してください。");
  } catch (e) {
    await notify(item, "error", "承認欄を差し込めませんでした：" + (e && e.message ? e.message : e));
  } finally {
    if (event && event.completed) event.completed();
  }
}

/* ---------- 承認欄のHTML ---------- */

function buildApprovalBlock(o) {
  const approveLink = buildMailto(o.requesterEmail,
    ANSWER_PREFIX + "承認｜" + o.subject + "（" + o.receiptId + "）",
    "承認します。\r\n\r\nコメント：\r\n");

  const rejectLink = buildMailto(o.requesterEmail,
    ANSWER_PREFIX + "却下｜" + o.subject + "（" + o.receiptId + "）",
    "却下します。\r\n\r\n理由：\r\n");

  const font = "font-family:'Yu Gothic UI','Meiryo UI',Meiryo,sans-serif;";

  return (
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" ' +
      'style="width:100%;max-width:560px;border:1px solid #9cc7b3;border-left:6px solid #166e4a;' +
      'background:#f3faf6;margin:0 0 16px 0;' + font + '">' +
      '<tr><td style="padding:14px 18px 18px 18px;color:#1f2a24;font-size:14px;line-height:1.6;">' +
        '<div style="font-size:15px;font-weight:bold;margin:0 0 10px 0;">承認依頼</div>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
          '<td style="padding:0 10px 0 0;">' +
            '<a href="' + approveLink + '" style="display:inline-block;padding:9px 22px;background:#166e4a;' +
            'color:#ffffff;text-decoration:none;font-weight:bold;font-size:14px;border-radius:4px;' + font + '">承認</a>' +
          '</td>' +
          '<td>' +
            '<a href="' + rejectLink + '" style="display:inline-block;padding:8px 22px;background:#ffffff;' +
            'color:#a1262b;text-decoration:none;font-weight:bold;font-size:14px;border:1px solid #a1262b;' +
            'border-radius:4px;' + font + '">却下</a>' +
          '</td>' +
        '</tr></table>' +
      '</td></tr>' +
    '</table>'
  );
}

/* ---------- 定型文 ---------- */

function buildMessageText(o) {
  const addressee = o.approvers
    .map((r) => familyName(r.displayName || r.emailAddress) + "様")
    .join("、");
  const sender = familyName(o.senderDisplayName) || "渡邊";

  const lines = [
    escapeHtml(addressee),
    "",
    escapeHtml(sender) + "です。",
    "",
    "標記の件、関連部署でのECN判定が完了しました。",
    "承認いただける際は、投票ボタンの「承認」をお願いします。",
    "",
    "　リンク：",
    "",
    "以上です。",
    "",
  ];
  return lines.map((l) => "<div>" + (l === "" ? "<br>" : l) + "</div>").join("");
}

// 「Jun Kitagawa(北川　淳)」→「北川」、括弧がなければ表示名をそのまま使う
function familyName(displayName) {
  const s = String(displayName || "").trim();
  const m = s.match(/[（(]([^）)]+)[）)]/);
  if (m) {
    return m[1].trim().split(/[\s　]+/)[0];
  }
  return s;
}

function buildMailto(to, subject, body) {
  return "mailto:" + to +
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

// 「Re:【ECN判定】xxx」→「xxx」
// 【ECN判定】が無い場合は、先頭の Re:/Fw: や既存の承認依頼の印だけを取り除く
function extractBaseSubject(s) {
  const idx = s.indexOf(SOURCE_MARKER);
  let rest;
  if (idx !== -1) {
    rest = s.slice(idx + SOURCE_MARKER.length);
  } else {
    rest = s
      .replace(/^\s*((re|fw|fwd|返信|転送)\s*[:：]\s*)+/i, "")
      .replace(/^【ECN承認依頼】|^【承認依頼】/, "");
  }
  return rest.trim();
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

function getBodyHtml(item) {
  return new Promise((resolve) => {
    item.body.getAsync(Office.CoercionType.Html, (r) =>
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
