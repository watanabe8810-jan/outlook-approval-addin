/*
 * \u627f\u8a8d\u4f9d\u983c\u30a2\u30c9\u30a4\u30f3\uff08\u65b0\u3057\u3044Outlook\uff0f\u30af\u30e9\u30b7\u30c3\u30afOutlook \u4e21\u5bfe\u5fdc\u7248\uff09
 * \u30e1\u30fc\u30eb\u4f5c\u6210\u753b\u9762\u3067\u300c\u627f\u8a8d\u4f9d\u983c\u300d\u30dc\u30bf\u30f3\u3092\u62bc\u3059\u3068\u3001
 *   1. To \u306e\u5b9b\u5148\u3092\u627f\u8a8d\u8005\u3068\u3057\u3066\u8aad\u307f\u53d6\u308a
 *   2. \u672c\u6587\u306e\u5148\u982d\u306b\u300c\u627f\u8a8d\uff0f\u5374\u4e0b\u300d\u30dc\u30bf\u30f3\u3068\u5b9a\u578b\u6587\u3092\u5dee\u3057\u8fbc\u307f
 *   3. \u4ef6\u540d\u306e\u300cRe:\u3010ECN\u5224\u5b9a\u3011\u300d\u307e\u3067\u3092\u524a\u9664\u3057\u3066\u3010ECN\u627f\u8a8d\u4f9d\u983c\u3011\u306b\u7f6e\u304d\u63db\u3048\u308b
 *
 * \u30af\u30e9\u30b7\u30c3\u30afOutlook\u306f\u74b0\u5883\u306b\u3088\u3063\u3066\u53e4\u3044\u30d6\u30e9\u30a6\u30b6\u90e8\u54c1\uff08Internet Explorer 11\uff09\u3067
 * \u30a2\u30c9\u30a4\u30f3\u3092\u52d5\u304b\u3059\u305f\u3081\u3001\u3053\u306e\u30d5\u30a1\u30a4\u30eb\u306f\u53e4\u3044\u66f8\u304d\u65b9\uff08ES5\uff09\u3060\u3051\u3067\u66f8\u3044\u3066\u3044\u308b\u3002
 *   \u4f7f\u308f\u306a\u3044\u3082\u306e\uff1a\u30a2\u30ed\u30fc\u95a2\u6570\u3001async/await\u3001Promise\u3001\u30c6\u30f3\u30d7\u30ec\u30fc\u30c8\u6587\u5b57\u5217\u3001
 *                 startsWith\u3001padStart\u3001const/let
 */

var SUBJECT_PREFIX = "\u3010ECN\u627f\u8a8d\u4f9d\u983c\u3011";
var SOURCE_MARKER = "\u3010ECN\u5224\u5b9a\u3011";
var ANSWER_PREFIX = "\u3010\u627f\u8a8d\u56de\u7b54\u3011";
// \u672c\u6587HTML\u5185\u306emailto\u30ea\u30f3\u30af\u306b\u542b\u307e\u308c\u308b\u6587\u5b57\u5217\u3067\u3001\u627f\u8a8d\u6b04\u306e\u91cd\u8907\u3092\u5224\u5b9a\u3059\u308b
var DUPLICATE_MARKER = encodeURIComponent(ANSWER_PREFIX);

Office.onReady(function () {
  if (Office.actions && Office.actions.associate) {
    Office.actions.associate("insertApproval", insertApproval);
  }
});

// \u30af\u30e9\u30b7\u30c3\u30afOutlook\u306f\u30b0\u30ed\u30fc\u30d0\u30eb\u95a2\u6570\u540d\u3067\u547c\u3073\u51fa\u3059
window.insertApproval = insertApproval;

function insertApproval(event) {
  var item = Office.context.mailbox.item;
  var done = false;
  function finish() {
    if (done) return;
    done = true;
    if (event && event.completed) event.completed();
  }
  function fail(e) {
    notify(item, "error", "\u627f\u8a8d\u6b04\u3092\u5dee\u3057\u8fbc\u3081\u307e\u305b\u3093\u3067\u3057\u305f\uff1a" + errorText(e), finish);
  }

  try {
    item.to.getAsync(function (rTo) {
      if (rTo.status !== Office.AsyncResultStatus.Succeeded) return fail(rTo.error);
      var toList = rTo.value || [];
      if (toList.length === 0) {
        return notify(item, "error", "To \u306b\u627f\u8a8d\u8005\u3092\u5165\u308c\u3066\u304b\u3089\u300c\u627f\u8a8d\u4f9d\u983c\u300d\u3092\u62bc\u3057\u3066\u304f\u3060\u3055\u3044\u3002", finish);
      }

      item.subject.getAsync(function (rSub) {
        if (rSub.status !== Office.AsyncResultStatus.Succeeded) return fail(rSub.error);
        var subject = rSub.value || "";

        item.body.getAsync(Office.CoercionType.Html, function (rBody) {
          var bodyHtml = rBody.status === Office.AsyncResultStatus.Succeeded ? (rBody.value || "") : "";
          if (bodyHtml.indexOf(DUPLICATE_MARKER) !== -1 || bodyHtml.indexOf(ANSWER_PREFIX) !== -1) {
            return notify(item, "error",
              "\u627f\u8a8d\u6b04\u306f\u3059\u3067\u306b\u5165\u3063\u3066\u3044\u307e\u3059\u3002\u3084\u308a\u76f4\u3059\u5834\u5408\u306f\u3001\u672c\u6587\u306e\u627f\u8a8d\u6b04\u3092\u524a\u9664\u3057\u3066\u304b\u3089\u62bc\u3057\u3066\u304f\u3060\u3055\u3044\u3002", finish);
          }

          var me = Office.context.mailbox.userProfile;
          var receiptId = makeReceiptId();
          var baseSubject = extractBaseSubject(subject);

          var html =
            buildApprovalBlock({
              requesterEmail: me.emailAddress,
              subject: baseSubject,
              receiptId: receiptId
            }) +
            buildMessageText({
              approvers: toList,
              senderDisplayName: me.displayName
            });

          item.body.prependAsync(html, { coercionType: Office.CoercionType.Html }, function (rPre) {
            if (rPre.status !== Office.AsyncResultStatus.Succeeded) return fail(rPre.error);

            item.subject.setAsync(SUBJECT_PREFIX + baseSubject, function (rSet) {
              if (rSet.status !== Office.AsyncResultStatus.Succeeded) return fail(rSet.error);
              notify(item, "info",
                "\u627f\u8a8d\u6b04\u3092\u5dee\u3057\u8fbc\u307f\u307e\u3057\u305f\uff08" + receiptId + "\uff09\u3002\u5199\u771f\u3068\u30ea\u30f3\u30af\u3092\u8ffd\u52a0\u3057\u3066\u9001\u4fe1\u3057\u3066\u304f\u3060\u3055\u3044\u3002", finish);
            });
          });
        });
      });
    });
  } catch (e) {
    fail(e);
  }
}

/* ---------- \u627f\u8a8d\u6b04\u306eHTML ---------- */
/* \u30af\u30e9\u30b7\u30c3\u30afOutlook\u306fWord\u3067\u30e1\u30fc\u30eb\u3092\u8868\u793a\u3059\u308b\u305f\u3081\u3001<a>\u306e\u4f59\u767d\u3084\u89d2\u4e38\u304c\u52b9\u304b\u306a\u3044\u3002
   \u30dc\u30bf\u30f3\u306f\u300c\u80cc\u666f\u8272\u4ed8\u304d\u306e\u8868\u306e\u30bb\u30eb\u300d\u3067\u4f5c\u308b\uff08\u3069\u3061\u3089\u306eOutlook\u3067\u3082\u540c\u3058\u898b\u305f\u76ee\u306b\u306a\u308b\uff09\u3002 */

function buildApprovalBlock(o) {
  var approveLink = buildMailto(o.requesterEmail,
    ANSWER_PREFIX + "\u627f\u8a8d\uff5c" + o.subject + "\uff08" + o.receiptId + "\uff09",
    "\u627f\u8a8d\u3057\u307e\u3059\u3002\r\n\r\n\u30b3\u30e1\u30f3\u30c8\uff1a\r\n");

  var rejectLink = buildMailto(o.requesterEmail,
    ANSWER_PREFIX + "\u5374\u4e0b\uff5c" + o.subject + "\uff08" + o.receiptId + "\uff09",
    "\u5374\u4e0b\u3057\u307e\u3059\u3002\r\n\r\n\u7406\u7531\uff1a\r\n");

  var font = "font-family:'Yu Gothic UI','Meiryo UI',Meiryo,sans-serif;";

  return (
    '<table role="presentation" cellpadding="0" cellspacing="0" border="0" width="560" ' +
      'style="width:560px;max-width:100%;border:1px solid #9cc7b3;border-left:6px solid #166e4a;' +
      'background:#f3faf6;' + font + '">' +
      '<tr><td bgcolor="#f3faf6" style="padding:14px 18px 18px 18px;color:#1f2a24;font-size:14px;' + font + '">' +
        '<p style="margin:0 0 10px 0;font-size:15px;font-weight:bold;' + font + '">\u627f\u8a8d\u4f9d\u983c</p>' +
        '<table role="presentation" cellpadding="0" cellspacing="0" border="0"><tr>' +
          button(approveLink, "\u627f\u8a8d", "#166e4a", "#ffffff", "#166e4a", font) +
          '<td width="10" style="width:10px;font-size:1px;">&nbsp;</td>' +
          button(rejectLink, "\u5374\u4e0b", "#ffffff", "#a1262b", "#a1262b", font) +
        '</tr></table>' +
      '</td></tr>' +
    '</table>' +
    '<p style="margin:0;">&nbsp;</p>'
  );
}

function button(href, label, bg, fg, borderColor, font) {
  return (
    '<td bgcolor="' + bg + '" align="center" ' +
      'style="background:' + bg + ';border:1px solid ' + borderColor + ';padding:8px 24px;' + font + '">' +
      '<a href="' + href + '" style="color:' + fg + ';text-decoration:none;font-weight:bold;' +
      'font-size:14px;display:inline-block;' + font + '">' + label + '</a>' +
    '</td>'
  );
}

/* ---------- \u5b9a\u578b\u6587 ---------- */

function buildMessageText(o) {
  var names = [];
  for (var i = 0; i < o.approvers.length; i++) {
    var r = o.approvers[i];
    names.push(familyName(r.displayName || r.emailAddress) + "\u69d8");
  }
  var sender = familyName(o.senderDisplayName) || "\u6e21\u908a";

  var lines = [
    escapeHtml(names.join("\u3001")),
    "",
    escapeHtml(sender) + "\u3067\u3059\u3002",
    "",
    "\u6a19\u8a18\u306e\u4ef6\u3001\u95a2\u9023\u90e8\u7f72\u3067\u306eECN\u5224\u5b9a\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f\u3002",
    "\u627f\u8a8d\u3044\u305f\u3060\u3051\u308b\u969b\u306f\u3001\u6295\u7968\u30dc\u30bf\u30f3\u306e\u300c\u627f\u8a8d\u300d\u3092\u304a\u9858\u3044\u3057\u307e\u3059\u3002",
    "",
    "\u3000\u30ea\u30f3\u30af\uff1a",
    "",
    "\u4ee5\u4e0a\u3067\u3059\u3002",
    ""
  ];
  var out = "";
  for (var j = 0; j < lines.length; j++) {
    out += "<div>" + (lines[j] === "" ? "<br>" : lines[j]) + "</div>";
  }
  return out;
}

// \u300cJun Kitagawa(\u5317\u5ddd\u3000\u6df3)\u300d\u2192\u300c\u5317\u5ddd\u300d\u3001\u62ec\u5f27\u304c\u306a\u3051\u308c\u3070\u8868\u793a\u540d\u3092\u305d\u306e\u307e\u307e\u4f7f\u3046
function familyName(displayName) {
  var s = trim(String(displayName || ""));
  var m = s.match(/[\uff08(]([^\uff09)]+)[\uff09)]/);
  if (m) {
    return trim(m[1]).split(/[\s\u3000]+/)[0];
  }
  return s;
}

function buildMailto(to, subject, body) {
  return "mailto:" + to +
    "?subject=" + encodeURIComponent(subject) +
    "&body=" + encodeURIComponent(body);
}

/* ---------- \u4ef6\u540d ---------- */

// \u300cRe:\u3010ECN\u5224\u5b9a\u3011xxx\u300d\u2192\u300cxxx\u300d
// \u3010ECN\u5224\u5b9a\u3011\u304c\u7121\u3044\u5834\u5408\u306f\u3001\u5148\u982d\u306e Re:/Fw: \u3084\u65e2\u5b58\u306e\u627f\u8a8d\u4f9d\u983c\u306e\u5370\u3060\u3051\u3092\u53d6\u308a\u9664\u304f
function extractBaseSubject(s) {
  var idx = s.indexOf(SOURCE_MARKER);
  var rest;
  if (idx !== -1) {
    rest = s.slice(idx + SOURCE_MARKER.length);
  } else {
    rest = s
      .replace(/^\s*((re|fw|fwd|\u8fd4\u4fe1|\u8ee2\u9001)\s*[:\uff1a]\s*)+/i, "")
      .replace(/^\u3010ECN\u627f\u8a8d\u4f9d\u983c\u3011|^\u3010\u627f\u8a8d\u4f9d\u983c\u3011/, "");
  }
  return trim(rest);
}

/* ---------- \u30e6\u30fc\u30c6\u30a3\u30ea\u30c6\u30a3 ---------- */

function makeReceiptId() {
  var d = new Date();
  return "AP-" + d.getFullYear() + pad2(d.getMonth() + 1) + pad2(d.getDate()) +
    "-" + pad2(d.getHours()) + pad2(d.getMinutes()) + pad2(d.getSeconds());
}

function pad2(n) {
  return (n < 10 ? "0" : "") + n;
}

function trim(s) {
  return String(s).replace(/^[\s\u3000]+|[\s\u3000]+$/g, "");
}

function escapeHtml(s) {
  return String(s || "")
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function errorText(e) {
  if (!e) return "\u4e0d\u660e\u306a\u30a8\u30e9\u30fc";
  return e.message || String(e);
}

function notify(item, kind, message, callback) {
  var details;
  if (kind === "error") {
    details = { type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage, message: message };
  } else {
    details = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: message,
      icon: "Icon.16",
      persistent: false
    };
  }
  try {
    item.notificationMessages.replaceAsync("approvalAddin", details, function () {
      if (callback) callback();
    });
  } catch (e) {
    if (callback) callback();
  }
}
