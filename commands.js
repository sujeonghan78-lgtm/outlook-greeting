// 메일 인사말 자동 삽입
// 새 메일 / 회신 / 전달 열 때 본문 맨 위에 인사말을 넣어줌
// 문구 우선순위: ① 본인이 "인사말 설정"에서 저장한 문구 → ② config.json → ③ 아래 비상용
// (여기는 수정할 필요 없음)

// config.json 을 못 불러왔을 때 쓰는 비상용 문구
var FALLBACK_LINES = [
  "수신 :",
  "발신 : {이름}",
  "",
  "연일 업무로 노고 많으십니다.",
  "",
  "감사합니다."
];

function escapeHtml(s) {
  return String(s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function linesToHtml(lines, name) {
  var html = "";
  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i]).replace(/\{이름\}/g, name);
    html += line === "" ? "<div><br></div>" : "<div>" + escapeHtml(line) + "</div>";
  }
  return html + "<div><br></div>";
}

// 아웃룩 표시 이름으로 config 에서 그 사람 문구 찾기
function pickLines(config, name) {
  if (config[name]) return config[name];
  for (var key in config) {
    if (key !== "기본" && name.indexOf(key) !== -1) return config[key];
  }
  return config["기본"] || FALLBACK_LINES;
}

function insertGreeting(event) {
  var name = "";
  try { name = Office.context.mailbox.userProfile.displayName || ""; } catch (e) {}

  function done(lines) {
    Office.context.mailbox.item.body.prependAsync(
      linesToHtml(lines, name),
      { coercionType: Office.CoercionType.Html },
      function () {
        if (event && event.completed) event.completed();
      }
    );
  }

  // 1순위: 본인이 "인사말 설정"에서 저장한 문구
  var saved = null;
  try { saved = Office.context.roamingSettings.get("greeting"); } catch (e) {}
  if (saved) {
    done(String(saved).split("\n"));
    return;
  }

  // 2순위: config.json (관리자가 관리하는 문구)
  fetch("config.json?t=" + Date.now(), { cache: "no-store" })
    .then(function (res) { return res.json(); })
    .then(function (config) { done(pickLines(config, name)); })
    .catch(function () { done(FALLBACK_LINES); });
}

// 새 메일/회신/전달 열릴 때 자동 실행
function onNewMessageComposeHandler(event) {
  insertGreeting(event);
}

Office.onReady(function () {});

// 아웃룩에 함수 연결
if (Office.actions && Office.actions.associate) {
  Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
  Office.actions.associate("insertGreeting", insertGreeting);
}
