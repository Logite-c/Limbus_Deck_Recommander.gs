// ==========================================
// last update 2026.06.11
// version 1.0
// Made by 이산코사인변환됨(로지트) with Gemini
// ==========================================

/**
 * 스프레드시트가 열릴 때 자동으로 실행되는 기본 내장 함수입니다.
 * 사용자가 스크립트를 쉽게 실행할 수 있도록 상단 메뉴바에 [💾 데이터 내보내기/불러오기] 커스텀 메뉴를 추가합니다.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("💾 데이터 내보내기/불러오기")
    .addItem("인격/E.G.O 내보내기", "exportIdentitiesUI")
    .addItem("인격/E.G.O 불러오기", "importIdentitiesUI")
    .addSeparator()
    .addItem("파티 데이터 내보내기", "exportPartyUI")
    .addItem("파티 데이터 불러오기", "importPartyUI")
    .addToUi();
}

// ==========================================
// 1. 인격 / E.G.O 데이터 처리
// ==========================================

function getIdentitiesAndEgos() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("보유 인격 체크");
  if (!sheet) throw new Error("'보유 인격 체크' 시트를 찾을 수 없습니다.");

  // 데이터 업데이트로 인해 행이 늘어나도 유연하게 대응할 수 있도록,
  // 시트의 마지막 데이터 행(getLastRow)까지 전체 범위를 동적으로 로드합니다.
  const range = sheet.getRange("B1:S" + Math.max(10, sheet.getLastRow()));
  const data = range.getValues();

  // 시트 레이아웃이 4개의 주요 구간으로 나뉘어 있으므로, 각 구간의 시작점을 찾기 위한 변수입니다.
  let rowIdx_Id_1 = -1; // 인격 1부 시작 (첫 번째 "이상" 등장 위치)
  let rowIdx_Id_2 = -1; // 인격 2부 시작 (첫 번째 "히스클리프" 등장 위치)
  let rowIdx_Ego_1 = -1; // E.G.O 1부 시작 (두 번째 "이상" 등장 위치)
  let rowIdx_Ego_2 = -1; // E.G.O 2부 시작 (두 번째 "히스클리프" 등장 위치)

  // B열(수감자 이름 열)을 위에서 아래로 순회하며 위 4가지 기준점의 행 인덱스(0-based)를 탐색합니다.
  for (let r = 0; r < data.length; r++) {
    const val = String(data[r][0]).trim(); // B열의 값
    if (val === "이상") {
      if (rowIdx_Id_1 === -1) {
        rowIdx_Id_1 = r;
      } else if (rowIdx_Ego_1 === -1) {
        rowIdx_Ego_1 = r;
      }
    } else if (val === "히스클리프") {
      if (rowIdx_Id_2 === -1) {
        rowIdx_Id_2 = r;
      } else if (rowIdx_Ego_2 === -1) {
        rowIdx_Ego_2 = r;
      }
    }
  }

  if (
    rowIdx_Id_1 === -1 ||
    rowIdx_Id_2 === -1 ||
    rowIdx_Ego_1 === -1 ||
    rowIdx_Ego_2 === -1
  ) {
    throw new Error(
      "시트의 '이상' 또는 '히스클리프' 행을 찾을 수 없습니다. 시트 레이아웃을 확인하세요.",
    );
  }

  // 한 행에는 총 6명의 수감자 데이터가 가로로 나열되어 있습니다.
  // 각각의 수감자 데이터를 추출하기 위해 [이름 열, 항목명 열, 체크박스 열]의 위치(인덱스)를 지정해둔 배열입니다.
  const groupCols = [
    { charColIdx: 0, nameColIdx: 1, cbColIdx: 2 }, // B, C, D 열
    { charColIdx: 3, nameColIdx: 4, cbColIdx: 5 }, // E, F, G 열
    { charColIdx: 6, nameColIdx: 7, cbColIdx: 8 }, // H, I, J 열
    { charColIdx: 9, nameColIdx: 10, cbColIdx: 11 }, // K, L, M 열
    { charColIdx: 12, nameColIdx: 13, cbColIdx: 14 }, // N, O, P 열
    { charColIdx: 15, nameColIdx: 16, cbColIdx: 17 }, // Q, R, S 열
  ];

  const exportData = [];

  /**
   * 지정된 행 범위(startRow ~ endRow)를 스캔하여 체크박스가 켜진(TRUE) 항목만 수집하는 내부 함수입니다.
   * @param {number} headerRow - 수감자 이름이 적혀있는 헤더 행 인덱스
   */
  function scanRange(startRow, endRow, headerRow) {
    for (let r = startRow; r <= endRow; r++) {
      if (r >= data.length) break;
      for (let g = 0; g < groupCols.length; g++) {
        const group = groupCols[g];
        const charName = String(data[headerRow][group.charColIdx]).trim();
        const itemName = String(data[r][group.nameColIdx]).trim();
        const cbValue = data[r][group.cbColIdx];

        // 항목 이름이 비어있거나 구분을 위한 더미 라벨(예: "3성", "WAW" 등)인 경우 추출 대상에서 제외합니다.
        if (
          itemName !== "" &&
          ![
            "1성",
            "2성",
            "3성",
            "ZAYIN",
            "TETH",
            "HE",
            "WAW",
            "ALEPH",
          ].includes(itemName)
        ) {
          if (cbValue === true || String(cbValue).toUpperCase() === "TRUE") {
            exportData.push({ character: charName, name: itemName });
          }
        }
      }
    }
  }

  // 위에서 선언한 scanRange 함수를 활용하여 4구간의 데이터를 순차적으로 추출합니다.
  scanRange(rowIdx_Id_1 + 1, rowIdx_Id_2 - 1, rowIdx_Id_1);

  // 2. 인격 2부 (히스클리프 ~ E.G.O 이상 직전)
  scanRange(rowIdx_Id_2 + 1, rowIdx_Ego_1 - 1, rowIdx_Id_2);

  // 3. E.G.O 1부 (이상 ~ 히스클리프 직전)
  scanRange(rowIdx_Ego_1 + 1, rowIdx_Ego_2 - 1, rowIdx_Ego_1);

  // 4. E.G.O 2부 (히스클리프부터 시작해서 표의 끝인 빈 행이 나올 때까지 탐색)
  let r = rowIdx_Ego_2 + 1;
  while (r < data.length) {
    // B열부터 S열까지 완전히 빈 행인지 확인
    let isRowEmpty = true;
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim() !== "") {
        isRowEmpty = false;
        break;
      }
    }
    if (isRowEmpty) break; // 모든 데이터가 없는 빈 행을 만나면 E.G.O 2부 스캔 종료

    for (let g = 0; g < groupCols.length; g++) {
      const group = groupCols[g];
      const charName = String(data[rowIdx_Ego_2][group.charColIdx]).trim();
      const itemName = String(data[r][group.nameColIdx]).trim();
      const cbValue = data[r][group.cbColIdx];

      if (
        itemName !== "" &&
        !["ZAYIN", "TETH", "HE", "WAW", "ALEPH"].includes(itemName)
      ) {
        if (cbValue === true || String(cbValue).toUpperCase() === "TRUE") {
          exportData.push({ character: charName, name: itemName });
        }
      }
    }
    r++;
  }

  return exportData;
}

/**
 * 추출한 인격/E.G.O 데이터를 JSON 문자열로 변환하여
 * 사용자가 복사/다운로드할 수 있게 다이얼로그(팝업창)를 띄웁니다.
 */
function exportIdentitiesUI() {
  try {
    const exportData = getIdentitiesAndEgos();
    showExportHtml(JSON.stringify(exportData, null, 2), "인격/E.G.O 내보내기");
  } catch (err) {
    SpreadsheetApp.getUi().alert("오류가 발생했습니다: " + err.message);
  }
}

/**
 * 사용자가 텍스트 창에 붙여넣은 JSON 문자열을 분석하여,
 * 시트의 체크박스 상태를 저장 당시의 상태로 덮어쓰기(불러오기) 하는 함수입니다.
 */
function importIdentitiesData(jsonString) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("보유 인격 체크");
  if (!sheet)
    return SpreadsheetApp.getUi().alert(
      "'보유 인격 체크' 시트를 찾을 수 없습니다.",
    );

  const range = sheet.getRange("B1:S" + Math.max(10, sheet.getLastRow()));
  const data = range.getValues();

  const importedList = JSON.parse(jsonString);

  // 무결성 검사: 파티 데이터를 인격 시트에 잘못 불러오는 경우를 방지
  if (
    !Array.isArray(importedList) ||
    (importedList.length > 0 && importedList[0].character === undefined)
  ) {
    return SpreadsheetApp.getUi().alert(
      "❌ 올바른 데이터가 아닙니다. '인격/E.G.O' 저장 데이터가 맞는지 확인해 주세요.",
    );
  }

  // 데이터 처리 속도를 높이기 위해, 불러온 목록을 Set 형태(예: "이상|개화 E.G.O::동백")로 캐싱해둡니다.
  // 이후 시트를 순회하며 이 Set에 포함되어 있는지만 빠르게 확인하여 체크박스 상태를 전환합니다.
  const ownedSet = new Set(
    importedList.map((item) => item.character + "|" + item.name),
  );

  let rowIdx_Id_1 = -1;
  let rowIdx_Id_2 = -1;
  let rowIdx_Ego_1 = -1;
  let rowIdx_Ego_2 = -1;

  // 불러오기를 수행하기 위해 각 데이터의 기준 행 위치를 다시 찾습니다.
  for (let r = 0; r < data.length; r++) {
    const val = String(data[r][0]).trim();
    if (val === "이상") {
      if (rowIdx_Id_1 === -1) rowIdx_Id_1 = r;
      else if (rowIdx_Ego_1 === -1) rowIdx_Ego_1 = r;
    } else if (val === "히스클리프") {
      if (rowIdx_Id_2 === -1) rowIdx_Id_2 = r;
      else if (rowIdx_Ego_2 === -1) rowIdx_Ego_2 = r;
    }
  }

  const groupCols = [
    { charColIdx: 0, nameColIdx: 1, cbColIdx: 2 },
    { charColIdx: 3, nameColIdx: 4, cbColIdx: 5 },
    { charColIdx: 6, nameColIdx: 7, cbColIdx: 8 },
    { charColIdx: 9, nameColIdx: 10, cbColIdx: 11 },
    { charColIdx: 12, nameColIdx: 13, cbColIdx: 14 },
    { charColIdx: 15, nameColIdx: 16, cbColIdx: 17 },
  ];

  /**
   * 지정된 범위 내의 시트 데이터를 스캔하면서,
   * ownedSet에 존재하는 항목이면 체크박스를 TRUE로, 없으면 FALSE로 업데이트합니다.
   */
  function applyImport(startRow, endRow, headerRow) {
    for (let r = startRow; r <= endRow; r++) {
      if (r >= data.length) break;
      for (let g = 0; g < groupCols.length; g++) {
        const group = groupCols[g];
        const charName = String(data[headerRow][group.charColIdx]).trim();
        const itemName = String(data[r][group.nameColIdx]).trim();

        if (
          itemName !== "" &&
          ![
            "1성",
            "2성",
            "3성",
            "ZAYIN",
            "TETH",
            "HE",
            "WAW",
            "ALEPH",
          ].includes(itemName)
        ) {
          const key = charName + "|" + itemName;
          data[r][group.cbColIdx] = ownedSet.has(key); // 가져온 목록에 있으면 true, 없으면 false
        }
      }
    }
  }

  // 데이터 덮어쓰기 적용
  applyImport(rowIdx_Id_1 + 1, rowIdx_Id_2 - 1, rowIdx_Id_1);
  applyImport(rowIdx_Id_2 + 1, rowIdx_Ego_1 - 1, rowIdx_Id_2);
  applyImport(rowIdx_Ego_1 + 1, rowIdx_Ego_2 - 1, rowIdx_Ego_1);

  // 4구간(E.G.O 2부) 불러오기 진행: 데이터베이스 끝부분(완전 빈 행)까지 탐색
  let r = rowIdx_Ego_2 + 1;
  while (r < data.length) {
    let isRowEmpty = true;
    for (let c = 0; c < data[r].length; c++) {
      if (String(data[r][c]).trim() !== "") {
        isRowEmpty = false;
        break;
      }
    }
    if (isRowEmpty) break;

    for (let g = 0; g < groupCols.length; g++) {
      const group = groupCols[g];
      const charName = String(data[rowIdx_Ego_2][group.charColIdx]).trim();
      const itemName = String(data[r][group.nameColIdx]).trim();

      if (
        itemName !== "" &&
        !["ZAYIN", "TETH", "HE", "WAW", "ALEPH"].includes(itemName)
      ) {
        const key = charName + "|" + itemName;
        data[r][group.cbColIdx] = ownedSet.has(key);
      }
    }
    r++;
  }

  // 메모리 상에서 변경된 배열 값들을 실제 스프레드시트 셀에 단 한 번의 명령으로 일괄 저장합니다 (성능 최적화)
  range.setValues(data);
  SpreadsheetApp.getUi().alert("인격/E.G.O 불러오기가 완료되었습니다.");
}

// ==========================================
// 2. 파티 편집기 데이터 처리 (4행부터, 2행 6열 묶음)
// ==========================================

/**
 * '파티 편집기' 시트에 저장된 파티 구성 데이터를 추출합니다.
 * 각 파티는 2개 행(상단/하단)과 6개 열(슬롯)로 구성되어 있습니다.
 */
function exportPartyUI() {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("파티 편집기");
  if (!sheet)
    return SpreadsheetApp.getUi().alert(
      "'파티 편집기' 시트를 찾을 수 없습니다.",
    );

  const data = sheet
    .getRange("B4:H" + Math.max(4, sheet.getLastRow()))
    .getValues();
  const exportData = [];

  for (let r = 0; r < data.length - 1; r += 2) {
    const rowLabel = String(data[r][0]).trim();

    if (rowLabel.includes("파티")) {
      const row1Slots = data[r].slice(1, 7); // C ~ H (윗행 6개 슬롯)
      const row2Slots = data[r + 1].slice(1, 7); // C ~ H (아랫행 6개 슬롯)

      exportData.push({
        partyLabel: rowLabel,
        row1: row1Slots,
        row2: row2Slots,
      });
    }
  }

  showExportHtml(JSON.stringify(exportData, null, 2), "파티 데이터 내보내기");
}

/**
 * 붙여넣은 파티 JSON 데이터를 바탕으로 파티 편집기 시트의 구성을 불러옵니다.
 * 라벨 이름("파티1", "파티2" 등)을 비교하며 위치를 매칭합니다.
 */
function importPartyData(jsonString) {
  const sheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("파티 편집기");
  const range = sheet.getRange("B4:H" + Math.max(4, sheet.getLastRow()));
  const data = range.getValues();

  const importedData = JSON.parse(jsonString);

  // 무결성 검사: 인격 데이터를 파티 시트에 잘못 붙여넣는 경우를 방지
  if (
    !Array.isArray(importedData) ||
    (importedData.length > 0 && importedData[0].partyLabel === undefined)
  ) {
    return SpreadsheetApp.getUi().alert(
      "❌ 올바른 데이터가 아닙니다. '파티' 저장 데이터가 맞는지 확인해 주세요.",
    );
  }

  let importIndex = 0;

  for (let r = 0; r < data.length - 1; r += 2) {
    const rowLabel = String(data[r][0]).trim();

    if (rowLabel.includes("파티") && importIndex < importedData.length) {
      // JSON 저장 데이터와 현재 시트의 파티 라벨이 일치할 때 불러오기를 진행합니다.
      if (rowLabel === importedData[importIndex].partyLabel) {
        for (let c = 0; c < 6; c++) {
          data[r][c + 1] = importedData[importIndex].row1[c]; // 윗행 불러오기
          data[r + 1][c + 1] = importedData[importIndex].row2[c]; // 아랫행 불러오기
        }
        importIndex++;
      }
    }
  }
  range.setValues(data);
  SpreadsheetApp.getUi().alert("파티 데이터 불러오기가 완료되었습니다.");
}

// ==========================================
// 3. HTML UI 다이얼로그 생성 함수
// ==========================================

/**
 * 데이터를 내보낼 때(Export) 화면에 보여줄 HTML 팝업 UI를 생성합니다.
 * 보안을 위해 특수기호를 이스케이프하고, 복사하기/다운로드 자바스크립트가 포함되어 있습니다.
 */
function showExportHtml(jsonString, title) {
  // HTML 태그 충돌(특히 </textarea>)을 방지하기 위한 안전한 이스케이프 처리
  const safeJsonString = jsonString
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

  const html = `
    <html>
      <head>
        <style>
          body { font-family: sans-serif; padding: 10px; }
          textarea { width: 100%; height: 200px; margin-bottom: 15px; resize: none; font-family: monospace; }
          button { padding: 8px 16px; margin-right: 10px; cursor: pointer; border: 1px solid #ccc; background: #f8f9fa; border-radius: 4px; }
          button:hover { background: #e2e6ea; }
        </style>
      </head>
      <body>
        <textarea id="data" readonly>${safeJsonString}</textarea>
        <div>
          <button onclick="copyData()">📋 복사하기</button>
          <button onclick="downloadData()">💾 다운로드 (.json)</button>
        </div>
        <script>
          function copyData() {
            var copyText = document.getElementById("data");
            copyText.select();
            document.execCommand("copy");
            alert("클립보드에 복사되었습니다!");
          }
          function downloadData() {
            var data = document.getElementById("data").value;
            var blob = new Blob([data], {type: "application/json"});
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = "${title === "인격/E.G.O 내보내기" ? "limbus_identity_save.json" : "limbus_party_save.json"}";
            a.click();
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(320);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}

/**
 * 인격/E.G.O 메뉴에서 불러오기를 클릭했을 때 실행
 */
function importIdentitiesUI() {
  showImportHtml("identity");
}
/**
 * 파티 데이터 메뉴에서 불러오기를 클릭했을 때 실행
 */
function importPartyUI() {
  showImportHtml("party");
}

/**
 * 외부 데이터를 붙여넣기(Import) 할 수 있는 텍스트 상자가 포함된 HTML UI를 띄웁니다.
 */
function showImportHtml(type) {
  const title =
    type === "identity" ? "인격/E.G.O 불러오기" : "파티 데이터 불러오기";
  const funcName =
    type === "identity" ? "importIdentitiesData" : "importPartyData";

  const html = `
    <html>
      <head>
        <style>
          body { font-family: sans-serif; padding: 10px; }
          textarea { width: 100%; height: 200px; margin-bottom: 15px; resize: none; font-family: monospace; }
          button { padding: 8px 16px; cursor: pointer; border: 1px solid #ccc; background: #4285f4; color: white; border-radius: 4px; border: none;}
          button:hover { background: #3367d6; }
        </style>
      </head>
      <body>
        <p style="margin-top:0; font-size: 14px; color: #555;">저장한 JSON 데이터를 아래에 붙여넣고 [불러오기]를 누르세요.</p>
        <textarea id="data" placeholder="[ { ... } ]"></textarea>
        <div>
          <button onclick="submitData()">📥 불러오기</button>
        </div>
        <script>
          function submitData() {
            var data = document.getElementById("data").value.trim();
            if (!data) return alert("데이터를 입력해주세요.");
            
            // 불러오기 전 괄호 누락이나 데이터 형식 오류를 사전에 검사
            try {
              JSON.parse(data);
            } catch (e) {
              return alert("❌ 데이터 형식이 올바르지 않습니다.\\n복사 과정에서 누락되거나 깨진 부분이 없는지 확인하세요.");
            }

            var btn = document.querySelector("button");
            btn.innerText = "처리 중...";
            btn.disabled = true;

            google.script.run
              .withSuccessHandler(function() {
                google.script.host.close();
              })
              .withFailureHandler(function(err) {
                alert("오류가 발생했습니다: " + err.message);
                btn.innerText = "📥 불러오기";
                btn.disabled = false;
              })
              .${funcName}(data);
          }
        </script>
      </body>
    </html>
  `;
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(450)
    .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, title);
}
