function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('학생 관리')
    .addItem('로그인/기록 열기', 'showStudentPanel')
    .addToUi();
}

function showStudentPanel() {
  const html = HtmlService.createHtmlOutputFromString(getStudentPanelHtml())
    .setWidth(640)
    .setHeight(760)
    .setTitle('학생 로그인 / 학습 기록');

  SpreadsheetApp.getUi().showModalDialog(html, '학생 로그인 / 학습 기록');
}

function getStudentPanelHtml() {
  return `
    <!DOCTYPE html>
    <html>
      <head>
        <base target="_top">
        <meta charset="utf-8">
        <style>
          body { font-family: Arial, sans-serif; padding: 16px; }
          .box { border: 1px solid #ddd; border-radius: 8px; padding: 14px; margin-bottom: 14px; }
          label { display: block; margin-top: 8px; font-weight: bold; }
          input, select, textarea { width: 100%; box-sizing: border-box; padding: 8px; margin-top: 4px; }
          button { margin-top: 12px; padding: 8px 12px; }
          .message { margin-top: 10px; padding: 8px; border-radius: 6px; }
          .success { background: #e8f5e9; color: #1b5e20; }
          .error { background: #ffebee; color: #c62828; }
          .hidden { display: none; }
        </style>
      </head>
      <body>
        <h3>학생 로그인</h3>
        <div class="box">
          <label>이름</label>
          <input id="studentName" placeholder="학생 이름">
          <label>비밀번호 (6자리 숫자)</label>
          <input id="studentPassword" type="password" inputmode="numeric" maxlength="6" placeholder="123456">
          <button onclick="loginStudent()">로그인</button>
          <button onclick="logoutStudent()" type="button">다른 학생으로 로그인</button>
          <div id="loginMessage" class="message"></div>
        </div>

        <div id="recordBox" class="box hidden">
          <h3>학습 기록 저장</h3>
          <label>과목</label>
          <input id="subject" value="국어">
          <label>단원</label>
          <input id="unit" placeholder="예: 1단원">
          <label>모드</label>
          <select id="mode">
            <option value="기본학습">기본학습</option>
            <option value="인출학습">인출학습</option>
            <option value="타임어택">타임어택</option>
          </select>
          <label>정답수</label>
          <input id="correctCount" type="number" value="0">
          <label>전체수</label>
          <input id="totalCount" type="number" value="0">
          <label>소요시간(초)</label>
          <input id="timeSec" type="number" value="0">
          <label>오답단어목록</label>
          <textarea id="wrongWords" rows="3" placeholder="예: 꿈, 민족"></textarea>
          <label>타임어택 점수</label>
          <input id="timeAttackScore" type="number" value="0">
          <label>누적학습시간(초)</label>
          <input id="cumulativeTime" type="number" value="0">
          <button onclick="saveRecord()">기록 저장</button>
          <div id="recordMessage" class="message"></div>
        </div>

        <script>
          const nameInput = document.getElementById('studentName');
          const passwordInput = document.getElementById('studentPassword');
          const recordBox = document.getElementById('recordBox');
          const loginMessage = document.getElementById('loginMessage');
          const recordMessage = document.getElementById('recordMessage');

          function showMessage(el, message, isError) {
            if (!el) return;
            el.className = 'message ' + (isError ? 'error' : 'success');
            el.textContent = message;
          }

          function restoreSession() {
            const savedName = localStorage.getItem('studentName') || '';
            if (savedName) {
              nameInput.value = savedName;
              nameInput.disabled = true;
              recordBox.classList.remove('hidden');
            }
          }

          function logoutStudent() {
            localStorage.removeItem('studentName');
            nameInput.value = '';
            nameInput.disabled = false;
            passwordInput.value = '';
            recordBox.classList.add('hidden');
            showMessage(loginMessage, '다른 학생으로 다시 로그인할 수 있습니다.', false);
            showMessage(recordMessage, '', false);
          }

          restoreSession();

          function loginStudent() {
            const name = (nameInput.value || '').trim();
            const password = (passwordInput.value || '').trim();

            if (!name) {
              showMessage(loginMessage, '이름을 입력해주세요.', true);
              return;
            }
            if (!/^\d{6}$/.test(password)) {
              showMessage(loginMessage, '비밀번호는 6자리 숫자로 입력해주세요.', true);
              return;
            }

            google.script.run.withSuccessHandler(function(result) {
              if (result.success) {
                localStorage.setItem('studentName', result.studentName);
                nameInput.value = result.studentName;
                nameInput.disabled = true;
                recordBox.classList.remove('hidden');
                const extra = result.created ? ' 새 학생으로 등록되었습니다.' : '';
                showMessage(loginMessage, '로그인 성공.' + extra + ' 기록을 저장할 수 있습니다.', false);
              } else {
                showMessage(loginMessage, result.message, true);
              }
            }).validateLogin(name, password);
          }

          function saveRecord() {
            const name = (nameInput.value || '').trim();
            const password = (passwordInput.value || '').trim();
            const payload = {
              name: name,
              password: password,
              subject: document.getElementById('subject').value,
              unit: document.getElementById('unit').value,
              mode: document.getElementById('mode').value,
              correctCount: Number(document.getElementById('correctCount').value || 0),
              totalCount: Number(document.getElementById('totalCount').value || 0),
              timeSec: Number(document.getElementById('timeSec').value || 0),
              wrongWords: document.getElementById('wrongWords').value,
              timeAttackScore: Number(document.getElementById('timeAttackScore').value || 0),
              cumulativeTime: Number(document.getElementById('cumulativeTime').value || 0)
            };

            google.script.run.withSuccessHandler(function(result) {
              if (result.success) {
                showMessage(recordMessage, result.message, false);
                document.getElementById('wrongWords').value = '';
                document.getElementById('correctCount').value = '0';
                document.getElementById('totalCount').value = '0';
                document.getElementById('timeSec').value = '0';
                document.getElementById('timeAttackScore').value = '0';
                document.getElementById('cumulativeTime').value = '0';
              } else {
                showMessage(recordMessage, result.message, true);
              }
            }).saveRecord(payload);
          }
        </script>
      </body>
    </html>
  `;
}

function ensureStudentSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Sheet1');
  if (!sheet) {
    sheet = ss.insertSheet('Sheet1');
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 2).setValues([['이름', '비밀번호']]);
  } else if (sheet.getRange(1, 1, 1, 2).getDisplayValues()[0][0] !== '이름' || sheet.getRange(1, 1, 1, 2).getDisplayValues()[0][1] !== '비밀번호') {
    sheet.getRange(1, 1, 1, 2).setValues([['이름', '비밀번호']]);
  }

  return sheet;
}

function ensureRecordSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Sheet2');
  if (!sheet) {
    sheet = ss.insertSheet('Sheet2');
  }

  const headers = [
    '이름', '과목', '단원', '모드', '정답수',
    '전체수', '정답률', '소요시간(초)', '오답단어목록', '완료시각',
    '타임어택점수', '누적학습시간(초)'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else if (sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].join('') !== headers.join('')) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sheet;
}

function validateLogin(studentName, studentPassword) {
  const sheet = ensureStudentSheet();
  const name = String(studentName || '').trim();
  const password = String(studentPassword || '').trim();

  if (!name) {
    return { success: false, message: '이름을 입력해주세요.' };
  }

  if (!/^\d{6}$/.test(password)) {
    return { success: false, message: '비밀번호는 6자리 숫자로 입력해주세요.' };
  }

  const data = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 0), 2).getDisplayValues();

  for (let i = 0; i < data.length; i++) {
    const storedName = String(data[i][0] || '').trim();
    const storedPassword = String(data[i][1] || '').trim();

    if (storedName.toLowerCase() === name.toLowerCase()) {
      if (storedPassword === password) {
        return { success: true, studentName: storedName, created: false };
      }
      return { success: false, message: '비밀번호가 일치하지 않습니다.' };
    }
  }

  sheet.appendRow([name, password]);
  return { success: true, studentName: name, created: true };
}

function saveRecord(payload) {
  const loginResult = validateLogin(payload && payload.name, payload && payload.password);
  if (!loginResult.success) {
    return loginResult;
  }

  const sheet = ensureRecordSheet();
  const correctCount = Number(payload && payload.correctCount ? payload.correctCount : 0);
  const totalCount = Number(payload && payload.totalCount ? payload.totalCount : 0);
  const accuracy = totalCount > 0 ? (correctCount / totalCount * 100).toFixed(1) + '%' : '0.0%';
  const completedAt = new Date().toLocaleString('ko-KR');

  const row = [
    loginResult.studentName,
    payload && payload.subject ? payload.subject : '',
    payload && payload.unit ? payload.unit : '',
    payload && payload.mode ? payload.mode : '',
    correctCount,
    totalCount,
    accuracy,
    payload && payload.timeSec ? payload.timeSec : 0,
    payload && payload.wrongWords ? payload.wrongWords : '',
    completedAt,
    payload && payload.timeAttackScore ? payload.timeAttackScore : 0,
    payload && payload.cumulativeTime ? payload.cumulativeTime : 0
  ];

  sheet.appendRow(row);
  return { success: true, message: '기록이 Sheet2에 저장되었습니다.' };
}
