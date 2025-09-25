function getDataSetting() {
  var data = {}
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.setting);
    // cài đặt tọa độ, IP
    let rows = s.getRange(5, 1, s.getLastRow() - 6, s.getMaxColumns() - 1).getDisplayValues();
    for (let [k, v] of Object.entries(listTitle)) {
      let indexRow = rows.findIndex(row => {
        let idx = row.indexOf(v);
        if (idx === -1) return false
        return true
      });

      if (indexRow === -1) continue
      let indexValue = rows[indexRow].findIndex(x => x !== v && x !== '');
      if (indexValue === -1) data[k] = null
      else data[k] = rows[indexRow][indexValue];
    }
    let setting = data;

    // danh sách sheet vote 
    let listVote = s.getRange(17, 2, s.getLastRow() - 6, s.getMaxColumns() - 1).getDisplayValues();
    let keys = {};
    let headers = null;
    headers = listVote.shift();
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
    }
    data = { setting, keys, listVote }
  } catch (error) {
    console.log(`Error: ${error}`)
  } finally {
    return data
  }
}

function getUrlTop() {
  let dataSetting = getDataSetting();
  return dataSetting.setting.full_url;
}

function getDataFromSheet(sheetName) {
  var data = {};
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    let rows = [];
    let keys = null;
    let headers = null;
    var regex = /【[^】]+】(.*)/;
    var match = regex.exec(sheetName);
    sheetName = match[1];
    switch (sheetName) {
      case listSheet.questions:
        // hàng 5 cột 2 lấy 4 hàng và số cột = max - 1
        let rowSetting = s.getRange(5, 2, 4, s.getLastColumn() - 1).getDisplayValues();
        let surveySetting = {};

        rows = s.getRange(9, 2, s.getLastRow() - 6, s.getLastColumn() - 1).getDisplayValues();
        headers = rows.shift();
        keys = {};

        for (const [k, v] of Object.entries(listTitle)) {
          keys[k] = headers.indexOf(v);

          // get survey setting
          for (let i = 0; i < rowSetting.length; i++) {
            let row = rowSetting[i];
            let idx = row.findIndex(x => x === v);
            if (idx > -1) {
              surveySetting[k] = row[idx + 3];
              break;
            }
          }
        }

        // filter rows
        rows = rows.filter(row => row[keys.question] !== '')

        data = { keys, rows, surveySetting }
        break;
      default:
        rows = s.getRange(5, 2, s.getLastRow() - 6, s.getLastColumn() - 1).getDisplayValues();
        headers = rows.shift();
        keys = {};

        for (const [k, v] of Object.entries(listTitle)) {
          keys[k] = headers.indexOf(v);
        }

        data = { keys, rows }
        break;
    }
  } catch (error) {
    console.log(`Error: ${error}`)
  } finally {
    return data
  }
}

function getDataFromSheet2(sheetName, listTitle) {
  var data = {};
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!s) throw 'Sheet not found'

    let rows = [];
    let keys = null;
    let headers = null;
    switch (true) {
      default:
        console.log("getDataFromSheet run default")
        rows = s.getRange(5, 2, s.getLastRow() - 4, s.getLastColumn() - 1).getDisplayValues();
        headers = rows.shift();
        keys = {};

        for (const [k, v] of Object.entries(listTitle)) {
          keys[k] = headers.indexOf(v);
        }
        break;
    }
    data = { keys, rows }
  } catch (error) {
    console.log(`Error: ${error}`)
  } finally {
    return data
  }
}

function getDataSettingByKey(key) {
  return headerSetting.getDataByKey(key)
}

function sendOTP(email) {
  var res = {
    status: true,
    data: null,
    msg: null
  }
  try {
    let check = false;
    let data = getDataFromSheet2(listSheet.verifyManager, headerVerifyManager);
    if (data) {
      let managers = data.rows
      //Kiểm tra xem email có trong danh sách không
      for (let i = 0; i < managers.length; i++) {
        if (email == managers[i][data.keys.email] && managers[i][data.keys.status] == headerVerifyManager.active) {
          check = true;
          let user = managers[i];
          let otpCode = Math.floor(1000 + Math.random() * 9000);
          let urlTop = getDataSettingByKey('gas_url');
          // Lấy ra giá trị lifetime OTP
          let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.setting);
          let OTPlifetime = getDataSettingByKey('OTPlifetime');
          let timeFrom = new Date();
          let timeTo = new Date(timeFrom.getTime() + OTPlifetime * 24 * 60 * 60 * 1000);
          // Xóa dữ liệu OTP cũ
          s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.loginOTP);
          let otpDataSheet = getDataFromSheet2(listSheet.loginOTP, headerLoginOTP);
          if (otpDataSheet) {
            let optData = otpDataSheet.rows;
            for (let y = 0; y < optData.length; y++) {
              if (email == optData[y][otpDataSheet.keys.email]) {
                console.log(Number(optData[y][otpDataSheet.keys.no]) + 5);
                s.deleteRow(Number(optData[y][otpDataSheet.keys.no]) + 5);
              }
            }
          }
          // Ghi otpCode vào sheet
          let newRow = [
            '=ROW() - 5',
            user[data.keys.fullname],
            email,
            otpCode,
            formatJapaneseDate(timeFrom),
            formatJapaneseDate(timeTo)
          ];
          s.insertRowBefore(6);
          s.getRange(6, 2, 1, s.getMaxColumns() - 2).setBorder(false, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID)
          s.getRange(6, 2, 1, newRow.length).setValues([newRow]);
          // Gửi OTP tới email
          let pageObject = { page: "manage", email: email, otp: otpCode };
          let encodedPageParam = encodeURIComponent(JSON.stringify(pageObject));
          let url = `https://myportal.sateraito.jp/gas?url=${urlTop}?page=${encodedPageParam}`;
          var subject = "「" + otpCode + "」" + "投票管理のログインコード";
          let htmlMessage = `<html>
                                <head>
                                    <meta charset="UTF-8">
                                    <title>Email OTP</title>
                                    <style>
                                        body {
                                            font-family: Arial, sans-serif;
                                            margin: 0;
                                            padding: 0;
                                            background-color: #f4f4f4;
                                        }
                                        .container {
                                            width: 100%;
                                            max-width: 600px;
                                            margin: 0 auto;
                                            padding: 20px;
                                            background-color: #ffffff;
                                            box-shadow: 0 0 10px rgba(0, 0, 0, 0.1);
                                        }
                                        .header {
                                            background-color: #4CAF50;
                                            color: white;
                                            padding: 10px;
                                            text-align: center;
                                        }
                                        .content {
                                
                                            text-align: center;
                                        }
                                        .otp {
                                            font-size: 24px;
                                            font-weight: bold;
                                            margin: 20px 0;
                                            color: #4CAF50;
                                        }
                                        .footer {
                                            text-align: center;
                                            font-size: 12px;
                                        }
                                    </style>
                                </head>
                                <body>
                                    <div class="container">
                                        <div class="header">
                                            <h1>投票管理画面へアクセスのログインコード</h1>
                                        </div>
                                        <div class="content">
                                            <p>ログイン画面に下のログインコードを入力してください。</p>
                                            <div class="otp">${otpCode}</div>
                                            <p>または、<a href="${url}">こちらURL</a>をクリックしてアクセスしてください。</p>
                                            <p>このログインコード有効期限はコードは ${OTPlifetime} 日間なります。</p>
                                        </div>
                                        <div class="footer">
                                            <p>よろしくお願いします</p>
                                        </div>
                                    </div>
                                </body>
                                </html>`
          MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: htmlMessage,
            name: "サテライトオフィスGAS開発"
          });
          break;
        }
      }
    }
    if (!check) {
      res.status = false;
    }
    return res;
  } catch (error) {
    console.log(`Error: ${error}`)
    res.status = false
    res.msg = error
  } finally {
    return res
  }
}

function submitVerify(email, otp) {
  var res = {
    status: true,
    data: null,
    msg: null
  }
  try {
    let check = false;
    let data = getDataFromSheet2(listSheet.loginOTP, headerLoginOTP);
    if (data) {
      let opts = data.rows;
      for (let i = 0; i < opts.length; i++) {
        if (email == opts[i][data.keys.email] && otp == opts[i][data.keys.otp]) {
          let fullname = opts[i][data.keys.fullname];
          res.data = (fullname) ? fullname : '';
          let timeFrom = new Date(opts[i][data.keys.fromTime].replace(/(\d{4})年(\d{2})月(\d{2})日 (\d{2}):(\d{2}):(\d{2})/, '$1-$2-$3T$4:$5:$6'));
          let timeTo = new Date(opts[i][data.keys.toTime].replace(/(\d{4})年(\d{2})月(\d{2})日 (\d{2}):(\d{2}):(\d{2})/, '$1-$2-$3T$4:$5:$6'));
          if (new Date() > timeFrom && new Date() <= timeTo) {
            check = true;
          }
          break;
        }
      }
    }
    if (!check) {
      res.status = false
    }
    return res;
  } catch (error) {
    console.log(`Error: ${error}`)
    res.status = false
    res.msg = error
  } finally {
    return res
  }
}

function formatJapaneseDate(date) {
  let year = date.getFullYear();
  let month = String(date.getMonth() + 1).padStart(2, '0');
  let day = String(date.getDate()).padStart(2, '0');
  let hours = String(date.getHours()).padStart(2, '0');
  let minutes = String(date.getMinutes()).padStart(2, '0');
  let seconds = String(date.getSeconds()).padStart(2, '0');
  return `${year}年${month}月${day}日 ${hours}:${minutes}:${seconds}`;
}

function getQuestions(surveyID) {
  var res = {
    status: true,
    data: null,
    colors: {},
    msg: null
  }
  try {
    var vote = getVoteById(surveyID);
    var combinedQuestion = `【${surveyID}】${listSheet.questions}`;
    var { keys, rows, surveySetting } = getDataFromSheet(combinedQuestion);
    console.log('vvv');
    console.log(combinedQuestion);
    surveySetting['display_type'] = display_type.getDisplayType(surveySetting['display_type']);
    surveySetting['titles'] = getVoteListTitle();
    

    for (let row of rows) {
      row[keys.question_type] = question_type.getQuestionType(row[keys.question_type]);
      if (row[keys.question_hasNote] !== undefined) {
        row[keys.question_hasNote] = row[keys.question_hasNote] === 'あり';
      }
      row[keys.question_required] = (row[keys.question_required] == 'はい') ? true : false;
      row[keys.question_answers] = _split(row[keys.question_answers]);
      row[keys.criterias] = row[keys.criterias] === '' ? [] : _split(row[keys.criterias]);
      row[keys.colors] = row[keys.colors] === '' ? [] : _split(row[keys.colors]);
      // row[keys.voteOrder] = row[keys.voteOrder] === '' ? [] : _split(row[keys.voteOrder]);
      // row[keys.max] = row[keys.max] === '' ? [] : _split(row[keys.max]);
      // row[keys.min] = row[keys.min] === '' ? [] : _split(row[keys.min]);
      row[keys.voteMethod] = row[keys.voteMethod] || '単純多数決';
      row[keys.voteThreshold] = row[keys.voteThreshold] || '50.1%';
    }

    // Lấy sheet theo tên
    let sheetColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.colors);
    // Kiểm tra nếu sheet tồn tại
    if (sheetColors) {
        // Lấy dữ liệu từ cột C (từ C6 trở đi) và cột D (từ D6 trở đi)
        let columnC = sheetColors.getRange('C6:C').getValues();
        // Lấy màu nền của các ô trong cột D (từ D6 trở đi)
        let backgroundColorsD = sheetColors.getRange('D6:D').getBackgrounds();
        // Khởi tạo đối tượng colors trong res.data
        res.colors = {};

        // Lặp qua dữ liệu từ cột C và D
        for (let i = 0; i < columnC.length; i++) {
            let key = columnC[i][0];  // Lấy giá trị từ cột C
            let value = backgroundColorsD[i][0];

            // Kiểm tra nếu key hoặc value không trống trước khi thêm vào res.data.colors
            if (key && value) {
                res.colors[key] = value; // Gán giá trị vào res.data.colors
            }
        }
    }
    console.log('qqq');
    console.log(rows);
    console.log(keys);

    // Responsive
    res.data = { keys, rows, surveySetting, vote };
    res.msg = "Success";

  } catch (error) {
    console.log("getQuestions", { error })
    res.status = false
    res.msg = error.name
  } finally {
    return res
  }
}

function getSheetIndex(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName() === sheetName) {
      return i + 1; // Trả về vị trí của sheet, bắt đầu từ 1
    }
  }
  return -1; // Trả về -1 nếu không tìm thấy sheet
}


function newSheetResponse(nameSheet, surveyID) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  try {
    // Get the source sheet
    var sourceSheet = spreadsheet.getSheetByName(listSheet.responseForm);

    // Create a new sheet with the same name
    var nameSheetQuestion = `【${surveyID}】${listSheet.questions}`;
    var indexSheetQuestion = getSheetIndex(nameSheetQuestion);
    var newSheet = spreadsheet.insertSheet(nameSheet, indexSheetQuestion - 1, { template: sourceSheet });
    // set Title sheet
    newSheet.getRange(2, 2).setValue(listSheet.responseTotal);

    // set header
    var { keys, rows, surveySetting } = getDataFromSheet(nameSheetQuestion);
    var amount_question = 0;
    let colStart = 5;
    for (let i = 0; i < rows.length; i++) {
      let row = rows[i];
      // insert column question
      let colName = `${i + 1}．${row[keys.question]}`;
      if (i !== rows.length - 1) newSheet.insertColumnAfter(colStart + amount_question);  // insert new column
      amount_question += 1;
      newSheet.getRange(5, colStart + amount_question).setValue(colName);  // set value column question
      newSheet.setColumnWidth(colStart + amount_question, 200);  // set width column
      newSheet.getRange(6, colStart + amount_question, newSheet.getMaxRows() - 5, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      newSheet.getRange(6, colStart + amount_question, newSheet.getMaxRows() - 5, 1).setHorizontalAlignment('left');
      newSheet.getRange(5, colStart + amount_question).setHorizontalAlignment('left');

      // insert a column at right column question
      let hasNote = row[keys.question_hasNote] === 'あり' ? true : false;
      if (hasNote) {
        newSheet.insertColumnAfter(colStart + amount_question)
        amount_question += 1;
        newSheet.setColumnWidth(colStart + amount_question, 200);  // set width column

        newSheet.getRange(5, colStart + amount_question - 1, 1, 2).merge();
      }
    }

    // set link at menu to newsheet

    //let sMenu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.menu);
    //let { row, col, value } = getLinkSheetResponse();
    //sMenu.getRange(row, col).setValue(`=HYPERLINK("#gid=${newSheet.getSheetId()}&range=A2", "${listTitle.linkToResponse}")`)
    //sMenu.getRange(row, col).setFontColor("#03a62f")
    //sMenu.getRange(row, col).setFontLine('none')

    return newSheet
  } catch (error) {
    console.log({ error })
    spreadsheet.deleteSheet(newSheet);
    return null
  }
}

function saveResponse(form, surveyID) {
  console.log('saveResponse');
  console.log(form);
  var res = {
    status: true,
    msg: ''
  }
  // Gets a script lock before modifying a shared resource.
  var lock = LockService.getScriptLock();
  var vote = getVoteById(surveyID);
  // Waits for up to 30 seconds for other processes to finish.
  lock.waitLock(30000);
  try {
    // get latest sheet response
    // let latestSheet = getLatestResponseSheet();
    let latestSheet = checkSheetAnswer(surveyID);
    if (!latestSheet) throw "シートが見つかりません。"

    let rows = latestSheet.getRange(5, 2, latestSheet.getLastRow() - 4, latestSheet.getLastColumn() - 1).getDisplayValues();
    let headers = rows.shift();

    // Tạo mảng cột và đặt lệnh số thứ tự cho cột đầu tiên
    let newValue = new Array(latestSheet.getLastColumn() - 1).fill('')
    newValue[0] = '=ROW()-5';
    for (const [k, v] of Object.entries(form)) {
      if (['answer_created', 'address', 'name', 'phone_number', 'email'].includes(k)) {
        newValue[headers.indexOf(listTitle[k])] = v.value
        continue
      }

      let idx = headers.indexOf(k);
      if (idx === -1) continue

      if (v.criteria == true) {
        newValue[idx] = '';
        if (v.type === 'multiSelect') {
          for (const key in v.criterias) {
            let valuesString = '';
            for (const index in v.criterias[key].value) {
              if (v.criterias[key].value[index] !== '') {
                valuesString += `[${v.criterias[key].value[index]}]`;
              }
            }
            if (valuesString) {
              newValue[idx] += `- ${key}: ${valuesString}\n`;
            }
            // newValue[idx] += `- ${key}: `;
            // for (const index in v.criterias[key].value) {
            //   newValue[idx] += `[${v.criterias[key].value[index]}]`;
            // }
            // newValue[idx] += `\n`;
          }
        } else {
          for (const key in v.criterias) {
            if (v.criterias[key]['value'] !== '') {
              newValue[idx] += `- ${key}: ${v.criterias[key]['value']}\n`;
            }
          }
          // Lặp qua các entries của participants và thay thế với dữ liệu mới
          // for (const [idxp, v] of Object.entries(form.participants)) {
          //   for (const key in v) {
          //     if (v[key] !== '') {
          //       newValue[idx] += `- ${key}: ${v[key]}\n`;
          //     }
          //   }
          // }
        }
      } else {
        if (v.value == '') {
          newValue[idx] = ''
        } else {
          newValue[idx] = v.type === 'multiSelect' ? `- ${v.value.join('\n- ')}` : v.value;
        }
      }
      if (v.note) newValue[idx + 1] = v.note
    }
    // check user answered
    console.log(form);
    let idx = rows.findIndex(row => {
      let status = false;
      for (const [k, v] of Object.entries(form)) {
        if (!['answer_created', 'name', 'phoneNumber', 'email'].includes(k)) continue
        console.log(row[headers.indexOf(listTitle['email'])]);
        if (k == 'email') {
          if (row[headers.indexOf(listTitle['email'])].includes(v.value)) status = true;
        }
      }
      return status
    })
    if(form.email.value == '') idx = -1;
    console.log({ idx });
    if (idx === -1) {
      // create new response
      latestSheet.insertRowBefore(6);
      latestSheet.getRange(6, 2, 1, latestSheet.getMaxColumns() - 2).setBorder(false, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID)
      latestSheet.getRange(6, 2, 1, latestSheet.getMaxColumns() - 2).setValues([newValue]);
    } else {
      // edit exist response
      latestSheet.getRange(idx + 6, 2, 1, latestSheet.getMaxColumns() - 2).setValues([newValue]);
    }
    // Cộng thêm voted
    // Tìm trong sheet master và thay đổi
    let nameListResponse = `【${surveyID}】${listSheet.responseTotal}`;
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListResponse);
    let numberRows = (sheet) ? sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues().length - 1 : 0;
    let numberVoted = (numberRows == 0) ? '0' : numberRows;
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    let keys = {};
    rows = sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues();
    headers = rows.shift();
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
    }
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][keys.id] == surveyID) {
        sheet.getRange("O" + (i + 6)).setValue(numberVoted);
      }
    }
    res.msg = "成功しました。"
  } catch (error) {
    console.log({ error })
    res.status = false;
    res.msg = error;
  } finally {
    lock.releaseLock();
    res.data = { vote }
    return res
  }
}

function getLatestResponseSheet() {
  try {
    let { row, col, value } = getLinkSheetResponse();
    if (!row || !col) throw "not found cell"
    const regex = /#gid=([^&]+)/;
    const match = regex.exec(value);
    let sId = null

    if (match) sId = match[1];
    else throw "No match found";

    let sheet = null;
    try {
      sheet = getSheetById(sId);
    } catch (error) {
      sheet = newSheetResponse(listSheet.responseTotal);
    } finally {
      return sheet
    }
  } catch (error) {
    console.log({ error })
    return null
  }
}

function checkSheetAnswer(surveyID) {
  let latestSheet = null;
  let nameSheetAnswer = `【${surveyID}】${listSheet.responseTotal}`;
  try {
    latestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetAnswer);
    if (latestSheet) {
      return latestSheet
    } else {
      latestSheet = newSheetResponse(nameSheetAnswer, surveyID);
    }
  } catch (error) {
    console.log({ error, nameSheetAnswer })
  }
  return latestSheet;
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function (s) { return String(s.getSheetId()) === id; }
  )[0];
}

function urlencode(data) {
  // Create a string in the form key1=value1&key2=value2&...
  const queryString = Object.keys(data)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(data[key])}`)
    .join('&');

  // Create a blob from the query string
  const blob = Utilities.newBlob(queryString, 'application/x-www-form-urlencoded');

  // Get the data as a string (URL-encoded)
  const encodedString = blob.getDataAsString();

  return encodedString;
}

function createShortURL(url = null) {
  if (!url) return ''
  let post_data = {
    'api_key': 'b2d669730e864d1f9422f2d926674c9e',
    'url': url,
  }

  let options = {
    'method': 'post',
    'payload': urlencode(post_data),
    'deadline': 30,
    'follow_redirects': true
  }
  let result = null
  let retry_cnt = 0
  while (retry_cnt < 3) {
    try {
      result = UrlFetchApp.fetch(url = 'https://shurl.jp/api/shortenurl/create', options)
      retry_cnt = 3
    } catch (error) {
      retry_cnt += 1
      console.log({ error })
    }
  }

  if (result.getResponseCode() !== 200) return '';
  let response = JSON.parse(result.getContentText("UTF-8"));
  if (response.code !== 0) return '';
  return response.shorten_url;
}

function saveShortLink() {
  console.log("run save short-link")
  var data = getDataSetting(listSheet.settings)
  let { setting, keys, listVote } = data;
  console.log({ data })
  let url = data.setting.full_url;
  if (url === '') return false
  // Mã hóa object thành chuỗi JSON và sau đó mã hóa URI
  let pageObject = { page: "manage" };
  let encodedPageParam = encodeURIComponent(JSON.stringify(pageObject));
  url = `${url}?page=${encodedPageParam}`;
  let shortLink = createShortURL(`https://myportal.sateraito.jp/gas?url=${url}`)
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.setting)
  ss.getRange("D7").setValue(shortLink)

  // for (let i = 0; i < listVote.length; i++) {
  //   for (const [k, v] of Object.entries(keys)) {
  //     if (k == 'id') {
  //       id = listVote[i][v];
  //     }
  //     if (k == 'url' && id) {
  //       url = `${url}?page=form&surveyID=${id}`;
  //       let bitly = createShortURL(url)
  //       let shortLink = createShortURL(`?url=${bitly}`)
  //       var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.setting)
  //       let indexRow = 17 + parseInt(id);
  //       // ss.getRange("G" + indexRow.toString()).setValue(shortLink)
  //     }
  //   }
  //   url = data.setting.full_url;
  // }

}

function isAnswered(form) {
  var res = false

  try {
    let user_code = form.user_code.value;
    let name = form.name.value;
    if (user_code === '' || name === '') throw false

    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.responseTotal)
    let rows = s.getRange(5, 2, s.getLastRow() - 4, s.getLastColumn() - 1).getDisplayValues();
    let headers = rows.shift();
    console.log({ user_code, name })

    let idx = rows.findIndex(row => row[headers.indexOf(listTitle.user_code)] == user_code && row[headers.indexOf(listTitle.name)] == name);
    console.log({ idx })
    if (idx === -1) throw false
    res = true

  } catch (error) {
    console.log({ error })
    res = false
  } finally {
    return res
  }
}

function isAnswered2(surveyID) {
  var res = false
  try {
    let sheetName = `【${surveyID}】回答一覧`;
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (s) {
      let value = s.getRange('C6').getDisplayValues();
      if (value != '') {
        res = true;
      }
    }
    console.log(res);
  } catch (error) {
    console.log({ error })
    res = true
  } finally {
    return res
  }
}

function getLinkSheetResponse() {
  var data = {};
  try {
    var s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.menu);
    var rows = s.getRange(5, 1, s.getLastRow() - 4, s.getMaxColumns() - 1).getDisplayValues();

    for (let i = 0; i < rows.length; i++) {
      let row = rows[i];
      let idx = row.findIndex(x => x === listTitle.linkToResponse);
      if (idx === -1) continue
      data['row'] = i + 5;
      data['col'] = idx + 1;
      data['value'] = s.getRange(i + 5, idx + 1).getFormula()
    }
  } catch (error) {
    console.log({ error })
  } finally {
    return data
  }
}

function genLinkVote(id) {
  var data = getDataSetting(listSheet.settings)
  let url = data.setting.full_url;
  if (url === '') return false
  let pageObject = {
    page: "form",
    surveyID: id
  };
  let encodedPageParam = encodeURIComponent(JSON.stringify(pageObject));
  url = `${url}?page=${encodedPageParam}`;
  let shortLink = createShortURL(`https://myportal.sateraito.jp/gas?url=${url}`)
  return shortLink;
}

function genLinkVoteQRCode(linkVote) {
  var data = getDataSetting(listSheet.settings)
  let url = data.setting.full_url;
  if (url === '') return false
  url = `${url}?page=qrCode&linkVote=${linkVote}`;
  let bitly = createShortURL(url)
  return createShortURL(`${bitly}`);
}

function getListVote() {
  var res = {
    status: true,
    data: null,
    msg: null
  }
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    rows = s.getRange(5, 2, s.getLastRow() - 4, s.getMaxColumns() - 2).getDisplayValues();
    let headers = rows.shift();
    let keys = {};
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
    }

    // keys.numberVoted = master_sheet.getLength();
    // for (let i = 0; i < rows.length; i++) {
    //   let nameListResponse = `【${rows[i][0]}】${listSheet.responseTotal}`;
    //   let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListResponse);
    //   let numberRows = (sheet) ? sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues().length - 1 : 0;
    //   rows[i].push(numberRows);
    // }

    // keys.numberQuestion = master_sheet.getLength() + 1;
    // keys.description = master_sheet.getLength() + 2;
    // for (let i = 0; i < rows.length; i++) {
    //   let nameListQuestions = `【${rows[i][0]}】${listSheet.questions}`;
    //   let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListQuestions);
    //   let numberRows = (sheet) ? sheet.getRange(11, 2, sheet.getLastRow() - 10, sheet.getMaxColumns() - 2).getDisplayValues().length - 1 : 0;
    //   let description = sheet.getRange('E7').getValue();
    //   rows[i].push(numberRows);
    //   rows[i].push(description);
    // }

    res.data = { keys, rows };
  } catch (error) {
    console.log({ error })
    res.status = false
    res.msg = error
  } finally {
    return res
  }
}

// get vote id
function getVoteById(surveyID) {
  let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
  let rows = s.getRange(5, 2, s.getLastRow() - 4, s.getMaxColumns() - 2).getDisplayValues();
  let headers = rows.shift();
  let keys = {};
  for (const [k, v] of Object.entries(master_sheet)) {
    keys[k] = headers.indexOf(v);
  }
  for (let i = 0; i < rows.length; i++) {
    let rowData = rows[i];
    if (rowData[keys.id] === surveyID) {
      let arrayKey = headers;
      const voteObj = arrayKey.reduce((acc, currentKey, index) => {
        acc[currentKey] = rowData[index];
        return acc;
      }, {});

      // map value with master_sheet
      const voteObjMaster = Object.keys(master_sheet).reduce((acc, key) => {
        const dataKey = master_sheet[key];
        acc[key] = voteObj[dataKey];
        return acc;
      }, {});
      return voteObjMaster;
    }
  }
}

// get vote title
function getVoteListTitle() {
  var res = {
    status: true,
    data: null,
    msg: null
  }
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    rows = s.getRange(5, 2, s.getLastRow() - 4, s.getMaxColumns() - 2).getDisplayValues();
    let headers = rows.shift();
    let keys = {};
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
    }

    return headers;
  } catch (error) {
    console.log(error);
  }
  return null;
}

function createVote(form) {
  var res = {
    status: true,
    msg: '',
    data: {},
  };
  try {
    form.statistics = (form.statistics) ? '表示' : '非表示';
    form.usePassCode = (form.usePassCode) ? 'はい' : 'いいえ';
    form.informationRequired = (form.informationRequired) ? 'あり' : 'なし';
    // Lấy trang tính để thực hiện thêm mới vote
    let id = Math.random().toString(36).substring(2, 6);
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    rows = sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues();
    let headers = rows.shift();
    let keys = {};
    let newValue = [];
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
      if (form[k]) {
        newValue[keys[k]] = form[k];
      }
    }
    // fixed value id, url, passcode
    let shortLinkLinkVote = genLinkVote(id);
    newValue[0] = id;
    newValue[master_sheet.getIndex('url')] = shortLinkLinkVote;
    newValue[master_sheet.getIndex('passcode')] = Math.floor(Math.random() * 90000) + 10000;
    // Get position insert data
    sheet.insertRowBefore(6);
    sheet.getRange(6, 2, 1, sheet.getMaxColumns() - 2).setBorder(false, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 2, 1, sheet.getMaxColumns() - 2).setValues([newValue]);
    // Tạo sheet câu hỏi mới
    var nameSheet = "【" + id + "】" + listSheet.questions;
    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.templateVoteSetting);
    var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(nameSheet, 5, { template: sourceSheet });
    newSheet.getRange("E5").setValue(form.nameVote);
    newSheet.getRange("E7").setValue(form.description);
    createQuestionDataVote(form, nameSheet);
    // Tạo sheet câu trả lời
    let sheetResponse = checkSheetAnswer(id);
    // Gán hyperlink cho list
    var newSheetQuestionsUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + newSheet.getSheetId();
    sheet.getRange("L6").setFormula('=HYPERLINK("' + newSheetQuestionsUrl + '", "投票設定")');
    var newSheetResponseUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sheetResponse.getSheetId();
    sheet.getRange("M6").setFormula('=HYPERLINK("' + newSheetResponseUrl + '", "回答一覧")');
    form.id = id;
    form.url = newValue[master_sheet.getIndex('url')];
    form.passcode = newValue[master_sheet.getIndex('passcode')];
    res.data = form
    res.msg = "新しい投票が正常に作成されました。";
    // res.data = { newValue, data };
  } catch (error) {
    res.status = false;
    res.msg = error;
    console.log(error);
  }
  return res;
}

function createQuestionDataVote(form, sheetName) {
  try {
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    for (let i = form.questions.length - 1; i >= 0; i--) {
      // Xử lý tiêu chí
      let criteria = '';
      for (let y = 0; y < form.questions[i].criterias.length; y++) {
        criteria += form.questions[i].criterias[y].title;
        if (form.questions[i].criterias[y].media != '') {
          criteria += '<' + form.questions[i].criterias[y].media + '>'
        }
        criteria += '\n'
      }
      criteria = criteria.endsWith('\n') ? criteria.slice(0, -1) : criteria;
      // Xử lý lựa chọn
      let answer = '';
      for (let y = 0; y < form.questions[i].answers.length; y++) {
        answer += form.questions[i].answers[y].title;
        if (form.questions[i].answers[y].media != '') {
          answer += '<' + form.questions[i].answers[y].media + '>'
        }
        answer += '\n'
      }
      answer = answer.endsWith('\n') ? answer.slice(0, -1) : answer;
      
      let color = '';
      for (let y = 0; y < form.questions[i].answers.length; y++) {
        color += form.questions[i].answers[y].color;
        color += '\n'
      }
      color = color.endsWith('\n') ? color.slice(0, -1) : color;

      let newValue = [];
      newValue[0] = '=ROW() - 9';
      newValue[1] = form.questions[i].question;
      newValue[2] = criteria;
      newValue[3] = form.questions[i].typeQuestion;
      newValue[4] = (form.questions[i].questionRequire == true) ? 'はい' : 'いいえ';
      newValue[5] = answer;
      newValue[6] = color;
      newValue[7] = form.questions[i].voteOrder;
      newValue[8] = form.questions[i].max;
      newValue[9] = form.questions[i].min;
      newValue[10] = (form.questions[i].typeSubject != 'なし') ? form.questions[i].subject : '';
      newValue[11] = form.questions[i].voteMethod || '単純多数決';
      newValue[12] = form.questions[i].voteThreshold || '50.1%';
      s.insertRowBefore(10);
      s.getRange(10, 2, 1, s.getLastColumn() - 1).setBorder(false, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID)
      s.getRange(10, 2, 1, s.getMaxColumns() - 2).setValues([newValue])
    }
  } catch (error) {
    console.log(error);
  }
}

function makeAcopy(vote) {
  var res = {
    status: true,
    msg: '',
    data: null
  }
  try {
    // chỉnh sửa các giá trị
    let surveyIDOld = vote.id;
    vote.id = Math.random().toString(36).substring(2, 6);
    vote.nameVote = vote.nameVote + "のコピー";
    vote.status = "作成中";
    let nameSheetSetting = "【" + surveyIDOld + "】投票設定";
    // copy sheet setting
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetSetting)
    if (s) {
      let copiedSheetSetting = s.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      copiedSheetSetting.setName("【" + vote.id + "】投票設定");
      vote.voteSetting = `=HYPERLINK("#gid=${copiedSheetSetting.getSheetId()}!A2", "投票設定")`
    } else {
      res.status = false;
      res.msg = 'ワークシート ' + nameSheetSetting + 'が見つかりません';
      return res;
    }
    // copy sheet respponse
    let nameSheetRespponse = "【" + surveyIDOld + "】回答一覧";
    s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetRespponse)
    if (s) {
      let copiedSheetResponse = s.copyTo(SpreadsheetApp.getActiveSpreadsheet());
      copiedSheetResponse.setName("【" + vote.id + "】回答一覧");
      // Xóa dữ liệu trả lời cũ
      let row6Values = copiedSheetResponse.getRange(6, 2, 1, copiedSheetResponse.getLastColumn() - 1).getValues();
      let hasDataInRow6 = row6Values[0].some(cell => cell !== "");
      if (hasDataInRow6) {
        copiedSheetResponse.getRange(6, 2, copiedSheetResponse.getLastRow() - 5, copiedSheetResponse.getLastColumn()).clearContent();
      }
      vote.voteResponse = `=HYPERLINK("#gid=${copiedSheetResponse.getSheetId()}!A2", "回答一覧")`
    } else {
      res.status = false;
      res.msg = 'ワークシート ' + nameSheetSetting + 'が見つかりません';
      return res;
    }

    // tạo passcode, shortLink
    vote.passcode = Math.floor(Math.random() * 90000) + 10000;
    let url = headerSetting.getDataByKey('gas_url');
    if (url) {
      let pageObject = {
        page: "form",
        surveyID: vote.id
      };
      let encodedPageParam = encodeURIComponent(JSON.stringify(pageObject));
      url = `${url}?page=${encodedPageParam}`;
      let shortLink = createShortURL(`https://myportal.sateraito.jp/gas?url=${url}`)
      vote.url = shortLink;
    }

    // Lấy master sheet và lưu thông tin
    vote.numberVoted = "0";
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    rows = sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues();
    let headers = rows.shift();
    let keys = {};
    let newValue = [];
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
      if (vote[k]) {
        newValue[keys[k]] = vote[k];
      }
    }
    sheet.insertRowBefore(6);
    sheet.getRange(6, 2, 1, sheet.getMaxColumns() - 2).setBorder(false, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
    sheet.getRange(6, 2, 1, sheet.getMaxColumns() - 2).setValues([newValue]);
    res.msg = "成功しました。"
    res.data = vote;
  } catch (error) {
    console.log(`registervote error: ${error}`);
    res.status = false;
    res.msg = error
  } finally {
    return res
  }
}

function getDataQuestionsVote(surveyID) {
  var res = {
    status: true,
    msg: '',
    data: null
  }
  try {
    let data = getQuestions(surveyID);
    let { keys, rows, surveySetting, vote } = data.data;
    let form = vote;
    form.description = surveySetting.survey_description;
    form.questions = [];
    for (let i = 0; i < rows.length; i++) {
      let row = {};
      row.question = rows[i][keys.question];
      if (rows[i][keys.question_type] == "multiSelect") {
        row.typeQuestion = '複数選択';
      } else if (rows[i][keys.question_type] == "oneSelect") {
        row.typeQuestion = '1 つだけ選択';
      } else if (rows[i][keys.question_type] == "score") {
        row.typeQuestion = '点数';
      } else if (rows[i][keys.question_type] == "input") {
        row.typeQuestion = '入力';
      } else if (rows[i][keys.question_type] == "target") {
        row.typeQuestion = '対象者選択';
      }
      row.questionRequire = rows[i][keys.question_required];
      row.subject = rows[i][keys.description_for_answer];
      row.voteOrder = rows[i][keys.voteOrder];
      row.max = rows[i][keys.max];
      row.min = rows[i][keys.min];
      row.voteMethod = rows[i][keys.voteMethod] || '単純多数決';
      row.voteThreshold = rows[i][keys.voteThreshold] || '50.1%';
      if (row.subject != '') {
        row.typeSubject = '補足テキスト'
      } else {
        row.typeSubject = 'なし'
      }
      row.answers = [];
      for (let y = 0; y < rows[i][keys.question_answers].length; y++) {
        if (rows[i][keys.question_answers][y] != '') {
          let answer = {};
          let parts = rows[i][keys.question_answers][y].split('<');
          answer.title = parts[0];
          answer.media = parts[1] ? parts[1].slice(0, -1) : '';
          console.log('bbb');
          console.log(rows[i][keys.colors][y]);
          if (rows[i][keys.colors][y] != '' && rows[i][keys.colors][y] != undefined) {
            let parts_color = rows[i][keys.colors][y].split('<');
            // answer.color = 'bg-info';
            answer.color = parts_color[0];
          } else {
            answer.color = 'bg-info';
          }
          row.answers.push(answer);
        }
      }
      row.criterias = [];
      for (let y = 0; y < rows[i][keys.criterias].length; y++) {
        if (rows[i][keys.criterias][y] != '') {
          let criteria = {};
          let parts = rows[i][keys.criterias][y].split('<');
          criteria.title = parts[0];
          criteria.media = parts[1] ? parts[1].slice(0, -1) : '';
          row.criterias.push(criteria);
        }
      }
      // row.colors = [];
      // console.log('ccc');
      // console.log(rows);
      // for (let y = 0; y < rows[i][keys.colors].length; y++) {
      //   if (rows[i][keys.colors][y] != '') {
      //     let color = {};
      //     let parts = rows[i][keys.colors][y].split('<');
      //     answer.color = parts[0];
      //     row.answers.push(color);
      //   }
      // }
      form.questions.push(row)
    }
    let nameListResponse = `【${surveyID}】${listSheet.responseTotal}`;
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameListResponse);
    let numberRows = (sheet) ? sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues().length - 1 : 0;
    form.numberVoted = (numberRows == 0) ? '0' : numberRows;
    res.data = form;
    // Lấy sheet theo tên
    let sheetColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.colors);
    // Kiểm tra nếu sheet tồn tại
    if (sheetColors) {
        // Lấy dữ liệu từ cột C (từ C6 trở đi) và cột D (từ D6 trở đi)
        let columnC = sheetColors.getRange('C6:C').getValues();
        // Lấy màu nền của các ô trong cột D (từ D6 trở đi)
        let backgroundColorsD = sheetColors.getRange('D6:D').getBackgrounds();
        // Khởi tạo đối tượng colors trong res.data
        res.data.colors = {};

        // Lặp qua dữ liệu từ cột C và D
        for (let i = 0; i < columnC.length; i++) {
            let key = columnC[i][0];  // Lấy giá trị từ cột C
            let value = backgroundColorsD[i][0];

            // Kiểm tra nếu key hoặc value không trống trước khi thêm vào res.data.colors
            if (key && value) {
                res.data.colors[key] = value; // Gán giá trị vào res.data.colors
            }
        }
    }
    console.log('eee');
    console.log(res.data.questions);
    console.log(keys.question);
  } catch (error) {
    console.log(`getQuizBycode error: ${error}`)
    res.status = false;
    res.msg = error;
  } finally {
    return res
  }
}

function getDataColors(){
  console.log('aaa');
  var res = {
    status: true,
    data: {
      colors: {}  // Khởi tạo res.data.colors như một đối tượng rỗng
    },
    msg: null
  };
  
  // Lấy sheet theo tên
  let sheetColors = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.colors);
  
  // Kiểm tra nếu sheet tồn tại
  if (sheetColors) {
      // Lấy dữ liệu từ cột C (từ C6 trở đi) và cột D (từ D6 trở đi)
      let columnC = sheetColors.getRange('C6:C').getValues();
      // Lấy màu nền của các ô trong cột D (từ D6 trở đi)
      let backgroundColorsD = sheetColors.getRange('D6:D').getBackgrounds();

      // Lặp qua dữ liệu từ cột C và D
      for (let i = 0; i < columnC.length; i++) {
          let key = columnC[i][0];  // Lấy giá trị từ cột C
          let value = backgroundColorsD[i][0];

          // Kiểm tra nếu key hoặc value không trống trước khi thêm vào res.data.colors
          if (key && value) {
              res.data.colors[key] = value; // Gán giá trị vào res.data.colors
          }
      }
  }
  console.log(res);
  return res;
}


function updateVote(form, surveyID) {
  var res = {
    status: true,
    msg: null,
    data: null,
  };
  console.log('rrr');
  console.log(form);
  // Giả sử `questions` nằm trong `form`
  form.questions.forEach((question, index) => {
      console.log(`Question ${index + 1}:`, question.question);
      console.log('Answers:', question.answers);
  });
  var lock = LockService.getScriptLock();
  lock.waitLock(6 * 60 * 1000);
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    let headers;
    let keys = {};
    let newValue = [];
    if (form.isQuestChange) {
      // Thay đổi dữ liệu sheet câu hỏi
      let nameSheetQuestion = '【' + surveyID + '】' + listSheet.questions;
      let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetQuestion);
      let startRow = 10;
      let lastRow = sheetQuestion.getMaxRows();
      let numRows = lastRow - startRow;
      if (numRows > 0) {
        sheetQuestion.deleteRows(startRow, numRows);
      }
      createQuestionDataVote(form, nameSheetQuestion);
      // Thay đổi dữ liệu sheet câu trả lời
      let nameSheetResponse = '【' + surveyID + '】' + listSheet.responseTotal;
      let sheetResponse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetResponse);
      SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetResponse);
      checkSheetAnswer(form.id);
    }

    form.statistics = (form.statistics) ? '表示' : '非表示';
    form.usePassCode = (form.usePassCode) ? 'はい' : 'いいえ';
    form.informationRequired = (form.informationRequired) ? 'あり' : 'なし';
    console.log(form.informationRequired);
    // Tìm trong sheet master và thay đổi
    rows = sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues();
    headers = rows.shift();
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
      if (form[k]) {
        newValue[keys[k]] = form[k];
      }
    }
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][keys.id] == surveyID) {
        sheet.getRange(i + 6, 2, 1, newValue.length).setValues([newValue]);
        // Gán lại hyperlink cho list
        let nameSheetQuestion = '【' + surveyID + '】' + listSheet.questions;
        let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetQuestion);
        let urlSheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sheetQuestion.getSheetId();
        let nameSheetResponse = '【' + surveyID + '】' + listSheet.responseTotal;
        let sheetResponse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetResponse);
        let urlSheetResponse = SpreadsheetApp.getActiveSpreadsheet().getUrl() + '#gid=' + sheetResponse.getSheetId();
        sheet.getRange("L" + (i + 6)).setFormula('=HYPERLINK("' + urlSheetQuestion + '", "' + listSheet.questions + '")');
        sheet.getRange("M" + (i + 6)).setFormula('=HYPERLINK("' + urlSheetResponse + '", "' + listSheet.responseTotal + '")');
        // Thay đổi thông tin trong sheet câu hỏi
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetQuestion);
        sheet.getRange("E5").setValue(form.nameVote);
        sheet.getRange("E7").setValue(form.description);
      }
    }

    res.msg = "編集プロセスは成功しました。";
    res.data = form;
  } catch (error) {
    console.log({ error })
    res.status = false;
    res.msg = error.name;
  }
  return res;
}

function delVote(surveyID) {
  let res = {
    status: true,
    msg: ''
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(6 * 60 * 1000);
  try {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    let headers;
    let keys = {};
    let newValue = [];
    // Xóa trong master sheet
    rows = sheet.getRange(5, 2, sheet.getLastRow() - 4, sheet.getMaxColumns() - 2).getDisplayValues();
    headers = rows.shift();
    for (const [k, v] of Object.entries(master_sheet)) {
      keys[k] = headers.indexOf(v);
    }
    for (let i = 0; i < rows.length; i++) {
      if (rows[i][keys.id] == surveyID) {
        sheet.deleteRow(i + 6);
      }
    }
    // Xóa sheet câu hỏi và trả lời
    let nameSheetQuestion = '【' + surveyID + '】' + listSheet.questions;
    let sheetQuestion = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetQuestion);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetQuestion);
    let nameSheetResponse = '【' + surveyID + '】' + listSheet.responseTotal;
    let sheetResponse = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetResponse);
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheetResponse);
  } catch (error) {
    console.log(`delVote error: ${error}`)
    res.status = false
    res.msg = error
  } finally {
    return res
  }
}

function dataStatistics(surveyID) {
  var res = {
    status: true,
    data: null,
    msg: 'Success'
  }

  let keyAnswers = {};
  let results = {};
  let vote = getVoteById(surveyID);
  let linkRegex = /<([^>]+)>/;
  try {
    // Danh sách câu hỏi
    let { keys, rows, surveySetting } = getQuestions(surveyID).data;
    let keyQuestions = keys;
    let rowQuestions = rows;
    // Danh sách câu trả lời
    let nameSheetAnswer = `【${surveyID}】${listSheet.responseTotal}`;
    let s = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nameSheetAnswer);
    rowAnswers = s.getRange(5, 2, s.getLastRow() - 4, s.getMaxColumns() - 2).getDisplayValues();
    let headerAnswers = rowAnswers.shift();
    //
    for (let i = 0; i < rowQuestions.length; i++) {
      let nameQuestions = rowQuestions[i][keyQuestions.no] + '．' + rowQuestions[i][keyQuestions.question]
      keyAnswers[nameQuestions] = headerAnswers.indexOf(nameQuestions);
      let typeQuestion = rowQuestions[i][keyQuestions.question_type];
      let criteria = (rowQuestions[i][keyQuestions.criterias].length > 0) ? true : false;
      results[i] = {};
      results[i]['typeQuestion'] = typeQuestion;
      results[i]['criteria'] = criteria;
      results[i]['question_answers'] = {};
      // Duyệt mảng câu trả lời và tổng hợp
      if (typeQuestion == 'oneSelect' && criteria == false) {
        // Tách link cho câu trả lời
        let question_answers = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.question_answers].length; key++) {
          let text = rowQuestions[i][keyQuestions.question_answers][key];
          let match = text.match(linkRegex);
          if (match) {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key].replace(match[0], ''));
          } else {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let answersContent = [];
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            if (str != '') {
              answersContent = answersContent.concat(['[' + str + ']']);
            }
          }
        }
        // Đếm số lượng theo đáp án
        for (let a = 0; a < question_answers.length; a++) {
          results[i]['question_answers'][a] = { name: question_answers[a], total: 0 };
          total = answersContent.filter(item => item === '[' + question_answers[a] + ']').length;
          results[i]['question_answers'][a]['total'] = total;
        }
      } else if (typeQuestion == 'multiSelect' && criteria == false) {
        // Tách link cho câu trả lời
        let question_answers = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.question_answers].length; key++) {
          let text = rowQuestions[i][keyQuestions.question_answers][key];
          let match = text.match(linkRegex);
          if (match) {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key].replace(match[0], ''));
          } else {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let answersContent = [];
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            let regex = /^-\s(.+)/;
            let match = regex.exec(str);
            if (str != '') {
              answersContent = answersContent.concat(['[' + match[1] + ']']);
            }
          }
        }
        // Đếm số lượng theo đáp án
        for (let a = 0; a < question_answers.length; a++) {
          results[i]['question_answers'][a] = { name: question_answers[a], total: 0 };
          total = answersContent.filter(item => item === '[' + question_answers[a] + ']').length;
          results[i]['question_answers'][a]['total'] = total;
        }
      } else if (typeQuestion == 'score' && criteria == false) {
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let answersContent = [];
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            if (str != '') {
              answersContent = answersContent.concat([str]);
            }
          }
        }
        // Đếm số lượng theo đáp án
        results[i]['question_answers'][0] = { total: 0 };
        total = answersContent.reduce((acc, curr) => acc + parseFloat(curr), 0);
        results[i]['question_answers'][0]['total'] = total;
      } else if (typeQuestion == 'input' && criteria == false) {
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let answersContent = [];
        for (let y = 0; y < rowAnswers.length; y++) {
          if (rowAnswers[y][keyAnswers[nameQuestions]] != '') {
            answersContent.push(rowAnswers[y][keyAnswers[nameQuestions]]);
          }
        }
        // Tổng hợp câu trả lời
        results[i]['question_answers'][0] = answersContent;
      }
      else if (typeQuestion == 'oneSelect' && criteria == true) {
        // Tách link cho câu trả lời
        let question_answers = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.question_answers].length; key++) {
          let text = rowQuestions[i][keyQuestions.question_answers][key];
          let match = text.match(linkRegex);
          if (match) {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key].replace(match[0], ''));
          } else {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key]);
          }
        }
        // Tách link cho câu tiêu chí
        let question_criteria = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.criterias].length; key++) {
          let text = rowQuestions[i][keyQuestions.criterias][key];
          let match = text.match(linkRegex);
          if (match) {
            question_criteria.push(rows[i][keys.criterias][key].replace(match[0], ''));
          } else {
            question_criteria.push(rows[i][keys.criterias][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let criteriaContent = {};
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            let regex = /^-\s(.*?)(?=:)/;
            let match = regex.exec(str);
            if (str != '') {
              for (let c = 0; c < question_criteria.length; c++) {
                if (match[1] == question_criteria[c]) {
                  let regex = /: (.+)/;
                  let matches = str.match(regex);
                  if (!criteriaContent[question_criteria[c]]) {
                    criteriaContent[question_criteria[c]] = [];
                  }
                  criteriaContent[question_criteria[c]] = criteriaContent[question_criteria[c]].concat(['[' + matches[1] + ']']);
                }
              }
            }
          }
        }
        // Đếm số lượng theo đáp án
        for (let a = 0; a < question_answers.length; a++) {
          results[i]['question_answers'][a] = { name: question_answers[a], criteria: {} };
          for (const [k, v] of Object.entries(criteriaContent)) {
            total = v.filter(item => item === '[' + question_answers[a] + ']').length;
            results[i]['question_answers'][a]['criteria'][k] = total;
          }
        }
      } else if (typeQuestion == 'multiSelect' && criteria == true) {
        // Tách link cho câu trả lời
        let question_answers = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.question_answers].length; key++) {
          let text = rowQuestions[i][keyQuestions.question_answers][key];
          let match = text.match(linkRegex);
          if (match) {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key].replace(match[0], ''));
          } else {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key]);
          }
        }
        // Tách link cho câu tiêu chí
        let question_criteria = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.criterias].length; key++) {
          let text = rowQuestions[i][keyQuestions.criterias][key];
          let match = text.match(linkRegex);
          if (match) {
            question_criteria.push(rows[i][keys.criterias][key].replace(match[0], ''));
          } else {
            question_criteria.push(rows[i][keys.criterias][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let criteriaContent = {};
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            let regex = /- (.+?): \[.+\]/;
            let match = regex.exec(str);
            if (str != '') {
              for (let c = 0; c < question_criteria.length; c++) {
                if (match[1] == question_criteria[c]) {
                  let regex = /\[.*?\]/g;
                  let matches = str.match(regex);
                  if (!criteriaContent[question_criteria[c]]) {
                    criteriaContent[question_criteria[c]] = [];
                  }
                  criteriaContent[question_criteria[c]] = criteriaContent[question_criteria[c]].concat(matches);
                }
              }
            }
          }
        }
        // Đếm số lượng theo đáp án
        for (let a = 0; a < question_answers.length; a++) {
          results[i]['question_answers'][a] = { name: question_answers[a], criteria: {} };
          for (const [k, v] of Object.entries(criteriaContent)) {
            total = v.filter(item => item === '[' + question_answers[a] + ']').length;
            results[i]['question_answers'][a]['criteria'][k] = total;
          }
        }
      } else if (typeQuestion == 'score' && criteria == true) {
        // Tách link cho câu trả lời
        let question_answers = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.question_answers].length; key++) {
          let text = rowQuestions[i][keyQuestions.question_answers][key];
          let match = text.match(linkRegex);
          if (match) {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key].replace(match[0], ''));
          } else {
            question_answers.push(rowQuestions[i][keyQuestions.question_answers][key]);
          }
        }
        // Tách link cho câu tiêu chí
        let question_criteria = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.criterias].length; key++) {
          let text = rowQuestions[i][keyQuestions.criterias][key];
          let match = text.match(linkRegex);
          if (match) {
            question_criteria.push(rows[i][keys.criterias][key].replace(match[0], ''));
          } else {
            question_criteria.push(rows[i][keys.criterias][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let criteriaContent = {};
        for (let y = 0; y < rowAnswers.length; y++) {
          rowAnswers[y][keyAnswers[nameQuestions]] = _split(rowAnswers[y][keyAnswers[nameQuestions]]);
          for (let a = 0; a < rowAnswers[y][keyAnswers[nameQuestions]].length; a++) {
            let str = rowAnswers[y][keyAnswers[nameQuestions]][a];
            let regex = /^-\s(.*?)(?=:)/;
            let match = regex.exec(str);
            if (str != '') {
              for (let c = 0; c < question_criteria.length; c++) {
                if (match[1] == question_criteria[c]) {
                  let regex = /: (.+)/;
                  let matches = str.match(regex);
                  if (!criteriaContent[question_criteria[c]]) {
                    criteriaContent[question_criteria[c]] = [];
                  }
                  criteriaContent[question_criteria[c]] = criteriaContent[question_criteria[c]].concat([matches[1]]);
                }
              }
            }
          }
        }
        // Tính tổng điểm theo tất cả đáp án
        results[i]['question_answers'][0] = { criteria: {} };
        for (const [k, v] of Object.entries(criteriaContent)) {
          total = v.reduce((acc, curr) => acc + parseFloat(curr), 0);
          results[i]['question_answers'][0]['criteria'][k] = total;
        }
      } else if (typeQuestion == 'input' && criteria == true) {
        // Tách link cho câu tiêu chí
        let question_criteria = [];
        for (let key = 0; key < rowQuestions[i][keyQuestions.criterias].length; key++) {
          let text = rowQuestions[i][keyQuestions.criterias][key];
          let match = text.match(linkRegex);
          if (match) {
            question_criteria.push(rows[i][keys.criterias][key].replace(match[0], ''));
          } else {
            question_criteria.push(rows[i][keys.criterias][key]);
          }
        }
        //Duyệt tất cả câu trả lời của câu hỏi hiện tại
        let criterias = question_criteria.map(criterion => `- ${criterion}:`);
        let result = {};
        for (let y = 0; y < rowAnswers.length; y++) {
          let content = rowAnswers[y][keyAnswers[nameQuestions]];
          criterias.forEach(criterion => {
            const regex = new RegExp(`${criterion}\\s*([\\s\\S]*?)(?=\n-\\s*(?:${question_criteria.join('|')}):|$)`, 'g');
            let match;

            while ((match = regex.exec(content)) !== null) {
              const criterionName = criterion.substring(2, criterion.length - 1);

              // Nếu tiêu chí chưa tồn tại trong đối tượng kết quả, khởi tạo mảng
              if (!result[criterionName]) {
                result[criterionName] = [];
              }

              // Thêm nội dung vào mảng tương ứng với tiêu chí
              result[criterionName].push(match[1].trim());
            }
          });
        }
        // Tổng hợp các trả lời
        results[i]['question_answers'][0] = { criteria: result };
      }
    }
  } catch (error) {
    console.log("dataStatistics:", { error })
    res.status = false
    res.msg = error.name
  } finally {
    res.data = results;
    res.vote = vote;
    return res;
  }
}

// QR Code
function getQRCodeUrl(text = "https://www.google.com/", size = 300, config = {}) {
  var res = {
    status: true,
    data: null,
    msg: null
  }
  try {
    res.data = { text, size, config };
  } catch (error) {
    console.log({ error })
    res.status = false
    res.msg = error
  } finally {
    return res
  }
}


/* Private functions */

function _split(textStr, regex = '\n') {
  return (textStr !== undefined && textStr !== null) ? textStr.split(regex) : '';
}

function recreateShortLink() {
  let res = true;
  try {
    let { keys, rows } = getDataFromSheet2(listSheet.listVote, master_sheet);
    // console.log(keys, rows);
    let url = headerSetting.getDataByKey('gas_url');
    console.log({ url })
    let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(listSheet.listVote);
    for (let i = 0; i < rows.length; i++) {
      let surveyID = rows[i][keys.id];
      let pageObject = { page: "form", surveyID: surveyID };
      let encodedPageParam = encodeURIComponent(JSON.stringify(pageObject));
      let urlSurvey = `https://myportal.sateraito.jp/gas?url=${url}?page=${encodedPageParam}`;
      // let urlQuiz = `${url}?page=form&code=${quizCode}`;
      let shortLink = createShortURL(urlSurvey);
      sheet.getRange('K' + Number(i + 6)).setValue(shortLink);
      console.log(shortLink);
    }
  } catch (error) {
    console.log(`recreateShortLink error: ${error}`)
    res = false
  } finally {
    return res
  }
}