

const SHEET_NAME = 'Expenses';
const CHAT_SHEET_NAME = 'Chat';
const USERS_SHEET_NAME = 'Users';

const SPREADSHEET_ID = '';

function getSpreadsheet() {
  if (SPREADSHEET_ID && SPREADSHEET_ID !== '') {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
  }
  return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.getActive();
}

function setupSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);

    const initialHeaders = ['Timestamp', 'Meal Name', 'Category', 'Due Date', 'Total Amount', 'Payer', 'Status', 'Attachment'];
    sheet.appendRow(initialHeaders);
    sheet.getRange(1, 1, 1, initialHeaders.length).setFontWeight('bold');
  }
  return sheet;
}

function setupChatSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(CHAT_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CHAT_SHEET_NAME);
    const headers = ['Timestamp', 'Sender', 'Message', 'Attachment'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  alignChatSheetHeaders(sheet);
  return sheet;
}

function alignUsersSheetHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([['Username', 'Password', 'FriendName', 'LineUserId']]);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    return;
  }
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.trim());
  if (!headers.includes('LineUserId')) {
    sheet.getRange(1, lastCol + 1).setValue('LineUserId');
    sheet.getRange(1, 1, 1, lastCol + 1).setFontWeight('bold');
  }
}

function setupUsersSheet() {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(USERS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USERS_SHEET_NAME);
    const headers = ['Username', 'Password', 'FriendName', 'LineUserId'];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  } else {
    alignUsersSheetHeaders(sheet);
  }
  return sheet;
}

function getSetting(key) {
  let defaultVal = '';
  if (key === 'LINE_CHANNEL_ID') {
    defaultVal = '2010744890';
  } else if (key === 'LINE_CHANNEL_SECRET') {
    defaultVal = '04cc8057667b95ae8658bc750653c5af';
  } else if (key === 'LINE_NOTIFY_TOKEN') {
    defaultVal = 'aLu3cQ8JuLSCfI8BP31Pmzr7V0ni2vFrvWGi2C3Kt2IPOSJ6nDzQM6skwyTd7a9i2iUEQVs2bBBxSqF1UQVaAY1GndiNieyDfuZoKv6eWI934g7ynRBmEPW7ykVVaRcic2S0y+gGbciqloKUeZ9TfQdB04t89/1O/w1cDnyilFU=';
  }

  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      return defaultVal;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return defaultVal;
    }
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    for (let i = 0; i < values.length; i++) {
      if (values[i][0].toString() === key) {
        const val = values[i][1].toString().trim();
        return val || defaultVal;
      }
    }
  } catch (e) {
    console.error('getSetting error:', e);
  }
  return defaultVal;
}

function setSetting(key, value) {
  try {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName('Settings');
    if (!sheet) {
      sheet = ss.insertSheet('Settings');
      sheet.appendRow(['Key', 'Value']);
      sheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    }
    const lastRow = sheet.getLastRow();
    let rowIdx = -1;
    if (lastRow > 1) {
      const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < values.length; i++) {
        if (values[i][0].toString() === key) {
          rowIdx = i + 2;
          break;
        }
      }
    }
    if (rowIdx !== -1) {
      sheet.getRange(rowIdx, 2).setValue(value);
    } else {
      sheet.appendRow([key, value]);
    }
  } catch (e) {
    console.error('setSetting error:', e);
  }
}

function sendLineNotify(message, imageUrl) {
  const token = getSetting('LINE_NOTIFY_TOKEN');
  if (!token) return;
  try {
    const url = 'https://api.line.me/v2/bot/message/broadcast';
    const messages = [
      {
        type: 'text',
        text: message
      }
    ];

    if (imageUrl) {
      const directUrl = getDirectImageLink(imageUrl);
      if (directUrl) {
        messages.push({
          type: 'image',
          originalContentUrl: directUrl,
          previewImageUrl: directUrl
        });
      }
    }

    const payload = {
      messages: messages
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    console.log('Broadcast response:', response.getContentText());
  } catch (err) {
    console.error('Error sending LINE Broadcast: ' + err.toString());
  }
}

function getLineUserIdByFriendName(friendName) {
  try {
    const userSheet = setupUsersSheet();
    const lastRow = userSheet.getLastRow();
    if (lastRow <= 1) return null;
    const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0].map(h => h.trim());
    const lineUserIdColIdx = headers.indexOf('LineUserId');
    const friendNameColIdx = headers.indexOf('FriendName');
    if (lineUserIdColIdx === -1 || friendNameColIdx === -1) return null;
    
    const values = userSheet.getRange(2, 1, lastRow - 1, userSheet.getLastColumn()).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowFriendName = values[i][friendNameColIdx];
      if (rowFriendName && rowFriendName.toString().toLowerCase() === friendName.toLowerCase()) {
        return values[i][lineUserIdColIdx] || null;
      }
    }
  } catch (e) {
    console.error('getLineUserIdByFriendName error:', e);
  }
  return null;
}

function getDirectImageLink(driveUrl) {
  if (!driveUrl) return null;
  const match = driveUrl.match(/id=([^&]+)/) || driveUrl.match(/\/d\/([^/]+)/);
  if (match) {
    const fileId = match[1];
    return "https://drive.google.com/uc?export=view&id=" + fileId;
  }
  return driveUrl;
}

function sendLinePush(userId, message, imageUrl) {
  if (!userId) return;
  const token = getSetting('LINE_NOTIFY_TOKEN');
  if (!token) return;
  try {
    const url = 'https://api.line.me/v2/bot/message/push';
    const messages = [
      {
        type: 'text',
        text: message
      }
    ];

    if (imageUrl) {
      const directUrl = getDirectImageLink(imageUrl);
      if (directUrl) {
        messages.push({
          type: 'image',
          originalContentUrl: directUrl,
          previewImageUrl: directUrl
        });
      }
    }

    const payload = {
      to: userId,
      messages: messages
    };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + token },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    const response = UrlFetchApp.fetch(url, options);
    console.log('Push response:', response.getContentText());
  } catch (err) {
    console.error('Error sending LINE Push: ' + err.toString());
  }
}


function alignChatSheetHeaders(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = ['Timestamp', 'Sender', 'Message', 'Attachment'];
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.trim());
  if (!currentHeaders.includes('Attachment')) {
    sheet.getRange(1, 4).setValue('Attachment');
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
  }
}

function alignSheetHeaders(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  const fixedHeaders = ['Timestamp', 'Meal Name', 'Category', 'Due Date', 'Total Amount', 'Payer', 'Status', 'Attachment'];

  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, fixedHeaders.length).setValues([fixedHeaders]);
    sheet.getRange(1, 1, 1, fixedHeaders.length).setFontWeight('bold');
    return;
  }

  const currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.trim());

  let needsAlignment = false;
  for (let i = 0; i < fixedHeaders.length; i++) {
    if (currentHeaders[i] !== fixedHeaders[i]) {
      needsAlignment = true;
      break;
    }
  }

  if (!needsAlignment) return; 

  const oldData = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  const friendHeaders = currentHeaders.filter(h => !fixedHeaders.includes(h) && h !== 'Participants' && h !== '');

  if (friendHeaders.length === 0 && currentHeaders.includes('Participants')) {
    migrateOldSheetIfNeeded(sheet);
    return;
  }

  const newHeaders = [...fixedHeaders, ...friendHeaders];
  const newRows = [];

  oldData.forEach(row => {
    const timestamp = currentHeaders.indexOf('Timestamp') !== -1 ? row[currentHeaders.indexOf('Timestamp')] : new Date();
    const meal = currentHeaders.indexOf('Meal Name') !== -1 ? row[currentHeaders.indexOf('Meal Name')] : '';
    const category = currentHeaders.indexOf('Category') !== -1 ? row[currentHeaders.indexOf('Category')] : 'Food';
    const dueDate = currentHeaders.indexOf('Due Date') !== -1 ? row[currentHeaders.indexOf('Due Date')] : '';
    const amount = currentHeaders.indexOf('Total Amount') !== -1 ? parseFloat(row[currentHeaders.indexOf('Total Amount')]) || 0 : 0;
    const payer = currentHeaders.indexOf('Payer') !== -1 ? row[currentHeaders.indexOf('Payer')] : '';
    const status = currentHeaders.indexOf('Status') !== -1 ? row[currentHeaders.indexOf('Status')] : 'Confirmed';
    const attachment = currentHeaders.indexOf('Attachment') !== -1 ? row[currentHeaders.indexOf('Attachment')] : '';

    const newRow = [timestamp, meal, category, dueDate, amount, payer, status, attachment];

    friendHeaders.forEach(friend => {
      const idx = currentHeaders.indexOf(friend);
      newRow.push(idx !== -1 ? row[idx] : 0);
    });

    newRows.push(newRow);
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');
  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
  }
}

function migrateOldSheetIfNeeded(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow <= 1 || lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.trim());
  const participantsIndex = headers.indexOf('Participants');

  if (participantsIndex === -1) {
    return;
  }

  const mealNameIndex = headers.indexOf('Meal Name');
  const amountIndex = headers.indexOf('Total Amount');
  const payerIndex = headers.indexOf('Payer');
  const timestampIndex = headers.indexOf('Timestamp');

  const oldDataRange = sheet.getRange(2, 1, lastRow - 1, lastCol);
  const oldValues = oldDataRange.getValues();

  const allFriendsSet = new Set();
  oldValues.forEach(row => {
    const pStr = row[participantsIndex] || '';
    if (pStr) {
      pStr.split(',').forEach(name => {
        const trimmed = name.trim();
        if (trimmed) allFriendsSet.add(trimmed);
      });
    }
    const payer = row[payerIndex];
    if (payer) allFriendsSet.add(payer.trim());
  });

  const friendsList = Array.from(allFriendsSet);
  if (friendsList.length === 0) {
    friendsList.push('นาย A', 'นาย B', 'นาย C');
  }

  const newHeaders = ['Timestamp', 'Meal Name', 'Category', 'Due Date', 'Total Amount', 'Payer', 'Status', 'Attachment', ...friendsList];

  const newRows = [];
  oldValues.forEach(row => {
    const timestamp = row[timestampIndex];
    const meal = row[mealNameIndex];
    const amount = parseFloat(row[amountIndex]) || 0;
    const payer = row[payerIndex];
    const pStr = row[participantsIndex] || '';
    const participants = pStr.split(',').map(p => p.trim());

    const newRow = [
      timestamp,
      meal,
      'Food', 
      '',     
      amount,
      payer,
      'Confirmed', 
      ''           
    ];

    const share = Math.round((amount / participants.length) * 100) / 100;

    friendsList.forEach(friend => {
      newRow.push(participants.includes(friend) ? share : 0);
    });

    newRows.push(newRow);
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  sheet.getRange(1, 1, 1, newHeaders.length).setFontWeight('bold');

  if (newRows.length > 0) {
    sheet.getRange(2, 1, newRows.length, newHeaders.length).setValues(newRows);
  }
}

function getChatHistory() {
  const sheet = setupChatSheet();
  const lastRow = sheet.getLastRow();
  const chat = [];
  if (lastRow > 1) {
    const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues();
    values.forEach(row => {
      const timestamp = row[0];
      const sender = row[1];
      const message = row[2];
      const attachment = row[3] || '';
      chat.push({
        timestamp: timestamp instanceof Date ? timestamp.toISOString() : timestamp,
        sender: sender,
        message: message,
        attachment: attachment
      });
    });
  }
  return chat.slice(-100);
}

function getBalancesAndHistory(friendsList) {
  const sheet = setupSheet();
  alignSheetHeaders(sheet);

  const lastRow = sheet.getLastRow();
  const finalLastCol = sheet.getLastColumn();
  const finalHeaders = sheet.getRange(1, 1, 1, finalLastCol).getValues()[0].map(h => h.trim());

  const fixedHeaders = ['Timestamp', 'Meal Name', 'Category', 'Due Date', 'Total Amount', 'Payer', 'Status', 'Attachment'];

  let activeFriends = friendsList || [];
  if (!Array.isArray(activeFriends)) activeFriends = [];

  const discoveredNames = finalHeaders.slice(fixedHeaders.length);
  discoveredNames.forEach(name => {
    if (name && !activeFriends.some(f => f.name.trim().toLowerCase() === name.trim().toLowerCase())) {
      activeFriends.push({ name: name.trim(), initial: 0 });
    }
  });

  if (activeFriends.length === 0) {
    activeFriends = [
      { name: 'นาย A', initial: 0 },
      { name: 'นาย B', initial: 0 },
      { name: 'นาย C', initial: 0 }
    ];
  }

  const friendsNames = activeFriends.map(f => f.name.trim());

  let headersUpdated = false;
  friendsNames.forEach(name => {
    if (!finalHeaders.includes(name)) {
      finalHeaders.push(name);
      headersUpdated = true;
    }
  });

  if (headersUpdated) {
    sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
    sheet.getRange(1, 1, 1, finalHeaders.length).setFontWeight('bold');
  }

  const balances = {};
  activeFriends.forEach(friend => {
    balances[friend.name.trim()] = parseFloat(friend.initial) || 0;
  });

  const history = [];

  if (lastRow > 1) {
    const dataRange = sheet.getRange(2, 1, lastRow - 1, finalHeaders.length);
    const values = dataRange.getValues();

    values.forEach((row, index) => {
      const timestamp = row[0];
      const meal = row[1];
      const category = row[2] || 'Food';
      const dueDate = row[3] ? (row[3] instanceof Date ? row[3].toISOString().split('T')[0] : row[3].toString()) : '';
      const amount = parseFloat(row[4]) || 0;
      const payer = row[5];
      const status = row[6] || '';
      const attachment = row[7] || '';

      if (!payer || amount <= 0) return;

      const participants = [];
      for (let c = fixedHeaders.length; c < finalHeaders.length; c++) {
        const friendName = finalHeaders[c];
        const val = row[c];
        const parsedVal = parseFloat(val);
        const hasJoined = val === true || val === 'TRUE' || (!isNaN(parsedVal) && parsedVal > 0);
        if (hasJoined) {
          participants.push(friendName);
        }
      }

      if (participants.length === 0) return;

      if (status !== 'Pending') {
        const share = amount / participants.length;

        if (balances[payer] !== undefined) {
          balances[payer] += amount;
        }

        participants.forEach(p => {
          if (balances[p] !== undefined) {
            balances[p] -= share;
          }
        });
      }

      history.push({
        id: index + 2, 
        timestamp: timestamp instanceof Date ? timestamp.toISOString() : timestamp,
        meal: meal,
        category: category,
        dueDate: dueDate,
        amount: amount,
        payer: payer,
        status: status,
        attachment: attachment,
        participants: participants
      });
    });
  }

  const balancesArray = friendsNames.map(name => {
    return {
      name: name,
      balance: Math.round((balances[name] || 0) * 100) / 100
    };
  });

  history.reverse();

  const chatHistory = getChatHistory();

  const settings = {
    lineChannelId: getSetting('LINE_CHANNEL_ID'),
    hasLineLogin: getSetting('LINE_CHANNEL_ID') !== '' && getSetting('LINE_CHANNEL_SECRET') !== '',
    hasLineNotify: getSetting('LINE_NOTIFY_TOKEN') !== ''
  };

  return {
    status: 'success',
    balances: balancesArray,
    history: history.slice(0, 100),
    chat: chatHistory,
    settings: settings
  };
}

function saveFileToDrive(base64Data, filename) {
  if (!base64Data) return '';
  try {
    const folderIterator = DriveApp.getFoldersByName("Gin-Gun Slips");
    let folder;
    if (folderIterator.hasNext()) {
      folder = folderIterator.next();
    } else {
      folder = DriveApp.createFolder("Gin-Gun Slips");
    }

    const commaIdx = base64Data.indexOf(',');
    const dataPart = commaIdx !== -1 ? base64Data.substring(commaIdx + 1) : base64Data;
    const mimePart = commaIdx !== -1 ? base64Data.substring(5, base64Data.indexOf(';')) : 'image/png';

    const bytes = Utilities.newBlob(Utilities.base64Decode(dataPart), mimePart, filename);
    const file = folder.createFile(bytes);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    return 'Error saving file to Drive: ' + err.toString();
  }
}

function handleLineCallback(e) {
  try {
    const code = e.parameter.code;
    const stateParam = e.parameter.state;
    const channelId = getSetting('LINE_CHANNEL_ID');
    const channelSecret = getSetting('LINE_CHANNEL_SECRET');

    const tokenUrl = 'https://api.line.me/oauth2/v2.1/token';
    const tokenPayload = {
      grant_type: 'authorization_code',
      code: code,
      redirect_uri: ScriptApp.getService().getUrl(),
      client_id: channelId,
      client_secret: channelSecret
    };

    const tokenResponse = UrlFetchApp.fetch(tokenUrl, {
      method: 'post',
      payload: tokenPayload,
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      muteHttpExceptions: true
    });

    const tokenData = JSON.parse(tokenResponse.getContentText());
    if (tokenData.error) {
      throw new Error('LINE Token Exchange Error: ' + tokenData.error_description);
    }

    const accessToken = tokenData.access_token;
    const profileUrl = 'https://api.line.me/v2/profile';
    const profileResponse = UrlFetchApp.fetch(profileUrl, {
      headers: { Authorization: 'Bearer ' + accessToken }
    });
    const profileData = JSON.parse(profileResponse.getContentText());
    const lineUserId = profileData.userId;
    const lineDisplayName = profileData.displayName;


    let linkUsername = '';
    if (stateParam) {
      const match = stateParam.match(/[?&]link_username=([^&]+)/);
      if (match) {
        linkUsername = decodeURIComponent(match[1]);
      }
    }

    let appUrl = 'https://gin-gun.vercel.app/';
    if (stateParam) {
      const qIdx = stateParam.indexOf('?');
      const hIdx = stateParam.indexOf('#');
      let cleanUrl = stateParam;
      if (qIdx !== -1 && hIdx !== -1) {
        cleanUrl = stateParam.substring(0, Math.min(qIdx, hIdx));
      } else if (qIdx !== -1) {
        cleanUrl = stateParam.substring(0, qIdx);
      } else if (hIdx !== -1) {
        cleanUrl = stateParam.substring(0, hIdx);
      }
      appUrl = cleanUrl;
    }
    const connector = appUrl.indexOf('?') === -1 ? '?' : '&';
    let finalUrl = appUrl + connector + 'line_login_success=true&lineUserId=' + encodeURIComponent(lineUserId) +
      '&lineDisplayName=' + encodeURIComponent(lineDisplayName);
    if (linkUsername) {
      finalUrl += '&link_username=' + encodeURIComponent(linkUsername);
    }

    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <meta http-equiv="refresh" content="2;url=${finalUrl}">
        <title>กำลังพาท่านกลับ...</title>
        <style>
          body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #0f172a; color: #f8fafc; margin: 0; }
          .card { background: rgba(30, 41, 59, 0.7); padding: 2rem; border-radius: 1rem; border: 1px solid #334155; text-align: center; max-width: 400px; }
          .spinner { border: 4px solid rgba(255,255,255,0.1); width: 36px; height: 36px; border-radius: 50%; border-left-color: #22c55e; animation: spin 1s linear infinite; margin: 1rem auto; }
          @keyframes spin { 0% { transform: rotate(0deg); } 100% { transform: rotate(360deg); } }
          a { color: #4ade80; text-decoration: underline; font-weight: bold; }
          a:hover { color: #86efac; }
        </style>
      </head>
      <body>
        <div class="card">
          <div class="spinner"></div>
          <h3>✅ เข้าสู่ระบบ LINE สำเร็จแล้ว</h3>
          <p>ยินดีต้อนรับ <strong>${lineDisplayName}</strong></p>
          <p style="font-size:13px; margin-top:12px;">กำลังพาท่านกลับสู่แอปพลิเคชัน...</p>
          <p style="font-size:11px; color:#94a3b8; margin-top:16px;">หากไม่ถูกพาไปโดยอัตโนมัติ <a href="${finalUrl}">คลิกที่นี่</a></p>
        </div>
        <script>
          try { window.top.location.href = ${JSON.stringify(finalUrl)}; } catch(e) {}
          setTimeout(function() {
            try { window.top.location.href = ${JSON.stringify(finalUrl)}; } catch(e) {}
            window.location.href = ${JSON.stringify(finalUrl)};
          }, 1500);
        </script>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(html)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const returnUrl = (e && e.parameter && e.parameter.state) || 'https://gin-gun.vercel.app/';
    const errorHtml = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="utf-8">
        <title>เกิดข้อผิดพลาด</title>
        <style>
          body { font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; background-color: #0f172a; color: #f8fafc; margin: 0; }
          .card { background: rgba(30, 41, 59, 0.7); padding: 2rem; border-radius: 1rem; border: 1px solid #f87171; text-align: center; max-width: 500px; }
          .error { color: #fca5a5; font-size: 12px; word-break: break-all; margin-top: 12px; padding: 12px; background: rgba(0,0,0,0.3); border-radius: 8px; }
          a { color: #4ade80; text-decoration: underline; font-weight: bold; }
        </style>
      </head>
      <body>
        <div class="card">
          <h3 style="color:#f87171;">❌ เกิดข้อผิดพลาดในการเข้าสู่ระบบด้วย LINE</h3>
          <div class="error">${err.toString()}</div>
          <p style="margin-top:16px;font-size:13px;"><a href="${returnUrl}">กลับไปยังหน้าหลัก</a></p>
        </div>
      </body>
      </html>
    `;
    return HtmlService.createHtmlOutput(errorHtml)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
}

function doGet(e) {
  try {
    if (e && e.parameter && e.parameter.code) {
      return handleLineCallback(e);
    }

    let friendsList = [];
    if (e && e.parameter && e.parameter.friends) {
      try {
        friendsList = JSON.parse(decodeURIComponent(e.parameter.friends));
      } catch (err) {

      }
    }

    const result = getBalancesAndHistory(friendsList);
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("ไม่มีข้อมูลส่งมาประมวลผล");
    }

    const postData = JSON.parse(e.postData.contents);
    const action = postData.action || 'addExpense';
    const friendsList = postData.friends || [];

    if (action === 'saveSettings') {
      const lineChannelId = postData.lineChannelId;
      const lineChannelSecret = postData.lineChannelSecret;
      const lineNotifyToken = postData.lineNotifyToken;

      if (lineChannelId !== undefined) setSetting('LINE_CHANNEL_ID', lineChannelId);
      if (lineChannelSecret !== undefined) setSetting('LINE_CHANNEL_SECRET', lineChannelSecret);
      if (lineNotifyToken !== undefined) setSetting('LINE_NOTIFY_TOKEN', lineNotifyToken);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'บันทึกการตั้งค่า LINE สำเร็จ!'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'register') {
      const username = postData.username ? postData.username.trim().toLowerCase() : '';
      const password = postData.password ? postData.password.trim() : '';
      const friendName = postData.friendName ? postData.friendName.trim() : '';

      if (!username || !password || !friendName) {
        throw new Error("ข้อมูลไม่ครบถ้วนสำหรับการลงทะเบียน");
      }

      const userSheet = setupUsersSheet();
      const lastRow = userSheet.getLastRow();

      if (lastRow > 1) {
        const values = userSheet.getRange(2, 1, lastRow - 1, 1).getValues();
        const exists = values.some(row => row[0].toString().toLowerCase() === username);
        if (exists) {
          throw new Error("ชื่อผู้ใช้งาน (Username) นี้ถูกใช้ไปแล้ว");
        }
      }

      userSheet.appendRow([username, password, friendName]);

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'ลงทะเบียนผู้ใช้สำเร็จ!',
        user: { username, friendName }
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'login') {
      const username = postData.username ? postData.username.trim().toLowerCase() : '';
      const password = postData.password ? postData.password.trim() : '';

      if (!username || !password) {
        throw new Error("กรุณากรอกผู้ใช้งานและรหัสผ่าน");
      }

      const userSheet = setupUsersSheet();
      const lastRow = userSheet.getLastRow();
      let matchedUser = null;

      if (lastRow > 1) {
        const headers = userSheet.getRange(1, 1, 1, userSheet.getLastColumn()).getValues()[0].map(h => h.trim());
        const lineUserIdColIdx = headers.indexOf('LineUserId');
        const usernameColIdx = headers.indexOf('Username');
        const passwordColIdx = headers.indexOf('Password');
        const friendNameColIdx = headers.indexOf('FriendName');

        const values = userSheet.getRange(2, 1, lastRow - 1, userSheet.getLastColumn()).getValues();
        for (let i = 0; i < values.length; i++) {
          if (values[i][usernameColIdx].toString().toLowerCase() === username && values[i][passwordColIdx].toString() === password) {
            matchedUser = {
              username: values[i][usernameColIdx],
              friendName: values[i][friendNameColIdx],
              hasLineLinked: lineUserIdColIdx !== -1 && values[i][lineUserIdColIdx] !== '' && values[i][lineUserIdColIdx] !== null && values[i][lineUserIdColIdx] !== undefined
            };
            break;
          }
        }
      }

      if (!matchedUser) {
        throw new Error("ชื่อผู้ใช้งานหรือรหัสผ่านไม่ถูกต้อง");
      }

      return ContentService.createTextOutput(JSON.stringify({
        status: 'success',
        message: 'เข้าสู่ระบบสำเร็จ!',
        user: matchedUser
      })).setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'postChat') {
      const sender = postData.sender;
      const message = postData.message || '';
      const image = postData.image || ''; 
      const imageName = postData.imageName || 'chat_attachment_' + Date.now() + '.png';

      if (!sender) {
        throw new Error("กรุณาระบุผู้ส่งข้อความ");
      }

      let fileUrl = '';
      if (image) {
        fileUrl = saveFileToDrive(image, imageName);
      }

      const chatSheet = setupChatSheet();
      chatSheet.appendRow([new Date(), sender, message, fileUrl]);

      // Send LINE Notify
      let notifyMsg = `💬 [กินกัน] ${sender}: ${message}`;
      if (fileUrl) {
        notifyMsg += `\n📎 แนบรูปภาพ: ${fileUrl}`;
      }
      sendLineNotify(notifyMsg, fileUrl);

      const updatedResult = getBalancesAndHistory(friendsList);
      return ContentService.createTextOutput(JSON.stringify(updatedResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (action === 'confirmPayment') {
      const rowId = parseInt(postData.id);
      const friendName = postData.friendName; 

      if (isNaN(rowId) || rowId < 2) {
        throw new Error("ID ของรายการรอยืนยันไม่ถูกต้อง");
      }
      if (!friendName) {
        throw new Error("ไม่พบข้อมูลสิทธิ์ผู้กดยืนยัน");
      }

      const sheet = setupSheet();
      alignSheetHeaders(sheet);

      const lastRow = sheet.getLastRow();
      if (rowId > lastRow) {
        throw new Error("ไม่พบรายการที่ระบุ");
      }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => h.trim());
      const rowValues = sheet.getRange(rowId, 1, 1, sheet.getLastColumn()).getValues()[0];

      const friendIdx = headers.indexOf(friendName);
      if (friendIdx === -1) {
        throw new Error("ชื่อคุณไม่ได้อยู่ในรายชื่อเพื่อนในระบบ");
      }

      const isRecipient = parseFloat(rowValues[friendIdx]) > 0;
      if (!isRecipient) {
        throw new Error("คุณไม่ใช่ผู้รับเงินของรายการแจ้งโอนนี้ จึงไม่มีสิทธิ์กดยืนยันการรับเงิน");
      }

      sheet.getRange(rowId, 7).setValue('Confirmed');

      // Send LINE Notify
      const mealName = rowValues[1];
      const payer = rowValues[5];
      const amount = parseFloat(rowValues[4]) || 0;
      const msg = `✅ [กินกัน] ยืนยันการรับเงินโอนเรียบร้อยแล้ว!\n👤 ผู้ยืนยัน (ผู้รับเงิน): ${friendName}\n👤 ผู้โอน: ${payer}\n💰 ยอดเงิน: ${amount} บาท\n📝 รายการ: ${mealName}`;
      sendLineNotify(msg);

      // Send direct push notification to the payer
      const payerLineUserId = getLineUserIdByFriendName(payer);
      if (payerLineUserId) {
        const pushMsg = `✅ [กินกัน] การโอนเงินของคุณได้รับการยืนยันแล้ว!\n👤 ผู้รับเงิน: ${friendName}\n💰 ยอดเงิน: ${amount} บาท\n📝 รายการ: ${mealName}`;
        sendLinePush(payerLineUserId, pushMsg);
      }

      const updatedResult = getBalancesAndHistory(friendsList);
      return ContentService.createTextOutput(JSON.stringify(updatedResult))
        .setMimeType(ContentService.MimeType.JSON);

    }

    if (action === 'addExpense') {
      const meal = postData.meal;
      const amount = parseFloat(postData.amount);
      const payer = postData.payer;
      const participants = postData.participants || [];
      const category = postData.category || 'Food';
      const dueDate = postData.dueDate || '';
      const status = postData.status || 'Confirmed';

      if (!meal || isNaN(amount) || amount <= 0 || !payer || participants.length === 0) {
        throw new Error("ข้อมูลรายการไม่ถูกต้องหรือไม่ครบถ้วน");
      }

      const sheet = setupSheet();
      alignSheetHeaders(sheet);

      const finalLastCol = sheet.getLastColumn();
      const finalHeaders = sheet.getRange(1, 1, 1, finalLastCol).getValues()[0].map(h => h.trim());

      const fixedHeaders = ['Timestamp', 'Meal Name', 'Category', 'Due Date', 'Total Amount', 'Payer', 'Status', 'Attachment'];
      const friendsNames = friendsList.map(f => f.name.trim());

      let headersUpdated = false;
      friendsNames.forEach(name => {
        if (!finalHeaders.includes(name)) {
          finalHeaders.push(name);
          headersUpdated = true;
        }
      });

      if (headersUpdated) {
        sheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
        sheet.getRange(1, 1, 1, finalHeaders.length).setFontWeight('bold');
      }

      let fileUrl = '';
      if (postData.image) {
        const imageName = postData.imageName || 'slip_' + Date.now() + '.png';
        fileUrl = saveFileToDrive(postData.image, imageName);
      }

      const newRow = new Array(finalHeaders.length);
      newRow[0] = new Date();
      newRow[1] = meal;
      newRow[2] = category;
      newRow[3] = dueDate;
      newRow[4] = amount;
      newRow[5] = payer;
      newRow[6] = status;
      newRow[7] = fileUrl || postData.attachment || '';

      const share = Math.round((amount / participants.length) * 100) / 100;

      for (let c = fixedHeaders.length; c < finalHeaders.length; c++) {
        const friendName = finalHeaders[c];
        newRow[c] = participants.includes(friendName) ? share : 0;
      }

      sheet.appendRow(newRow);

      // Send LINE Notify
      let notifyMsg = '';
      if (category === 'Settlement') {
        notifyMsg = `💸 [กินกัน] มีการแจ้งโอนเงินคืนค้างชำระ!\n👤 ผู้โอน: ${payer}\n👤 ผู้รับ: ${participants[0]}\n💰 ยอดเงิน: ${amount} บาท\nℹ️ สถานะ: รอยืนยันการรับเงิน`;
      } else {
        notifyMsg = `💸 [กินกัน] มีรายการบันทึกค่าใช้จ่ายใหม่!\n🍔 มื้อ: ${meal}\n💰 ยอดรวม: ${amount} บาท\n👤 ผู้จ่าย: ${payer}\n👥 คนหาร: ${participants.join(', ')}\n💵 ตกคนละ: ${share} บาท`;
      }
      if (fileUrl) {
        notifyMsg += `\n📎 แนบหลักฐาน (สลิป): ${fileUrl}`;
      }
      sendLineNotify(notifyMsg, fileUrl);

      // Send direct push notification to the recipient of the settlement
      if (category === 'Settlement' && participants && participants.length > 0) {
        const recipientName = participants[0];
        const recipientLineUserId = getLineUserIdByFriendName(recipientName);
        if (recipientLineUserId) {
          let pushMsg = `💸 [กินกัน] มีการแจ้งโอนเงินคืนให้คุณ!\n👤 ผู้โอน: ${payer}\n💰 ยอดเงิน: ${amount} บาท\nℹ️ สถานะ: รอยืนยันการรับเงิน`;
          if (fileUrl) {
            pushMsg += `\n📎 แนบหลักฐาน (สลิป): ${fileUrl}`;
          }
          pushMsg += `\n\nกรุณาเข้าสู่ระบบเพื่อยืนยันยอดเงินค้างชำระ`;
          sendLinePush(recipientLineUserId, pushMsg, fileUrl);
        }
      }

      const updatedResult = getBalancesAndHistory(friendsList);
      return ContentService.createTextOutput(JSON.stringify(updatedResult))
        .setMimeType(ContentService.MimeType.JSON);
    }

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
