/**
 * CODE.GS - MORPEST API MODE
 * Untuk dijalankan di Vercel
 * Hanya return JSON, tidak serve HTML
 */

var REQUEST_SHEET_ID = '1pAOs2c68DkZfQvl18nzaPYHpFcEDDOWMFGn6kS1tNlU';
var REQUEST_SHEET_NAME = 'Incoming Requests';

function getRequestSpreadsheet() {
  return SpreadsheetApp.openById(REQUEST_SHEET_ID);
}

function getRequestSheet() {
  return getRequestSpreadsheet().getSheetByName(REQUEST_SHEET_NAME);
}

// ============================================
// BUG FIX: Menambahkan doPost(e) agar bisa 
// menerima request POST dari Vercel/Frontend
// ============================================
function doPost(e) {
  try {
    var body = JSON.parse(e.postData.contents);
    var func = body.function;
    var params = body.parameters || [];
    var result;
    
    if (func === "processForm") {
      result = processForm(params[0]);
    } else if (func === "checkRequestStatus") {
      result = checkRequestStatus(params[0]);
    } else if (func === "getAdminData") {
      result = getAdminData();
    } else if (func === "getCMODashboardData") {
      result = getCMODashboardData();
    } else if (func === "getPICData") {
      result = getPICData();
    } else if (func === "updateTicketData") {
      result = updateTicketData(params[0], params[1], params[2], params[3], params[4]);
    } else if (func === "deleteRequest") {
      result = deleteRequest(params[0]);
    } else {
      result = { error: "Function not found" };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================
// API: Process Form Submission
// ============================================
function processForm(formObject) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(30000);
  } catch (e) {
    return { 
      status: "error", 
      message: "Server sedang sibuk, silakan coba lagi dalam beberapa detik." 
    };
  }
  
  try {
    var ss = getRequestSpreadsheet();
    var sheet = getRequestSheet();
    
    if (!sheet) {
      return { 
        status: "error", 
        message: "Sheet 'Incoming Requests' tidak ditemukan!" 
      };
    }

    var timestamp = new Date();
    var lastRow = sheet.getLastRow();
    
    // Validasi form
    if (!formObject.reqName || !formObject.reqNickname || !formObject.reqContact || !formObject.reqPosition) {
      return { status: "error", message: "Semua kolom wajib diisi!" };
    }
    
    // Generate Ticket ID (cari max ID)
    var newIdNumber = 1001;
    if (lastRow > 1) {
      var idValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var maxId = 1000;
      for (var i = 0; i < idValues.length; i++) {
        var idStr = idValues[i][0];
        if (idStr && typeof idStr === 'string' && idStr.indexOf("-") >= 0) {
          var num = parseInt(idStr.split("-")[1]);
          if (!isNaN(num) && num > maxId) {
            maxId = num;
          }
        }
      }
      newIdNumber = maxId + 1;
    }
    var ticketId = "MKT-" + newIdNumber;

    // Process request details berdasarkan type
    var jenisRequest = formObject.requestType || "-";
    var finalBrief = "";
    var formatOrPlatform = "-";
    var toneVoice = "-";
    var targetAudience = "-";
    var duration = "-";

    if (jenisRequest.indexOf("upload konten") >= 0) {
      finalBrief = formObject.smsArahan || "-";
      formatOrPlatform = formObject.smsPlatform || "-";
    } else if (jenisRequest.indexOf("Request Design") >= 0) {
      finalBrief = formObject.gdArahan || "-";
      formatOrPlatform = formObject.gdFormat || "-";
    } else if (jenisRequest.indexOf("Penulisan Caption") >= 0) {
      finalBrief = formObject.cwArahan || "-";
      toneVoice = formObject.cwTone || "-";
      targetAudience = formObject.cwAudience || "-";
    } else if (jenisRequest.indexOf("konten video") >= 0) {
      finalBrief = formObject.ccArahan || "-";
      formatOrPlatform = formObject.ccPlatform || "-";
      duration = formObject.ccDuration || "-";
    }

    // Create calendar link
    var deadlineDate = new Date(formObject.reqDeadline);
    var timeZone = ss.getSpreadsheetTimeZone();
    var dateString = Utilities.formatDate(deadlineDate, timeZone, "yyyyMMdd");
    
    var calendarLink = "https://www.google.com/calendar/render?action=TEMPLATE" +
      "&text=" + encodeURIComponent("[Deadline] " + (formObject.reqTitle || "Marketing Request")) +
      "&dates=" + dateString + "/" + dateString +
      "&details=" + encodeURIComponent("Ticket ID: " + ticketId + "\nJenis: " + jenisRequest + "\n\n" + finalBrief) +
      "&location=Job On Yours";

    // Prepare row data
    var rowData = [
      ticketId,
      timestamp,
      formObject.reqName || "-",
      formObject.reqNickname || "-",
      formObject.reqContact || "-",
      formObject.reqPosition || "-",
      formObject.reqPriority || "-",
      jenisRequest,
      formObject.reqDeadline || "-",
      formObject.reqTitle || "-",
      finalBrief,
      formObject.reqLink || "-",
      formatOrPlatform,
      toneVoice,
      targetAudience,
      duration,
      "",
      "Pending",
      "",
      "",
      "",
      ""
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();
    
    return { 
      status: "success", 
      ticketId: ticketId, 
      calendarLink: calendarLink 
    };

  } catch (e) {
    return { 
      status: "error", 
      message: "Terjadi kesalahan sistem: " + e.toString() 
    };
  } finally {
    lock.releaseLock();
  }
}

// ============================================
// API: Check Request Status
// ============================================
function checkRequestStatus(ticketId) {
  try {
    var ss = getRequestSpreadsheet();
    var sheet = getRequestSheet();
    
    if (!sheet) {
      return { 
        found: false, 
        message: "Sheet tidak ditemukan" 
      };
    }
    
    var data = sheet.getDataRange().getValues();

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().toLowerCase() === ticketId.toString().toLowerCase()) {
        return {
          found: true,
          ticketId: data[i][0],
          judul: data[i][9] || "-",
          jenis: data[i][7] || "-",
          status: data[i][17] || "Pending",
          requester: data[i][2] || "-",
          reqContact: data[i][4] || "-",
          pic: data[i][18] || "Belum di-assign",
          deadline: Utilities.formatDate(new Date(data[i][8]), ss.getSpreadsheetTimeZone(), "dd MMM yyyy"),
          linkResult: data[i][20] || ""
        };
      }
    }
    
    return { 
      found: false,
      message: "Ticket ID tidak ditemukan"
    };
    
  } catch (e) {
    return { 
      found: false, 
      message: "Error: " + e.toString() 
    };
  }
}

// ============================================
// API: Get Admin Data
// ============================================
function getAdminData() {
  try {
    var ss = getRequestSpreadsheet();
    var sheet = getRequestSheet();
    
    if (!sheet) return { error: "Sheet tidak ditemukan" };
    
    var data = sheet.getDataRange().getValues();
    var tz = ss.getSpreadsheetTimeZone();
    var result = [];
    
    for (var i = 1; i < data.length; i++) {
      if (!data[i][0] || data[i][0].toString().trim() === "") continue;
      
      var deadlineDate = data[i][8] ? new Date(data[i][8]) : null;
      var status = data[i][17] || "Pending";
      var countdown = "";
      
      if (deadlineDate) {
        var today = new Date();
        today.setHours(0,0,0,0);
        var target = new Date(deadlineDate);
        target.setHours(0,0,0,0);
        var diffDays = Math.round((target - today) / (1000 * 60 * 60 * 24));
        if (status === "Done") {
          countdown = "Done";
        } else if (diffDays < 0) {
          countdown = "TERLEWAT " + Math.abs(diffDays) + " Hari";
        } else if (diffDays === 0) {
          countdown = "HARI INI";
        } else {
          countdown = diffDays + " Hari lagi";
        }
      }
      
      var submittedAt = "";
      if (data[i][1]) {
         try {
           submittedAt = Utilities.formatDate(new Date(data[i][1]), tz, "dd MMM yyyy, HH:mm");
         } catch(e) {}
      }
      var fullDeadline = "";
      if (deadlineDate) {
         try {
           fullDeadline = Utilities.formatDate(deadlineDate, tz, "dd MMM yyyy");
         } catch(e) {}
      }
      
      result.push({
        id: data[i][0],
        timestamp: data[i][1],
        submittedAt: submittedAt,
        requester: data[i][2] || "-",
        nickname: data[i][3],
        reqContact: data[i][4] || "-",
        divisi: data[i][5] || "-",
        priority: data[i][6],
        jenis: data[i][7] || "-",
        deadline: data[i][8],
        fullDeadline: fullDeadline,
        judul: data[i][9] || "-",
        brief: data[i][10],
        linkRef: data[i][11],
        format: data[i][12],
        tone: data[i][13],
        audience: data[i][14],
        duration: data[i][15],
        status: status,
        picName: data[i][18] || "-",
        picContact: data[i][19] || "-",
        finalLink: data[i][20] || "",
        countdown: countdown
      });
    }
    
    return { status: "success", data: result };
    
  } catch(e) {
    return { error: e.toString() };
  }
}

// ============================================
// API: Get PIC Data
// ============================================
function getPICData() {
  try {
    var ss = getRequestSpreadsheet();
    var sheets = ss.getSheets();
    var picSheet = null;
    for (var i=0; i<sheets.length; i++) {
      var name = sheets[i].getName().toLowerCase();
      if (name.indexOf("karyawan") >= 0 || name.indexOf("pic") >= 0 || name.indexOf("database") >= 0) {
        picSheet = sheets[i];
        break;
      }
    }
    
    var result = { sms: [], gd: [], cw: [], cc: [], cmo: [] };
    
    if (picSheet) {
      var data = picSheet.getDataRange().getValues();
      if (data.length > 0) {
        var headers = data[0];
        var nameIdx = 0, nickIdx = 1, divIdx = 2, phoneIdx = 3;
        for (var j=0; j<headers.length; j++) {
          var h = headers[j].toString().toLowerCase();
          if (h.indexOf("nama") >= 0 && h.indexOf("lengkap") >= 0) nameIdx = j;
          else if (h.indexOf("panggilan") >= 0 || h.indexOf("nick") >= 0) nickIdx = j;
          else if (h.indexOf("divisi") >= 0 || h.indexOf("posisi") >= 0) divIdx = j;
          else if (h.indexOf("wa") >= 0 || h.indexOf("hp") >= 0 || h.indexOf("kontak") >= 0) phoneIdx = j;
        }
        
        for (var k=1; k<data.length; k++) {
          if (!data[k][nameIdx]) continue;
          var picObj = {
            name: data[k][nameIdx] || "-",
            nickname: data[k][nickIdx] || "-",
            phone: data[k][phoneIdx] || "-"
          };
          var div = (data[k][divIdx] || "").toString().toLowerCase();
          
          if (div.indexOf("cmo") >= 0) result.cmo.push(picObj);
          if (div.indexOf("sms") >= 0) result.sms.push(picObj);
          if (div.indexOf("gd") >= 0 || div.indexOf("desain") >= 0 || div.indexOf("design") >= 0) result.gd.push(picObj);
          if (div.indexOf("cw") >= 0 || div.indexOf("copy") >= 0 || div.indexOf("caption") >= 0) result.cw.push(picObj);
          if (div.indexOf("cc") >= 0 || div.indexOf("video") >= 0 || div.indexOf("content") >= 0) result.cc.push(picObj);
        }
      }
    }
    
    // Fallback if empty
    if (result.sms.length === 0 && result.gd.length === 0 && result.cw.length === 0 && result.cc.length === 0) {
       var fallback = {name: "Morpest Admin", nickname: "Admin", phone: "6285172283505"};
       result.sms.push(fallback); result.gd.push(fallback); result.cw.push(fallback); result.cc.push(fallback);
    }
    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}

// ============================================
// API: Update Ticket Data
// ============================================
function updateTicketData(ticketId, picName, picContact, status, finalLink) {
  try {
    var sheet = getRequestSheet();
    if (!sheet) return "Error: Sheet not found";
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === ticketId) {
        if (picName) sheet.getRange(i + 1, 19).setValue(picName);
        if (picContact) sheet.getRange(i + 1, 20).setValue(picContact);
        if (status) sheet.getRange(i + 1, 18).setValue(status);
        if (finalLink !== undefined && finalLink !== null && finalLink !== '') {
           sheet.getRange(i + 1, 21).setValue(finalLink);
           sheet.getRange(i + 1, 18).setValue("Done");
        }
        return "Success";
      }
    }
    return "Error: Ticket not found";
  } catch(e) {
    return "Error: " + e.toString();
  }
}

// ============================================
// API: Delete Request
// ============================================
function deleteRequest(ticketId) {
  try {
    var sheet = getRequestSheet();
    if (!sheet) return { success: false, message: "Sheet not found" };
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString() === ticketId) {
        sheet.deleteRow(i + 1);
        return { success: true };
      }
    }
    return { success: false, message: "Ticket not found" };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

// ============================================
// API: Get CMO Dashboard Data
// ============================================
function getCMODashboardData() {
  try {
    var ss = getRequestSpreadsheet();
    var sheet = getRequestSheet();
    if (!sheet) return { error: "Sheet tidak ditemukan" };
    
    var data = sheet.getDataRange().getValues();
    var tz = ss.getSpreadsheetTimeZone();
    
    var totalRequests = 0;
    var totalDone = 0;
    var totalOnProcess = 0;
    var totalPending = 0;
    var totalOverdue = 0;
    var byType = { "SMS": 0, "GD": 0, "CW": 0, "CC": 0 };
    var byMonth = {};
    var picPerf = {};
    var urgentItems = [];
    var today = new Date();
    today.setHours(0, 0, 0, 0);
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row[0] || row[0].toString().trim() === "") continue;
      
      totalRequests++;
      var status = (row[17] || "Pending").toString().trim();
      var jenis = (row[7] || "").toString();
      var picName = (row[18] || "").toString().trim();
      
      if (status === "Done") totalDone++;
      else if (status === "On Process") totalOnProcess++;
      else totalPending++;
      
      var rawDeadline = row[8];
      var deadlineDate = null;
      if (rawDeadline) {
        deadlineDate = (rawDeadline instanceof Date) ? rawDeadline : new Date(rawDeadline);
        deadlineDate.setHours(0, 0, 0, 0);
        var diffDays = Math.round((deadlineDate - today) / (1000 * 60 * 60 * 24));
        if (diffDays < 0 && status !== "Done") totalOverdue++;
      }
      
      if (jenis.indexOf("SMS") >= 0 || jenis.indexOf("upload konten") >= 0) byType.SMS++;
      else if (jenis.indexOf("GD") >= 0 || jenis.indexOf("Design") >= 0) byType.GD++;
      else if (jenis.indexOf("CW") >= 0 || jenis.indexOf("Caption") >= 0) byType.CW++;
      else if (jenis.indexOf("CC") >= 0 || jenis.indexOf("konten video") >= 0) byType.CC++;
      
      var rawTimestamp = row[1];
      if (rawTimestamp) {
        var ts = (rawTimestamp instanceof Date) ? rawTimestamp : new Date(rawTimestamp);
        if (!isNaN(ts.getTime())) {
          var monthKey = Utilities.formatDate(ts, tz, "yyyy-MM");
          byMonth[monthKey] = (byMonth[monthKey] || 0) + 1;
        }
      }
      
      if (picName) {
        if (!picPerf[picName]) picPerf[picName] = { done: 0, total: 0 };
        picPerf[picName].total++;
        if (status === "Done") picPerf[picName].done++;
      }
    }
    
    var monthLabels = [];
    var monthValues = [];
    var monthKeys = Object.keys(byMonth).sort();
    for (var m = 0; m < monthKeys.length; m++) {
      monthLabels.push(monthKeys[m]);
      monthValues.push(byMonth[monthKeys[m]]);
    }
    
    var picPerfArray = [];
    for (var pn in picPerf) {
      picPerfArray.push({ name: pn, done: picPerf[pn].done, total: picPerf[pn].total });
    }
    picPerfArray.sort(function(a, b) { return b.total - a.total; });
    var topPIC = picPerfArray.slice(0, 10);
    
    return {
      summary: {
        total: totalRequests,
        done: totalDone,
        onProcess: totalOnProcess,
        pending: totalPending,
        overdue: totalOverdue,
        doneRate: totalRequests > 0 ? Math.round((totalDone / totalRequests) * 100) : 0
      },
      byType: byType,
      byMonth: {
        labels: monthLabels,
        values: monthValues
      },
      topPIC: topPIC
    };
    
  } catch(e) {
    return { error: e.toString() };
  }
}
