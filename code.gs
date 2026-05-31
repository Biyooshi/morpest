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
// SETUP & AUTHORIZATION
// ============================================
// PENTING: Jalankan fungsi ini SATU KALI dari editor Apps Script (Run -> setupEmailAuth)
// Ini diperlukan untuk memberikan izin (Otorisasi) agar script bisa mengirim email.
function setupEmailAuth() {
  MailApp.sendEmail(Session.getActiveUser().getEmail(), "Morpest Email Setup", "Otorisasi pengiriman email telah berhasil diaktifkan!");
  Logger.log("Otorisasi email berhasil.");
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
      formObject.reqEmail || "-",
      "Pending",
      "",
      "",
      "",
      ""
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();
    
    // --- SEND NOTIFICATIONS ---
    var adminMsg = "*NEW REQUEST MORPEST*\n\nTicket ID: " + ticketId + "\nDari: " + (formObject.reqName || "-") + " (" + (formObject.reqPosition || "-") + ")\nJudul: " + (formObject.reqTitle || "-") + "\nJenis: " + jenisRequest + "\nDeadline: " + (formObject.reqDeadline || "-") + "\n\nHarap segera di-assign ke PIC yang bertugas.";
    sendWhatsAppMessage("6285233142178", adminMsg); // Admin Obi
    sendEmailNotification("qolbimuhammad00@gmail.com", "New Request: " + ticketId, adminMsg.replace(/\n/g, '<br>')); // CMO gets explicit admin email
    
    var reqMsg = "*REQUEST BERHASIL DIBUAT*\n\nHalo " + (formObject.reqNickname || "-") + ",\nRequest kamu dengan judul *" + (formObject.reqTitle || "-") + "* telah diterima.\nTicket ID: *" + ticketId + "*\n\nKamu dapat mengecek status requestmu di halaman Request Status Checker. Terima kasih!\n_Job On Yours Marketing_";
    sendWhatsAppMessage(formObject.reqContact, reqMsg);
    if (formObject.reqEmail) {
      sendEmailNotification(formObject.reqEmail, "Request Berhasil Dibuat: " + ticketId, reqMsg.replace(/\n/g, '<br>'), "qolbimuhammad00@gmail.com");
    }

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
    
    result.reverse();
    return { status: "success", data: result };
    
  } catch(e) {
    return { error: e.toString() };
  }
}

// ============================================
// API: Get PIC Data
// ============================================
function getPICData() {
  var picData = {
    "sms": [
      { name: "Riska Stephanie", nickname: "Ika", phone: "62895352730008", email: "riskastphnie28@gmail.com" },
      { name: "Naya Azani", nickname: "Naya", phone: "62895391527014", email: "nayaazani13@gmail.com" },
      { name: "Meitia Kurniawati", nickname: "Mei", phone: "6282247765358", email: "meitiakurniawati@gmail.com" },
      { name: "Muhamad Farhan", nickname: "Aan", phone: "6289531974196", email: "aan.hans12@gmail.com" },
      { name: "Kladya Khoirunisa' Hapsari", nickname: "Kladya", phone: "6285645111043", email: "kladyakhoirunisakh@gmail.com" },
      { name: "Shellbitav Azazel", nickname: "Shelby", phone: "6285172283505", email: "ch3llb14z@gmail.com" }
    ],
    "gd": [
      { name: "Head Graphic Designer", nickname: "Head-GD", phone: "6280808080808", email: "headgd@gmail.com" },
      { name: "Co-Head Graphic Designer", nickname: "Co-Head-GD", phone: "6280808080808", email: "coheadgd@gmai.com" },
      { name: "Yusmita Alya Melanie", nickname: "Alya", phone: "6285792300256", email: "yusmitaalya@gmail.com" },
      { name: "Syafa Salsabila", nickname: "Syafa", phone: "62882003120378", email: "syafasalsabila226@gmail.com" },
      { name: "Lidia Siregar", nickname: "Lidia", phone: "6281316207014", email: "kembaroktober06@gmail.com" },
      { name: "Mulqy Azzam", nickname: "Azzam", phone: "6281770094378", email: "mulqyazzam41@gmail.com" }
    ],
    "cw": [
      { name: "Sidik Permana", nickname: "Sidik", phone: "6285321200416", email: "sidiksipengelana@gmail.com" },
      { name: "Ridatasa Nadiawati", nickname: "Rida", phone: "6289656144248", email: "Ridatasa@gmail.com" },
      { name: "Sevia Rahmadani", nickname: "Via", phone: "6285709598764", email: "seviarahmadani9@gmail.com" },
      { name: "Mayang Anggraini", nickname: "Mayang", phone: "628812756505", email: "mayanganggraini242@gmail.com" },
      { name: "Sri Rahayu Mulyaningsih", nickname: "Ayu", phone: "62895414845637", email: "srirahayuuu2937@gmail.com" },
      { name: "Benedict Jemima Cecilia Pietersz", nickname: "Ben", phone: "6282114887824", email: "bjcpietersz47@gmail.com" }
    ],
    "cc": [],
    "cmo": [
      { name: "Muhammad Nurul Qolbi", nickname: "Obi", phone: "6285233142178", email: "qolbi@joy.internal" }
    ]
  };
  
  picData.legacy = getLegacyPICData();
  return picData;
}

function getLegacyPICData() {
  var picData = {
    "sms": [
      { name: "Dio", nickname: "Dio", phone: "6285814783478", email: "dio@joy.internal" },
      { name: "Muslihah (Biru)", nickname: "Biru", phone: "6285776786184", email: "biru@joy.internal" },
      { name: "Juhariah Ningrum", nickname: "Juju", phone: "6285251549269", email: "juju@joy.internal" },
      { name: "Nichell", nickname: "Nichell", phone: "601123362713", email: "nichell@joy.internal" },
      { name: "Sauma Wulandari", nickname: "Sauma", phone: "6285174076340", email: "sauma@joy.internal" },
      { name: "Zahra Sri Rahmah", nickname: "Zahra", phone: "6282321454263", email: "zahra@joy.internal" },
      { name: "Naili Zumna", nickname: "Naili", phone: "6285924846577", email: "naili@joy.internal" }
    ],
    "gd": [
      { name: "Naufal Zuhri", nickname: "Nopal", phone: "6281331189718", email: "naufal.zuhri@joy.internal" },
      { name: "Shafni Virginia Putri", nickname: "Shafni", phone: "6281394759578", email: "shafni@joy.internal" },
      { name: "Nuri Ibadi Rahmania", nickname: "Nuri", phone: "628115232805", email: "nuri@joy.internal" },
      { name: "Rosdiana Dewi", nickname: "Rosdiana", phone: "6282228190006", email: "rosdiana@joy.internal" },
      { name: "Shellbiav Azazel", nickname: "Shelby", phone: "6285172283505", email: "shelby@joy.internal" },
      { name: "Naufal Zuhdi", nickname: "Nopal", phone: "6281242528932", email: "naufal.zuhdi@joy.internal" },
      { name: "Bayu Pratama", nickname: "Bayu", phone: "6282111742892", email: "bayu@joy.internal" }
    ],
    "cw": [
      { name: "Astri Kania Dewi", nickname: "Astri", phone: "6281220517281", email: "astri@joy.internal" },
      { name: "Afifah Fauziah", nickname: "Fifah", phone: "6282268973187", email: "fifahzyelianna@gmail.com" },
      { name: "Intan Nurul Fasha", nickname: "Asha", phone: "6289501941526", email: "asha@joy.internal" },
      { name: "Nadiyah Sabilah Bintang", nickname: "Nadiya", phone: "6282283822925", email: "nadiya@joy.internal" },
      { name: "Klosse Ignatius Siringo Ringo", nickname: "Klosse", phone: "6281290141232", email: "klosse@joy.internal" },
      { name: "Danis Ahnaf", nickname: "Danis", phone: "6285941137220", email: "danis@joy.internal" },
      { name: "R. AJ. Afra Aurelia Luqyandysa", nickname: "Afra", phone: "6281213403556", email: "afra@joy.internal" }
    ],
    "cc": [
      { name: "Refan Regika Renggo", nickname: "Refan", phone: "6285697039805", email: "refan@joy.internal" },
      { name: "Desy fitriyanti", nickname: "Desy", phone: "6289630585603", email: "desy@joy.internal" },
      { name: "Caitlin Gabriella Rahardjan", nickname: "Caitlin", phone: "62818862341", email: "caitlin@joy.internal" },
      { name: "Falisha Azzahra Dinarsanti", nickname: "Fale", phone: "6285157341440", email: "fale@joy.internal" },
      { name: "Mia Kultsum Safitri", nickname: "Mia", phone: "6288980237068", email: "mia@joy.internal" },
      { name: "Qonita Tsaltsa Earlyana", nickname: "Qonita", phone: "6285156903647", email: "qonita@joy.internal" }
    ]
  };
  
  return picData;
}

function getPICInfo(nameToFind) {
  var active = getPICData();
  var legacy = getLegacyPICData();
  var allData = [active, legacy];
  
  for (var i = 0; i < allData.length; i++) {
    for (var div in allData[i]) {
       var arr = allData[i][div];
       for (var j = 0; j < arr.length; j++) {
         if (arr[j].name === nameToFind) return arr[j];
       }
    }
  }
  return null;
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
        var oldStatus = data[i][17];
        var reqName = data[i][2];
        var reqContact = data[i][4];
        var reqTitle = data[i][9];
        var reqEmail = data[i][16]; // Kolom 17
        
        var picVal = picName !== undefined ? picName : data[i][18];
        var finalLinkVal = finalLink !== undefined ? finalLink : data[i][20];
        
        var newStatus = "Pending";
        if (finalLinkVal && finalLinkVal.toString().trim() !== '') {
           newStatus = "Done";
        } else if (picVal && picVal.toString().trim() !== '') {
           newStatus = "On Process";
        }
        
        if (picName !== undefined) sheet.getRange(i + 1, 19).setValue(picName);
        if (picContact !== undefined) sheet.getRange(i + 1, 20).setValue(picContact);
        sheet.getRange(i + 1, 18).setValue(newStatus);
        if (finalLinkVal !== undefined) sheet.getRange(i + 1, 21).setValue(finalLinkVal);
        
        if (newStatus === "Done" && oldStatus !== "Done") {
           sheet.getRange(i + 1, 22).setValue(new Date()); // Save Completed At
        }
        
        // Notify Requester and Admin on status change
        if (newStatus !== oldStatus && newStatus === "Done") {
           var doneMsg = "*REQUEST SELESAI*\n\nHalo " + reqName + ",\nRequest kamu *" + reqTitle + "* (Ticket: " + ticketId + ") telah selesai dikerjakan!\n\nLink Hasil: " + (finalLinkVal || "-") + "\n\nTerima kasih,\n_Job On Yours Marketing_";
           sendWhatsAppMessage(reqContact, doneMsg);
           if (reqEmail && reqEmail !== "-") {
             sendEmailNotification(reqEmail, "Request Selesai: " + ticketId, doneMsg.replace(/\n/g, '<br>'), "qolbimuhammad00@gmail.com");
           } else {
             sendEmailNotification("qolbimuhammad00@gmail.com", "Request Selesai: " + ticketId, doneMsg.replace(/\n/g, '<br>'));
           }
           
           if (picVal) {
              var picInfo = getPICInfo(picVal);
              if (picInfo && picInfo.email) {
                 var picMsg = "Halo " + picInfo.nickname + ",\nTugas *" + reqTitle + "* (Ticket: " + ticketId + ") telah ditandai SELESAI.";
                 sendEmailNotification(picInfo.email, "Tugas Selesai: " + ticketId, picMsg.replace(/\n/g, '<br>'), "qolbimuhammad00@gmail.com");
              }
           }
        } else if (newStatus !== oldStatus && newStatus === "On Process") {
           var procMsg = "*REQUEST DIPROSES*\n\nHalo " + reqName + ",\nRequest kamu *" + reqTitle + "* (Ticket: " + ticketId + ") saat ini sedang dikerjakan oleh PIC: *" + picVal + "*.\n\nHarap ditunggu hasilnya!\n_Job On Yours Marketing_";
           sendWhatsAppMessage(reqContact, procMsg);
           if (reqEmail && reqEmail !== "-") {
             sendEmailNotification(reqEmail, "Request Diproses: " + ticketId, procMsg.replace(/\n/g, '<br>'), "qolbimuhammad00@gmail.com");
           } else {
             sendEmailNotification("qolbimuhammad00@gmail.com", "Request Diproses: " + ticketId, procMsg.replace(/\n/g, '<br>'));
           }
           
           if (picVal) {
              var picInfo = getPICInfo(picVal);
              if (picInfo) {
                 var picMsg = "Halo " + picInfo.nickname + ",\nKamu telah ditugaskan untuk mengerjakan *" + reqTitle + "* (Ticket: " + ticketId + ").\nMohon segera dikerjakan sesuai deadline.";
                 sendWhatsAppMessage(picInfo.phone, picMsg);
                 if (picInfo.email) {
                    sendEmailNotification(picInfo.email, "Tugas Baru: " + ticketId, picMsg.replace(/\n/g, '<br>'), "qolbimuhammad00@gmail.com");
                 }
              }
           }
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
    var totalLeadTimeDays = 0;
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
      
      if (status === "Done") {
        totalDone++;
        var startTs = (row[1] instanceof Date) ? row[1] : new Date(row[1]);
        var endTs = row[21] ? ((row[21] instanceof Date) ? row[21] : new Date(row[21])) : new Date(); // Col 22 is index 21
        if (!isNaN(startTs.getTime()) && !isNaN(endTs.getTime())) {
            var diff = Math.round((endTs - startTs) / (1000 * 60 * 60 * 24));
            if (diff < 0) diff = 0;
            totalLeadTimeDays += diff;
        }
      }
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
    
    // WORKLOAD & URGENT ITEMS
    var activePICs = getPICData();
    var legacyPICs = getLegacyPICData();
    
    var allActivePICs = [];
    for (var div in activePICs) {
       if (!Array.isArray(activePICs[div])) continue;
       activePICs[div].forEach(function(p) { allActivePICs.push({name: p.name, activeCount: 0, doneCount: 0, totalCount: 0, activeTickets: []}); });
    }
    
    var allLegacyPICs = [];
    for (var div in legacyPICs) {
       legacyPICs[div].forEach(function(p) { allLegacyPICs.push({name: p.name, activeCount: 0, doneCount: 0, totalCount: 0, activeTickets: []}); });
    }

    for (var i = 1; i < data.length; i++) {
        var row = data[i];
        if (!row[0] || row[0].toString().trim() === "") continue;
        
        var status = (row[17] || "Pending").toString().trim();
        var picName = (row[18] || "").toString().trim();
        var priority = (row[6] || "").toString();
        
        // Populate workloads
        if (picName) {
            var foundActive = null;
            for (var j = 0; j < allActivePICs.length; j++) {
                if (allActivePICs[j].name === picName) { foundActive = allActivePICs[j]; break; }
            }
            if (!foundActive) {
                var foundLegacy = null;
                for (var j = 0; j < allLegacyPICs.length; j++) {
                    if (allLegacyPICs[j].name === picName) { foundLegacy = allLegacyPICs[j]; break; }
                }
            }
            
            var targetPIC = foundActive || foundLegacy;
            if (targetPIC) {
                if (status === "On Process") {
                    targetPIC.activeCount++;
                    targetPIC.activeTickets.push(row[0]);
                } else if (status === "Done") {
                    targetPIC.doneCount++;
                }
                targetPIC.totalCount = targetPIC.activeCount + targetPIC.doneCount;
            }
        }

        // Populate urgent items
        var rawDeadline = row[8];
        if (status !== "Done" && rawDeadline) {
            var deadlineDate = (rawDeadline instanceof Date) ? rawDeadline : new Date(rawDeadline);
            deadlineDate.setHours(0, 0, 0, 0);
            var diffDays = Math.round((deadlineDate - today) / (1000 * 60 * 60 * 24));
            
            if (diffDays <= 1 || priority.indexOf("Urgent") >= 0) {
                var fullDeadline = Utilities.formatDate(deadlineDate, tz, "dd MMM yyyy");
                urgentItems.push({
                    id: row[0],
                    judul: row[9] || "-",
                    picName: picName || "-",
                    deadline: fullDeadline,
                    status: status,
                    diffDays: diffDays
                });
            }
        }
    }
    
    // Sort workloads descending by totalCount
    allActivePICs.sort(function(a,b) { return b.totalCount - a.totalCount; });
    allLegacyPICs.sort(function(a,b) { return b.totalCount - a.totalCount; });
    urgentItems.sort(function(a,b) { return a.diffDays - b.diffDays; });

    return {
      summary: {
        total: totalRequests,
        done: totalDone,
        onProcess: totalOnProcess,
        pending: totalPending,
        overdue: totalOverdue,
        doneRate: totalRequests > 0 ? Math.round((totalDone / totalRequests) * 100) : 0,
        avgCompletionTime: totalDone > 0 ? (totalLeadTimeDays / totalDone).toFixed(1) : 0
      },
      byType: byType,
      byMonth: {
        labels: monthLabels,
        values: monthValues
      },
      topPIC: topPIC,
      workloads: {
        active: allActivePICs,
        legacy: allLegacyPICs
      },
      urgentItems: urgentItems
    };
    
  } catch(e) {
    return { error: e.toString() };
  }
}

// ============================================
// API: Notifications (Email & WhatsApp)
// ============================================
var FONNTE_TOKEN = '9PkBs4SoEG15Qbw8mVBd';

function sendEmailNotification(to, subject, bodyHtml, ccEmail) {
  try {
    var options = {
      to: to,
      subject: subject,
      htmlBody: bodyHtml,
      name: "Morpest JOY"
    };
    if (ccEmail) {
      options.cc = ccEmail;
    }
    MailApp.sendEmail(options);
    return true;
  } catch(e) {
    Logger.log("Email Error: " + e.toString());
    return false;
  }
}

function sendWhatsAppMessage(targetNumber, message) {
  if (!targetNumber || targetNumber === "-") return false;
  try {
    var cleanNumber = targetNumber.toString().replace(/\D/g, '');
    if (cleanNumber.startsWith('0')) {
      cleanNumber = '62' + cleanNumber.substring(1);
    }
    
    var url = "https://api.fonnte.com/send";
    var payloadObj = {
      "target": cleanNumber,
      "message": message,
      "delay": "2"
    };
    
    var options = {
      "method": "post",
      "headers": {
        "Authorization": FONNTE_TOKEN,
        "Content-Type": "application/json"
      },
      "payload": JSON.stringify(payloadObj),
      "muteHttpExceptions": true
    };
    
    var response = UrlFetchApp.fetch(url, options);
    Logger.log("WA Response: " + response.getContentText());
    
    return true;
  } catch(e) {
    Logger.log("WA Error: " + e.toString());
    return false;
  }
}
