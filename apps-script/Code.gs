/**
 * Food Line — серверна частина (Google Apps Script)
 * V10 — PDF на Диск / PDF на пошту / налаштування у вкладці таблиці
 *       + СПІЛЬНА БАЗА ЧЕРНЕТОК: інструкцію можна дописувати з будь-якого
 *         телефону, а фото не вивантажуються повторно.
 *
 * ПІСЛЯ ВСТАВКИ ЦЬОГО КОДУ ОБОВ'ЯЗКОВО:
 * 1) Запустити один раз функцію setupPermissions() і дати дозволи.
 * 2) Розгорнути → Керувати розгортаннями → Редагувати (олівець)
 *    → Версія: НОВА → Розгорнути.
 *    Виконувати як: Я.  Хто має доступ: Будь-хто (Anyone).
 * 3) URL розгортання вставити в index.html у GOOGLE_SCRIPT_URL.
 *
 * Вкладки «Налаштування» і «Чернетки» створюються самі при першому запиті.
 */

// ─────────────────────────────────────────────────────────────
// НАЛАШТУВАННЯ
// ─────────────────────────────────────────────────────────────

// Таблиця-журнал. Залиште "" (порожньо), щоб узагалі не писати в таблицю.
// УВАГА: для налаштувань і для спільної бази чернеток цей ID обов'язковий.
var SPREADSHEET_ID = "1vzAYros4jLTQFZQR-QHko8WvdB8vEhVkquNQ_Khn5tw";

// Назви службових вкладок.
var SETTINGS_SHEET_NAME = "Налаштування";
var DRAFTS_SHEET_NAME   = "Чернетки";

// Назва папки, у якій живуть чернетки, якщо в налаштуваннях не вказано свою.
var DRAFTS_FOLDER_NAME = "Чернетки Food Line";

// Мінімальний розмір коректного PDF (байт). Менше — вважаємо порожнім аркушем.
var MIN_PDF_BYTES = 5000;

// true — кожному завантаженому фото відкривається доступ "усім, хто має посилання".
// Потрібно, якщо таблицю переглядають люди без доступу до папки з фото.
var MAKE_PHOTOS_LINK_VIEWABLE = false;

// Скільки байтів фото максимум віддаємо в одній відповіді (ліміт Apps Script ~10 МБ).
var MAX_DOWNLOAD_BYTES = 6 * 1024 * 1024;

// Простий пароль додатку. Поки "" — перевірки немає (як було раніше).
// Якщо заповните — той самий рядок треба вписати в index.html у APP_TOKEN.
// УВАГА: це лише бар'єр від випадкових запитів, а не справжня безпека:
// index.html лежить публічно, тому токен з нього можна вичитати.
// Справжнє обмеження — розгортання «Тільки для мене / моєї організації».
var ACCESS_TOKEN = "";

// ─────────────────────────────────────────────────────────────
// ТОЧКИ ВХОДУ
// ─────────────────────────────────────────────────────────────

function doGet(e) {
  // Дозволяємо прочитати налаштування і простим GET-запитом:
  //   ...exec?action=get_settings
  if (e && e.parameter && e.parameter.action === "get_settings") {
    try {
      checkToken(e.parameter);
      return json(getSettings());
    } catch (err) {
      return json({ status: "error", message: err.message });
    }
  }

  return json({
    status: "success",
    message: "Food Line script V10 працює. Час сервера: " +
             Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy HH:mm")
  });
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error("Порожній запит. Перевірте, що розгортання виконано як веб-додаток з доступом «Будь-хто».");
    }

    var data = JSON.parse(e.postData.contents);
    checkToken(data);

    switch (data.action) {
      // готовий документ
      case "send_report":         return json(sendReport(data));
      case "upload_photos":       return json(uploadPhotos(data));
      case "save_instruction":    return json(saveInstruction(data));

      // налаштування
      case "get_settings":        return json(getSettings());
      case "save_settings":       return json(saveSettings(data));

      // спільна база чернеток
      case "draft_save":          return json(draftSave(data));
      case "draft_load":          return json(draftLoad(data));
      case "draft_list":          return json(draftList(data));
      case "draft_delete":        return json(draftDelete(data));
      case "draft_upload_photos": return json(draftUploadPhotos(data));
      case "draft_photos":        return json(draftPhotos(data));

      default:
        throw new Error("Невідома дія: " + data.action);
    }
  } catch (error) {
    return json({ status: "error", message: (error && error.message) ? error.message : String(error) });
  }
}

function checkToken(data) {
  if (!ACCESS_TOKEN) return;
  if (String((data && data.token) || "") !== ACCESS_TOKEN) {
    throw new Error("Немає доступу: невірний пароль додатку. Звірте ACCESS_TOKEN у скрипті та APP_TOKEN в index.html.");
  }
}

// ─────────────────────────────────────────────────────────────
// ДІЯ 1: PDF НА ПОШТУ
// ─────────────────────────────────────────────────────────────

function sendReport(data) {
  var email = (data.email || "").trim();
  if (!email) throw new Error("Не вказано Email у налаштуваннях додатку.");
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error("Email виглядає некоректним: " + email);

  var blob = pdfBlobFrom(data);

  // Захист від дурної несподіванки: квота Gmail вичерпана
  var quota = MailApp.getRemainingDailyQuota();
  if (quota <= 0) throw new Error("Вичерпано денну квоту на відправку пошти Google. Спробуйте завтра.");

  MailApp.sendEmail({
    to: email,
    subject: "Нова інструкція: " + (data.title || "Без назви"),
    body: "Добрий день!\n\n" +
          "У додатку Food Line згенеровано нову інструкцію.\n\n" +
          "Назва: " + (data.title || "—") + "\n" +
          "Обладнання: " + (data.equipment || "—") + "\n" +
          "Дата: " + (data.date || "—") + "\n\n" +
          "PDF у вкладенні.",
    attachments: [blob]
  });

  return { status: "success", message: "Звіт відправлено на " + email };
}

// ─────────────────────────────────────────────────────────────
// ДІЯ 2: ФОТО В ПАПКУ ДИСКУ (пачками по кілька штук)
// ─────────────────────────────────────────────────────────────

function uploadPhotos(data) {
  var photos = data.photos || [];
  if (!photos.length) return { status: "success", urls: {}, photoFolderId: data.photoFolderId || null };

  // Перша пачка створює підпапку під цю інструкцію, наступні — просто в неї пишуть
  var folder;
  if (data.photoFolderId) {
    try {
      folder = DriveApp.getFolderById(data.photoFolderId);
    } catch (e) {
      throw new Error("Втрачено доступ до підпапки з фото. Спробуйте зберегти ще раз.");
    }
  } else {
    var parent = resolveFolder(data.photoDriveId, "фото");
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH-mm");
    var folderName = stamp + " " + sanitizeName(data.title || "Інструкція", 60);
    folder = parent.createFolder(folderName);
  }

  var urls = {};
  for (var i = 0; i < photos.length; i++) {
    var p = photos[i];
    if (!p || !p.key) continue;

    var name = sanitizeName(p.name || ("photo_" + i + ".jpg"), 90);
    var file = null;

    try {
      if (p.copyFileId) {
        // Фото вже лежить на Диску (вивантажене як чернетка) — копіюємо
        // на сервері, замість того щоб качати мегабайти з телефона вдруге.
        file = DriveApp.getFileById(p.copyFileId).makeCopy(name, folder);
      } else if (p.dataUrl) {
        var blob = dataUrlToBlob(p.dataUrl, name);
        if (blob) file = folder.createFile(blob);
      }

      if (!file) continue;

      if (MAKE_PHOTOS_LINK_VIEWABLE) {
        try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (shareErr) {}
      }
      urls[p.key] = file.getUrl();
    } catch (err) {
      Logger.log("Фото не збереглось (" + name + "): " + err);
    }
  }

  return { status: "success", urls: urls, photoFolderId: folder.getId(), folderUrl: folder.getUrl() };
}

// ─────────────────────────────────────────────────────────────
// ДІЯ 3: PDF У ПАПКУ GOOGLE ДИСКУ
// ─────────────────────────────────────────────────────────────

function saveInstruction(data) {
  var blob = pdfBlobFrom(data);
  var folder = resolveFolder(data.driveId, "PDF");

  // ГОЛОВНЕ — файл на Диску. Робимо це ПЕРШИМ.
  var file;
  try {
    file = folder.createFile(blob);
  } catch (err) {
    throw new Error("Не вдалося записати файл у папку. Перевірте, що акаунт-власник скрипта має доступ на редагування цієї папки. Деталі: " + err.message);
  }

  var message = "PDF збережено у папку «" + folder.getName() + "»";

  // Журнал у таблиці — другорядне. Якщо впаде, файл на Диску вже є,
  // тому помилку не піднімаємо, а лише дописуємо в повідомлення.
  if (SPREADSHEET_ID) {
    try {
      var sheetName = writeLogSheet(data, file.getUrl());
      message += ". Створено вкладку «" + sheetName + "» у таблиці";
    } catch (logErr) {
      message += ". (Увага: у таблицю не записалось — " + logErr.message + ")";
    }
  }

  // Чернетку позначаємо як завершену, щоб у спільній базі було видно,
  // що робота з неї вже перетворилась на готовий PDF.
  if (data.draftId) {
    try { markDraftDone(data.draftId, file.getUrl()); } catch (e) {
      Logger.log("Не вдалося позначити чернетку готовою: " + e);
    }
  }

  return { status: "success", message: message + ".", fileUrl: file.getUrl() };
}

// ─────────────────────────────────────────────────────────────
// ДІЇ 4-5: НАЛАШТУВАННЯ ДОДАТКУ У ВКЛАДЦІ ТАБЛИЦІ
// ─────────────────────────────────────────────────────────────

/** Рядки, які створюються у вкладці «Налаштування» при першому запуску. */
var SETTINGS_ROWS = [
  ["email",           "",     "Email для відправки звіту"],
  ["driveId",         "",     "Папка для PDF — ID або повне посилання"],
  ["photoDriveId",    "",     "Папка для фото — ID або повне посилання"],
  ["draftDriveId",    "",     "Папка для чернеток спільної бази (порожньо = створиться сама)"],
  ["singlePage",      "TRUE", "PDF одним суцільним аркушем: TRUE / FALSE"],
  ["cloudAutoSave",   "TRUE", "Автозапис у спільну базу: TRUE / FALSE"],
  ["cloudAutoSaveSec","120",  "Як часто автозапис, секунд (не менше 30)"],
  ["apiKey",          "",     "Ключ Gemini API"],
  ["aiModel",         "",     "Модель ШІ — заповнюється автоматично, руками не чіпати"]
];

/** Створює вкладку з ключами при першому зверненні. */
function ensureSettingsSheet() {
  if (!SPREADSHEET_ID) throw new Error("У скрипті не вказано SPREADSHEET_ID — нема де зберігати налаштування.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(SETTINGS_SHEET_NAME);
  if (sh) {
    ensureSettingsRows(sh);
    return sh;
  }

  sh = ss.insertSheet(SETTINGS_SHEET_NAME, 0);
  sh.setColumnWidth(1, 150);
  sh.setColumnWidth(2, 430);
  sh.setColumnWidth(3, 360);

  sh.getRange("A1:C1")
    .setValues([["Ключ", "Значення", "Що це"]])
    .setFontWeight("bold")
    .setBackground("#e2e8f0");

  sh.getRange(2, 1, SETTINGS_ROWS.length, 3).setValues(SETTINGS_ROWS);

  sh.setFrozenRows(1);
  // Колонка значень — суворо текст, щоб довгі ID не перетворювались на числа
  sh.getRange("B2:B").setNumberFormat("@");

  return sh;
}

/** Дописує ключі, яких ще немає (щоб старі таблиці отримали нові налаштування). */
function ensureSettingsRows(sh) {
  var values = sh.getDataRange().getValues();
  var have = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || "").trim();
    if (k) have[k] = true;
  }

  var missing = SETTINGS_ROWS.filter(function (r) { return !have[r[0]]; });
  if (!missing.length) return;

  var start = sh.getLastRow() + 1;
  sh.getRange(start, 1, missing.length, 3).setValues(missing);
  sh.getRange(start, 2, missing.length, 1).setNumberFormat("@");
}

/** Читає всі пари «ключ → значення» з вкладки. */
function getSettings() {
  var values = ensureSettingsSheet().getDataRange().getValues();
  var out = {};

  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || "").trim();
    if (!k) continue;
    var v = values[i][1];
    out[k] = (v === null || v === undefined) ? "" : v;
  }

  // Прапорці приймаємо і як булеве, і як текст «TRUE» / «так» / «1»
  out.singlePage    = asBool(out.singlePage, true);
  out.cloudAutoSave = asBool(out.cloudAutoSave, true);

  return { status: "success", settings: out };
}

function asBool(v, dflt) {
  if (v === true) return true;
  if (v === false) return false;
  var s = String(v === null || v === undefined ? "" : v).trim();
  if (!s) return !!dflt;
  return /^(true|1|так|yes)$/i.test(s);
}

/** Записує налаштування назад у вкладку. Невідомі ключі дописує знизу. */
function saveSettings(data) {
  var s = data.settings || {};
  var sh = ensureSettingsSheet();
  var values = sh.getDataRange().getValues();

  var rowByKey = {};
  for (var i = 1; i < values.length; i++) {
    var k = String(values[i][0] || "").trim();
    if (k) rowByKey[k] = i + 1;
  }

  Object.keys(s).forEach(function (key) {
    var v = s[key];
    if (typeof v === "boolean") v = v ? "TRUE" : "FALSE";
    if (v === null || v === undefined) v = "";

    var row = rowByKey[key];
    if (!row) {
      row = sh.getLastRow() + 1;
      sh.getRange(row, 1).setValue(key);
      sh.getRange(row, 2).setNumberFormat("@");
    }
    sh.getRange(row, 2).setValue(String(v));
  });

  SpreadsheetApp.flush();
  return { status: "success", message: "Налаштування збережено у таблицю." };
}

// ─────────────────────────────────────────────────────────────
// СПІЛЬНА БАЗА ЧЕРНЕТОК
//
// Одна чернетка = одна папка на Диску:
//   draft.json  — текст інструкції + ID файлів фото
//   *.jpg       — самі фото, кожне вивантажується РІВНО ОДИН РАЗ
// Вкладка «Чернетки» в таблиці — покажчик: де що лежить і хто коли правив.
// ─────────────────────────────────────────────────────────────

var DRAFT_HEADERS = ["ID", "Назва", "Обладнання", "Вузлів", "Фото", "Оновлено",
                     "Пристрій", "Версія", "Готово", "Папка", "folderId", "jsonId"];

function ensureDraftsSheet() {
  if (!SPREADSHEET_ID) throw new Error("У скрипті не вказано SPREADSHEET_ID — нема де тримати спільну базу.");

  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sh = ss.getSheetByName(DRAFTS_SHEET_NAME);
  if (sh) return sh;

  sh = ss.insertSheet(DRAFTS_SHEET_NAME, 1);
  sh.getRange(1, 1, 1, DRAFT_HEADERS.length)
    .setValues([DRAFT_HEADERS])
    .setFontWeight("bold")
    .setBackground("#e2e8f0");

  sh.setColumnWidth(1, 240);   // ID
  sh.setColumnWidth(2, 260);   // Назва
  sh.setColumnWidth(3, 180);   // Обладнання
  sh.setColumnWidth(6, 130);   // Оновлено
  sh.setColumnWidth(7, 140);   // Пристрій
  sh.setFrozenRows(1);
  sh.hideColumns(11, 2);       // службові folderId / jsonId

  return sh;
}

/** id -> { row, title, equipment, nodes, photos, updated, device, rev, done, folderId, jsonId } */
function draftIndex() {
  var sh = ensureDraftsSheet();
  var values = sh.getDataRange().getValues();
  var map = {};

  for (var i = 1; i < values.length; i++) {
    var id = String(values[i][0] || "").trim();
    if (!id) continue;

    var when = values[i][5];
    map[id] = {
      row:       i + 1,
      title:     String(values[i][1] || ""),
      equipment: String(values[i][2] || ""),
      nodes:     Number(values[i][3]) || 0,
      photos:    Number(values[i][4]) || 0,
      updated:   (when instanceof Date) ? when.getTime() : 0,
      device:    String(values[i][6] || ""),
      rev:       Number(values[i][7]) || 0,
      done:      String(values[i][8] || "") !== "",
      folderId:  String(values[i][10] || ""),
      jsonId:    String(values[i][11] || "")
    };
  }
  return { sheet: sh, map: map };
}

/** Коренева папка для всіх чернеток. */
function draftsRoot(data) {
  var raw = (data && data.draftDriveId || "").trim();
  if (raw) return resolveFolder(raw, "чернеток");

  // Не вказано окрему папку — робимо підпапку там, де вже лежать фото або PDF
  var parentId = (data && data.photoDriveId || "").trim() || (data && data.driveId || "").trim();
  if (!parentId) {
    throw new Error("Не вказано жодної папки Google Диску в налаштуваннях — спільній базі немає де жити.");
  }

  var parent = resolveFolder(parentId, "чернеток");
  return findOrCreateFolder(parent, DRAFTS_FOLDER_NAME);
}

function findOrCreateFolder(parent, name) {
  var it = parent.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent.createFolder(name);
}

/** Папка конкретної чернетки: беремо з покажчика, а якщо зникла — робимо нову. */
function draftFolder(data, id, title, rec) {
  if (rec && rec.folderId) {
    try { return DriveApp.getFolderById(rec.folderId); } catch (e) {}
  }
  var root = draftsRoot(data);
  var name = sanitizeName(title || "Інструкція", 60) + " [" + String(id).substring(0, 8) + "]";
  return findOrCreateFolder(root, name);
}

/** Фото чернетки. Кожне вивантажується один раз, далі живе на Диску. */
function draftUploadPhotos(data) {
  var id = String(data.draftId || "").trim();
  if (!id) throw new Error("Не вказано ID чернетки.");

  var photos = data.photos || [];
  if (!photos.length) return { status: "success", ids: {} };

  var idx = draftIndex();
  var folder = draftFolder(data, id, data.title, idx.map[id]);

  var ids = {};
  for (var i = 0; i < photos.length; i++) {
    var p = photos[i];
    if (!p || !p.key || !p.dataUrl) continue;

    var blob = dataUrlToBlob(p.dataUrl, sanitizeName(p.name || ("photo_" + i + ".jpg"), 90));
    if (!blob) continue;

    try {
      ids[p.key] = folder.createFile(blob).getId();
    } catch (err) {
      Logger.log("Фото чернетки не збереглось: " + err);
    }
  }

  return { status: "success", ids: ids, folderId: folder.getId(), folderUrl: folder.getUrl() };
}

/** Запис чернетки. rev захищає від того, щоб два телефони затерли один одного. */
function draftSave(data) {
  var d = data.draft || {};
  var id = String(d.id || "").trim();
  if (!id) throw new Error("Чернетка без ID — нічого зберігати.");

  var idx = draftIndex();
  var rec = idx.map[id];
  var device = String(d.device || "невідомий пристрій").substring(0, 60);

  // Хтось уже записав новішу версію — не затираємо молчки
  if (rec && !data.force && Number(d.rev || 0) < rec.rev) {
    return {
      status: "success",
      conflict: true,
      serverRev: rec.rev,
      serverUpdated: rec.updated,
      serverDevice: rec.device,
      message: "Цю інструкцію змінили з іншого пристрою (" + (rec.device || "невідомо") + ")."
    };
  }

  var rev = (rec ? rec.rev : 0) + 1;
  var now = new Date();

  var folder = draftFolder(data, id, d.title, rec);

  var payload = {
    id:        id,
    title:     d.title || "",
    equipment: d.equipment || "",
    mainPhoto: d.mainPhoto || null,
    nodes:     d.nodes || [],
    rev:       rev,
    device:    device,
    updated:   now.getTime()
  };

  var jsonFile = null;
  if (rec && rec.jsonId) {
    try { jsonFile = DriveApp.getFileById(rec.jsonId); } catch (e) { jsonFile = null; }
  }
  if (jsonFile) {
    jsonFile.setContent(JSON.stringify(payload));
  } else {
    var existing = folder.getFilesByName("draft.json");
    if (existing.hasNext()) {
      jsonFile = existing.next();
      jsonFile.setContent(JSON.stringify(payload));
    } else {
      jsonFile = folder.createFile("draft.json", JSON.stringify(payload), MimeType.PLAIN_TEXT);
    }
  }

  var row = rec ? rec.row : (idx.sheet.getLastRow() + 1);
  idx.sheet.getRange(row, 1, 1, 8).setValues([[
    id, payload.title, payload.equipment, (payload.nodes || []).length,
    countDraftPhotos(payload), now, device, rev
  ]]);
  idx.sheet.getRange(row, 6).setNumberFormat("dd.MM.yyyy HH:mm");
  idx.sheet.getRange(row, 10).setFormula('=HYPERLINK("' + folder.getUrl() + '";"Відкрити папку")');
  idx.sheet.getRange(row, 11).setValue(folder.getId());
  idx.sheet.getRange(row, 12).setValue(jsonFile.getId());
  if (!rec) idx.sheet.getRange(row, 1).setNumberFormat("@");

  SpreadsheetApp.flush();

  return {
    status: "success",
    id: id,
    rev: rev,
    updated: now.getTime(),
    folderUrl: folder.getUrl(),
    message: "Інструкцію записано у спільну базу (версія " + rev + ")."
  };
}

function countDraftPhotos(payload) {
  var n = payload.mainPhoto ? 1 : 0;
  (payload.nodes || []).forEach(function (node) {
    if (node.photo) n++;
    (node.steps || []).forEach(function (s) { n += (s.photos || []).length; });
  });
  return n;
}

function draftList(data) {
  var idx = draftIndex();
  var out = [];

  Object.keys(idx.map).forEach(function (id) {
    var r = idx.map[id];
    out.push({
      id: id, title: r.title, equipment: r.equipment,
      nodes: r.nodes, photos: r.photos, updated: r.updated,
      device: r.device, rev: r.rev, done: r.done
    });
  });

  out.sort(function (a, b) { return b.updated - a.updated; });
  return { status: "success", drafts: out };
}

function draftLoad(data) {
  var id = String(data.id || "").trim();
  if (!id) throw new Error("Не вказано ID чернетки.");

  var rec = draftIndex().map[id];
  if (!rec) throw new Error("У спільній базі немає інструкції з таким ID. Можливо, її видалили.");

  var file = null;
  if (rec.jsonId) {
    try { file = DriveApp.getFileById(rec.jsonId); } catch (e) {}
  }
  if (!file && rec.folderId) {
    try {
      var it = DriveApp.getFolderById(rec.folderId).getFilesByName("draft.json");
      if (it.hasNext()) file = it.next();
    } catch (e) {}
  }
  if (!file) throw new Error("Файл чернетки не знайдено на Диску — схоже, папку видалили вручну.");

  var draft;
  try {
    draft = JSON.parse(file.getBlob().getDataAsString("UTF-8"));
  } catch (e) {
    throw new Error("Файл чернетки пошкоджений і не читається.");
  }

  return { status: "success", draft: draft, rev: rec.rev, updated: rec.updated, device: rec.device };
}

/** Фото назад у телефон: віддаємо base64 пачкою, з оглядкою на ліміт відповіді. */
function draftPhotos(data) {
  var ids = data.ids || [];
  var photos = {};
  var served = [];
  var total = 0;

  for (var i = 0; i < ids.length; i++) {
    var fid = String(ids[i] || "").trim();
    if (!fid) continue;

    try {
      var blob = DriveApp.getFileById(fid).getBlob();
      var bytes = blob.getBytes();

      // Перше фото віддаємо завжди, інакше телефон ніколи не дочитає великий файл
      if (served.length && total + bytes.length > MAX_DOWNLOAD_BYTES) break;

      photos[fid] = "data:" + (blob.getContentType() || "image/jpeg") +
                    ";base64," + Utilities.base64Encode(bytes);
      served.push(fid);
      total += bytes.length;
    } catch (e) {
      Logger.log("Фото " + fid + " не читається: " + e);
      photos[fid] = null;   // щоб телефон не чекав його вічно
      served.push(fid);
    }
  }

  return { status: "success", photos: photos, served: served };
}

function draftDelete(data) {
  var id = String(data.id || "").trim();
  if (!id) throw new Error("Не вказано ID чернетки.");

  var idx = draftIndex();
  var rec = idx.map[id];
  if (!rec) return { status: "success", message: "У базі такої інструкції вже немає." };

  // Папку з фото не знищуємо назавжди, а кладемо в кошик Диску —
  // звідти її можна дістати, якщо видалили помилково.
  if (rec.folderId) {
    try { DriveApp.getFolderById(rec.folderId).setTrashed(true); }
    catch (e) { Logger.log("Папку чернетки не прибрано: " + e); }
  }

  idx.sheet.deleteRow(rec.row);
  SpreadsheetApp.flush();

  return { status: "success", message: "Інструкцію видалено зі спільної бази (папка — у кошику Диску)." };
}

/** Позначка «з чернетки зроблено PDF». */
function markDraftDone(id, pdfUrl) {
  var idx = draftIndex();
  var rec = idx.map[String(id || "").trim()];
  if (!rec) return;
  idx.sheet.getRange(rec.row, 9).setFormula('=HYPERLINK("' + pdfUrl + '";"Готовий PDF")');
  SpreadsheetApp.flush();
}

// ─────────────────────────────────────────────────────────────
// ДОПОМІЖНІ
// ─────────────────────────────────────────────────────────────

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Перевірка та розпакування PDF з base64 dataURL. */
function pdfBlobFrom(data) {
  var raw = data.pdfBase64;
  if (!raw) throw new Error("Файл PDF не надійшов від додатку.");

  // Приймаємо і "data:application/pdf;base64,XXXX", і чистий base64
  var payload = (raw.indexOf(",") > -1) ? raw.split(",")[1] : raw;

  var bytes;
  try {
    bytes = Utilities.base64Decode(payload);
  } catch (err) {
    throw new Error("Не вдалося декодувати PDF (пошкоджені дані).");
  }

  Logger.log("Отримано PDF байтів: " + bytes.length);

  if (bytes.length < MIN_PDF_BYTES) {
    throw new Error("Згенерований PDF занадто малий (" + bytes.length + " байт) — схоже на порожній аркуш.");
  }

  // Сигнатура %PDF
  if (bytes[0] !== 0x25 || bytes[1] !== 0x50 || bytes[2] !== 0x44 || bytes[3] !== 0x46) {
    throw new Error("Надісланий файл не є коректним PDF.");
  }

  var name = data.filename || ((data.title || "Інструкція") + ".pdf");
  if (name.slice(-4).toLowerCase() !== ".pdf") name += ".pdf";

  return Utilities.newBlob(bytes, MimeType.PDF, name);
}

/** Приймає і чистий ID папки, і повне посилання на неї. */
function resolveFolder(driveId, label) {
  var what = label ? ("папки для " + label) : "папки Google Диску";
  var raw = (driveId || "").trim();
  if (!raw) throw new Error("Не вказано ID " + what + " у налаштуваннях додатку.");

  var id = raw;
  var m = raw.match(/[-\w]{25,}/);
  if (m) id = m[0];

  try {
    return DriveApp.getFolderById(id);
  } catch (err) {
    throw new Error("Не знайдено " + what + " з ID «" + id + "» або немає доступу. Скопіюйте посилання на папку з Google Диску ще раз.");
  }
}

/** Вкладка-журнал: текст у A–C, посилання на фото у колонці D. */
function writeLogSheet(data, fileUrl) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  var safeTitle = (data.title || "Інструкція")
    .replace(/[^a-zA-Z0-9а-яА-ЯіІїЇєЄґҐ\s-]/g, "").trim().substring(0, 15);
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM HH:mm");
  var base = (safeTitle || "Інструкція") + " " + stamp;

  var sheetName = base;
  var counter = 1;
  while (ss.getSheetByName(sheetName)) {
    sheetName = base + " (" + counter + ")";
    counter++;
  }

  var sheet = ss.insertSheet(sheetName);
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(2, 60);
  sheet.setColumnWidth(3, 450);
  sheet.setColumnWidth(4, 200);

  sheet.getRange("A1").setValue("НАЗВА ІНСТРУКЦІЇ:").setFontWeight("bold");
  sheet.getRange("B1:D1").merge().setValue(data.title || "").setFontWeight("bold").setFontSize(14);
  sheet.getRange("A2").setValue("ОБЛАДНАННЯ:").setFontWeight("bold");
  sheet.getRange("B2:D2").merge().setValue(data.equipment || "");
  sheet.getRange("A3").setValue("ДАТА СТВОРЕННЯ:").setFontWeight("bold");
  sheet.getRange("B3:D3").merge().setValue(data.date || "");
  sheet.getRange("A4").setValue("PDF:").setFontWeight("bold");
  sheet.getRange("B4:D4").merge().setFormula('=HYPERLINK("' + fileUrl + '";"Відкрити PDF")');

  // Текст пишемо одним setValues, посилання — окремим проходом
  var rows = [["Вузол / Частина", "Крок №", "Опис дії", "Фото"]];
  var rowLinks = [null];

  (data.nodes || []).forEach(function (node) {
    var nodeText = String(node.title || "").toUpperCase();
    if (node.desc) nodeText += "\n" + node.desc;
    rows.push([nodeText, "", "", ""]);
    rowLinks.push(node.imageUrls || null);

    (node.steps || []).forEach(function (step, i) {
      rows.push(["", i + 1, step.text || "", ""]);
      rowLinks.push(step.imageUrls || null);
    });
  });

  var start = 6;
  sheet.getRange(start, 1, rows.length, 4).setValues(rows);
  sheet.getRange(start, 1, 1, 4).setFontWeight("bold").setBackground("#e2e8f0");
  sheet.getRange(start, 1, rows.length, 4)
    .setVerticalAlignment("top").setWrap(true)
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // Посилання на фото. Якщо фото кілька — кілька клікабельних рядків в одній комірці.
  for (var r = 0; r < rowLinks.length; r++) {
    var links = rowLinks[r];
    if (!links || !links.length) continue;

    var labels = [];
    for (var k = 0; k < links.length; k++) {
      labels.push(links.length > 1 ? ("📷 Фото " + (k + 1)) : "📷 Фото");
    }
    var text = labels.join("\n");

    var builder = SpreadsheetApp.newRichTextValue().setText(text);
    var pos = 0;
    for (var j = 0; j < labels.length; j++) {
      builder.setLinkUrl(pos, pos + labels[j].length, links[j]);
      pos += labels[j].length + 1; // +1 на символ переносу рядка
    }

    sheet.getRange(start + r, 4).setRichTextValue(builder.build());
  }

  SpreadsheetApp.flush();
  return sheetName;
}

/** Прибирає символи, недопустимі в іменах файлів Диску. */
function sanitizeName(name, max) {
  var clean = String(name || "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  if (max && clean.length > max) clean = clean.substring(0, max);
  return clean || "file";
}

/** dataURL ("data:image/jpeg;base64,...") -> Blob. Повертає null, якщо дані биті. */
function dataUrlToBlob(dataUrl, name) {
  if (!dataUrl || typeof dataUrl !== "string") return null;
  try {
    var parts = dataUrl.split(",");
    var payload = (parts.length > 1) ? parts[1] : parts[0];
    var mime = "image/jpeg";
    var m = dataUrl.match(/^data:([^;,]+)/);
    if (m) mime = m[1];
    return Utilities.newBlob(Utilities.base64Decode(payload), mime, name || "photo.jpg");
  } catch (e) {
    Logger.log("dataUrlToBlob: " + e);
    return null;
  }
}

/** Запустити ОДИН РАЗ вручну в редакторі, щоб видати дозволи. */
function setupPermissions() {
  MailApp.getRemainingDailyQuota();
  var tempFile = DriveApp.createFile("temp_auth_file.txt", "auth");
  tempFile.setTrashed(true);
  if (SPREADSHEET_ID) {
    SpreadsheetApp.openById(SPREADSHEET_ID).getName();
    ensureSettingsSheet();   // одразу створюємо вкладку «Налаштування»
    ensureDraftsSheet();     // ...і покажчик спільної бази
  }
  Logger.log("Дозволи видано успішно.");
}
