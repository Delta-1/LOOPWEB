const LOOP_CONFIG = Object.freeze({
  sheetName: "Produtos",
  spreadsheetIdProperty: "LOOP_SPREADSHEET_ID",
  folderIdProperty: "LOOP_FOLDER_ID"
});

const PRODUCT_KEYS = ["code", "name", "price", "stock", "category", "unit", "image"];

const HEADER_ALIASES = Object.freeze({
  code: ["code", "codigo", "cod", "cod registro", "codigo do item"],
  name: ["name", "nome", "produto", "nome do produto", "nome do item"],
  price: ["price", "preco", "preco base", "preco un medida", "valor", "valor unitario"],
  stock: ["stock", "estoque", "estoque fisico", "quantidade", "disponibilidade", "disponibilidade inicial"],
  category: ["category", "categoria", "categoria comercial"],
  unit: ["unit", "unidade", "medida", "unidade de medida"],
  image: ["image", "imagem", "foto", "fotografia", "fotografia de referencia", "url da imagem", "url da foto"]
});

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || "read").toLowerCase();
    if (action !== "read") {
      return jsonResponse_({ status: "error", message: "Ação GET não suportada." });
    }

    return jsonResponse_({ status: "success", data: readProducts_() });
  } catch (error) {
    return errorResponse_(error);
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(30000);
    const payload = parsePayload_(e);
    const action = String(payload.action || "").toLowerCase();

    if (action === "create") {
      return jsonResponse_({ status: "success", data: createProduct_(payload) });
    }
    if (action === "update") {
      return jsonResponse_({ status: "success", data: updateProduct_(payload) });
    }
    if (action === "delete") {
      deleteProduct_(payload.code);
      return jsonResponse_({ status: "success" });
    }

    return jsonResponse_({ status: "error", message: "Ação POST não suportada." });
  } catch (error) {
    return errorResponse_(error);
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function readProducts_() {
  const sheet = getProductSheet_();
  const columns = resolveColumns_(sheet);
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return [];

  const width = Math.max(sheet.getLastColumn(), PRODUCT_KEYS.length);
  const values = sheet.getRange(2, 1, lastRow - 1, width).getValues();

  return values
    .filter(row => String(row[columns.code] || "").trim())
    .map(row => ({
      code: String(row[columns.code] || "").trim(),
      name: String(row[columns.name] || "").trim(),
      price: toNumber_(row[columns.price]),
      stock: toNumber_(row[columns.stock]),
      category: String(row[columns.category] || "Geral").trim() || "Geral",
      unit: String(row[columns.unit] || "un").trim() || "un",
      image: String(row[columns.image] || "").trim()
    }));
}

function createProduct_(payload) {
  validateProduct_(payload);

  const sheet = getProductSheet_();
  const columns = resolveColumns_(sheet);
  const existingRow = findProductRow_(sheet, columns.code, payload.code);

  // Um código representa um único produto. Se ele já existir, atualiza a linha
  // em vez de criar uma duplicata que poderia apontar para uma foto antiga.
  if (existingRow) {
    return writeProduct_(sheet, columns, existingRow, payload, false);
  }

  const newRow = Math.max(2, sheet.getLastRow() + 1);
  return writeProduct_(sheet, columns, newRow, payload, true);
}

function updateProduct_(payload) {
  validateProduct_(payload);

  const sheet = getProductSheet_();
  const columns = resolveColumns_(sheet);
  const row = findProductRow_(sheet, columns.code, payload.code);

  if (!row) throw new Error("Produto não encontrado para atualização.");
  return writeProduct_(sheet, columns, row, payload, false);
}

function deleteProduct_(code) {
  const normalizedCode = String(code || "").trim();
  if (!normalizedCode) throw new Error("Código do produto não informado.");

  const sheet = getProductSheet_();
  const columns = resolveColumns_(sheet);
  const row = findProductRow_(sheet, columns.code, normalizedCode);

  if (!row) throw new Error("Produto não encontrado para exclusão.");
  sheet.deleteRow(row);
}

function writeProduct_(sheet, columns, row, payload, isNew) {
  const currentImage = isNew ? "" : String(sheet.getRange(row, columns.image + 1).getValue() || "");
  const imageUrl = payload.imageBase64
    ? saveProductImage_(payload)
    : currentImage;

  const product = {
    code: String(payload.code).trim(),
    name: String(payload.name).trim(),
    price: toNumber_(payload.price),
    stock: toNumber_(payload.stock),
    category: String(payload.category || "Geral").trim() || "Geral",
    unit: String(payload.unit || "un").trim() || "un",
    image: imageUrl
  };

  PRODUCT_KEYS.forEach(key => {
    sheet.getRange(row, columns[key] + 1).setValue(product[key]);
  });

  SpreadsheetApp.flush();
  return product;
}

function saveProductImage_(payload) {
  const folderId = String(
    payload.folderId ||
    PropertiesService.getScriptProperties().getProperty(LOOP_CONFIG.folderIdProperty) ||
    ""
  ).trim();

  if (!folderId) throw new Error("ID da pasta de fotos não configurado.");

  const contentType = String(payload.imageMimeType || "image/jpeg");
  const safeCode = sanitizeFilePart_(payload.code || "produto");
  const requestedName = sanitizeFilePart_(payload.imageName || "");
  const fileName = requestedName || `${safeCode}_IMG_${Date.now()}.jpg`;
  const bytes = Utilities.base64Decode(String(payload.imageBase64));
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = DriveApp.getFolderById(folderId).createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // O ID muda a cada substituição; assim o navegador nunca reaproveita
  // silenciosamente a foto antiga pelo cache.
  return `https://drive.google.com/thumbnail?id=${file.getId()}&sz=w1600`;
}

function getProductSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(LOOP_CONFIG.spreadsheetIdProperty);
  const spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      `Planilha não encontrada. Configure a propriedade ${LOOP_CONFIG.spreadsheetIdProperty}.`
    );
  }

  return spreadsheet.getSheetByName(LOOP_CONFIG.sheetName) || spreadsheet.getSheets()[0];
}

function resolveColumns_(sheet) {
  const width = Math.max(sheet.getLastColumn(), PRODUCT_KEYS.length);
  let headers = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];

  if (headers.every(value => !String(value).trim())) {
    headers = ["Código", "Nome", "Preço", "Estoque", "Categoria", "Unidade", "Imagem"];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  const normalizedHeaders = headers.map(normalizeText_);
  const columns = {};

  PRODUCT_KEYS.forEach((key, fallbackIndex) => {
    const aliases = HEADER_ALIASES[key].map(normalizeText_);
    const matchedIndex = normalizedHeaders.findIndex(header => aliases.includes(header));
    columns[key] = matchedIndex >= 0 ? matchedIndex : fallbackIndex;
  });

  return columns;
}

function findProductRow_(sheet, codeColumnIndex, code) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const target = String(code || "").trim().toLowerCase();
  const codes = sheet.getRange(2, codeColumnIndex + 1, lastRow - 1, 1).getDisplayValues();
  const offset = codes.findIndex(row => String(row[0]).trim().toLowerCase() === target);
  return offset < 0 ? 0 : offset + 2;
}

function parsePayload_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    throw new Error("Corpo da requisição não informado.");
  }
  return JSON.parse(e.postData.contents);
}

function validateProduct_(payload) {
  if (!String(payload.code || "").trim()) throw new Error("Código do produto não informado.");
  if (!String(payload.name || "").trim()) throw new Error("Nome do produto não informado.");
}

function toNumber_(value) {
  if (typeof value === "number") return value;
  const normalized = String(value || "0")
    .replace(/\s/g, "")
    .replace(/\.(?=\d{3}(?:\D|$))/g, "")
    .replace(",", ".");
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function normalizeText_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]+/g, " ")
    .trim()
    .toLowerCase();
}

function sanitizeFilePart_(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9_.-]/g, "_")
    .replace(/_+/g, "_")
    .slice(0, 120);
}

function jsonResponse_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorResponse_(error) {
  console.error(error && error.stack ? error.stack : error);
  return jsonResponse_({
    status: "error",
    message: error && error.message ? error.message : "Erro interno no Apps Script."
  });
}
