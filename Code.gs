/**
 * ═══════════════════════════════════════════════════════════════════════════
 * GENERADOR DE CERTIFICADOS & EMAILER - GOOGLE APPS SCRIPT
 * ═══════════════════════════════════════════════════════════════════════════
 * 
 * INSTRUCCIONES DE CONFIGURACIÓN:
 * 1. Abre tu Google Sheets con los datos de los certificados
 * 2. Ve a Extensiones > Apps Script
 * 3. Borra cualquier código previo y pega este script entero
 * 4. Guarda el proyecto (Ctrl+S o Cmd+S)
 * 5. Refresca tu Google Sheets (F5)
 * 6. Un nuevo menú "Sistema de Certificados" aparecerá
 * 7. Haz click para autorizar el primer uso del certificado (recomendación: dale al botón de ayuda para tener más info también)
 * 
 * USO:
 * Fase 1: Sistema de Certificados > Generar Certificados (crea PDFs sin firmar)
 * Fase 2: Sistema de Certificados > Mandar Certificados Firmados (manda los PDFs firmados)
 * 
 * REQUISITOS:
 * - La hoja debe contener columnas para: Nombre, Apellido(s), Email (mínimo)
 * - Plantilla en Google Docs con placeholders como {{NOMBRE}}, {{APELLIDOS}}, {{EMAIL}}
 * - Los PDFs firmados tienen que tener el siguiente nombre: NombreApellido_signed.pdf
 * 
 * ═══════════════════════════════════════════════════════════════════════════
 */

// ═══════════════════════════════════════════════════════════════════════════
// MENU SETUP
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Creates a personalized menu when the spreadsheet is open
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sistema de Certificados')
    .addItem('📄 Generar Certificados', 'generateCertificates')
    .addSeparator()
    .addItem('📧 Mandar Certificados Firmados', 'sendCertificates')
    .addSeparator()
    .addItem('🔁 Reintentar Certificados Fallidos', 'retryFailedCertificates')
    .addItem('📨 Reintentar Emails Fallidos', 'retryFailedEmails')
    .addSeparator()
    .addItem('🗑️ Limpiar Registro de Errores', 'clearErrorLog')
    .addSeparator()
    .addItem('ℹ️ Ayuda e Instrucciones', 'showHelp')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════
// ERROR LOGGING
// ═══════════════════════════════════════════════════════════════════════════

const ERROR_SHEET_NAME = '🔴 Registro de Errores';

/**
 * Returns (or creates) the error log sheet with proper headers.
 * Columns: Timestamp | Tipo | Fila | Nombre | Apellidos | Email | Nombre PDF | Motivo del Error | Estado
 */
function getOrCreateErrorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let errorSheet = ss.getSheetByName(ERROR_SHEET_NAME);

  if (!errorSheet) {
    errorSheet = ss.insertSheet(ERROR_SHEET_NAME);

    // Style the header row
    const headers = [
      'Timestamp', 'Tipo de Error', 'Fila (original)', 'Nombre', 'Apellidos',
      'Email', 'Nombre PDF Esperado', 'Motivo del Error', 'Estado'
    ];
    const headerRange = errorSheet.getRange(1, 1, 1, headers.length);
    headerRange.setValues([headers]);
    headerRange.setBackground('#c0392b');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    errorSheet.setFrozenRows(1);
    errorSheet.setColumnWidth(1, 160);
    errorSheet.setColumnWidth(2, 140);
    errorSheet.setColumnWidth(7, 200);
    errorSheet.setColumnWidth(8, 300);
    errorSheet.setColumnWidth(9, 100);
  }

  return errorSheet;
}

/**
 * Appends one or more error rows to the error log sheet.
 * Each error object: { type, rowIndex, name, surname, email, pdfName, reason }
 */
function logErrors(errors) {
  if (!errors || errors.length === 0) return;

  const errorSheet = getOrCreateErrorSheet();
  const timestamp = new Date().toLocaleString('es-ES');

  const rows = errors.map(e => [
    timestamp,
    e.type   || '',
    e.rowIndex !== undefined ? e.rowIndex + 1 : '',   // human-readable (1-based)
    e.name   || '',
    e.surname || '',
    e.email  || '',
    e.pdfName || '',
    e.reason || '',
    'Pendiente'
  ]);

  const firstNewRow = errorSheet.getLastRow() + 1;
  errorSheet.getRange(firstNewRow, 1, rows.length, 9).setValues(rows);

  // Highlight new rows in light red
  errorSheet.getRange(firstNewRow, 1, rows.length, 9).setBackground('#fdecea');
}

/**
 * Marks a row in the error log as resolved.
 */
function markErrorResolved(errorSheet, sheetRow) {
  errorSheet.getRange(sheetRow, 9).setValue('✅ Resuelto');
  errorSheet.getRange(sheetRow, 1, 1, 9).setBackground('#eafaf1');
}

/**
 * Clears all rows (except header) from the error log sheet.
 */
function clearErrorLog() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    ui.alert('Info', 'El registro de errores ya está vacío.', ui.ButtonSet.OK);
    return;
  }

  const confirm = ui.alert(
    'Limpiar Registro',
    '¿Seguro que quieres eliminar todas las entradas del registro de errores?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  errorSheet.deleteRows(2, errorSheet.getLastRow() - 1);
  ui.alert('Hecho', 'Registro de errores limpiado.', ui.ButtonSet.OK);
}

// ═══════════════════════════════════════════════════════════════════════════
// PHASE 1: CERTIFICATE GENERATION
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Main function to generate certificates
 */
function generateCertificates() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    const criticalColumns = identifyCriticalColumnsForGeneration(ui, sheet);
    if (!criticalColumns) return;

    const templateDoc = getTemplateDocument(ui);
    if (!templateDoc) return;

    const placeholders = detectPlaceholders(templateDoc);
    if (placeholders.length === 0) {
      ui.alert('Error', 'No se han encontrado placeholders en la plantilla. Usa el formato: {{PLACEHOLDER}}', ui.ButtonSet.OK);
      return;
    }

    ui.alert('Placeholders Encontrados',
      `Detectados ${placeholders.length} placeholders:\n${placeholders.join(', ')}\n\nSiguiente: Asígnalos a las columnas de tu hoja.`,
      ui.ButtonSet.OK);

    const columnMapping = mapPlaceholdersToColumns(ui, sheet, placeholders);
    if (!columnMapping) return;

    const outputFolder = getOutputFolder(ui);
    if (!outputFolder) return;

    const unsignedFolder = getOrCreateFolder(outputFolder, 'No firmados');

    const result = generateCertificatePDFs(sheet, templateDoc, columnMapping, unsignedFolder, criticalColumns);

    // Log any errors to the error sheet
    if (result.errors.length > 0) {
      logErrors(result.errors);
    }

    let message =
      `✅ Generados correctamente: ${result.success}\n` +
      `❌ Fallidos: ${result.failed}\n\n` +
      `PDFs guardados en:\n${unsignedFolder.getName()}\n\n` +
      `URL de la carpeta:\n${unsignedFolder.getUrl()}`;

    if (result.failed > 0) {
      message += `\n\n⚠️ Se han registrado ${result.failed} error(es) en la hoja "${ERROR_SHEET_NAME}".\n` +
                 `Revísala para identificar y corregir los certificados fallidos.`;
    }

    ui.alert('¡Generación Completa!', message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', `Un error ha ocurrido: ${error.toString()}`, ui.ButtonSet.OK);
    Logger.log('Error in generateCertificates: ' + error.toString());
  }
}

/**
 * Retry certificate generation only for rows that appear as pending in the error log.
 */
function retryFailedCertificates() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    ui.alert('Sin errores', 'No hay errores pendientes en el registro.', ui.ButtonSet.OK);
    return;
  }

  // Collect pending certificate-generation errors
  const errorData = errorSheet.getDataRange().getValues();
  const pendingRows = [];   // { sheetRow (1-based in errorSheet), originalRow (0-based in data) }

  for (let r = 1; r < errorData.length; r++) {
    const tipo   = errorData[r][1];
    const estado = errorData[r][8];
    if (tipo === 'Generación de Certificado' && estado === 'Pendiente') {
      pendingRows.push({ sheetRow: r + 1, originalDataRow: parseInt(errorData[r][2]) - 1 });
    }
  }

  if (pendingRows.length === 0) {
    ui.alert('Sin errores', 'No hay errores pendientes de generación de certificados.', ui.ButtonSet.OK);
    return;
  }

  ui.alert('Reintentar Certificados',
    `Se van a reintentar ${pendingRows.length} certificado(s) fallido(s).\n\nNecesitarás introducir de nuevo la plantilla y la carpeta de destino.`,
    ui.ButtonSet.OK);

  const mainSheet = ss.getActiveSheet();
  const criticalColumns = identifyCriticalColumnsForGeneration(ui, mainSheet);
  if (!criticalColumns) return;

  const templateDoc = getTemplateDocument(ui);
  if (!templateDoc) return;

  const placeholders = detectPlaceholders(templateDoc);
  if (placeholders.length === 0) {
    ui.alert('Error', 'No se han encontrado placeholders en la plantilla.', ui.ButtonSet.OK);
    return;
  }

  const columnMapping = mapPlaceholdersToColumns(ui, mainSheet, placeholders);
  if (!columnMapping) return;

  const outputFolder = getOutputFolder(ui);
  if (!outputFolder) return;

  const unsignedFolder = getOrCreateFolder(outputFolder, 'No firmados');

  // Only process the failed rows
  const allData = mainSheet.getDataRange().getValues();
  let success = 0;
  let stillFailing = 0;
  const newErrors = [];

  for (const pending of pendingRows) {
    const rowData = allData[pending.originalDataRow];
    try {
      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(columnMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }

      const name    = rowData[criticalColumns.nameCol]    || 'Unknown';
      const surname = rowData[criticalColumns.surnameCol] || 'Unknown';
      const filename = `${name}${surname}.pdf`.replace(/\s+/g, '');

      const personalizedDoc = createPersonalizedDocument(templateDoc, replacements);
      const pdfBlob = convertDocToPdf(personalizedDoc);
      pdfBlob.setName(filename);
      unsignedFolder.createFile(pdfBlob);
      DriveApp.getFileById(personalizedDoc.getId()).setTrashed(true);

      // Mark as resolved in the error sheet
      markErrorResolved(errorSheet, pending.sheetRow);
      success++;

    } catch (err) {
      Logger.log(`Retry failed for data row ${pending.originalDataRow + 1}: ${err}`);
      stillFailing++;
      newErrors.push({
        type: 'Generación de Certificado',
        rowIndex: pending.originalDataRow,
        name: rowData[criticalColumns.nameCol]    || '',
        surname: rowData[criticalColumns.surnameCol] || '',
        email: '',
        pdfName: '',
        reason: err.toString()
      });
    }
  }

  if (newErrors.length > 0) logErrors(newErrors);

  ui.alert('Reintento Completado',
    `✅ Resueltos: ${success}\n❌ Siguen fallando: ${stillFailing}` +
    (stillFailing > 0 ? `\n\nRevisa "${ERROR_SHEET_NAME}" para más detalles.` : ''),
    ui.ButtonSet.OK);
}

/**
 * Identifies critical columns at the start (NAME and SURNAME for PDF filenames)
 */
function identifyCriticalColumnsForGeneration(ui, sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersList = headers.map((h, i) => `${i + 1}. ${h}`).join('\n');

  ui.alert('Configuración: Columnas Críticas',
    'Primero, identifica las columnas que contienen NOMBRE y APELLIDOS.\n\n' +
    'Estas se necesitan para crear los nombres de los PDFs (formato: NombreApellidos.pdf)\n\n' +
    '¡Puedes usar cualquier idioma para los nombres de las columnas!',
    ui.ButtonSet.OK);

  const criticalColumns = {};

  const nameResponse = ui.prompt(
    'Columna de Nombre',
    `Columnas disponibles:\n${headersList}\n\n` +
    'Introduce el NÚMERO de la columna que contiene los NOMBRES:\n' +
    '(e.g., "Name", "Nombre", "Prénom", "Nome", etc.)',
    ui.ButtonSet.OK_CANCEL
  );

  if (nameResponse.getSelectedButton() !== ui.Button.OK) return null;
  const nameCol = parseInt(nameResponse.getResponseText().trim());
  if (isNaN(nameCol) || nameCol < 1 || nameCol > headers.length) {
    ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
    return null;
  }
  criticalColumns.nameCol = nameCol - 1;

  const surnameResponse = ui.prompt(
    'Columna de Apellidos',
    `Columnas disponibles:\n${headersList}\n\n` +
    'Introduce el NÚMERO de la columna que contiene los APELLIDOS:\n' +
    '(e.g., "Surname", "Apellidos", "Nom de famille", "Sobrenome", etc.)',
    ui.ButtonSet.OK_CANCEL
  );

  if (surnameResponse.getSelectedButton() !== ui.Button.OK) return null;
  const surnameCol = parseInt(surnameResponse.getResponseText().trim());
  if (isNaN(surnameCol) || surnameCol < 1 || surnameCol > headers.length) {
    ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
    return null;
  }
  criticalColumns.surnameCol = surnameCol - 1;

  return criticalColumns;
}

/**
 * Prompts user to provide template document
 */
function getTemplateDocument(ui) {
  const response = ui.prompt(
    'Documento Plantilla',
    'Introduce la URL o el ID del Google Docs:\n\n' +
    '(El ID se encuentra en la URL del Doc: docs.google.com/document/d/DOCUMENT_ID/edit)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;

  const input = response.getResponseText().trim();
  const docId = extractDocumentId(input);

  try {
    return DocumentApp.openById(docId);
  } catch (error) {
    ui.alert('Error', 'No se pudo abrir el documento. Comprueba la URL/ID y los permisos.', ui.ButtonSet.OK);
    return null;
  }
}

function extractDocumentId(input) {
  const urlMatch = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return urlMatch ? urlMatch[1] : input;
}

function detectPlaceholders(doc) {
  const body = doc.getBody();
  const text = body.getText();
  const regex = /\{\{([A-Za-z0-9_]+)\}\}/g;
  const placeholders = new Set();
  let match;
  while ((match = regex.exec(text)) !== null) {
    placeholders.add(match[1]);
  }
  return Array.from(placeholders).sort();
}

function mapPlaceholdersToColumns(ui, sheet, placeholders) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersList = headers.map((h, i) => `${i + 1}. ${h}`).join('\n');
  const mapping = {};

  for (const placeholder of placeholders) {
    const response = ui.prompt(
      `Asigna: {{${placeholder}}}`,
      `Columnas disponibles:\n${headersList}\n\n` +
      `Introduce el NÚMERO de columna para {{${placeholder}}}:`,
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() !== ui.Button.OK) return null;
    const colNum = parseInt(response.getResponseText().trim());
    if (isNaN(colNum) || colNum < 1 || colNum > headers.length) {
      ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
      return null;
    }
    mapping[placeholder] = colNum - 1;
  }

  return mapping;
}

function getOutputFolder(ui) {
  const response = ui.prompt(
    'Carpeta de Salida',
    'Introduce el ID o la URL de la Carpeta de Google Drive donde deberían guardarse los certificados:\n\n' +
    '(Se encuentra en la URL de la carpeta: drive.google.com/drive/folders/FOLDER_ID)',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const input = response.getResponseText().trim();
  const folderId = extractFolderId(input);

  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    ui.alert('Error', 'No se pudo acceder a la carpeta. Comprueba la URL/ID y los permisos.', ui.ButtonSet.OK);
    return null;
  }
}

function extractFolderId(input) {
  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  return urlMatch ? urlMatch[1] : input;
}

function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(folderName);
}

/**
 * Generates PDF certificates for all rows.
 * Returns { success, failed, errors[] }
 */
function generateCertificatePDFs(sheet, templateDoc, columnMapping, outputFolder, criticalColumns) {
  const data = sheet.getDataRange().getValues();
  let successCount = 0;
  let failedCount = 0;
  const errors = [];

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const name    = rowData[criticalColumns.nameCol]    || 'Unknown';
    const surname = rowData[criticalColumns.surnameCol] || 'Unknown';
    const filename = `${name}${surname}.pdf`.replace(/\s+/g, '');

    try {
      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(columnMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }

      const personalizedDoc = createPersonalizedDocument(templateDoc, replacements);
      const pdfBlob = convertDocToPdf(personalizedDoc);
      pdfBlob.setName(filename);
      outputFolder.createFile(pdfBlob);
      DriveApp.getFileById(personalizedDoc.getId()).setTrashed(true);

      successCount++;

    } catch (error) {
      Logger.log(`Failed to generate certificate for row ${i + 1}: ${error.toString()}`);
      failedCount++;
      errors.push({
        type: 'Generación de Certificado',
        rowIndex: i,
        name: name,
        surname: surname,
        email: '',         // not relevant at this stage
        pdfName: filename,
        reason: error.toString()
      });
    }
  }

  return { success: successCount, failed: failedCount, errors };
}

function createPersonalizedDocument(templateDoc, replacements) {
  const templateFile = DriveApp.getFileById(templateDoc.getId());
  const copyFile = templateFile.makeCopy('temp_certificate_' + Date.now());
  const copyDoc = DocumentApp.openById(copyFile.getId());
  const body = copyDoc.getBody();

  for (const [placeholder, value] of Object.entries(replacements)) {
    body.replaceText(`\\{\\{${placeholder}\\}\\}`, value);
  }

  copyDoc.saveAndClose();
  return copyDoc;
}

function convertDocToPdf(doc) {
  const docId = doc.getId();
  const url = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
  const token = ScriptApp.getOAuthToken();

  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token }
  });

  return response.getBlob();
}

// ═══════════════════════════════════════════════════════════════════════════
// PHASE 2: SEND SIGNED CERTIFICATES
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Main function to send signed certificates via email
 */
function sendCertificates() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    const criticalColumns = identifyCriticalColumnsForSending(ui, sheet);
    if (!criticalColumns) return;

    const signedFolder = getSignedFolder(ui);
    if (!signedFolder) return;

    const emailSubject = getEmailSubject(ui);
    if (!emailSubject) return;

    const emailBody = getEmailBody(ui);
    if (!emailBody) return;

    const emailPlaceholders = detectEmailPlaceholders(emailSubject + ' ' + emailBody);
    let emailMapping = {};
    if (emailPlaceholders.length > 0) {
      // Show detected placeholders for the user to verify before mapping
      const placeholderList = emailPlaceholders.map(p => `  • {{${p}}}`).join('\n');
      const confirmPlaceholders = ui.alert(
        '✅ Placeholders Detectados en el Email',
        `Se han encontrado ${emailPlaceholders.length} placeholder(s) en el asunto y cuerpo del email:\n\n` +
        `${placeholderList}\n\n` +
        `¿Son correctos? Pulsa OK para continuar asignándolos a columnas,\n` +
        `o Cancelar para volver a escribir el asunto/cuerpo del email.`,
        ui.ButtonSet.OK_CANCEL
      );
      if (confirmPlaceholders !== ui.Button.OK) return;

      emailMapping = mapPlaceholdersToColumns(ui, sheet, emailPlaceholders);
      if (!emailMapping) return;
    } else {
      // Warn the user if no placeholders were found at all
      const confirmNoPlaceholders = ui.alert(
        '⚠️ Sin Placeholders',
        'No se han detectado placeholders ({{...}}) en el asunto ni en el cuerpo del email.\n\n' +
        'El mismo texto se enviará a todos los destinatarios sin personalización.\n\n' +
        '¿Quieres continuar de todos modos?',
        ui.ButtonSet.YES_NO
      );
      if (confirmNoPlaceholders !== ui.Button.YES) return;
    }

    const result = sendCertificateEmails(
      sheet, signedFolder, emailSubject, emailBody, emailMapping, criticalColumns
    );

    // Log any errors
    if (result.errors.length > 0) {
      logErrors(result.errors);
    }

    let message =
      `✅ Mandados con éxito: ${result.success}\n` +
      `❌ Fallos de envío: ${result.failed}\n` +
      (result.notFound > 0 ? `⚠️ PDFs no encontrados: ${result.notFound}\n` : '');

    if (result.errors.length > 0) {
      message += `\n⚠️ Se han registrado ${result.errors.length} error(es) en la hoja "${ERROR_SHEET_NAME}".\n` +
                 `Revísala para identificar los emails fallidos y corregirlos o enviarlos manualmente.`;
    }

    ui.alert('¡Envío Completado!', message, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('Error', `Ha ocurrido un error: ${error.toString()}`, ui.ButtonSet.OK);
    Logger.log('Error in sendCertificates: ' + error.toString());
  }
}

/**
 * Retry sending emails only for rows that appear as pending in the error log.
 */
function retryFailedEmails() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(ERROR_SHEET_NAME);

  if (!errorSheet || errorSheet.getLastRow() <= 1) {
    ui.alert('Sin errores', 'No hay errores pendientes en el registro.', ui.ButtonSet.OK);
    return;
  }

  const errorData = errorSheet.getDataRange().getValues();
  const pendingRows = [];

  for (let r = 1; r < errorData.length; r++) {
    const tipo   = errorData[r][1];
    const estado = errorData[r][8];
    if ((tipo === 'Email No Enviado' || tipo === 'PDF No Encontrado') && estado === 'Pendiente') {
      pendingRows.push({ sheetRow: r + 1, originalDataRow: parseInt(errorData[r][2]) - 1 });
    }
  }

  if (pendingRows.length === 0) {
    ui.alert('Sin errores', 'No hay errores pendientes de envío de emails.', ui.ButtonSet.OK);
    return;
  }

  ui.alert('Reintentar Emails',
    `Se van a reintentar ${pendingRows.length} email(s) fallido(s).\n\nNecesitarás introducir de nuevo la carpeta de firmados y el contenido del email.`,
    ui.ButtonSet.OK);

  const mainSheet = ss.getActiveSheet();

  const criticalColumns = identifyCriticalColumnsForSending(ui, mainSheet);
  if (!criticalColumns) return;

  const signedFolder = getSignedFolder(ui);
  if (!signedFolder) return;

  const emailSubject = getEmailSubject(ui);
  if (!emailSubject) return;

  const emailBody = getEmailBody(ui);
  if (!emailBody) return;

  const emailPlaceholders = detectEmailPlaceholders(emailSubject + ' ' + emailBody);
  let emailMapping = {};
  if (emailPlaceholders.length > 0) {
    // Show detected placeholders for the user to verify before mapping
    const placeholderList = emailPlaceholders.map(p => `  • {{${p}}}`).join('\n');
    const confirmPlaceholders = ui.alert(
      '✅ Placeholders Detectados en el Email',
      `Se han encontrado ${emailPlaceholders.length} placeholder(s) en el asunto y cuerpo del email:\n\n` +
      `${placeholderList}\n\n` +
      `¿Son correctos? Pulsa OK para continuar asignándolos a columnas,\n` +
      `o Cancelar para volver a escribir el asunto/cuerpo del email.`,
      ui.ButtonSet.OK_CANCEL
    );
    if (confirmPlaceholders !== ui.Button.OK) return;

    emailMapping = mapPlaceholdersToColumns(ui, mainSheet, emailPlaceholders);
    if (!emailMapping) return;
  } else {
    const confirmNoPlaceholders = ui.alert(
      '⚠️ Sin Placeholders',
      'No se han detectado placeholders ({{...}}) en el asunto ni en el cuerpo del email.\n\n' +
      'El mismo texto se enviará a todos los destinatarios sin personalización.\n\n' +
      '¿Quieres continuar de todos modos?',
      ui.ButtonSet.YES_NO
    );
    if (confirmNoPlaceholders !== ui.Button.YES) return;
  }

  // Index files in signed folder
  const files = {};
  const fileIterator = signedFolder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files[file.getName()] = file;
  }

  const allData = mainSheet.getDataRange().getValues();
  let success = 0;
  let stillFailing = 0;
  const newErrors = [];

  for (const pending of pendingRows) {
    const rowData = allData[pending.originalDataRow];
    const email   = rowData[criticalColumns.emailCol];
    const name    = rowData[criticalColumns.nameCol];
    const surname = rowData[criticalColumns.surnameCol];
    const filename = `${name}${surname}_signed.pdf`.replace(/\s+/g, '');

    try {
      if (!email) throw new Error('Dirección de email vacía');

      if (!files[filename]) throw new Error(`PDF no encontrado: ${filename}`);

      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(emailMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }

      let personalizedSubject = emailSubject;
      let personalizedBody    = emailBody;
      for (const [placeholder, value] of Object.entries(replacements)) {
        const regex = new RegExp(`\\{\\{${placeholder}\\}\\}`, 'g');
        personalizedSubject = personalizedSubject.replace(regex, value);
        personalizedBody    = personalizedBody.replace(regex, value);
      }

      MailApp.sendEmail({
        to: email,
        subject: personalizedSubject,
        body: personalizedBody,
        attachments: [files[filename].getBlob()]
      });

      markErrorResolved(errorSheet, pending.sheetRow);
      success++;

    } catch (err) {
      Logger.log(`Retry email failed for data row ${pending.originalDataRow + 1}: ${err}`);
      stillFailing++;
      newErrors.push({
        type: 'Email No Enviado',
        rowIndex: pending.originalDataRow,
        name: name,
        surname: surname,
        email: email,
        pdfName: filename,
        reason: err.toString()
      });
    }
  }

  if (newErrors.length > 0) logErrors(newErrors);

  ui.alert('Reintento Completado',
    `✅ Resueltos: ${success}\n❌ Siguen fallando: ${stillFailing}` +
    (stillFailing > 0 ? `\n\nRevisa "${ERROR_SHEET_NAME}" para más detalles.` : ''),
    ui.ButtonSet.OK);
}

/**
 * Identifies critical columns for sending (NAME, SURNAME, EMAIL)
 */
function identifyCriticalColumnsForSending(ui, sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersList = headers.map((h, i) => `${i + 1}. ${h}`).join('\n');

  ui.alert('Configuración: Columnas Críticas',
    'Primero, identificad las columas que contienen NOMBRE, APELLIDOS y EMAIL.\n\n' +
    'Estas se necesitan para comparar con los PDFs firmados y mandar los emails.\n\n' +
    '¡Puedes usar cualquier idioma para los nombres de tus columnas!',
    ui.ButtonSet.OK);

  const criticalColumns = {};

  const nameResponse = ui.prompt(
    'Columna de Nombre',
    `Columnas disponibles:\n${headersList}\n\n` +
    'Introduce el NÚMERO de la columna que contiene los NOMBRES:\n' +
    '(e.g., "Name", "Nombre", "Prénom", "Nome", etc.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (nameResponse.getSelectedButton() !== ui.Button.OK) return null;
  const nameCol = parseInt(nameResponse.getResponseText().trim());
  if (isNaN(nameCol) || nameCol < 1 || nameCol > headers.length) {
    ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
    return null;
  }
  criticalColumns.nameCol = nameCol - 1;

  const surnameResponse = ui.prompt(
    'Columna de Apellidos',
    `Columnas disponibles:\n${headersList}\n\n` +
    'Introduce el NÚMERO de la columna que contiene los APELLIDOS:\n' +
    '(e.g., "Surname", "Apellidos", "Nom de famille", "Sobrenome", etc.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (surnameResponse.getSelectedButton() !== ui.Button.OK) return null;
  const surnameCol = parseInt(surnameResponse.getResponseText().trim());
  if (isNaN(surnameCol) || surnameCol < 1 || surnameCol > headers.length) {
    ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
    return null;
  }
  criticalColumns.surnameCol = surnameCol - 1;

  const emailResponse = ui.prompt(
    'Columna de Email',
    `Columnas disponibles:\n${headersList}\n\n` +
    'Introduce el NÚMERO de la columna que contiene las DIRECCIONES DE CORREO:\n' +
    '(e.g., "Email", "Correo", "E-mail", "Courriel", etc.)',
    ui.ButtonSet.OK_CANCEL
  );
  if (emailResponse.getSelectedButton() !== ui.Button.OK) return null;
  const emailCol = parseInt(emailResponse.getResponseText().trim());
  if (isNaN(emailCol) || emailCol < 1 || emailCol > headers.length) {
    ui.alert('Error', 'Número de columna inválido. Por favor inténtalo de nuevo.', ui.ButtonSet.OK);
    return null;
  }
  criticalColumns.emailCol = emailCol - 1;

  return criticalColumns;
}

function getSignedFolder(ui) {
  const response = ui.prompt(
    'Carpeta de Certificados Firmados',
    'Introduce la URL o ID de la Carpeta de Google Drive que contiene los PDFs firmados:\n\n' +
    '(Los archivos deben tener el nombre: NombreApellidos_signed.pdf)',
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) return null;
  const input = response.getResponseText().trim();
  const folderId = extractFolderId(input);
  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    ui.alert('Error', 'No se pudo acceder a la carpeta. Comprueba la URL/ID y los permisos.', ui.ButtonSet.OK);
    return null;
  }
}

function getEmailSubject(ui) {
  const response = ui.prompt(
    'Asunto del Email',
    'Introduce el asunto del email:\n\n' +
    '(Puedes usar placeholders como {{NOMBRE}}, no hace falta que coincida el nombre exactamente con el de la columna, luego los asignarás como en la plantilla)',
    ui.ButtonSet.OK_CANCEL
  );
  return response.getSelectedButton() === ui.Button.OK ? response.getResponseText() : null;
}

function getEmailBody(ui) {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      textarea { width: 100%; height: 300px; margin: 10px 0; padding: 10px; }
      .buttons { text-align: right; margin-top: 10px; }
      button { padding: 10px 20px; margin-left: 10px; }
    </style>
    <h3>Cuerpo del Email</h3>
    <p>Introduce el cuerpo del email. Puedes usar placeholders como {{NOMBRE}}, {{APELLIDOS}}, etc.</p>
    <textarea id="emailBody" placeholder="Querido {{NAME}},&#10;&#10;Te adjunto aquí tu certificado..."></textarea>
    <div class="buttons">
      <button onclick="cancel()">Cancelar</button>
      <button onclick="submit()" style="background: #4285f4; color: white; border: none;">OK</button>
    </div>
    <script>
      function submit() {
        const body = document.getElementById('emailBody').value;
        google.script.run.withSuccessHandler(() => google.script.host.close())
          .storeEmailBody(body);
      }
      function cancel() {
        google.script.host.close();
      }
    </script>
  `)
    .setWidth(500)
    .setHeight(450);

  ui.showModalDialog(htmlOutput, 'Configuración del cuerpo del Email');
  Utilities.sleep(1000);

  const startTime = Date.now();
  while (Date.now() - startTime < 120000) {
    const stored = PropertiesService.getScriptProperties().getProperty('temp_email_body');
    if (stored) {
      PropertiesService.getScriptProperties().deleteProperty('temp_email_body');
      return stored;
    }
    Utilities.sleep(500);
  }
  return null;
}

function storeEmailBody(body) {
  PropertiesService.getScriptProperties().setProperty('temp_email_body', body);
}

function detectEmailPlaceholders(text) {
  const regex = /\{\{([A-Za-z0-9_]+)\}\}/g;
  const placeholders = new Set();
  let match;
  while ((match = regex.exec(text)) !== null) {
    placeholders.add(match[1]);
  }
  return Array.from(placeholders).sort();
}

/**
 * Sends emails with signed certificate attachments.
 * Returns { success, failed, notFound, errors[] }
 */
function sendCertificateEmails(sheet, signedFolder, subject, body, placeholderMapping, criticalColumns) {
  const data = sheet.getDataRange().getValues();
  let successCount  = 0;
  let failedCount   = 0;
  let notFoundCount = 0;
  const errors = [];

  // Index all files for fast lookup
  const files = {};
  const fileIterator = signedFolder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files[file.getName()] = file;
  }

  for (let i = 1; i < data.length; i++) {
    const rowData = data[i];
    const email   = rowData[criticalColumns.emailCol];
    const name    = rowData[criticalColumns.nameCol];
    const surname = rowData[criticalColumns.surnameCol];
    const filename = `${name}${surname}_signed.pdf`.replace(/\s+/g, '');

    // --- Missing email ---
    if (!email) {
      Logger.log(`Row ${i + 1}: No email address`);
      failedCount++;
      errors.push({
        type: 'Email No Enviado',
        rowIndex: i,
        name: name,
        surname: surname,
        email: '',
        pdfName: filename,
        reason: 'Dirección de email vacía o ausente en la hoja'
      });
      continue;
    }

    // --- PDF not found ---
    if (!files[filename]) {
      Logger.log(`Row ${i + 1}: File not found: ${filename}`);
      notFoundCount++;
      errors.push({
        type: 'PDF No Encontrado',
        rowIndex: i,
        name: name,
        surname: surname,
        email: email,
        pdfName: filename,
        reason: `No se encontró el archivo "${filename}" en la carpeta de firmados`
      });
      continue;
    }

    // --- Send email ---
    try {
      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(placeholderMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }

      let personalizedSubject = subject;
      let personalizedBody    = body;
      for (const [placeholder, value] of Object.entries(replacements)) {
        const regex = new RegExp(`\\{\\{${placeholder}\\}\\}`, 'g');
        personalizedSubject = personalizedSubject.replace(regex, value);
        personalizedBody    = personalizedBody.replace(regex, value);
      }

      MailApp.sendEmail({
        to: email,
        subject: personalizedSubject,
        body: personalizedBody,
        attachments: [files[filename].getBlob()]
      });

      successCount++;

    } catch (error) {
      Logger.log(`Failed to send email for row ${i + 1}: ${error.toString()}`);
      failedCount++;
      errors.push({
        type: 'Email No Enviado',
        rowIndex: i,
        name: name,
        surname: surname,
        email: email,
        pdfName: filename,
        reason: error.toString()
      });
    }
  }

  return { success: successCount, failed: failedCount, notFound: notFoundCount, errors };
}

// ═══════════════════════════════════════════════════════════════════════════
// HELP & DOCUMENTATION
// ═══════════════════════════════════════════════════════════════════════════

function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
SISTEMA DE CERTIFICADOS - INSTRUCCIONES
═══════════════════════════════════════════

📋 PREPARACIÓN
1. Prepara tu hoja de cálculo con columnas: Nombre, Apellidos, Email (como mínimo)
2. Crea una plantilla de Google Docs con placeholders: {{NAME}}, {{SURNAME}}, {{EMAIL}}, etc.
3. Crea una carpeta de Google Drive para almacenar los certificados

📄 FASE 1: GENERAR CERTIFICADOS
1. Haz clic en «Sistema de certificados» > «Generar certificados»
2. Introduce el ID o la URL de tu plantilla
3. Revisa los placeholders detectados
4. Asigna cada placeholder a una columna de la hoja
5. Introduce el ID o la URL de la carpeta de salida
6. Los PDFs se guardarán en la subcarpeta «No firmados»

✍️ FIRMA MANUAL
1. Descarga los PDF de la carpeta «No firmados»
2. Fírmalos con la herramienta que prefieras
3. Renómbralos: NombreApellido_signed.pdf
4. Súbelos a una carpeta de Drive (por ejemplo, «Firmados»)

📧 FASE 2: ENVIAR CERTIFICADOS
1. Haz clic en «Sistema de certificados» > «Enviar certificados firmados»
2. Introduce el ID/URL de la carpeta con los PDFs firmados
3. Escribe el asunto y cuerpo del email (puedes usar placeholders)
4. Asigna los placeholders a las columnas
5. Espera a que se complete el envío

🔴 REGISTRO DE ERRORES
• Cuando hay fallos, se crea automáticamente la hoja "🔴 Registro de Errores"
• Muestra: fila original, nombre, apellidos, email, nombre del PDF esperado y motivo del error
• Usa «Reintentar Certificados Fallidos» o «Reintentar Emails Fallidos» para reprocessar solo los fallidos
• Los errores resueltos se marcan en verde ✅
• Usa «Limpiar Registro de Errores» para resetear el log

💡 CONSEJOS
• Los placeholders distinguen entre mayúsculas y minúsculas
• El nombre del PDF firmado debe ser exacto: NombreApellido_signed.pdf
• Límite de Gmail: ~1500 emails/día
• Prueba primero con 1-2 filas antes de procesar todas

❓ SOLUCIÓN DE PROBLEMAS
• «No se puede acceder al archivo»: comprueba los permisos para compartir
• «No se ha encontrado el placeholder»: comprueba el formato {{PLACEHOLDER}}
• «Error en el email»: comprueba los límites de cuota diarios
  `;

  ui.alert('Ayuda del Sistema de Certificados', helpText, ui.ButtonSet.OK);
}
