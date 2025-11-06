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
    .addItem('ℹ️ Ayuda e Instrucciones', 'showHelp')
    .addToUi();
}

// ═══════════════════════════════════════════════════════════════════════════
// 1: CERTIFICATE GENERATION
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Main function to generate certificates
 * Guides user through template selection, placeholder mapping, and PDF generation
 */
function generateCertificates() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  try {
    // Step 1: Identify critical columns (NAME and SURNAME for PDF naming)
    const criticalColumns = identifyCriticalColumnsForGeneration(ui, sheet);
    if (!criticalColumns) return;
    
    // Step 2: Get template document
    const templateDoc = getTemplateDocument(ui);
    if (!templateDoc) return;
    
    // Step 3: Detect placeholders in template
    const placeholders = detectPlaceholders(templateDoc);
    if (placeholders.length === 0) {
      ui.alert('Error', 'No se han encontrado placeholders en la plantilla. Usa el formato: {{PLACEHOLDER}}', ui.ButtonSet.OK);
      return;
    }
    
    ui.alert('Placeholders Encontrados', 
      `Detectados ${placeholders.length} placeholders:\n${placeholders.join(', ')}\n\nSiguiente: Asígnalos a las columnas de tu hoja.`, 
      ui.ButtonSet.OK);
    
    // Step 4: Map placeholders to columns
    const columnMapping = mapPlaceholdersToColumns(ui, sheet, placeholders);
    if (!columnMapping) return;
    
    // Step 5: Get output folder
    const outputFolder = getOutputFolder(ui);
    if (!outputFolder) return;
    
    // Step 6: Create "Non signed" subfolder
    const unsignedFolder = getOrCreateFolder(outputFolder, 'No firmados');
    
    // Step 7: Generate certificates
    const result = generateCertificatePDFs(sheet, templateDoc, columnMapping, unsignedFolder, criticalColumns);
    
    // Show results
    ui.alert('¡Generación Completa!', 
      `✅ Generados correctamente: ${result.success}\n` +
      `❌ Fallidos: ${result.failed}\n\n` +
      `PDFs guardados en:\n${unsignedFolder.getName()}\n\n` +
      `URL de la carpeta:\n${unsignedFolder.getUrl()}`, 
      ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Un error ha ocurrido: ${error.toString()}`, ui.ButtonSet.OK);
    Logger.log('Error in generateCertificates: ' + error.toString());
  }
}

/**
 * Identifies critical columns at the start (NAME and SURNAME for PDF filenames)
 * Returns {nameCol: index, surnameCol: index}
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
  
  // Get NAME column
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
  
  // Get SURNAME column
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
 * Returns DocumentApp Document object
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

/**
 * Extracts document ID from URL or returns ID if already provided
 */
function extractDocumentId(input) {
  const urlMatch = input.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return urlMatch ? urlMatch[1] : input;
}

/**
 * Detects all {{PLACEHOLDER}} patterns in document
 * Returns array of unique placeholder names
 */
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

/**
 * Maps placeholders to sheet columns interactively
 * Returns object: {PLACEHOLDER: columnIndex}
 */
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
    
    mapping[placeholder] = colNum - 1; // Convert to 0-based index
  }
  
  return mapping;
}

/**
 * Prompts user for output folder
 * Returns Folder object
 */
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

/**
 * Extracts folder ID from URL or returns ID if already provided
 */
function extractFolderId(input) {
  const urlMatch = input.match(/\/folders\/([a-zA-Z0-9-_]+)/);
  return urlMatch ? urlMatch[1] : input;
}

/**
 * Gets or creates a subfolder
 */
function getOrCreateFolder(parentFolder, folderName) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * Generates PDF certificates for all rows
 * Returns {success: count, failed: count}
 */
function generateCertificatePDFs(sheet, templateDoc, columnMapping, outputFolder, criticalColumns) {
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  let successCount = 0;
  let failedCount = 0;
  
  // Process each row (skip header)
  for (let i = 1; i < data.length; i++) {
    try {
      const rowData = data[i];
      
      // Build replacement map
      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(columnMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }
      
      // Generate filename using the critical columns identified at the start
      const name = rowData[criticalColumns.nameCol] || 'Unknown';
      const surname = rowData[criticalColumns.surnameCol] || 'Unknown';
      const filename = `${name}${surname}.pdf`.replace(/\s+/g, '');
      
      // Create personalized document
      const personalizedDoc = createPersonalizedDocument(templateDoc, replacements);
      
      // Export to PDF
      const pdfBlob = convertDocToPdf(personalizedDoc);
      pdfBlob.setName(filename);
      
      // Save to folder
      outputFolder.createFile(pdfBlob);
      
      // Delete temporary document
      DriveApp.getFileById(personalizedDoc.getId()).setTrashed(true);
      
      successCount++;
      
    } catch (error) {
      Logger.log(`Failed to generate certificate for row ${i + 1}: ${error.toString()}`);
      failedCount++;
    }
  }
  
  return { success: successCount, failed: failedCount };
}

/**
 * Creates a copy of template and replaces placeholders
 * Returns Document object
 */
function createPersonalizedDocument(templateDoc, replacements) {
  // Create a copy
  const templateFile = DriveApp.getFileById(templateDoc.getId());
  const copyFile = templateFile.makeCopy('temp_certificate_' + Date.now());
  const copyDoc = DocumentApp.openById(copyFile.getId());
  const body = copyDoc.getBody();
  
  // Replace all placeholders
  for (const [placeholder, value] of Object.entries(replacements)) {
    body.replaceText(`\\{\\{${placeholder}\\}\\}`, value);
  }
  
  copyDoc.saveAndClose();
  return copyDoc;
}

/**
 * Converts a Google Doc to PDF blob
 */
function convertDocToPdf(doc) {
  const docId = doc.getId();
  const url = `https://docs.google.com/document/d/${docId}/export?format=pdf`;
  const token = ScriptApp.getOAuthToken();
  
  const response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' + token
    }
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
    // Step 1: Identify critical columns (NAME, SURNAME, and EMAIL for sending)
    const criticalColumns = identifyCriticalColumnsForSending(ui, sheet);
    if (!criticalColumns) return;
    
    // Step 2: Get signed certificates folder
    const signedFolder = getSignedFolder(ui);
    if (!signedFolder) return;
    
    // Step 3: Get email subject
    const emailSubject = getEmailSubject(ui);
    if (!emailSubject) return;
    
    // Step 4: Get email body
    const emailBody = getEmailBody(ui);
    if (!emailBody) return;
    
    // Step 5: Detect placeholders in email content
    const emailPlaceholders = detectEmailPlaceholders(emailSubject + ' ' + emailBody);
    
    // Step 6: Map email placeholders to columns
    let emailMapping = {};
    if (emailPlaceholders.length > 0) {
      emailMapping = mapPlaceholdersToColumns(ui, sheet, emailPlaceholders);
      if (!emailMapping) return;
    }
    
    // Step 7: Send emails
    const result = sendCertificateEmails(
      sheet, 
      signedFolder, 
      emailSubject, 
      emailBody, 
      emailMapping, 
      criticalColumns
    );
    
    // Show results
    ui.alert('¡Envío Completado!', 
      `✅ Mandados con éxito: ${result.success}\n` +
      `❌ Fallos: ${result.failed}\n` +
      (result.notFound > 0 ? `⚠️ PDFs no encontrados: ${result.notFound}\n` : ''), 
      ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Error', `Ha ocurrido un error: ${error.toString()}`, ui.ButtonSet.OK);
    Logger.log('Error in sendCertificates: ' + error.toString());
  }
}

/**
 * Identifies critical columns for sending (NAME, SURNAME, EMAIL)
 * Returns {nameCol: index, surnameCol: index, emailCol: index}
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
  
  // Get NAME column
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
  
  // Get SURNAME column
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
  
  // Get EMAIL column
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

/**
 * Prompts user for signed certificates folder
 */
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

/**
 * Prompts user for email subject (supports placeholders)
 */
function getEmailSubject(ui) {
  const response = ui.prompt(
    'Asunto del Email',
    'Introduce el asunto del email:\n\n' +
    '(Puedes usar placeholders como {{NOMBRE}}, no hace falta que coincida el nombre exactamente con el de la columna, luego los asignarás como en la plantilla)',
    ui.ButtonSet.OK_CANCEL
  );
  
  return response.getSelectedButton() === ui.Button.OK ? response.getResponseText() : null;
}

/**
 * Prompts user for email body using HTML dialog
 */
function getEmailBody(ui) {
  const htmlOutput = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; }
      textarea { width: 100%; height: 300px; margin: 10px 0; padding: 10px; }
      .buttons { text-align: right; margin-top: 10px; }
      button { padding: 10px 20px; margin-left: 10px; }
    </style>
    <h3>Cuerpo del Email</h3>
    <p>Introduce el cuerpo del email. Puedes usar placeholders como {{NOMBRE}}, {{APELLIDOS}}, etc (no hace falta que coincida el nombre exactamente con el de la columna, luego los asignarás como en la plantilla).</p>
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
  
  // Wait for user input (stored in PropertiesService)
  Utilities.sleep(1000); // Give dialog time to open
  
  // Poll for stored value (max 2 minutes)
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

/**
 * Stores email body temporarily (called from HTML dialog)
 */
function storeEmailBody(body) {
  PropertiesService.getScriptProperties().setProperty('temp_email_body', body);
}

/**
 * Detects placeholders in email subject/body
 */
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
 * Sends emails with signed certificate attachments
 */
function sendCertificateEmails(sheet, signedFolder, subject, body, placeholderMapping, criticalColumns) {
  const data = sheet.getDataRange().getValues();
  let successCount = 0;
  let failedCount = 0;
  let notFoundCount = 0;
  
  // Get all files in folder for faster lookup
  const files = {};
  const fileIterator = signedFolder.getFiles();
  while (fileIterator.hasNext()) {
    const file = fileIterator.next();
    files[file.getName()] = file;
  }
  
  // Process each row (skip header)
  for (let i = 1; i < data.length; i++) {
    try {
      const rowData = data[i];
      
      // Get critical values using the columns identified at the start
      const email = rowData[criticalColumns.emailCol];
      const name = rowData[criticalColumns.nameCol];
      const surname = rowData[criticalColumns.surnameCol];
      
      if (!email) {
        Logger.log(`Row ${i + 1}: No email address`);
        failedCount++;
        continue;
      }
      
      // Build filename: NameSurname_signed.pdf
      const filename = `${name}${surname}_signed.pdf`.replace(/\s+/g, '');
      
      // Find file
      if (!files[filename]) {
        Logger.log(`Row ${i + 1}: File not found: ${filename}`);
        notFoundCount++;
        continue;
      }
      
      const pdfFile = files[filename];
      
      // Build replacement map for email
      const replacements = {};
      for (const [placeholder, colIndex] of Object.entries(placeholderMapping)) {
        replacements[placeholder] = rowData[colIndex] || '';
      }
      
      // Replace placeholders in subject and body
      let personalizedSubject = subject;
      let personalizedBody = body;
      
      for (const [placeholder, value] of Object.entries(replacements)) {
        const regex = new RegExp(`\\{\\{${placeholder}\\}\\}`, 'g');
        personalizedSubject = personalizedSubject.replace(regex, value);
        personalizedBody = personalizedBody.replace(regex, value);
      }
      
      // Send email
      MailApp.sendEmail({
        to: email,
        subject: personalizedSubject,
        body: personalizedBody,
        attachments: [pdfFile.getBlob()]
      });
      
      successCount++;
      
    } catch (error) {
      Logger.log(`Failed to send email for row ${i + 1}: ${error.toString()}`);
      failedCount++;
    }
  }
  
  return { success: successCount, failed: failedCount, notFound: notFoundCount };
}

// ═══════════════════════════════════════════════════════════════════════════
// HELP & DOCUMENTATION
// ═══════════════════════════════════════════════════════════════════════════

/**
 * Shows help dialog with instructions
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
SISTEMA DE CERTIFICADOS - INSTRUCCIONES
═══════════════════════════════════════════

📋 PREPARACIÓN
1. Prepara tu hoja de cálculo de Google con las columnas: Nombre, Apellidos, Correo electrónico (como mínimo)
2. Crea una plantilla de Google Docs con placeholders: {{NAME}}, {{SURNAME}}, {{EMAIL}}, etc.
3. Crea una carpeta de Google Drive para almacenar los certificados

📄 FASE 1: GENERAR CERTIFICADOS
1. Haz clic en «Sistema de certificados» > «Generar certificados».
2. Introduce el ID o la URL de tu plantilla.
3. Revisa los placeholders detectados.
4. Asigna cada placeholder a una columna de la hoja.
5. Introduce el ID o la URL de la carpeta de salida.
6. Espera a que se complete la generación.
7. Los PDF se guardarán en la subcarpeta «Sin firmar».

✍️ FIRMA MANUAL
1. Descarga los PDF de la carpeta «Sin firmar».
2. Fírmalos con la herramienta que prefieras.
3. Cambia el nombre de los archivos firmados: NombreApellido_firmado.pdf.
4. Súbelos a una nueva carpeta (por ejemplo, «Firmados»).

📧 FASE 2: ENVIAR CERTIFICADOS
1. Haz clic en «Sistema de certificados» > «Enviar certificados firmados».
2. Introduce el ID/URL de la carpeta que contiene los PDF firmados.
3. Introduce el asunto del correo electrónico (puedes utilizar marcadores de posición).
4. Introduce el cuerpo del correo electrónico (puedes utilizar marcadores de posición).
5. Asigna los placeholders a las columnas.
6. Espera a que se complete el envío.

💡 CONSEJOS
• Los placeholders distinguen entre mayúsculas y minúsculas.
• El nombre del PDF debe ser exacto: NombreApellido_firmado.pdf
• Compruebe los límites de cuota: ~1500 correos electrónicos al día para Gmail.
• Pruebe primero con 1-2 filas antes de procesar todas.

❓ SOLUCIÓN DE PROBLEMAS
• «No se puede acceder al archivo»: compruebe los permisos para compartir.
• «No se ha encontrado el marcador de posición»: compruebe el formato {{PLACEHOLDER}}.
• «Error en el correo electrónico»: compruebe los límites de cuota diarios.
  `;
  
  ui.alert('Ayuda del Sistema de Certificados', helpText, ui.ButtonSet.OK);
}