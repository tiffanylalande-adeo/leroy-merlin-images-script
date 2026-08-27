// ============================================================
//  CONFIGURATION — à modifier selon votre sheet
// ============================================================

var CONFIG = {
  // ID du dossier principal sur Google Drive contenant les sous-dossiers images
  // (visible dans l'URL Drive : drive.google.com/drive/folders/XXX)
  DRIVE_FOLDER_ID: '1YVgp4tf_trcgcoKTUKinkV70C6rNJsRi',

  // Colonne de matching (numéro) : la valeur qui correspond au nom du sous-dossier
  // Ex: 11 = colonne K (SKU), 9 = colonne I (EAN)
  MATCHING_COLUMN: 11,

  // Colonne où insérer les images (numéro) : 1 = colonne A
  IMAGE_COLUMN: 1,

  // Hauteur des lignes avec image (en pixels)
  ROW_HEIGHT: 110
};

// ============================================================


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖼️ Images produits')
    .addItem('Insérer les images (toutes les lignes)', 'insererImages')
    .addItem('Insérer les images (lignes sélectionnées)', 'insererImagesSelection')
    .addItem('⚙️ Voir la configuration', 'afficherConfig')
    .addToUi();
}


/**
 * Affiche la configuration actuelle.
 */
function afficherConfig() {
  var ui = SpreadsheetApp.getUi();
  var message =
    'Configuration actuelle :\n\n' +
    '📁 ID Dossier Drive : "' + CONFIG.DRIVE_FOLDER_ID + '"\n' +
    '🔑 Colonne de matching : ' + CONFIG.MATCHING_COLUMN + ' (colonne ' + colonneLettre(CONFIG.MATCHING_COLUMN) + ')\n' +
    '🖼️ Colonne image : ' + CONFIG.IMAGE_COLUMN + ' (colonne ' + colonneLettre(CONFIG.IMAGE_COLUMN) + ')\n\n' +
    'Pour modifier, changez les valeurs dans la section CONFIG en haut du script.';
  ui.alert('⚙️ Configuration', message, ui.ButtonSet.OK);
}


/**
 * Convertit un numéro de colonne en lettre (ex: 11 → K).
 */
function colonneLettre(n) {
  var lettre = '';
  while (n > 0) {
    var reste = (n - 1) % 26;
    lettre = String.fromCharCode(65 + reste) + lettre;
    n = Math.floor((n - 1) / 26);
  }
  return lettre;
}


/**
 * Récupère le dossier principal via son ID Google Drive.
 */
function getDossierPrincipal() {
  try {
    return DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
  } catch (e) {
    throw new Error(
      'Dossier introuvable avec l\'ID : "' + CONFIG.DRIVE_FOLDER_ID + '".\n' +
      'Vérifiez l\'ID dans la configuration.'
    );
  }
}


/**
 * Cherche un fichier image dans un sous-dossier nommé par la référence.
 * Retourne le fichier Drive ou null si introuvable.
 */
function trouverFichierImage(dossierPrincipal, reference) {
  reference = reference.toString().trim();

  var subFolders = dossierPrincipal.getFoldersByName(reference);
  if (!subFolders.hasNext()) return null;

  var subFolder = subFolders.next();

  // Cherche EAN.jpg / EAN.jpeg / EAN.png / EAN.webp
  var extensions = ['.jpg', '.jpeg', '.png', '.webp'];
  for (var i = 0; i < extensions.length; i++) {
    var files = subFolder.getFilesByName(reference + extensions[i]);
    if (files.hasNext()) return files.next();
  }

  // Fallback : premier fichier image du dossier
  var allFiles = subFolder.getFiles();
  while (allFiles.hasNext()) {
    var file = allFiles.next();
    var mime = file.getMimeType();
    if (mime === 'image/jpeg' || mime === 'image/png' || mime === 'image/webp') {
      return file;
    }
  }

  return null;
}


/**
 * Insère une image Drive dans une cellule via blob base64.
 */
function insererImageDansCellule(sheet, row, file) {
  var blob = file.getBlob();
  var imageData = Utilities.base64Encode(blob.getBytes());
  var mimeType = blob.getContentType();

  var imageUrl = 'data:' + mimeType + ';base64,' + imageData;

  var image = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageUrl)
    .setAltTextTitle('Image produit')
    .build();

  sheet.getRange(row, CONFIG.IMAGE_COLUMN).setValue(image);
  sheet.setRowHeight(row, CONFIG.ROW_HEIGHT);
}


/**
 * Fonction commune de traitement d'une liste de numéros de lignes.
 */
function traiterLignes(rows) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  var dossierPrincipal;
  try {
    dossierPrincipal = getDossierPrincipal();
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ ' + e.message);
    return;
  }

  var nbOK = 0;
  var nbErreur = 0;
  var nbSaute = 0;

  for (var i = 0; i < rows.length; i++) {
    var row = rows[i];
    if (row < 2) continue; // ignore l'en-tête

    var ref = sheet.getRange(row, CONFIG.MATCHING_COLUMN).getValue();
    if (!ref) continue;

    // Ne pas écraser une image existante
    var existingValue = sheet.getRange(row, CONFIG.IMAGE_COLUMN).getValue();
    if (existingValue && existingValue !== '') {
      nbSaute++;
      continue;
    }

    try {
      var file = trouverFichierImage(dossierPrincipal, ref);

      if (file) {
        insererImageDansCellule(sheet, row, file);
        nbOK++;
      } else {
        sheet.getRange(row, CONFIG.IMAGE_COLUMN).setValue('❌ Image introuvable');
        nbErreur++;
      }
    } catch (e) {
      sheet.getRange(row, CONFIG.IMAGE_COLUMN).setValue('❌ Erreur : ' + e.message);
      nbErreur++;
      console.log('Erreur ligne ' + row + ' : ' + e.message);
    }
  }

  SpreadsheetApp.getUi().alert(
    'Terminé !\n\n' +
    '✅ Images insérées : ' + nbOK + '\n' +
    '⏭️ Lignes ignorées (déjà remplies) : ' + nbSaute + '\n' +
    '❌ Images introuvables : ' + nbErreur
  );
}


/**
 * Traite toutes les lignes du sheet.
 */
function insererImages() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Aucune donnée à traiter.');
    return;
  }

  var rows = [];
  for (var i = 2; i <= lastRow; i++) {
    rows.push(i);
  }

  traiterLignes(rows);
}


/**
 * Traite uniquement les lignes sélectionnées (une ou plusieurs).
 */
function insererImagesSelection() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var selection = sheet.getActiveRange();

  var firstRow = selection.getRow();
  var lastRow = firstRow + selection.getNumRows() - 1;

  if (firstRow < 2 && lastRow < 2) {
    SpreadsheetApp.getUi().alert('Sélectionne au moins une ligne de données (pas l\'en-tête).');
    return;
  }

  var rows = [];
  for (var r = firstRow; r <= lastRow; r++) {
    rows.push(r);
  }

  traiterLignes(rows);
}
