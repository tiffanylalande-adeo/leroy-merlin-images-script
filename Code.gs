/**
 * Récupère automatiquement la première image du carrousel
 * Leroy Merlin à partir de l'URL en colonne P
 * et l'insère dans la colonne A.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🖼️ Images produits')
    .addItem('Récupérer les images', 'recupererImages')
    .addItem('Récupérer uniquement la ligne sélectionnée', 'recupererImageLigne')
    .addToUi();
}


/**
 * Traite toutes les lignes contenant une URL en colonne P.
 */
function recupererImages() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Aucune donnée à traiter.');
    return;
  }

  const urls = sheet.getRange(2, 16, lastRow - 1, 1).getValues();
  let nbOK = 0;
  let nbErreur = 0;

  for (let i = 0; i < urls.length; i++) {
    const row = i + 2;
    const url = urls[i][0];

    if (!url || typeof url !== 'string') continue;

    const existingValue = sheet.getRange(row, 1).getValue();
    if (existingValue && existingValue.valueType === SpreadsheetApp.ValueType.IMAGE) continue;

    try {
      const imageUrl = trouverPremiereImage(url);

      if (imageUrl) {
        const image = SpreadsheetApp.newCellImage()
          .setSourceUrl(imageUrl)
          .setAltTextTitle('Image produit Leroy Merlin')
          .build();
        sheet.getRange(row, 1).setValue(image);
        sheet.setRowHeight(row, 110);
        nbOK++;
      } else {
        sheet.getRange(row, 1).setValue('❌ Image introuvable');
        nbErreur++;
      }
    } catch (e) {
      sheet.getRange(row, 1).setValue('❌ Erreur');
      nbErreur++;
      console.log('Erreur ligne ' + row + ' : ' + e.message);
    }

    Utilities.sleep(500);
  }

  SpreadsheetApp.getUi().alert(
    'Terminé !\n\n✅ Images récupérées : ' + nbOK + '\n❌ Erreurs : ' + nbErreur
  );
}


/**
 * Traite uniquement la ligne actuellement sélectionnée.
 */
function recupererImageLigne() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();

  if (row < 2) {
    SpreadsheetApp.getUi().alert('Sélectionne une ligne contenant une URL produit.');
    return;
  }

  const url = sheet.getRange(row, 16).getValue();

  if (!url) {
    SpreadsheetApp.getUi().alert("Il n'y a pas d'URL en colonne P sur cette ligne.");
    return;
  }

  try {
    const imageUrl = trouverPremiereImage(url);

    if (!imageUrl) {
      SpreadsheetApp.getUi().alert("❌ Impossible de trouver l'image du produit.");
      return;
    }

    const image = SpreadsheetApp.newCellImage()
      .setSourceUrl(imageUrl)
      .setAltTextTitle('Image produit Leroy Merlin')
      .build();

    sheet.getRange(row, 1).setValue(image);
    sheet.setRowHeight(row, 110);

    SpreadsheetApp.getUi().alert('✅ Image récupérée pour la ligne ' + row);
  } catch (e) {
    SpreadsheetApp.getUi().alert('❌ Erreur : ' + e.message);
  }
}


/**
 * Analyse la page Leroy Merlin et récupère
 * la première image du carrousel produit.
 */
function trouverPremiereImage(url) {
  const options = {
    muteHttpExceptions: true,
    followRedirects: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
      'Accept-Language': 'fr-FR,fr;q=0.9,en-US;q=0.8,en;q=0.7',
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache',
      'Referer': 'https://www.leroymerlin.fr/',
      'Connection': 'keep-alive'
    }
  };

  const response = UrlFetchApp.fetch(url, options);

  if (response.getResponseCode() !== 200) {
    throw new Error('La page renvoie le code HTTP ' + response.getResponseCode());
  }

  const code = response.getContentText();

  const patterns = [
    /<img[^>]+(?:src|data-src)=["'](https?:\/\/media\.adeo\.com[^"']+)["']/gi,
    /["'](https?:\/\/media\.adeo\.com[^"']+\.(?:jpg|jpeg|png|webp)[^"']*)["']/gi,
    /["'](https?:\/\/media\.adeo\.com[^"']+)["']/gi
  ];

  for (const pattern of patterns) {
    const images = [];
    let match;

    while ((match = pattern.exec(code)) !== null) {
      let imageUrl = match[1];
      imageUrl = imageUrl
        .replace(/\\u0026/g, '&')
        .replace(/\\u003D/g, '=')
        .replace(/\\\//g, '/')
        .replace(/&amp;/g, '&');

      if (!images.includes(imageUrl)) {
        images.push(imageUrl);
      }
    }

    if (images.length > 0) {
      return images[0];
    }
  }

  return null;
}
