/**
 * ============================================================================
 * Project: QantIA Technology Suite - Master Data Cleaner - Unit Pro
 * Author: Steven - QantIA Services
 * Copyright: (c) 2026 QantIA Services. All rights reserved.
 * License: Commercial - Integration Service
 * Website: https://www.upwork.com/services/qantia
 * ============================================================================
 */

/**
 * Adds the QantIA Automation custom menu when the spreadsheet is opened.
 */
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('🚀 QantIA Automation')
      .addItem('Master Data Cleaner — Panneau', 'showMasterDataCleanerSidebar')
      .addSeparator()
      .addItem('Nettoyer les noms', 'cleanNames')
      .addItem('Normaliser les téléphones', 'cleanPhones')
      .addItem('Supprimer les doublons (colonne active)', 'removeDuplicates')
      .addItem('Supprimer espaces invisibles', 'trimWhiteSpaces')
      .addToUi();
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Menu',
      'Impossible d’ajouter le menu : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Opens the Master Data Cleaner sidebar (Material UI).
 */
function showMasterDataCleanerSidebar() {
  try {
    var html = HtmlService.createHtmlOutputFromFile('UI_Sidebar')
      .setTitle('Master Data Cleaner — Unit Pro')
      .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Panneau',
      'Ouverture du panneau impossible : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Runs selected cleaning steps in a safe order (trim → names → phones → duplicates).
 *
 * @param {Object} options Options from the sidebar.
 * @param {boolean} [options.trimWhiteSpaces]
 * @param {boolean} [options.cleanNames]
 * @param {boolean} [options.cleanPhones]
 * @param {boolean} [options.removeDuplicates]
 * @return {{ok: boolean, message: string}}
 */
function runMasterClean(options) {
  try {
    options = options || {};
    if (!SpreadsheetApp.getActiveSheet().getActiveRange()) {
      return { ok: false, message: 'Sélectionnez une plage de cellules à traiter.' };
    }
    var steps = [];
    if (options.trimWhiteSpaces) {
      trimWhiteSpaces(true);
      steps.push('espaces');
    }
    if (options.cleanNames) {
      cleanNames(true);
      steps.push('noms');
    }
    if (options.cleanPhones) {
      cleanPhones(true);
      steps.push('téléphones');
    }
    if (options.removeDuplicates) {
      removeDuplicates(true);
      steps.push('doublons');
    }
    if (steps.length === 0) {
      return { ok: false, message: 'Sélectionnez au moins une option de nettoyage.' };
    }
    return {
      ok: true,
      message: 'Traitement terminé : ' + steps.join(', ') + '.',
    };
  } catch (err) {
    return { ok: false, message: err.message || String(err) };
  }
}

/**
 * Title-cases name-like text (e.g. JEAN DUPONT → Jean Dupont).
 *
 * @param {boolean} [silent] If true, skips success alert (for batch runs from the sidebar).
 */
function cleanNames(silent) {
  try {
    var ctx = getActiveDataRangeOrAlert_();
    if (!ctx) {
      return;
    }
    var values = ctx.range.getValues();
    var out = values.map(function (row) {
      return row.map(function (cell) {
        if (cell === '' || cell === null || cell === undefined) {
          return cell;
        }
        if (typeof cell === 'string') {
          return titleCaseName_(cell);
        }
        if (typeof cell === 'number') {
          return cell;
        }
        return titleCaseName_(String(cell));
      });
    });
    ctx.range.setValues(out);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Noms',
        'Les noms de la plage sélectionnée ont été mis en forme.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Noms',
      'Erreur : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Normalizes phone values to (AAA) EEE-NNNN when 10 North American digits are detected.
 *
 * @param {boolean} [silent] If true, skips success alert (for batch runs from the sidebar).
 */
function cleanPhones(silent) {
  try {
    var ctx = getActiveDataRangeOrAlert_();
    if (!ctx) {
      return;
    }
    var values = ctx.range.getValues();
    var out = values.map(function (row) {
      return row.map(function (cell) {
        if (cell === '' || cell === null || cell === undefined) {
          return cell;
        }
        return formatNorthAmericanPhone_(cell);
      });
    });
    ctx.range.setValues(out);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Téléphones',
        'Les numéros reconnus (10 chiffres) ont été normalisés.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Téléphones',
      'Erreur : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Removes duplicate rows within the active range using the sheet’s active column as key.
 * The first occurrence (top) is kept.
 *
 * @param {boolean} [silent] If true, skips success alert (for batch runs from the sidebar).
 */
function removeDuplicates(silent) {
  try {
    var ctx = getActiveDataRangeOrAlert_();
    if (!ctx) {
      return;
    }
    var sheet = ctx.sheet;
    var range = ctx.range;
    var keyCol = sheet.getActiveCell().getColumn();
    var startRow = range.getRow();
    var startCol = range.getColumn();
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    if (numRows < 2) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Doublons',
        'Sélectionnez au moins deux lignes pour supprimer des doublons.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    var relKeyCol = keyCol - startCol;
    if (relKeyCol < 0 || relKeyCol >= numCols) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Doublons',
        'Placez la cellule active dans une colonne comprise dans la plage sélectionnée (colonne clé).',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
      return;
    }
    var values = range.getValues();
    var seen = Object.create(null);
    var rowsToDelete = [];
    var r;
    for (r = 0; r < values.length; r++) {
      var key = duplicateKey_(values[r][relKeyCol]);
      if (seen[key]) {
        rowsToDelete.push(startRow + r);
      } else {
        seen[key] = true;
      }
    }
    rowsToDelete.sort(function (a, b) {
      return b - a;
    });
    for (r = 0; r < rowsToDelete.length; r++) {
      sheet.deleteRow(rowsToDelete[r]);
    }
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Doublons',
        rowsToDelete.length + ' ligne(s) en double supprimée(s).',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Doublons',
      'Erreur : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * Trims cells and replaces invisible / exotic whitespace with a single ASCII space.
 *
 * @param {boolean} [silent] If true, skips success alert (for batch runs from the sidebar).
 */
function trimWhiteSpaces(silent) {
  try {
    var ctx = getActiveDataRangeOrAlert_();
    if (!ctx) {
      return;
    }
    var values = ctx.range.getValues();
    var out = values.map(function (row) {
      return row.map(function (cell) {
        if (cell === '' || cell === null || cell === undefined) {
          return cell;
        }
        if (typeof cell === 'string') {
          return normalizeWhitespace_(cell);
        }
        if (typeof cell === 'number') {
          return cell;
        }
        return normalizeWhitespace_(String(cell));
      });
    });
    ctx.range.setValues(out);
    if (!silent) {
      SpreadsheetApp.getUi().alert(
        'QantIA — Espaces',
        'Les espaces invisibles de la plage ont été nettoyés.',
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  } catch (err) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Espaces',
      'Erreur : ' + err.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

/**
 * @return {{sheet: GoogleAppsScript.Spreadsheet.Sheet, range: GoogleAppsScript.Spreadsheet.Range}|null}
 * @private
 */
function getActiveDataRangeOrAlert_() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var range = sheet.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert(
      'QantIA — Plage',
      'Sélectionnez une plage de cellules à traiter.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return null;
  }
  return { sheet: sheet, range: range };
}

/**
 * @param {string} s
 * @return {string}
 * @private
 */
function titleCaseName_(s) {
  var t = normalizeWhitespace_(s);
  if (!t) {
    return t;
  }
  return t
    .toLowerCase()
    .replace(/(^|[\s'-])(.)/gu, function (_, sep, ch) {
      return sep + ch.toUpperCase();
    });
}

/**
 * @param {*} value
 * @return {string}
 * @private
 */
function formatNorthAmericanPhone_(value) {
  var raw = String(value).replace(/\D/g, '');
  if (raw.length === 11 && raw.charAt(0) === '1') {
    raw = raw.substring(1);
  }
  if (raw.length === 10) {
    return '(' + raw.substring(0, 3) + ') ' + raw.substring(3, 6) + '-' + raw.substring(6);
  }
  return typeof value === 'string' ? value : String(value);
}

/**
 * @param {string} s
 * @return {string}
 * @private
 */
function normalizeWhitespace_(s) {
  if (!s) {
    return s;
  }
  var cleaned = String(s).replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, ' ');
  cleaned = cleaned.replace(/\s+/g, ' ').trim();
  return cleaned;
}

/**
 * @param {*} cellValue
 * @return {string}
 * @private
 */
function duplicateKey_(cellValue) {
  if (cellValue === null || cellValue === undefined) {
    return '';
  }
  return normalizeWhitespace_(String(cellValue)).toLowerCase();
}
