/*
 * Demo / sanitized Google Apps Script tool for localization QA workflows.
 
 * This script checks one selected locale column against the English source column.
 * It validates numbers, tier names, event names, and month names, then highlights
 * potentially inconsistent translations and adds review notes.
 */

const VALIDATOR_CONFIG = {

  SOURCE_START_ROW: 5,
  SOURCE_COLUMN: 1,

  FIRST_LOCALE_COLUMN: 5,
  SPECIAL_LEVEL_COLUMNS: [7, 8],
  NUMERIC_MONTH_LOCALE_COLUMNS: [10, 11],

  ERROR_COLOR: '#FF9999',
  CLEAR_COLOR: '#FFFFFF',

  NOTES: {
    NUMBERS: 'Check numbers',
    TIER_NAMES: 'Check tier names',
    TIER_ORDER: 'Check tier order',
    EVENT_NAMES: 'Check event names',
    MONTHS: 'Check months'
  }
};

function validateLocalizationConsistency() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Localization QA');

  if (!sheet) 
    {
      throw new Error(`Sheet not found: Localization QA`);
    }

  const lastRow = sheet.getDataRange().getLastRow();
  const rowCount = lastRow - VALIDATOR_CONFIG.SOURCE_START_ROW + 1;

  if (rowCount <= 0) 
    {
      return;
    }

  const selectedLocaleColumn = sheet.getActiveRange().getColumn();

  const sourceValues = sheet.getRange(VALIDATOR_CONFIG.SOURCE_START_ROW, VALIDATOR_CONFIG.SOURCE_COLUMN, rowCount, 1).getValues();

  const localeValues = sheet.getRange(VALIDATOR_CONFIG.SOURCE_START_ROW, selectedLocaleColumn, rowCount, 1).getValues();

  const localeRange = sheet.getRange(VALIDATOR_CONFIG.SOURCE_START_ROW, selectedLocaleColumn, rowCount, 1);
  const colors = localeRange.getBackgrounds();
  const notes = localeRange.getNotes();

  validateNumbers(sourceValues, localeValues, selectedLocaleColumn, colors, notes);
  validateTierNames(sourceValues, localeValues, selectedLocaleColumn, colors, notes);
  validateEventNames(sourceValues, localeValues, colors, notes);

  if (!VALIDATOR_CONFIG.NUMERIC_MONTH_LOCALE_COLUMNS.includes(selectedLocaleColumn)) 
    {
      validateMonthNames(sourceValues, localeValues, selectedLocaleColumn, colors, notes);
    }
  else
    {
      validateMonthNames(sourceValues, localeValues, selectedLocaleColumn, colors, notes);
    }

  localeRange.setBackgrounds(colors);
  localeRange.setNotes(notes);
}

//Validates numeric consistency between source and selected locale
function validateNumbers(sourceValues, localeValues, selectedLocaleColumn, colors, notes) {
  for (let i = 0; i < sourceValues.length; i++) 
    {
      const sourceText = sourceValues[i].toString();
      const localeText = localeValues[i].toString();

      const normalizedSourceText = normalizeLevelNames(sourceText);

      const sourceNumbers = VALIDATOR_CONFIG.SPECIAL_LEVEL_COLUMNS.includes(selectedLocaleColumn)
        ? extractNumbersFromText(normalizedSourceText)
        : extractNumbersFromText(sourceText);

      const localeNumbers = extractDigitNumbers(localeText);

      if (!arraysEqual(sourceNumbers, localeNumbers)) 
        {
          colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
          addNote(notes, i, VALIDATOR_CONFIG.NOTES.NUMBERS);
        } 
        else 
        {
          colors[i][0] = null;
          notes[i][0] = '';
        }
    }

  // Some locale columns use numbers instead of translated month names
  // If source month names are equivalent to locale numbers, clear false positives
  if (VALIDATOR_CONFIG.NUMERIC_MONTH_LOCALE_COLUMNS.includes(selectedLocaleColumn)) 
    {
      for (let i = 0; i < sourceValues.length; i++) 
        {
          const sourceNumbersWithMonths = extractNumbersFromTextWithMonths(sourceValues[i][0].toString());
          const localeNumbersForMonths = extractDigitNumbers(localeValues[i].toString());

          if (arraysEqual(sourceNumbersWithMonths, localeNumbersForMonths)) 
            {
              colors[i][0] = VALIDATOR_CONFIG.CLEAR_COLOR;

              if (notes[i][0] === VALIDATOR_CONFIG.NOTES.NUMBERS) 
                {
                  notes[i][0] = '';
                }
            }
        }
    }
}


function validateTierNames(sourceValues, localeValues, selectedLocaleColumn, colors, notes) {
  const tierTranslations = getTierTranslations();
  const localeIndex = selectedLocaleColumn - VALIDATOR_CONFIG.FIRST_LOCALE_COLUMN;

  for (let i = 0; i < sourceValues.length; i++) 
    {
      const sourceText = sourceValues[i].toString();
      const localeText = localeValues[i].toString();

      for (const key in tierTranslations) 
        {
          if (sourceText.includes(key)) 
            {
              const expectedTranslation = tierTranslations[key][localeIndex];

              if (!localeText.includes(expectedTranslation)) 
                {
                  colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
                  addNote(notes, i, VALIDATOR_CONFIG.NOTES.TIER_NAMES);
                } 
                else 
                {
                  for (const extraKey in tierTranslations) 
                    {
                      const extraTranslation = tierTranslations[extraKey][localeIndex];

                      if (localeText.includes(extraTranslation) && !sourceText.includes(extraKey)) 
                        {
                          colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
                          addNote(notes, i, VALIDATOR_CONFIG.NOTES.TIER_NAMES);
                        }
                    }
                }
            }
        }

    const sourceTierOrder = Object.keys(tierTranslations)
      .filter(key => sourceText.includes(key))
      .map(key => ({ key, index: sourceText.indexOf(key) }))
      .sort((a, b) => a.index - b.index)
      .map(item => item.key);

    const localeTierOrder = Object.keys(tierTranslations)
      .filter(key => localeText.includes(tierTranslations[key][localeIndex]))
      .map(key => ({ key, index: localeText.indexOf(tierTranslations[key][localeIndex]) }))
      .sort((a, b) => a.index - b.index)
      .map(item => item.key);

    for (let orderIndex = 0; orderIndex < Math.min(sourceTierOrder.length, localeTierOrder.length); orderIndex++) 
      {
        if (sourceTierOrder[orderIndex] !== localeTierOrder[orderIndex]) 
          {
            colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
            addNote(notes, i, VALIDATOR_CONFIG.NOTES.TIER_ORDER);
            break;
          }
      }
  }
}


function validateEventNames(sourceValues, localeValues, colors, notes) {
  const eventNameTranslations = getEventNameTranslations();

  for (let i = 0; i < sourceValues.length; i++) 
    {
      const sourceText = sourceValues[i].toString();
      const localeText = localeValues[i].toString().replace(/\s+/g, ' ').normalize('NFC');

      const match = sourceText.match(/"([^"]+)"/);

      if (!match) 
        {
          continue;
        }

      const eventName = match[1];

      if (!(eventName in eventNameTranslations)) 
        {
          continue;
        } 

      const validTranslations = eventNameTranslations[eventName];
      let foundValidTranslation = false;
      let extraTranslationFound = false;

      for (const validTranslation of validTranslations) 
        {
          const validRegex = new RegExp(`(^|\\s|[\\W])${escapeRegExp(validTranslation)}(?=$|\\s|[\\W])`, 'i');

          if (localeText.match(validRegex)) 
            {
              foundValidTranslation = true;
              break;
            }
        }

      for (const key in eventNameTranslations) 
        {
          for (const extraTranslation of eventNameTranslations[key]) 
            {
              const extraRegex = new RegExp(`(^|\\s|[\\W])${escapeRegExp(extraTranslation)}(?=$|\\s|[\\W])`, 'i');

              if (localeText.match(extraRegex) && key !== eventName && !sourceText.includes(key)) 
                {
                  extraTranslationFound = true;
                  break;
                }
            }

          if (extraTranslationFound) 
            {
              break;
            }
        }

      if (!foundValidTranslation || extraTranslationFound) 
        {
          colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
          addNote(notes, i, VALIDATOR_CONFIG.NOTES.EVENT_NAMES);
        }
    }
}


function validateMonthNames(sourceValues, localeValues, selectedLocaleColumn, colors, notes) {
  const monthTranslations = getMonthTranslations();
  const localeIndex = selectedLocaleColumn - VALIDATOR_CONFIG.FIRST_LOCALE_COLUMN;

  for (let i = 0; i < sourceValues.length; i++) 
    {
      const sourceText = sourceValues[i].toString();
      const localeText = localeValues[i].toString();

      for (const key in monthTranslations) 
        {
          if (sourceText.includes(key)) 
            {
              const expectedMonthTranslation = monthTranslations[key][localeIndex];

              if (!localeText.includes(expectedMonthTranslation)) 
                {
                  colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
                  addNote(notes, i, VALIDATOR_CONFIG.NOTES.MONTHS);
                } 
                else 
                {
                  for (const extraKey in monthTranslations) 
                    {
                      const extraMonthTranslation = monthTranslations[extraKey][localeIndex];

                      if (localeText.includes(extraMonthTranslation) && !sourceText.includes(extraKey)) 
                        {
                          colors[i][0] = VALIDATOR_CONFIG.ERROR_COLOR;
                          addNote(notes, i, VALIDATOR_CONFIG.NOTES.MONTHS);
                        }
                    }
                }
            }
        }
    }
}

//Adds a note like "Check numbers, Check tier names, Check event names, Check months"
function addNote(notes, rowIndex, message) {
  if (!notes[rowIndex][0]) {
    notes[rowIndex][0] = message;
    return;
  }

  if (notes[rowIndex][0].indexOf(message) === -1) {
    notes[rowIndex][0] += ',\n' + message;
  }
}

//Normalizes hyphenated level names before extracting numbers to make sure all the numbers are extracted correctly
function normalizeLevelNames(text) {
  return text
    .replace(/level[-\s]?five/gi, 'level five')
    .replace(/level[-\s]?four/gi, 'level four')
    .replace(/level[-\s]?three/gi, 'level three')
    .replace(/level[-\s]?two/gi, 'level two')
    .replace(/level[-\s]?one/gi, 'level one');
}


function extractDigitNumbers(text) {
  return (text.match(/\d+/g) || [])
    .map(Number)
    .sort((a, b) => a - b);
}

function extractNumbersFromText(text) {
  const numberWords = {
    'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
    'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
    'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14, 'fifteen': 15,
    'sixteen': 16, 'seventeen': 17, 'eighteen': 18, 'nineteen': 19, 'twenty': 20,
    'iv': 4, 'v': 5, 'vi': 6, 'vii': 7, 'viii': 8, 'ix': 9
  };

  const matches = text.toLowerCase().match(/\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|iv|v|vi|vii|viii|ix)\b|\d+/g) || [];

  return matches
    .map(word => numberWords[word] || parseInt(word))
    .sort((a, b) => a - b);
}

function extractNumbersFromTextWithMonths(text) {
  const numberWords = {
    'one': 1, 'two': 2, 'three': 3, 'four': 4, 'five': 5,
    'six': 6, 'seven': 7, 'eight': 8, 'nine': 9, 'ten': 10,
    'eleven': 11, 'twelve': 12, 'thirteen': 13, 'fourteen': 14, 'fifteen': 15,
    'sixteen': 16, 'seventeen': 17, 'eighteen': 18, 'nineteen': 19, 'twenty': 20,
    'january': 1, 'february': 2, 'march': 3, 'april': 4,
    'may': 5, 'june': 6, 'july': 7, 'august': 8, 'september': 9,
    'october': 10, 'november': 11, 'december': 12,
    'iv': 4, 'v': 5, 'vi': 6, 'vii': 7, 'viii': 8, 'ix': 9
  };

  const matches = text.toLowerCase().match(/\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty|january|february|march|april|may|june|july|august|september|october|november|december|iv|v|vi|vii|viii|ix)\b|\d+/g) || [];

  return matches
    .map(word => numberWords[word] || parseInt(word))
    .sort((a, b) => a - b);
}

function getTierTranslations() {
  return {
    'Tier 1': ['Tier 1 ES', 'Tier 1 FR', 'Tier 1 DE', 'Tier 1 IT', 'Tier 1 TR', 'Tier 1 KO', 'Tier 1 JA', 'Tier 1 PT'],
    'Tier 2': ['Tier 2 ES', 'Tier 2 FR', 'Tier 2 DE', 'Tier 2 IT', 'Tier 2 TR', 'Tier 2 KO', 'Tier 2 JA', 'Tier 2 PT'],
    'Tier 3': ['Tier 3 ES', 'Tier 3 FR', 'Tier 3 DE', 'Tier 3 IT', 'Tier 3 TR', 'Tier 3 KO', 'Tier 3 JA', 'Tier 3 PT'],
    'Tier 4': ['Tier 4 ES', 'Tier 4 FR', 'Tier 4 DE', 'Tier 4 IT', 'Tier 4 TR', 'Tier 4 KO', 'Tier 4 JA', 'Tier 4 PT'],
    'Tier 5': ['Tier 5 ES', 'Tier 5 FR', 'Tier 5 DE', 'Tier 5 IT', 'Tier 5 TR', 'Tier 5 KO', 'Tier 5 JA', 'Tier 5 PT']
  };
}

function getEventNameTranslations() {
  return {
    'Event Alpha': ['Event Alpha ES', 'Event Alpha FR', 'Event Alpha DE', 'Event Alpha IT', 'Event Alpha TR', 'Event Alpha KO', 'Event Alpha JA', 'Event Alpha PT'],
    'Celebratory Event Alpha': ['Celebratory Event Alpha ES', 'Celebratory Event Alpha FR', 'Celebratory Event Alpha DE', 'Celebratory Event Alpha IT', 'Celebratory Event Alpha TR', 'Celebratory Event Alpha KO', 'Celebratory Event Alpha JA', 'Celebratory Event Alpha PT'],
    'Blitz: Event Alpha': ['Blitz Event Alpha ES', 'Blitz Event Alpha FR', 'Blitz Event Alpha DE', 'Blitz Event Alpha IT', 'Blitz Event Alpha TR', 'Blitz Event Alpha KO', 'Blitz Event Alpha JA', 'Blitz Event Alpha PT'],
    'Event Beta: Attack': ['Event Beta Attack ES', 'Event Beta Attack FR', 'Event Beta Attack DE', 'Event Beta Attack IT', 'Event Beta Attack TR', 'Event Beta Attack KO', 'Event Beta Attack JA', 'Event Beta Attack PT'],
    'Seasonal Challenge': ['Seasonal Challenge ES', 'Seasonal Challenge FR', 'Seasonal Challenge DE', 'Seasonal Challenge IT', 'Seasonal Challenge TR', 'Seasonal Challenge KO', 'Seasonal Challenge JA', 'Seasonal Challenge PT'],
    'Celebratory Seasonal Challenge': ['Celebratory Seasonal Challenge ES', 'Celebratory Seasonal Challenge FR', 'Celebratory Seasonal Challenge DE', 'Celebratory Seasonal Challenge IT', 'Celebratory Seasonal Challenge TR', 'Celebratory Seasonal Challenge KO', 'Celebratory Seasonal Challenge JA', 'Celebratory Seasonal Challenge PT'],
    'Milestone Challenge': ['Milestone Challenge ES', 'Milestone Challenge FR', 'Milestone Challenge DE', 'Milestone Challenge IT', 'Milestone Challenge TR', 'Milestone Challenge KO', 'Milestone Challenge JA', 'Milestone Challenge PT'],
    'Resource Rush': ['Resource Rush ES', 'Resource Rush FR', 'Resource Rush DE', 'Resource Rush IT', 'Resource Rush TR', 'Resource Rush KO', 'Resource Rush JA', 'Resource Rush PT'],
    'Monster Hunt': ['Monster Hunt ES', 'Monster Hunt FR', 'Monster Hunt DE', 'Monster Hunt IT', 'Monster Hunt TR', 'Monster Hunt KO', 'Monster Hunt JA', 'Monster Hunt PT']
  };
}

//KO and JA have numbers instead of month names
function getMonthTranslations() {
  return {
    'January': ['enero', 'janvier', 'Januar', 'gennaio', 'Ocak', '', '', 'janeiro'],
    'February': ['febrero', 'février', 'Februar', 'febbraio', 'Şubat', '', '', 'fevereiro'],
    'March': ['marzo', 'mars', 'März', 'marzo', 'Mart', '', '', 'março'],
    'April': ['abril', 'avril', 'April', 'aprile', 'Nisan', '', '', 'abril'],
    'May': ['mayo', 'mai', 'Mai', 'maggio', 'Mayıs', '', '', 'maio'],
    'June': ['junio', 'juin', 'Juni', 'giugno', 'Haziran', '', '', 'junho'],
    'July': ['julio', 'juillet', 'Juli', 'luglio', 'Temmuz', '', '', 'julho'],
    'August': ['agosto', 'août', 'August', 'agosto', 'Ağustos', '', '', 'agosto'],
    'September': ['septiembre', 'septembre', 'September', 'settembre', 'Eylül', '', '', 'setembro'],
    'October': ['octubre', 'octobre', 'Oktober', 'ottobre', 'Ekim', '', '', 'outubro'],
    'November': ['noviembre', 'novembre', 'November', 'novembre', 'Kasım', '', '', 'novembro'],
    'December': ['diciembre', 'décembre', 'Dezember', 'dicembre', 'Aralık', '', '', 'dezembro']
  };
}

function escapeRegExp(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function arraysEqual(arr1, arr2) {
  if (!arr1 || !arr2) return false;
  if (!Array.isArray(arr1) || !Array.isArray(arr2)) return false;
  if (arr1.length !== arr2.length) return false;

  for (let i = 0; i < arr1.length; i++) {
    if (arr1[i] !== arr2[i]) return false;
  }

  return true;
}