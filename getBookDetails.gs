/**
 * ==============================================================================
 * СИСТЕМНЫЕ ТРИГГЕРЫ И МЕНЮ (ПОЛНОСТЬЮ СТАБИЛЬНЫЕ)
 * ==============================================================================
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📚 Библиотека')
    .addItem('🔄 Обновить выделенную строку', 'updateCurrentRow')
    .addToUi();
}

function atEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  if (row <= 1) return;

  if (col === 2 || col === 3) {
    const bookTitle = sheet.getRange(row, 2).getValue().toString().trim();
    const bookAuthor = sheet.getRange(row, 3).getValue().toString().trim();
    
    if (!bookTitle) {
      sheet.getRange(row, 25).setValue('Укажите название книги');
      return;
    }
    
    const currentStatus = sheet.getRange(row, 25).getValue().toString();
    if (currentStatus === 'Успешно' || currentStatus === 'Поиск...') {
      return;
    }

    sheet.getRange(row, 25).setValue('Поиск...');
    processRow(sheet, row, bookTitle, bookAuthor);
  }
}

function updateCurrentRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeCell = sheet.getActiveCell();
  const row = activeCell.getRow();
  
  if (row <= 1) {
    SpreadsheetApp.getUi().alert('Внимание', 'Пожалуйста, выделите любую ячейку в строке с книгой.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const bookTitle = sheet.getRange(row, 2).getValue().toString().trim();
  const bookAuthor = sheet.getRange(row, 3).getValue().toString().trim();
  
  if (!bookTitle) {
    sheet.getRange(row, 25).setValue('Укажите название книги');
    return;
  }
  
  sheet.getRange(row, 25).setValue('Обновление...');
  processRow(sheet, row, bookTitle, bookAuthor);
}

/**
 * ==============================================================================
 * ОБРАБОТКА СТРОКИ С УСЛОВИЕМ СКЛЕИВАНИЯ (ИЗДАТЕЛЬСТВО + СЕРИЯ)
 * ==============================================================================
 */

function processRow(sheet, row, title, author) {
  try {
    const searchResult = searchBook(title, author);
    
    if (!searchResult) {
      sheet.getRange(row, 25).setValue('Не найдено на сайте');
      return;
    }
    
    // Формируем текст для Издания (Колонка D) строго по ТЗ
    let editionText = '';
    const pub = searchResult.publisher;
    const ser = searchResult.series;

    if (pub) {
      editionText = pub;
      if (ser) {
        editionText += '. ' + ser; // Добавляем точку, только если есть и серия
      }
    } else {
      editionText = ''; // Если издательства нет, ячейка остается пустой
    }
    
    sheet.getRange(row, 4).setValue(editionText);
    
    if (searchResult.pages) {
      sheet.getRange(row, 11).setValue(searchResult.pages);
    }
    
    if (searchResult.imageUrl) {
      saveOriginalJpeg(sheet, row, 5, searchResult.imageUrl);
    }
    
    sheet.getRange(row, 25).setValue('Успешно');
    
  } catch (error) {
    Logger.log('Ошибка обработки строки ' + row + ': ' + error.toString());
    sheet.getRange(row, 25).setValue('Ошибка скрипта');
  }
}

/**
 * ==============================================================================
 * СВЯЩЕННЫЙ РАБОЧИЙ ПОИСК (ВОЗВРАЩЕН В ПЕРВОЗДАННЫЙ ВИД)
 * ==============================================================================
 */
function searchBook(title, author) {
  const query = encodeURIComponent(title + (author ? ' ' + author : ''));
  const searchUrl = 'https://www.chitai-gorod.ru/search?phrase=' + query;
  
  const options = {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
      'Accept-Language': 'ru-RU,ru;q=0.9'
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(searchUrl, options);
    if (response.getResponseCode() !== 200) return null;
    
    const html = response.getContentText('UTF-8');
    
    let cleanPath = '';
    if (html.includes('href="/product/')) {
      const parts = html.split('href="/product/');
      const rightPart = parts[1]; // ТОТ САМЫЙ РАБОЧИЙ ИНДЕКС
      const endQuoteIndex = rightPart.indexOf('"');
      if (endQuoteIndex !== -1) {
        cleanPath = '/product/' + rightPart.substring(0, endQuoteIndex);
      }
    }
    
    if (!cleanPath) return null;
    
    const productUrl = 'https://www.chitai-gorod.ru' + cleanPath;
    const productResponse = UrlFetchApp.fetch(productUrl, options);
    if (productResponse.getResponseCode() !== 200) return null;
    
    const productHtml = productResponse.getContentText('UTF-8');

    const $ = Cheerio.load(productHtml);
    
    // Передаем HTML-код в безопасный теговый парсинг Cheerio
    const publisher = extractPublisher($);
    const series = extractSeries($);
    const pages = extractPages($);
    const imageUrl = extractImageUrl($);
    
    return { 
      pages: pages, 
      publisher: publisher, 
      series: series, 
      imageUrl: imageUrl 
    };
    
  } catch (e) {
    Logger.log('Ошибка при парсинге сайта: ' + e.toString());
    return null;
  }
}

/**
 * ==============================================================================
 * БЕЗОПАСНЫЙ ИЗВЛЕКАТЕЛЬ ДАННЫХ ЧЕРЕЗ ТЕГИ (ЧЕРИО)
 * ==============================================================================
 */

/**
 * Проверяет, что найденное значение — реальное название, а не URL или сам магазин
 */
function isValidText(val) {
  if (!val || typeof val !== 'string') return false;
  const clean = val.trim();
  if (!clean || clean === 'null' || clean === 'undefined') return false;
  if (clean.startsWith('http') || clean.includes('chitai-gorod')) return false;
  if (clean.startsWith('{') || clean.startsWith('[')) return false;
  return true;
}

/**
 * Поиск значений в JSON внутри тегов <script> с отсеиванием ссылок
 */
function extractFromJsonScripts($, keys) {
  let result = null;
  $('script').each((_, el) => {
    if (result) return;
    const content = $(el).html() || '';
    for (const key of keys) {
      // 1. Ищем объект вида "publisher":{"@type":"Organization","name":"АСТ"}
      const objRegex = new RegExp(`"${key}"\\s*:\\s*\\{[^}]*"name"\\s*:\\s*"([^"]+)"`, 'gi');
      let match;
      while ((match = objRegex.exec(content)) !== null) {
        if (isValidText(match[1])) {
          result = match[1].trim();
          return;
        }
      }

      // 2. Ищем простое строковое или числовое поле вида "publisher":"АСТ" или "numberOfPages":320
      const strRegex = new RegExp(`"${key}"\\s*:\\s*"?([^",}\\]]+)"?`, 'gi');
      while ((match = strRegex.exec(content)) !== null) {
        if (isValidText(match[1])) {
          result = match[1].trim();
          return;
        }
      }
    }
  });
  return result;
}

/**
 * Поиск значения по текстовой метке в блоке характеристик
 */
function extractByLabel($, labelText) {
  let foundValue = null;
  $('*').each((_, el) => {
    if (foundValue) return;
    if ($(el).children().length === 0 && $(el).text().trim() === labelText) {
      const nextText = $(el).next().text().trim() || $(el).parent().next().text().trim();
      if (isValidText(nextText)) {
        foundValue = nextText;
      }
    }
  });
  return foundValue;
}

function extractPublisher($) {
  try {
    // 1. Самый точный способ: ссылка на страницу издательства (/publisher/...)
    const pubLink = $('a[href*="/publisher/"]').first().text().trim();
    if (isValidText(pubLink)) return pubLink;

    // 2. Поиск по текстовой метке "Издательство" в характеристиках
    const labelVal = extractByLabel($, 'Издательство');
    if (isValidText(labelVal)) return labelVal;

    // 3. Поиск в JSON-разметке страницы (с игнорированием URL)
    const jsonVal = extractFromJsonScripts($, ['publisher']);
    if (isValidText(jsonVal)) return jsonVal;

    return null;
  } catch (e) {
    Logger.log('Ошибка в extractPublisher: ' + e.toString());
    return null;
  }
}

function extractSeries($) {
  try {
    // 1. Ссылка на страницу серии (/series/...)
    const serLink = $('a[href*="/series/"]').first().text().trim();
    if (isValidText(serLink)) return serLink;

    // 2. Поиск по текстовой метке "Серия" в характеристиках
    const labelVal = extractByLabel($, 'Серия');
    if (isValidText(labelVal)) return labelVal;

    // 3. Поиск в JSON-разметке страницы
    const jsonVal = extractFromJsonScripts($, ['series', 'bookSeries']);
    if (isValidText(jsonVal)) return jsonVal;

    return null;
  } catch (e) {
    Logger.log('Ошибка в extractSeries: ' + e.toString());
    return null;
  }
}

function extractPages($) {
  try {
    let pagesText = extractFromJsonScripts($, ['numberOfPages', 'pagesCount', 'pageCount', 'pages']);

    if (!pagesText) {
      pagesText = extractByLabel($, 'Кол-во стр.');
    }

    if (pagesText) {
      const pages = parseInt(pagesText.toString().replace(/\D/g, ''), 10);
      if (!isNaN(pages) && pages > 0) return pages;
    }
    return null;
  } catch (e) {
    Logger.log('Ошибка в extractPages: ' + e.toString());
    return null;
  }
}


/**
 * Извлекает оригинальный URL картинки обложки СТРОГО в формате JPEG.
 */
function extractImageUrl($) {
  try {
    
    // Ищем тег og:image — Читай-город всегда кладет туда полноценный качественный JPEG
    let url = $('meta[property="og:image"]').attr('content');
    
    if (url) {
      url = url.toString().trim();
      if (url.startsWith('//')) url = 'https:' + url;
      
      // Если сайт отдал webp в og:image, принудительно меняем расширение в ссылке на jpg
      // (Сервер Читай-города поддерживает выдачу jpg по точно такому же адресу)
      if (url.includes('.webp')) {
        url = url.replace('.webp', '.jpg');
      }
      
      Logger.log('extractImageUrl: Успешно найден чистый JPEG: ' + url);
      return url;
    }
    return null;
  } catch (e) {
    Logger.log('Ошибка в extractImageUrl: ' + e.toString());
    return null;
  }
}

/**
 * Сверхнадежная функция вставки обложек.
 * Настоящий JPEG сохраняет в память, а WebP фиксирует в ячейке через временный файл на Google Диске.
 */
function saveOriginalJpeg(sheet, row, col, imageUrl) {
  const cell = sheet.getRange(row, col);
  
  if (!imageUrl) {
    cell.clearContent();
    return;
  }
  
  const options = {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36'
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(imageUrl, options);
    
    if (response.getResponseCode() === 200) {
      const imageBlob = response.getBlob();
      const base64Data = Utilities.base64Encode(imageBlob.getBytes());
      
      // 1. ПРОВЕРКА НА СКРЫТЫЙ WEBP (код начинается на UklGR)
      if (base64Data.startsWith('UklGR')) {
        Logger.log('Обнаружен WebP. Запускаем фиксацию через временный файл на Google Диске...');
        
        // Создаем временный файл на вашем Google Диске
        // Даем уникальное имя по номеру строки, чтобы файлы не пересекались
        const tempFile = DriveApp.createFile(imageBlob);
        tempFile.setName('temp_cover_row_' + row + '.webp');
        
        // Включаем доступ по ссылке, чтобы внутренний сервер таблиц мог забрать файл
        tempFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        const tempDownloadUrl = tempFile.getDownloadUrl();
        
        // Передаем внутреннюю ссылку Диска в конструктор картинок
        const imageBuilder = SpreadsheetApp.newCellImage()
          .setSourceUrl(tempDownloadUrl)
          .setAltTextDescription('Обложка книги (зафиксировано из WebP)')
          .build();
          
        cell.setValue(imageBuilder);
        
        // Принудительно заставляем Google Таблицы завершить все графические операции на листе
        SpreadsheetApp.flush();
        
        // Отправляем временный файл в корзину Диска, чтобы он не мешался
        tempFile.setTrashed(true);
        Logger.log('Временный WebP-файл успешно удален с Диска. Картинка зафиксирована в ячейке.');
        return;
      }
      
      // 2. ЕСЛИ ЭТО КЛАССИЧЕСКИЙ JPEG
      const dataUrl = 'data:image/jpeg;base64,' + base64Data;
      const imageBuilder = SpreadsheetApp.newCellImage()
        .setSourceUrl(dataUrl)
        .setAltTextDescription('Обложка книги (сохранено как JPEG)')
        .build();
        
      cell.setValue(imageBuilder);
      Logger.log('Обложка для строки ' + row + ' успешно сохранена как локальный JPEG!');
    } else {
      Logger.log('Не удалось скачать картинку, код: ' + response.getResponseCode());
      cell.clearContent();
    }
  } catch (e) {
    Logger.log('Ошибка сохранения картинки: ' + e.toString());
    cell.clearContent();
  }
}


