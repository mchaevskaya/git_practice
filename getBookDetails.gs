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

function extractPublisher($) {
  try {
    const pub = $('[itemprop="publisher"]');
    return (pub.attr('content') || pub.find('a').text() || pub.text()).trim();
  } catch (e) { return null; }
}

function extractSeries($) {
  try {
    const ser = $('[itemprop="series"]');
    return (ser.find('a').text() || ser.text()).trim();
  } catch (e) { return null; }
}

function extractPages($) {
  try {
    let pagesText = $('[itemprop="pageCount"]').attr('content') || 
                    $('meta[property="pageCount"]').attr('content') ||
                    $('[itemprop="numberOfPages"]').attr('content') ||
                    $('[itemprop="pages"]').attr('content') ||
                    $('span[itemprop="pages"]').text().trim();
    
    if (pagesText) {
      const pages = parseInt(pagesText, 10);
      if (!isNaN(pages) && pages > 0) return pages;
    }
  } catch (e) {
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


