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
      sheet.getRange(row, 12).setValue('Укажите название книги');
      return;
    }
    
    const currentStatus = sheet.getRange(row, 12).getValue().toString();
    if (currentStatus === 'Успешно' || currentStatus === 'Поиск...') {
      return;
    }

    sheet.getRange(row, 12).setValue('Поиск...');
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
    sheet.getRange(row, 12).setValue('Укажите название книги');
    return;
  }
  
  sheet.getRange(row, 12).setValue('Обновление...');
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
      sheet.getRange(row, 12).setValue('Не найдено на сайте');
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
      insertImageToCellDirect(sheet, row, 5, searchResult.imageUrl);
    }
    
    sheet.getRange(row, 12).setValue('Успешно');
    
  } catch (error) {
    Logger.log('Ошибка обработки строки ' + row + ': ' + error.toString());
    sheet.getRange(row, 12).setValue('Ошибка скрипта');
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
    
    // Передаем HTML-код в безопасный теговый парсинг Cheerio
    const publisher = extractPublisher(productHtml);
    const series = extractSeries(productHtml);
    const pages = extractPages(productHtml);
    const imageUrl = extractImageUrl(productHtml);
    
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

function extractPublisher(html) {
  try {
    const $ = Cheerio.load(html);
    const publisherSpan = $('[itemprop="publisher"]');
    let publisher = publisherSpan.attr('content'); 
    
    if (!publisher) {
      publisher = publisherSpan.find('a').text() || publisherSpan.text();
    }
    return publisher ? publisher.toString().trim() : null;
  } catch (e) {
    return null;
  }
}

function extractSeries(html) {
  try {
    const $ = Cheerio.load(html);
    const seriesSpan = $('[itemprop="series"]');
    const series = seriesSpan.find('a').text() || seriesSpan.text();
    return series ? series.toString().trim() : null;
  } catch (e) {
    return null;
  }
}

function extractPages(html) {
  try {
    const $ = Cheerio.load(html);
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

function extractImageUrl(html) {
  try {
    const $ = Cheerio.load(html);
    let url = $('meta[property="og:image"]').attr('content') || $('[itemprop="image"]').attr('content');
    if (url) {
      url = url.toString().trim();
      if (url.startsWith('//')) url = 'https:' + url;
      return url;
    }
    return null;
  } catch (e) {
    return null;
  }
}

function insertImageToCellDirect(sheet, row, col, imageUrl) {
  const cell = sheet.getRange(row, col);
  const imageBuilder = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageUrl)
    .setAltTextDescription('Обложка книги')
    .build();
    
  cell.setValue(imageBuilder);
}
