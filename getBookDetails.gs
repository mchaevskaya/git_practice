/**
 * Функция-триггер: срабатывает автоматически при любом изменении в таблице.
 */
function atEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // Исключаем шапку (строка 1)
  if (row <= 1) return;

  // Проверяем, что изменения произошли в колонке B (2 - Название) ИЛИ в колонке C (3 - Автор)
  if (col === 2 || col === 3) {
    const bookTitle = sheet.getRange(row, 2).getValue().toString().trim();
    const bookAuthor = sheet.getRange(row, 3).getValue().toString().trim();
    
    // Если название пука пустое, искать нет смысла (даже если автора заполнили)
    if (!bookTitle) {
      sheet.getRange(row, 12).setValue('Укажите название книги');
      return;
    }
    
    // Если скрипт уже успешно отработал по этой строке ранее, не мучаем сайт повторно.
    // Если хотите принудительно переискать — просто сотрите статус "Успешно" в колонке L (12).
    const currentStatus = sheet.getRange(row, 12).getValue().toString();
    if (currentStatus === 'Успешно' || currentStatus === 'Поиск...') {
      return;
    }

    // Запускаем процесс парсинга
    sheet.getRange(row, 12).setValue('Поиск...');
    processRow(sheet, row, bookTitle, bookAuthor);
  }
}

/**
 * 1. Исправленная логика склеивания Издательства и Серии
 */
function processRow(sheet, row, title, author) {
  try {
    const searchResult = searchBook(title, author); // Вызывает ваш рабочий поиск
    
    if (!searchResult) {
      sheet.getRange(row, 12).setValue('Не найдено на сайте');
      return;
    }
    
    // Формируем текст для колонки D (Издание) строго по условиям
    let editionText = '';
    const pub = searchResult.publisher;
    const ser = searchResult.series;

    if (pub) {
      editionText = pub;
      // Добавляем точку, ТОЛЬКО если нашлись И издательство, И серия
      if (ser) {
        editionText += '. ' + ser;
      }
    } else {
      // Если Издательство не найдено, ячейка остается пустой (даже если серия есть)
      editionText = ''; 
    }
    
    // Записываем результат в колонку D (4 - Издание)
    sheet.getRange(row, 4).setValue(editionText);
    
    // Записываем Количество страниц в колонку K (11)
    if (searchResult.pages) {
      sheet.getRange(row, 11).setValue(searchResult.pages);
    }
    
    // Вставляем обложку напрямую в ячейку E (5)
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
 * Ищет книгу на Читай-Городе через GET-запрос страницы поиска
 */
function searchBook(title, author) {
  // 1. Формируем поисковый URL (строго кодируем русские буквы)
  const searchPhrase = (title + (author ? ' ' + author : '')).trim();
  const encodedPhrase = encodeURIComponent(searchPhrase);
  const searchUrl = 'https://www.chitai-gorod.ru/search?phrase=' + encodedPhrase;
  
  Logger.log('1. Создан URL поиска: ' + searchUrl);
  
  const options = {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
      'Accept-Language': 'ru-RU,ru;q=0.9',
      'Cache-Control': 'no-cache',
      'Pragma': 'no-cache'
    }
  };
  
  try {
    const response = UrlFetchApp.fetch(searchUrl, options);
    const responseCode = response.getResponseCode();
    Logger.log('2. Ответ сервера поиска. Код: ' + responseCode);
    
    if (responseCode !== 200) {
      Logger.log('Ошибка: Сайт вернул код ' + responseCode);
      return null;
    }
    
    const html = response.getContentText('UTF-8');
    
    // Проверяем, не подсунул ли Cloudflare страницу проверки на робота
    if (html.includes('error-code') || html.includes('cloudflare') || html.length < 5000) {
      Logger.log('Защита Cloudflare заблокировала чтение страницы поиска (код слишком короткий).');
      return null;
    }
    
    // 2. Ищем ссылку на книгу. Используем метод split вместо хрупкого regex, чтобы исключить баг склеивания строк
    let cleanPath = '';
    if (html.includes('href="/product/')) {
      const parts = html.split('href="/product/');
      // Берем первую попавшуюся ссылку после разделителя
      const rightPart = parts[1];
      const endQuoteIndex = rightPart.indexOf('"');
      if (endQuoteIndex !== -1) {
        cleanPath = '/product/' + rightPart.substring(0, endQuoteIndex);
      }
    }
    
    // Если через split не нашлось, пробуем стандартный Regex, но БЕЗ перевода массива в строку
    if (!cleanPath) {
      const regex = /href="(\/product\/[^"]+)"/i;
      const match = regex.exec(html);
      if (match && match[1]) {
        cleanPath = String(match[1]).trim();
      }
    }
    
    // Жесткая проверка: если путь не найден или в него каким-то чудом попал исходный текст — останавливаем скрипт
    if (!cleanPath || cleanPath.includes('search') || cleanPath.length > 200) {
      Logger.log('3. Ссылка на книгу на странице результатов не найдена.');
      return null;
    }
    
    // Собираем финальный URL карточки товара
    const productUrl = 'https://www.chitai-gorod.ru' + cleanPath;
    Logger.log('4. Успешно собрана ссылка на карточку книги: ' + productUrl);
    
    // 3. Загружаем страницу самой книги
    const productResponse = UrlFetchApp.fetch(productUrl, options);
    if (productResponse.getResponseCode() !== 200) {
      Logger.log('Ошибка загрузки карточки товара. Код: ' + productResponse.getResponseCode());
      return null;
    }
    
    const productHtml = productResponse.getContentText('UTF-8');
    Logger.log('5. Страница книги успешно загружена. Начинаем парсинг параметров...');
    
    // ВЫЗЫВАЕМ ОБНОВЛЕННЫЕ ФУНКЦИИ ПАРСИНГА
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
    Logger.log('Критический сбой в searchBook: ' + e.toString());
    return null;
  }
}

/**
 * Извлекает количество страниц (комбинированный метод: теги + текст)
 */
function extractPages(html) {
  try {
    // Способ 1: Ищем через библиотеку Cheerio по тегам и микроразметке
    const $ = Cheerio.load(html);
    let pagesText = $('[itemprop="pageCount"]').attr('content') || 
                    $('meta[property="pageCount"]').attr('content') ||
                    $('[itemprop="numberOfPages"]').attr('content');
    
    if (pagesText) {
      const pages = parseInt(pagesText, 10);
      if (!isNaN(pages) && pages > 0) {
        Logger.log('Страницы найдены через Cheerio по тегам: ' + pages);
        return pages;
      }
    }

    // Способ 2: Наш старый проверенный метод поиска по визуальному маркеру
    const lowerHtml = html.toLowerCase();
    const targetPhrases = ['кол-во стр.', 'количество страниц', 'страниц:'];
    
    for (let i = 0; i < targetPhrases.length; i++) {
      const phrase = targetPhrases[i];
      
      if (lowerHtml.includes(phrase)) {
        const parts = lowerHtml.split(phrase);
        const textAfterPhrase = parts[1].substring(0, 150);
        const numberMatch = textAfterPhrase.match(/\d+/);
        
        if (numberMatch) {
          const pagesFallback = parseInt(numberMatch[0], 10);
          if (!isNaN(pagesFallback) && pagesFallback > 0 && pagesFallback < 10000) {
            Logger.log('Страницы найдены через текстовый маркер "' + phrase + '": ' + pagesFallback);
            return pagesFallback;
          }
        }
      }
    }
    
    Logger.log('Количество страниц не удалось найти ни одним способом.');
    return null;

  } catch (e) {
    Logger.log('Ошибка при парсинге страниц: ' + e.toString());
    return null;
  }
}

/**
 * Извлекает название издательства с помощью поиска по HTML-тегам
 */
function extractPublisher(html) {
  try {
    // Загружаем HTML-код страницы в парсер тегов
    const $ = Cheerio.load(html);
    
    // Ищем тег, который отвечает за разметку издательства (itemprop="publisher")
    const publisherSpan = $('[itemprop="publisher"]');
    
    // 1. Сначала проверяем, записано ли издательство в атрибут content (как видно в вашем логе: content="аст")
    let publisher = publisherSpan.attr('content');
    
    // 2. Если атрибута нет, пробуем забрать чистый текст из ссылки внутри этого тега
    if (!publisher) {
      publisher = publisherSpan.find('a').text() || publisherSpan.text();
    }
    
    if (publisher) {
      const cleanPub = publisher.toString().trim();
      Logger.log('Успешно нашли Издательство по тегам: "' + cleanPub + '"');
      return cleanPub;
    }
    
    Logger.log('Тег издательства не найден на странице книги.');
    return null;
  } catch (e) {
    Logger.log('Ошибка при поиске тега издательства: ' + e.toString());
    return null;
  }
}

/**
 * Извлекает название серии с помощью поиска по HTML-тегам
 */
function extractSeries(html) {
  try {
    const $ = Cheerio.load(html);
    
    // Ищем тег, отвечающий за серию книги (itemprop="series")
    const seriesSpan = $('[itemprop="series"]');
    
    // Забираем текст ссылки <a> внутри этого тега (например: "Эксклюзивная классика")
    let series = seriesSpan.find('a').text() || seriesSpan.text();
    
    if (series) {
      const cleanSer = series.toString().trim();
      Logger.log('Успешно нашли Серию по тегам: "' + cleanSer + '"');
      return cleanSer;
    }
    
    Logger.log('Тег серии не найден на странице книги.');
    return null;
  } catch (e) {
    Logger.log('Ошибка при поиске тега серии: ' + e.toString());
    return null;
  }
}

/**
 * Извлекает чистый URL картинки обложки
 */
function extractImageUrl(html) {
  // Читай-город отдает картинку в теге og:image для соцсетей — это идеальный источник
  const match = html.match(/<meta[^>]*property="og:image"[^>]*content="([^"]+)"/i) || 
                html.match(/"image"\s*:\s*"([^"]+)"/i);
  if (match) {
    let url = match[1].toString().trim();
    if (url.startsWith('//')) url = 'https:' + url;
    return url;
  }
  return null;
}

/**
 * Современный и безопасный метод вставки картинки прямо внутрь ячейки.
 * Больше не требует папки на Google Диске и не ломается при удалении файлов.
 */
function insertImageToCellDirect(sheet, row, col, imageUrl) {
  const cell = sheet.getRange(row, col);
  
  // Создаем объект изображения Google Таблиц на основе URL сайта
  const imageBuilder = SpreadsheetApp.newCellImage()
    .setSourceUrl(imageUrl)
    .setAltTextDescription('Обложка книги')
    .build();
    
  cell.setValue(imageBuilder);
}
