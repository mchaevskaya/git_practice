/**
 * Триггер, срабатывающий при редактировании таблицы
 */
function onEdit(e) {
  try {
    // Проверяем, есть ли параметр e
    if (!e) {
      // Если функция запущена вручную, получаем активный лист
      const sheet = SpreadsheetApp.getActiveSheet();
      const row = SpreadsheetApp.getActiveRange().getRow();
      return processRow(sheet, row);
    }

    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const row = range.getRow();
    const col = range.getColumn();

    // Обрабатываем только изменения в колонке C (автор), начиная со 2‑й строки
    if (col !== 3 || row < 2) return;

    const title = sheet.getRange(row, 2).getValue(); // колонка B — название
    const author = sheet.getRange(row, 3).getValue(); // колонка C — автор

    // Если название или автор не заполнены — выходим
    if (!title || !author) return;

    // Проверяем, что в колонке E (обложка) ещё нет данных
    const coverCell = sheet.getRange(row, 5); // колонка E
    if (coverCell.getValue() !== '') return;

    console.log(`Обрабатывается новая строка ${row}: "${title}" - "${author}"`);

    updateStatus(sheet, row - 1, 'Обрабатывается');

    // Поиск книги и получение URL страницы
    const bookUrl = searchBook(title, author);
    if (!bookUrl) {
      updateStatus(sheet, row - 1, 'Не найдено на сайте');
      return;
    }

    // Загрузка страницы книги
    const bookHtml = fetchBookPage(bookUrl);
    if (!bookHtml) {
      updateStatus(sheet, row - 1, 'Ошибка загрузки страницы');
      return;
    }

    // Получение и вставка изображения обложки (в колонку E)
    const imageUrl = extractImageUrl(bookHtml);
    insertImageToCell(sheet, row - 1, imageUrl);

    // Получение и запись количества страниц (в колонку K)
    const pages = extractPages(bookHtml);
    sheet.getRange(row, 11).setValue(pages); // колонка K = 11
    console.log(`Количество страниц: ${pages}`);

    updateStatus(sheet, row - 1, 'Успешно');
  } catch (error) {
    updateStatus(sheet, row - 1, `Ошибка: ${error.message}`);
    console.error(`Ошибка в строке ${row}:`, error);
  }
}

/**
 * Функция для ручного запуска скрипта
 */
function runManually() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const row = SpreadsheetApp.getActiveRange().getRow();
  onEdit({ source: sheet, range: sheet.getRange(row, 3) });
}

/**
 * Создаем меню для удобного запуска
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Обработка')
    .addItem('Запустить обработку', 'runManually')
    .addToUi();
}

/**
 * Основная функция обработки строки
 * @param {Sheet} sheet — лист для обработки
 * @param {number} row — номер строки
 */
function processRow(sheet, row) {
  try {
    const title = sheet.getRange(row, 2).getValue(); // колонка B — название
    const author = sheet.getRange(row, 3).getValue(); // колонка C — автор

    // Проверки данных
    if (!title || !author) return;

    // Поиск книги
    const bookUrl = searchBook(title, author);
    if (!bookUrl) {
      updateStatus(sheet, row - 1, 'Не найдено на сайте');
      return;
    }

    // Загрузка страницы
    const bookHtml = fetchBookPage(bookUrl);
    if (!bookHtml) {
      updateStatus(sheet, row - 1, 'Ошибка загрузки страницы');
      return;
    }

    // Извлекаем издательство и серию
    const { publisher, series } = extractPublisherAndSeries(bookHtml);
    
    // Формируем значение для колонки D
    let edition = '';
    if (publisher) {
      edition = publisher;
      if (series) {
        edition += '. ' + series;
      }
    } else if (series) {
      edition = series;
    }
    
    // Записываем результат
    sheet.getRange(row, 4).setValue(edition);
    console.log(`В колонку D записано: ${edition}`);

    // Остальные операции
    const imageUrl = extractImageUrl(bookHtml);
    insertImageToCell(sheet, row - 1, imageUrl);
    
    const pages = extractPages(bookHtml);
    sheet.getRange(row, 11).setValue(pages); // колонка K = 11
    
    updateStatus(sheet, row - 1, 'Успешно');
  } catch (error) {
    updateStatus(sheet, row - 1, `Ошибка: ${error.message}`);
    console.error(`Ошибка в строке ${row}:`, error);
  }
}

/**
 * Извлекает издательство и серию из HTML страницы книги
 * @param {string} html — HTML страницы книги
 * @return {Object} — объект с полями publisher и series
 */
function extractPublisherAndSeries(html) {
  console.log('Начинаем извлечение издательства и серии');
  
  // Обновленные регулярные выражения для поиска издательства и серии
  const publisherMatch = html.match(/Издательство[^>]*>([^<]+)</i);
  const seriesMatch = html.match(/Серия[^>]*>([^<]+)</i);
  
  console.log(`Найдено издательство: ${publisherMatch ? publisherMatch[1] : 'не найдено'}`);
  console.log(`Найдено серия: ${seriesMatch ? seriesMatch[1] : 'не найдено'}`);
  
  return {
    publisher: publisherMatch ? publisherMatch[1].trim() : '',
    series: seriesMatch ? seriesMatch[1].trim() : ''
  };
}

/**
 * Поиск книги на сайте
 * @param {string} title — название книги
 * @param {string} author — автор книги
 * @return {string|null} — URL страницы книги или null
 */
function searchBook(title, author) {
  const searchTerm = encodeURIComponent(`${title} ${author}`);
  const url = `https://www.chitai-gorod.ru/search?phrase=${searchTerm}`;
  console.log(`Выполняется поиск: ${url}`);

  try {
    const response = UrlFetchApp.fetch(url, getRequestOptions());
    if (response.getResponseCode() !== 200) {
      throw new Error(`HTTP ${response.getResponseCode()}`);
    }
    const html = response.getContentText();
    const match = html.match(/(\/product\/[^"'>]*)/i);
    return match ? `https://www.chitai-gorod.ru${match[1]}` : null;
  } catch (error) {
    console.error('Ошибка поиска книги:', error);
    return null;
  }
}

/**
 * Загрузка страницы книги
 * @param {string} url — URL страницы книги
 * @return {string|null} — HTML страницы или null
 */
function fetchBookPage(url) {
  try {
    const response = UrlFetchApp.fetch(url, getRequestOptions());
    if (response.getResponseCode() !== 200) {
      throw new Error('Ошибка загрузки страницы книги');
    }
    return response.getContentText();
  } catch (error) {
    console.error('Ошибка загрузки страницы:', error);
    return null;
  }
}

/**
 * Извлекает URL изображения обложки из HTML
 * @param {string} html — HTML страницы книги
 * @return {string|null} — URL изображения или null
 */
function extractImageUrl(html) {
  const selectors = [
    /data-src=["']([^"']*\.(?:jpg|jpeg|png|webp))/i,
    /<img[^>]*src=["']([^"']*\.(?:jpg|jpeg|png|webp))/i,
    /<meta[^>]*property=["']og:image["'][^>]*content=["']([^"']*)/i,
    /"image"\s*:\s*["']([^"']*\.(?:jpg|jpeg|png|webp))/i
  ];
  for (const regex of selectors) {
    const match = html.match(regex);
    if (match) {
      let url = match[1];
      if (!url.startsWith('http')) {
        url = url.startsWith('/') ? url.substring(1) : url;
        return `https://www.chitai-gorod.ru/${url}`;
      }
      return url;
    }
  }
  return null;
}

/**
 * Извлекает количество страниц из HTML
 * @param {string} html — HTML страницы книги
 * @return {string} — количество страниц или сообщение об отсутствии данных
 */
function extractPages(html) {
  const match = html.match(/(\d+)(?=\s*(?:стр\.?|страниц?))/i);
  if (match && match[1]) return match[1].trim();
  return 'Объём не указан';
}

/**
 * Вставляет изображение в ячейку (скачивает, конвертирует в JPEG и сжимает)
 * @param {Sheet} sheet — лист, куда вставляем
 * @param {number} row — номер строки (0‑индексация)
 * @param {string} imageUrl — URL изображения
 */
function insertImageToCell(sheet, row, imageUrl) {
  if (!imageUrl) {
    sheet.getRange(row + 1, 5).setValue('Изображение не найдено'); // колонка E = 5
    console.log('Обложка не найдена');
    return;
  }

  try {
    // Скачиваем изображение
    const imageResponse = UrlFetchApp.fetch(imageUrl, {
      muteHttpExceptions: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
      }
    });

    if (imageResponse.getResponseCode() !== 200) {
      sheet.getRange(row + 1, 5).setValue('Ошибка загрузки изображения'); // колонка E = 5
      console.log('Ошибка HTTP при загрузке изображения:', imageResponse.getResponseCode());
      return;
    }

    // Получаем бинарные данные изображения
    const originalBlob = imageResponse.getBlob();

    // Конвертируем в JPEG с запасным планом
    const jpegBlob = convertToJpegWithFallback(originalBlob);
    if (!jpegBlob) {
      sheet.getRange(row + 1, 5).setValue('Ошибка конвертации в JPEG'); // колонка E = 5
      console.log('Не удалось конвертировать изображение в JPEG');
      return;
    }

    // Сохраняем сжатое изображение во временный файл в Google Drive с публичным доступом
    const tempImageFile = DriveApp.createFile(jpegBlob.setName('temp_image.jpg'));
    tempImageFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // Формируем правильный URL для формулы IMAGE()
    const fileId = tempImageFile.getId();
    const publicUrl = `https://drive.google.com/uc?id=${fileId}&export=download`;

    // Вставляем изображение в ячейку с помощью формулы IMAGE() (колонка E)
    const cell = sheet.getRange(row + 1, 5); // колонка E, строка row+1
    cell.setFormula(`=IMAGE("${publicUrl}", "1")`); // "1" — масштабирование по размеру ячейки

    // Удаляем временный файл через задержку (даём время таблице загрузить изображение)
    Utilities.sleep(2000); // пауза 2 секунды
    tempImageFile.setTrashed(true);

    console.log(`Обложка успешно вставлена в ячейку: ${imageUrl}`);
  } catch (error) {
    sheet.getRange(row + 1, 5).setValue('Ошибка вставки изображения'); // колонка E = 5
    console.error('Ошибка при вставке изображения:', error);
  }
}

/**
 * Конвертирует изображение в JPEG с несколькими уровнями запасного плана
 * @param {Blob} blob — исходное изображение любого формата
 * @return {Blob|null} — JPEG‑изображение или null при ошибке
 */
function convertToJpegWithFallback(blob) {
  try {
    // Метод 1: Прямое преобразование через Drive API
    const directJpeg = tryDirectConversion(blob);
    if (directJpeg) return directJpeg;

    // Метод 2: Через миниатюру с уменьшением размера
    const thumbnailJpeg = tryThumbnailConversion(blob);
    if (thumbnailJpeg) return thumbnailJpeg;

    // Метод 3: Экспорт с явным указанием параметров
    const exportJpeg = tryExportConversion(blob);
    if (exportJpeg) return exportJpeg;

    // Если все методы не сработали
    console.error('Все методы конвертации провалились');
    return null;
  } catch (error) {
    console.error('Критическая ошибка в конвертации:', error);
    return null;
  }
}

/**
 * Попытка прямого преобразования в JPEG
 * @param {Blob} blob — исходное изображение
 * @return {Blob|null}
 */
function tryDirectConversion(blob) {
  try {
    const tempFile = DriveApp.createFile(blob);
    const jpegBlob = tempFile.getAs(MimeType.JPEG)
      .setContentType(MimeType.JPEG)
      .setName('converted.jpg');
    tempFile.setTrashed(true);
    return jpegBlob;
  } catch (error) {
    console.warn('Прямое преобразование не удалось:', error);
    return null;
  }
}

/**
 * Попытка конвертации через миниатюру (автоматически уменьшает размер)
 * @param {Blob} blob — исходное изображение
 * @return {Blob|null}
 */
function tryThumbnailConversion(blob) {
  try {
    const tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();
    const thumbnailUrl = `https://drive.google.com/thumbnail?id=${fileId}&sz=w400-h400`;
    const response = UrlFetchApp.fetch(thumbnailUrl, { muteHttpExceptions: true });
    if (response.getResponseCode() === 200) {
      const thumbnailBlob = response.getBlob()
        .setContentType(MimeType.JPEG)
        .setName('thumbnail.jpg');
      tempFile.setTrashed(true);
      return thumbnailBlob;
    }
    tempFile.setTrashed(true);
    return null;
  } catch (error) {
    console.warn('Конвертация через миниатюру не удалась:', error);
    return null;
  }
}

/**
 * Попытка экспорта с явным указанием параметров
 * @param {Blob} blob — исходное изображение
 * @return {Blob|null}
 */
function tryExportConversion(blob) {
  try {
    const tempFile = DriveApp.createFile(blob);
    const jpegBlob = tempFile.getAs('image/jpeg')
      .setContentType('image/jpeg')
      .setName('exported.jpg');
    tempFile.setTrashed(true);
    return jpegBlob;
  } catch (error) {
    console.warn('Экспорт с явным указанием не удался:', error);
    return null;
  }
}

/**
 * Обновляет статус в колонке F (6)
 * @param {Sheet} sheet — лист, куда вставляем
 * @param {number} row — номер строки (0‑индексация)
 * @param {string} status — текст статуса
 */
function updateStatus(sheet, row, status) {
  sheet.getRange(row + 1, 6).setValue(status); // колонка F = 6
}

/**
 * Возвращает настройки для HTTP‑запросов
 * @return {Object} — параметры запроса
 */
function getRequestOptions() {
  return {
    muteHttpExceptions: true,
    headers: {
      'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
      'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
      'Accept-Language': 'ru-RU,ru;q=0.9'
    }
  };
}
