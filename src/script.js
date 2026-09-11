/**
 * Внедряйте в ОСНОВНОЙ JavaScript-контекст iframe ДО скриптов приложения.
 * После полной навигации внедряйте заново; во вложенные iframe — отдельно.
 *
 * Версия 2: распознанные скачивания отменяются синхронно. Затем скрипт читает
 * байты, определяет формат Office и сохраняет именно ПРОВЕРЕННЫЕ данные.
 * XLS: сигнатура CFB, поток Excel в корне контейнера и записи BOF/EOF BIFF5/8.
 * XLSX/PPTX: сигнатура ZIP, связи и типы содержимого пакета, главный XML-документ.
 * Это определение формата, а не антивирус или полная проверка документа Office.
 *
 * Лимиты: 50 МиБ на файл, 3 одновременные проверки, 30 секунд на чтение.
 * ZIP64 и зашифрованные документы Office отклоняются. Распознанные скачивания
 * через POST блокируются; скрипт не повторяет эти запросы как GET.
 * HTTP-файлы должны быть доступны через fetch с соблюдением CORS.
 * Непрозрачные ответы и ошибки чтения приводят к блокировке скачивания.
 * У iframe в sandbox с непрозрачным origin чтение HTTP-файлов часто требует CORS.
 * Для Blob-ссылок, созданных после установки скрипта, объект хранится до отзыва
 * ссылки. Начатая проверка удерживает Blob, даже если приложение сразу отзывает URL.
 *
 * Перехват остаётся ЧАСТИЧНЫМ: неизвестные маршруты без расширения, сохранённые
 * ссылки на нативные методы, другие JS-контексты, интерфейс браузера и недоступные
 * наблюдению переходы могут обойти фильтр.
 * Переходы по Blob/Data-ссылкам без имени файла блокируются, включая просмотр.
 * Асинхронное сохранение может ограничиваться политикой браузера, требующей
 * пользовательского действия или разрешения на скачивание.
 * Для распознанного скачивания window.open возвращает null;
 * после проверки файл сохраняется в текущем iframe.
 *
 * Настройки ДО внедрения:
 * window.IFRAME_DOWNLOAD_GUARD_CONFIG = {
 *   isDownloadUrl: url => url.pathname.startsWith('/export/'),
 *   maxFileBytes: 50 * 1024 * 1024,
 *   fetchCredentials: 'same-origin', // для 'include' нужна поддержка CORS сервером
 *   onAttempt: entry => console.log(entry)
 * };
 *
 * API: iframeDownloadGuard.history / .uninstall() / .whenIdle()
 * await iframeDownloadGuard.validateFile(blob, 'report.xlsx') // ошибка при отказе
 * await iframeDownloadGuard.download(url, 'report.xlsx') // проверка и сохранение
 * Событие: window.addEventListener('iframe-download-attempt', e => ...e.detail)
 * Статусы: validating / allowed / blocked / unable-to-block.
 * Статус 'allowed' означает, что файл проверен и запрос на сохранение передан
 * браузеру; завершение скачивания этим статусом не подтверждается.
 */
(() => {
  'use strict';
  if (window.iframeDownloadGuard?.version === 2) return;
  if (window.iframeDownloadGuard) {
    if (typeof window.iframeDownloadGuard.uninstall !== 'function') return;
    window.iframeDownloadGuard.uninstall();
  }

  // НАЧАЛО ПРОВЕРКИ СОДЕРЖИМОГО
  // Определяем формат Office; это не полный разбор документа и не антивирус.
  function requireValid(condition, reason) {
    if (!condition) throw new Error(reason);
  }

  function binaryView(bytes) {
    return new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  }

  async function readBounded(stream, maxBytes) {
    requireValid(stream && typeof stream.getReader === 'function', 'unreadable-body');
    const reader = stream.getReader();
    const chunks = [];
    let size = 0;
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        size += value.byteLength;
        requireValid(size <= maxBytes, 'file-or-part-too-large');
        chunks.push(value);
      }
    } catch (error) {
      await reader.cancel().catch(() => {});
      throw error;
    } finally { reader.releaseLock(); }
    const bytes = new Uint8Array(size);
    let offset = 0;
    for (const chunk of chunks) { bytes.set(chunk, offset); offset += chunk.byteLength; }
    return bytes;
  }

  const crcTable = Uint32Array.from({ length: 256 }, (_, index) => {
    let value = index;
    for (let bit = 0; bit < 8; bit++) value = (value >>> 1) ^ (value & 1 ? 0xedb88320 : 0);
    return value >>> 0;
  });
  function crc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 255];
    return (crc ^ 0xffffffff) >>> 0;
  }

  function parseOfficeXML(bytes) {
    let encoding = 'utf-8';
    if ((bytes[0] === 0xff && bytes[1] === 0xfe) || (bytes[0] === 0x3c && bytes[1] === 0)) encoding = 'utf-16le';
    if ((bytes[0] === 0xfe && bytes[1] === 0xff) || (bytes[0] === 0 && bytes[1] === 0x3c)) encoding = 'utf-16be';
    const text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
    // /<!DOCTYPE|<!ENTITY/i ищет начало объявления DOCTYPE или ENTITY: | означает «или», i игнорирует регистр.
    // Такие объявления отклоняются, чтобы исключить DTD и определения XML-сущностей.
    requireValid(!/<!DOCTYPE|<!ENTITY/i.test(text), 'unsupported-xml-declaration');
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    requireValid(doc.documentElement && !doc.getElementsByTagName('parsererror').length, 'invalid-office-xml');
    return doc;
  }

  function zipDirectory(bytes) {
    const view = binaryView(bytes);
    requireValid(bytes.length >= 22 && view.getUint32(0, true) === 0x04034b50, 'invalid-zip-magic');
    let end = -1;
    for (let pos = bytes.length - 22; pos >= Math.max(0, bytes.length - 65557); pos--) {
      if (view.getUint32(pos, true) === 0x06054b50 && pos + 22 + view.getUint16(pos + 20, true) === bytes.length) {
        end = pos; break;
      }
    }
    requireValid(end >= 0, 'invalid-zip-directory');
    const count = view.getUint16(end + 10, true);
    const size = view.getUint32(end + 12, true);
    const offset = view.getUint32(end + 16, true);
    requireValid(view.getUint16(end + 4, true) === 0 && view.getUint16(end + 6, true) === 0
      && view.getUint16(end + 8, true) === count, 'multi-volume-zip-not-supported');
    requireValid(count > 0 && count < 0xffff && count <= 20000 && size !== 0xffffffff
      && offset !== 0xffffffff && offset + size === end, 'zip64-or-invalid-directory');
    const entries = new Map();
    const ranges = [];
    let cursor = offset;
    for (let index = 0; index < count; index++) {
      requireValid(cursor + 46 <= end && view.getUint32(cursor, true) === 0x02014b50, 'invalid-zip-entry');
      const flags = view.getUint16(cursor + 8, true);
      const method = view.getUint16(cursor + 10, true);
      const compressed = view.getUint32(cursor + 20, true);
      const uncompressed = view.getUint32(cursor + 24, true);
      const nameLength = view.getUint16(cursor + 28, true);
      const extraLength = view.getUint16(cursor + 30, true);
      const commentLength = view.getUint16(cursor + 32, true);
      const local = view.getUint32(cursor + 42, true);
      requireValid(cursor + 46 + nameLength + extraLength + commentLength <= end, 'truncated-zip-entry');
      requireValid(!(flags & 0x41) && view.getUint16(cursor + 34, true) === 0, 'encrypted-or-multi-volume-zip');
      requireValid(compressed !== 0xffffffff && uncompressed !== 0xffffffff && local !== 0xffffffff, 'zip64-not-supported');
      // Имена частей OPC записаны как URI в ASCII или UTF-8; ошибки декодирования блокируют файл.
      const nameBytes = bytes.subarray(cursor + 46, cursor + 46 + nameLength);
      const name = new TextDecoder('utf-8', { fatal: true }).decode(nameBytes);
      // /[\\\x00-\x1f]/ ищет обратную косую черту или управляющий ASCII-символ с кодом 0x00–0x1F.
      // Квадратные скобки задают набор запрещённых символов; обычная / остаётся разделителем частей пути.
      requireValid(name && !name.startsWith('/') && !/[\\\x00-\x1f]/.test(name)
        && !name.split('/').some(part => part === '..' || part === '.') && !entries.has(name), 'invalid-or-duplicate-part');
      requireValid(local + 30 <= offset && view.getUint32(local, true) === 0x04034b50, 'invalid-local-zip-entry');
      requireValid(view.getUint16(local + 6, true) === flags && view.getUint16(local + 8, true) === method, 'zip-header-mismatch');
      const localNameLength = view.getUint16(local + 26, true);
      const dataOffset = local + 30 + localNameLength + view.getUint16(local + 28, true);
      requireValid(localNameLength === nameLength && dataOffset + compressed <= offset, 'zip-entry-out-of-bounds');
      requireValid(nameBytes.every((byte, i) => bytes[local + 30 + i] === byte), 'zip-filename-mismatch');
      const crc = view.getUint32(cursor + 16, true);
      if (!(flags & 8)) {
        requireValid(view.getUint32(local + 14, true) === crc
          && view.getUint32(local + 18, true) === compressed
          && view.getUint32(local + 22, true) === uncompressed, 'zip-size-or-crc-mismatch');
      }
      entries.set(name, { dataOffset, compressed, uncompressed, method, crc });
      ranges.push([local, dataOffset + compressed]);
      cursor += 46 + nameLength + extraLength + commentLength;
    }
    requireValid(cursor === end, 'invalid-central-directory-size');
    ranges.sort((a, b) => a[0] - b[0]);
    for (let i = 1; i < ranges.length; i++) requireValid(ranges[i][0] >= ranges[i - 1][1], 'overlapping-zip-parts');
    return entries;
  }

  async function readZipXML(bytes, entries, name) {
    const entry = entries.get(name);
    requireValid(entry, 'missing-office-part');
    // Ограничиваем размер сжатых и распакованных метаданных; весь ZIP не распаковываем.
    requireValid(entry.uncompressed > 0 && entry.uncompressed <= 4 * 1024 * 1024
      && entry.compressed <= 4 * 1024 * 1024, 'office-metadata-too-large');
    const compressed = bytes.subarray(entry.dataOffset, entry.dataOffset + entry.compressed);
    let content;
    if (entry.method === 0) content = compressed;
    else {
      requireValid(entry.method === 8 && typeof DecompressionStream === 'function', 'zip-compression-not-supported');
      const stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
      content = await readBounded(stream, entry.uncompressed);
    }
    requireValid(content.length === entry.uncompressed && crc32(content) === entry.crc, 'office-part-integrity-failed');
    return parseOfficeXML(content);
  }

  async function identifyOOXML(bytes) {
    const entries = zipDirectory(bytes);
    const typesNS = 'http://schemas.openxmlformats.org/package/2006/content-types';
    const relsNS = 'http://schemas.openxmlformats.org/package/2006/relationships';
    const types = await readZipXML(bytes, entries, '[Content_Types].xml');
    requireValid(types.documentElement.localName === 'Types' && types.documentElement.namespaceURI === typesNS, 'not-office-content-types');
    const parts = Array.from(types.getElementsByTagNameNS(typesNS, 'Override'));
    const declarations = [...parts, ...Array.from(types.getElementsByTagNameNS(typesNS, 'Default'))];
    // /macroenabled|vbaproject/i ищет любой из двух признаков макросов в MIME-типе, без учёта регистра.
    // /(^|\/)vbaproject\.bin$/i ищет файл vbaProject.bin в корне или после / в пути архива.
    // ^ — начало строки, $ — конец, \. — буквальная точка; суффиксы вроде .bin.txt не совпадут.
    requireValid(!declarations.some(node => /macroenabled|vbaproject/i.test(node.getAttribute('ContentType')))
      && !Array.from(entries.keys()).some(name => /(^|\/)vbaproject\.bin$/i.test(name)), 'macro-enabled-ooxml-not-allowed');
    const rels = await readZipXML(bytes, entries, '_rels/.rels');
    requireValid(rels.documentElement.localName === 'Relationships' && rels.documentElement.namespaceURI === relsNS, 'invalid-package-relationships');
    const officeRelations = Array.from(rels.getElementsByTagNameNS(relsNS, 'Relationship')).filter(node => [
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
      'http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument'
    ].includes(node.getAttribute('Type')));
    requireValid(officeRelations.length === 1 && officeRelations[0].getAttribute('TargetMode') !== 'External', 'invalid-office-main-relationship');
    const target = officeRelations[0].getAttribute('Target');
    // /[\\?#]/ ищет обратную косую черту, ? или #: запрещаем неоднозначный путь, query-параметры и фрагмент.
    requireValid(target && !/[\\?#]/.test(target), 'invalid-office-main-target');
    const mainURL = new URL(target, 'https://office-package.invalid/');
    requireValid(mainURL.origin === 'https://office-package.invalid', 'external-office-main-part');
    const mainName = decodeURIComponent(mainURL.pathname.slice(1));
    const mainTypes = parts.filter(node => {
      try { return decodeURIComponent(node.getAttribute('PartName')) === '/' + mainName; }
      catch { return false; }
    });
    requireValid(mainTypes.length === 1, 'missing-or-ambiguous-main-content-type');
    const mainType = mainTypes[0].getAttribute('ContentType');
    const formats = {
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml': {
        format: 'xlsx', root: 'workbook', namespaces: ['http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main']
      },
      'application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml': {
        format: 'pptx', root: 'presentation', namespaces: ['http://schemas.openxmlformats.org/presentationml/2006/main', 'http://purl.oclc.org/ooxml/presentationml/main']
      }
    };
    const format = formats[mainType];
    requireValid(format, 'unsupported-office-format');
    const main = await readZipXML(bytes, entries, mainName);
    requireValid(main.documentElement.localName === format.root && format.namespaces.includes(main.documentElement.namespaceURI), 'office-main-part-mismatch');
    return format.format;
  }

  function identifyXLS(bytes) {
    // CFB используется также в DOC/PPT: требуем поток Excel Workbook/Book в корне.
    const view = binaryView(bytes);
    requireValid(bytes.length >= 512, 'truncated-cfb');
    const version = view.getUint16(26, true);
    const sectorSize = 2 ** view.getUint16(30, true);
    requireValid(view.getUint16(28, true) === 0xfffe && view.getUint16(32, true) === 6
      && ((version === 3 && sectorSize === 512) || (version === 4 && sectorSize === 4096))
      && view.getUint32(56, true) === 4096 && bytes.length % sectorSize === 0, 'invalid-cfb-header');
    const sectorCount = bytes.length / sectorSize - 1;
    const END = 0xfffffffe;
    const FREE = 0xffffffff;
    const sector = id => {
      requireValid(Number.isInteger(id) && id >= 0 && id < sectorCount, 'cfb-sector-out-of-bounds');
      return bytes.subarray((id + 1) * sectorSize, (id + 2) * sectorSize);
    };
    const fatCount = view.getUint32(44, true);
    requireValid(fatCount > 0 && fatCount <= sectorCount, 'invalid-cfb-fat-count');
    const fatIds = [];
    const addFatIds = data => {
      const dataView = binaryView(data);
      for (let p = 0; p < data.length; p += 4) {
        const id = dataView.getUint32(p, true);
        if (id !== FREE) fatIds.push(id);
      }
    };
    addFatIds(bytes.subarray(76, 512));
    let difat = view.getUint32(68, true);
    const difatCount = view.getUint32(72, true);
    requireValid(difatCount <= sectorCount, 'invalid-cfb-difat-count');
    const seenDIFAT = new Set();
    for (let i = 0; i < difatCount; i++) {
      requireValid(!seenDIFAT.has(difat), 'cyclic-cfb-difat');
      seenDIFAT.add(difat);
      const data = sector(difat);
      addFatIds(data.subarray(0, sectorSize - 4));
      difat = binaryView(data).getUint32(sectorSize - 4, true);
    }
    requireValid(fatIds.length === fatCount && new Set(fatIds).size === fatIds.length, 'invalid-cfb-fat-list');
    requireValid(difatCount === 0 || difat === END, 'invalid-cfb-difat-end');
    const fat = new Uint32Array(fatCount * sectorSize / 4);
    fatIds.forEach((id, index) => {
      const data = binaryView(sector(id));
      for (let p = 0; p < sectorSize; p += 4) fat[index * sectorSize / 4 + p / 4] = data.getUint32(p, true);
    });
    function readChain(start, table, blockSize, getBlock, size = null) {
      requireValid(size === null || (Number.isSafeInteger(size) && size >= 0 && size <= bytes.length), 'invalid-cfb-stream-size');
      const chunks = [];
      const seen = new Set();
      let id = start;
      let length = 0;
      while (id !== END) {
        requireValid(id < table.length && !seen.has(id), 'invalid-or-cyclic-cfb-chain');
        seen.add(id);
        const block = getBlock(id);
        chunks.push(block);
        length += blockSize;
        requireValid(length <= bytes.length, 'cfb-chain-too-large');
        id = table[id];
        if (size !== null && length >= size) {
          requireValid(id === END, 'cfb-chain-length-mismatch');
          break;
        }
      }
      requireValid(size === null || length >= size, 'truncated-cfb-stream');
      const result = new Uint8Array(size ?? length);
      for (let i = 0; i < chunks.length; i++) result.set(chunks[i].subarray(0, Math.min(blockSize, result.length - i * blockSize)), i * blockSize);
      return result;
    }
    const directory = readChain(view.getUint32(48, true), fat, sectorSize, sector);
    requireValid(directory.length >= 128, 'missing-cfb-directory');
    const dirView = binaryView(directory);
    function directoryEntry(id) {
      const base = id * 128;
      requireValid(base + 128 <= directory.length, 'invalid-cfb-directory-id');
      const nameLength = dirView.getUint16(base + 64, true);
      requireValid(nameLength >= 2 && nameLength <= 64 && nameLength % 2 === 0, 'invalid-cfb-name');
      const size = Number(dirView.getBigUint64(base + 120, true));
      return {
        name: new TextDecoder('utf-16le').decode(directory.subarray(base, base + nameLength - 2)),
        type: directory[base + 66], left: dirView.getUint32(base + 68, true),
        right: dirView.getUint32(base + 72, true), child: dirView.getUint32(base + 76, true),
        start: dirView.getUint32(base + 116, true), size
      };
    }
    const root = directoryEntry(0);
    requireValid(root.type === 5, 'missing-cfb-root');
    const queue = [root.child];
    const seenEntries = new Set();
    const workbooks = [];
    while (queue.length) {
      const id = queue.pop();
      if (id === FREE) continue;
      requireValid(!seenEntries.has(id), 'cyclic-cfb-directory');
      seenEntries.add(id);
      const entry = directoryEntry(id);
      if (entry.type === 2 && ['Workbook', 'Book'].includes(entry.name)) workbooks.push(entry);
      requireValid(!['EncryptedPackage', 'EncryptionInfo'].includes(entry.name), 'encrypted-office-not-supported');
      queue.push(entry.left, entry.right);
    }
    requireValid(workbooks.length === 1, 'not-an-excel-compound-file');
    const workbook = workbooks[0];
    requireValid(workbook.size >= 12, 'truncated-excel-workbook');
    let data;
    if (workbook.size >= 4096) data = readChain(workbook.start, fat, sectorSize, sector, workbook.size);
    else {
      const miniFatBytes = readChain(view.getUint32(60, true), fat, sectorSize, sector, view.getUint32(64, true) * sectorSize);
      const miniFatView = binaryView(miniFatBytes);
      const miniFat = Uint32Array.from({ length: miniFatBytes.length / 4 }, (_, i) => miniFatView.getUint32(i * 4, true));
      const miniStream = readChain(root.start, fat, sectorSize, sector, root.size);
      data = readChain(workbook.start, miniFat, 64, id => {
        requireValid((id + 1) * 64 <= miniStream.length, 'invalid-cfb-mini-sector');
        return miniStream.subarray(id * 64, (id + 1) * 64);
      }, workbook.size);
    }
    const workbookView = binaryView(data);
    requireValid(workbookView.getUint16(0, true) === 0x0809
      && [0x0500, 0x0600].includes(workbookView.getUint16(4, true))
      && workbookView.getUint16(6, true) === 0x0005
      && workbookView.getUint16(2, true) >= 8, 'invalid-excel-bof');
    let cursor = 0;
    let hasEOF = false;
    while (cursor + 4 <= data.length) {
      const type = workbookView.getUint16(cursor, true);
      const length = workbookView.getUint16(cursor + 2, true);
      requireValid(cursor + 4 + length <= data.length, 'truncated-excel-record');
      requireValid(type !== 0x002f, 'encrypted-xls-not-supported');
      if (type === 0x000a) { requireValid(length === 0, 'invalid-excel-eof'); hasEOF = true; break; }
      cursor += 4 + length;
    }
    requireValid(hasEOF, 'missing-excel-eof');
    return 'xls';
  }

  async function identifyOffice(bytes) {
    requireValid(bytes.length >= 8, 'file-too-short');
    const oleMagic = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
    if (oleMagic.every((byte, i) => bytes[i] === byte)) return identifyXLS(bytes);
    if (binaryView(bytes).getUint32(0, true) === 0x04034b50) return identifyOOXML(bytes);
    throw new Error('unsupported-magic-bytes');
  }
  // КОНЕЦ ПРОВЕРКИ СОДЕРЖИМОГО


  const config = window.IFRAME_DOWNLOAD_GUARD_CONFIG || {};
  const allowed = new Set(['xls', 'xlsx', 'pptx']);
  const mimeTypes = {
    xls: 'application/vnd.ms-excel',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
  };
  const pageExtensions = new Set(['html', 'htm', 'php', 'asp', 'aspx', 'jsp', 'cgi']);
  // Эти схемы открывают внешнее приложение и не скачивают файлы.
  const contactProtocols = new Set(['mailto:', 'tel:', 'sms:']);
  const positiveInteger = (value, fallback) => Number.isSafeInteger(value) && value > 0 ? value : fallback;
  const maxFileBytes = positiveInteger(config.maxFileBytes, 50 * 1024 * 1024);
  const maxPendingChecks = positiveInteger(config.maxPendingChecks, 3);
  const timeoutMs = positiveInteger(config.timeoutMs, 30000);
  const fetchCredentials = ['same-origin', 'include', 'omit'].includes(config.fetchCredentials) ? config.fetchCredentials : 'same-origin';
  const history = [];
  const cleanup = [];
  const blobs = new Map();
  const warnings = [];
  const pending = new Set();
  const controllers = new Set();
  const approvedLinks = new WeakSet();
  const approvedURLs = new Map();
  let active = true;

  const nativeFetch = window.fetch?.bind(window);
  const nativeClick = HTMLElement.prototype.click;
  const nativeCreateURL = URL.createObjectURL.bind(URL);
  const nativeRevokeURL = URL.revokeObjectURL.bind(URL);
  const nativeDispatch = EventTarget.prototype.dispatchEvent;
  // /\.([a-z0-9]+)$/i извлекает расширение после буквальной точки в конце имени.
  // Группа (...) сохраняет одну или больше ASCII-букв/цифр; $ привязывает к концу, i игнорирует регистр.
  // Например, для report.XLSX группа [1] содержит XLSX, а для report.xlsx.exe — exe.
  const extension = name => /\.([a-z0-9]+)$/i.exec(name)?.[1].toLowerCase() || '';
  // /[\x00-\x1f\x7f/\\]/ ищет управляющие ASCII-символы 0x00–0x1F, DEL (0x7F), / или обратную косую черту.
  // /[.\s]$/ проверяет последний символ: точку или пробельный символ; \s включает пробелы, табуляцию и переводы строк.
  // Обе проверки используются с !, поэтому совпадение делает имя недопустимым.
  const safeFilename = name => typeof name === 'string' && name.length > 0 && name.length <= 255
    && !/[\x00-\x1f\x7f/\\]/.test(name) && !/[.\s]$/.test(name);
  const isLink = node => node instanceof HTMLAnchorElement || node instanceof HTMLAreaElement;

  function assess(rawUrl, explicitDownload = false, suggestedName = '', trigger = '') {
    let url;
    try { url = new URL(rawUrl, document.baseURI); }
    catch {
      return explicitDownload ? { trigger, url: '[invalid]', filename: suggestedName, eligible: false, reason: 'invalid-url' } : null;
    }
    // Иначе домен в mailto:info@example.ru распознавался бы как расширение .ru и ссылка блокировалась.
    if (!explicitDownload && contactProtocols.has(url.protocol)) return null;
    const local = url.protocol === 'blob:' || url.protocol === 'data:';
    let pathName = '';
    if (!local) {
      try { pathName = decodeURIComponent(url.pathname.split('/').pop() || ''); } catch {}
    }
    const filename = suggestedName || pathName;
    let customDownload = false;
    try { customDownload = Boolean(config.isDownloadUrl?.(url)); } catch { customDownload = true; }
    const pathExt = extension(pathName);
    if (!explicitDownload && !local && !(pathExt && !pageExtensions.has(pathExt)) && !customDownload) return null;
    const protocolOK = ['http:', 'https:', 'blob:', 'data:'].includes(url.protocol);
    const nameOK = safeFilename(filename);
    const extOK = allowed.has(extension(filename));
    return {
      trigger, requestUrl: url.href,
      url: url.protocol === 'data:' ? 'data:[omitted]' : url.href,
      filename, eligible: protocolOK && nameOK && extOK,
      reason: !protocolOK ? 'unsupported-protocol' : !nameOK ? 'missing-or-invalid-filename'
        : !extOK ? 'blocked-extension' : 'content-check-required'
    };
  }

  function report(decision, status, extra = {}) {
    const { requestUrl, eligible, ...publicFields } = decision;
    const entry = Object.freeze({ ...publicFields, ...extra, status,
      allowed: status === 'allowed' ? true : status === 'validating' ? null : false,
      time: new Date().toISOString() });
    history.push(entry);
    if (history.length > 200) history.shift();
    try { config.onAttempt?.(entry); } catch (error) { console.error(error); }
    Reflect.apply(nativeDispatch, window, [new CustomEvent('iframe-download-attempt', { detail: entry })]);
    if (!config.onAttempt) console.info('[iframe-download-guard]', entry);
  }

  async function validateBytes(bytes, filename) {
    requireValid(safeFilename(filename) && allowed.has(extension(filename)), 'blocked-extension');
    requireValid(bytes.byteLength <= maxFileBytes, 'file-too-large');
    const format = await identifyOffice(bytes);
    requireValid(format === extension(filename), 'extension-content-mismatch');
    return { format, size: bytes.byteLength };
  }

  async function validateFile(blob, filename) {
    requireValid(blob && typeof blob.arrayBuffer === 'function' && Number.isSafeInteger(blob.size), 'expected-blob');
    requireValid(blob.size <= maxFileBytes, 'file-too-large');
    return validateBytes(new Uint8Array(await blob.arrayBuffer()), filename);
  }

  async function loadBytes(decision, capturedBlob, controller) {
    if (capturedBlob) {
      requireValid(capturedBlob.size <= maxFileBytes, 'file-too-large');
      return new Uint8Array(await capturedBlob.arrayBuffer());
    }
    requireValid(nativeFetch, 'fetch-not-available');
    const response = await nativeFetch(decision.requestUrl, {
      method: 'GET', credentials: fetchCredentials, signal: controller.signal,
      redirect: 'follow', mode: 'cors'
    });
    requireValid(response.ok && !['opaque', 'opaqueredirect'].includes(response.type), 'unreadable-http-response');
    const lengthHeader = response.headers.get('content-length');
    requireValid(!lengthHeader || Number(lengthHeader) <= maxFileBytes, 'file-too-large');
    return readBounded(response.body, maxFileBytes);
  }

  function releaseApprovedURL(url) {
    clearTimeout(approvedURLs.get(url));
    approvedURLs.delete(url);
    nativeRevokeURL(url);
  }

  function saveCheckedBytes(bytes, filename, format) {
    requireValid(active, 'guard-uninstalled');
    const url = nativeCreateURL(new Blob([bytes], { type: mimeTypes[format] }));
    approvedURLs.set(url, setTimeout(() => releaseApprovedURL(url), 60000));
    const link = document.createElement('a');
    link.href = url; link.download = filename; link.hidden = true;
    approvedLinks.add(link);
    document.documentElement.append(link);
    try { Reflect.apply(nativeClick, link, []); }
    catch (error) { releaseApprovedURL(url); throw error; }
    finally { link.remove(); }
  }

  function startCheck(decision) {
    if (!active) return Promise.resolve(false);
    if (!decision.eligible) { report(decision, 'blocked'); return Promise.resolve(false); }
    if (pending.size >= maxPendingChecks) {
      report(decision, 'blocked', { reason: 'too-many-pending-checks' });
      return Promise.resolve(false);
    }
    // Сохраняем Blob синхронно: FileSaver может отозвать URL сразу после клика.
    const capturedBlob = blobs.get(decision.requestUrl);
    const controller = new AbortController();
    controllers.add(controller);
    const timer = setTimeout(() => controller.abort(), timeoutMs);
    const operation = Promise.resolve().then(async () => {
      try {
        requireValid(active && !controller.signal.aborted, 'check-aborted');
        report(decision, 'validating');
        const bytes = await loadBytes(decision, capturedBlob, controller);
        const result = await validateBytes(bytes, decision.filename);
        requireValid(active && !controller.signal.aborted, 'check-aborted');
        saveCheckedBytes(bytes, decision.filename, result.format);
        report(decision, 'allowed', { reason: 'content-verified', ...result });
        return true;
      } catch (error) {
        report(decision, 'blocked', { reason: controller.signal.aborted ? 'check-aborted-or-timeout' : error.message || 'content-check-failed' });
        return false;
      } finally {
        clearTimeout(timer);
        controller.abort(); // Прерываем чтение ответа, если проверка завершилась раньше.
        controllers.delete(controller);
      }
    });
    pending.add(operation);
    void operation.then(() => pending.delete(operation), () => pending.delete(operation));
    return operation;
  }

  function interceptEvent(event, decision) {
    if (!decision || event.defaultPrevented) return;
    if (!event.cancelable) {
      report(decision, 'unable-to-block', { reason: 'event-not-cancelable' });
      return;
    }
    event.preventDefault();
    event.stopImmediatePropagation();
    void startCheck(decision);
  }

  function listen(target, name, fn) {
    target.addEventListener(name, fn, { capture: true, passive: false });
    cleanup.push(() => target.removeEventListener(name, fn, true));
  }
  function patch(target, name, factory) {
    const descriptor = Object.getOwnPropertyDescriptor(target, name);
    const original = target[name];
    if (typeof original !== 'function') return;
    const replacement = factory(original);
    try {
      Object.defineProperty(target, name, descriptor ? { ...descriptor, value: replacement }
        : { value: replacement, configurable: true, writable: true });
      cleanup.push(() => {
        if (target[name] !== replacement) return;
        if (descriptor) Object.defineProperty(target, name, descriptor); else delete target[name];
      });
    } catch (error) { warnings.push(`Could not patch ${name}: ${error.message}`); }
  }
  const linkDecision = (link, trigger) => assess(link.href, link.hasAttribute('download'), link.getAttribute('download') || '', trigger);

  for (const type of ['click', 'auxclick']) listen(window, type, event => {
    if (type === 'auxclick' && event.button !== 1) return;
    const link = event.composedPath().find(isLink);
    if (link && !approvedLinks.has(link)) interceptEvent(event, linkDecision(link, type));
  });

  patch(HTMLElement.prototype, 'click', native => function (...args) {
    if (isLink(this) && !approvedLinks.has(this)) {
      const decision = linkDecision(this, 'element.click');
      if (decision) { void startCheck(decision); return; }
    }
    return Reflect.apply(native, this, args);
  });

  patch(EventTarget.prototype, 'dispatchEvent', native => function (event) {
    if (isLink(this) && !approvedLinks.has(this) && (event?.type === 'click'
      || (event?.type === 'auxclick' && event.button === 1))) {
      const decision = linkDecision(this, 'dispatchEvent');
      if (decision) {
        if (event.cancelable) event.preventDefault();
        void startCheck(decision); return false;
      }
    }
    return Reflect.apply(native, this, [event]);
  });

  patch(window, 'open', native => function (url, ...args) {
    const decision = assess(url ?? 'about:blank', false, '', 'window.open');
    if (decision) { void startCheck(decision); return null; }
    return Reflect.apply(native, this, [url, ...args]);
  });

  function formDecision(form, submitter, trigger) {
    const action = submitter?.hasAttribute('formaction') ? submitter.formAction : form.action;
    const method = String(submitter?.hasAttribute('formmethod') ? submitter.formMethod : form.method).toLowerCase();
    if (method === 'dialog') return null;
    const decision = assess(action, false, '', trigger);
    if (!decision) return null;
    if (method !== 'get') return { ...decision, eligible: false, reason: 'post-download-not-supported' };
    try {
      const data = submitter ? new FormData(form, submitter) : new FormData(form);
      const query = new URLSearchParams();
      for (const [key, value] of data) query.append(key, typeof value === 'string' ? value : value.name);
      const url = new URL(action, document.baseURI);
      url.search = query.toString();
      return { ...decision, requestUrl: url.href, url: url.href };
    } catch { return { ...decision, eligible: false, reason: 'unreadable-form' }; }
  }
  listen(window, 'submit', event => {
    if (event.target instanceof HTMLFormElement) interceptEvent(event, formDecision(event.target, event.submitter, 'submit'));
  });
  patch(HTMLFormElement.prototype, 'submit', native => function (...args) {
    if (this instanceof HTMLFormElement) {
      const decision = formDecision(this, null, 'form.submit');
      if (decision) { void startCheck(decision); return; }
    }
    return Reflect.apply(native, this, args);
  });

  patch(URL, 'createObjectURL', native => function (object) {
    const url = Reflect.apply(native, this, [object]);
    if (object instanceof Blob) blobs.set(url, object);
    return url;
  });
  patch(URL, 'revokeObjectURL', native => function (url) {
    const result = Reflect.apply(native, this, [url]);
    blobs.delete(String(url));
    return result;
  });

  if (window.navigation && window.navigation.currentEntry) {
    listen(window.navigation, 'navigate', event => {
      if (approvedURLs.has(event.destination.url)) return;
      const decision = assess(event.destination.url, typeof event.downloadRequest === 'string', event.downloadRequest || '', 'navigate');
      if (decision && event.formData) { decision.eligible = false; decision.reason = 'post-download-not-supported'; }
      interceptEvent(event, decision);
    });
  } else warnings.push('Navigation API is unavailable in this document; navigation checks are inactive.');

  const api = Object.freeze({
    version: 2,
    allowedExtensions: Object.freeze([...allowed]), warnings: Object.freeze(warnings),
    get history() { return history.slice(); },
    validateFile,
    download(url, filename) { return startCheck(assess(url, true, filename, 'api.download')); },
    async whenIdle() { while (pending.size) await Promise.allSettled([...pending]); },
    uninstall() {
      active = false;
      for (const controller of controllers) controller.abort();
      cleanup.splice(0).reverse().forEach(fn => fn());
      for (const url of [...approvedURLs.keys()]) releaseApprovedURL(url);
      blobs.clear();
      if (window.iframeDownloadGuard === api) delete window.iframeDownloadGuard;
    }
  });
  window.iframeDownloadGuard = api;
  for (const warning of warnings) console.warn('[iframe-download-guard]', warning);
})();
