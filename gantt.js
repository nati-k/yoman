(() => {
    'use strict';
    const DAYS = ['ראשון', 'שני', 'שלישי', 'רביעי', 'חמישי'];
    const DEFAULTS = [
        ['עבודה בגנים', '#FADBD8'], ['מעבר בין גנים', '#FFF2CC'],
        ["סופרוויז'ן קבוצתי / ישיבת צוות", '#E2EFDA'],
        ['קמפוס סייעות / הכשרת תומכות', '#D9E1F2'],
        ["סופרוויז'ן דואט / נמל", '#E8DDF2'], ['פגישות ממשקים', '#D1F2EB'],
        ['ישיבות צוות בגנים', '#BDD7EE'], ['הדרכה אישית לאשת צוות', '#FCE4D6'],
        ["סופרוויז'ן אישי חודשי", '#E2EFDA'],
        ["סופרוויז'ן ללמוד לחיות ביחד/הורים", '#E8DDF2'],
        ['העברת הכשרות לצוותים והורים', '#D9E1F2']
    ].map(([name, color], index) => ({ id: `type-${index}`, name, color }));
    const state = { owner: null, accountId: null, model: null, file: null, dirty: false, busy: false, timer: null, saving: null, day: new Date().getDay() % 5, exportFile: null, exportPromise: null, exportVersion: 0, exportError: false };
    const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const html = value => escapeHtml(String(value ?? ''));
    const icon = name => `<i data-lucide="${name}" aria-hidden="true"></i>`;
    const time = minutes => `${String(Math.floor(minutes / 60)).padStart(2, '0')}:${String(minutes % 60).padStart(2, '0')}`;
    const minutes = value => Number(value.split(':')[0]) * 60 + Number(value.split(':')[1]);
    const iso = date => `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
    const dateAt = value => new Date(`${value}T12:00:00`);
    const shift = (week, count) => { const date = dateAt(week); date.setDate(date.getDate() + count); return iso(date); };
    const sunday = value => { const date = dateAt(value); date.setDate(date.getDate() - date.getDay()); return iso(date); };
    const shortDate = value => dateAt(value).toLocaleDateString('he-IL', { day: '2-digit', month: '2-digit', year: 'numeric' });
    const title = week => `גאנט למדריכה ${shortDate(week)}–${shortDate(shift(week, 4))}`;
    const blank = week => ({ version: 1, week, activities: structuredClone(DEFAULTS), events: [] });
    const draftKey = week => `yoman_gantt:${state.owner}:${week}`;
    const isIOS = () => /iPad|iPhone|iPod/.test(navigator.userAgent) || (navigator.platform === 'MacIntel' && navigator.maxTouchPoints > 1);
    const redrawIcons = () => globalThis.lucide?.createIcons();
    document.addEventListener('DOMContentLoaded', redrawIcons, { once: true });

    function lanes(events) {
        const ends = [];
        const arranged = [...events].sort((first, second) => first.start - second.start || first.end - second.end).map(event => {
            let lane = ends.findIndex(end => end <= event.start);
            if (lane < 0) lane = ends.length;
            ends[lane] = event.end;
            return { ...event, lane };
        });
        return { count: Math.max(1, ends.length), events: arranged };
    }

    function validateModel(model, week) {
        if (model.version !== 1 || model.week !== week || !Array.isArray(model.activities) || !Array.isArray(model.events)) throw new Error('קובץ השבוע אינו בפורמט של הפורטל.');
        if (model.activities.some(activity => !activity.id || typeof activity.name !== 'string' || !/^#[0-9a-f]{6}$/i.test(activity.color))) throw new Error('מקראת הפעילויות אינה תקינה.');
        if (model.events.some(event => !event.id || !Number.isInteger(event.day) || event.day < 0 || event.day > 4 || !Number.isInteger(event.start) || !Number.isInteger(event.end) || event.start < 480 || event.end > 1200 || event.start >= event.end || event.start % 30 || event.end % 30 || !model.activities.some(activity => activity.id === event.activity))) throw new Error('נתוני הלו״ז אינם תקינים.');
        return model;
    }

    function legendFor(model) {
        return model.activities.filter(activity => {
            const original = DEFAULTS.find(item => item.id === activity.id);
            return !original || Number(activity.id.slice(5)) < 8 || activity.name !== original.name || activity.color.toUpperCase() !== original.color;
        });
    }

    function dimensionsFor(model) {
        const layouts = DAYS.map((day, index) => lanes(model.events.filter(event => event.day === index)));
        const heights = [35, 20, 12, 30, ...Array(24).fill(24), 15, 25, 10, ...Array(Math.ceil(legendFor(model).length / 4)).fill(26)];
        const widths = [16, ...layouts.flatMap(layout => Array(layout.count).fill(25 / layout.count))];
        return { heights, widths, layouts };
    }

    function dimensionRequests(model) {
        const dimensions = dimensionsFor(model);
        // Adjacent rows or columns with the same size share one request.
        const grouped = (dimension, sizes) => {
            const requests = [];
            sizes.forEach((size, index) => {
                const last = requests.at(-1)?.updateDimensionProperties;
                if (last && last.properties.pixelSize === size) last.range.endIndex = index + 1;
                else requests.push({ updateDimensionProperties: { range: { sheetId: 0, dimension, startIndex: index, endIndex: index + 1 }, properties: { pixelSize: size }, fields: 'pixelSize' } });
            });
            return requests;
        };
        return [
            ...grouped('ROWS', dimensions.heights.map(height => Math.round(height * 96 / 72))),
            ...grouped('COLUMNS', dimensions.widths.map(width => Math.floor(width * 7 + 5)))
        ];
    }

    function sheetsFor(model) {
        const layouts = DAYS.map((day, index) => lanes(model.events.filter(event => event.day === index)));
        const columns = 1 + layouts.reduce((sum, layout) => sum + layout.count, 0);
        const rgb = color => ({ red: parseInt(color.slice(1, 3), 16) / 255, green: parseInt(color.slice(3, 5), 16) / 255, blue: parseInt(color.slice(5, 7), 16) / 255 });
        const profiles = {
            plain: { textFormat: { fontFamily: 'Arial', fontSize: 11, bold: false } },
            title: { textFormat: { fontFamily: 'Calibri', fontSize: 16, bold: true, foregroundColor: rgb('#203764') } },
            subtitle: { textFormat: { fontFamily: 'Calibri', fontSize: 11, bold: false, foregroundColor: rgb('#595959') } },
            heading: { textFormat: { fontFamily: 'Calibri', fontSize: 11, bold: true, foregroundColor: rgb('#203764') } },
            body: { textFormat: { fontFamily: 'Calibri', fontSize: 10, bold: false, foregroundColor: rgb('#000000') } },
            activity: { textFormat: { fontFamily: 'Calibri', fontSize: 10, bold: true, foregroundColor: rgb('#000000') } },
            legendTitle: { textFormat: { fontFamily: 'Calibri', fontSize: 12, bold: true, foregroundColor: rgb('#203764') } }
        };
        const cell = (text = '', color = null, profile = 'plain') => {
            const format = structuredClone(profiles[profile]);
            if (profile !== 'plain') Object.assign(format, { horizontalAlignment: profile === 'legendTitle' ? 'RIGHT' : 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: profile === 'legendTitle' ? 'OVERFLOW_CELL' : 'WRAP' });
            if (color) Object.assign(format, { backgroundColor: rgb(color), borders: Object.fromEntries(['top', 'bottom', 'left', 'right'].map(side => [side, { style: 'SOLID', color: rgb('#D3D3D3') }])) });
            return { userEnteredValue: { stringValue: String(text) }, userEnteredFormat: format };
        };
        const legend = legendFor(model);
        const rows = Array.from({ length: 31 + Math.ceil(legend.length / 4) }, () => ({ values: Array.from({ length: columns }, () => cell()) }));
        rows[0].values[0] = cell('מערכת שעות שבועית', null, 'title');
        rows[1].values[0] = cell('מלווה גני ילדים - תכנון שבועי מלא', null, 'subtitle');
        rows[3].values[0] = cell('שעות', '#C9D0F8', 'heading');
        const merges = [{ sheetId: 0, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: columns }, { sheetId: 0, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 0, endColumnIndex: columns }, { sheetId: 0, startRowIndex: 29, endRowIndex: 30, startColumnIndex: 0, endColumnIndex: columns }];
        const displayTime = value => `${Math.floor(value / 60)}:${String(value % 60).padStart(2, '0')}`;
        for (let slot = 0; slot < 24; slot++) rows[slot + 4].values[0] = cell(`${displayTime(480 + slot * 30)}-${displayTime(510 + slot * 30)}`, '#FFFFFF', 'heading');
        let column = 1;
        layouts.forEach((layout, day) => {
            for (let offset = 0; offset < layout.count; offset++) {
                rows[3].values[column + offset] = cell(offset ? '' : DAYS[day], '#C9D0F8', 'heading');
                for (let row = 4; row < 28; row++) rows[row].values[column + offset] = cell('', '#FFFFFF', 'body');
            }
            if (layout.count > 1) merges.push({ sheetId: 0, startRowIndex: 3, endRowIndex: 4, startColumnIndex: column, endColumnIndex: column + layout.count });
            layout.events.forEach(event => {
                const activity = model.activities.find(item => item.id === event.activity);
                const startRow = 4 + (event.start - 480) / 30;
                const endRow = 4 + (event.end - 480) / 30;
                const label = [event.place, event.text].filter(Boolean).join('\n') || activity.name;
                for (let row = startRow; row < endRow; row++) rows[row].values[column + event.lane] = cell(row === startRow ? label : '', activity.color, activity.id === 'type-1' ? 'body' : 'activity');
                if (endRow - startRow > 1) merges.push({ sheetId: 0, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: column + event.lane, endColumnIndex: column + event.lane + 1 });
            });
            column += layout.count;
        });
        rows[29].values[0] = cell('מקרא הצבעים', null, 'legendTitle');
        legend.forEach((activity, index) => {
            const day = index % 4;
            const startColumn = 1 + layouts.slice(0, day).reduce((sum, layout) => sum + layout.count, 0);
            const endColumn = startColumn + layouts[day].count;
            const row = 31 + Math.floor(index / 4);
            for (let offset = startColumn; offset < endColumn; offset++) rows[row].values[offset] = cell(offset === startColumn ? activity.name : '', activity.color, 'activity');
            if (endColumn - startColumn > 1) merges.push({ sheetId: 0, startRowIndex: row, endRowIndex: row + 1, startColumnIndex: startColumn, endColumnIndex: endColumn });
        });
        const dataRows = ['yoman-gantt-v1', model.week, JSON.stringify(model.activities), ...model.events.map(event => JSON.stringify(event))].map(value => ({ values: [cell(value)] }));
        if (dataRows.some(row => row.values[0].userEnteredValue.stringValue.length > 45000)) throw new Error('הנתונים ארוכים מדי לשמירה.');
        return [
            { properties: { sheetId: 0, title: 'מערכת שעות שבועית', rightToLeft: true, gridProperties: { rowCount: rows.length, columnCount: columns, frozenRowCount: 0, frozenColumnCount: 0, hideGridlines: true } }, data: [{ startRow: 0, startColumn: 0, rowData: rows }], merges },
            { properties: { sheetId: 1, title: '_portal_data', hidden: true, gridProperties: { rowCount: Math.max(100, dataRows.length), columnCount: 1 } }, data: [{ startRow: 0, startColumn: 0, rowData: dataRows }] }
        ];
    }

    let templateBytes = null;

    async function templateXlsx(model) {
        if (!globalThis.JSZip) throw new Error('ספריית האקסל עדיין לא נטענה. נסה שוב בעוד רגע.');
        if (!templateBytes) {
            templateBytes = fetch(new URL('גאנט למדריכה ריק_8220.xlsx', document.baseURI)).then(response => {
                if (!response.ok) throw new Error('לא ניתן לטעון את תבנית האקסל המקורית.');
                return response.arrayBuffer();
            }).catch(error => { templateBytes = null; throw error; });
        }
        const zip = await JSZip.loadAsync(await templateBytes);
        const parse = async path => {
            const xml = new DOMParser().parseFromString(await zip.file(path).async('string'), 'application/xml');
            if (xml.getElementsByTagName('parsererror').length) throw new Error('תבנית האקסל אינה תקינה.');
            return xml;
        };
        const worksheet = await parse('xl/worksheets/sheet1.xml');
        const styles = await parse('xl/styles.xml');
        const strings = await parse('xl/sharedStrings.xml');
        const namespace = worksheet.documentElement.namespaceURI;
        const element = (xml, name, attributes = {}) => {
            const node = xml.createElementNS(namespace, name);
            Object.entries(attributes).forEach(([key, value]) => node.setAttribute(key, String(value)));
            return node;
        };
        const serialize = xml => {
            const content = new XMLSerializer().serializeToString(xml).replace(/^<\?xml[^?]*\?>\s*/i, '');
            return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${content}`;
        };
        const fills = styles.getElementsByTagName('fills')[0];
        const formats = styles.getElementsByTagName('cellXfs')[0];
        const originalFillCount = fills.children.length;
        const originalFormatCount = formats.children.length;
        const colors = color => [color.red || 0, color.green || 0, color.blue || 0].map(channel => Math.round(channel * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
        const styleFor = cell => {
            const format = cell.userEnteredFormat;
            const font = format.textFormat;
            if (!format.backgroundColor) return font.fontSize === 16 ? 23 : font.fontSize === 12 ? 22 : font.fontFamily === 'Calibri' ? 15 : 16;
            const color = `FF${colors(format.backgroundColor)}`;
            let fillIndex = [...fills.children].findIndex(fill => fill.getElementsByTagName('fgColor')[0]?.getAttribute('rgb') === color);
            if (fillIndex < 0) {
                const fill = fills.children[4].cloneNode(true);
                fill.getElementsByTagName('fgColor')[0].setAttribute('rgb', color);
                fill.getElementsByTagName('bgColor')[0].setAttribute('rgb', color);
                fillIndex = fills.children.length;
                fills.append(fill);
            }
            const fontId = font.fontSize === 11 ? 3 : font.bold ? 5 : 4;
            let styleIndex = [...formats.children].findIndex(style => Number(style.getAttribute('fontId')) === fontId && Number(style.getAttribute('fillId')) === fillIndex && style.getAttribute('borderId') === '1');
            if (styleIndex < 0) {
                const style = formats.children[font.bold ? 13 : 4].cloneNode(true);
                style.setAttribute('fontId', String(fontId));
                style.setAttribute('fillId', String(fillIndex));
                styleIndex = formats.children.length;
                formats.append(style);
            }
            return styleIndex;
        };
        const shared = [...strings.documentElement.children].map(item => [...item.getElementsByTagName('t')].map(text => text.textContent).join(''));
        const stringIndices = new Map(shared.map((text, index) => [text, index]));
        let stringCount = 0;
        const stringFor = text => {
            stringCount++;
            if (!stringIndices.has(text)) {
                const item = element(strings, 'si');
                const value = element(strings, 't');
                value.textContent = text;
                if (text.trim() !== text) value.setAttributeNS('http://www.w3.org/XML/1998/namespace', 'xml:space', 'preserve');
                item.append(value);
                strings.documentElement.append(item);
                stringIndices.set(text, stringIndices.size);
            }
            return stringIndices.get(text);
        };
        const columnName = index => {
            let name = '';
            for (let number = index + 1; number; number = Math.floor((number - 1) / 26)) name = String.fromCharCode(65 + (number - 1) % 26) + name;
            return name;
        };
        const sheet = sheetsFor(model)[0];
        const dimensions = dimensionsFor(model);
        const sheetData = worksheet.getElementsByTagName('sheetData')[0];
        const originalRows = [...sheetData.children];
        const output = element(worksheet, 'sheetData');
        sheet.data[0].rowData.forEach((rowData, rowIndex) => {
            const row = (originalRows[rowIndex] || originalRows[32]).cloneNode(false);
            row.setAttribute('r', String(rowIndex + 1));
            row.setAttribute('ht', String(dimensions.heights[rowIndex]));
            row.setAttribute('spans', `1:${dimensions.widths.length}`);
            rowData.values.forEach((cell, column) => {
                const text = cell.userEnteredValue.stringValue;
                const borderless = !cell.userEnteredFormat.backgroundColor;
                if (!text && borderless && ![0, 1, 29].includes(rowIndex)) return;
                const node = element(worksheet, 'c', { r: `${columnName(column)}${rowIndex + 1}`, s: text || !borderless ? styleFor(cell) : 16 });
                if (text) {
                    node.setAttribute('t', 's');
                    const value = element(worksheet, 'v');
                    value.textContent = String(stringFor(text));
                    node.append(value);
                }
                row.append(node);
            });
            output.append(row);
        });
        sheetData.replaceWith(output);
        worksheet.getElementsByTagName('dimension')[0].setAttribute('ref', `A1:${columnName(dimensions.widths.length - 1)}${dimensions.heights.length}`);
        if (dimensions.widths.length !== 6) {
            const columns = element(worksheet, 'cols');
            dimensions.widths.forEach((width, index) => columns.append(element(worksheet, 'col', { min: index + 1, max: index + 1, width, customWidth: 1 })));
            worksheet.getElementsByTagName('cols')[0].replaceWith(columns);
        }
        const ranges = sheet.merges.map(merge => {
            const start = `${columnName(merge.startColumnIndex)}${merge.startRowIndex + 1}`;
            const end = `${columnName(merge.endColumnIndex - 1)}${merge.endRowIndex}`;
            return start === end ? start : `${start}:${end}`;
        });
        dimensions.layouts.forEach((layout, day) => {
            const firstColumn = 1 + dimensions.layouts.slice(0, day).reduce((sum, item) => sum + item.count, 0);
            layout.events.filter(event => event.end - event.start === 30).forEach(event => ranges.push(`${columnName(firstColumn + event.lane)}${5 + (event.start - 480) / 30}`));
        });
        const merges = element(worksheet, 'mergeCells', { count: ranges.length });
        ranges.forEach(ref => merges.append(element(worksheet, 'mergeCell', { ref })));
        worksheet.getElementsByTagName('mergeCells')[0].replaceWith(merges);
        strings.documentElement.setAttribute('count', String(stringCount));
        strings.documentElement.setAttribute('uniqueCount', String(stringIndices.size));
        zip.file('xl/worksheets/sheet1.xml', serialize(worksheet));
        zip.file('xl/sharedStrings.xml', serialize(strings));
        if (fills.children.length !== originalFillCount || formats.children.length !== originalFormatCount) {
            fills.setAttribute('count', String(fills.children.length));
            formats.setAttribute('count', String(formats.children.length));
            zip.file('xl/styles.xml', serialize(styles));
        }
        return zip.generateAsync({ type: 'blob', mimeType: XLSX_MIME, compression: 'DEFLATE' });
    }

    async function api(url, options = {}) {
        const response = await googleFetch(url, options);
        if (!response.ok) {
            const error = await response.json().catch(() => ({}));
            throw new Error(error.error?.message || `שגיאת Google (${response.status})`);
        }
        return response.status === 204 ? {} : response.json();
    }
    const jsonOptions = (method, body) => ({ method, headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
    const driveFile = id => api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=id,name,appProperties,parents`);

    // New files are owned by the Google account so they survive recreating the journal; files owned by the journal id stay valid.
    const ownerValue = () => state.accountId || state.owner;
    const isOwnedByUser = file => Boolean(file.appProperties?.yomanOwner) && [state.owner, state.accountId].includes(file.appProperties.yomanOwner);

    async function loadAccountId() {
        try {
            const about = await api('https://www.googleapis.com/drive/v3/about?fields=user(permissionId)');
            return about.user?.permissionId ? `acct:${about.user.permissionId}` : null;
        } catch { return null; }
    }

    // Duplicates (e.g. two devices creating the same week) resolve to the most recently modified copy.
    async function findFile(week) {
        const query = `name='${title(week)}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
        const result = await api(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,modifiedTime,appProperties,parents)&orderBy=modifiedTime%20desc&pageSize=100`);
        return (result.files || []).find(isOwnedByUser) || null;
    }

    async function readModel(file, week) {
        const result = await api(`https://sheets.googleapis.com/v4/spreadsheets/${file.id}/values/${encodeURIComponent("'_portal_data'!A:A")}`);
        const values = (result.values || []).map(row => row[0]);
        if (values[0] !== 'yoman-gantt-v1' || values[1] !== week) throw new Error('קובץ השבוע אינו בפורמט של הפורטל.');
        return validateModel({ version: 1, week, activities: JSON.parse(values[2]), events: values.slice(3).filter(Boolean).map(value => JSON.parse(value)) }, week);
    }

    async function folder() {
        const query = "name='גאנט למדריכה' and mimeType='application/vnd.google-apps.folder' and trashed=false";
        const result = await api(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,appProperties)&pageSize=100`);
        const existing = result.files.find(isOwnedByUser);
        return existing || api('https://www.googleapis.com/drive/v3/files?fields=id', jsonOptions('POST', { name: 'גאנט למדריכה', mimeType: 'application/vnd.google-apps.folder', appProperties: { yomanOwner: ownerValue() } }));
    }

    function persist() {
        try { localStorage.setItem(draftKey(state.model.week), JSON.stringify({ model: state.model, fileId: state.file?.id || null })); }
        catch { showToast('אין מקום לשמירת טיוטה במכשיר. השאר את הדף פתוח עד לסיום השמירה.', true); }
    }
    function status(message, error = false) {
        const label = document.getElementById('gantt-status');
        if (label) { label.textContent = message; label.classList.toggle('is-error', error); }
    }
    function mutate(change) {
        if (state.busy || !state.model) return;
        change(state.model);
        state.exportVersion++;
        state.exportFile = null;
        state.exportError = false;
        state.dirty = true;
        persist();
        render();
        status('ממתין לשמירה…');
        clearTimeout(state.timer);
        state.timer = setTimeout(() => flush().catch(reportError), 700);
    }
    function reportError(error) { status(error.message, true); showToast(error.message, true); renderButtons(); }

    async function prepareIosDownload() {
        if (!isIOS() || !state.model || !state.file || state.dirty || state.saving || state.exportFile || state.exportPromise) return;
        const version = state.exportVersion;
        const owner = state.owner;
        const model = structuredClone(state.model);
        const promise = (async () => new File([await templateXlsx(model)], `${title(model.week)}.xlsx`, { type: XLSX_MIME }))();
        state.exportPromise = promise;
        state.exportError = false;
        renderButtons();
        try {
            const file = await promise;
            if (state.exportVersion === version && state.owner === owner && state.model?.week === model.week && !state.dirty) state.exportFile = file;
        } catch (error) {
            state.exportError = true;
            throw error;
        } finally {
            if (state.exportPromise === promise) state.exportPromise = null;
            renderButtons();
        }
        if (state.exportVersion !== version && state.file && !state.dirty) queueMicrotask(() => prepareIosDownload().catch(reportError));
    }

    async function flush() {
        clearTimeout(state.timer);
        if (state.saving) { await state.saving; return state.dirty ? flush() : undefined; }
        if (!state.dirty) return;
        if (!accessToken || state.owner !== spreadsheetId) throw new Error('יש להתחבר מחדש כדי לשמור את הלו״ז.');
        const snapshot = structuredClone(state.model);
        state.saving = (async () => {
            status('שומר בדרייב…');
            renderButtons();
            const sheets = sheetsFor(validateModel(snapshot, snapshot.week));
            let createdFile = false;
            if (!state.file) {
                const existing = await findFile(snapshot.week);
                if (existing) state.file = existing;
                else {
                    const parent = await folder();
                    const created = await api('https://sheets.googleapis.com/v4/spreadsheets', jsonOptions('POST', { properties: { title: title(snapshot.week) }, sheets }));
                    state.file = { id: created.spreadsheetId };
                    createdFile = true;
                    persist();
                    await api(`https://sheets.googleapis.com/v4/spreadsheets/${state.file.id}:batchUpdate`, jsonOptions('POST', { requests: dimensionRequests(snapshot) }));
                    state.file = await api(`https://www.googleapis.com/drive/v3/files/${state.file.id}?addParents=${parent.id}&fields=id,name,appProperties,parents`, jsonOptions('PATCH', { appProperties: { yomanOwner: ownerValue(), yomanWeek: snapshot.week } }));
                }
            }
            if (!createdFile) {
                let remote = await driveFile(state.file.id);
                if (!remote.appProperties?.yomanOwner) {
                    const parent = await folder();
                    remote = await api(`https://www.googleapis.com/drive/v3/files/${remote.id}?addParents=${parent.id}&fields=id,name,appProperties,parents`, jsonOptions('PATCH', { appProperties: { yomanOwner: ownerValue(), yomanWeek: snapshot.week } }));
                }
                if (!isOwnedByUser(remote)) throw new Error('קובץ השבוע שייך ליומן אחר.');
                if (state.accountId && remote.appProperties.yomanOwner !== state.accountId) {
                    remote = await api(`https://www.googleapis.com/drive/v3/files/${remote.id}?fields=id,name,appProperties,parents`, jsonOptions('PATCH', { appProperties: { yomanOwner: state.accountId } }));
                }
                const requests = [
                    { unmergeCells: { range: { sheetId: 0 } } },
                    ...sheets.map(sheet => ({ updateSheetProperties: { properties: sheet.properties, fields: 'gridProperties,rightToLeft,hidden' } })),
                    ...sheets.map(sheet => ({ updateCells: { range: { sheetId: sheet.properties.sheetId }, rows: sheet.data[0].rowData, fields: 'userEnteredValue,userEnteredFormat' } })),
                    ...sheets[0].merges.map(range => ({ mergeCells: { range, mergeType: 'MERGE_ALL' } })),
                    ...dimensionRequests(snapshot)
                ];
                await api(`https://sheets.googleapis.com/v4/spreadsheets/${state.file.id}:batchUpdate`, jsonOptions('POST', { requests }));
                state.file = remote;
            }
            state.dirty = JSON.stringify(state.model) !== JSON.stringify(snapshot);
            if (state.dirty) persist();
            else localStorage.removeItem(draftKey(snapshot.week));
            status(state.dirty ? 'ממתין לשמירה…' : 'נשמר בדרייב');
        })();
        try { await state.saving; }
        finally { state.saving = null; renderButtons(); }
        if (state.dirty) return flush();
        prepareIosDownload().catch(reportError);
    }

    async function loadWeek(value) {
        if (state.busy) return;
        const week = sunday(value);
        const oldWeek = state.model?.week;
        state.busy = true;
        renderButtons();
        try {
            await flush();
            status('טוען שבוע…');
            const stored = localStorage.getItem(draftKey(week));
            const draft = stored ? JSON.parse(stored) : null;
            let file = await findFile(week);
            if (!file && draft?.fileId) {
                const recovered = await driveFile(draft.fileId);
                if (recovered.name !== title(week) || (recovered.appProperties?.yomanOwner && !isOwnedByUser(recovered))) throw new Error('לא ניתן לשייך את קובץ הטיוטה לשבוע הזה.');
                file = recovered;
            }
            let model = file ? await readModel(file, week) : blank(week);
            let dirty = false;
            if (draft) {
                model = validateModel(draft.model, week);
                dirty = true;
            }
            state.exportVersion++;
            Object.assign(state, { file, model, dirty, exportFile: null, exportError: false });
            render();
            status(dirty ? 'טיוטה מקומית שטרם נשמרה' : file ? 'נשמר בדרייב' : 'שבוע חדש');
        } catch (error) {
            if (oldWeek) document.getElementById('gantt-week').value = oldWeek;
            reportError(error);
        } finally {
            state.busy = false;
            renderButtons();
            if (state.file && !state.dirty) prepareIosDownload().catch(reportError);
        }
    }

    async function open() {
        if (!accessToken || !spreadsheetId || state.busy) return;
        if (state.owner !== spreadsheetId) reset();
        state.owner = spreadsheetId;
        if (state.model) { render(); return; }
        try {
            if (!state.accountId) state.accountId = await loadAccountId();
            await loadWeek(iso(new Date()));
        } catch (error) { reportError(error); }
    }

    function reset() {
        clearTimeout(state.timer);
        state.exportVersion++;
        Object.assign(state, { owner: null, accountId: null, model: null, file: null, dirty: false, exportFile: null, exportPromise: null, exportError: false });
        document.getElementById('gantt-editor')?.close();
        if (document.getElementById('gantt-board')) document.getElementById('gantt-board').innerHTML = '';
    }

    function renderButtons() {
        const blocked = state.busy || Boolean(state.saving) || !state.model;
        document.querySelectorAll('#view-gantt button, #view-gantt input').forEach(button => { button.disabled = blocked; });
        ['gantt-download', 'gantt-drive'].forEach(id => { document.getElementById(id).disabled = blocked || (!state.file && !state.dirty); });
        const downloadButton = document.getElementById('gantt-download');
        if (isIOS() && state.model && !state.exportFile && !state.exportError) downloadButton.disabled = true;
        document.getElementById('gantt-copy').hidden = Boolean(state.file || state.model?.events.length);
        const retry = document.getElementById('gantt-retry');
        retry.hidden = Boolean(state.model && !state.dirty);
        retry.disabled = state.busy || Boolean(state.saving);
    }

    function render() {
        if (!state.model) return;
        const model = state.model;
        document.getElementById('gantt-week').value = model.week;
        document.getElementById('gantt-range').textContent = `${shortDate(model.week)}–${shortDate(shift(model.week, 4))}`;
        document.getElementById('gantt-days').innerHTML = DAYS.map((day, index) => `<button type="button" data-day="${index}" class="${state.day === index ? 'is-active' : ''}" aria-pressed="${state.day === index}">${day}</button>`).join('');
        const hours = Array.from({ length: 24 }, (_, slot) => `<div class="gantt-time">${time(480 + slot * 30)}</div>`).join('');
        const layouts = DAYS.map((day, index) => lanes(model.events.filter(event => event.day === index)));
        const columns = DAYS.map((day, index) => {
            const layout = layouts[index];
            const slots = Array.from({ length: 24 }, (_, slot) => `<button class="gantt-slot" type="button" data-slot="${slot}" data-day="${index}" aria-label="הוסף פעילות ביום ${day} בשעה ${time(480 + slot * 30)}"></button>`).join('');
            const events = layout.events.map(event => {
                const activity = model.activities.find(item => item.id === event.activity);
                const description = [activity.name, event.place, `${time(event.start)}–${time(event.end)}`, event.text].filter(Boolean).join(' | ');
                return `<div class="gantt-event${event.end - event.start === 30 ? ' is-compact' : ''}" style="--event-color:${activity.color};top:${(event.start - 480) / 30 * 48}px;height:${(event.end - event.start) / 30 * 48}px;inset-inline-start:${event.lane / layout.count * 100}%;width:${100 / layout.count}%" data-event="${html(event.id)}"><button class="gantt-event-edit" type="button" data-edit="${html(event.id)}" title="${html(description)}" aria-label="${html(description)}"><strong>${html(event.place || activity.name)}</strong><span class="gantt-event-time" dir="ltr">${time(event.start)}–${time(event.end)}</span><span class="gantt-event-detail">${html(event.place ? activity.name : event.text)}</span></button><button class="gantt-resize gantt-resize-start" type="button" title="שינוי שעת התחלה" aria-label="שינוי שעת התחלה" data-resize="start" data-event-id="${html(event.id)}"></button><button class="gantt-resize gantt-resize-end" type="button" title="שינוי שעת סיום" aria-label="שינוי שעת סיום" data-resize="end" data-event-id="${html(event.id)}"></button><button class="gantt-move" type="button" title="הזזת פעילות" aria-label="הזזת פעילות" data-move="${html(event.id)}">${icon('grip-vertical')}</button></div>`;
            }).join('');
            return `<div class="gantt-day ${index === state.day ? 'is-selected' : ''}"><div class="gantt-day-head"><strong>${day}</strong><span>${shortDate(shift(model.week, index))}</span></div><div class="gantt-day-body" data-column="${index}">${slots}${events}</div></div>`;
        }).join('');
        const board = document.getElementById('gantt-board');
        const tracks = layouts.map(layout => `minmax(${Math.max(140, layout.count * 115)}px, ${layout.count}fr)`).join(' ');
        board.innerHTML = `<div class="gantt-times"><div class="gantt-day-head">שעות</div>${hours}</div><div class="gantt-week" style="grid-template-columns:${tracks}" tabindex="0" role="region" aria-label="לוח השבוע">${columns}</div>`;
        document.getElementById('gantt-legend').innerHTML = model.activities.map(activity => `<button type="button" class="gantt-legend-item" data-type="${html(activity.id)}" title="עריכת סוג פעילות"><span style="background:${activity.color}"></span>${html(activity.name)}</button>`).join('') + `<button type="button" class="btn-secondary" id="gantt-new-type">${icon('plus')} סוג פעילות</button>`;
        renderButtons();
        redrawIcons();
    }

    function updatePlaceFields(form) {
        const isGardenActivity = form.elements.activity.value === 'type-0';
        const isFreeText = form.elements.gardenSelect.value === '__free_text__';
        document.getElementById('gantt-garden-field').hidden = !isGardenActivity;
        document.getElementById('gantt-free-garden-field').hidden = !isGardenActivity || !isFreeText;
        document.getElementById('gantt-free-garden').required = isGardenActivity && isFreeText;
        document.getElementById('gantt-place-field').hidden = isGardenActivity;
    }

    function editEvent(id, day = state.day, start = 480) {
        if (state.busy || !state.model) return;
        const event = state.model.events.find(item => item.id === id);
        const form = document.getElementById('gantt-form');
        form.reset();
        form.dataset.id = event?.id || '';
        document.getElementById('gantt-editor-title').textContent = event ? 'עריכת פעילות' : 'פעילות חדשה';
        form.elements.day.innerHTML = DAYS.map((name, index) => `<option value="${index}">${name}</option>`).join('');
        form.elements.activity.innerHTML = state.model.activities.map(activity => `<option value="${html(activity.id)}">${html(activity.name)}</option>`).join('');
        form.elements.day.value = event?.day ?? day;
        form.elements.start.value = time(event?.start ?? start);
        form.elements.end.value = time(event?.end ?? Math.min(start + 60, 1200));
        form.elements.activity.value = event?.activity || state.model.activities[0].id;
        form.elements.place.value = event?.place || '';
        form.elements.gardenSelect.innerHTML = appData.kindergartens.map((name, index) => `<option value="garden:${index}">${html(name)}</option>`).join('') + '<option value="__free_text__">טקסט חופשי</option>';
        form.elements.gardenSelect.value = '';
        form.elements.customGarden.value = '';
        if (form.elements.activity.value === 'type-0') {
            const gardenIndex = appData.kindergartens.indexOf(event?.place || '');
            if (gardenIndex >= 0) form.elements.gardenSelect.value = `garden:${gardenIndex}`;
            else if (event?.place || !appData.kindergartens.length) {
                form.elements.gardenSelect.value = '__free_text__';
                form.elements.customGarden.value = event?.place || '';
            } else form.elements.gardenSelect.value = 'garden:0';
        }
        updatePlaceFields(form);
        form.elements.text.value = event?.text || '';
        document.getElementById('gantt-gardens').innerHTML = appData.kindergartens.map(name => `<option value="${html(name)}"></option>`).join('');
        document.getElementById('gantt-delete').hidden = !event;
        document.getElementById('gantt-editor').showModal();
    }

    async function editType(id) {
        const existing = state.model.activities.find(activity => activity.id === id);
        const form = document.getElementById('gantt-type-form');
        form.dataset.id = id || '';
        form.elements.name.value = existing?.name || '';
        form.elements.color.value = existing?.color || '#FADBD8';
        document.getElementById('gantt-type-editor').showModal();
    }

    async function copyPrevious() {
        if (state.busy) return;
        const week = state.model.week;
        state.busy = true;
        renderButtons();
        try {
            const previous = shift(week, -7);
            const file = await findFile(previous);
            if (!file) throw new Error('לא נמצא לו״ז לשבוע הקודם.');
            const model = await readModel(file, previous);
            if (!await openDialog({ title: 'העתקה משבוע קודם', text: `להעתיק את התכנון של ${shortDate(previous)} לשבוע הנבחר?`, confirmLabel: 'העתק' })) return;
            state.busy = false;
            mutate(target => { target.activities = model.activities; target.events = model.events.map(event => ({ ...event, id: crypto.randomUUID() })); });
        } catch (error) { reportError(error); }
        finally { state.busy = false; renderButtons(); }
    }

    async function download() {
        if (state.busy || state.saving) return;
        if (isIOS()) {
            if (!state.exportFile) {
                if (!state.exportPromise) prepareIosDownload().catch(reportError);
                return;
            }
            if (navigator.share && navigator.canShare?.({ files: [state.exportFile] })) {
                const sharing = navigator.share({ files: [state.exportFile] });
                return sharing.then(() => status('הקובץ נשלח לשמירה'), error => {
                    if (error.name !== 'AbortError') reportError(error);
                });
            }
            showToast('שיתוף קובץ אינו נתמך בדפדפן הזה. פתח את הפורטל ב-Safari.', true);
            return;
        }
        try {
            state.dirty = true;
            persist();
            await flush();
            state.busy = true;
            renderButtons();
            const file = new File([await templateXlsx(state.model)], `${title(state.model.week)}.xlsx`, { type: XLSX_MIME });
            const url = URL.createObjectURL(file);
            const link = document.createElement('a');
            link.href = url; link.download = file.name;
            document.body.append(link); link.click(); link.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
        } catch (error) { if (error.name !== 'AbortError') reportError(error); }
        finally { state.busy = false; renderButtons(); }
    }

    function setup() {
        const section = document.getElementById('view-gantt');
        section.innerHTML = `<div class="gantt-toolbar"><div><h2>גאנט למדריכה</h2><div id="gantt-range" class="muted-note"></div></div><div class="gantt-actions"><button id="gantt-drive" class="btn-secondary" type="button" title="פתיחה בדרייב" aria-label="פתיחה בדרייב">${icon('external-link')} פתיחה בדרייב</button><button id="gantt-download" class="btn-secondary" type="button" title="הורדת אקסל" aria-label="הורדת אקסל">${icon('download')} הורדה</button></div></div><div class="gantt-weekbar"><button id="gantt-prev" class="icon-btn" title="השבוע הקודם" aria-label="השבוע הקודם">${icon('chevron-right')}</button><label class="field"><span>שבוע שמתחיל ביום ראשון</span><input id="gantt-week" type="date"></label><button id="gantt-next" class="icon-btn" title="השבוע הבא" aria-label="השבוע הבא">${icon('chevron-left')}</button><button id="gantt-today" class="btn-secondary" type="button">השבוע</button></div><div class="gantt-controls"><button id="gantt-add" class="btn" type="button">${icon('plus')} פעילות</button><button id="gantt-copy" class="btn-secondary" type="button">${icon('copy')} העתק משבוע קודם</button><span id="gantt-status" role="status" aria-live="polite">בחר שבוע</span><button id="gantt-retry" class="btn-secondary" type="button" hidden>${icon('cloud-upload')} שמור שוב</button></div><div id="gantt-days" class="gantt-days" aria-label="יום לתצוגה"></div><div id="gantt-board" class="gantt-board"></div><h3 class="gantt-legend-title">מקרא פעילויות</h3><div id="gantt-legend" class="gantt-legend"></div>`;
        document.body.insertAdjacentHTML('beforeend', `<dialog id="gantt-editor" class="gantt-dialog" dir="rtl"><form id="gantt-form"><div class="sheet-head"><h3 id="gantt-editor-title">פעילות חדשה</h3><button class="icon-btn" type="button" data-close="gantt-editor" title="סגירה" aria-label="סגירה">${icon('x')}</button></div><div class="field-row"><label class="field"><span>יום</span><select name="day"></select></label><label class="field"><span>סוג פעילות</span><select name="activity"></select></label></div><div class="field-row"><label class="field"><span>התחלה</span><input name="start" type="time" min="08:00" max="19:30" step="1800" required></label><label class="field"><span>סיום</span><input name="end" type="time" min="08:30" max="20:00" step="1800" required></label></div><label class="field"><span>גן / מקום</span><input name="place" list="gantt-gardens" maxlength="200"><datalist id="gantt-gardens"></datalist></label><label class="field gantt-text"><span>תוכן הפעילות</span><textarea name="text" maxlength="2000"></textarea></label><div class="gantt-dialog-actions"><button id="gantt-delete" class="btn-danger" type="button">מחק</button><button class="btn" type="submit">שמור</button></div></form></dialog><dialog id="gantt-type-editor" class="gantt-dialog" dir="rtl"><form id="gantt-type-form"><div class="sheet-head"><h3>סוג פעילות</h3><button class="icon-btn" type="button" data-close="gantt-type-editor" title="סגירה" aria-label="סגירה">${icon('x')}</button></div><label class="field"><span>שם</span><input name="name" maxlength="150" required></label><label class="field gantt-color"><span>צבע</span><input name="color" type="color" required></label><div class="gantt-dialog-actions"><button class="btn" type="submit">שמור</button></div></form></dialog>`);
        const eventForm = document.getElementById('gantt-form');
        const originalPlaceField = eventForm.elements.place.closest('.field');
        originalPlaceField.id = 'gantt-place-field';
        const gardenField = document.createElement('label');
        gardenField.id = 'gantt-garden-field';
        gardenField.className = 'field';
        gardenField.innerHTML = '<span>גן</span><select name="gardenSelect"></select>';
        const freeGardenField = document.createElement('label');
        freeGardenField.id = 'gantt-free-garden-field';
        freeGardenField.className = 'field';
        freeGardenField.innerHTML = '<span>שם הגן</span><input id="gantt-free-garden" name="customGarden" type="text" maxlength="200">';
        originalPlaceField.before(gardenField, freeGardenField);
        eventForm.elements.activity.addEventListener('change', () => {
            if (eventForm.elements.activity.value === 'type-0' && !eventForm.elements.gardenSelect.value) {
                eventForm.elements.gardenSelect.value = appData.kindergartens.length ? 'garden:0' : '__free_text__';
            }
            updatePlaceFields(eventForm);
        });
        eventForm.elements.gardenSelect.addEventListener('change', () => updatePlaceFields(eventForm));
        const safe = action => () => Promise.resolve().then(action).catch(reportError);
        document.getElementById('gantt-week').addEventListener('change', event => { if (event.target.value) loadWeek(event.target.value); });
        document.getElementById('gantt-prev').onclick = () => loadWeek(shift(state.model.week, -7));
        document.getElementById('gantt-next').onclick = () => loadWeek(shift(state.model.week, 7));
        document.getElementById('gantt-today').onclick = () => loadWeek(iso(new Date()));
        document.getElementById('gantt-add').onclick = () => editEvent();
        document.getElementById('gantt-copy').onclick = safe(copyPrevious);
        document.getElementById('gantt-retry').onclick = safe(() => state.model ? flush() : open());
        document.getElementById('gantt-download').onclick = () => { download().catch(reportError); };
        document.getElementById('gantt-drive').onclick = safe(async () => {
            await flush();
            const fileId = state.file.id;
            if (await openDialog({ title: 'פתיחה בדרייב', text: 'העריכה נשמרת דרך הפורטל. שינויים ישירות בשיטס אינם מיובאים אוטומטית ועלולים להתנגש עם התכנון בפורטל.', confirmLabel: 'פתח' })) window.open(`https://docs.google.com/spreadsheets/d/${fileId}/edit`, '_blank', 'noopener');
        });
        section.addEventListener('click', event => {
            const button = event.target.closest('button');
            if (!button || button.disabled || state.busy || !state.model) return;
            if (button.dataset.edit) editEvent(button.dataset.edit);
            else if (button.dataset.slot !== undefined) editEvent(null, Number(button.dataset.day), 480 + Number(button.dataset.slot) * 30);
            else if (button.closest('#gantt-days')) { state.day = Number(button.dataset.day); render(); }
            else if (button.dataset.type) editType(button.dataset.type);
            else if (button.id === 'gantt-new-type') editType();
        });
        document.querySelectorAll('[data-close]').forEach(button => { button.onclick = () => document.getElementById(button.dataset.close).close(); });
        document.getElementById('gantt-form').addEventListener('submit', event => {
            event.preventDefault();
            const form = event.currentTarget;
            const start = minutes(form.elements.start.value), end = minutes(form.elements.end.value);
            if (start >= end || start < 480 || end > 1200 || start % 30 || end % 30) { showToast('בחר שעות בחצאי שעה בין 08:00 ל־20:00, עם סיום אחרי ההתחלה.', true); return; }
            const activity = form.elements.activity.value;
            let place = form.elements.place.value.trim();
            if (activity === 'type-0') {
                const gardenChoice = form.elements.gardenSelect.value;
                place = gardenChoice === '__free_text__' ? form.elements.customGarden.value.trim() : appData.kindergartens[Number(gardenChoice.slice(7))] || '';
                if (!place) { showToast('נא לבחור גן או להזין שם גן.', true); return; }
            }
            const saved = { id: form.dataset.id || crypto.randomUUID(), day: Number(form.elements.day.value), start, end, activity, place, text: form.elements.text.value.trim() };
            mutate(model => { const index = model.events.findIndex(item => item.id === saved.id); if (index < 0) model.events.push(saved); else model.events[index] = saved; });
            document.getElementById('gantt-editor').close();
        });
        document.getElementById('gantt-delete').onclick = async () => {
            const id = document.getElementById('gantt-form').dataset.id;
            if (!confirm('למחוק את הפעילות מהתכנון השבועי?')) return;
            mutate(model => { model.events = model.events.filter(event => event.id !== id); });
            document.getElementById('gantt-editor').close();
        };
        document.getElementById('gantt-type-form').addEventListener('submit', event => {
            event.preventDefault();
            const form = event.currentTarget;
            const name = form.elements.name.value.trim();
            if (!name) return;
            const activity = { id: form.dataset.id || crypto.randomUUID(), name, color: form.elements.color.value };
            mutate(model => { const index = model.activities.findIndex(item => item.id === activity.id); if (index < 0) model.activities.push(activity); else model.activities[index] = activity; });
            document.getElementById('gantt-type-editor').close();
        });
        let dragging = null;
        section.addEventListener('pointerdown', event => {
            const resizeHandle = event.target.closest('[data-resize]');
            if (resizeHandle && !state.busy && !state.saving) {
                const modelEvent = state.model.events.find(item => item.id === resizeHandle.dataset.eventId);
                if (!modelEvent) return;
                dragging = { mode: 'resize', edge: resizeHandle.dataset.resize, event: modelEvent, element: resizeHandle.closest('.gantt-event'), y: event.clientY, pointerId: event.pointerId };
                dragging.element.classList.add('is-resizing');
                resizeHandle.setPointerCapture(event.pointerId);
                event.preventDefault();
                return;
            }
            const handle = event.target.closest('[data-move]');
            if (!handle || state.busy || state.saving) return;
            const modelEvent = state.model.events.find(item => item.id === handle.dataset.move);
            if (!modelEvent) return;
            dragging = { mode: 'move', event: modelEvent, element: handle.closest('.gantt-event'), x: event.clientX, y: event.clientY, pointerId: event.pointerId };
            handle.setPointerCapture(event.pointerId);
            event.preventDefault();
        });
        section.addEventListener('pointermove', event => {
            if (!dragging || event.pointerId !== dragging.pointerId) return;
            if (dragging.mode === 'move') {
                dragging.element.style.transform = `translate(${event.clientX - dragging.x}px,${event.clientY - dragging.y}px)`;
                dragging.element.classList.add('is-dragging');
                return;
            }
            const delta = Math.round((event.clientY - dragging.y) / 48) * 30;
            const start = dragging.edge === 'start'
                ? Math.max(480, Math.min(dragging.event.end - 30, dragging.event.start + delta))
                : dragging.event.start;
            const end = dragging.edge === 'end'
                ? Math.min(1200, Math.max(dragging.event.start + 30, dragging.event.end + delta))
                : dragging.event.end;
            dragging.preview = { start, end };
            dragging.element.style.top = `${(start - 480) / 30 * 48}px`;
            dragging.element.style.height = `${(end - start) / 30 * 48}px`;
            dragging.element.classList.add('is-resizing');
            dragging.element.querySelector('.gantt-event-time').textContent = `${time(start)}–${time(end)}`;
        });
        const drop = event => {
            if (!dragging || event.pointerId !== dragging.pointerId) return;
            const current = dragging; dragging = null;
            current.element.style.transform = '';
            current.element.style.top = '';
            current.element.style.height = '';
            current.element.classList.remove('is-dragging', 'is-resizing');
            if (event.type === 'pointercancel') {
                current.element.querySelector('.gantt-event-time').textContent = `${time(current.event.start)}–${time(current.event.end)}`;
                return;
            }
            if (current.mode === 'resize') {
                if (current.preview && (current.preview.start !== current.event.start || current.preview.end !== current.event.end)) {
                    mutate(model => {
                        const resized = model.events.find(item => item.id === current.event.id);
                        if (resized) Object.assign(resized, current.preview);
                    });
                } else {
                    current.element.querySelector('.gantt-event-time').textContent = `${time(current.event.start)}–${time(current.event.end)}`;
                }
                return;
            }
            const target = [...document.querySelectorAll('[data-column]')].find(column => { const rect = column.getBoundingClientRect(); return rect.width && event.clientX >= rect.left && event.clientX <= rect.right; });
            if (!target || Math.abs(event.clientY - current.y) + Math.abs(event.clientX - current.x) < 8) return;
            const duration = current.event.end - current.event.start;
            const start = Math.max(480, Math.min(1200 - duration, current.event.start + Math.round((event.clientY - current.y) / 48) * 30));
            mutate(model => { const moved = model.events.find(item => item.id === current.event.id); Object.assign(moved, { day: Number(target.dataset.column), start, end: start + duration }); });
        };
        section.addEventListener('pointerup', drop);
        section.addEventListener('pointercancel', drop);
        window.addEventListener('beforeunload', event => { if (state.dirty || state.saving) { event.preventDefault(); event.returnValue = ''; } });
        renderButtons();
        redrawIcons();
    }

    globalThis.GuideGantt = { open, reset, setup, isSaving: () => Boolean(state.saving || state.busy) };
    if (document.getElementById('view-gantt')) setup();
})();