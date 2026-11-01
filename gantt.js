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
    const state = { owner: null, model: null, file: null, dirty: false, busy: false, timer: null, saving: null, day: new Date().getDay() % 5, conflict: false };
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

    function sheetsFor(model) {
        const layouts = DAYS.map((day, index) => lanes(model.events.filter(event => event.day === index)));
        const columns = 1 + layouts.reduce((sum, layout) => sum + layout.count, 0);
        const cell = (text = '', color = '#FFFFFF', bold = false) => ({ userEnteredValue: { stringValue: String(text) }, userEnteredFormat: { backgroundColor: { red: parseInt(color.slice(1, 3), 16) / 255, green: parseInt(color.slice(3, 5), 16) / 255, blue: parseInt(color.slice(5, 7), 16) / 255 }, horizontalAlignment: 'CENTER', verticalAlignment: 'MIDDLE', wrapStrategy: 'WRAP', textFormat: { fontFamily: 'Calibri', fontSize: 11, bold }, borders: Object.fromEntries(['top', 'bottom', 'left', 'right'].map(side => [side, { style: 'SOLID', color: { red: .83, green: .83, blue: .83 } }])) } });
        const rows = Array.from({ length: 29 + model.activities.length }, () => ({ values: Array.from({ length: columns }, () => cell()) }));
        rows[0].values[1] = cell('מערכת שעות שבועית', '#FFFFFF', true);
        rows[1].values[1] = cell(`מלווה גני ילדים - תכנון שבועי מלא | ${shortDate(model.week)}–${shortDate(shift(model.week, 4))}`);
        rows[3].values[0] = cell('שעות', '#C9D0F8', true);
        const merges = [{ sheetId: 0, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 1, endColumnIndex: columns }, { sheetId: 0, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 1, endColumnIndex: columns }];
        for (let slot = 0; slot < 24; slot++) rows[slot + 4].values[0] = cell(`${time(480 + slot * 30)}–${time(510 + slot * 30)}`, '#FFFFFF', true);
        let column = 1;
        layouts.forEach((layout, day) => {
            for (let offset = 0; offset < layout.count; offset++) rows[3].values[column + offset] = cell(offset ? '' : `${DAYS[day]}\n${shortDate(shift(model.week, day))}`, '#C9D0F8', true);
            if (layout.count > 1) merges.push({ sheetId: 0, startRowIndex: 3, endRowIndex: 4, startColumnIndex: column, endColumnIndex: column + layout.count });
            layout.events.forEach(event => {
                const activity = model.activities.find(item => item.id === event.activity);
                const startRow = 4 + (event.start - 480) / 30;
                const endRow = 4 + (event.end - 480) / 30;
                for (let row = startRow; row < endRow; row++) rows[row].values[column + event.lane] = cell(row === startRow ? [activity.name, event.place, event.text].filter(Boolean).join('\n') : '', activity.color, true);
                if (endRow - startRow > 1) merges.push({ sheetId: 0, startRowIndex: startRow, endRowIndex: endRow, startColumnIndex: column + event.lane, endColumnIndex: column + event.lane + 1 });
            });
            column += layout.count;
        });
        rows[28].values[0] = cell('מקרא הצבעים', '#FFFFFF', true);
        model.activities.forEach((activity, index) => {
            rows[29 + index].values[1] = cell(activity.name, activity.color, true);
            merges.push({ sheetId: 0, startRowIndex: 29 + index, endRowIndex: 30 + index, startColumnIndex: 1, endColumnIndex: columns });
        });
        const dataRows = ['yoman-gantt-v1', model.week, JSON.stringify(model.activities), ...model.events.map(event => JSON.stringify(event))].map(value => ({ values: [cell(value)] }));
        if (dataRows.some(row => row.values[0].userEnteredValue.stringValue.length > 45000)) throw new Error('הנתונים ארוכים מדי לשמירה.');
        return [
            { properties: { sheetId: 0, title: 'מערכת שעות שבועית', rightToLeft: true, gridProperties: { rowCount: rows.length, columnCount: columns, frozenRowCount: 4, frozenColumnCount: 1 } }, data: [{ startRow: 0, startColumn: 0, rowData: rows }], merges },
            { properties: { sheetId: 1, title: '_portal_data', hidden: true, gridProperties: { rowCount: Math.max(100, dataRows.length), columnCount: 1 } }, data: [{ startRow: 0, startColumn: 0, rowData: dataRows }] }
        ];
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
    const driveFile = id => api(`https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?fields=id,name,modifiedTime,appProperties,parents`);

    async function findFile(week) {
        const query = `name='${title(week)}' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`;
        const result = await api(`https://www.googleapis.com/drive/v3/files?q=${encodeURIComponent(query)}&fields=files(id,name,modifiedTime,appProperties,parents)&pageSize=100`);
        const matches = (result.files || []).filter(file => file.appProperties?.yomanOwner === state.owner);
        if (matches.length > 1) throw new Error('נמצאו כמה קבצים לשבוע הזה. יש להסיר את העותק המיותר בדרייב.');
        return matches[0] || null;
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
        const existing = result.files.find(file => file.appProperties?.yomanOwner === state.owner);
        return existing || api('https://www.googleapis.com/drive/v3/files?fields=id', jsonOptions('POST', { name: 'גאנט למדריכה', mimeType: 'application/vnd.google-apps.folder', appProperties: { yomanOwner: state.owner } }));
    }

    function persist() {
        try { localStorage.setItem(draftKey(state.model.week), JSON.stringify({ model: state.model, fileId: state.file?.id || null, modifiedTime: state.file?.modifiedTime || null })); }
        catch { showToast('אין מקום לשמירת טיוטה במכשיר. השאר את הדף פתוח עד לסיום השמירה.', true); }
    }
    function status(message, error = false) {
        const label = document.getElementById('gantt-status');
        if (label) { label.textContent = message; label.classList.toggle('is-error', error); }
    }
    function mutate(change) {
        if (state.busy || !state.model) return;
        change(state.model);
        state.dirty = true;
        persist();
        render();
        status('ממתין לשמירה…');
        clearTimeout(state.timer);
        state.timer = setTimeout(() => flush().catch(reportError), 700);
    }
    function reportError(error) { status(error.message, true); showToast(error.message, true); renderButtons(); }

    async function flush() {
        clearTimeout(state.timer);
        if (state.saving) { await state.saving; return state.dirty ? flush() : undefined; }
        if (!state.dirty) return;
        if (!accessToken || state.owner !== spreadsheetId) throw new Error('יש להתחבר מחדש כדי לשמור את הלו״ז.');
        if (state.conflict) throw new Error('הקובץ השתנה מחוץ לפורטל. לחץ על פתרון התנגשות לפני שמירה.');
        const snapshot = structuredClone(state.model);
        state.saving = (async () => {
            status('שומר בדרייב…');
            renderButtons();
            const sheets = sheetsFor(validateModel(snapshot, snapshot.week));
            if (!state.file) {
                const existing = await findFile(snapshot.week);
                if (existing) { state.file = existing; state.conflict = true; throw new Error('השבוע נוצר בחלון אחר. לחץ על פתרון התנגשות.'); }
                const parent = await folder();
                const created = await api('https://sheets.googleapis.com/v4/spreadsheets', jsonOptions('POST', { properties: { title: title(snapshot.week) }, sheets }));
                state.file = { id: created.spreadsheetId };
                persist();
                await api(`https://sheets.googleapis.com/v4/spreadsheets/${state.file.id}:batchUpdate`, jsonOptions('POST', { requests: [
                    { updateDimensionProperties: { range: { sheetId: 0, dimension: 'ROWS', startIndex: 4, endIndex: 28 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 0, endIndex: sheets[0].properties.gridProperties.columnCount }, properties: { pixelSize: 160 }, fields: 'pixelSize' } }
                ] }));
                state.file = await api(`https://www.googleapis.com/drive/v3/files/${state.file.id}?addParents=${parent.id}&fields=id,name,modifiedTime,appProperties,parents`, jsonOptions('PATCH', { appProperties: { yomanOwner: state.owner, yomanWeek: snapshot.week } }));
            } else {
                let remote = await driveFile(state.file.id);
                if (state.file.modifiedTime && remote.modifiedTime !== state.file.modifiedTime) { state.conflict = true; throw new Error('הקובץ השתנה מחוץ לפורטל. לחץ על פתרון התנגשות.'); }
                if (!remote.appProperties?.yomanOwner) {
                    const parent = await folder();
                    remote = await api(`https://www.googleapis.com/drive/v3/files/${remote.id}?addParents=${parent.id}&fields=id,name,modifiedTime,appProperties,parents`, jsonOptions('PATCH', { appProperties: { yomanOwner: state.owner, yomanWeek: snapshot.week } }));
                }
                if (remote.appProperties.yomanOwner !== state.owner) throw new Error('קובץ השבוע שייך ליומן אחר.');
                const requests = [
                    { unmergeCells: { range: { sheetId: 0 } } },
                    ...sheets.map(sheet => ({ updateSheetProperties: { properties: sheet.properties, fields: 'gridProperties,rightToLeft,hidden' } })),
                    ...sheets.map(sheet => ({ updateCells: { range: { sheetId: sheet.properties.sheetId }, rows: sheet.data[0].rowData, fields: 'userEnteredValue,userEnteredFormat' } })),
                    ...sheets[0].merges.map(range => ({ mergeCells: { range, mergeType: 'MERGE_ALL' } })),
                    { updateDimensionProperties: { range: { sheetId: 0, dimension: 'ROWS', startIndex: 4, endIndex: 28 }, properties: { pixelSize: 40 }, fields: 'pixelSize' } },
                    { updateDimensionProperties: { range: { sheetId: 0, dimension: 'COLUMNS', startIndex: 0, endIndex: sheets[0].properties.gridProperties.columnCount }, properties: { pixelSize: 160 }, fields: 'pixelSize' } }
                ];
                await api(`https://sheets.googleapis.com/v4/spreadsheets/${state.file.id}:batchUpdate`, jsonOptions('POST', { requests }));
                state.file = await driveFile(state.file.id);
            }
            state.dirty = JSON.stringify(state.model) !== JSON.stringify(snapshot);
            if (state.dirty) persist();
            else localStorage.removeItem(draftKey(snapshot.week));
            status(state.dirty ? 'ממתין לשמירה…' : 'נשמר בדרייב');
        })();
        try { await state.saving; }
        finally { state.saving = null; renderButtons(); }
        if (state.dirty) return flush();
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
                if (recovered.name !== title(week) || (recovered.appProperties?.yomanOwner && recovered.appProperties.yomanOwner !== state.owner)) throw new Error('לא ניתן לשייך את קובץ הטיוטה לשבוע הזה.');
                file = recovered;
            }
            let model = file ? await readModel(file, week) : blank(week);
            let dirty = false;
            let conflict = false;
            if (draft) {
                model = validateModel(draft.model, week);
                dirty = true;
                conflict = Boolean(file && (draft.fileId !== file.id || (draft.modifiedTime && draft.modifiedTime !== file.modifiedTime)));
            }
            Object.assign(state, { file, model, dirty, conflict });
            render();
            status(conflict ? 'טיוטה מקומית ושינוי בדרייב — נדרש פתרון התנגשות' : dirty ? 'טיוטה מקומית שטרם נשמרה' : file ? 'נשמר בדרייב' : 'שבוע חדש');
        } catch (error) {
            if (oldWeek) document.getElementById('gantt-week').value = oldWeek;
            reportError(error);
        } finally { state.busy = false; renderButtons(); }
    }

    async function open() {
        if (!accessToken || !spreadsheetId || state.busy) return;
        if (state.owner !== spreadsheetId) reset();
        state.owner = spreadsheetId;
        if (state.model) { render(); return; }
        try {
            await loadWeek(iso(new Date()));
        } catch (error) { reportError(error); }
    }

    function reset() {
        clearTimeout(state.timer);
        Object.assign(state, { owner: null, model: null, file: null, dirty: false, conflict: false });
        document.getElementById('gantt-editor')?.close();
        if (document.getElementById('gantt-board')) document.getElementById('gantt-board').innerHTML = '';
    }

    function renderButtons() {
        const blocked = state.busy || Boolean(state.saving) || !state.model;
        document.querySelectorAll('#view-gantt button, #view-gantt input').forEach(button => { button.disabled = blocked; });
        ['gantt-download', 'gantt-drive'].forEach(id => { document.getElementById(id).disabled = blocked || (!state.file && !state.dirty); });
        document.getElementById('gantt-copy').hidden = Boolean(state.file || state.model?.events.length);
        document.getElementById('gantt-conflict').hidden = !state.conflict;
        const retry = document.getElementById('gantt-retry');
        retry.hidden = Boolean(state.model && (!state.dirty || state.conflict));
        retry.disabled = state.busy || Boolean(state.saving);
    }

    function render() {
        if (!state.model) return;
        const model = state.model;
        document.getElementById('gantt-week').value = model.week;
        document.getElementById('gantt-range').textContent = `${shortDate(model.week)}–${shortDate(shift(model.week, 4))}`;
        document.getElementById('gantt-days').innerHTML = DAYS.map((day, index) => `<button type="button" data-day="${index}" class="${state.day === index ? 'is-active' : ''}" aria-pressed="${state.day === index}">${day}</button>`).join('');
        const hours = Array.from({ length: 24 }, (_, slot) => `<div class="gantt-time">${time(480 + slot * 30)}</div>`).join('');
        const columns = DAYS.map((day, index) => {
            const layout = lanes(model.events.filter(event => event.day === index));
            const slots = Array.from({ length: 24 }, (_, slot) => `<button class="gantt-slot" type="button" data-slot="${slot}" data-day="${index}" aria-label="הוסף פעילות ביום ${day} בשעה ${time(480 + slot * 30)}"></button>`).join('');
            const events = layout.events.map(event => {
                const activity = model.activities.find(item => item.id === event.activity);
                return `<div class="gantt-event" style="--event-color:${activity.color};top:${(event.start - 480) / 30 * 48}px;height:${(event.end - event.start) / 30 * 48}px;inset-inline-start:${event.lane / layout.count * 100}%;width:${100 / layout.count}%" data-event="${html(event.id)}"><button class="gantt-event-edit" type="button" data-edit="${html(event.id)}" title="עריכת פעילות"><strong>${html(event.place || activity.name)}</strong><span>${time(event.start)}–${time(event.end)}</span><span>${html(event.place ? activity.name : event.text)}</span></button><button class="gantt-move" type="button" title="הזזת פעילות" aria-label="הזזת פעילות" data-move="${html(event.id)}">${icon('grip-vertical')}</button></div>`;
            }).join('');
            return `<div class="gantt-day ${index === state.day ? 'is-selected' : ''}"><div class="gantt-day-head"><strong>${day}</strong><span>${shortDate(shift(model.week, index))}</span></div><div class="gantt-day-body" data-column="${index}">${slots}${events}</div></div>`;
        }).join('');
        document.getElementById('gantt-board').innerHTML = `<div class="gantt-times"><div class="gantt-day-head">שעות</div>${hours}</div>${columns}`;
        document.getElementById('gantt-legend').innerHTML = model.activities.map(activity => `<button type="button" class="gantt-legend-item" data-type="${html(activity.id)}" title="עריכת סוג פעילות"><span style="background:${activity.color}"></span>${html(activity.name)}</button>`).join('') + `<button type="button" class="btn-secondary" id="gantt-new-type">${icon('plus')} סוג פעילות</button>`;
        renderButtons();
        redrawIcons();
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

    async function resolveConflict() {
        const overwrite = await openDialog({ title: 'שינוי מחוץ לפורטל', text: 'שמירת הטיוטה מהפורטל תחליף את התכנון והעיצוב הנוכחיים בשיטס. בחר ״שמור את הטיוטה״ להחלפה, או ״ביטול״ כדי להשאיר את הטיוטה ללא שינוי.', confirmLabel: 'שמור את הטיוטה', danger: true });
        if (!overwrite) return;
        try { state.file = await driveFile(state.file.id); state.conflict = false; await flush(); }
        catch (error) { reportError(error); }
    }

    async function download() {
        try {
            await flush();
            state.busy = true;
            renderButtons();
            const response = await googleFetch(`https://www.googleapis.com/drive/v3/files/${state.file.id}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}`);
            if (!response.ok) throw new Error('הורדת האקסל נכשלה. נסה שוב.');
            const url = URL.createObjectURL(await response.blob());
            const link = document.createElement('a');
            link.href = url; link.download = `${title(state.model.week)}.xlsx`;
            document.body.append(link); link.click(); link.remove();
            setTimeout(() => URL.revokeObjectURL(url), 60000);
        } catch (error) { reportError(error); }
        finally { state.busy = false; renderButtons(); }
    }

    function setup() {
        const section = document.getElementById('view-gantt');
        section.innerHTML = `<div class="gantt-toolbar"><div><h2>גאנט למדריכה</h2><div id="gantt-range" class="muted-note"></div></div><div class="gantt-actions"><button id="gantt-drive" class="btn-secondary" type="button" title="פתיחה בדרייב" aria-label="פתיחה בדרייב">${icon('external-link')} פתיחה בדרייב</button><button id="gantt-download" class="btn-secondary" type="button" title="הורדת אקסל" aria-label="הורדת אקסל">${icon('download')} הורדה</button></div></div><div class="gantt-weekbar"><button id="gantt-prev" class="icon-btn" title="השבוע הקודם" aria-label="השבוע הקודם">${icon('chevron-right')}</button><label class="field"><span>שבוע שמתחיל ביום ראשון</span><input id="gantt-week" type="date"></label><button id="gantt-next" class="icon-btn" title="השבוע הבא" aria-label="השבוע הבא">${icon('chevron-left')}</button><button id="gantt-today" class="btn-secondary" type="button">השבוע</button></div><div class="gantt-controls"><button id="gantt-add" class="btn" type="button">${icon('plus')} פעילות</button><button id="gantt-copy" class="btn-secondary" type="button">${icon('copy')} העתק משבוע קודם</button><span id="gantt-status" role="status" aria-live="polite">בחר שבוע</span><button id="gantt-retry" class="btn-secondary" type="button" hidden>${icon('cloud-upload')} שמור שוב</button><button id="gantt-conflict" class="btn-secondary" type="button" hidden>פתרון התנגשות</button></div><div id="gantt-days" class="gantt-days" aria-label="יום לתצוגה"></div><div id="gantt-board" class="gantt-board"></div><h3 class="gantt-legend-title">מקרא פעילויות</h3><div id="gantt-legend" class="gantt-legend"></div>`;
        document.body.insertAdjacentHTML('beforeend', `<dialog id="gantt-editor" class="gantt-dialog" dir="rtl"><form id="gantt-form"><div class="sheet-head"><h3 id="gantt-editor-title">פעילות חדשה</h3><button class="icon-btn" type="button" data-close="gantt-editor" title="סגירה" aria-label="סגירה">${icon('x')}</button></div><div class="field-row"><label class="field"><span>יום</span><select name="day"></select></label><label class="field"><span>סוג פעילות</span><select name="activity"></select></label></div><div class="field-row"><label class="field"><span>התחלה</span><input name="start" type="time" min="08:00" max="19:30" step="1800" required></label><label class="field"><span>סיום</span><input name="end" type="time" min="08:30" max="20:00" step="1800" required></label></div><label class="field"><span>גן / מקום</span><input name="place" list="gantt-gardens" maxlength="200"><datalist id="gantt-gardens"></datalist></label><label class="field gantt-text"><span>תוכן הפעילות</span><textarea name="text" maxlength="2000"></textarea></label><div class="gantt-dialog-actions"><button id="gantt-delete" class="btn-danger" type="button">מחק</button><button class="btn" type="submit">שמור</button></div></form></dialog><dialog id="gantt-type-editor" class="gantt-dialog" dir="rtl"><form id="gantt-type-form"><div class="sheet-head"><h3>סוג פעילות</h3><button class="icon-btn" type="button" data-close="gantt-type-editor" title="סגירה" aria-label="סגירה">${icon('x')}</button></div><label class="field"><span>שם</span><input name="name" maxlength="150" required></label><label class="field gantt-color"><span>צבע</span><input name="color" type="color" required></label><div class="gantt-dialog-actions"><button class="btn" type="submit">שמור</button></div></form></dialog>`);
        const safe = action => () => Promise.resolve().then(action).catch(reportError);
        document.getElementById('gantt-week').addEventListener('change', event => { if (event.target.value) loadWeek(event.target.value); });
        document.getElementById('gantt-prev').onclick = () => loadWeek(shift(state.model.week, -7));
        document.getElementById('gantt-next').onclick = () => loadWeek(shift(state.model.week, 7));
        document.getElementById('gantt-today').onclick = () => loadWeek(iso(new Date()));
        document.getElementById('gantt-add').onclick = () => editEvent();
        document.getElementById('gantt-copy').onclick = safe(copyPrevious);
        document.getElementById('gantt-retry').onclick = safe(() => state.model ? flush() : open());
        document.getElementById('gantt-conflict').onclick = safe(resolveConflict);
        document.getElementById('gantt-download').onclick = safe(download);
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
            const saved = { id: form.dataset.id || crypto.randomUUID(), day: Number(form.elements.day.value), start, end, activity: form.elements.activity.value, place: form.elements.place.value.trim(), text: form.elements.text.value.trim() };
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
            const handle = event.target.closest('[data-move]');
            if (!handle || state.busy || state.saving) return;
            const modelEvent = state.model.events.find(item => item.id === handle.dataset.move);
            dragging = { event: modelEvent, element: handle.closest('.gantt-event'), x: event.clientX, y: event.clientY, pointerId: event.pointerId };
            handle.setPointerCapture(event.pointerId);
            event.preventDefault();
        });
        section.addEventListener('pointermove', event => {
            if (!dragging || event.pointerId !== dragging.pointerId) return;
            dragging.element.style.transform = `translate(${event.clientX - dragging.x}px,${event.clientY - dragging.y}px)`;
            dragging.element.classList.add('is-dragging');
        });
        const drop = event => {
            if (!dragging || event.pointerId !== dragging.pointerId) return;
            const current = dragging; dragging = null;
            current.element.style.transform = ''; current.element.classList.remove('is-dragging');
            if (event.type === 'pointercancel') return;
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