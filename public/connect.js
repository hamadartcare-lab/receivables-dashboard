const status = document.getElementById('sync-status');
const choose = document.getElementById('choose-workbook');
const upload = document.getElementById('workbook-input');
let handle, remembered, signature, timer, selection = 0;

function workbookStore(mode, action) {
  return new Promise((resolve, reject) => {
    const open = indexedDB.open('receivables-dashboard', 1);
    open.onupgradeneeded = () => open.result.createObjectStore('settings');
    open.onerror = () => reject(open.error);
    open.onsuccess = () => {
      const db = open.result;
      const tx = db.transaction('settings', mode);
      const request = action(tx.objectStore('settings'));
      tx.oncomplete = () => { db.close(); resolve(request.result); };
      tx.onabort = () => { db.close(); reject(tx.error); };
    };
  });
}

async function refresh(id, file) {
  try {
    file ??= await handle.getFile();
    const marker = file.lastModified + ':' + file.size;
    if (marker !== signature) {
      const data = parseWorkbook(await file.arrayBuffer(), file.lastModified);
      if (id !== selection) return;
      renderDashboard(data);
      for (const field of ['p-search','n-search','m-search']) {
        const input = document.getElementById(field);
        if (input.value) input.oninput();
      }
      signature = marker;
      status.textContent = 'آخر حفظ للملف ' + data.asOf;
    }
  } catch (error) {
    if (id === selection) {
      signature = undefined;
      if (error.name === 'NotAllowedError' || error.name === 'NotFoundError') {
        if (error.name === 'NotFoundError') {
          remembered = undefined;
          workbookStore('readwrite', store => store.delete('workbook')).catch(() => {});
        }
        handle = undefined;
        choose.textContent = remembered ? 'السماح بقراءة الملف المحفوظ' : 'ربط ملف الإكسل';
      }
      status.textContent = 'تعذرت قراءة الملف · ' + error.message;
    }
  } finally {
    if (id === selection && handle) timer = setTimeout(() => refresh(id), 2000);
  }
}

function connect(nextHandle, file) {
  clearTimeout(timer);
  handle = nextHandle;
  signature = undefined;
  choose.textContent = 'تغيير ملف الإكسل';
  refresh(++selection, file);
}

choose.onclick = async () => {
  if (!window.showOpenFilePicker) return upload.click();
  try {
    if (remembered && !handle) {
      if (await remembered.requestPermission({mode: 'read'}) === 'granted') {
        connect(remembered);
        return;
      }
      status.textContent = 'المتصفح لم يسمح بقراءة الملف المحفوظ';
      return;
    }
    const [file] = await window.showOpenFilePicker({
      types: [{description: 'Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]
    });
    remembered = file;
    connect(file);
    workbookStore('readwrite', store => store.put(file, 'workbook')).catch(() => {});
  } catch (error) {
    if (error.name === 'SecurityError') upload.click();
    else if (error.name !== 'AbortError') status.textContent = 'تعذر فتح الملف';
  }
};

upload.onchange = () => {
  if (upload.files[0]) connect(undefined, upload.files[0]);
  upload.value = '';
};

async function restoreWorkbook() {
  if (!window.showOpenFilePicker) {
    status.textContent = 'اختر ملف الإكسل. لحفظ الربط والتحديث التلقائي استخدم Chrome أو Edge.';
    return;
  }
  try {
    const saved = await workbookStore('readonly', store => store.get('workbook'));
    if (!saved || selection || remembered) return;
    const permission = await saved.queryPermission({mode: 'read'});
    if (selection || remembered) return;
    remembered = saved;
    if (permission === 'granted') {
      connect(saved);
    } else {
      choose.textContent = 'السماح بقراءة الملف المحفوظ';
      status.textContent = 'ملف الإكسل محفوظ. اسمح للمتصفح بقراءته.';
    }
  } catch {
    // The file picker remains usable if browser storage is unavailable.
  }
}

restoreWorkbook();
