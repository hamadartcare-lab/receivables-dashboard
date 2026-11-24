const status = document.getElementById('sync-status');
const choose = document.getElementById('choose-workbook');
const upload = document.getElementById('workbook-input');
let handle, signature, timer, selection = 0;
if (!window.showOpenFilePicker) {
  status.textContent = 'اختر ملف الإكسل. للتحديث التلقائي استخدم Chrome أو Edge.';
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
  refresh(++selection, file);
}
choose.onclick = async () => {
  if (!window.showOpenFilePicker) return upload.click();
  try {
    const [file] = await window.showOpenFilePicker({
      types: [{description: 'Excel', accept: {'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':['.xlsx']}}]
    });
    connect(file);
  } catch (error) {
    if (error.name === 'SecurityError') upload.click();
    else if (error.name !== 'AbortError') status.textContent = 'تعذر فتح الملف';
  }
};
upload.onchange = () => {
  if (upload.files[0]) connect(undefined, upload.files[0]);
  upload.value = '';
};
