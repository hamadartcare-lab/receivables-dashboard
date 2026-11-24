const clean = value => String(value ?? '').normalize('NFKC').trim();
const money = value => Math.round(Number(value || 0) * 100) / 100;
const statuses = new Map([
  ['دفعت','paid'], ['مدفوعة','paid'], ['تم الدفع','paid'],
  ['لم تدفع','unpaid'], ['غير مدفوعة','unpaid'],
  ['لم تسلم','undelivered'], ['لم تُسلَّم','undelivered'],
]);

function readInvoices(sheet, marketplace) {
  const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, blankrows: true, range: {s: {r: 0, c: 0}, e: XLSX.utils.decode_range(sheet['!ref']).e}});
  for (const {s, e} of sheet['!merges'] || []) {
    if (s.c !== e.c) continue;
    for (let row = s.r + 1; row <= e.r; row++) {
      if (rows[row]) rows[row][s.c] ??= rows[s.r]?.[s.c];
    }
  }
  const shift = Number(marketplace);
  const invoices = [];
  const groups = new Map();
  rows.forEach((row, index) => {
    const inv = clean(row[2 + shift]);
    const qty = Number(row[4 + shift]), price = Number(row[5 + shift]);
    if (!inv || !Number.isFinite(qty) || !Number.isFinite(price)) return;
    const status = statuses.get(clean(row[7 + shift]));
    if (!status || (marketplace && status === 'undelivered')) {
      throw new Error('حالة فاتورة غير معروفة في الصف ' + (index + 1));
    }
    const po = clean(row[1 + shift]), invNew = clean(row[3 + shift]).replace(/^-+$/, '');
    const line = {qty, price, total: money(row[6 + shift] ?? qty * price)};
    if (!marketplace) {
      invoices.push({po, inv, invNew, status, ...line});
      return;
    }
    const key = JSON.stringify([po, inv]);
    if (!groups.has(key)) {
      groups.set(key, {po, inv, invNew, status, submitted: ['1','نعم','true'].includes(clean(row[9])),
        note: clean(row[1]), lines: [], total: 0});
    }
    const group = groups.get(key);
    if (group.status !== status) throw new Error('حالات مختلفة لبنود الفاتورة ' + inv);
    group.lines.push(line);
    group.total = money(group.total + line.total);
  });
  return marketplace ? [...groups.values()] : invoices;
}

function readCustomers(sheet) {
  const rows = XLSX.utils.sheet_to_json(sheet, {header: 1, blankrows: true, range: {s: {r: 0, c: 0}, e: XLSX.utils.decode_range(sheet['!ref']).e}});
  const customers = [];
  rows.forEach((header, index) => header.forEach((heading, col) => {
    if (clean(heading) !== 'النوع') return;
    const start = clean(header[col - 1]) === 'التاريخ' ? col - 1 : col;
    const fields = new Map();
    for (let c = start; c < col + 7; c++) {
      const field = clean(header[c]);
      fields.set(field, c);
      if (field === 'الرصيد') break;
    }
    if (!['المرجع','مدين','دائن','الرصيد'].every(field => fields.has(field))) return;
    const branchColumn = fields.get('وصف العملية');
    const name = branchColumn === undefined
      ? clean(rows[index - 1]?.[start])
      : clean(rows[index + 1]?.[branchColumn]).split(' - ')[0];
    if (!name) return;
    const tx = [];
    let balance;
    for (let r = index + 1; r < rows.length; r++) {
      const row = rows[r];
      const get = field => row[fields.get(field)];
      if (row.slice(start, fields.get('الرصيد')).some(value => clean(value).includes('الرصيد الختامي'))) {
        balance = get('الرصيد');
        break;
      }
      const raw = ['النوع','المرجع','مدين','دائن','الرصيد'].map(get);
      if (raw.every(value => value == null)) continue;
      const [type, ref, dr, cr, bal] = raw;
      const dateValue = get('التاريخ');
      const serial = typeof dateValue === 'number' ? XLSX.SSF.parse_date_code(dateValue) : null;
      const date = serial ? [serial.y, serial.m, serial.d].map(n => String(n).padStart(2, '0')).join('-') : clean(dateValue);
      tx.push({type: clean(type), ref: clean(ref), dr: Number(dr || 0), cr: Number(cr || 0),
        bal: bal == null ? null : Number(bal), branch: clean(get('وصف العملية')), date});
    }
    if (tx.length || balance != null) {
      customers.push({name, tx, balance: money(balance ?? tx.findLast(t => t.bal != null)?.bal)});
    }
  }));
  return customers;
}

function parseWorkbook(bytes, modified) {
  const {Sheets} = XLSX.read(bytes, {bookVBA: false});
  function sheet(name) {
    if (!Sheets[name]) throw new Error('ورقة مفقودة: ' + name);
    return Sheets[name];
  }
  const regions = ['الوسطى','الغربية','الشرقية','الشمالية'].map(name => {
    const source = name === 'الشمالية' && !Sheets[name] ? 'الشماليه' : name;
    const customers = readCustomers(sheet(source));
    if (!customers.length) throw new Error('لم يتم العثور على حسابات ' + name);
    return {name, customers, total: money(customers.reduce((sum, c) => sum + c.balance, 0))};
  });
  return {asOf: new Date(modified).toLocaleString('ar-JO'), regions,
    nupco: readInvoices(sheet('نبكو'), false), market: readInvoices(sheet('ماركت بليس'), true)};
}
