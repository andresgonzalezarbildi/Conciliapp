const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const context = {
  console,
  Intl,
  Date,
  Math,
  Set,
  Map,
  Promise,
  Object,
  Array,
  Number,
  String,
  Boolean,
  RegExp,
  JSON,
  Error,
  Blob: class {},
  URL: { createObjectURL(){ return ''; }, revokeObjectURL(){} },
  FormData: class {},
  window: {},
  document: {
    addEventListener(){},
    getElementById(){ return null; },
    querySelectorAll(){ return []; },
    querySelector(){ return null; },
    createElement(){ return { append(){}, click(){}, style:{}, remove(){} }; },
    body: { append(){} }
  },
  setTimeout,
  clearTimeout,
};
context.window = context;
vm.createContext(context);
vm.runInContext(fs.readFileSync(require('path').join(__dirname, '..', 'app.js'), 'utf8'), context, { filename: 'app.js' });
const app = context.ReconciliationApp;
assert(app, 'exports missing');

const day = d => new Date(`${d}T12:00:00Z`).getTime();
const mov = (id, source, date, description, amount) => ({ id, source, row: 1, dateTime: day(date), description, amount, type: amount >= 0 ? 'debit' : 'credit' });

(async () => {
  const base = { ...app.DEFAULT_CONFIG, dateTolerance: 1, amountAbsTolerance: 0.1, amountPercentTolerance: 0, autoThreshold: 70, possibleThreshold: 55, searchOneToOne: true, searchOneToMany: false, searchManyToOne: false, searchInternalOffsets: false };
  const result = await app.reconciliationEngine({
    system: [
      mov('s-transfer','system','2026-06-01','TRASPASO DE USD',210000),
      mov('s-loan','system','2026-06-02','PRESTAMO A BLUE PETER',210000),
    ],
    bank: [
      mov('b-loan','bank','2026-06-01','PRESTAMO A BLUE PETER',210000),
      mov('b-transfer','bank','2026-06-02','IngresoBan Traspaso ITAU USD',210000),
    ],
    config: base,
    excludedSignatures: []
  });
  const pairs = new Set(result.reconciliations.map(r => `${r.systemIds[0]}:${r.bankIds[0]}`));
  assert(pairs.has('s-transfer:b-transfer'), `global transfer pair missing: ${[...pairs]}`);
  assert(pairs.has('s-loan:b-loan'), `global loan pair missing: ${[...pairs]}`);

  const mixed = await app.reconciliationEngine({
    system: [mov('s-ute','system','2026-07-17','UTE',100)],
    bank: [mov('b-1','bank','2026-07-17','UTE factura',150), mov('b-2','bank','2026-07-17','UTE devolucion IVA',-50)],
    config: { ...base, searchOneToOne: false, searchOneToMany: true, allowMixedGroupSigns: true },
    excludedSignatures: []
  });
  assert(mixed.reconciliations.length >= 1, 'mixed group not found');
  assert(mixed.reconciliations.every(r => r.status === 'possible'), 'mixed-sign group auto-confirmed');

  const internal = await app.reconciliationEngine({
    system: [mov('s-guille','system','2026-07-01','Cobranza Buenos Aires Guille',1700), mov('s-piotto','system','2026-07-01','Pago Agustin Piotto',-1700)],
    bank: [],
    config: { ...base, searchOneToOne: false, searchInternalOffsets: true },
    excludedSignatures: []
  });
  assert(internal.reconciliations.length >= 1, 'internal offset not found');
  assert(internal.reconciliations.every(r => r.status === 'possible'), 'unrelated internal offset auto-confirmed');

  const snapshotA = {
    schemaVersion: 1,
    exportFileName: 'Caja pesos',
    workspace: { id: 'a', name: 'Caja pesos' },
    sources: {
      system: { movements: [{ id:'system-1', source:'system', row:47, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Recibo P UTE', amount:287227.43, type:'debit', status:'' }] },
      bank: { movements: [] }
    },
    results: { reconciliations: [] },
    review: { periodFilter: { from:'', to:'' } }
  };
  const snapshotB = {
    schemaVersion: 1,
    exportFileName: 'Itau pesos',
    workspace: { id: 'b', name: 'Itau pesos' },
    sources: { system: { movements: [] }, bank: { movements: [] } },
    results: { reconciliations: [] },
    review: { periodFilter: { from:'', to:'' } }
  };
  assert(app.moveSnapshotMovement(snapshotA, snapshotB, 'system', 'system-1'), 'transfer failed');
  assert.equal(snapshotA.sources.system.movements.length, 0);
  assert.equal(snapshotB.sources.system.movements.length, 1);
  assert.equal(snapshotB.sources.system.movements[0].transferOriginAccount, 'Caja pesos');
  assert.notEqual(snapshotB.sources.system.movements[0].id, 'system-1');

  console.log('All tests passed');
})().catch(err => { console.error(err); process.exit(1); });
