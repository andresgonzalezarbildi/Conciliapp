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
assert.equal(app.getState().movementEditor.statusFilter, 'pending', 'movement editor must open with pending movements');
const editorCss = fs.readFileSync(require('path').join(__dirname, '..', 'styles.css'), 'utf8');
assert(/\.movement-editor-card\s*\{[^}]*display:\s*flex;[^}]*flex-direction:\s*column;/s.test(editorCss), 'movement editor card must constrain the scrollable content');
assert(/\.movement-editor-content\s*\{[^}]*overflow-y:\s*scroll;/s.test(editorCss), 'movement editor content must have vertical scrolling');

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

  const snapshotC = {
    schemaVersion: 1,
    exportFileName: 'Caja dólares',
    workspace: { id: 'c', name: 'Caja dólares' },
    sources: {
      system: { movements: [{ id:'system-credit', source:'system', row:9, date:'2026-07-18T00:00:00.000Z', dateKey:'2026-07-18', description:'Movimiento invertido', amount:500, debitAmount:500, creditAmount:0, type:'debit', status:'' }] },
      bank: { movements: [] }
    },
    results: { reconciliations: [] },
    review: { periodFilter: { from:'', to:'' } }
  };
  const snapshotD = {
    schemaVersion: 1,
    exportFileName: 'Itau dólares',
    workspace: { id: 'd', name: 'Itau dólares' },
    sources: { system: { movements: [] }, bank: { movements: [] } },
    results: { reconciliations: [] },
    review: { periodFilter: { from:'', to:'' } }
  };
  assert(app.moveSnapshotMovement(snapshotC, snapshotD, 'system', 'system-credit', 'credit'), 'typed transfer failed');
  const transferred = snapshotD.sources.system.movements[0];
  assert.equal(transferred.type, 'credit');
  assert.equal(transferred.amount, -500);
  assert.equal(transferred.debitAmount, 0);
  assert.equal(transferred.creditAmount, 500);
  assert.equal(snapshotC.transferLog[0].originalType, 'debit');
  assert.equal(snapshotD.transferLog[0].destinationType, 'credit');
  assert.equal(app.moveSnapshotMovement(snapshotC, snapshotD, 'system', 'system-credit', 'credit'), false, 'same movement transferred twice');
  assert.equal(snapshotD.sources.system.movements.length, 1);

  const editorCaja = {
    schemaVersion: 1,
    workspace: { id: 'editor-caja', name: 'Caja pesos' },
    exportFileName: 'Caja pesos',
    sources: {
      system: { movements: [{ id:'system-ute', source:'system', row:47, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Recibo P UTE', amount:287227.43, debitAmount:287227.43, creditAmount:0, type:'debit', status:'' }] },
      bank: { movements: [] }
    },
    results: { reconciliations: [{ id:'CON-UTE', status:'confirmed', systemIds:['system-ute'], bankIds:[] }] },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [],
    movementEditLog: []
  };
  const editorItau = {
    schemaVersion: 1,
    workspace: { id: 'editor-itau', name: 'ITAU pesos' },
    exportFileName: 'ITAU pesos',
    sources: { system: { movements: [] }, bank: { movements: [] } },
    results: { reconciliations: [] },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [],
    movementEditLog: []
  };
  const editedMove = app.editMovementAcrossSnapshots([editorCaja, editorItau], {
    accountId: 'editor-caja', sourceKey: 'system', movementId: 'system-ute'
  }, {
    accountId: 'editor-itau', sourceKey: 'bank', date: '2026-07-18', description: 'UTE corregido', amount: 300000, type: 'credit', status: 'ajustado'
  });
  assert.equal(editorCaja.sources.system.movements.length, 0, 'editor did not remove from origin');
  assert.equal(editorCaja.results.reconciliations.length, 0, 'related reconciliation was not invalidated');
  assert.equal(editedMove.invalidated, 1, 'invalidated count mismatch');
  const editedUte = editorItau.sources.bank.movements[0];
  assert.equal(editedUte.amount, -300000, 'credit sign not applied');
  assert.equal(editedUte.debitAmount, 0, 'credit kept debit value');
  assert.equal(editedUte.creditAmount, 300000, 'credit column not set');
  assert.equal(editedUte.type, 'credit', 'credit type not set');
  assert.equal(editedUte.description, 'UTE corregido');
  assert.equal(editedUte.dateKey, '2026-07-18');
  assert.equal(editedUte.source, 'bank');
  assert.equal(editorCaja.movementEditLog[0].action, 'move');
  assert.equal(editorItau.transferLog[0].destinationType, 'credit');

  const sameAccountEdit = app.editMovementAcrossSnapshots([editorItau], {
    accountId: 'editor-itau', sourceKey: 'bank', movementId: editedUte.id
  }, {
    accountId: 'editor-itau', sourceKey: 'bank', date: '2026-07-19', description: 'UTE débito corregido', amount: 1500.5, type: 'debit', status: ''
  });
  assert.equal(sameAccountEdit.movement.amount, 1500.5, 'debit sign not applied');
  assert.equal(sameAccountEdit.movement.debitAmount, 1500.5);
  assert.equal(sameAccountEdit.movement.creditAmount, 0);
  assert.equal(sameAccountEdit.movement.type, 'debit');

  const deleted = app.deleteMovementAcrossSnapshots([editorItau], {
    accountId: 'editor-itau', sourceKey: 'bank', movementId: sameAccountEdit.movement.id
  });
  assert.equal(editorItau.sources.bank.movements.length, 0, 'editor delete failed');
  assert.equal(deleted.movement.description, 'UTE débito corregido');

  const crossSystemAccount = {
    schemaVersion: 1,
    workspace: { id: 'cross-system-account', name: 'ITAU pesos' },
    exportFileName: 'ITAU pesos',
    sources: {
      system: { movements: [{ id:'cross-system-movement', source:'system', row:10, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'UTE pago factura julio', amount:287227.43, debitAmount:287227.43, creditAmount:0, type:'debit', status:'' }] },
      bank: { movements: [] }
    },
    results: { reconciliations: [], nextId: 1 },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [],
    movementEditLog: []
  };
  const crossBankAccount = {
    schemaVersion: 1,
    workspace: { id: 'cross-bank-account', name: 'Caja pesos' },
    exportFileName: 'Caja pesos',
    sources: {
      system: { movements: [] },
      bank: { movements: [{ id:'cross-bank-movement', source:'bank', row:47, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Recibo P UTE', amount:287227.43, debitAmount:287227.43, creditAmount:0, type:'debit', status:'' }] }
    },
    results: { reconciliations: [], nextId: 1 },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [],
    movementEditLog: []
  };
  const crossCandidates = app.findCrossAccountCandidates([{ snapshot: crossSystemAccount }, { snapshot: crossBankAccount }]);
  assert.equal(crossCandidates.length, 1, 'cross-account pending match was not detected');
  assert.equal(crossCandidates[0].systemAccountId, 'cross-system-account');
  assert.equal(crossCandidates[0].bankAccountId, 'cross-bank-account');

  const runtimeState = app.getState();
  runtimeState.movementEditor.accounts = [{ snapshot: crossSystemAccount }, { snapshot: crossBankAccount }];
  runtimeState.movementEditor.modifiedAccountIds = new Set();
  runtimeState.movementEditor.selectedAccountId = '__summary__';
  const crossResolution = app.resolveMovementEditorCross('cross-system-account', 'cross-system-movement', 'cross-bank-account', 'cross-bank-movement', 'system');
  assert.equal(crossResolution.targetAccountId, 'cross-system-account');
  assert.equal(crossBankAccount.sources.bank.movements.length, 0, 'cross resolution did not remove movement from origin account');
  assert.equal(crossSystemAccount.sources.bank.movements.length, 1, 'cross resolution did not move bank movement to target account');
  assert.equal(crossSystemAccount.results.reconciliations.length, 1, 'cross resolution did not create reconciliation');
  assert.equal(crossSystemAccount.results.reconciliations[0].status, 'confirmed');
  assert.equal(crossSystemAccount.results.reconciliations[0].crossAccountResolved, true);
  assert(runtimeState.movementEditor.modifiedAccountIds.has('cross-system-account'), 'target account not marked for download');
  assert(runtimeState.movementEditor.modifiedAccountIds.has('cross-bank-account'), 'origin account not marked for download');

  const groupedSystemAccount = {
    schemaVersion: 1,
    workspace: { id: 'grouped-system-account', name: 'ITAU pesos agrupado' },
    exportFileName: 'ITAU pesos agrupado',
    sources: {
      system: { movements: [
        { id:'ute-1', source:'system', row:1, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Ute', amount:1000, debitAmount:1000, creditAmount:0, type:'debit', status:'' },
        { id:'ute-2', source:'system', row:2, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Ute', amount:500, debitAmount:500, creditAmount:0, type:'debit', status:'' },
        { id:'ute-3', source:'system', row:3, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Ute rediva', amount:-100, debitAmount:0, creditAmount:100, type:'credit', status:'' }
      ] },
      bank: { movements: [] }
    },
    results: { reconciliations: [], nextId: 1 },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [], movementEditLog: []
  };
  const groupedBankAccount = {
    schemaVersion: 1,
    workspace: { id: 'grouped-bank-account', name: 'Caja pesos agrupado' },
    exportFileName: 'Caja pesos agrupado',
    sources: {
      system: { movements: [] },
      bank: { movements: [{ id:'ute-bank', source:'bank', row:47, date:'2026-07-17T00:00:00.000Z', dateKey:'2026-07-17', description:'Recibo P UTE', amount:-1400, debitAmount:0, creditAmount:1400, type:'credit', status:'' }] }
    },
    results: { reconciliations: [], nextId: 1 },
    review: { periodFilter: { from:'', to:'' } },
    transferLog: [], movementEditLog: []
  };
  const groupedAccounts = [{ snapshot: groupedSystemAccount }, { snapshot: groupedBankAccount }];
  const groupedCandidates = app.findCrossAccountCandidates(groupedAccounts);
  const groupedCandidate = groupedCandidates.find(item => item.systemMovements.length === 3 && item.bankMovements.length === 1);
  assert(groupedCandidate, 'grouped cross-account candidate was not detected');
  assert.equal(groupedCandidate.signCorrectionNeeded, true, 'opposite sign was not flagged');
  runtimeState.movementEditor.accounts = groupedAccounts;
  runtimeState.movementEditor.modifiedAccountIds = new Set();
  const groupedResolution = app.resolveMovementEditorCrossCandidate(groupedCandidate, 'system');
  assert.equal(groupedResolution.adjustedSigns, 1, 'destination sign was not corrected');
  const groupedReconciliation = groupedSystemAccount.results.reconciliations[0];
  assert.equal(groupedReconciliation.systemIds.length, 3);
  assert.equal(groupedReconciliation.bankIds.length, 1);
  assert.equal(groupedReconciliation.totalSystem, 1400);
  assert.equal(groupedReconciliation.totalBank, 1400);
  assert.equal(groupedReconciliation.difference, 0);
  assert.equal(groupedSystemAccount.sources.bank.movements[0].type, 'debit');
  assert.equal(groupedSystemAccount.sources.bank.movements[0].amount, 1400);

  console.log('All tests passed');
})().catch(err => { console.error(err); process.exit(1); });
