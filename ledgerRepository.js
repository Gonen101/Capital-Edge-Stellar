// ledgerRepository.js
//
// The entire risk surface of the relational ledger migration lives in this
// one file. App.jsx never knows accounts/transactions came from real
// tables instead of a JSON blob - this reconstructs the identical shape it
// has always received, and translates changes back into targeted writes.
//
// rowsToAccounts / rowsToTransactions: read path (database rows -> the
// exact data.accounts / data.transactions arrays the app expects).
//
// diffAndBuildWrites: write path (old array, new array -> the minimal set
// of inserts/upserts/deletes needed), built on the app's own
// diffCollection function so "what changed" is computed exactly the same
// way the audit log already computes it - not a second, separate
// implementation that could disagree with the first.

// Exact copy of App.jsx's own diffCollection - not imported, since App.jsx
// intentionally exports nothing but the App component itself (kept
// byte-identical to the tested standalone artifact). Verified identical to
// the original as part of this migration's test suite.
function diffCollection(prevArr, nextArr) {
  const a = prevArr || [], b = nextArr || [];
  if (a === b) return null;
  const prevMap = new Map(a.map(x => [x.id, x]));
  const nextMap = new Map(b.map(x => [x.id, x]));
  const added = [...nextMap.keys()].filter(id => !prevMap.has(id));
  const removed = [...prevMap.keys()].filter(id => !nextMap.has(id));
  const changed = [...nextMap.keys()].filter(id => prevMap.has(id) && JSON.stringify(prevMap.get(id)) !== JSON.stringify(nextMap.get(id)));
  if (added.length === 0 && removed.length === 0 && changed.length === 0) return null;
  return { added, removed, changed, prevMap, nextMap };
}

export function rowsToAccounts(accountRows) {
  return (accountRows || []).map((r) => {
    const a = {
      id: r.id,
      code: r.code,
      name: r.name,
      type: r.type,
      subtype: r.subtype ?? undefined,
      category: r.category ?? undefined,
      status: r.status || "active",
      parentId: r.parent_id ?? null,
      description: r.description ?? "",
      taxAccountId: r.tax_account_id ?? "",
      currency: r.currency ?? "",
    };
    if (r.normal) a.normal = r.normal;
    if (r.contra) a.contra = true;
    return a;
  });
}

export function rowsToTransactions(txnRows, lineRows) {
  const linesByTxn = new Map();
  for (const l of lineRows || []) {
    const key = l.transaction_id;
    if (!linesByTxn.has(key)) linesByTxn.set(key, []);
    linesByTxn.get(key).push(l);
  }
  return (txnRows || []).map((t) => {
    const lines = (linesByTxn.get(t.id) || [])
      .slice()
      .sort((a, b) => (a.line_order ?? 0) - (b.line_order ?? 0))
      .map((l) => ({ accountId: l.account_id, debit: Number(l.debit) || 0, credit: Number(l.credit) || 0 }));
    return { id: t.id, date: t.date, memo: t.memo || "", source: t.source || "manual", docId: t.doc_id ?? null, lines };
  });
}

export function accountToRow(companyId, a) {
  return {
    id: a.id, company_id: companyId, code: String(a.code), name: a.name, type: a.type,
    subtype: a.subtype ?? null, category: a.category ?? null, status: a.status || "active",
    parent_id: a.parentId ?? null, description: a.description ?? "", tax_account_id: a.taxAccountId ?? "",
    currency: a.currency ?? "", normal: a.normal ?? null, contra: !!a.contra,
  };
}

export function transactionToRow(companyId, t) {
  return { id: t.id, company_id: companyId, date: t.date, memo: t.memo || "", source: t.source || "manual", doc_id: t.docId ?? null };
}

export function transactionLinesToRows(companyId, t) {
  return (t.lines || []).map((l, i) => ({
    company_id: companyId, transaction_id: t.id, account_id: l.accountId,
    debit: Number(l.debit) || 0, credit: Number(l.credit) || 0, line_order: i,
  }));
}

// Builds the exact set of database operations needed to bring the
// accounts table in line with a new accounts array, given what it was
// before. Returns null if nothing changed (skips a round of empty writes).
export function diffAccountWrites(companyId, prevAccounts, nextAccounts) {
  const d = diffCollection(prevAccounts, nextAccounts);
  if (!d) return null;
  const upserts = [...d.added, ...d.changed].map((id) => accountToRow(companyId, d.nextMap.get(id)));
  return { removedIds: d.removed, upserts };
}

// Same idea for transactions, but a transaction has both a header and a
// variable number of lines - changed transactions get their lines fully
// replaced (delete then insert) rather than diffed line-by-line, since
// lines don't carry their own stable id in the app's data model. This is
// the deliberately simple, correct choice: transactions are small and
// change relatively rarely once posted, so a full replace of a changed
// transaction's lines is safe and easy to reason about, instead of a more
// intricate per-line diff that would be harder to get right.
export function diffTransactionWrites(companyId, prevTransactions, nextTransactions) {
  const d = diffCollection(prevTransactions, nextTransactions);
  if (!d) return null;
  const changedIds = [...d.added, ...d.changed];
  const txnUpserts = changedIds.map((id) => transactionToRow(companyId, d.nextMap.get(id)));
  const lineInserts = changedIds.flatMap((id) => transactionLinesToRows(companyId, d.nextMap.get(id)));
  return { removedIds: d.removed, changedIds, txnUpserts, lineInserts };
}
