// Redondeo defensivo para evitar arrastres de punto flotante en los cálculos de liquidación
function round1(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

// Misma lógica que computeSettlement() en la app (src/App.jsx). Se mantiene
// duplicada acá porque las funciones de Netlify corren en su propio runtime
// (Node) y no pueden importar directamente el bundle de React.
function computeSettlement(game, roster) {
  const players = game.playerIds.map((pid) => {
    const p = roster.find((r) => r.id === pid);
    const purchases = game.purchases.filter((pu) => pu.playerId === pid);
    const cashAmount = purchases.filter((pu) => pu.type === "cash").reduce((s, pu) => s + pu.amount, 0);
    const virtualAmount = purchases.filter((pu) => pu.type === "virtual").reduce((s, pu) => s + pu.amount, 0);
    const totalBuyIn = cashAmount + virtualAmount;
    const cashOut = Number((game.finalChips || {})[pid]) || 0;
    const balance = round1(cashOut - totalBuyIn);

    let pagoCash = 0;
    let pagoTransfer = 0;

    if (balance > 0) {
      const afterVirtual = cashOut - virtualAmount;
      if (virtualAmount > 0 && afterVirtual < 0) {
        pagoCash = 0;
        pagoTransfer = round1(cashOut - totalBuyIn);
      } else {
        pagoCash = round1(Math.min(cashAmount, Math.max(0, afterVirtual)));
        pagoTransfer = round1(balance);
      }
    } else {
      pagoCash = 0;
      pagoTransfer = balance;
    }

    return {
      playerId: pid,
      name: p ? p.name : "?",
      cashAmount, virtualAmount, totalBuyIn,
      cashOut, balance,
      pagoCash, pagoTransfer,
    };
  });

  const totalCashAll = players.reduce((s, p) => s + p.cashAmount, 0);
  const rake = Math.min(game.rake || 0, totalCashAll);

  const winners = players.filter((p) => p.balance > 0);
  const losers = players.filter((p) => p.balance < 0);

  const creditors = winners
    .filter((p) => p.pagoTransfer > 0)
    .map((p) => ({ playerId: p.playerId, name: p.name, amount: p.pagoTransfer }))
    .sort((a, b) => b.amount - a.amount);
  const debtors = losers
    .filter((p) => p.pagoTransfer < 0)
    .map((p) => ({ playerId: p.playerId, name: p.name, amount: -p.pagoTransfer }))
    .sort((a, b) => b.amount - a.amount);

  const transfers = [];
  let i = 0, j = 0;
  while (i < debtors.length && j < creditors.length) {
    const d = debtors[i], c = creditors[j];
    const amt = round1(Math.min(d.amount, c.amount));
    if (amt > 0) transfers.push({ from: d.name, fromId: d.playerId, to: c.name, toId: c.playerId, amount: amt });
    d.amount = round1(d.amount - amt);
    c.amount = round1(c.amount - amt);
    if (d.amount <= 0) i++;
    if (c.amount <= 0) j++;
  }

  return { players, rake, totalCashAll, transfers };
}

module.exports = { round1, computeSettlement };
