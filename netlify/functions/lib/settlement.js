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
    const balance = round1(cashOut - totalBuyIn); // balance neto informativo (gana/pierde en total)

    // Regla: el cash out primero salda el buy-in virtual. Lo que sobra de eso
    // ("netClaim") es lo que el jugador realmente puede reclamar del pozo de
    // cash real — no el balance total. Si netClaim <= 0, ni siquiera alcanzó
    // para saldar el virtual, y esa diferencia se debe por transferencia.
    const netClaim = round1(cashOut - virtualAmount);

    let pagoCash = 0;
    let pagoTransfer = 0;
    if (netClaim > 0) {
      pagoCash = round1(Math.min(cashAmount, netClaim));
      pagoTransfer = round1(netClaim - pagoCash);
    } else {
      pagoCash = 0;
      pagoTransfer = netClaim; // negativo: debe transferir
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
  const cashPool = round1(totalCashAll - rake);

  // Todo el cash disponible debe entregarse. Si el tope individual (nadie
  // cobra en cash más de lo que él mismo puso) deja un remanente sin asignar,
  // se le da al jugador con mayor balance positivo, convirtiendo esa parte de
  // su transferencia pendiente en cash físico.
  let leftoverCash = round1(cashPool - players.reduce((s, p) => s + p.pagoCash, 0));
  if (leftoverCash > 0) {
    const byBalanceDesc = players.filter((p) => p.balance > 0).sort((a, b) => b.balance - a.balance);
    for (const w of byBalanceDesc) {
      if (leftoverCash <= 0) break;
      const shift = round1(Math.min(leftoverCash, w.pagoTransfer));
      if (shift > 0) {
        w.pagoCash = round1(w.pagoCash + shift);
        w.pagoTransfer = round1(w.pagoTransfer - shift);
        leftoverCash = round1(leftoverCash - shift);
      }
    }
  }

  // Transferencias sugeridas: se emparejan los que más reciben con los que más
  // deben. Es un circuito cerrado entre jugadores (no involucra el rake, que
  // ya se descontó del pozo de cash), así que suman exactamente 0.
  const creditors = players
    .filter((p) => p.pagoTransfer > 0)
    .map((p) => ({ playerId: p.playerId, name: p.name, amount: p.pagoTransfer }))
    .sort((a, b) => b.amount - a.amount);
  const debtors = players
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
