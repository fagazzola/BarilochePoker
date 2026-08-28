import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Plus, Minus, Trash2, Pencil, Users, UtensilsCrossed, Wine,
  Trophy, ArrowRightLeft, Save, X, Check, ChevronDown, ChevronUp, ChevronLeft,
  Banknote, Landmark, Flame, History, UserPlus, UserX, UserCheck,
  Play, Square, AlertCircle, Crown, DollarSign, CircleDollarSign, Coins,
  BarChart3, Activity, Settings
} from "lucide-react";

/* ----------------------------------------------------------------------
   THEME
---------------------------------------------------------------------- */
const FONT_IMPORT =
  "@import url('https://fonts.googleapis.com/css2?family=Bebas+Neue&family=IBM+Plex+Mono:wght@400;500;600;700&family=Inter:wght@400;500;600;700&display=swap');";

const C = {
  felt: "#0c3527",
  feltDeep: "#082019",
  feltLine: "rgba(212,175,55,0.14)",
  card: "#f4ead6",
  cardDim: "#e7dabd",
  ink: "#241d12",
  inkSoft: "#5c5240",
  gold: "#d8ad3f",
  goldSoft: "#f0d888",
  cash: "#2fae66",
  cashDeep: "#1d7d49",
  virtual: "#f2883c",
  virtualDeep: "#c9631f",
  win: "#3fbf72",
  loss: "#e2634f",
  panel: "#123f30",
  panelLine: "rgba(212,175,55,0.22)",
};

const displayFont = { fontFamily: "'Bebas Neue', sans-serif", letterSpacing: "0.04em" };
const monoFont = { fontFamily: "'IBM Plex Mono', monospace" };
const bodyFont = { fontFamily: "'Inter', sans-serif" };

/* ----------------------------------------------------------------------
   STORAGE
---------------------------------------------------------------------- */
const KEYS = { roster: "poker-roster", games: "poker-games", active: "poker-active-game", adminPassword: "poker-admin-password" };

// Guardado vía funciones de Netlify -> Microsoft Graph -> Excel (OneDrive).
// Reemplaza el window.storage propio de los artifacts de Claude, que no
// existe fuera de ese entorno.
async function loadKey(key, fallback) {
  try {
    const res = await fetch(`/api/store?key=${encodeURIComponent(key)}`);
    if (!res.ok) return fallback;
    const data = await res.json();
    return data && data.value !== undefined && data.value !== null ? data.value : fallback;
  } catch (e) {
    console.error("load failed", key, e);
    return fallback;
  }
}
async function saveKey(key, value) {
  try {
    const res = await fetch(`/api/store?key=${encodeURIComponent(key)}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ value }),
    });
    if (!res.ok) {
      const t = await res.text();
      console.error("storage set failed", key, t);
    }
  } catch (e) {
    console.error("storage set failed", key, e);
  }
}

const uid = () => Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
const money = (n) =>
  "$" + Math.round(n || 0).toLocaleString("en-US", { maximumFractionDigits: 0 });
const todayISO = () => new Date().toISOString().slice(0, 10);
// Redondeo defensivo para evitar arrastres de punto flotante (ej. 0.1+0.2) en los cálculos de liquidación
const round1 = (n) => Math.round((n + Number.EPSILON) * 100) / 100;
// Unidad mínima de valor: $100. roundTo100 para redondear al más cercano
// (usado al "aterrizar" montos ingresados), ceilTo100 para redondear siempre
// hacia arriba (usado en cargos repartidos, para no cobrar de menos).
const roundTo100 = (n) => Math.round((Number(n) || 0) / 100) * 100;
// Pide la contraseña de administrador (cargada a mano en la hoja "Meta" del
// Excel) antes de dejar pasar una acción destructiva. No es seguridad real
// (no hay login), es solo una traba para evitar borrados accidentales.
// La UI es un modal propio (ver <AdminPasswordModal/>, montado una sola vez
// en la raíz de la app) con input enmascarado tipo password — a diferencia
// de window.prompt(), que no se puede estilizar ni ocultar el texto tipeado.
let _showAdminPasswordModal = null;
function requestAdminPassword(adminPassword, actionLabel) {
  return new Promise((resolve) => {
    if (!adminPassword) {
      alert(`Todavía no hay una contraseña de administrador configurada en el Excel (hoja "Meta", fila con key=admin_password). Agregala ahí para poder ${actionLabel}.`);
      resolve(false);
      return;
    }
    if (!_showAdminPasswordModal) {
      resolve(false);
      return;
    }
    _showAdminPasswordModal(actionLabel, (entered) => {
      if (entered === null) { resolve(false); return; }
      if (entered !== adminPassword) {
        alert("Contraseña incorrecta.");
        resolve(false);
        return;
      }
      resolve(true);
    });
  });
}
function AdminPasswordModal() {
  const [pending, setPending] = useState(null); // { actionLabel, onSubmit }
  const [value, setValue] = useState("");

  useEffect(() => {
    _showAdminPasswordModal = (actionLabel, onSubmit) => {
      setValue("");
      setPending({ actionLabel, onSubmit });
    };
    return () => { _showAdminPasswordModal = null; };
  }, []);

  if (!pending) return null;
  const submit = (entered) => {
    const cb = pending.onSubmit;
    setPending(null);
    cb(entered);
  };

  return (
    <div
      style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)",
        display: "flex", alignItems: "center", justifyContent: "center",
        zIndex: 1000, padding: 16,
      }}
      onClick={() => submit(null)}
    >
      <div
        style={{
          background: C.panel, border: `1px solid ${C.panelLine}`, borderRadius: 14,
          padding: 20, width: "min(340px, 100%)", boxShadow: "0 12px 32px rgba(0,0,0,0.4)",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ ...displayFont, fontSize: 17, color: C.card, marginBottom: 4 }}>Contraseña de administrador</div>
        <div style={{ fontSize: 12.5, color: "rgba(244,234,214,0.6)", marginBottom: 14 }}>Para {pending.actionLabel}:</div>
        <input
          type="password"
          autoFocus
          style={inputStyle}
          value={value}
          onChange={(e) => setValue(e.target.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter") submit(value);
            if (e.key === "Escape") submit(null);
          }}
        />
        <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
          <GhostBtn onClick={() => submit(null)}>Cancelar</GhostBtn>
          <PrimaryBtn onClick={() => submit(value)}>Confirmar</PrimaryBtn>
        </div>
      </div>
    </div>
  );
}
const ceilTo100 = (n) => Math.ceil((Number(n) || 0) / 50) * 50;

// Confirmación con el look de la app, para reemplazar el confirm() nativo del
// navegador (que no se puede estilizar) en acciones destructivas.
let _showConfirmModal = null;
function requestConfirm({ title, message, confirmLabel = "Confirmar", cancelLabel = "Cancelar", danger = true }) {
  return new Promise((resolve) => {
    if (!_showConfirmModal) { resolve(false); return; }
    _showConfirmModal({ title, message, confirmLabel, cancelLabel, danger }, resolve);
  });
}
function ConfirmModal() {
  const [pending, setPending] = useState(null); // { opts, onSubmit }

  useEffect(() => {
    _showConfirmModal = (opts, onSubmit) => setPending({ opts, onSubmit });
    return () => { _showConfirmModal = null; };
  }, []);

  if (!pending) return null;
  const { opts, onSubmit } = pending;
  const submit = (result) => {
    setPending(null);
    onSubmit(result);
  };

  return (
    <div
      style={{
        position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)",
        display: "flex", alignItems: "center", justifyContent: "center",
        zIndex: 1000, padding: 16,
      }}
      onClick={() => submit(false)}
    >
      <div
        style={{
          background: C.panel, border: `1px solid ${opts.danger ? "rgba(226,99,79,0.5)" : C.panelLine}`, borderRadius: 14,
          padding: 20, width: "min(360px, 100%)", boxShadow: "0 12px 32px rgba(0,0,0,0.4)",
        }}
        onClick={(e) => e.stopPropagation()}
      >
        <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
          {opts.danger && <AlertCircle size={18} color={C.loss} />}
          <div style={{ ...displayFont, fontSize: 19, color: C.card }}>{opts.title}</div>
        </div>
        <div style={{ fontSize: 13.5, color: "rgba(244,234,214,0.75)", lineHeight: 1.5, marginBottom: 18 }}>{opts.message}</div>
        <div style={{ display: "flex", gap: 8, justifyContent: "flex-end" }}>
          <GhostBtn onClick={() => submit(false)}>{opts.cancelLabel}</GhostBtn>
          <PrimaryBtn
            onClick={() => submit(true)}
            style={opts.danger ? { background: `linear-gradient(180deg, #e2634f, #c94d3b)`, color: "#fff" } : undefined}
          >
            {opts.confirmLabel}
          </PrimaryBtn>
        </div>
      </div>
    </div>
  );
}

/* Icon choices for player avatars */
const AVATAR_ICONS = [
  "🂡", "🃏", "🎴", "🀄", "🎲",
  "♠️", "♥️", "♦️", "♣️",
  "😎", "🐺", "🦁", "🐯", "🐍", "🍀", "🔥", "⚡", "🎯", "🥇", "👑", "🏆",
  "⚽", "🏀", "🏈", "⚾", "🎾", "🏐", "🏉", "🎱",
  "🍺", "🥂", "🥃", "🍻", "🧉",
  "🍔", "🥩", "🍕", "🌮", "🍗", "🥓", "🍤",
];

function readFileAsDataURL(file) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result);
    r.onerror = () => reject(new Error("no se pudo leer el archivo"));
    r.readAsDataURL(file);
  });
}

function Avatar({ player, size = 30 }) {
  const av = player?.avatar;
  const base = { width: size, height: size, borderRadius: 99, flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center", overflow: "hidden" };
  if (av?.type === "photo" && av.value) {
    return <img src={av.value} alt={player.name} style={{ ...base, objectFit: "cover", border: `1px solid ${C.panelLine}` }} />;
  }
  if (av?.type === "icon" && av.value) {
    return <div style={{ ...base, background: "rgba(216,173,63,0.18)", fontSize: size * 0.55 }}>{av.value}</div>;
  }
  return (
    <div style={{ ...base, background: C.gold, color: C.ink, fontWeight: 800, fontSize: size * 0.42, ...bodyFont }}>
      {(player?.name || "?").slice(0, 1).toUpperCase()}
    </div>
  );
}

/* Break an amount (assumed multiple of 100) into bills of 500/200/100 */
function billsFor(amount) {
  let a = Math.max(0, Math.round(amount / 100) * 100);
  const b500 = Math.floor(a / 500); a -= b500 * 500;
  const b200 = Math.floor(a / 200); a -= b200 * 200;
  const b100 = Math.round(a / 100);
  return { 500: b500, 200: b200, 100: b100 };
}
function billsLabel(bd) {
  const parts = [];
  if (bd[500]) parts.push(`${bd[500]}×$500`);
  if (bd[200]) parts.push(`${bd[200]}×$200`);
  if (bd[100]) parts.push(`${bd[100]}×$100`);
  return parts.length ? parts.join(" + ") : "—";
}

/* ----------------------------------------------------------------------
   SETTLEMENT ENGINE
---------------------------------------------------------------------- */
function computeSettlement(game, roster) {
  const players = game.playerIds.map((pid) => {
    const p = roster.find((r) => r.id === pid);
    const purchases = game.purchases.filter((pu) => pu.playerId === pid);
    const cashAmount = purchases.filter((pu) => pu.type === "cash").reduce((s, pu) => s + pu.amount, 0);
    const virtualAmount = purchases.filter((pu) => pu.type === "virtual").reduce((s, pu) => s + pu.amount, 0);
    const totalBuyIn = cashAmount + virtualAmount;
    const cashOut = Number(game.finalChips[pid]) || 0; // lo reportado en fichas
    const balance = round1(cashOut - totalBuyIn); // balance neto informativo (gana/pierde en total)

    // Regla: el cash out primero salda el buy-in virtual. Lo que sobra de eso
    // ("netClaim") es lo que el jugador realmente puede reclamar del pozo de
    // cash real — no el balance total. Si netClaim <= 0, ni siquiera alcanzó
    // para saldar el virtual, y esa diferencia se debe por transferencia.
    const netClaim = round1(cashOut - virtualAmount);

    let pagoCash = 0;
    let pagoTransfer = 0;
    if (netClaim > 0) {
      // Puede cobrar en cash hasta lo que él mismo puso en cash; el resto del
      // reclamo (si lo hay) se cobra por transferencia.
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

  // Todo el cash disponible debe entregarse. Como el tope de cada jugador es
  // "hasta lo que él mismo puso en cash", puede quedar un remanente sin
  // asignar si alguien reclama más de lo que puso en cash. Ese remanente se
  // le da al jugador con mayor balance positivo, convirtiendo esa parte de
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
  // deben, para minimizar el número de transacciones. Las transferencias son
  // un circuito cerrado entre jugadores (no involucran el rake, que ya se
  // descontó del pozo de cash), así que en teoría suman exactamente 0.
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

/* ----------------------------------------------------------------------
   ROOT APP
---------------------------------------------------------------------- */
export default function PokerLedger() {
  const [loading, setLoading] = useState(true);
  const [roster, setRoster] = useState([]);
  const [games, setGames] = useState([]);
  const [activeGame, setActiveGame] = useState(null);
  const [tab, setTab] = useState("partida");
  const [adminPassword, setAdminPassword] = useState("");

  // Si hay una partida en curso, la pestaña "Jugadores" queda oculta (no se debe
  // editar el roster a mitad de una partida); si alguien estaba justo ahí cuando
  // arrancó una partida, lo mandamos de vuelta a la pestaña de la partida.
  useEffect(() => {
    if (activeGame && !activeGame.finished && tab === "jugadores") setTab("partida");
  }, [activeGame, tab]);

  useEffect(() => {
    (async () => {
      const [r, g, a, pw] = await Promise.all([
        loadKey(KEYS.roster, []),
        loadKey(KEYS.games, []),
        loadKey(KEYS.active, null),
        loadKey(KEYS.adminPassword, ""),
      ]);
      setRoster(r); setGames(g); setActiveGame(a); setAdminPassword(pw || "");
      setLoading(false);
    })();
  }, []);

  // Evita que, apenas termina la carga inicial, se dispare un guardado con
  // los datos recién leídos (o con un estado intermedio) y se sobreescriba
  // el Excel innecesariamente. Solo guardamos ante cambios reales, hechos
  // por el usuario después de que la carga inicial ya terminó.
  const skipNextRosterSave = useRef(true);
  const skipNextGamesSave = useRef(true);
  const skipNextActiveSave = useRef(true);

  useEffect(() => {
    if (loading) return;
    if (skipNextRosterSave.current) { skipNextRosterSave.current = false; return; }
    saveKey(KEYS.roster, roster);
  }, [roster, loading]);
  useEffect(() => {
    if (loading) return;
    if (skipNextGamesSave.current) { skipNextGamesSave.current = false; return; }
    saveKey(KEYS.games, games);
  }, [games, loading]);
  useEffect(() => {
    if (loading) return;
    if (skipNextActiveSave.current) { skipNextActiveSave.current = false; return; }
    saveKey(KEYS.active, activeGame);
  }, [activeGame, loading]);

  const playerStats = useCallback(
    (playerId) => {
      let played = 0, wins = 0, balance = 0, cashInTotal = 0, virtualInTotal = 0;
      games.forEach((g) => {
        const res = computeSettlement(g, roster).players.find((p) => p.playerId === playerId);
        if (res) {
          played++; balance += res.balance;
          if (res.balance > 0) wins++;
          cashInTotal += res.cashAmount; virtualInTotal += res.virtualAmount;
        }
      });
      return { played, wins, balance, cashInTotal, virtualInTotal };
    },
    [games, roster]
  );

  if (loading) {
    return (
      <div style={{ background: C.felt, minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <div style={{ ...displayFont, color: C.goldSoft, fontSize: 28 }}>Repartiendo cartas…</div>
      </div>
    );
  }

  return (
    <div style={{ minHeight: "100vh", background: `radial-gradient(ellipse at top, ${C.panel} 0%, ${C.felt} 45%, ${C.feltDeep} 100%)`, ...bodyFont }}>
      <style>{`
        ${FONT_IMPORT}
        * { box-sizing: border-box; }
        body { margin:0; }
        input[type=number]::-webkit-inner-spin-button { opacity: 1; }
        .felt-line { border-color: ${C.feltLine}; }
        ::selection { background: ${C.gold}; color: ${C.ink}; }
        button:focus-visible, input:focus-visible, select:focus-visible {
          outline: 2px solid ${C.gold}; outline-offset: 2px;
        }
        .scrollbar-thin::-webkit-scrollbar { width: 6px; height:6px; }
        .scrollbar-thin::-webkit-scrollbar-thumb { background: ${C.panelLine}; border-radius: 3px; }
      `}</style>

      <Header tab={tab} setTab={setTab} hasActive={!!activeGame && !activeGame.finished} />

      <main style={{ maxWidth: 980, margin: "0 auto", padding: "18px 14px 60px" }}>
        {!!activeGame && !activeGame.finished && (
          <div style={{
            display: "flex", alignItems: "center", gap: 8, marginBottom: 14,
            background: "rgba(216,173,63,0.14)", border: `1px solid ${C.gold}`,
            borderRadius: 10, padding: "9px 12px",
          }}>
            <span style={{ width: 8, height: 8, borderRadius: 99, background: C.win, flexShrink: 0, boxShadow: "0 0 0 3px rgba(63,191,114,0.25)" }} />
            <span style={{ ...displayFont, fontSize: 15, color: C.goldSoft, letterSpacing: "0.04em" }}>Jugada en Curso</span>
          </div>
        )}
        {tab === "jugadores" && !(!!activeGame && !activeGame.finished) && (
          <PlayersTab roster={roster} setRoster={setRoster} playerStats={playerStats} adminPassword={adminPassword} />
        )}
        {tab === "partida" && (
          <GameTab
            roster={roster}
            activeGame={activeGame}
            setActiveGame={setActiveGame}
            games={games}
            setGames={setGames}
          />
        )}
        {tab === "historial" && <HistoryTab games={games} roster={roster} setGames={setGames} adminPassword={adminPassword} />}
      </main>

      <AdminPasswordModal />
      <ConfirmModal />
    </div>
  );
}

/* ----------------------------------------------------------------------
   HEADER / TABS
---------------------------------------------------------------------- */
function Header({ tab, setTab, hasActive }) {
  const tabs = hasActive
    ? [
        { id: "partida", label: "Estatus jugada", icon: Activity },
        { id: "historial", label: "Información histórica", icon: History },
      ]
    : [
        { id: "partida", label: "Partida", icon: Flame },
        { id: "jugadores", label: "Jugadores", icon: Users },
        { id: "historial", label: "Historial", icon: History },
      ];
  return (
    <header style={{ borderBottom: `1px solid ${C.panelLine}`, background: "rgba(0,0,0,0.15)", position: "sticky", top: 0, zIndex: 20, backdropFilter: "blur(6px)" }}>
      <div style={{ maxWidth: 980, margin: "0 auto", padding: "14px 14px 0" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
          <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
            <span style={{ ...displayFont, fontSize: 30, color: C.gold, lineHeight: 1 }}>BARILOCHE</span>
            <span style={{ ...bodyFont, fontSize: 12, color: "rgba(240,216,136,0.55)", letterSpacing: "0.14em", textTransform: "uppercase" }}>
              registro de poker
            </span>
          </div>
          {/* Página aparte para el organizador: no es un tab más, así los
              jugadores no se topan de casualidad con rachas o comparativas. */}
          <a
            href="/dashboard.html"
            target="_blank"
            rel="noopener noreferrer"
            title="Dashboard del organizador"
            style={{
              display: "flex", alignItems: "center", gap: 6, textDecoration: "none",
              border: `1px solid ${C.panelLine}`, borderRadius: 8, padding: "6px 10px",
              color: "rgba(244,234,214,0.65)", fontSize: 12, fontWeight: 600, ...bodyFont,
            }}
          >
            <BarChart3 size={13} /> Organizador
          </a>
        </div>
        <div style={{ display: "flex", gap: 4, marginTop: 12 }}>
          {tabs.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 6,
                  padding: "8px 14px", border: "none", cursor: "pointer",
                  background: "transparent",
                  color: active ? C.goldSoft : "rgba(244,234,214,0.55)",
                  borderBottom: active ? `2px solid ${C.gold}` : "2px solid transparent",
                  ...bodyFont, fontWeight: 600, fontSize: 13.5,
                }}
              >
                <Icon size={15} />
                {t.label}
                {t.id === "partida" && hasActive && (
                  <span style={{ width: 6, height: 6, borderRadius: 99, background: C.win, display: "inline-block" }} />
                )}
              </button>
            );
          })}
        </div>
      </div>
    </header>
  );
}

/* ----------------------------------------------------------------------
   REUSABLE UI
---------------------------------------------------------------------- */
function Panel({ children, style }) {
  return (
    <div style={{ background: C.panel, border: `1px solid ${C.panelLine}`, borderRadius: 14, padding: 16, ...style }}>
      {children}
    </div>
  );
}
function SectionTitle({ icon: Icon, children, right }) {
  return (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, color: C.goldSoft }}>
        {Icon && <Icon size={16} />}
        <span style={{ ...displayFont, fontSize: 19, letterSpacing: "0.06em" }}>{children}</span>
      </div>
      {right}
    </div>
  );
}
function Field({ label, children }) {
  return (
    <label style={{ display: "block" }}>
      <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.6)", marginBottom: 5, textTransform: "uppercase", letterSpacing: "0.06em" }}>{label}</div>
      {children}
    </label>
  );
}
const inputStyle = {
  width: "100%", background: "rgba(0,0,0,0.25)", border: `1px solid ${C.panelLine}`,
  borderRadius: 8, padding: "9px 10px", color: C.card, fontSize: 14.5, ...bodyFont,
};
function PrimaryBtn({ children, onClick, disabled, style, icon: Icon }) {
  return (
    <button
      onClick={onClick} disabled={disabled}
      style={{
        display: "inline-flex", alignItems: "center", gap: 7, justifyContent: "center",
        background: disabled ? "rgba(212,175,55,0.25)" : `linear-gradient(180deg, ${C.goldSoft}, ${C.gold})`,
        color: C.ink, border: "none", borderRadius: 9, padding: "10px 16px",
        fontWeight: 700, fontSize: 14, cursor: disabled ? "not-allowed" : "pointer", ...bodyFont, ...style,
      }}
    >
      {Icon && <Icon size={15} />}
      {children}
    </button>
  );
}
function GhostBtn({ children, onClick, style, icon: Icon, color }) {
  return (
    <button
      onClick={onClick}
      style={{
        display: "inline-flex", alignItems: "center", gap: 6, justifyContent: "center",
        background: "transparent", color: color || "rgba(244,234,214,0.75)",
        border: `1px solid ${C.panelLine}`, borderRadius: 8, padding: "7px 11px",
        fontWeight: 600, fontSize: 12.5, cursor: "pointer", ...bodyFont, ...style,
      }}
    >
      {Icon && <Icon size={13} />}
      {children}
    </button>
  );
}
function Badge({ children, tone }) {
  const bg = tone === "cash" ? C.cash : tone === "virtual" ? C.virtual : "rgba(255,255,255,0.1)";
  return (
    <span style={{ background: bg, color: "#fff", padding: "3px 8px", borderRadius: 99, fontSize: 11.5, fontWeight: 700, ...monoFont }}>
      {children}
    </span>
  );
}

/* ----------------------------------------------------------------------
   PLAYERS TAB
---------------------------------------------------------------------- */
function PlayersTab({ roster, setRoster, playerStats, adminPassword }) {
  const [selId, setSelId] = useState("");
  const [newName, setNewName] = useState("");
  const [newAvatar, setNewAvatar] = useState(null);
  const [editName, setEditName] = useState("");
  const [editAvatar, setEditAvatar] = useState(null);

  const sel = roster.find((p) => p.id === selId) || null;
  useEffect(() => { setEditName(sel ? sel.name : ""); setEditAvatar(sel ? sel.avatar || null : null); }, [selId]); // eslint-disable-line

  const addPlayer = () => {
    const name = newName.trim();
    if (!name) return;
    const p = { id: uid(), name, active: true, avatar: newAvatar, createdAt: Date.now() };
    setRoster((r) => [...r, p]);
    setNewName(""); setNewAvatar(null);
  };
  const saveEdit = () => {
    if (!sel || !editName.trim()) return;
    setRoster((r) => r.map((p) => (p.id === sel.id ? { ...p, name: editName.trim(), avatar: editAvatar } : p)));
  };
  const toggleActive = () => {
    if (!sel) return;
    setRoster((r) => r.map((p) => (p.id === sel.id ? { ...p, active: !p.active } : p)));
  };
  const removePlayer = async () => {
    if (!sel) return;
    if (!(await requestAdminPassword(adminPassword, "eliminar jugadores"))) return;
    if (!confirm(`¿Eliminar a ${sel.name} del roster? Sus estadísticas históricas se conservarán en partidas guardadas, pero dejará de aparecer en la lista.`)) return;
    setRoster((r) => r.filter((p) => p.id !== sel.id));
    setSelId("");
  };

  const sorted = [...roster].sort((a, b) => a.name.localeCompare(b.name));

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <Panel>
        <SectionTitle icon={UserPlus}>Alta de jugador</SectionTitle>
        <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
          <input
            style={inputStyle} placeholder="Nombre del jugador" value={newName}
            onChange={(e) => setNewName(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && addPlayer()}
          />
          <PrimaryBtn onClick={addPlayer} icon={Plus}>Agregar</PrimaryBtn>
        </div>
        <AvatarPicker avatar={newAvatar} setAvatar={setNewAvatar} previewName={newName} />
      </Panel>

      <Panel>
        <SectionTitle icon={Pencil}>Editar / dar de baja</SectionTitle>
        <Field label="Seleccionar jugador">
          <select style={{ ...inputStyle, appearance: "auto" }} value={selId} onChange={(e) => setSelId(e.target.value)}>
            <option value="">— elegir del combo —</option>
            {sorted.map((p) => (
              <option key={p.id} value={p.id}>{p.name}{!p.active ? " (baja)" : ""}</option>
            ))}
          </select>
        </Field>

        {sel && (
          <div style={{ marginTop: 14, display: "grid", gap: 10 }}>
            <div style={{ display: "flex", gap: 8 }}>
              <input style={inputStyle} value={editName} onChange={(e) => setEditName(e.target.value)} />
              <GhostBtn onClick={saveEdit} icon={Check} color={C.win}>Guardar</GhostBtn>
            </div>
            <AvatarPicker avatar={editAvatar} setAvatar={setEditAvatar} previewName={editName} />
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <GhostBtn onClick={toggleActive} icon={sel.active ? UserX : UserCheck} color={sel.active ? C.loss : C.win}>
                {sel.active ? "Dar de baja" : "Reactivar"}
              </GhostBtn>
              <GhostBtn onClick={removePlayer} icon={Trash2} color={C.loss}>Eliminar del roster</GhostBtn>
            </div>
            <PlayerStatRow stats={playerStats(sel.id)} />
          </div>
        )}
      </Panel>

      <Panel>
        <SectionTitle icon={Users} right={<span style={{ color: "rgba(244,234,214,0.5)", fontSize: 12 }}>{roster.filter(p=>p.active).length} activos</span>}>
          Roster completo
        </SectionTitle>
        <div style={{ display: "grid", gap: 8 }}>
          {sorted.length === 0 && <Empty>Aún no hay jugadores. Agrega el primero arriba.</Empty>}
          {sorted.map((p) => {
            const st = playerStats(p.id);
            return (
              <div key={p.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: "10px 12px", opacity: p.active ? 1 : 0.5 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                  <Avatar player={p} size={30} />
                  <div>
                    <div style={{ color: C.card, fontWeight: 600, fontSize: 14 }}>{p.name}{!p.active && <span style={{ color: C.loss, fontSize: 11, marginLeft: 6 }}>BAJA</span>}</div>
                    <div style={{ color: "rgba(244,234,214,0.55)", fontSize: 11.5, ...monoFont }}>{st.played} partidas</div>
                  </div>
                </div>
              </div>
            );
          })}
        </div>
      </Panel>
    </div>
  );
}
function AvatarPicker({ avatar, setAvatar, previewName }) {
  const [customEmoji, setCustomEmoji] = useState("");
  const onPhoto = async (e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    try {
      const dataUrl = await readFileAsDataURL(file);
      setAvatar({ type: "photo", value: dataUrl });
    } catch {
      alert("No se pudo cargar la imagen.");
    }
    e.target.value = "";
  };
  const useCustomEmoji = () => {
    const val = customEmoji.trim();
    if (!val) return;
    setAvatar({ type: "icon", value: val });
    setCustomEmoji("");
  };
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
      <Avatar player={{ name: previewName, avatar }} size={44} />
      <div style={{ display: "flex", gap: 5, flexWrap: "wrap", maxWidth: 280 }}>
        {AVATAR_ICONS.map((ic) => (
          <button key={ic} onClick={() => setAvatar({ type: "icon", value: ic })}
            style={{
              width: 28, height: 28, borderRadius: 7, cursor: "pointer", fontSize: 15,
              background: avatar?.type === "icon" && avatar.value === ic ? "rgba(216,173,63,0.35)" : "rgba(0,0,0,0.2)",
              border: `1px solid ${avatar?.type === "icon" && avatar.value === ic ? C.gold : C.panelLine}`,
            }}>{ic}</button>
        ))}
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <input
          type="text"
          inputMode="text"
          placeholder="😊 tu emoji"
          value={customEmoji}
          maxLength={8}
          onChange={(e) => setCustomEmoji(e.target.value)}
          onKeyDown={(e) => e.key === "Enter" && useCustomEmoji()}
          style={{ ...inputStyle, width: 96, textAlign: "center", fontSize: 17, padding: "6px 6px" }}
        />
        <GhostBtn onClick={useCustomEmoji} icon={Check} color={C.win}>Usar</GhostBtn>
      </div>
      <label style={{ cursor: "pointer" }}>
        <span style={{ display: "inline-flex", alignItems: "center", gap: 5, background: "transparent", border: `1px solid ${C.panelLine}`, borderRadius: 8, padding: "6px 10px", fontSize: 12, color: "rgba(244,234,214,0.75)", ...bodyFont, fontWeight: 600 }}>
          Subir foto
        </span>
        <input type="file" accept="image/*" onChange={onPhoto} style={{ display: "none" }} />
      </label>
      {avatar && (
        <GhostBtn onClick={() => setAvatar(null)} icon={X} color={C.loss}>Quitar</GhostBtn>
      )}
    </div>
  );
}
function PlayerStatRow({ stats }) {
  return (
    <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginTop: 4 }}>
      <MiniStat label="Partidas" value={stats.played} />
      <MiniStat label="Ganadas" value={stats.wins} />
      <MiniStat label="Balance histórico" value={money(stats.balance)} tone={stats.balance >= 0 ? "win" : "loss"} />
    </div>
  );
}
function MiniStat({ label, value, tone }) {
  const color = tone === "win" ? C.win : tone === "loss" ? C.loss : C.card;
  return (
    <div style={{ background: "rgba(0,0,0,0.2)", borderRadius: 8, padding: "8px 10px", textAlign: "center" }}>
      <div style={{ fontSize: 10.5, color: "rgba(244,234,214,0.55)", textTransform: "uppercase" }}>{label}</div>
      <div style={{ ...monoFont, fontWeight: 700, fontSize: 15, color }}>{value}</div>
    </div>
  );
}
function Empty({ children }) {
  return (
    <div style={{ textAlign: "center", padding: "22px 10px", color: "rgba(244,234,214,0.45)", fontSize: 13.5 }}>
      {children}
    </div>
  );
}

/* ----------------------------------------------------------------------
   GAME TAB
---------------------------------------------------------------------- */
function GameTab({ roster, activeGame, setActiveGame, games, setGames }) {
  if (!activeGame) return <NewGameSetup roster={roster} setActiveGame={setActiveGame} />;
  if (activeGame.finished) return <FinalizedGame game={activeGame} roster={roster} onClose={() => setActiveGame(null)} setActiveGame={setActiveGame} setGames={setGames} />;
  return <ActiveGameScreen game={activeGame} setGame={setActiveGame} roster={roster} setGames={setGames} />;
}

function NewGameSetup({ roster, setActiveGame }) {
  const active = roster.filter((p) => p.active);
  const [date, setDate] = useState(todayISO());
  const [loteValue, setLoteValue] = useState(1000);
  const [rake, setRake] = useState(1500); // arranca en 1500, pero se puede editar
  const [selected, setSelected] = useState([]);
  const [hostId, setHostId] = useState("");

  const toggle = (id) => {
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
    if (hostId === id) setHostId("");
  };

  const start = () => {
    if (selected.length < 2 || !loteValue || !hostId) return;
    setActiveGame({
      id: uid(), date, loteValue: Number(loteValue), rake: Number(rake) || 0,
      playerIds: selected, hostId, startedAt: Date.now(),
      purchases: [],
      dinner: { total: 0, sidesFee: 50, waiter: 50, alcoholFee: 50, alcohol: {}, paid: {}, paymentMethod: {} },
      dinnerSetupDone: false, // fuerza a pasar por la pantalla de cena antes de comprar lotes, una sola vez
      finalChips: {}, finished: false, results: null,
    });
  };

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <Panel>
        <SectionTitle icon={Play}>Nueva partida</SectionTitle>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="Fecha">
            <input type="date" style={inputStyle} value={date} onChange={(e) => setDate(e.target.value)} />
          </Field>
          <Field label="Valor del lote">
            <input type="number" min="0" style={inputStyle} value={loteValue} onChange={(e) => setLoteValue(e.target.value)} onFocus={(e) => e.target.select()} />
          </Field>
        </div>
        <div style={{ marginTop: 10 }}>
          <Field label="Rake">
            <input type="number" min="0" style={inputStyle} value={rake === 0 ? "" : rake} onChange={(e) => setRake(e.target.value === "" ? 0 : Number(e.target.value))} onFocus={(e) => e.target.select()} />
          </Field>
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={Users} right={<span style={{ fontSize: 12, color: "rgba(244,234,214,0.5)" }}>{selected.length} seleccionados</span>}>
          Jugadores en la mesa
        </SectionTitle>
        {active.length === 0 ? (
          <Empty>No hay jugadores activos. Ve a la pestaña Jugadores para dar de alta.</Empty>
        ) : (
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {active.map((p) => {
              const on = selected.includes(p.id);
              return (
                <button key={p.id} onClick={() => toggle(p.id)}
                  style={{
                    display: "flex", alignItems: "center", gap: 8, textAlign: "left",
                    background: on ? "rgba(216,173,63,0.16)" : "rgba(0,0,0,0.18)",
                    border: `1px solid ${on ? C.gold : C.panelLine}`, borderRadius: 9,
                    padding: "9px 10px", cursor: "pointer", ...bodyFont,
                  }}>
                  <div style={{ width: 18, height: 18, borderRadius: 5, border: `1.5px solid ${on ? C.gold : "rgba(244,234,214,0.4)"}`, background: on ? C.gold : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    {on && <Check size={12} color={C.ink} />}
                  </div>
                  <Avatar player={p} size={22} />
                  <span style={{ color: C.card, fontSize: 13.5, fontWeight: 600 }}>{p.name}</span>
                </button>
              );
            })}
          </div>
        )}
      </Panel>

      {selected.length >= 1 && (
        <Panel>
          <SectionTitle icon={Crown}>Host / administrador de la partida</SectionTitle>
          <div style={{ color: "rgba(244,234,214,0.55)", fontSize: 12, marginBottom: 10 }}>
            Selecciona un único jugador que administrará la partida.
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
            {selected.map((id) => {
              const p = active.find((x) => x.id === id) || roster.find((x) => x.id === id);
              if (!p) return null;
              const on = hostId === id;
              return (
                <button key={id} onClick={() => setHostId(id)}
                  style={{
                    display: "flex", alignItems: "center", gap: 8, textAlign: "left",
                    background: on ? "rgba(216,173,63,0.22)" : "rgba(0,0,0,0.18)",
                    border: `1px solid ${on ? C.gold : C.panelLine}`, borderRadius: 9,
                    padding: "9px 10px", cursor: "pointer", ...bodyFont,
                  }}>
                  <div style={{ width: 16, height: 16, borderRadius: 99, border: `1.5px solid ${on ? C.gold : "rgba(244,234,214,0.4)"}`, background: on ? C.gold : "transparent", flexShrink: 0 }} />
                  <Avatar player={p} size={22} />
                  <span style={{ color: C.card, fontSize: 13.5, fontWeight: 600 }}>{p.name}</span>
                  {on && <Crown size={13} color={C.gold} style={{ marginLeft: "auto" }} />}
                </button>
              );
            })}
          </div>
        </Panel>
      )}

      <PrimaryBtn onClick={start} disabled={selected.length < 2 || !loteValue || !hostId} icon={Play} style={{ padding: "13px 18px", fontSize: 15 }}>
        Iniciar partida
      </PrimaryBtn>
      {selected.length < 2 && <div style={{ color: "rgba(244,234,214,0.5)", fontSize: 12.5, textAlign: "center" }}>Selecciona al menos 2 jugadores.</div>}
      {selected.length >= 2 && !hostId && <div style={{ color: "rgba(244,234,214,0.5)", fontSize: 12.5, textAlign: "center" }}>Selecciona quién será el host.</div>}
    </div>
  );
}

/* ----- Active game: buy-ins, dinner, finalize ----- */
function useTicker(intervalMs) {
  const [, tick] = useState(0);
  useEffect(() => {
    const t = setInterval(() => tick((n) => n + 1), intervalMs);
    return () => clearInterval(t);
  }, [intervalMs]);
}
function formatElapsed(ms) {
  const totalMin = Math.max(0, Math.floor(ms / 60000));
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  const pad = (n) => String(n).padStart(2, "0");
  return `${pad(h)}:${pad(m)}`;
}
function formatClock(ts) {
  if (!ts) return "—";
  return new Date(ts).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit", hour12: false });
}
// Recuadro combinado: Rake y Número de jugadores juntos en una sola tarjeta.
function RakeAndPlayersBox({ rake, playerCount }) {
  return (
    <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "8px 6px", display: "flex" }}>
      <div style={{ flex: 1, textAlign: "center", borderRight: `1px solid ${C.panelLine}` }}>
        <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase" }}>Rake</div>
        <div style={{ ...monoFont, fontWeight: 700, fontSize: 14.5, color: C.goldSoft }}>{money(rake)}</div>
      </div>
      <div style={{ flex: 1, textAlign: "center" }}>
        <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase" }}>Jugadores</div>
        <div style={{ ...monoFont, fontWeight: 700, fontSize: 14.5, color: C.goldSoft }}>{playerCount}</div>
      </div>
    </div>
  );
}
function GameStatusBlock({ game, players, totals }) {
  useTicker(1000);
  const lotesTotal = game.purchases.reduce((s, p) => s + p.lotes, 0);
  const lotesProm = players.length ? round1(lotesTotal / players.length) : 0;
  const elapsed = game.startedAt ? Date.now() - game.startedAt : null;

  const perPlayer = players.map((p) => {
    const cash = game.purchases.filter((pu) => pu.playerId === p.id && pu.type === "cash").reduce((s, pu) => s + pu.amount, 0);
    const virtual = game.purchases.filter((pu) => pu.playerId === p.id && pu.type === "virtual").reduce((s, pu) => s + pu.amount, 0);
    return { player: p, cash, virtual, total: cash + virtual };
  }).sort((a, b) => b.total - a.total);

  return (
    <Panel>
      <SectionTitle icon={Activity}>Estatus de la jugada</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 10 }}>
        <RakeAndPlayersBox rake={game.rake} playerCount={players.length} />
        <ScoreBox label="Hora inicio" value={formatClock(game.startedAt)} />
        <ScoreBox label="Tiempo jugado" value={elapsed == null ? "—" : formatElapsed(elapsed)} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 12 }}>
        <ScoreBox label="Total cash" value={money(totals.cash)} tone="cash" />
        <ScoreBox label="Total virtual" value={money(totals.virtual)} tone="virtual" />
        <ScoreBox label="Lotes prom./jugador" value={lotesProm} />
      </div>
      <div style={{ display: "flex", gap: 12, marginBottom: 6, fontSize: 10, color: "rgba(244,234,214,0.4)", justifyContent: "flex-end" }}>
        <span>cash</span><span>virtual</span><span>total</span>
      </div>
      <div style={{ display: "grid", gap: 6 }}>
        {perPlayer.map((row) => (
          <div key={row.player.id} style={{ background: "rgba(0,0,0,0.18)", borderRadius: 8, padding: "8px 10px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 5 }}>
              <Avatar player={row.player} size={22} />
              <span style={{ color: C.card, fontWeight: 700, fontSize: 13.5, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.player.name}</span>
            </div>
            <div style={{ display: "flex", gap: 14 }}>
              <span style={{ ...monoFont, fontSize: 12.5, color: C.cash, fontWeight: 700 }}>{money(row.cash)}</span>
              <span style={{ ...monoFont, fontSize: 12.5, color: C.virtual, fontWeight: 700 }}>{money(row.virtual)}</span>
              <span style={{ ...monoFont, fontSize: 12.5, color: C.goldSoft, fontWeight: 800 }}>{money(row.total)}</span>
            </div>
          </div>
        ))}
      </div>
    </Panel>
  );
}

function ActiveGameScreen({ game, setGame, roster, setGames }) {
  const [finalizing, setFinalizing] = useState(false);
  const [view, setView] = useState(() => (game.dinnerSetupDone ? "purchase" : "dinner")); // "config" | "purchase" | "dinner"

  const players = game.playerIds.map((id) => roster.find((r) => r.id === id)).filter(Boolean);
  const availableToAdd = roster.filter((p) => p.active && !game.playerIds.includes(p.id));

  // update acepta un objeto (mezcla directa) o una función (g) => patch, que
  // recibe el estado MÁS RECIENTE del juego. Usar la forma función evita
  // "carreras" cuando se disparan varios cambios rápido seguidos (por
  // ejemplo, tocar +Cash varias veces, o editar dos campos de la cena
  // rápido): cada patch se arma sobre el estado real, no sobre una copia
  // vieja capturada en el momento del render.
  const update = (patch) =>
    setGame((g) => ({ ...g, ...(typeof patch === "function" ? patch(g) : patch) }));

  const addPurchase = (playerId, type) => {
    update((g) => {
      const entry = { id: uid(), playerId, type, lotes: 1, amount: g.loteValue, ts: Date.now() };
      return { purchases: [...g.purchases, entry] };
    });
  };
  const removeLastPurchase = (playerId, type) => {
    update((g) => {
      const entries = g.purchases.filter((p) => p.playerId === playerId && p.type === type);
      if (entries.length === 0) return {};
      const last = entries[entries.length - 1];
      return { purchases: g.purchases.filter((p) => p.id !== last.id) };
    });
  };

  const totals = useMemo(() => {
    let cash = 0, virtual = 0;
    game.purchases.forEach((p) => { if (p.type === "cash") cash += p.amount; else virtual += p.amount; });
    return { cash, virtual };
  }, [game.purchases]);
  const unpaidCount = players.filter((p) => !game.dinner.paid?.[p.id]).length;

  const setRake = (v) => update({ rake: Number(v) || 0 });
  const setLote = (v) => update({ loteValue: Number(v) || 0 });
  const rakeLocked = Object.keys(game.finalChips || {}).length > 0;

  const addPlayerToGame = (id) => {
    if (game.playerIds.includes(id)) return;
    update((g) => ({ playerIds: [...g.playerIds, id] }));
  };
  const removePlayerFromGame = (id) => {
    const hasPurchases = game.purchases.some((p) => p.playerId === id);
    const name = roster.find((r) => r.id === id)?.name || "este jugador";
    if (hasPurchases) {
      if (!confirm(`⚠️ ${name} ya tiene lotes comprados registrados en esta partida. Si lo quitas, esas compras se van a eliminar también. ¿Continuar?`)) return;
    } else if (!confirm(`¿Quitar a ${name} de esta partida?`)) {
      return;
    }
    update((g) => ({
      playerIds: g.playerIds.filter((pid) => pid !== id),
      purchases: g.purchases.filter((p) => p.playerId !== id),
      hostId: g.hostId === id ? "" : g.hostId,
    }));
  };

  if (finalizing) {
    return (
      <FinalizeGame
        game={game} roster={roster} update={update}
        onBack={() => setFinalizing(false)}
        onConfirm={(finalChips) => {
          const g2 = { ...game, finalChips };
          const results = computeSettlement(g2, roster);
          const finished = { ...g2, finished: true, results };
          setGame(finished);
          // Si esta partida ya estaba guardada (por ejemplo, volviste a
          // editarla desde "Resultados" y la volviste a cerrar), reemplazamos
          // la entrada vieja en vez de duplicarla.
          setGames((gs) => [...gs.filter((x) => x.id !== finished.id), finished]);
        }}
      />
    );
  }

  if (view === "dinner") {
    const goToPurchase = () => {
      update({ dinnerSetupDone: true });
      setView("purchase");
    };
    return (
      <div style={{ display: "grid", gap: 16 }}>
        <GameStatusBlock game={game} players={players} totals={totals} />
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, marginBottom: 4 }}>
            <SectionTitle icon={UtensilsCrossed}>Cena y servicio</SectionTitle>
          </div>
          <DinnerSection game={game} players={players} update={update} />
        </Panel>
        <PrimaryBtn onClick={goToPurchase} icon={Banknote} style={{ padding: "13px 18px", fontSize: 15 }}>
          {game.dinnerSetupDone ? "Volver a compra de lotes" : "Continuar a compra de lotes"}
        </PrimaryBtn>
      </div>
    );
  }

  if (view === "config") {
    return (
      <div style={{ display: "grid", gap: 16 }}>
        <GameStatusBlock game={game} players={players} totals={totals} />
        <Panel>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
            <div>
              <div style={{ ...displayFont, fontSize: 24, color: C.goldSoft }}>Configuración de la partida</div>
              <div style={{ color: "rgba(244,234,214,0.55)", fontSize: 12.5, ...monoFont }}>{game.date}</div>
            </div>
            <GhostBtn
              icon={X} color={C.loss}
              onClick={async () => {
                const ok = await requestConfirm({
                  title: "¿Cancelar esta partida?",
                  message: "Se va a perder todo el progreso de esta partida (lotes comprados, cena, configuración). Esta acción no se puede deshacer.",
                  confirmLabel: "Sí, cancelar partida",
                  cancelLabel: "Seguir jugando",
                });
                if (ok) setGame(null);
              }}
            >
              Cancelar partida
            </GhostBtn>
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 14 }}>
            <Field label="Valor de lote"><input type="number" style={inputStyle} value={game.loteValue} onChange={(e) => setLote(e.target.value)} onFocus={(e) => e.target.select()} /></Field>
            <Field label={rakeLocked ? "Rake 🔒" : "Rake"}>
              <input
                type="number" style={{ ...inputStyle, opacity: rakeLocked ? 0.55 : 1, cursor: rakeLocked ? "not-allowed" : "text" }}
                value={game.rake} disabled={rakeLocked}
                onChange={(e) => setRake(e.target.value)} onFocus={(e) => e.target.select()}
                title={rakeLocked ? "El rake queda fijo una vez que se hizo la entrega de fichas." : undefined}
              />
            </Field>
          </div>
        </Panel>

        <Panel>
          <SectionTitle icon={Users}>Jugadores en la mesa ({players.length})</SectionTitle>
          <div style={{ display: "grid", gap: 8 }}>
            {players.map((p) => {
              const hasPurchases = game.purchases.some((pu) => pu.playerId === p.id);
              return (
                <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(0,0,0,0.18)", borderRadius: 9, padding: "8px 10px" }}>
                  <Avatar player={p} size={24} />
                  <span style={{ color: C.card, fontWeight: 600, fontSize: 13.5, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</span>
                  {game.hostId === p.id && <Badge tone="gold"><Crown size={10} /> Host</Badge>}
                  {hasPurchases && <span title="Ya tiene lotes comprados" style={{ fontSize: 10.5, color: "rgba(244,234,214,0.4)" }}>tiene lotes</span>}
                  <button onClick={() => removePlayerFromGame(p.id)} title="Quitar de la partida" style={{ background: "transparent", border: "none", cursor: "pointer", color: "rgba(226,99,79,0.75)", padding: 4, flexShrink: 0 }}>
                    <Trash2 size={15} />
                  </button>
                </div>
              );
            })}
          </div>
          {availableToAdd.length > 0 && (
            <div style={{ marginTop: 12 }}>
              <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.5)", marginBottom: 6 }}>Agregar a la mesa:</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8 }}>
                {availableToAdd.map((p) => (
                  <button key={p.id} onClick={() => addPlayerToGame(p.id)}
                    style={{ display: "flex", alignItems: "center", gap: 8, textAlign: "left", background: "rgba(0,0,0,0.18)", border: `1px solid ${C.panelLine}`, borderRadius: 9, padding: "8px 10px", cursor: "pointer", ...bodyFont }}>
                    <Avatar player={p} size={20} />
                    <span style={{ color: C.card, fontSize: 13 }}>{p.name}</span>
                    <Plus size={13} color={C.goldSoft} style={{ marginLeft: "auto", flexShrink: 0 }} />
                  </button>
                ))}
              </div>
            </div>
          )}
        </Panel>

        <PrimaryBtn onClick={() => setView("purchase")} icon={Banknote} style={{ padding: "13px 18px", fontSize: 15 }}>
          Continuar a compra de lotes
        </PrimaryBtn>
      </div>
    );
  }

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <GameStatusBlock game={game} players={players} totals={totals} />

      <Panel>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 10 }}>
          <div>
            <div style={{ ...displayFont, fontSize: 24, color: C.goldSoft }}>Partida en curso</div>
            <div style={{ color: "rgba(244,234,214,0.55)", fontSize: 12.5, ...monoFont, display: "flex", alignItems: "center", gap: 6, flexWrap: "wrap" }}>
              <span>{game.date} · {players.length} jugadores</span>
              {game.hostId && (
                <span style={{ display: "inline-flex", alignItems: "center", gap: 4, background: "rgba(216,173,63,0.16)", color: C.goldSoft, padding: "2px 7px", borderRadius: 99 }}>
                  <Crown size={11} /> Host: {roster.find((r) => r.id === game.hostId)?.name || "—"}
                </span>
              )}
            </div>
          </div>
          <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
            <GhostBtn icon={UtensilsCrossed} onClick={() => setView("dinner")} color={unpaidCount > 0 ? C.loss : undefined}>
              Cena y servicio{unpaidCount > 0 ? ` (${unpaidCount} sin pagar)` : ""}
            </GhostBtn>
            <GhostBtn icon={Settings} onClick={() => setView("config")}>
              Configuración
            </GhostBtn>
          </div>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginTop: 14 }}>
          <ScoreBox label="Lote" value={money(game.loteValue)} />
          <ScoreBox label="Cash" value={money(totals.cash)} tone="cash" />
          <ScoreBox label="Virtual" value={money(totals.virtual)} tone="virtual" />
          <ScoreBox label="Rake" value={money(game.rake)} />
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={Banknote}>Compra de lotes por jugador</SectionTitle>
        <div style={{ display: "grid", gap: 10 }}>
          {players.map((p) => (
            <PlayerBuyRow
              key={p.id} player={p} game={game} onAdd={addPurchase} onRemoveLast={removeLastPurchase}
              dinnerPaid={!!game.dinner.paid?.[p.id]}
              onGoToDinner={() => setView("dinner")}
            />
          ))}
        </div>
      </Panel>

      <PrimaryBtn onClick={() => setFinalizing(true)} icon={Square} style={{ padding: "13px 18px", fontSize: 15 }}>
        Finalizar partida
      </PrimaryBtn>
    </div>
  );
}

function ScoreBox({ label, value, tone }) {
  const color = tone === "cash" ? C.cash : tone === "virtual" ? C.virtual : C.goldSoft;
  return (
    <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "8px 6px", textAlign: "center" }}>
      <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase" }}>{label}</div>
      <div style={{ ...monoFont, fontWeight: 700, fontSize: 14.5, color }}>{value}</div>
    </div>
  );
}

function buyBtnStyle(color, subtract, disabled) {
  return {
    background: subtract ? "transparent" : color,
    border: `1.5px solid ${color}`,
    color: subtract ? color : "#fff",
    borderRadius: 8,
    padding: "6px 9px",
    display: "flex",
    alignItems: "center",
    gap: 4,
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.35 : 1,
    fontWeight: 700,
    fontSize: 12,
    ...bodyFont,
  };
}

function PlayerBuyRow({ player, game, onAdd, onRemoveLast, dinnerPaid, onGoToDinner }) {
  const entries = game.purchases.filter((p) => p.playerId === player.id);
  const cashLotes = entries.filter((e) => e.type === "cash").reduce((s, e) => s + e.lotes, 0);
  const virtualLotes = entries.filter((e) => e.type === "virtual").reduce((s, e) => s + e.lotes, 0);
  const cashAmount = cashLotes * game.loteValue;
  const virtualAmount = virtualLotes * game.loteValue;
  const blocked = !dinnerPaid;

  return (
    <div style={{ background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Avatar player={player} size={26} />
        <span style={{ color: C.card, fontWeight: 700, fontSize: 14.5 }}>{player.name}</span>
      </div>

      {blocked && (
        <button
          onClick={onGoToDinner}
          style={{
            display: "flex", alignItems: "center", gap: 7, width: "100%", textAlign: "left",
            background: "rgba(226,99,79,0.14)", border: `1px solid ${C.loss}`, borderRadius: 8,
            padding: "8px 10px", marginBottom: 10, cursor: "pointer", color: "#f0c9c2", fontSize: 12, ...bodyFont,
          }}
        >
          <UtensilsCrossed size={14} color={C.loss} style={{ flexShrink: 0 }} />
          Falta confirmar el pago de la cena de {player.name} para poder comprar lotes. Toca para ir a Cena y servicio.
        </button>
      )}

      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
        <button onClick={() => onAdd(player.id, "cash")} disabled={blocked} style={buyBtnStyle(C.cash, false, blocked)}>
          <Plus size={12} /> Cash
        </button>
        <button onClick={() => onRemoveLast(player.id, "cash")} disabled={cashLotes === 0} style={buyBtnStyle(C.cash, true, cashLotes === 0)}>
          <Minus size={12} /> Cash
        </button>
        <span style={{ marginLeft: "auto", ...monoFont, fontSize: 12.5, fontWeight: 700, color: C.cash, whiteSpace: "nowrap" }}>
          {cashLotes} lotes · {money(cashAmount)}
        </span>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
        <button onClick={() => onAdd(player.id, "virtual")} disabled={blocked} style={buyBtnStyle(C.virtual, false, blocked)}>
          <Plus size={12} /> Virtual
        </button>
        <button onClick={() => onRemoveLast(player.id, "virtual")} disabled={virtualLotes === 0} style={buyBtnStyle(C.virtual, true, virtualLotes === 0)}>
          <Minus size={12} /> Virtual
        </button>
        <span style={{ marginLeft: "auto", ...monoFont, fontSize: 12.5, fontWeight: 700, color: C.virtual, whiteSpace: "nowrap" }}>
          {virtualLotes} lotes · {money(virtualAmount)}
        </span>
      </div>
    </div>
  );
}

function DinnerSection({ game, players, update }) {
  const d = game.dinner;
  const sidesFee = d.sidesFee || 0; // compatibilidad con partidas viejas sin este campo
  const setD = (patch) =>
    update((g) => ({
      dinner: { ...g.dinner, ...(typeof patch === "function" ? patch(g.dinner) : patch) },
    }));
  const numPlayers = players.length || 1;

  // Costo base por persona (cena repartida + guarniciones + servicio), sin
  // alcohol porque ese varía según quién bebió. Es el número que se muestra
  // como referencia.
  const costoPorPersona = d.total / numPlayers + sidesFee + d.waiter;
  const nominalTotal = players.reduce((s, p) => {
    const alcohol = !!d.alcohol[p.id];
    return s + (d.total / numPlayers + sidesFee + d.waiter + (alcohol ? d.alcoholFee : 0));
  }, 0);
  const totalRecaudado = players.reduce((s, p) => {
    const alcohol = !!d.alcohol[p.id];
    return s + ceilTo100(d.total / numPlayers + sidesFee + d.waiter + (alcohol ? d.alcoholFee : 0));
  }, 0);
  const redondeoExtra = round1(totalRecaudado - nominalTotal);

  return (
    <div style={{ marginTop: 6 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Monto total de la cena">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.total === 0 ? "" : d.total} onChange={(e) => setD({ total: e.target.value === "" ? 0 : Number(e.target.value) })} onFocus={(e) => e.target.select()} />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(d.total)}</span>
          </div>
        </Field>
        <Field label="Guarniciones y complementos (por jugador)">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={sidesFee === 0 ? "" : sidesFee} onChange={(e) => setD({ sidesFee: e.target.value === "" ? 0 : Number(e.target.value) })} onFocus={(e) => e.target.select()} />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(sidesFee)}</span>
          </div>
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
        <Field label="Servicio de mesero (por jugador)">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.waiter === 0 ? "" : d.waiter} onChange={(e) => setD({ waiter: e.target.value === "" ? 0 : Number(e.target.value) })} onFocus={(e) => e.target.select()} />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(d.waiter)}</span>
          </div>
        </Field>
        <Field label="Cargo extra de alcohol (a quienes bebieron)">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.alcoholFee === 0 ? "" : d.alcoholFee} onChange={(e) => setD({ alcoholFee: e.target.value === "" ? 0 : Number(e.target.value) })} onFocus={(e) => e.target.select()} />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(d.alcoholFee)}</span>
          </div>
        </Field>
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginTop: 12 }}>
        <ScoreBox label="Costo por persona (cena + servicio)" value={money(costoPorPersona)} />
        <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "8px 6px", textAlign: "center" }}>
          <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase" }}>Total recabado</div>
          <div style={{ ...monoFont, fontWeight: 700, fontSize: 14.5, color: C.goldSoft }}>{money(totalRecaudado)}</div>
          <div style={{ fontSize: 9.5, ...monoFont, color: "rgba(244,234,214,0.45)" }}>
            {redondeoExtra > 0 ? `+${money(redondeoExtra)} de redondeo` : "sin redondeo extra"}
          </div>
        </div>
      </div>
      <div style={{ marginTop: 12, display: "grid", gap: 6 }}>
        {players.map((p) => {
          const alcohol = !!d.alcohol[p.id];
          const charge = ceilTo100(d.total / numPlayers + sidesFee + d.waiter + (alcohol ? d.alcoholFee : 0));
          const paid = !!d.paid?.[p.id];
          return (
            <div key={p.id} style={{ background: "rgba(0,0,0,0.16)", borderRadius: 8, padding: "8px 10px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                <button onClick={() => setD((dd) => ({ alcohol: { ...dd.alcohol, [p.id]: !dd.alcohol[p.id] } }))}
                  style={{ display: "flex", alignItems: "center", gap: 7, background: "transparent", border: "none", cursor: "pointer", color: C.card, flex: 1, minWidth: 0 }}>
                  <Avatar player={p} size={22} />
                  <Wine size={14} color={alcohol ? C.virtual : "rgba(244,234,214,0.3)"} style={{ flexShrink: 0 }} />
                  <span style={{ fontSize: 13, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</span>
                </button>
                <span style={{ ...monoFont, fontSize: 13, color: "rgba(244,234,214,0.75)", whiteSpace: "nowrap" }}>{money(charge)}</span>
              </div>
              <div style={{ marginTop: 7 }}>
                <button onClick={() => setD((dd) => ({ paid: { ...dd.paid, [p.id]: !dd.paid?.[p.id] } }))}
                  style={{
                    display: "flex", alignItems: "center", justifyContent: "center", gap: 7, width: "100%",
                    background: paid ? "rgba(63,191,114,0.16)" : "rgba(0,0,0,0.2)",
                    border: `1px solid ${paid ? C.win : C.panelLine}`, borderRadius: 7, padding: "8px 9px",
                    cursor: "pointer", color: paid ? C.win : "rgba(244,234,214,0.6)", fontSize: 12.5, fontWeight: 700, ...bodyFont,
                  }}>
                  <div style={{ width: 15, height: 15, borderRadius: 4, border: `1.5px solid ${paid ? C.win : "rgba(244,234,214,0.4)"}`, background: paid ? C.win : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                    {paid && <Check size={11} color="#08251a" />}
                  </div>
                  {paid ? "Pagado" : "Confirmar pagado"}
                </button>
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.45)", marginTop: 8 }}>
        Este cargo de cena es independiente del conteo de lotes y no afecta el balance de la partida. Un jugador no puede comprar lotes hasta confirmar que pagó la cena.
      </div>
    </div>
  );
}

/* ----- Finalize: enter chips returned per player ----- */
function FinalizeGame({ game, roster, onBack, onConfirm, update }) {
  const players = game.playerIds.map((id) => roster.find((r) => r.id === id)).filter(Boolean);
  const [values, setValues] = useState(() =>
    Object.fromEntries(
      players.map((p) => {
        const existing = game.finalChips ? game.finalChips[p.id] : undefined;
        return [p.id, existing !== undefined && existing !== null ? String(existing) : ""];
      })
    )
  );

  // Al volver a la pantalla anterior, guardamos lo ya tipeado en la propia
  // partida (aunque no esté completo ni cuadre todavía) para no perderlo si
  // se vuelve a entrar más tarde a cerrar la partida.
  const handleBack = () => {
    const draft = {};
    players.forEach((p) => {
      if (values[p.id] !== "") draft[p.id] = Number(values[p.id]) || 0;
    });
    update({ finalChips: draft });
    onBack();
  };

  const buyIns = useMemo(() => {
    const map = {};
    players.forEach((p) => {
      const entries = game.purchases.filter((pu) => pu.playerId === p.id);
      const cash = entries.filter((e) => e.type === "cash").reduce((s, e) => s + e.amount, 0);
      const virtual = entries.filter((e) => e.type === "virtual").reduce((s, e) => s + e.amount, 0);
      map[p.id] = { cash, virtual, total: cash + virtual };
    });
    return map;
  }, [players, game.purchases]);

  const totalCash = useMemo(() => Object.values(buyIns).reduce((s, b) => s + b.cash, 0), [buyIns]);
  const totalVirtual = useMemo(() => Object.values(buyIns).reduce((s, b) => s + b.virtual, 0), [buyIns]);
  const rake = Number(game.rake) || 0;
  const targetTotal = totalCash + totalVirtual;

  const totalFinalValue = players.reduce((s, p) => s + (Number(values[p.id]) || 0), 0) + rake;
  const enteredCount = players.filter((p) => values[p.id] !== "").length;
  const runningDiff = round1(totalFinalValue - targetTotal);
  const diff = totalFinalValue - targetTotal;

  const ready = players.every((p) => values[p.id] === "" || !isNaN(Number(values[p.id])));
  // Regla de la app: no se puede cerrar la partida si lo entregado (fichas +
  // rake) no cuadra exactamente contra el total comprado (cash + virtual).
  const canConfirm = ready && enteredCount > 0 && runningDiff === 0;

  const cuadra = runningDiff === 0;
  const pendingLabel = enteredCount === 0
    ? "Ingresa los montos de cierre"
    : cuadra
      ? "Cuadra ✓"
      : runningDiff > 0
        ? `Sobran ${money(runningDiff)}`
        : `Faltan ${money(-runningDiff)}`;

  return (
    <div style={{ display: "grid", gap: 16, paddingBottom: 74 }}>
      <Panel>
        <SectionTitle icon={Trophy}>Entrega de fichas</SectionTitle>
        <div style={{ color: "rgba(244,234,214,0.6)", fontSize: 12.5, marginBottom: 10 }}>
          Cada jugador entrega el valor en dinero de las fichas que tiene en su poder al cierre (monto libre). El rake se considera como un jugador más dentro del total.
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 14 }}>
          <ScoreBox label="Cash" value={money(totalCash)} tone="cash" />
          <ScoreBox label="Virtual" value={money(totalVirtual)} tone="virtual" />
          <ScoreBox label="Total general" value={money(targetTotal)} />
          <div style={{
            background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "8px 6px", textAlign: "center",
            border: `1px solid ${enteredCount === 0 ? C.panelLine : runningDiff === 0 ? C.win : C.loss}`,
          }}>
            <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase" }}>Ingresado</div>
            <div style={{ ...monoFont, fontWeight: 700, fontSize: 14.5, color: enteredCount === 0 ? C.goldSoft : runningDiff === 0 ? C.win : C.loss }}>
              {money(totalFinalValue)}
            </div>
          </div>
        </div>
        <div style={{ display: "grid", gap: 8 }}>
          <div style={{ background: "rgba(216,173,63,0.1)", border: `1px dashed ${C.panelLine}`, borderRadius: 9, padding: "9px 10px", display: "flex", alignItems: "center", justifyContent: "space-between" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 26, height: 26, borderRadius: 99, background: "rgba(216,173,63,0.22)", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                <Flame size={13} color={C.gold} />
              </div>
              <span style={{ color: C.goldSoft, fontWeight: 700, fontSize: 13.5 }}>Rake (casa)</span>
              <span style={{ ...monoFont, fontSize: 10.5, color: "rgba(244,234,214,0.4)" }}>se resta automáticamente, no se captura</span>
            </div>
            <span style={{ ...monoFont, fontSize: 13.5, color: C.goldSoft, fontWeight: 700 }}>{money(rake)}</span>
          </div>
          {players.map((p) => {
            const bi = buyIns[p.id];
            return (
              <div key={p.id} style={{ background: "rgba(0,0,0,0.18)", borderRadius: 9, padding: "9px 10px" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <Avatar player={p} size={26} />
                    <span style={{ color: C.card, fontWeight: 600, fontSize: 13.5 }}>{p.name}</span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                    <span style={{ color: "rgba(244,234,214,0.5)", fontSize: 13 }}>$</span>
                    <input type="number" min="0" style={{ ...inputStyle, width: 110, textAlign: "right" }} placeholder="0"
                      value={values[p.id]} onChange={(e) => setValues((c) => ({ ...c, [p.id]: e.target.value }))}
                      onFocus={(e) => e.target.select()}
                      onBlur={(e) => { if (e.target.value !== "") setValues((c) => ({ ...c, [p.id]: String(roundTo100(e.target.value)) })); }} step="100" />
                  </div>
                </div>
                <div style={{ display: "flex", gap: 6, marginTop: 7, marginLeft: 34 }}>
                  <Badge tone="cash">Cash {money(bi.cash)}</Badge>
                  <Badge tone="virtual">Virtual {money(bi.virtual)}</Badge>
                  <span style={{ ...monoFont, fontSize: 11, color: "rgba(244,234,214,0.5)", alignSelf: "center" }}>
                    Total buy-in {money(bi.total)}
                  </span>
                </div>
              </div>
            );
          })}
        </div>
        {ready && enteredCount > 0 && runningDiff !== 0 && (
          <div style={{ display: "flex", gap: 7, alignItems: "flex-start", marginTop: 12, background: "rgba(226,99,79,0.12)", border: `1px solid ${C.loss}`, borderRadius: 8, padding: "8px 10px" }}>
            <AlertCircle size={15} color={C.loss} style={{ flexShrink: 0, marginTop: 1 }} />
            <div style={{ fontSize: 12, color: "rgba(244,234,214,0.85)" }}>
              <strong>No se puede cerrar la partida hasta que el total cuadre exactamente</strong> — ajustá los montos de cierre.
            </div>
          </div>
        )}
      </Panel>

      {/* Barra fija: siempre visible mientras bajás por la lista de jugadores,
          para ver de un vistazo cuánto falta o sobra sin tener que volver a
          scrollear hasta arriba. */}
      <div style={{
        position: "fixed", left: 0, right: 0, bottom: 0, zIndex: 40,
        background: "rgba(8,32,25,0.97)", borderTop: `1px solid ${cuadra ? C.win : enteredCount === 0 ? C.panelLine : C.loss}`,
        backdropFilter: "blur(6px)", padding: "10px 14px calc(10px + env(safe-area-inset-bottom, 0px))",
      }}>
        <div style={{ maxWidth: 980, margin: "0 auto", display: "flex", alignItems: "center", gap: 12 }}>
          <div style={{ flex: 1, minWidth: 0 }}>
            <div style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase", letterSpacing: "0.04em" }}>Pendiente para cuadrar</div>
            <div style={{
              ...displayFont, fontSize: 22, lineHeight: 1.1,
              color: enteredCount === 0 ? C.goldSoft : cuadra ? C.win : C.loss,
              overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap",
            }}>
              {pendingLabel}
            </div>
          </div>
          <GhostBtn onClick={handleBack} style={{ flexShrink: 0 }}>Volver</GhostBtn>
          <PrimaryBtn
            disabled={!canConfirm}
            onClick={() => onConfirm(Object.fromEntries(players.map((p) => [p.id, Number(values[p.id]) || 0])))}
            icon={Trophy} style={{ flexShrink: 0, padding: "12px 16px" }}
          >
            Calcular
          </PrimaryBtn>
        </div>
      </div>
    </div>
  );
}

/* ----- Finalized game: results table + transfers + save/close ----- */
function ChampionBanner({ winner, player }) {
  if (!winner) return null;
  return (
    <div style={{
      display: "flex", flexDirection: "column", gap: 8,
      background: "linear-gradient(135deg, rgba(255,205,90,0.18), rgba(255,205,90,0.05))",
      border: `1px solid ${C.gold}`, borderRadius: 12,
      padding: "10px 14px", margin: "6px 0 12px",
    }}>
      <div style={{
        fontSize: 10.5, letterSpacing: "0.08em", textTransform: "uppercase",
        color: "rgba(244,234,214,0.55)", fontWeight: 700,
        whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis",
      }}>
        Próximos postres, cortesía de:
      </div>
      <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
        <Trophy size={48} color={C.gold} style={{ flexShrink: 0 }} />
        {player ? <Avatar player={player} size={30} /> : null}
        <div style={{ ...displayFont, fontSize: 22, color: C.gold, letterSpacing: "0.02em", lineHeight: 1.1, minWidth: 0 }}>
          {winner.name}
        </div>
        <span style={{ marginLeft: "auto", ...monoFont, fontSize: 15, fontWeight: 800, color: C.win }}>
          {(winner.balance > 0 ? "+" : "") + money(winner.balance)}
        </span>
      </div>
    </div>
  );
}
function ResultRow({ label, value, tone, strong }) {
  const color = tone === "cash" ? C.cash : tone === "virtual" ? C.virtual : tone === "win" ? C.win : tone === "loss" ? C.loss : "rgba(244,234,214,0.85)";
  return (
    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "6px 0", borderBottom: `1px solid rgba(255,255,255,0.06)` }}>
      <span style={{ fontSize: 12.5, color: "rgba(244,234,214,0.55)" }}>{label}</span>
      <span style={{ ...monoFont, fontSize: 14, fontWeight: strong ? 800 : 600, color }}>{value}</span>
    </div>
  );
}
function PlayerResultCard({ p, player }) {
  const win = p.balance > 0;
  const flat = p.balance === 0;
  const cashLabel = p.pagoCash > 0 ? "Recibe Cash" : p.pagoCash < 0 ? "Debe Cash" : "Cash";
  const transferLabel = p.pagoTransfer > 0 ? "Recibe Transfer" : p.pagoTransfer < 0 ? "Debe Transfer" : "Transfer";
  return (
    <div style={{ background: "rgba(0,0,0,0.2)", border: `1px solid ${win ? "rgba(63,191,114,0.35)" : flat ? C.panelLine : "rgba(226,99,79,0.35)"}`, borderRadius: 12, padding: "12px 14px" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 6 }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          {player ? <Avatar player={player} size={30} /> : null}
          <span style={{ ...displayFont, fontSize: 19, color: C.card, letterSpacing: "0.03em" }}>{p.name}</span>
        </div>
        <span style={{
          ...monoFont, fontWeight: 800, fontSize: 15,
          color: win ? C.win : flat ? "rgba(244,234,214,0.6)" : C.loss,
        }}>
          {(p.balance > 0 ? "+" : "") + money(p.balance)}
        </span>
      </div>
      <ResultRow label="Buy-in cash" value={money(p.cashAmount)} tone="cash" />
      <ResultRow label="Buy-in virtual" value={money(p.virtualAmount)} tone="virtual" />
      <ResultRow label="Total buy-in" value={money(p.totalBuyIn)} />
      <ResultRow label="Cash out (fichas)" value={money(p.cashOut)} />
      <ResultRow label="Balance" value={(p.balance > 0 ? "+" : "") + money(p.balance)} tone={win ? "win" : flat ? undefined : "loss"} strong />
      <ResultRow label={cashLabel} value={money(Math.abs(p.pagoCash))} tone={p.pagoCash > 0 ? "cash" : undefined} />
      <ResultRow label={transferLabel} value={money(Math.abs(p.pagoTransfer))} tone={p.pagoTransfer > 0 ? "win" : p.pagoTransfer < 0 ? "loss" : undefined} />
    </div>
  );
}
function FinalizedGame({ game, roster, onClose, setActiveGame, setGames }) {
  const r = useMemo(() => computeSettlement(game, roster), [game, roster]);
  const players = game.playerIds.map((id) => roster.find((p) => p.id === id)).filter(Boolean);
  const numPlayers = players.length || 1;
  const d = game.dinner;
  const sidesFee = d.sidesFee || 0;
  const winner = [...r.players].sort((a, b) => b.balance - a.balance)[0];
  const winnerPlayer = players.find((pl) => pl.id === winner?.playerId);

  // La cena no afecta el balance de lotes, así que es seguro seguir editando
  // quién pagó/con qué método incluso después de haber cerrado la partida.
  const updateDinner = (patch) => {
    const applyPatch = (dinner) => ({ ...dinner, ...(typeof patch === "function" ? patch(dinner) : patch) });
    setActiveGame((g) => ({ ...g, dinner: applyPatch(g.dinner) }));
    setGames((gs) => gs.map((x) => (x.id === game.id ? { ...x, dinner: applyPatch(x.dinner) } : x)));
  };

  // Vuelve a dejar la partida "en curso" para poder corregir lotes, fichas
  // de cierre, rake, etc. La sacamos del historial hasta que se vuelva a
  // cerrar, para no dejar una copia vieja e inconsistente dando vueltas.
  const handleBack = () => {
    if (!confirm("¿Volver a editar esta partida? Vas a poder corregir lotes, rake y fichas de cierre, y después cerrarla de nuevo.")) return;
    setGames((gs) => gs.filter((g) => g.id !== game.id));
    setActiveGame({ ...game, finished: false });
  };

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <GhostBtn onClick={handleBack} icon={ChevronLeft} color={C.goldSoft}>Volver a editar partida</GhostBtn>

      <Panel>
        <SectionTitle icon={Trophy}>Resultados de la partida</SectionTitle>
        <div style={{ ...displayFont, fontSize: 18, color: C.goldSoft, marginTop: -4 }}>{game.date}</div>
        <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.5)", display: "flex", alignItems: "center", gap: 4, marginBottom: 4, marginTop: 2 }}>
          <Crown size={12} color={C.gold} /> Operó: {roster.find((pl) => pl.id === game.hostId)?.name || "—"}
        </div>
        <ChampionBanner winner={winner} player={winnerPlayer} />
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8, marginBottom: 14 }}>
          <ScoreBox label="Rake" value={money(r.rake)} />
          <ScoreBox label="Total buy-in cash" value={money(r.players.reduce((s, p) => s + p.cashAmount, 0))} tone="cash" />
          <ScoreBox label="Total buy-in virtual" value={money(r.players.reduce((s, p) => s + p.virtualAmount, 0))} tone="virtual" />
        </div>

        <div style={{ display: "grid", gap: 10 }}>
          {[...r.players].sort((a, b) => b.cashOut - a.cashOut).map((p) => (
            <PlayerResultCard key={p.playerId} p={p} player={players.find((pl) => pl.id === p.playerId)} />
          ))}
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={Banknote}>Reparto de efectivo</SectionTitle>
        <div style={{ display: "grid", gap: 8 }}>
          {r.players.filter((p) => p.pagoCash > 0).map((p) => (
            <div key={p.playerId} style={{ display: "flex", justifyContent: "space-between", background: "rgba(0,0,0,0.18)", borderRadius: 8, padding: "8px 10px" }}>
              <span style={{ color: C.card, fontSize: 13.5, fontWeight: 600 }}>{p.name}</span>
              <span style={{ ...monoFont, fontSize: 12.5, color: C.cash }}>{money(p.pagoCash)} · {billsLabel(billsFor(p.pagoCash))}</span>
            </div>
          ))}
          {r.players.every((p) => p.pagoCash === 0) && <Empty>No hay efectivo para repartir.</Empty>}
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={ArrowRightLeft}>Transferencias sugeridas</SectionTitle>
        <div style={{ display: "grid", gap: 8 }}>
          {r.transfers.length === 0 && <Empty>No se requieren transferencias.</Empty>}
          {r.transfers.map((t, i) => (
            <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(0,0,0,0.18)", borderRadius: 8, padding: "9px 12px" }}>
              <span style={{ color: C.card, fontWeight: 700, fontSize: 13.5 }}>{t.from}</span>
              <ArrowRightLeft size={13} color={C.gold} />
              <span style={{ color: C.card, fontWeight: 700, fontSize: 13.5 }}>{t.to}</span>
              <span style={{ marginLeft: "auto", ...monoFont, color: C.gold, fontWeight: 700 }}>{money(t.amount)}</span>
            </div>
          ))}
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={UtensilsCrossed}>Cargo de cena (no afecta el balance de lotes)</SectionTitle>
        <div style={{ display: "grid", gap: 6 }}>
          {players.map((p) => {
            const alcohol = !!d.alcohol[p.id];
            const charge = ceilTo100(d.total / numPlayers + sidesFee + d.waiter + (alcohol ? d.alcoholFee : 0));
            const paid = !!d.paid?.[p.id];
            return (
              <div key={p.id} style={{ background: "rgba(0,0,0,0.16)", borderRadius: 8, padding: "8px 10px" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                  <button onClick={() => updateDinner((dd) => ({ alcohol: { ...dd.alcohol, [p.id]: !dd.alcohol[p.id] } }))}
                    style={{ display: "flex", alignItems: "center", gap: 7, background: "transparent", border: "none", cursor: "pointer", color: C.card, flex: 1, minWidth: 0 }}>
                    <Avatar player={p} size={22} />
                    <Wine size={14} color={alcohol ? C.virtual : "rgba(244,234,214,0.3)"} style={{ flexShrink: 0 }} />
                    <span style={{ fontSize: 13, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</span>
                  </button>
                  <span style={{ ...monoFont, fontSize: 13, color: "rgba(244,234,214,0.75)", whiteSpace: "nowrap" }}>{money(charge)}</span>
                </div>
                <div style={{ marginTop: 7 }}>
                  <button onClick={() => updateDinner((dd) => ({ paid: { ...dd.paid, [p.id]: !dd.paid?.[p.id] } }))}
                    style={{
                      display: "flex", alignItems: "center", justifyContent: "center", gap: 7, width: "100%",
                      background: paid ? "rgba(63,191,114,0.16)" : "rgba(0,0,0,0.2)",
                      border: `1px solid ${paid ? C.win : C.panelLine}`, borderRadius: 7, padding: "8px 9px",
                      cursor: "pointer", color: paid ? C.win : "rgba(244,234,214,0.6)", fontSize: 12.5, fontWeight: 700, ...bodyFont,
                    }}>
                    <div style={{ width: 15, height: 15, borderRadius: 4, border: `1.5px solid ${paid ? C.win : "rgba(244,234,214,0.4)"}`, background: paid ? C.win : "transparent", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                      {paid && <Check size={11} color="#08251a" />}
                    </div>
                    {paid ? "Pagado" : "Confirmar pagado"}
                  </button>
                </div>
              </div>
            );
          })}
        </div>
      </Panel>

      <PrimaryBtn onClick={onClose} icon={Save} style={{ padding: "13px 18px", fontSize: 15 }}>
        Guardar y cerrar partida
      </PrimaryBtn>
    </div>
  );
}
function Th({ children, align }) {
  return <th style={{ textAlign: align || "right", padding: "0 8px 8px", whiteSpace: "nowrap" }}>{children}</th>;
}
function Td({ children, align, bold, tone, signed, value, strong }) {
  let color = "rgba(244,234,214,0.85)";
  if (tone === "cash") color = C.cash;
  if (tone === "virtual") color = C.virtual;
  if (signed && typeof value === "number") color = value > 0 ? C.win : value < 0 ? C.loss : color;
  return (
    <td style={{ textAlign: align || "right", padding: "8px", ...monoFont, fontSize: 13, color, fontWeight: bold || strong ? 700 : 500, whiteSpace: "nowrap" }}>
      {children}
    </td>
  );
}

/* ----------------------------------------------------------------------
   HISTORY TAB
---------------------------------------------------------------------- */
function SwipeableRow({ children, onDelete }) {
  const [dragX, setDragX] = useState(0);
  const [open, setOpen] = useState(false);
  const startX = useRef(0);
  const dragging = useRef(false);
  const OPEN_X = -84;

  const onTouchStart = (e) => {
    startX.current = e.touches[0].clientX;
    dragging.current = true;
  };
  const onTouchMove = (e) => {
    if (!dragging.current) return;
    const delta = e.touches[0].clientX - startX.current;
    const base = open ? OPEN_X : 0;
    const next = Math.max(OPEN_X, Math.min(0, base + delta));
    setDragX(next);
  };
  const onTouchEnd = () => {
    dragging.current = false;
    if (dragX < OPEN_X / 2) { setDragX(OPEN_X); setOpen(true); }
    else { setDragX(0); setOpen(false); }
  };

  return (
    <div style={{ position: "relative", overflow: "hidden", borderRadius: 14 }}>
      <div style={{
        position: "absolute", top: 0, right: 0, bottom: 0, width: 84,
        display: "flex", alignItems: "center", justifyContent: "center", background: C.loss,
      }}>
        <button
          onClick={() => { onDelete(); setDragX(0); setOpen(false); }}
          style={{ background: "transparent", border: "none", color: "#fff", display: "flex", flexDirection: "column", alignItems: "center", gap: 3, cursor: "pointer", width: "100%", height: "100%" }}
        >
          <Trash2 size={18} />
          <span style={{ fontSize: 10.5, fontWeight: 700, ...bodyFont }}>Eliminar</span>
        </button>
      </div>
      <div
        onTouchStart={onTouchStart}
        onTouchMove={onTouchMove}
        onTouchEnd={onTouchEnd}
        onClick={() => { if (open) { setDragX(0); setOpen(false); } }}
        style={{ transform: `translateX(${dragX}px)`, transition: dragging.current ? "none" : "transform 0.2s ease" }}
      >
        {children}
      </div>
    </div>
  );
}

function HistoryTab({ games, roster, setGames, adminPassword }) {
  const [openId, setOpenId] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const sorted = [...games].sort((a, b) => (a.date < b.date ? 1 : -1));

  const handleDelete = async (g) => {
    if (!(await requestAdminPassword(adminPassword, "eliminar partidas"))) return;
    if (!confirm(`¿Eliminar definitivamente la partida del ${g.date}? Esta acción no se puede deshacer.`)) return;
    setGames((gs) => gs.filter((x) => x.id !== g.id));
  };

  // Fuerza un reguardado de todas las partidas sin cambiar ningún dato — sirve
  // para que la hoja "Resultados" del Excel se regenere con la fórmula de
  // liquidación más reciente, sin tener que reabrir y reingresar cada partida.
  const handleResync = async () => {
    if (!(await requestAdminPassword(adminPassword, "recalcular y sincronizar la base de datos"))) return;
    setSyncing(true);
    setGames((gs) => [...gs]);
    setTimeout(() => setSyncing(false), 1200);
  };

  if (sorted.length === 0) {
    return <Panel><Empty>Aún no hay partidas guardadas.</Empty></Panel>;
  }

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <div style={{ display: "flex", justifyContent: "center" }}>
        <GhostBtn onClick={handleResync} icon={ArrowRightLeft} color={C.goldSoft}>
          {syncing ? "Sincronizando…" : "Recalcular y sincronizar con la base de datos"}
        </GhostBtn>
      </div>
      <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.45)", textAlign: "center" }}>
        Deslizá una partida hacia la izquierda para eliminarla (pide contraseña).
      </div>
      {sorted.map((g) => {
        const open = openId === g.id;
        const results = computeSettlement(g, roster);
        const winner = results.players.slice().sort((a, b) => b.balance - a.balance)[0];
        return (
          <SwipeableRow key={g.id} onDelete={() => handleDelete(g)}>
            <Panel>
              <button onClick={() => setOpenId(open ? null : g.id)} style={{ width: "100%", background: "transparent", border: "none", cursor: "pointer", padding: 0 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <div style={{ textAlign: "left" }}>
                    <div style={{ ...displayFont, fontSize: 18, color: C.goldSoft }}>{g.date}</div>
                    <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.5)" }}>
                      {g.playerIds.length} jugadores · lote {money(g.loteValue)} · rake {money(g.rake)}
                    </div>
                    <div style={{ fontSize: 11, color: "rgba(244,234,214,0.4)", marginTop: 2, display: "flex", alignItems: "center", gap: 4 }}>
                      <Crown size={11} color={C.gold} /> Operó: {roster.find((r) => r.id === g.hostId)?.name || "—"}
                    </div>
                  </div>
                  {open ? <ChevronUp size={16} color={C.goldSoft} /> : <ChevronDown size={16} color={C.goldSoft} />}
                </div>
              </button>
              <ChampionBanner winner={winner} player={roster.find((pl) => pl.id === winner?.playerId)} />
              {open && (
                <div style={{ display: "grid", gap: 8, marginTop: 12 }}>
                  {[...results.players].sort((a, b) => b.cashOut - a.cashOut).map((p) => (
                    <PlayerResultCard key={p.playerId} p={p} player={roster.find((pl) => pl.id === p.playerId)} />
                  ))}

                  <div style={{ marginTop: 6 }}>
                    <SectionTitle icon={Banknote}>Reparto de efectivo</SectionTitle>
                    <div style={{ display: "grid", gap: 8 }}>
                      {results.players.filter((p) => p.pagoCash > 0).map((p) => (
                        <div key={p.playerId} style={{ display: "flex", justifyContent: "space-between", background: "rgba(0,0,0,0.18)", borderRadius: 8, padding: "8px 10px" }}>
                          <span style={{ color: C.card, fontSize: 13.5, fontWeight: 600 }}>{p.name}</span>
                          <span style={{ ...monoFont, fontSize: 12.5, color: C.cash }}>{money(p.pagoCash)} · {billsLabel(billsFor(p.pagoCash))}</span>
                        </div>
                      ))}
                      {results.players.every((p) => p.pagoCash === 0) && <Empty>No hay efectivo para repartir.</Empty>}
                    </div>
                  </div>

                  <div style={{ marginTop: 6 }}>
                    <SectionTitle icon={ArrowRightLeft}>Transferencias sugeridas</SectionTitle>
                    <div style={{ display: "grid", gap: 8 }}>
                      {results.transfers.length === 0 && <Empty>No se requieren transferencias.</Empty>}
                      {results.transfers.map((t, i) => (
                        <div key={i} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(0,0,0,0.18)", borderRadius: 8, padding: "9px 12px" }}>
                          <span style={{ color: C.card, fontWeight: 700, fontSize: 13.5 }}>{t.from}</span>
                          <ArrowRightLeft size={13} color={C.gold} />
                          <span style={{ color: C.card, fontWeight: 700, fontSize: 13.5 }}>{t.to}</span>
                          <span style={{ marginLeft: "auto", ...monoFont, color: C.gold, fontWeight: 700 }}>{money(t.amount)}</span>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </Panel>
          </SwipeableRow>
        );
      })}
    </div>
  );
}
