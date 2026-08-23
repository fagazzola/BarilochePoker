import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Plus, Minus, Trash2, Pencil, Users, UtensilsCrossed, Wine,
  Trophy, ArrowRightLeft, Save, X, Check, ChevronDown, ChevronUp, ChevronLeft,
  Banknote, Landmark, Flame, History, UserPlus, UserX, UserCheck,
  Play, Square, AlertCircle, Crown, DollarSign, CircleDollarSign, Coins,
  BarChart3
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
function requestAdminPassword(adminPassword, actionLabel) {
  if (!adminPassword) {
    alert(`Todavía no hay una contraseña de administrador configurada en el Excel (hoja "Meta", fila con key=admin_password). Agregala ahí para poder ${actionLabel}.`);
    return false;
  }
  const entered = window.prompt(`Contraseña de administrador para ${actionLabel}:`);
  if (entered === null) return false;
  if (entered !== adminPassword) {
    alert("Contraseña incorrecta.");
    return false;
  }
  return true;
}
const ceilTo100 = (n) => Math.ceil((Number(n) || 0) / 50) * 50;

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

      <Header tab={tab} setTab={setTab} hasActive={!!activeGame} />

      <main style={{ maxWidth: 980, margin: "0 auto", padding: "18px 14px 60px" }}>
        {tab === "jugadores" && (
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
    </div>
  );
}

/* ----------------------------------------------------------------------
   HEADER / TABS
---------------------------------------------------------------------- */
function Header({ tab, setTab, hasActive }) {
  const tabs = [
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
  const removePlayer = () => {
    if (!sel) return;
    if (!requestAdminPassword(adminPassword, "eliminar jugadores")) return;
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
  const [loteValue, setLoteValue] = useState(500);
  const [rake, setRake] = useState(1000);
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
      playerIds: selected, hostId,
      purchases: [],
      dinner: { total: 0, waiter: 0, alcoholFee: 0, alcohol: {}, paid: {}, paymentMethod: {} },
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
          <Field label="Rake (se puede ajustar después)">
            <input type="number" min="0" style={inputStyle} value={rake} onChange={(e) => setRake(e.target.value)} onFocus={(e) => e.target.select()} />
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
function ActiveGameScreen({ game, setGame, roster, setGames }) {
  const [dinnerOpen, setDinnerOpen] = useState(false);
  const [finalizing, setFinalizing] = useState(false);

  const players = game.playerIds.map((id) => roster.find((r) => r.id === id)).filter(Boolean);

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

  const setRake = (v) => update({ rake: Number(v) || 0 });
  const setLote = (v) => update({ loteValue: Number(v) || 0 });

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

  return (
    <div style={{ display: "grid", gap: 16 }}>
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
          <GhostBtn icon={X} color={C.loss} onClick={() => { if (confirm("¿Cancelar esta partida? Se perderá el progreso.")) setGame(null); }}>
            Cancelar partida
          </GhostBtn>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginTop: 14 }}>
          <ScoreBox label="Lote" value={money(game.loteValue)} />
          <ScoreBox label="Cash" value={money(totals.cash)} tone="cash" />
          <ScoreBox label="Virtual" value={money(totals.virtual)} tone="virtual" />
          <ScoreBox label="Rake" value={money(game.rake)} />
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 12 }}>
          <Field label="Valor de lote"><input type="number" style={inputStyle} value={game.loteValue} onChange={(e) => setLote(e.target.value)} onFocus={(e) => e.target.select()} /></Field>
          <Field label="Rake"><input type="number" style={inputStyle} value={game.rake} onChange={(e) => setRake(e.target.value)} onFocus={(e) => e.target.select()} /></Field>
        </div>
      </Panel>

      <Panel>
        <SectionTitle icon={Banknote}>Compra de lotes por jugador</SectionTitle>
        <div style={{ display: "grid", gap: 10 }}>
          {players.map((p) => (
            <PlayerBuyRow key={p.id} player={p} game={game} onAdd={addPurchase} onRemoveLast={removeLastPurchase} />
          ))}
        </div>
      </Panel>

      <Panel>
        <button onClick={() => setDinnerOpen((o) => !o)} style={{ width: "100%", background: "transparent", border: "none", cursor: "pointer", padding: 0 }}>
          <SectionTitle icon={UtensilsCrossed} right={dinnerOpen ? <ChevronUp size={16} color={C.goldSoft} /> : <ChevronDown size={16} color={C.goldSoft} />}>
            Cena y servicio
          </SectionTitle>
        </button>
        {dinnerOpen && <DinnerSection game={game} players={players} update={update} onClose={() => setDinnerOpen(false)} />}
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

function PlayerBuyRow({ player, game, onAdd, onRemoveLast }) {
  const entries = game.purchases.filter((p) => p.playerId === player.id);
  const cashLotes = entries.filter((e) => e.type === "cash").reduce((s, e) => s + e.lotes, 0);
  const virtualLotes = entries.filter((e) => e.type === "virtual").reduce((s, e) => s + e.lotes, 0);
  const cashAmount = cashLotes * game.loteValue;
  const virtualAmount = virtualLotes * game.loteValue;

  return (
    <div style={{ background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: 12 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
        <Avatar player={player} size={26} />
        <span style={{ color: C.card, fontWeight: 700, fontSize: 14.5 }}>{player.name}</span>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 8 }}>
        <button onClick={() => onAdd(player.id, "cash")} style={buyBtnStyle(C.cash, false)}>
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
        <button onClick={() => onAdd(player.id, "virtual")} style={buyBtnStyle(C.virtual, false)}>
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

function DinnerSection({ game, players, update, onClose }) {
  const d = game.dinner;
  const setD = (patch) =>
    update((g) => ({
      dinner: { ...g.dinner, ...(typeof patch === "function" ? patch(g.dinner) : patch) },
    }));
  const numPlayers = players.length || 1;

  // Costo base por persona (cena repartida + servicio), sin alcohol porque
  // ese varía según quién bebió. Es el número que se muestra como referencia.
  const costoPorPersona = d.total / numPlayers + d.waiter;
  const nominalTotal = players.reduce((s, p) => {
    const alcohol = !!d.alcohol[p.id];
    return s + (d.total / numPlayers + d.waiter + (alcohol ? d.alcoholFee : 0));
  }, 0);
  const totalRecaudado = players.reduce((s, p) => {
    const alcohol = !!d.alcohol[p.id];
    return s + ceilTo100(d.total / numPlayers + d.waiter + (alcohol ? d.alcoholFee : 0));
  }, 0);
  const redondeoExtra = round1(totalRecaudado - nominalTotal);

  return (
    <div style={{ marginTop: 6 }}>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Monto total de la cena">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.total} onChange={(e) => setD({ total: Number(e.target.value) || 0 })} onFocus={(e) => e.target.select()} onBlur={(e) => setD({ total: roundTo100(e.target.value) })} step="100" />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(d.total)}</span>
          </div>
        </Field>
        <Field label="Servicio de mesero (por jugador)">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.waiter} onChange={(e) => setD({ waiter: Number(e.target.value) || 0 })} onFocus={(e) => e.target.select()} />
            <span style={{ ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.5)", whiteSpace: "nowrap" }}>{money(d.waiter)}</span>
          </div>
        </Field>
      </div>
      <div style={{ marginTop: 10 }}>
        <Field label="Cargo extra de alcohol (se suma a la cena de quienes bebieron)">
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <input type="number" style={inputStyle} value={d.alcoholFee} onChange={(e) => setD({ alcoholFee: Number(e.target.value) || 0 })} onFocus={(e) => e.target.select()} onBlur={(e) => setD({ alcoholFee: roundTo100(e.target.value) })} step="100" />
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
          const charge = ceilTo100(d.total / numPlayers + d.waiter + (alcohol ? d.alcoholFee : 0));
          const paid = !!d.paid?.[p.id];
          const method = d.paymentMethod?.[p.id] || "fichas";
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
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginTop: 7 }}>
                <select
                  value={method}
                  onChange={(e) => setD((dd) => ({ paymentMethod: { ...dd.paymentMethod, [p.id]: e.target.value } }))}
                  style={{ ...inputStyle, appearance: "auto", padding: "5px 8px", fontSize: 12, width: "auto", flex: 1 }}
                >
                  <option value="fichas">Pago con fichas</option>
                  <option value="cash">Pago con cash</option>
                  <option value="mixto">Combinado (fichas + cash)</option>
                </select>
                <button onClick={() => setD((dd) => ({ paid: { ...dd.paid, [p.id]: !dd.paid?.[p.id] } }))}
                  style={{
                    display: "flex", alignItems: "center", gap: 6, background: paid ? "rgba(63,191,114,0.16)" : "rgba(0,0,0,0.2)",
                    border: `1px solid ${paid ? C.win : C.panelLine}`, borderRadius: 7, padding: "5px 9px",
                    cursor: "pointer", color: paid ? C.win : "rgba(244,234,214,0.6)", fontSize: 11.5, fontWeight: 700, ...bodyFont, flexShrink: 0,
                  }}>
                  <div style={{ width: 14, height: 14, borderRadius: 4, border: `1.5px solid ${paid ? C.win : "rgba(244,234,214,0.4)"}`, background: paid ? C.win : "transparent", display: "flex", alignItems: "center", justifyContent: "center" }}>
                    {paid && <Check size={10} color="#08251a" />}
                  </div>
                  Pagado
                </button>
              </div>
            </div>
          );
        })}
      </div>
      <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.45)", marginTop: 8 }}>
        Este cargo de cena es independiente del conteo de lotes y no afecta el balance de la partida.
      </div>
      <div style={{ marginTop: 12 }}>
        <GhostBtn onClick={onClose} icon={ChevronUp} color={C.goldSoft}>Cerrar cena</GhostBtn>
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

  return (
    <div style={{ display: "grid", gap: 16 }}>
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
            <div style={{ fontSize: 9.5, ...monoFont, color: "rgba(244,234,214,0.45)" }}>
              {enteredCount === 0 ? "—" : runningDiff === 0 ? "cuadra ✓" : (runningDiff > 0 ? "+" : "") + money(runningDiff)}
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
              El valor total entregado, incluyendo el rake ({money(totalFinalValue)}), no coincide con el total general (cash + virtual: {money(targetTotal)}). Diferencia: {money(runningDiff)}. <strong>No se puede cerrar la partida hasta que el total cuadre exactamente</strong> — ajustá los montos de cierre.
            </div>
          </div>
        )}
      </Panel>
      <div style={{ display: "flex", gap: 10 }}>
        <GhostBtn onClick={handleBack}>Volver</GhostBtn>
        <PrimaryBtn
          disabled={!canConfirm}
          onClick={() => onConfirm(Object.fromEntries(players.map((p) => [p.id, Number(values[p.id]) || 0])))}
          icon={Trophy} style={{ flex: 1, padding: "12px 16px" }}
        >
          Calcular resultados
        </PrimaryBtn>
      </div>
    </div>
  );
}

/* ----- Finalized game: results table + transfers + save/close ----- */
function ChampionBanner({ winner, player }) {
  if (!winner) return null;
  return (
    <div style={{
      display: "flex", alignItems: "center", gap: 10,
      background: "linear-gradient(135deg, rgba(255,205,90,0.18), rgba(255,205,90,0.05))",
      border: `1px solid ${C.gold}`, borderRadius: 12,
      padding: "10px 14px", margin: "6px 0 12px",
    }}>
      <Trophy size={24} color={C.gold} style={{ flexShrink: 0 }} />
      {player ? <Avatar player={player} size={30} /> : null}
      <div style={{ minWidth: 0 }}>
        <div style={{ fontSize: 10.5, letterSpacing: "0.08em", textTransform: "uppercase", color: "rgba(244,234,214,0.55)", fontWeight: 700 }}>
          Campeón de la noche
        </div>
        <div style={{ ...displayFont, fontSize: 22, color: C.gold, letterSpacing: "0.02em", lineHeight: 1.1 }}>
          {winner.name}
        </div>
      </div>
      <span style={{ marginLeft: "auto", ...monoFont, fontSize: 15, fontWeight: 800, color: C.win }}>
        {(winner.balance > 0 ? "+" : "") + money(winner.balance)}
      </span>
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
            const charge = ceilTo100(d.total / numPlayers + d.waiter + (alcohol ? d.alcoholFee : 0));
            const paid = !!d.paid?.[p.id];
            const method = d.paymentMethod?.[p.id] || "fichas";
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
                <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginTop: 7 }}>
                  <select
                    value={method}
                    onChange={(e) => updateDinner((dd) => ({ paymentMethod: { ...dd.paymentMethod, [p.id]: e.target.value } }))}
                    style={{ ...inputStyle, appearance: "auto", padding: "5px 8px", fontSize: 12, width: "auto", flex: 1 }}
                  >
                    <option value="fichas">Pago con fichas</option>
                    <option value="cash">Pago con cash</option>
                    <option value="mixto">Combinado (fichas + cash)</option>
                  </select>
                  <button onClick={() => updateDinner((dd) => ({ paid: { ...dd.paid, [p.id]: !dd.paid?.[p.id] } }))}
                    style={{
                      display: "flex", alignItems: "center", gap: 6, background: paid ? "rgba(63,191,114,0.16)" : "rgba(0,0,0,0.2)",
                      border: `1px solid ${paid ? C.win : C.panelLine}`, borderRadius: 7, padding: "5px 9px",
                      cursor: "pointer", color: paid ? C.win : "rgba(244,234,214,0.6)", fontSize: 11.5, fontWeight: 700, ...bodyFont, flexShrink: 0,
                    }}>
                    <div style={{ width: 14, height: 14, borderRadius: 4, border: `1.5px solid ${paid ? C.win : "rgba(244,234,214,0.4)"}`, background: paid ? C.win : "transparent", display: "flex", alignItems: "center", justifyContent: "center" }}>
                      {paid && <Check size={10} color="#08251a" />}
                    </div>
                    {paid ? "Pagado" : "Pendiente"}
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

  const handleDelete = (g) => {
    if (!requestAdminPassword(adminPassword, "eliminar partidas")) return;
    if (!confirm(`¿Eliminar definitivamente la partida del ${g.date}? Esta acción no se puede deshacer.`)) return;
    setGames((gs) => gs.filter((x) => x.id !== g.id));
  };

  // Fuerza un reguardado de todas las partidas sin cambiar ningún dato — sirve
  // para que la hoja "Resultados" del Excel se regenere con la fórmula de
  // liquidación más reciente, sin tener que reabrir y reingresar cada partida.
  const handleResync = () => {
    if (!requestAdminPassword(adminPassword, "recalcular y sincronizar la base de datos")) return;
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
