import React, { useState, useEffect, useMemo, useCallback, useRef } from "react";
import {
  Plus, Minus, Trash2, Pencil, Users, UtensilsCrossed, Wine,
  Trophy, ArrowRightLeft, Save, X, Check, ChevronDown, ChevronUp, ChevronLeft,
  Banknote, Landmark, Flame, History, UserPlus, UserX, UserCheck,
  Play, Square, AlertCircle, Crown, DollarSign, CircleDollarSign, Coins,
  BarChart3, Activity
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
// Versión de la app, formato 1.AA.BB.CCC:
//   AA  = cantidad de rondas de cambios entregadas a src/App.jsx (esta app)
//   BB  = cantidad de rondas de cambios entregadas a public/dashboard.html
//   CCC = total acumulado de rondas de entrega (incluye AA + BB + cualquier
//         otro archivo, p. ej. netlify/functions) — nunca baja.
// Se actualiza a mano en cada ronda de cambios que Claude entrega.
const APP_VERSION = "3.03.07.040";

// Identidad del jugador en este dispositivo: se guarda en localStorage, así
// que persiste aunque cierres y vuelvas a abrir la app en el mismo celular.
// No es un login "de verdad" (no hay servidor de autenticación), es un PIN
// simple para que la app sepa quién eres tú y así poder distinguir al host
// del resto durante una partida.
const MY_ID_STORAGE_KEY = "bariloche_myPlayerId";
let _showIdentityModal = null;
function requestIdentity(roster) {
  return new Promise((resolve) => {
    if (!_showIdentityModal) { resolve(null); return; }
    _showIdentityModal(roster, resolve);
  });
}
function IdentityModal({ setRoster }) {
  const [pending, setPending] = useState(null); // { roster, onSubmit }
  const [pickedId, setPickedId] = useState("");
  const [pin, setPin] = useState("");
  const [newPin, setNewPin] = useState("");
  const [newPinConfirm, setNewPinConfirm] = useState("");
  const [err, setErr] = useState("");

  useEffect(() => {
    _showIdentityModal = (roster, onSubmit) => {
      setPickedId(""); setPin(""); setNewPin(""); setNewPinConfirm(""); setErr("");
      setPending({ roster, onSubmit });
    };
    return () => { _showIdentityModal = null; };
  }, []);

  if (!pending) return null;
  const players = pending.roster.filter((p) => p.active).sort((a, b) => a.name.localeCompare(b.name));
  const picked = players.find((p) => p.id === pickedId);

  const finish = (id) => {
    const cb = pending.onSubmit;
    setPending(null);
    cb(id);
  };
  const submitPin = () => {
    if (!picked) return;
    if (picked.pin && picked.pin !== pin) { setErr("PIN incorrecto."); return; }
    finish(picked.id);
  };
  // El Excel (hoja "roster") es la fuente de verdad de los PIN. Si este
  // jugador todavía no tiene uno ahí, lo obligamos a crearlo aquí mismo antes
  // de dejarlo entrar, para que quede registrado desde ya.
  const submitNewPin = () => {
    if (!picked) return;
    if (!/^\d{4}$/.test(newPin)) { setErr("El PIN debe ser numérico, de exactamente 4 dígitos."); return; }
    if (newPin !== newPinConfirm) { setErr("Los PIN no coinciden."); return; }
    setRoster((r) => r.map((p) => (p.id === picked.id ? { ...p, pin: newPin } : p)));
    finish(picked.id);
  };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.6)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 16 }} onClick={() => finish(null)}>
      <div style={{ background: C.panel, border: `1px solid ${C.panelLine}`, borderRadius: 14, padding: 20, width: "min(360px, 100%)", boxShadow: "0 12px 32px rgba(0,0,0,0.4)" }} onClick={(e) => e.stopPropagation()}>
        <div style={{ ...displayFont, fontSize: 19, color: C.card, marginBottom: 4 }}>¿QUIÉN ERES?</div>
        {!picked ? (
          <>
            <div style={{ fontSize: 12.5, color: "rgba(244,234,214,0.6)", marginBottom: 12 }}>
              Elige tu nombre para identificarte en este celular. Así la app sabe cuando eres tú el host de la partida.
            </div>
            <div style={{ display: "grid", gap: 6, maxHeight: 320, overflowY: "auto" }}>
              {players.map((p) => (
                <button key={p.id} onClick={() => setPickedId(p.id)}
                  style={{ display: "flex", alignItems: "center", gap: 10, textAlign: "left", background: "rgba(0,0,0,0.18)", border: `1px solid ${C.panelLine}`, borderRadius: 9, padding: "9px 10px", cursor: "pointer", ...bodyFont }}>
                  <Avatar player={p} size={26} />
                  <span style={{ color: C.card, fontSize: 14 }}>{p.name}</span>
                  {p.pin && <span title="Tiene PIN configurado" style={{ marginLeft: "auto", fontSize: 12, color: "rgba(244,234,214,0.35)" }}>🔒</span>}
                </button>
              ))}
            </div>
            <div style={{ marginTop: 14 }}>
              <GhostBtn onClick={() => finish(null)}>Cancelar</GhostBtn>
            </div>
          </>
        ) : (
          <>
            <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
              <Avatar player={picked} size={32} />
              <span style={{ color: C.card, fontSize: 16, fontWeight: 700 }}>{picked.name}</span>
            </div>
            {picked.pin ? (
              <>
                <div style={{ fontSize: 12.5, color: "rgba(244,234,214,0.6)", marginBottom: 8 }}>Ingresa tu PIN:</div>
                <input
                  type="password" inputMode="numeric" maxLength={4} autoFocus style={inputStyle}
                  value={pin} onChange={(e) => { setPin(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                  onKeyDown={(e) => { if (e.key === "Enter") submitPin(); if (e.key === "Escape") finish(null); }}
                />
                {err && <div style={{ color: C.loss, fontSize: 12, marginTop: 6 }}>{err}</div>}
                <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
                  <GhostBtn onClick={() => { setPickedId(""); setPin(""); setErr(""); }}>Atrás</GhostBtn>
                  <PrimaryBtn onClick={submitPin}>Entrar</PrimaryBtn>
                </div>
              </>
            ) : (
              <>
                <div style={{ fontSize: 12.5, color: "rgba(244,234,214,0.6)", marginBottom: 8 }}>
                  Todavía no tienes un PIN registrado. Crea uno de 4 dígitos para poder identificarte en tus celulares:
                </div>
                <div style={{ display: "grid", gap: 8 }}>
                  <input
                    type="password" inputMode="numeric" maxLength={4} autoFocus style={inputStyle} placeholder="Nuevo PIN (4 dígitos)"
                    value={newPin} onChange={(e) => { setNewPin(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                  />
                  <input
                    type="password" inputMode="numeric" maxLength={4} style={inputStyle} placeholder="Confirma tu PIN"
                    value={newPinConfirm} onChange={(e) => { setNewPinConfirm(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                    onKeyDown={(e) => { if (e.key === "Enter") submitNewPin(); if (e.key === "Escape") finish(null); }}
                  />
                </div>
                {err && <div style={{ color: C.loss, fontSize: 12, marginTop: 6 }}>{err}</div>}
                <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
                  <GhostBtn onClick={() => { setPickedId(""); setNewPin(""); setNewPinConfirm(""); setErr(""); }}>Atrás</GhostBtn>
                  <PrimaryBtn onClick={submitNewPin}>Crear PIN y entrar</PrimaryBtn>
                </div>
              </>
            )}
          </>
        )}
      </div>
    </div>
  );
}

// Pantalla completa (no un modal descartable): mientras nadie se identificó
// en este celular, no se puede navegar por la app en absoluto — a diferencia
// del modal de "cambiar identidad" (que sí se puede cancelar porque ahí ya
// estabas identificado), esta pantalla no tiene forma de saltearse.
function IdentityGateScreen({ roster, setRoster, onIdentified }) {
  const [pickedId, setPickedId] = useState("");
  const [pin, setPin] = useState("");
  const [newPin, setNewPin] = useState("");
  const [newPinConfirm, setNewPinConfirm] = useState("");
  const [err, setErr] = useState("");

  const players = roster.filter((p) => p.active).sort((a, b) => a.name.localeCompare(b.name));
  const picked = players.find((p) => p.id === pickedId);

  const submitPin = () => {
    if (!picked) return;
    if (picked.pin && picked.pin !== pin) { setErr("PIN incorrecto."); return; }
    onIdentified(picked.id);
  };
  // Igual que en el modal de identidad: si el jugador no tiene PIN en el
  // Excel, lo obligamos a crear uno antes de dejarlo pasar.
  const submitNewPin = () => {
    if (!picked) return;
    if (!/^\d{4}$/.test(newPin)) { setErr("El PIN debe ser numérico, de exactamente 4 dígitos."); return; }
    if (newPin !== newPinConfirm) { setErr("Los PIN no coinciden."); return; }
    setRoster((r) => r.map((p) => (p.id === picked.id ? { ...p, pin: newPin } : p)));
    onIdentified(picked.id);
  };

  return (
    <div style={{ minHeight: "100vh", background: `radial-gradient(ellipse at top, ${C.panel} 0%, ${C.felt} 45%, ${C.feltDeep} 100%)`, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, ...bodyFont }}>
      <div style={{ width: "min(420px, 100%)" }}>
        <div style={{ textAlign: "center", marginBottom: 18 }}>
          <div style={{ ...displayFont, fontSize: 36, color: C.gold, lineHeight: 1 }}>BARILOCHE</div>
          <div style={{ ...bodyFont, fontSize: 12, color: "rgba(240,216,136,0.55)", letterSpacing: "0.14em", textTransform: "uppercase", marginTop: 4 }}>
            registro de poker
          </div>
        </div>
        <div style={{ background: C.panel, border: `1px solid ${C.panelLine}`, borderRadius: 14, padding: 22, boxShadow: "0 12px 32px rgba(0,0,0,0.4)" }}>
          <div style={{ ...displayFont, fontSize: 22, color: C.card, marginBottom: 4 }}>¿QUIÉN ERES?</div>
          {!picked ? (
            <>
              <div style={{ fontSize: 13, color: "rgba(244,234,214,0.6)", marginBottom: 14 }}>
                Elige tu nombre para poder usar la app en este celular. Así la app sabe cuando eres tú el host de la partida.
              </div>
              <div style={{ display: "grid", gap: 8, maxHeight: 380, overflowY: "auto" }}>
                {players.map((p) => (
                  <button key={p.id} onClick={() => setPickedId(p.id)}
                    style={{ display: "flex", alignItems: "center", gap: 10, textAlign: "left", background: "rgba(0,0,0,0.18)", border: `1px solid ${C.panelLine}`, borderRadius: 9, padding: "10px 12px", cursor: "pointer", ...bodyFont }}>
                    <Avatar player={p} size={28} />
                    <span style={{ color: C.card, fontSize: 15 }}>{p.name}</span>
                    {p.pin && <span title="Tiene PIN configurado" style={{ marginLeft: "auto", fontSize: 13, color: "rgba(244,234,214,0.35)" }}>🔒</span>}
                  </button>
                ))}
              </div>
            </>
          ) : (
            <>
              <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12 }}>
                <Avatar player={picked} size={34} />
                <span style={{ color: C.card, fontSize: 17, fontWeight: 700 }}>{picked.name}</span>
              </div>
              {picked.pin ? (
                <>
                  <div style={{ fontSize: 13, color: "rgba(244,234,214,0.6)", marginBottom: 8 }}>Ingresa tu PIN:</div>
                  <input
                    type="password" inputMode="numeric" maxLength={4} autoFocus style={inputStyle}
                    value={pin} onChange={(e) => { setPin(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                    onKeyDown={(e) => { if (e.key === "Enter") submitPin(); }}
                  />
                  {err && <div style={{ color: C.loss, fontSize: 12, marginTop: 6 }}>{err}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
                    <GhostBtn onClick={() => { setPickedId(""); setPin(""); setErr(""); }}>Atrás</GhostBtn>
                    <PrimaryBtn onClick={submitPin}>Entrar</PrimaryBtn>
                  </div>
                </>
              ) : (
                <>
                  <div style={{ fontSize: 13, color: "rgba(244,234,214,0.6)", marginBottom: 8 }}>
                    Todavía no tienes un PIN registrado. Crea uno de 4 dígitos para poder identificarte en tus celulares:
                  </div>
                  <div style={{ display: "grid", gap: 8 }}>
                    <input
                      type="password" inputMode="numeric" maxLength={4} autoFocus style={inputStyle} placeholder="Nuevo PIN (4 dígitos)"
                      value={newPin} onChange={(e) => { setNewPin(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                    />
                    <input
                      type="password" inputMode="numeric" maxLength={4} style={inputStyle} placeholder="Confirma tu PIN"
                      value={newPinConfirm} onChange={(e) => { setNewPinConfirm(e.target.value.replace(/[^\d]/g, "").slice(0, 4)); setErr(""); }}
                      onKeyDown={(e) => { if (e.key === "Enter") submitNewPin(); }}
                    />
                  </div>
                  {err && <div style={{ color: C.loss, fontSize: 12, marginTop: 6 }}>{err}</div>}
                  <div style={{ display: "flex", gap: 8, marginTop: 16, justifyContent: "flex-end" }}>
                    <GhostBtn onClick={() => { setPickedId(""); setNewPin(""); setNewPinConfirm(""); setErr(""); }}>Atrás</GhostBtn>
                    <PrimaryBtn onClick={submitNewPin}>Crear PIN y entrar</PrimaryBtn>
                  </div>
                </>
              )}
            </>
          )}
        </div>
      </div>
    </div>
  );
}

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

    // Ajuste manual capturado al entregar fichas: corrige las fichas
    // realmente entregadas (un billete/ficha mal contado, etc.), sin tocar
    // el campo principal de fichas ni el cuadre total de dinero de la
    // partida. Como las fichas entregadas son lo que primero salda el
    // buy-in virtual, el ajuste impacta ese cálculo (y por lo tanto cuánto
    // termina cobrando en cash vs. transferencia).
    const cashAdjust = Number((game.finalChipsAdjust || {})[pid]) || 0;
    const fichasAjustadas = round1(cashOut + cashAdjust);

    // Regla: el cash out (ya ajustado) primero salda el buy-in virtual. Lo
    // que sobra de eso ("netClaim") es lo que el jugador realmente puede
    // reclamar del pozo de cash real — no el balance total. Si netClaim <= 0,
    // ni siquiera alcanzó para saldar el virtual, y esa diferencia se debe
    // por transferencia.
    const netClaim = round1(fichasAjustadas - virtualAmount);

    let pagoCash = 0;
    let pagoTransfer = 0;
    if (netClaim > 0) {
      // Puede cobrar en cash hasta lo que él mismo puso en cash; el resto
      // del reclamo (si lo hay) se cobra por transferencia.
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

  // Identidad del dispositivo: quién eres tú, persistida en localStorage.
  const [myPlayerId, setMyPlayerIdState] = useState(() => {
    try { return localStorage.getItem(MY_ID_STORAGE_KEY) || ""; } catch { return ""; }
  });
  const setMyPlayerId = (id) => {
    setMyPlayerIdState(id || "");
    try { if (id) localStorage.setItem(MY_ID_STORAGE_KEY, id); else localStorage.removeItem(MY_ID_STORAGE_KEY); } catch {}
  };
  const identify = async () => {
    const id = await requestIdentity(roster);
    if (id) setMyPlayerId(id);
  };
  const logout = async () => {
    const ok = await requestConfirm({
      title: "¿Cerrar tu sesión en este celular?",
      message: "Vas a dejar de estar identificado en este dispositivo. La próxima vez que quieras hacer cambios durante una partida, vas a tener que volver a identificarte con tu PIN.",
      confirmLabel: "Sí, cerrar sesión",
      cancelLabel: "Cancelar",
      danger: false,
    });
    if (ok) setMyPlayerId("");
  };
  const hasActive = !!activeGame && !activeGame.finished;
  const isHost = hasActive && !!myPlayerId && myPlayerId === activeGame.hostId;

  // Si una partida en curso termina o se cancela mientras alguien está parado
  // en una de las pestañas exclusivas de "partida en curso" (Cena, Lote y
  // Rake, o Jugadores-de-la-partida), lo mandamos de vuelta a la pestaña
  // principal para que no quede en una pestaña que ya no existe. Lo mismo si
  // quien no es host queda parado en una pestaña que ahora es solo del host.
  useEffect(() => {
    if (!hasActive && (tab === "cena" || tab === "loterake" || tab === "compra" || tab === "finalizar")) setTab("partida");
    if (hasActive && tab === "historial") setTab("partida");
    if (hasActive && !isHost && (tab === "cena" || tab === "loterake" || tab === "jugadores" || tab === "finalizar")) setTab("partida");
  }, [hasActive, isHost, tab]);

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

  // Mientras este dispositivo tiene un cambio propio recién guardándose, el
  // polling (más abajo) no debe pisarlo con lo que todavía estaba en el
  // Excel un instante antes — si no, se ve como que el cambio "no pegó" por
  // un segundo. Cada guardado local extiende su propio "candado" unos
  // segundos; pasado ese tiempo, el polling ya puede volver a sincronizar
  // libremente desde el Excel (por ejemplo, para traer los cambios que hizo
  // OTRO dispositivo).
  const rosterLockUntil = useRef(0);
  const gamesLockUntil = useRef(0);
  const activeLockUntil = useRef(0);
  const LOCK_MS = 3000;

  useEffect(() => {
    if (loading) return;
    if (skipNextRosterSave.current) { skipNextRosterSave.current = false; return; }
    rosterLockUntil.current = Date.now() + LOCK_MS;
    saveKey(KEYS.roster, roster);
  }, [roster, loading]);
  useEffect(() => {
    if (loading) return;
    if (skipNextGamesSave.current) { skipNextGamesSave.current = false; return; }
    gamesLockUntil.current = Date.now() + LOCK_MS;
    saveKey(KEYS.games, games);
  }, [games, loading]);
  useEffect(() => {
    if (loading) return;
    if (skipNextActiveSave.current) { skipNextActiveSave.current = false; return; }
    activeLockUntil.current = Date.now() + LOCK_MS;
    saveKey(KEYS.active, activeGame);
  }, [activeGame, loading]);

  // Este ledger no tiene tiempo real (no hay websockets): cada celular
  // guarda sus cambios al Excel, pero para ENTERARSE de lo que hicieron los
  // demás dispositivos (el host inició la partida, otro jugador pidió
  // fichas, el host la aceptó, etc.) hace falta ir a preguntarle al Excel
  // de vez en cuando. Sin este polling, cada quien se quedaba viendo una
  // foto congelada de cuando entró a la app.
  const POLL_MS = 4000;
  useEffect(() => {
    if (loading) return;
    const interval = setInterval(async () => {
      try {
        const [r, g, a] = await Promise.all([
          loadKey(KEYS.roster, []),
          loadKey(KEYS.games, []),
          loadKey(KEYS.active, null),
        ]);
        const now = Date.now();
        if (now >= rosterLockUntil.current) {
          setRoster((prev) => {
            if (JSON.stringify(prev) === JSON.stringify(r)) return prev;
            skipNextRosterSave.current = true;
            return r;
          });
        }
        if (now >= gamesLockUntil.current) {
          setGames((prev) => {
            if (JSON.stringify(prev) === JSON.stringify(g)) return prev;
            skipNextGamesSave.current = true;
            return g;
          });
        }
        if (now >= activeLockUntil.current) {
          setActiveGame((prev) => {
            if (JSON.stringify(prev) === JSON.stringify(a)) return prev;
            skipNextActiveSave.current = true;
            return a;
          });
        }
      } catch {
        // Si falla una vuelta de polling (por ejemplo, sin conexión un
        // instante), simplemente se reintenta en la siguiente — no vale la
        // pena molestar al usuario por esto.
      }
    }, POLL_MS);
    return () => clearInterval(interval);
  }, [loading]);

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

  // Sin identificarte, no se puede navegar por ninguna parte de la app.
  // Excepción: si el roster todavía está vacío (recién desplegada, sin
  // jugadores dados de alta), no hay nadie para elegir — se deja pasar para
  // que se pueda cargar al primer jugador desde la pestaña Jugadores.
  if (roster.length > 0 && !myPlayerId) {
    return <IdentityGateScreen roster={roster} setRoster={setRoster} onIdentified={setMyPlayerId} />;
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

      <Header tab={tab} setTab={setTab} hasActive={hasActive} isHost={isHost} me={roster.find((p) => p.id === myPlayerId) || null} onIdentify={identify} onLogout={logout} />

      <main style={{ maxWidth: 980, margin: "0 auto", padding: "18px 14px 60px" }}>
        {tab === "jugadores" && !hasActive && (
          <PlayersTab roster={roster} setRoster={setRoster} playerStats={playerStats} adminPassword={adminPassword} myPlayerId={myPlayerId} />
        )}
        {(tab === "partida" || tab === "cena" || tab === "loterake" || tab === "compra" || tab === "finalizar" || (tab === "jugadores" && hasActive)) && (
          <GameTab
            roster={roster}
            activeGame={activeGame}
            setActiveGame={setActiveGame}
            games={games}
            setGames={setGames}
            myPlayerId={myPlayerId}
            onIdentify={identify}
            adminPassword={adminPassword}
            setTab={setTab}
            subView={tab === "cena" ? "cena" : tab === "loterake" ? "loterake" : tab === "jugadores" ? "jugadoresPartida" : tab === "compra" ? "compra" : tab === "finalizar" ? "finalizar" : "estatus"}
          />
        )}
        {tab === "historial" && <HistoryTab games={games} roster={roster} setGames={setGames} adminPassword={adminPassword} activeGame={activeGame} setActiveGame={setActiveGame} />}
      </main>

      <AdminPasswordModal />
      <ConfirmModal />
      <IdentityModal setRoster={setRoster} />
    </div>
  );
}

/* ----------------------------------------------------------------------
   HEADER / TABS
---------------------------------------------------------------------- */
function Header({ tab, setTab, hasActive, isHost, me, onIdentify, onLogout }) {
  const tabs = hasActive
    ? (isHost
        ? [
            { id: "partida", label: "Estatus jugada", icon: Activity },
            { id: "compra", label: "Compra de lotes", icon: Banknote },
            { id: "jugadores", label: "Jugadores", icon: Users },
            { id: "cena", label: "Cena y servicio", icon: UtensilsCrossed },
            { id: "loterake", label: "Lote y Rakes", icon: Coins },
          ]
        : [
            { id: "partida", label: "Estatus jugada", icon: Activity },
            { id: "compra", label: "Compra de lotes", icon: Banknote },
          ])
    : [
        { id: "partida", label: "Partida", icon: Flame },
        { id: "jugadores", label: "Jugadores", icon: Users },
        { id: "historial", label: "Histórico", icon: History },
      ];
  const handleIdentityClick = () => { if (me) onLogout(); else onIdentify(); };
  return (
    <header style={{ borderBottom: `1px solid ${C.panelLine}`, background: "rgba(0,0,0,0.15)", position: "sticky", top: 0, zIndex: 20, backdropFilter: "blur(6px)" }}>
      <div style={{ maxWidth: 980, margin: "0 auto", padding: "14px 14px 0" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10 }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 1 }}>
            <div style={{ display: "flex", alignItems: "baseline", gap: 10 }}>
              <span style={{ ...displayFont, fontSize: 30, color: C.gold, lineHeight: 1 }}>BARILOCHE</span>
              <span style={{ ...bodyFont, fontSize: 12, color: "rgba(240,216,136,0.55)", letterSpacing: "0.14em", textTransform: "uppercase" }}>
                registro de poker
              </span>
            </div>
            <span style={{ ...monoFont, fontSize: 10, color: "rgba(244,234,214,0.35)" }}>v{APP_VERSION}</span>
          </div>
          {/* Página aparte para el organizador: no es un tab más, así los
              jugadores no se topan de casualidad con rachas o comparativas. */}
          <div style={{ display: "flex", alignItems: "center", gap: 6, flexShrink: 0 }}>
            <button
              onClick={handleIdentityClick} title={me ? "Cerrar tu sesión en este celular" : "Identifícate para poder ser host"}
              style={{
                display: "flex", alignItems: "center", gap: 6, cursor: "pointer",
                background: me ? "rgba(0,0,0,0.2)" : "rgba(216,173,63,0.16)",
                border: `1px solid ${me ? C.panelLine : C.gold}`, borderRadius: 8, padding: "5px 9px",
                color: me ? "rgba(244,234,214,0.75)" : C.goldSoft, fontSize: 12, fontWeight: 600, ...bodyFont,
              }}
            >
              {me ? (
                <>
                  <Avatar player={me} size={17} />
                  <span>{me.name}</span>
                </>
              ) : (
                <>👤 ¿Quién eres?</>
              )}
            </button>
            <a
              href="/dashboard.html"
              target="_blank"
              rel="noopener noreferrer"
              title="Ver estadísticas históricas"
              style={{
                display: "flex", alignItems: "center", gap: 6, textDecoration: "none",
                border: `1px solid ${C.panelLine}`, borderRadius: 8, padding: "6px 10px",
                color: "rgba(244,234,214,0.65)", fontSize: 12, fontWeight: 600, ...bodyFont,
              }}
            >
              <BarChart3 size={13} /> Estadísticas
            </a>
          </div>
        </div>
        {hasActive && (
          <div style={{
            display: "flex", alignItems: "center", gap: 8, marginTop: 12,
            background: "rgba(216,173,63,0.14)", border: `1px solid ${C.gold}`,
            borderRadius: 10, padding: "9px 12px",
          }}>
            <span style={{ width: 8, height: 8, borderRadius: 99, background: C.win, flexShrink: 0, boxShadow: "0 0 0 3px rgba(63,191,114,0.25)" }} />
            <span style={{ ...displayFont, fontSize: 15, color: C.goldSoft, letterSpacing: "0.04em" }}>Jugada en Curso</span>
          </div>
        )}
        <div className="scrollbar-thin" style={{ display: "flex", gap: 4, marginTop: 12, flexWrap: "nowrap", overflowX: "auto", overflowY: "hidden" }}>
          {tabs.map((t) => {
            const Icon = t.icon;
            const active = tab === t.id;
            return (
              <button
                key={t.id}
                onClick={() => setTab(t.id)}
                style={{
                  display: "flex", alignItems: "center", gap: 6, flexShrink: 0, whiteSpace: "nowrap",
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
// Campo para montos en pesos (lote, rake, cena, etc.): mientras se está
// editando se ve el número "pelón" para no pelear con el cursor, pero en
// cuanto se sale del campo (blur) se muestra formateado como $#,##0.
function MoneyInput({ value, onChange, disabled, placeholder, style }) {
  const [focused, setFocused] = useState(false);
  const display = focused
    ? (value || value === 0 ? String(value) : "")
    : (value || value === 0 ? money(value) : "");
  return (
    <input
      type="text" inputMode="numeric"
      disabled={disabled}
      style={{ ...inputStyle, opacity: disabled ? 0.55 : 1, cursor: disabled ? "not-allowed" : "text", ...style }}
      placeholder={placeholder}
      value={display}
      onFocus={(e) => { setFocused(true); requestAnimationFrame(() => e.target.select()); }}
      onChange={(e) => {
        const digits = e.target.value.replace(/[^\d]/g, "");
        onChange(digits === "" ? 0 : Number(digits));
      }}
      onBlur={() => setFocused(false)}
    />
  );
}
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
function PlayersTab({ roster, setRoster, playerStats, adminPassword, myPlayerId }) {
  const [selId, setSelId] = useState("");
  const [newName, setNewName] = useState("");
  const [newAvatar, setNewAvatar] = useState(null);
  const [editName, setEditName] = useState("");
  const [editAvatar, setEditAvatar] = useState(null);
  const [editPin, setEditPin] = useState("");
  const [pinMsg, setPinMsg] = useState("");
  const [savedMsg, setSavedMsg] = useState(false);

  const sel = roster.find((p) => p.id === selId) || null;
  useEffect(() => {
    setEditName(sel ? sel.name : ""); setEditAvatar(sel ? sel.avatar || null : null);
    setEditPin(sel ? sel.pin || "" : ""); setPinMsg("");
  }, [selId]); // eslint-disable-line

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
    setSelId(""); // limpia el combo — vuelve a "— elegir del combo —"
    setSavedMsg(true);
    setTimeout(() => setSavedMsg(false), 2500);
  };
  const savePin = async () => {
    if (!sel) return;
    const trimmed = editPin.trim();
    if (trimmed && !/^\d{4}$/.test(trimmed)) { setPinMsg("El PIN debe ser numérico, de exactamente 4 dígitos."); return; }
    const isSelf = !!myPlayerId && myPlayerId === sel.id;
    if (!isSelf) {
      if (!(await requestAdminPassword(adminPassword, `cambiar el PIN de ${sel.name}`))) return;
    }
    setRoster((r) => r.map((p) => (p.id === sel.id ? { ...p, pin: trimmed } : p)));
    setPinMsg(trimmed ? "PIN guardado." : "PIN eliminado — este jugador queda sin protección al identificarse.");
    setTimeout(() => setPinMsg(""), 3000);
  };
  const toggleActive = async () => {
    if (!sel) return;
    if (sel.active) {
      // Solo pedimos contraseña para dar de baja (acción sensible), no para reactivar.
      if (!(await requestAdminPassword(adminPassword, "dar de baja jugadores"))) return;
    }
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
        {savedMsg && (
          <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 10, background: "rgba(63,191,114,0.14)", border: `1px solid ${C.win}`, borderRadius: 8, padding: "8px 10px" }}>
            <Check size={14} color={C.win} />
            <span style={{ fontSize: 12.5, color: C.win, fontWeight: 700 }}>Jugador actualizado correctamente.</span>
          </div>
        )}
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
            <Field label="PIN (4 dígitos, para identificarte en un celular)">
              <div style={{ display: "flex", gap: 8 }}>
                <input
                  type="password" inputMode="numeric" maxLength={4} style={inputStyle} placeholder="Sin PIN configurado" value={editPin}
                  onChange={(e) => setEditPin(e.target.value.replace(/[^\d]/g, "").slice(0, 4))}
                />
                <GhostBtn onClick={savePin} icon={Check} color={C.gold}>Guardar PIN</GhostBtn>
              </div>
              {pinMsg && <div style={{ fontSize: 11.5, color: pinMsg.includes("eliminado") ? C.goldSoft : C.win, marginTop: 5 }}>{pinMsg}</div>}
            </Field>
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
function GameTab({ roster, activeGame, setActiveGame, games, setGames, myPlayerId, onIdentify, adminPassword, setTab, subView }) {
  if (!activeGame) return <NewGameSetup roster={roster} setActiveGame={setActiveGame} />;
  if (activeGame.finished) return <FinalizedGame game={activeGame} roster={roster} onClose={() => setActiveGame(null)} setActiveGame={setActiveGame} setGames={setGames} adminPassword={adminPassword} />;
  const isHost = !!myPlayerId && myPlayerId === activeGame.hostId;
  return (
    <ActiveGameScreen
      game={activeGame} setGame={setActiveGame} roster={roster} setGames={setGames}
      isHost={isHost} onIdentify={onIdentify} myPlayerId={myPlayerId}
      view={subView} setTab={setTab}
    />
  );
}

function NewGameSetup({ roster, setActiveGame }) {
  const active = roster.filter((p) => p.active);
  const [date, setDate] = useState(todayISO());
  const [loteValue, setLoteValue] = useState(1000);
  const [rakeHost, setRakeHost] = useState(1500);
  const [rakeAutosCount, setRakeAutosCount] = useState(0);
  const [rakeAutoAmount, setRakeAutoAmount] = useState(250);
  const [selected, setSelected] = useState([]);
  const [hostId, setHostId] = useState("");

  const rakeTotal = round1((Number(rakeHost) || 0) + (Number(rakeAutosCount) || 0) * (Number(rakeAutoAmount) || 0));

  const toggle = (id) => {
    setSelected((s) => (s.includes(id) ? s.filter((x) => x !== id) : [...s, id]));
    if (hostId === id) setHostId("");
  };

  const start = () => {
    if (selected.length < 2 || !loteValue || !hostId) return;
    setActiveGame({
      id: uid(), date, loteValue: Number(loteValue),
      rakeHost: Number(rakeHost) || 0,
      rakeAutosCount: Number(rakeAutosCount) || 0,
      rakeAutoAmount: Number(rakeAutoAmount) || 0,
      rake: rakeTotal,
      playerIds: selected, hostId, startedAt: Date.now(),
      purchases: [],
      requests: [],
      dinner: { amountNoAlcohol: 0, amountAlcohol: 0, alcohol: {}, paid: {}, paymentMethod: {} },
      finalChips: {}, finalizeDraft: {}, finished: false, results: null,
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
            <MoneyInput value={loteValue} onChange={setLoteValue} />
          </Field>
        </div>
        <div style={{ marginTop: 16, borderTop: `1px solid ${C.panelLine}`, paddingTop: 12 }}>
          <div style={{ ...displayFont, fontSize: 15, color: C.goldSoft, marginBottom: 8, letterSpacing: "0.05em" }}>RAKE</div>
          <Field label="Rake para anfitrión">
            <MoneyInput value={rakeHost} onChange={setRakeHost} />
          </Field>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
            <Field label="Rake para autos — # de autos">
              <input type="number" min="0" style={inputStyle} value={rakeAutosCount === 0 ? "" : rakeAutosCount} onChange={(e) => setRakeAutosCount(e.target.value === "" ? 0 : Number(e.target.value))} onFocus={(e) => e.target.select()} />
            </Field>
            <Field label="Monto por auto">
              <MoneyInput value={rakeAutoAmount} onChange={setRakeAutoAmount} />
            </Field>
          </div>
          <div style={{ marginTop: 10 }}>
            <ScoreBox label="Rake total" value={money(rakeTotal)} />
          </div>
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
function formatClock(ts) {
  if (!ts) return "—";
  return new Date(ts).toLocaleTimeString("es-MX", { hour: "2-digit", minute: "2-digit", hour12: false });
}
function GameStatusBlock({ game, players, totals }) {
  const perPlayer = players.map((p) => {
    const cash = game.purchases.filter((pu) => pu.playerId === p.id && pu.type === "cash").reduce((s, pu) => s + pu.amount, 0);
    const virtual = game.purchases.filter((pu) => pu.playerId === p.id && pu.type === "virtual").reduce((s, pu) => s + pu.amount, 0);
    return { player: p, cash, virtual, total: cash + virtual };
  }).sort((a, b) => a.player.name.localeCompare(b.player.name));

  return (
    <Panel>
      <SectionTitle icon={Activity}>Estatus de la jugada</SectionTitle>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 8, marginBottom: 14 }}>
        <ScoreBox label="Total cash" value={money(totals.cash)} tone="cash" />
        <ScoreBox label="Total virtual" value={money(totals.virtual)} tone="virtual" />
        <ScoreBox label="Total" value={money(totals.cash + totals.virtual)} />
        <ScoreBox label="Hora inicio" value={formatClock(game.startedAt)} />
      </div>
      <div style={{ display: "grid", gap: 8 }}>
        {perPlayer.map((row) => (
          <div key={row.player.id} style={{ background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: "10px 12px" }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 8 }}>
              <Avatar player={row.player} size={22} />
              <span style={{ color: C.card, fontWeight: 700, fontSize: 14.5, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{row.player.name}</span>
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 4, textAlign: "center" }}>
              <div>
                <div style={{ fontSize: 9.5, color: "rgba(244,234,214,0.4)", textTransform: "uppercase", letterSpacing: "0.04em" }}>Cash</div>
                <div style={{ ...monoFont, fontWeight: 800, fontSize: 19, color: C.cash }}>{money(row.cash)}</div>
              </div>
              <div>
                <div style={{ fontSize: 9.5, color: "rgba(244,234,214,0.4)", textTransform: "uppercase", letterSpacing: "0.04em" }}>Virtual</div>
                <div style={{ ...monoFont, fontWeight: 800, fontSize: 19, color: C.virtual }}>{money(row.virtual)}</div>
              </div>
              <div>
                <div style={{ fontSize: 9.5, color: "rgba(244,234,214,0.4)", textTransform: "uppercase", letterSpacing: "0.04em" }}>Total</div>
                <div style={{ ...monoFont, fontWeight: 800, fontSize: 20, color: C.goldSoft }}>{money(row.total)}</div>
              </div>
            </div>
          </div>
        ))}
      </div>
    </Panel>
  );
}

function ActiveGameScreen({ game, setGame, roster, setGames, isHost, onIdentify, myPlayerId, view, setTab }) {
  // "Finalizar partida" es una pantalla más, seleccionada igual que cualquier
  // otra pestaña (a través de "view", que en definitiva viene del estado
  // "tab" de más arriba) — así, tocar cualquier pestaña del Header mientras
  // se está finalizando la partida navega ahí directamente en vez de quedar
  // atascado en esta pantalla (antes esto vivía en un estado local aparte,
  // que se podía desincronizar de la pestaña activa).
  const effectiveView = view || "estatus";

  const players = game.playerIds.map((id) => roster.find((r) => r.id === id)).filter(Boolean);
  const availableToAdd = roster.filter((p) => p.active && !game.playerIds.includes(p.id));

  const totals = useMemo(() => {
    let cash = 0, virtual = 0;
    game.purchases.forEach((p) => { if (p.type === "cash") cash += p.amount; else virtual += p.amount; });
    return { cash, virtual };
  }, [game.purchases]);

  // update acepta un objeto (mezcla directa) o una función (g) => patch, que
  // recibe el estado MÁS RECIENTE del juego. Usar la forma función evita
  // "carreras" cuando se disparan varios cambios rápido seguidos (por
  // ejemplo, tocar +Cash varias veces, o editar dos campos de la cena
  // rápido): cada patch se arma sobre el estado real, no sobre una copia
  // vieja capturada en el momento del render.
  const update = (patch) =>
    setGame((g) => ({ ...g, ...(typeof patch === "function" ? patch(g) : patch) }));

  // Solicitudes de fichas: cualquier jugador identificado en la mesa puede
  // pedir un lote (cash o virtual) o un monto libre; solo el host puede
  // aceptarla (se vuelve una compra normal) o rechazarla.
  const myRequests = (game.requests || []).filter((r) => r.playerId === myPlayerId);
  const submitRequest = (type, amount) => {
    const amt = Math.round(Number(amount) || 0);
    if (!myPlayerId || amt <= 0) return;
    update((g) => ({ requests: [...(g.requests || []), { id: uid(), playerId: myPlayerId, type, amount: amt, status: "pending", ts: Date.now() }] }));
  };
  // Alguien que ya se identificó con su PIN pero todavía no está sentado en
  // esta mesa (no llegó a tiempo para el alta inicial, o se sumó a mitad de
  // la noche) no puede simplemente entrar solo: le manda al host una
  // solicitud de "agrégame a la mesa", igual que cuando pide fichas.
  const submitJoinRequest = () => {
    if (!myPlayerId) return;
    update((g) => {
      const already = (g.requests || []).some((r) => r.playerId === myPlayerId && r.type === "join" && r.status === "pending");
      if (already) return {};
      return { requests: [...(g.requests || []), { id: uid(), playerId: myPlayerId, type: "join", amount: 0, status: "pending", ts: Date.now() }] };
    });
  };
  const resolveRequest = (reqId, approve) => {
    update((g) => {
      const req = (g.requests || []).find((r) => r.id === reqId);
      if (!req || req.status !== "pending") return {};
      if (approve) {
        if (req.type === "join") {
          return {
            playerIds: g.playerIds.includes(req.playerId) ? g.playerIds : [...g.playerIds, req.playerId],
            requests: g.requests.map((r) => (r.id === reqId ? { ...r, status: "approved" } : r)),
          };
        }
        const entry = { id: uid(), playerId: req.playerId, type: req.type, lotes: g.loteValue ? round1(req.amount / g.loteValue) : 0, amount: req.amount, ts: Date.now() };
        return {
          purchases: [...g.purchases, entry],
          requests: g.requests.map((r) => (r.id === reqId ? { ...r, status: "approved" } : r)),
        };
      }
      return { requests: g.requests.map((r) => (r.id === reqId ? { ...r, status: "rejected" } : r)) };
    });
  };

  // Solo el host de la partida puede modificar algo (comprar lotes, tocar la
  // cena, cambiar configuración, cerrarla). Cualquier otro dispositivo ve
  // exactamente el mismo bloque de estatus (y puede pedir fichas), pero de
  // solo lectura para todo lo demás.
  if (!isHost) {
    const hostPlayer = roster.find((r) => r.id === game.hostId);
    const iAmInGame = players.some((p) => p.id === myPlayerId);
    // El aviso de "modo lectura" solo tiene sentido al entrar — en cuanto
    // el jugador navega a otra pestaña, ya lo entendió y estorba tenerlo
    // repitiéndose en cada pantalla.
    const firstViewRef = useRef(effectiveView);
    const [bannerDismissed, setBannerDismissed] = useState(false);
    useEffect(() => {
      if (effectiveView !== firstViewRef.current) setBannerDismissed(true);
    }, [effectiveView]);
    const banner = !bannerDismissed && (
      <div style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(226,99,79,0.12)", border: `1px solid ${C.loss}`, borderRadius: 10, padding: "8px 12px" }}>
        <AlertCircle size={14} color={C.loss} style={{ flexShrink: 0 }} />
        <div style={{ fontSize: 12, color: "rgba(244,234,214,0.85)" }}>
          Solo <strong>{hostPlayer ? hostPlayer.name : "el host"}</strong> puede modificar la partida — modo lectura.
        </div>
      </div>
    );

    if (effectiveView === "cena") {
      return (
        <div style={{ display: "grid", gap: 16 }}>
          {banner}
          <Panel>
            <SectionTitle icon={UtensilsCrossed}>Cena y servicio</SectionTitle>
            <DinnerReadOnly game={game} players={players} />
          </Panel>
        </div>
      );
    }
    if (effectiveView === "loterake") {
      return (
        <div style={{ display: "grid", gap: 16 }}>
          {banner}
          <Panel>
            <SectionTitle icon={Coins}>Lote y Rakes</SectionTitle>
            <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 8 }}>
              <ScoreBox label="Valor de lote" value={money(game.loteValue)} />
              <ScoreBox label="Rake total" value={money(game.rake)} />
              <ScoreBox label="Rake anfitrión" value={money(game.rakeHost || 0)} />
              <ScoreBox label={`Rake autos (${game.rakeAutosCount || 0} × ${money(game.rakeAutoAmount ?? 250)})`} value={money((game.rakeAutosCount || 0) * (game.rakeAutoAmount ?? 250))} />
            </div>
          </Panel>
        </div>
      );
    }
    if (effectiveView === "jugadoresPartida") {
      return (
        <div style={{ display: "grid", gap: 16 }}>
          {banner}
          <Panel>
            <SectionTitle icon={Users}>Jugadores en la mesa ({players.length})</SectionTitle>
            <div style={{ display: "grid", gap: 8 }}>
              {players.map((p) => (
                <div key={p.id} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(0,0,0,0.18)", borderRadius: 9, padding: "8px 10px" }}>
                  <Avatar player={p} size={24} />
                  <span style={{ color: C.card, fontWeight: 600, fontSize: 13.5, flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{p.name}</span>
                  {game.hostId === p.id && <Badge tone="gold"><Crown size={10} /> Host</Badge>}
                </div>
              ))}
            </div>
          </Panel>
        </div>
      );
    }

    if (effectiveView === "compra") {
      return (
        <div style={{ display: "grid", gap: 16 }}>
          {banner}
          {iAmInGame ? (
            <RequestChipsPanel loteValue={game.loteValue} myRequests={myRequests} onSubmit={submitRequest} />
          ) : (
            myPlayerId && <JoinRequestPanel myRequests={myRequests} onSubmit={submitJoinRequest} />
          )}
        </div>
      );
    }

    return (
      <div style={{ display: "grid", gap: 16 }}>
        {banner}
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
          </div>
          <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 8, marginTop: 14 }}>
            <ScoreBox label="Monto del lote" value={money(game.loteValue)} />
            <ScoreBox label="Rake + estacionamiento" value={money(game.rake)} />
          </div>
        </Panel>
        <GameStatusBlock game={game} players={players} totals={totals} />
      </div>
    );
  }

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

  const setLote = (v) => update({ loteValue: Number(v) || 0 });
  // El rake queda fijo apenas se empieza a capturar la entrega de fichas —
  // ya sea que se haya llegado a cerrar la partida (finalChips) o que
  // todavía esté en borrador en la pantalla de "Entrega de fichas"
  // (finalizeDraft), que ahora se guarda solo en cuanto se escribe algo,
  // sin esperar a tocar "Volver".
  const hasFinalizeDraft = Object.values(game.finalizeDraft || {}).some(
    (d) => d && ((d.pagaVirtual !== undefined && d.pagaVirtual !== "") || (d.remanente !== undefined && d.remanente !== ""))
  );
  const rakeLocked = Object.keys(game.finalChips || {}).length > 0 || hasFinalizeDraft;
  const setRakeParts = (patch) =>
    update((g) => {
      const rakeHost = patch.rakeHost !== undefined ? Number(patch.rakeHost) || 0 : (Number(g.rakeHost) || 0);
      const rakeAutosCount = patch.rakeAutosCount !== undefined ? Number(patch.rakeAutosCount) || 0 : (Number(g.rakeAutosCount) || 0);
      const rakeAutoAmount = patch.rakeAutoAmount !== undefined ? Number(patch.rakeAutoAmount) || 0 : (g.rakeAutoAmount ?? 250);
      return { rakeHost, rakeAutosCount, rakeAutoAmount, rake: round1(rakeHost + rakeAutosCount * rakeAutoAmount) };
    });
  const setRakeHost = (v) => setRakeParts({ rakeHost: v });
  const setRakeAutosCount = (v) => setRakeParts({ rakeAutosCount: v });
  const setRakeAutoAmount = (v) => setRakeParts({ rakeAutoAmount: v });

  const cancelPartida = async () => {
    const ok = await requestConfirm({
      title: "¿Cancelar esta partida?",
      message: "Se va a perder todo el progreso de esta partida (lotes comprados, cena, configuración). Esta acción no se puede deshacer.",
      confirmLabel: "Sí, cancelar partida",
      cancelLabel: "Seguir jugando",
    });
    if (ok) setGame(null);
  };

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

  if (effectiveView === "finalizar") {
    return (
      <FinalizeGame
        game={game} roster={roster} update={update}
        onBack={() => setTab && setTab("compra")}
        onConfirm={(finalChips, finalChipsAdjust) => {
          const g2 = { ...game, finalChips, finalChipsAdjust };
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

  if (effectiveView === "cena") {
    return (
      <div style={{ display: "grid", gap: 16 }}>
        <Panel>
          <SectionTitle icon={UtensilsCrossed}>Cena y servicio</SectionTitle>
          <DinnerSection game={game} players={players} update={update} />
        </Panel>
      </div>
    );
  }

  if (effectiveView === "compra") {
    return (
      <div style={{ display: "grid", gap: 16 }}>
        <PendingRequestsPanel game={game} roster={roster} onResolve={resolveRequest} />
        <Panel>
          <SectionTitle icon={Banknote}>Compra de lotes por jugador</SectionTitle>
          <div style={{ display: "grid", gap: 10 }}>
            {players.map((p) => (
              <PlayerBuyRow
                key={p.id} player={p} game={game} onAdd={addPurchase} onRemoveLast={removeLastPurchase}
                dinnerPaid={!!game.dinner.paid?.[p.id]}
                onGoToDinner={() => setTab && setTab("cena")}
              />
            ))}
          </div>
        </Panel>
        <PrimaryBtn onClick={() => setTab && setTab("finalizar")} icon={Square} style={{ padding: "13px 18px", fontSize: 15 }}>
          Finalizar partida
        </PrimaryBtn>
      </div>
    );
  }

  if (effectiveView === "loterake") {
    return (
      <div style={{ display: "grid", gap: 16 }}>
        <Panel>
          <div>
            <div style={{ ...displayFont, fontSize: 24, color: C.goldSoft }}>Lote y Rakes</div>
            <div style={{ color: "rgba(244,234,214,0.55)", fontSize: 12.5, ...monoFont }}>{game.date}</div>
          </div>
          <div style={{ marginTop: 14 }}>
            <Field label="Valor de lote">
              <MoneyInput value={game.loteValue} onChange={setLote} />
            </Field>
          </div>
          <div style={{ marginTop: 16, borderTop: `1px solid ${C.panelLine}`, paddingTop: 14 }}>
            <div style={{ ...displayFont, fontSize: 16, color: C.goldSoft, marginBottom: 8, letterSpacing: "0.05em" }}>RAKE</div>
            <Field label={rakeLocked ? "Rake para anfitrión 🔒" : "Rake para anfitrión"}>
              <MoneyInput value={game.rakeHost || 0} onChange={setRakeHost} disabled={rakeLocked} />
            </Field>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginTop: 10 }}>
              <Field label="Rake para autos — # de autos">
                <input type="number" min="0" style={{ ...inputStyle, opacity: rakeLocked ? 0.55 : 1 }} value={game.rakeAutosCount || 0} disabled={rakeLocked} onChange={(e) => setRakeAutosCount(e.target.value)} onFocus={(e) => e.target.select()} />
              </Field>
              <Field label="Monto por auto">
                <MoneyInput value={game.rakeAutoAmount ?? 250} onChange={setRakeAutoAmount} disabled={rakeLocked} />
              </Field>
            </div>
            <div style={{ marginTop: 10 }}>
              <ScoreBox label="Rake total (anfitrión + autos)" value={money(game.rake)} />
            </div>
            {rakeLocked && <div style={{ fontSize: 11, color: "rgba(244,234,214,0.4)", marginTop: 8 }}>El rake queda fijo una vez que se hizo la entrega de fichas.</div>}
          </div>
        </Panel>
      </div>
    );
  }

  if (effectiveView === "jugadoresPartida") {
    return (
      <div style={{ display: "grid", gap: 16 }}>
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
      </div>
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
          <GhostBtn icon={X} color={C.loss} onClick={cancelPartida}>Cancelar partida</GhostBtn>
        </div>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(2,1fr)", gap: 8, marginTop: 14 }}>
          <ScoreBox label="Monto del lote" value={money(game.loteValue)} />
          <ScoreBox label="Rake + estacionamiento" value={money(game.rake)} />
        </div>
      </Panel>

      <GameStatusBlock game={game} players={players} totals={totals} />
    </div>
  );
}

function PendingRequestsPanel({ game, roster, onResolve }) {
  const pending = (game.requests || []).filter((r) => r.status === "pending");
  if (pending.length === 0) return null;
  return (
    <Panel style={{ border: `1px solid ${C.gold}` }}>
      <SectionTitle icon={AlertCircle}>Solicitudes pendientes ({pending.length})</SectionTitle>
      <div style={{ display: "grid", gap: 8 }}>
        {pending.map((r) => {
          // Ojo: para una solicitud de "unirse a la mesa" el jugador todavía
          // NO está en playerIds, así que hay que buscarlo en el roster
          // completo (no solo entre los ya sentados) para poder mostrar su
          // nombre y avatar.
          const p = roster.find((pl) => pl.id === r.playerId);
          const isJoin = r.type === "join";
          return (
            <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 8, background: isJoin ? "rgba(216,173,63,0.14)" : "rgba(0,0,0,0.18)", borderRadius: 9, padding: "9px 10px", flexWrap: "wrap" }}>
              {p && <Avatar player={p} size={24} />}
              <div style={{ flex: 1, minWidth: 120 }}>
                <div style={{ color: C.card, fontWeight: 700, fontSize: 13.5 }}>{p ? p.name : "?"}</div>
                {isJoin ? (
                  <div style={{ ...monoFont, fontSize: 12, color: C.goldSoft, fontWeight: 700 }}>Pide unirse a la mesa</div>
                ) : (
                  <div style={{ ...monoFont, fontSize: 12, color: r.type === "cash" ? C.cash : C.virtual, fontWeight: 700 }}>
                    {r.type === "cash" ? "Cash" : "Virtual"} · {money(r.amount)}
                  </div>
                )}
              </div>
              <GhostBtn icon={Check} color={C.win} onClick={() => onResolve(r.id, true)}>{isJoin ? "Agregar a la mesa" : "Aceptar"}</GhostBtn>
              <GhostBtn icon={X} color={C.loss} onClick={() => onResolve(r.id, false)}>Rechazar</GhostBtn>
            </div>
          );
        })}
      </div>
    </Panel>
  );
}

function JoinRequestPanel({ myRequests, onSubmit }) {
  const joinRequests = myRequests.filter((r) => r.type === "join");
  const pending = joinRequests.some((r) => r.status === "pending");
  const rejected = [...joinRequests].reverse().find((r) => r.status === "rejected");

  return (
    <Panel>
      <SectionTitle icon={Users}>Unirme a la mesa</SectionTitle>
      <div style={{ fontSize: 12.5, color: "rgba(244,234,214,0.6)", marginBottom: 12 }}>
        Todavía no estás en esta partida. Pídele al host que te agregue para poder comprar lotes.
      </div>
      {pending ? (
        <div style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(216,173,63,0.1)", border: `1px dashed ${C.panelLine}`, borderRadius: 8, padding: "9px 10px" }}>
          <span style={{ fontSize: 12.5, color: "rgba(244,234,214,0.7)" }}>Esperando a que el host te agregue a la mesa…</span>
        </div>
      ) : (
        <>
          {rejected && (
            <div style={{ fontSize: 12, color: C.loss, marginBottom: 8 }}>El host rechazó tu solicitud anterior. Puedes volver a pedirlo.</div>
          )}
          <PrimaryBtn icon={Plus} onClick={onSubmit}>Solicitar unirme a la mesa</PrimaryBtn>
        </>
      )}
    </Panel>
  );
}

function RequestChipsPanel({ loteValue, myRequests, onSubmit }) {
  const [customAmount, setCustomAmount] = useState("");
  const [customType, setCustomType] = useState("cash");
  const pending = myRequests.filter((r) => r.status === "pending");
  const recent = myRequests.filter((r) => r.status !== "pending").slice(-4).reverse();

  const submitCustom = () => {
    onSubmit(customType, customAmount);
    setCustomAmount("");
  };

  return (
    <Panel>
      <SectionTitle icon={Banknote}>Solicitar fichas</SectionTitle>
      <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        <PrimaryBtn icon={Plus} onClick={() => onSubmit("cash", loteValue)} style={{ background: `linear-gradient(180deg, ${C.cash}, ${C.cashDeep})`, color: "#fff" }}>
          1 lote cash
        </PrimaryBtn>
        <PrimaryBtn icon={Plus} onClick={() => onSubmit("virtual", loteValue)} style={{ background: `linear-gradient(180deg, ${C.virtual}, ${C.virtualDeep})`, color: "#fff" }}>
          1 lote virtual
        </PrimaryBtn>
      </div>
      <div style={{ fontSize: 11.5, color: "rgba(244,234,214,0.5)", marginBottom: 6 }}>O pide un monto libre (no tiene que ser un lote completo):</div>
      <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
        <div style={{ display: "flex", background: "rgba(0,0,0,0.24)", borderRadius: 8, padding: 2 }}>
          <button onClick={() => setCustomType("cash")}
            style={{ border: "none", cursor: "pointer", borderRadius: 6, padding: "6px 10px", fontSize: 12, fontWeight: 700, ...bodyFont, background: customType === "cash" ? C.cash : "transparent", color: customType === "cash" ? "#fff" : "rgba(244,234,214,0.6)" }}>
            Cash
          </button>
          <button onClick={() => setCustomType("virtual")}
            style={{ border: "none", cursor: "pointer", borderRadius: 6, padding: "6px 10px", fontSize: 12, fontWeight: 700, ...bodyFont, background: customType === "virtual" ? C.virtual : "transparent", color: customType === "virtual" ? "#fff" : "rgba(244,234,214,0.6)" }}>
            Virtual
          </button>
        </div>
        <input type="number" min="0" placeholder="Monto libre" style={{ ...inputStyle, width: 130 }} value={customAmount} onChange={(e) => setCustomAmount(e.target.value)} onFocus={(e) => e.target.select()} />
        <PrimaryBtn onClick={submitCustom} disabled={!customAmount || Number(customAmount) <= 0}>Solicitar</PrimaryBtn>
      </div>
      {pending.length > 0 && (
        <div style={{ marginTop: 12, display: "grid", gap: 6 }}>
          {pending.map((r) => (
            <div key={r.id} style={{ display: "flex", alignItems: "center", gap: 8, background: "rgba(216,173,63,0.1)", border: `1px dashed ${C.panelLine}`, borderRadius: 8, padding: "7px 10px" }}>
              <span style={{ ...monoFont, fontSize: 12, color: r.type === "cash" ? C.cash : C.virtual, fontWeight: 700 }}>{r.type === "cash" ? "Cash" : "Virtual"} {money(r.amount)}</span>
              <span style={{ fontSize: 11.5, color: "rgba(244,234,214,0.5)", marginLeft: "auto" }}>Esperando aprobación del host…</span>
            </div>
          ))}
        </div>
      )}
      {recent.length > 0 && (
        <div style={{ marginTop: 10, display: "flex", gap: 6, flexWrap: "wrap" }}>
          {recent.map((r) => (
            <span key={r.id} style={{ fontSize: 10.5, ...monoFont, padding: "3px 7px", borderRadius: 6, background: r.status === "approved" ? "rgba(63,191,114,0.14)" : "rgba(226,99,79,0.14)", color: r.status === "approved" ? C.win : C.loss }}>
              {r.type === "cash" ? "Cash" : "Virtual"} {money(r.amount)} · {r.status === "approved" ? "aprobada" : "rechazada"}
            </span>
          ))}
        </div>
      )}
    </Panel>
  );
}

function DinnerReadOnly({ game, players }) {
  const d = game.dinner;
  const totalRecaudado = players.reduce((s, p) => {
    const alcohol = !!d.alcohol?.[p.id];
    return s + (alcohol ? (d.amountAlcohol || 0) : (d.amountNoAlcohol || 0));
  }, 0);
  return (
    <div>
      <div style={{ marginBottom: 12 }}>
        <ScoreBox label="Total recabado" value={money(totalRecaudado)} />
      </div>
      <div style={{ display: "grid", gap: 6 }}>
        {players.map((p) => {
          const alcohol = !!d.alcohol?.[p.id];
          const charge = alcohol ? (d.amountAlcohol || 0) : (d.amountNoAlcohol || 0);
          const paid = !!d.paid?.[p.id];
          return (
            <div key={p.id} style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, background: "rgba(0,0,0,0.16)", borderRadius: 8, padding: "8px 10px" }}>
              <div style={{ display: "flex", alignItems: "center", gap: 7, minWidth: 0 }}>
                <Avatar player={p} size={22} />
                <Wine size={13} color={alcohol ? C.virtual : "rgba(244,234,214,0.3)"} />
                <span style={{ fontSize: 13, color: C.card, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</span>
              </div>
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ ...monoFont, fontSize: 13, color: "rgba(244,234,214,0.75)" }}>{money(charge)}</span>
                <span style={{ fontSize: 10.5, fontWeight: 700, color: paid ? C.win : C.loss, ...monoFont }}>{paid ? "PAGADO" : "PENDIENTE"}</span>
              </div>
            </div>
          );
        })}
      </div>
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
    border: `2px solid ${color}`,
    color: subtract ? color : "#fff",
    borderRadius: 10,
    padding: "14px 20px",
    display: "flex",
    alignItems: "center",
    gap: 8,
    cursor: disabled ? "not-allowed" : "pointer",
    opacity: disabled ? 0.35 : 1,
    fontWeight: 800,
    fontSize: 20,
    ...bodyFont,
  };
}

function PlayerBuyRow({ player, game, onAdd, onRemoveLast, dinnerPaid, onGoToDinner }) {
  const entries = game.purchases.filter((p) => p.playerId === player.id);
  const cashLotes = entries.filter((e) => e.type === "cash").reduce((s, e) => s + e.lotes, 0);
  const virtualLotes = entries.filter((e) => e.type === "virtual").reduce((s, e) => s + e.lotes, 0);
  const cashAmount = entries.filter((e) => e.type === "cash").reduce((s, e) => s + e.amount, 0);
  const virtualAmount = entries.filter((e) => e.type === "virtual").reduce((s, e) => s + e.amount, 0);
  const blocked = !dinnerPaid;

  return (
    <div style={{ background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: 14 }}>
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12 }}>
        <Avatar player={player} size={30} />
        <span style={{ color: C.card, fontWeight: 700, fontSize: 17 }}>{player.name}</span>
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

      <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 12, flexWrap: "wrap" }}>
        <button onClick={() => onAdd(player.id, "cash")} disabled={blocked} style={buyBtnStyle(C.cash, false, blocked)}>
          <Plus size={20} /> Cash
        </button>
        <button onClick={() => onRemoveLast(player.id, "cash")} disabled={cashLotes === 0} style={buyBtnStyle(C.cash, true, cashLotes === 0)}>
          <Minus size={20} /> Cash
        </button>
        <span style={{ marginLeft: "auto", ...monoFont, fontSize: 20, fontWeight: 800, color: C.cash, whiteSpace: "nowrap" }}>
          {cashLotes} lotes · {money(cashAmount)}
        </span>
      </div>

      <div style={{ display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <button onClick={() => onAdd(player.id, "virtual")} disabled={blocked} style={buyBtnStyle(C.virtual, false, blocked)}>
          <Plus size={20} /> Virtual
        </button>
        <button onClick={() => onRemoveLast(player.id, "virtual")} disabled={virtualLotes === 0} style={buyBtnStyle(C.virtual, true, virtualLotes === 0)}>
          <Minus size={20} /> Virtual
        </button>
        <span style={{ marginLeft: "auto", ...monoFont, fontSize: 20, fontWeight: 800, color: C.virtual, whiteSpace: "nowrap" }}>
          {virtualLotes} lotes · {money(virtualAmount)}
        </span>
      </div>
    </div>
  );
}

function DinnerSection({ game, players, update }) {
  const d = game.dinner;
  const setD = (patch) =>
    update((g) => ({
      dinner: { ...g.dinner, ...(typeof patch === "function" ? patch(g.dinner) : patch) },
    }));

  const totalRecaudado = players.reduce((s, p) => {
    const alcohol = !!d.alcohol[p.id];
    return s + (alcohol ? (d.amountAlcohol || 0) : (d.amountNoAlcohol || 0));
  }, 0);

  return (
    <div style={{ marginTop: 6 }}>
      <div style={{ marginBottom: 14 }}>
        <ScoreBox label="Total recabado" value={money(totalRecaudado)} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Cena sin alcohol (por jugador)">
          <MoneyInput value={d.amountNoAlcohol || 0} onChange={(v) => setD({ amountNoAlcohol: v })} />
        </Field>
        <Field label="Cena con alcohol (por jugador)">
          <MoneyInput value={d.amountAlcohol || 0} onChange={(v) => setD({ amountAlcohol: v })} />
        </Field>
      </div>
      <div style={{ marginTop: 14, display: "grid", gap: 6 }}>
        {players.map((p) => {
          const alcohol = !!d.alcohol[p.id];
          const charge = alcohol ? (d.amountAlcohol || 0) : (d.amountNoAlcohol || 0);
          const paid = !!d.paid?.[p.id];
          return (
            <div key={p.id} style={{ background: "rgba(0,0,0,0.16)", borderRadius: 8, padding: "8px 10px" }}>
              <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                <div style={{ display: "flex", alignItems: "center", gap: 7, flex: 1, minWidth: 0 }}>
                  <Avatar player={p} size={22} />
                  <span style={{ fontSize: 13, color: C.card, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{p.name}</span>
                </div>
                <div style={{ display: "flex", background: "rgba(0,0,0,0.24)", borderRadius: 7, padding: 2 }}>
                  <button onClick={() => setD((dd) => ({ alcohol: { ...dd.alcohol, [p.id]: false } }))}
                    style={{ border: "none", cursor: "pointer", borderRadius: 5, padding: "4px 8px", fontSize: 11, fontWeight: 700, ...bodyFont, background: !alcohol ? C.gold : "transparent", color: !alcohol ? C.ink : "rgba(244,234,214,0.6)" }}>
                    Sin alcohol
                  </button>
                  <button onClick={() => setD((dd) => ({ alcohol: { ...dd.alcohol, [p.id]: true } }))}
                    style={{ border: "none", cursor: "pointer", borderRadius: 5, padding: "4px 8px", fontSize: 11, fontWeight: 700, display: "flex", alignItems: "center", gap: 3, ...bodyFont, background: alcohol ? C.virtual : "transparent", color: alcohol ? "#fff" : "rgba(244,234,214,0.6)" }}>
                    <Wine size={11} /> Con alcohol
                  </button>
                </div>
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
// Color fijo para "Virtual pendiente" (bloque 2) y las columnas relacionadas
// con virtuales — no es ninguno de los tonos ya usados (cash/virtual/win/loss).
const ORANGE = "#f2883c";

function FinalizeGame({ game, roster, onBack, onConfirm, update }) {
  // Orden alfabético, igual que el resto de las pantallas de la partida.
  const players = useMemo(
    () => game.playerIds.map((id) => roster.find((r) => r.id === id)).filter(Boolean).sort((a, b) => a.name.localeCompare(b.name)),
    [game.playerIds, roster]
  );

  // Todo lo que se va tecleando en esta pantalla (paga virtual, fichas
  // remanentes, ajuste manual) se guarda de una vez en la propia partida
  // (game.finalizeDraft), no solo al tocar "Volver". Así, sin importar por
  // dónde se salga de esta pantalla, lo ya capturado nunca se pierde al
  // volver a entrar a cerrar la partida.
  const draft = game.finalizeDraft || {};
  const draftFor = (pid, field) => {
    const v = draft[pid] ? draft[pid][field] : undefined;
    return v !== undefined && v !== null ? String(v) : "";
  };

  const [pagaVirtual, setPagaVirtualState] = useState(() =>
    Object.fromEntries(players.map((p) => [p.id, draftFor(p.id, "pagaVirtual")]))
  );
  const [remanente, setRemanenteState] = useState(() =>
    Object.fromEntries(players.map((p) => [p.id, draftFor(p.id, "remanente")]))
  );
  // Ajuste manual por jugador (+/-): corrige el monto de fichas con el que se
  // presenta a cobrar (un billete mal contado, una ficha que apareció
  // después, etc.) sin tener que tocar los campos principales.
  const [adjustments, setAdjustmentsState] = useState(() =>
    Object.fromEntries(players.map((p) => [p.id, draftFor(p.id, "adjust")]))
  );
  // Qué jugador tiene un campo enfocado ahora mismo — el aviso de "pendiente
  // por cuadrar" se muestra pegado a esa fila, para que se vea sin que el
  // teclado del celular lo tape (a diferencia de una barra fija abajo de la
  // pantalla, que el teclado sí bloquea).
  const [focusedId, setFocusedId] = useState(null);
  const moneySigned = (n) => (n < 0 ? "-" : n > 0 ? "+" : "") + "$" + Math.abs(Math.round(n)).toLocaleString("en-US");

  const persistField = (playerId, field, value) => {
    update((g) => ({
      finalizeDraft: {
        ...(g.finalizeDraft || {}),
        [playerId]: { ...((g.finalizeDraft || {})[playerId] || {}), [field]: value },
      },
    }));
  };
  const setPagaVirtual = (pid, v) => { setPagaVirtualState((c) => ({ ...c, [pid]: v })); persistField(pid, "pagaVirtual", v); };
  const setRemanente = (pid, v) => { setRemanenteState((c) => ({ ...c, [pid]: v })); persistField(pid, "remanente", v); };
  const setAdjustment = (pid, v) => { setAdjustmentsState((c) => ({ ...c, [pid]: v })); persistField(pid, "adjust", v); };

  // "Guardar en Excel ahora": empuja un snapshot on-demand a la hoja
  // "EntregaFichas" (aparte del guardado automático de "active", que tiene
  // su propia latencia y candado de 3s). Sirve para tener la certeza de que
  // lo capturado hasta este momento ya quedó en el Excel, y deja un
  // historial con timestamp para auditorías posteriores.
  const [saveState, setSaveState] = useState("idle"); // idle | saving | ok | error
  const [saveErrorMsg, setSaveErrorMsg] = useState("");

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
  const cashDisponible = round1(totalCash - rake);
  const targetTotal = totalCash + totalVirtual;

  // "Paga virtual" que ya se fue capturando por jugador — es lo que reduce
  // el virtual pendiente a nivel de toda la partida (bloque 2).
  const pagaVirtualTotal = players.reduce((s, p) => s + (Number(pagaVirtual[p.id]) || 0), 0);
  const virtualPendiente = round1(totalVirtual - pagaVirtualTotal);

  // El dinero total en juego (para el "cuadre") se basa en las fichas tal
  // cual se entregaron — paga virtual + fichas remanentes + ajuste manual.
  // El ajuste manual SÍ altera las fichas disponibles del jugador (positiva
  // o negativamente) y por lo tanto el total general: si corrige un mal
  // conteo, ese dinero corregido tiene que reflejarse en el cuadre.
  const enteredCount = players.filter((p) => pagaVirtual[p.id] !== "" || remanente[p.id] !== "").length;
  const totalFinalValue = players.reduce(
    (s, p) => s + (Number(pagaVirtual[p.id]) || 0) + (Number(remanente[p.id]) || 0) + (Number(adjustments[p.id]) || 0),
    0
  ) + rake;
  const runningDiff = round1(totalFinalValue - targetTotal);

  const ready = players.every((p) => (pagaVirtual[p.id] === "" || !isNaN(Number(pagaVirtual[p.id]))) && (remanente[p.id] === "" || !isNaN(Number(remanente[p.id]))));
  // Regla de la app: no se puede cerrar la partida si lo entregado (fichas +
  // rake) no cuadra exactamente contra el total comprado (cash + virtual).
  const canConfirm = ready && enteredCount > 0 && runningDiff === 0;

  const cuadra = runningDiff === 0;

  // Filas concentradas por jugador para los bloques 4 (VIRTUALES) y 5 (CASH):
  // - Total A (por jugador) = lotes virtuales que le quedan sin pagar.
  // - Total B (por jugador) = fichas remanentes + ajuste manual, es decir,
  //   lo que le queda a ese jugador una vez separado lo que paga de virtual
  //   (que ya se contabilizó en Total A). Sumando todos los Total A y Total B
  //   se puede volver a armar la misma ecuación de cuadre que ya usa la
  //   pantalla (Cash remanente + Total A = Total B), pero mirándola desde el
  //   ángulo de cash vs. virtual en vez del total genérico.
  const rows = useMemo(() => players.map((p) => {
    const bi = buyIns[p.id];
    const pv = Number(pagaVirtual[p.id]) || 0;
    const rem = Number(remanente[p.id]) || 0;
    const adj = Number(adjustments[p.id]) || 0;
    const fichasTotales = round1(pv + rem + adj);
    const totalA = round1(bi.virtual - pv);
    const fichasEntregadas = round1(rem);
    const fichasRemanentesPostVirtual = round1(adj);
    const totalB = round1(fichasEntregadas + fichasRemanentesPostVirtual);
    return { p, bi, pv, rem, adj, fichasTotales, totalA, fichasEntregadas, fichasRemanentesPostVirtual, totalB };
  }), [players, buyIns, pagaVirtual, remanente, adjustments]);

  const grandTotalA = round1(rows.reduce((s, r) => s + r.totalA, 0));
  const grandTotalB = round1(rows.reduce((s, r) => s + r.totalB, 0));
  const totalesCuadran = round1(cashDisponible + grandTotalA - grandTotalB) === 0;

  const saveExcelNow = async () => {
    setSaveState("saving");
    setSaveErrorMsg("");
    try {
      const payload = {
        gameId: game.id,
        gameDate: game.date,
        rows: rows.map((r) => ({
          playerId: r.p.id,
          playerName: r.p.name,
          debeVirtual: r.bi.virtual,
          pagaVirtual: r.pv,
          fichasRemanentes: r.rem,
          ajusteManual: r.adj,
          fichasTotales: r.fichasTotales,
          rake,
          cashDisponible,
          virtualPendiente,
          totalA: r.totalA,
          totalB: r.totalB,
        })),
      };
      const res = await fetch("/api/store?key=entregaFichas", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      if (!res.ok) {
        const t = await res.text();
        throw new Error(t || `HTTP ${res.status}`);
      }
      setSaveState("ok");
    } catch (e) {
      console.error("guardar en Excel (EntregaFichas) falló", e);
      setSaveState("error");
      setSaveErrorMsg(String((e && e.message) || e));
    }
  };

  const rowLabelStyle = { fontSize: 10, color: "rgba(244,234,214,0.45)", textTransform: "uppercase", letterSpacing: "0.04em" };
  const blockLabelStyle = { fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 4 };
  const blockValueStyle = (color, big) => ({ ...monoFont, fontWeight: 800, fontSize: big ? 30 : 18, color });
  const colLabelStyle = { fontSize: 9.5, color: "rgba(244,234,214,0.4)", textTransform: "uppercase", letterSpacing: "0.04em" };
  const colValueStyle = (color) => ({ ...monoFont, fontWeight: 800, fontSize: 16, color });
  const thStyle = { textAlign: "right", fontSize: 10, textTransform: "uppercase", letterSpacing: "0.03em", color: "rgba(244,234,214,0.42)", fontWeight: 600, padding: "0 8px 8px", whiteSpace: "nowrap" };
  const tdStyle = { textAlign: "right", padding: "7px 8px", ...monoFont, fontSize: 12.5, color: "rgba(244,234,214,0.85)", borderTop: "1px solid rgba(255,255,255,0.06)", whiteSpace: "nowrap" };
  const tdNameStyle = { ...tdStyle, textAlign: "left", ...bodyFont, fontWeight: 600, color: C.card };

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <Panel>
        <SectionTitle
          icon={Trophy}
          right={
            <div style={{ display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" }}>
              {saveState === "ok" && <span style={{ fontSize: 11, color: C.win, ...monoFont }}>Guardado ✓</span>}
              {saveState === "error" && <span style={{ fontSize: 11, color: C.loss, ...monoFont }} title={saveErrorMsg}>Error al guardar</span>}
              <GhostBtn onClick={saveExcelNow} icon={Save} color={C.goldSoft}>
                {saveState === "saving" ? "Guardando…" : "Guardar en Excel"}
              </GhostBtn>
            </div>
          }
        >
          Entrega de fichas
        </SectionTitle>
        <div style={{ color: "rgba(244,234,214,0.6)", fontSize: 12.5, marginBottom: 10 }}>
          De las fichas que entrega cada jugador, primero se pagan los lotes virtuales pendientes; lo que sobra son sus fichas remanentes (a cobrar en cash o transferencia). El botón "Guardar en Excel" empuja este avance al Excel al instante, sin esperar a la sincronización automática.
        </div>

        {/* Bloques 1 y 2: Cash (arriba) y Virtual (justo debajo, alineados por columna) */}
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1.3fr", gap: 8, marginBottom: 14 }}>
          <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>Cash</div>
            <div style={blockValueStyle(C.card)}>{money(totalCash)}</div>
          </div>
          <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>menos Rake+Estacionamiento</div>
            <div style={blockValueStyle(C.card)}>-{money(rake)}</div>
          </div>
          <div style={{ background: "rgba(47,174,102,0.14)", border: `1px solid ${C.cashDeep}`, borderRadius: 9, padding: "10px 8px", textAlign: "center", display: "flex", flexDirection: "column", justifyContent: "center" }}>
            <div style={blockLabelStyle}>Cash disponible</div>
            <div style={blockValueStyle(C.cash, true)}>{money(cashDisponible)}</div>
          </div>

          <div style={{ gridColumn: "1 / 3", background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>Virtual</div>
            <div style={blockValueStyle(C.card)}>{money(totalVirtual)}</div>
          </div>
          <div style={{ background: "rgba(0,0,0,0.22)", border: `1px solid ${C.panelLine}`, borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>Virtual pendiente</div>
            <div style={blockValueStyle(ORANGE)}>{money(virtualPendiente)}</div>
          </div>
        </div>

        <div style={{ display: "grid", gap: 10 }}>
          {players.map((p) => {
            const bi = buyIns[p.id];
            const focused = focusedId === p.id;
            const pv = Number(pagaVirtual[p.id]) || 0;
            const rem = Number(remanente[p.id]) || 0;
            const adj = Number(adjustments[p.id]) || 0;
            // El ajuste manual corrige las fichas realmente entregadas (paga
            // virtual + remanentes), sin tocar esos campos principales, y
            // esas fichas ajustadas son las que saldan primero el buy-in
            // virtual (igual que antes).
            const fichasAjustadas = round1(pv + rem + adj);
            const virtualPagado = Math.max(0, Math.min(fichasAjustadas, bi.virtual));
            const cashOutPendiente = round1(fichasAjustadas - bi.virtual);
            const debeVirtual = cashOutPendiente < 0;
            const cashAjustado = round1(bi.cash + adj);
            return (
              <div key={p.id} style={{ background: "rgba(0,0,0,0.18)", borderRadius: 10, padding: "12px 14px" }}>
                <div style={{ display: "flex", alignItems: "center", justifyContent: "center", gap: 8, marginBottom: 10 }}>
                  <Avatar player={p} size={22} />
                  <span style={{ color: C.card, fontWeight: 700, fontSize: 14.5 }}>{p.name}</span>
                </div>

                <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 6, textAlign: "center" }}>
                  <div>
                    <div style={colLabelStyle}>Debe virtual</div>
                    <div style={colValueStyle(C.virtual)}>{money(bi.virtual)}</div>
                  </div>
                  <div>
                    <div style={colLabelStyle}>Paga virtual</div>
                    <input
                      type="number" min="0" placeholder="0"
                      style={{
                        width: "100%", background: "transparent", border: "none",
                        borderBottom: `1px solid ${C.panelLine}`, textAlign: "center",
                        ...monoFont, fontWeight: 800, fontSize: 16, color: C.card, padding: "0 0 2px",
                      }}
                      value={pagaVirtual[p.id]}
                      onChange={(e) => setPagaVirtual(p.id, e.target.value)}
                      onFocus={(e) => { e.target.select(); setFocusedId(p.id); }}
                      onBlur={() => setFocusedId((cur) => (cur === p.id ? null : cur))}
                    />
                  </div>
                  <div>
                    <div style={colLabelStyle}>Fichas remanentes</div>
                    <input
                      type="number" min="0" placeholder="0"
                      style={{
                        width: "100%", background: "transparent", border: "none",
                        borderBottom: `1px solid ${C.panelLine}`, textAlign: "center",
                        ...monoFont, fontWeight: 800, fontSize: 16, color: C.card, padding: "0 0 2px",
                      }}
                      value={remanente[p.id]}
                      onChange={(e) => setRemanente(p.id, e.target.value)}
                      onFocus={(e) => { e.target.select(); setFocusedId(p.id); }}
                      onBlur={(e) => {
                        if (e.target.value !== "") setRemanente(p.id, String(roundTo100(e.target.value)));
                        setFocusedId((cur) => (cur === p.id ? null : cur));
                      }}
                      step="100"
                    />
                  </div>
                  <div>
                    <div style={colLabelStyle}>Fichas totales</div>
                    <div style={colValueStyle(C.goldSoft)}>{money(fichasAjustadas)}</div>
                  </div>
                </div>

                <div style={{ display: "grid", gap: 8, borderTop: `1px solid ${C.panelLine}`, marginTop: 10, paddingTop: 10 }}>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                    <span style={rowLabelStyle}>Cash</span>
                    <span style={{ ...monoFont, fontSize: 16, fontWeight: 800, color: C.cash }}>
                      {money(bi.cash)}
                      {adj !== 0 && (
                        <>
                          <span style={{ color: "rgba(244,234,214,0.35)" }}> · </span>
                          <span style={{ fontSize: 10, color: "rgba(244,234,214,0.5)", textTransform: "uppercase", fontWeight: 700 }}>Cash ajustado </span>
                          {money(cashAjustado)}
                        </>
                      )}
                    </span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                    <span style={rowLabelStyle}>Virtual pagado</span>
                    <span style={{ ...monoFont, fontSize: 16, fontWeight: 800, color: C.win }}>{money(virtualPagado)}</span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                    <span style={rowLabelStyle}>{debeVirtual ? "Debe virtual" : "Cash out pendiente"}</span>
                    <span style={{ ...monoFont, fontSize: 22, fontWeight: 800, color: debeVirtual ? C.loss : C.win }}>
                      {money(Math.abs(cashOutPendiente))}
                    </span>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, flexWrap: "wrap" }}>
                    <span style={rowLabelStyle}>Ajuste manual (+/-)</span>
                    <input
                      type="number" placeholder="0" style={{ ...inputStyle, width: 100, textAlign: "right" }}
                      value={adjustments[p.id]} onChange={(e) => setAdjustment(p.id, e.target.value)}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                </div>

                {focused && (
                  <div style={{
                    marginTop: 8, display: "inline-flex", alignItems: "center", gap: 6,
                    background: cuadra ? "rgba(63,191,114,0.14)" : "rgba(226,99,79,0.14)",
                    border: `1px solid ${cuadra ? C.win : C.loss}`, borderRadius: 8, padding: "5px 10px",
                  }}>
                    {cuadra ? (
                      <span style={{ ...bodyFont, fontWeight: 700, fontSize: 13, color: C.win }}>Cuadrado</span>
                    ) : (
                      <>
                        <span style={{ ...bodyFont, fontSize: 11.5, color: "rgba(244,234,214,0.7)" }}>Pendiente por cuadrar:</span>
                        <span style={{ ...monoFont, fontWeight: 700, fontSize: 13, color: C.loss }}>{moneySigned(runningDiff)}</span>
                      </>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
        {ready && enteredCount > 0 && runningDiff !== 0 && (
          <div style={{ display: "flex", gap: 7, alignItems: "flex-start", marginTop: 12, background: "rgba(226,99,79,0.12)", border: `1px solid ${C.loss}`, borderRadius: 8, padding: "8px 10px" }}>
            <AlertCircle size={15} color={C.loss} style={{ flexShrink: 0, marginTop: 1 }} />
            <div style={{ fontSize: 12, color: "rgba(244,234,214,0.85)" }}>
              <strong>No se puede cerrar la partida hasta que el total cuadre exactamente</strong> — ajustá los montos de cierre. Total entregado: {money(totalFinalValue)} · Total general: {money(targetTotal)}
            </div>
          </div>
        )}
      </Panel>

      {/* Bloque 4: VIRTUALES — tabla concentrada, no editable, para cuadrar */}
      <Panel>
        <SectionTitle icon={Coins}>Virtuales</SectionTitle>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={{ ...thStyle, textAlign: "left" }}>Jugador</th>
              <th style={thStyle}>Lotes virtuales</th>
              <th style={thStyle}>Lotes virtuales pagados</th>
              <th style={thStyle}>Total A</th>
            </tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.p.id}>
                  <td style={tdNameStyle}>{r.p.name}</td>
                  <td style={{ ...tdStyle, color: C.virtual }}>{money(r.bi.virtual)}</td>
                  <td style={tdStyle}>{money(r.pv)}</td>
                  <td style={{ ...tdStyle, color: r.totalA === 0 ? C.win : ORANGE, fontWeight: 700 }}>{money(r.totalA)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ ...tdNameStyle, fontWeight: 800 }}>Total A</td>
                <td style={tdStyle} />
                <td style={tdStyle} />
                <td style={{ ...tdStyle, fontWeight: 800, color: grandTotalA === 0 ? C.win : ORANGE }}>{money(grandTotalA)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Panel>

      {/* Bloque 5: CASH — tabla concentrada, no editable, para cuadrar */}
      <Panel>
        <SectionTitle icon={Banknote}>Cash</SectionTitle>
        <div style={{ overflowX: "auto" }}>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr>
              <th style={{ ...thStyle, textAlign: "left" }}>Jugador</th>
              <th style={thStyle}>Fichas entregadas</th>
              <th style={thStyle}>Fichas remanentes (post-virtual)</th>
              <th style={thStyle}>Total B</th>
            </tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.p.id}>
                  <td style={tdNameStyle}>{r.p.name}</td>
                  <td style={{ ...tdStyle, color: C.cash }}>{money(r.fichasEntregadas)}</td>
                  <td style={tdStyle}>{money(r.fichasRemanentesPostVirtual)}</td>
                  <td style={{ ...tdStyle, fontWeight: 700, color: C.goldSoft }}>{money(r.totalB)}</td>
                </tr>
              ))}
              <tr>
                <td style={{ ...tdNameStyle, fontWeight: 800 }}>Total B</td>
                <td style={tdStyle} />
                <td style={tdStyle} />
                <td style={{ ...tdStyle, fontWeight: 800, color: C.goldSoft }}>{money(grandTotalB)}</td>
              </tr>
            </tbody>
          </table>
        </div>
      </Panel>

      {/* Bloque 6: TOTALES — Cash remanente + Total A debe ser igual a Total B */}
      <Panel>
        <SectionTitle icon={CircleDollarSign}>Totales</SectionTitle>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 8 }}>
          <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>Cash remanente</div>
            <div style={blockValueStyle(C.cash)}>{money(cashDisponible)}</div>
          </div>
          <div style={{ background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center" }}>
            <div style={blockLabelStyle}>Total A</div>
            <div style={blockValueStyle(ORANGE)}>{money(grandTotalA)}</div>
          </div>
          <div style={{
            background: "rgba(0,0,0,0.22)", borderRadius: 9, padding: "9px 8px", textAlign: "center",
            border: `1px solid ${totalesCuadran ? C.win : C.loss}`,
          }}>
            <div style={blockLabelStyle}>Total B</div>
            <div style={blockValueStyle(totalesCuadran ? C.win : C.loss)}>{money(grandTotalB)}</div>
          </div>
        </div>
        <div style={{ marginTop: 10, fontSize: 12, textAlign: "center", color: totalesCuadran ? C.win : "rgba(244,234,214,0.7)" }}>
          {totalesCuadran
            ? "Cash remanente + Total A = Total B — cuadrado ✓"
            : `Cash remanente + Total A (${money(round1(cashDisponible + grandTotalA))}) debería ser igual a Total B (${money(grandTotalB)}) — todavía no cuadra.`}
        </div>
      </Panel>

      <div style={{ display: "flex", gap: 10 }}>
        <GhostBtn onClick={onBack}>Volver</GhostBtn>
        <PrimaryBtn
          disabled={!canConfirm}
          onClick={() => onConfirm(
            Object.fromEntries(players.map((p) => [p.id, (Number(pagaVirtual[p.id]) || 0) + (Number(remanente[p.id]) || 0)])),
            Object.fromEntries(players.map((p) => [p.id, Number(adjustments[p.id]) || 0]).filter(([, v]) => v !== 0))
          )}
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
function FinalizedGame({ game, roster, onClose, setActiveGame, setGames, adminPassword }) {
  const r = useMemo(() => computeSettlement(game, roster), [game, roster]);
  const players = game.playerIds.map((id) => roster.find((p) => p.id === id)).filter(Boolean);
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
  // Se pide la contraseña de administrador porque la partida ya se finalizó
  // "formalmente" — reabrirla es una acción sensible.
  const handleBack = async () => {
    if (!(await requestAdminPassword(adminPassword, "reabrir esta partida"))) return;
    setGames((gs) => gs.filter((g) => g.id !== game.id));
    setActiveGame({ ...game, finished: false });
  };

  return (
    <div style={{ display: "grid", gap: 16 }}>
      <GhostBtn onClick={handleBack} icon={ChevronLeft} color={C.goldSoft}>Reabrir partida</GhostBtn>

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
              <span style={{ ...monoFont, fontSize: 12.5, color: C.cash }}>{money(p.pagoCash)}</span>
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
            const charge = alcohol ? (d.amountAlcohol || 0) : (d.amountNoAlcohol || 0);
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

function HistoryTab({ games, roster, setGames, adminPassword, activeGame, setActiveGame }) {
  const [openId, setOpenId] = useState(null);
  const [syncing, setSyncing] = useState(false);
  const sorted = [...games].sort((a, b) => (a.date < b.date ? 1 : -1));

  const handleDelete = async (g) => {
    if (!(await requestAdminPassword(adminPassword, "eliminar partidas"))) return;
    if (!confirm(`¿Eliminar definitivamente la partida del ${g.date}? Esta acción no se puede deshacer.`)) return;
    setGames((gs) => gs.filter((x) => x.id !== g.id));
  };

  // Reabre una partida ya cerrada (y ya reemplazada por otra en curso, o
  // sin ninguna partida activa) para poder ajustar algo pendiente. Pide
  // contraseña de administrador porque es una acción sensible: se pierde
  // temporalmente el resultado ya calculado hasta volver a cerrarla.
  const handleReopen = async (g) => {
    if (activeGame && !activeGame.finished) {
      alert("Ya hay una partida en curso. Finalízala o cancélala antes de reabrir otra.");
      return;
    }
    if (!(await requestAdminPassword(adminPassword, "reabrir esta partida"))) return;
    setGames((gs) => gs.filter((x) => x.id !== g.id));
    setActiveGame({ ...g, finished: false });
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
                          <span style={{ ...monoFont, fontSize: 12.5, color: C.cash }}>{money(p.pagoCash)}</span>
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

                  <div style={{ display: "flex", justifyContent: "center", marginTop: 4 }}>
                    <GhostBtn onClick={() => handleReopen(g)} icon={ChevronLeft} color={C.goldSoft}>Reabrir partida</GhostBtn>
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
