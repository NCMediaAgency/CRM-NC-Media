import React, { Suspense, lazy, useEffect, useState } from "react";
import { supabase } from "./lib/supabase";
import { archivioCloud, archivioLocale } from "./lib/storage";
/* Il CRM è il pezzo pesante: si scarica solo dopo l'accesso, così il login appare subito */
const CRM = lazy(() => import("./CRM.jsx"));

/* In sviluppo, con VITE_LOCALE=1, si salta il login e si salva nel browser */
const LOCALE = import.meta.env.DEV && import.meta.env.VITE_LOCALE === "1";

const STILE = `
body{margin:0;background:#0B0B0E}
.accesso{min-height:100vh;display:flex;align-items:center;justify-content:center;padding:16px;background:#0B0B0E;
  color:#F2F2F5;font-family:"Inter",-apple-system,"Segoe UI",Roboto,system-ui,sans-serif;font-size:14px}
.accesso form,.accesso .box{width:100%;max-width:360px;background:#131317;border:1px solid #26262E;border-radius:14px;padding:28px;
  display:flex;flex-direction:column;gap:14px}
.accesso h1{font-size:20px;margin:0}
.accesso p{margin:0;color:#A0A0AC;line-height:1.5}
.accesso label{display:flex;flex-direction:column;gap:6px;color:#A0A0AC;font-size:12px}
.accesso input{background:#0B0B0E;border:1px solid #26262E;border-radius:8px;padding:11px 12px;color:#F2F2F5;font:inherit}
.accesso input:focus{outline:2px solid #8E2039;outline-offset:-1px}
.accesso button{background:#8E2039;color:#fff;border:0;border-radius:8px;padding:12px;font:inherit;font-weight:600;cursor:pointer}
.accesso button:disabled{opacity:.6;cursor:default}
.accesso .err{color:#D8455F}
.esci{position:fixed;right:12px;bottom:12px;z-index:50;background:#191920;color:#A0A0AC;border:1px solid #26262E;
  border-radius:8px;padding:6px 10px;font:12px "Inter",system-ui,sans-serif;cursor:pointer;opacity:.75}
.esci:hover{opacity:1;color:#F2F2F5}
`;

function Schermo({ children }) {
  return <div className="accesso"><style>{STILE}</style>{children}</div>;
}

function Accesso() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [attesa, setAttesa] = useState(false);
  const [errore, setErrore] = useState("");

  const entra = async (e) => {
    e.preventDefault();
    setAttesa(true); setErrore("");
    const { error } = await supabase.auth.signInWithPassword({ email: email.trim(), password });
    setAttesa(false);
    if (error) setErrore(error.message === "Invalid login credentials" ? "Email o password non corrette." : error.message);
  };

  return (
    <Schermo>
      <form onSubmit={entra}>
        <h1>Accedi al CRM</h1>
        <label>Email<input type="email" autoComplete="email" required value={email} onChange={(e) => setEmail(e.target.value)} /></label>
        <label>Password<input type="password" autoComplete="current-password" required value={password} onChange={(e) => setPassword(e.target.value)} /></label>
        {errore && <p className="err">{errore}</p>}
        <button type="submit" disabled={attesa}>{attesa ? "Accesso in corso…" : "Accedi"}</button>
      </form>
    </Schermo>
  );
}

export default function App() {
  const [sessione, setSessione] = useState(undefined);
  const [pronto, setPronto] = useState(LOCALE);
  const [errore, setErrore] = useState("");

  useEffect(() => {
    if (LOCALE) { window.storage = archivioLocale(); return; }
    supabase.auth.getSession().then(({ data }) => setSessione(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_evento, s) => setSessione(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  /* Dopo l'accesso: trovo (o creo al primo accesso) lo spazio di lavoro e vi aggancio l'archivio */
  const utente = sessione && sessione.user ? sessione.user.id : "";
  useEffect(() => {
    if (LOCALE) return;
    setPronto(false); setErrore("");
    if (!utente) return;
    let vivo = true;
    supabase.rpc("my_workspace").then(({ data, error }) => {
      if (!vivo) return;
      if (error || !data) { setErrore("Non riesco a aprire il tuo spazio di lavoro. Ricarica la pagina; se il problema resta, controlla la configurazione del database."); return; }
      window.storage = archivioCloud(data);
      setPronto(true);
    });
    return () => { vivo = false; };
  }, [utente]);

  const caricamento = <Schermo><div className="box"><p>Carico il CRM…</p></div></Schermo>;

  if (LOCALE) return <Suspense fallback={caricamento}><CRM /></Suspense>;
  if (sessione === undefined) return <Schermo><div className="box"><p>Caricamento…</p></div></Schermo>;
  if (!sessione) return <Accesso />;
  if (errore) return <Schermo><div className="box"><p className="err">{errore}</p><button onClick={() => supabase.auth.signOut()}>Esci</button></div></Schermo>;
  if (!pronto) return <Schermo><div className="box"><p>Apro il tuo spazio di lavoro…</p></div></Schermo>;

  return (
    <>
      <style>{STILE}</style>
      {/* key: cambiando utente il CRM riparte da zero e rilegge il suo archivio */}
      <Suspense fallback={caricamento}><CRM key={utente} /></Suspense>
      <button className="esci" onClick={() => supabase.auth.signOut()} title={sessione.user.email}>Esci</button>
    </>
  );
}
