"use client";
import React, { useState, useEffect } from "react";
import { supabase } from "../lib/supabase.js";
import LaborApp, { setAuthInfo, MASTER_EMAIL } from "./LaborApp.jsx";

// 로그인한 이메일의 접속 권한(hq/cell/store)을 조회. 마스터 계정은 테이블 상태와 무관하게 항상 hq로 취급.
// approvers 테이블에 code 컬럼이 아직 없어도(마이그레이션 전) 로그인 자체는 깨지지 않도록 폴백 조회함.
async function resolveAuth(session) {
  const email = session?.user?.email?.toLowerCase() || "";
  if (!email) return { email: "", role: null, cell: null, code: null };
  if (email === MASTER_EMAIL.toLowerCase()) return { email, role: "hq", cell: null, code: null };
  try {
    let { data, error } = await supabase.from("approvers").select("role,cell,code").eq("email", email).maybeSingle();
    if (error) ({ data, error } = await supabase.from("approvers").select("role,cell").eq("email", email).maybeSingle());
    if (error) throw error;
    return data ? { email, role: data.role, cell: data.cell, code: data.code || null } : { email, role: null, cell: null, code: null };
  } catch (e) {
    console.warn("권한 조회 실패", e);
    return { email, role: null, cell: null, code: null };
  }
}

// 로그인 + 접속 권한(approvers 등록 여부) 확인까지 담당 — 미등록 계정은 denied 화면
export default function Gate() {
  const [phase, setPhase] = useState("checking"); // checking | login | denied | ready
  const [mode, setMode] = useState("login"); // login | signup
  const [email, setEmail] = useState("");
  const [pw, setPw] = useState("");
  const [pw2, setPw2] = useState("");
  const [err, setErr] = useState("");
  const [notice, setNotice] = useState(""); // 가입 완료 등 안내(에러 아님)
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    let alive = true;
    const safety = setTimeout(() => { if (alive) setPhase((p) => (p === "checking" ? "login" : p)); }, 8000);
    const settle = async (session) => {
      if (!session) { if (alive) { setAuthInfo(null); setPhase("login"); } return; }
      const info = await resolveAuth(session);
      if (!alive) return;
      setAuthInfo(info);
      setPhase(info.role ? "ready" : "denied");
    };
    supabase.auth.getSession().then(({ data }) => settle(data.session)).catch(() => { if (alive) setPhase("login"); });
    const { data: sub } = supabase.auth.onAuthStateChange((_e, sess) => { settle(sess); });
    return () => { alive = false; clearTimeout(safety); sub.subscription.unsubscribe(); };
  }, []);

  async function doLogin(e) {
    e.preventDefault(); setErr(""); setNotice(""); setBusy(true);
    try {
      const { error } = await supabase.auth.signInWithPassword({ email: email.trim().toLowerCase(), password: pw });
      if (error) setErr("이메일 또는 비밀번호를 확인해 주세요.");
    } catch { setErr("로그인 처리 중 문제가 발생했습니다. 다시 시도해 주세요."); }
    finally { setBusy(false); }
  }

  // 자가 가입 — 로그인 계정만 만듦(승인 권한은 별도로 본사 관리자가 "권한 관리"에서 부여해야 접속 가능)
  async function doSignup(e) {
    e.preventDefault(); setErr(""); setNotice(""); setBusy(true);
    try {
      const em = email.trim().toLowerCase();
      if (!em || !em.includes("@")) { setErr("이메일을 정확히 입력해 주세요."); return; }
      if (pw.length < 6) { setErr("비밀번호는 6자 이상이어야 합니다."); return; }
      if (pw !== pw2) { setErr("비밀번호 확인이 서로 다릅니다."); return; }
      const { data, error } = await supabase.auth.signUp({ email: em, password: pw });
      if (error) {
        setErr(/already registered|already exists/i.test(error.message) ? "이미 가입된 이메일입니다. 로그인해 주세요." : "가입 처리 중 문제가 발생했습니다: " + error.message);
        return;
      }
      if (data.session) {
        // 이메일 인증이 꺼져 있으면 가입 즉시 로그인됨 → onAuthStateChange가 자동으로 접속권한 확인(denied 화면)으로 넘김
      } else {
        // 이메일 인증이 켜져 있으면 확인 메일을 눌러야 로그인 가능
        setMode("login"); setPw(""); setPw2("");
        setNotice("가입 신청이 완료됐습니다. 이메일로 온 인증 링크를 확인한 뒤 로그인해 주세요.");
      }
    } catch { setErr("가입 처리 중 문제가 발생했습니다. 다시 시도해 주세요."); }
    finally { setBusy(false); }
  }
  async function doLogout() { try { await supabase.auth.signOut(); } catch {} }

  if (phase === "checking") return <div style={S.center}>불러오는 중…</div>;

  if (phase === "login") return (
    <div style={S.wrap}>
      <form onSubmit={mode === "signup" ? doSignup : doLogin} style={S.card}>
        <div style={S.brand}>애슐리 인건비 모니터</div>
        <div style={S.sub}>{mode === "signup" ? "이메일로 계정을 만드세요" : "이메일로 로그인하세요"}</div>
        <input style={S.input} type="email" placeholder="이메일" value={email} onChange={(e) => setEmail(e.target.value)} autoFocus />
        <input style={S.input} type="password" placeholder="비밀번호" value={pw} onChange={(e) => setPw(e.target.value)} />
        {mode === "signup" && <input style={S.input} type="password" placeholder="비밀번호 확인" value={pw2} onChange={(e) => setPw2(e.target.value)} />}
        {err && <div style={S.err}>{err}</div>}
        {notice && <div style={S.notice}>{notice}</div>}
        <button style={S.btn} disabled={busy} type="submit">{busy ? "확인 중…" : mode === "signup" ? "가입하기" : "로그인"}</button>
        <button type="button" style={S.linkBtn} onClick={() => { setMode(mode === "signup" ? "login" : "signup"); setErr(""); setNotice(""); setPw(""); setPw2(""); }}>
          {mode === "signup" ? "이미 계정이 있으신가요? 로그인" : "계정이 없으신가요? 가입하기"}
        </button>
        {mode === "signup" && <div style={S.hint}>가입 후에도 본사 관리자가 "권한 관리"에서 권한을 부여해야 접속할 수 있습니다.</div>}
      </form>
    </div>
  );

  if (phase === "denied") return (
    <div style={S.wrap}>
      <div style={S.card}>
        <div style={S.brand}>애슐리 인건비 모니터</div>
        <div style={S.sub}>이 계정은 아직 접속 권한이 없습니다.<br />본사 관리자에게 계정 등록(권한 관리)을 요청해 주세요.</div>
        <button style={S.btn} onClick={doLogout}>다른 계정으로 로그인</button>
      </div>
    </div>
  );

  return (
    <div>
      <div style={S.topbar}>
        <span style={{ marginLeft: "auto" }}>
          <button style={S.smallBtn} onClick={doLogout}>로그아웃</button>
        </span>
      </div>
      <LaborApp />
    </div>
  );
}

const S = {
  center: { minHeight: "70vh", display: "flex", alignItems: "center", justifyContent: "center", color: "#68737E", fontFamily: "system-ui" },
  wrap: { minHeight: "100vh", display: "flex", alignItems: "center", justifyContent: "center", background: "#F3F5F2", fontFamily: "system-ui" },
  card: { background: "#fff", borderRadius: 16, padding: "32px 28px", width: 340, boxShadow: "0 12px 40px rgba(0,0,0,.12)", display: "flex", flexDirection: "column", gap: 10 },
  brand: { fontSize: 20, fontWeight: 800, color: "#182027" },
  sub: { fontSize: 13, color: "#68737E", marginBottom: 8 },
  input: { padding: "11px 12px", border: "1px solid #E1E5DE", borderRadius: 9, fontSize: 14, outline: "none" },
  btn: { padding: "11px", background: "#24478F", color: "#fff", border: "none", borderRadius: 9, fontSize: 14, fontWeight: 700, cursor: "pointer" },
  err: { fontSize: 13, color: "#C2402A" },
  notice: { fontSize: 13, color: "#2E7D32" },
  hint: { fontSize: 12, color: "#9AA3A0", marginTop: 4, textAlign: "center" },
  linkBtn: { background: "none", border: "none", color: "#24478F", fontSize: 12, cursor: "pointer", padding: "2px 0", textAlign: "center" },
  topbar: { display: "flex", alignItems: "center", padding: "8px 16px", background: "#fff", borderBottom: "1px solid #E1E5DE", fontFamily: "system-ui", fontSize: 13 },
  smallBtn: { padding: "5px 12px", background: "#F3F5F2", border: "1px solid #E1E5DE", borderRadius: 7, fontSize: 12, cursor: "pointer" },
};
