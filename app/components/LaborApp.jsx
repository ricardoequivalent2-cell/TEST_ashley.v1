"use client";
import React, { useState, useMemo } from "react";
import { supabase } from "../lib/supabase.js";
import * as XLSX from "xlsx";
import ScheduleDiagnosisFrame from "./ScheduleDiagnosisFrame.jsx";

const SUPA_BASE = "https://efqbcqlrxwgygobvlvcs.supabase.co";
const SNAP_BUCKET = "snapshots";
const DATA_URL = `${SUPA_BASE}/storage/v1/object/public/${SNAP_BUCKET}/ashley_week_2026-07-1.json`; // 폴백

let DATA_SOURCE = { name: "", url: "", loadedAt: "" }; // 현재 보고 있는 데이터 출처
// 권한 3단계: hq(본사관리자) / cell(셀 승인자, cell 필요) / store(일반매장, code 필요) — null이면 미승인(접속 불가)
let AUTH = { email: "", role: null, cell: null, code: null };
export function setAuthInfo(info) { AUTH = info || { email: "", role: null, cell: null, code: null }; }
// 마스터 계정 — approvers 테이블 상태와 무관하게 항상 본사관리자 권한으로 접속되고, 권한 관리 화면에서 회수 대상이 되지 않음
export const MASTER_EMAIL = "ahn_taehyuk01@eland.co.kr";

// 업로드 시점 기준 최신 스냅샷 URL 찾기
async function latestSnapshotUrl() {
  try {
    const { data, error } = await supabase.storage.from(SNAP_BUCKET).list("", { limit: 200, sortBy: { column: "updated_at", order: "desc" } });
    if (error || !data) return DATA_URL;
    const jsons = data.filter((f) => /^ashley_week_.+\.json$/i.test(f.name));
    if (!jsons.length) return DATA_URL;
    jsons.sort((a, b) => new Date(b.updated_at || b.created_at || 0) - new Date(a.updated_at || a.created_at || 0));
    const f = jsons[0];
    DATA_SOURCE = { name: f.name, url: `${SUPA_BASE}/storage/v1/object/public/${SNAP_BUCKET}/${f.name}`, loadedAt: (f.updated_at || f.created_at || "").slice(0, 19).replace("T", " ") };
    return DATA_SOURCE.url;
  } catch (e) { console.warn("최신 스냅샷 탐지 실패, 폴백 사용", e); return DATA_URL; }
}
let DATA = { stores: [], wages: [], helpers: [], peerClass: [], peerSlot: {}, peerWage: {}, peerRating: {}, peerWE: {}, peerTrend: {}, peerMate: {}, peerSlotHK: {} };

const GUIDE_HTML = `<!DOCTYPE html>
<html lang="ko"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>애슐리 인건비 · 데이터 계보 설명서</title>
<style>
:root{--paper:#F3F5F2;--surface:#fff;--ink:#182027;--muted:#68737E;--line:#E1E5DE;--cobalt:#24478F;--up:#24478F;--up-bg:#EEF3FB;--staff:#7A3E9D;--staff-bg:#F5EEFA;--helper:#1E7A5B;--helper-bg:#EAF5F0;--mon:#B5751A;--mon-bg:#FCF4E4;--ref:#4A5560;--ref-bg:#EFF1EE;--derive:#8A9199;--derive-bg:#F2F4F1;}
*{box-sizing:border-box;margin:0}
body{background:var(--paper);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Malgun Gothic",sans-serif;font-size:14px;line-height:1.55;padding:32px 20px}
.wrap{max-width:920px;margin:0 auto}
h1{font-size:22px;letter-spacing:-.02em}
.sub{color:var(--muted);font-size:13px;margin-top:4px}
.legend{display:flex;flex-wrap:wrap;gap:10px;margin:18px 0 24px;padding:14px 16px;background:var(--surface);border:1px solid var(--line);border-radius:12px}
.legend .t{font-size:12px;color:var(--muted);font-weight:700;width:100%;margin-bottom:2px}
.lg{display:inline-flex;align-items:center;gap:6px;font-size:12px}
.dot{width:11px;height:11px;border-radius:3px;display:inline-block}
.pipeline{display:flex;gap:0;margin-bottom:28px;background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.stage{flex:1;padding:14px 16px}
.stage+.stage{border-left:1px solid var(--line)}
.stage .n{font-size:11px;color:var(--muted);font-weight:700;margin-bottom:6px}
.stage .b{font-weight:700;font-size:13px}
.stage .d{font-size:12px;color:var(--muted);margin-top:2px}
.group{margin-bottom:26px}
.group h2{font-size:14px;font-weight:800;letter-spacing:-.01em;margin-bottom:12px;padding-bottom:6px;border-bottom:2px solid var(--ink)}
.cards{display:grid;gap:12px}
.card{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px;break-inside:avoid}
.card-h{display:flex;justify-content:space-between;align-items:baseline;gap:12px;margin-bottom:10px;flex-wrap:wrap}
.card-name{font-weight:800;font-size:15px}
.card-shows{color:var(--muted);font-size:12px;text-align:right}
.calc-row{background:var(--paper);border-radius:8px;padding:9px 12px;font-size:13px;color:#2C343C;margin-bottom:12px}
.calc-label{font-size:10px;color:var(--muted);font-weight:700;margin-right:8px;letter-spacing:.03em}
.srcgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(210px,1fr));gap:10px}
.srcbox{border:1px solid var(--line);border-radius:10px;padding:11px 12px;border-left-width:4px}
.srcbox.up{border-left-color:var(--up);background:var(--up-bg)}
.srcbox.staff{border-left-color:var(--staff);background:var(--staff-bg)}
.srcbox.helper{border-left-color:var(--helper);background:var(--helper-bg)}
.srcbox.mon{border-left-color:var(--mon);background:var(--mon-bg)}
.srcbox.ref{border-left-color:var(--ref);background:var(--ref-bg)}
.srcbox.derive{border-left-color:var(--derive);background:var(--derive-bg)}
.srcbox-h{display:flex;flex-direction:column;gap:1px;margin-bottom:7px}
.src-name{font-weight:800;font-size:13px}
.src-edit{font-size:11px;color:var(--muted)}
.src-items{margin:0;padding-left:16px;display:grid;gap:3px}
.src-items li{font-size:12px;color:#3D4650}
footer{color:var(--muted);font-size:12px;margin-top:24px;padding-top:14px;border-top:1px solid var(--line)}
@media print{body{background:#fff;padding:0}.card,.legend,.pipeline{border-color:#ccc}@page{margin:12mm}}
@media(max-width:640px){.pipeline{flex-direction:column}.stage+.stage{border-left:0;border-top:1px solid var(--line)}}
</style></head>
<body><div class="wrap">
<header>
<h1>매장 진단 · 데이터 계보 설명서</h1>
<div class="sub">매장 진단 화면의 각 칸에 들어가는 수치가 어떤 데이터의 어떤 항목에서 나오는지 세부 항목까지 정리한 문서입니다. 숫자에 이상이 있을 때 어느 항목을 확인·수정해야 하는지 진단하세요.</div>
</header>

<div class="legend">
<span class="t">데이터를 어디서 수정하나 (박스 색상)</span>
<span class="lg"><span class="dot" style="background:var(--up)"></span>모니터링실 · 차E 업로드</span>
<span class="lg"><span class="dot" style="background:var(--staff)"></span>매장 정직원 탭</span>
<span class="lg"><span class="dot" style="background:var(--helper)"></span>인건비 직접 보정 탭</span>
<span class="lg"><span class="dot" style="background:var(--mon)"></span>모니터링실 보정</span>
<span class="lg"><span class="dot" style="background:var(--ref)"></span>기준정보</span>
<span class="lg"><span class="dot" style="background:var(--derive)"></span>자동 계산(파생)</span>
</div>

<div class="pipeline">
<div class="stage"><div class="n">1단계 · 원천</div><div class="b">차E 업로드 데이터</div><div class="d">매출 · 출퇴근 · 급여 · 근무자관리 · 퇴사발령 · 평점</div></div>
<div class="stage"><div class="n">2단계 · 가공</div><div class="b">지표 계산</div><div class="d">피어그룹 분류 · 백판(정직원급여) · 헬퍼/특이사항 보정</div></div>
<div class="stage"><div class="n">3단계 · 화면</div><div class="b">매장 진단 카드</div><div class="d">아래 각 칸에 최종 수치로 표시</div></div>
</div>

<section class="group"><h2>상단 요약 카드</h2><div class="cards">
        <div class="card">
          <div class="card-h">
            <div class="card-name">주매출(이번 주)</div>
            <div class="card-shows">매장 누적 실매출 · 인건비율의 분모</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>일자별 매출을 누적 합산</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">일자별매출</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>일자별(월~일) 실매출</li><li>주간 합계 · 누적 합계</li><li>영업일수(예상매출 환산용)</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">주간 인건비율</div>
            <div class="card-shows">정직원+메이트 인건비를 매출로 나눈 값</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>(정직원급여주 + 메이트급여주 + 메이트OT주) ÷ 주매출</div>
          <div class="srcgrid"><div class="srcbox staff">
      <div class="srcbox-h"><span class="src-name">백판 (정직원급여)</span><span class="src-edit">매장 정직원 탭에서 수정</span></div>
      <ul class="src-items"><li>직위별 정직원 인원수</li><li>직위별 월임금(통상임금표 연동)</li><li>월 근무일수 반영분</li></ul>
    </div><div class="srcbox ref">
      <div class="srcbox-h"><span class="src-name">통상임금표</span><span class="src-edit">기준정보(통상임금 등)</span></div>
      <ul class="src-items"><li>직위별 월임금 단가</li><li>직위별 통상임금·시급</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트급여</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>기본급(시급 × 근무시간)</li><li>주휴수당</li><li>연차수당</li><li>야간·연장(OT)수당</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">일자별매출</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>주매출 7/1~7/5 (분모)</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">피어그룹 대비 gap</div>
            <div class="card-shows">같은 매출대 피어그룹 평균과의 인건비율 차이</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>우리 인건비율 − 피어그룹 평균(최대·최소 제외)</div>
          <div class="srcgrid"><div class="srcbox mon">
      <div class="srcbox-h"><span class="src-name">피어그룹분류</span><span class="src-edit">모니터링실 보정</span></div>
      <ul class="src-items"><li>매출구간별 피어그룹 배정</li><li>신규·특수매장 제외 플래그</li></ul>
    </div><div class="srcbox derive">
      <div class="srcbox-h"><span class="src-name">주간 인건비율</span><span class="src-edit">다른 카드에서 자동 계산</span></div>
      <ul class="src-items"><li>우리 매장 인건비율</li><li>피어그룹 내 타 매장 인건비율 평균</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">계약 대비 초과(분)</div>
            <div class="card-shows">메이트 계약시간 대비 실제 사용시간 초과분</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>주간 실사용시간(메이트) − 메이트 계약시간</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">출퇴근기록부(메이트)</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>메이트 실근무시간(출근~퇴근, 휴게 제외)</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트 퇴사발령</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>발령일 기준 계약시간 산정</li><li>퇴사자 계약시간 차감</li></ul>
    </div></div>
        </div></div></section><section class="group"><h2>이 수치는 어떻게 나왔나 · 종합판 세부</h2><div class="cards">
        <div class="card">
          <div class="card-h">
            <div class="card-name">인건비율 구성 (정직원 / 메이트 / OT)</div>
            <div class="card-shows">인건비율을 세 파트로 분해</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>각 파트 급여 ÷ 매출</div>
          <div class="srcgrid"><div class="srcbox staff">
      <div class="srcbox-h"><span class="src-name">백판 (정직원급여)</span><span class="src-edit">매장 정직원 탭에서 수정</span></div>
      <ul class="src-items"><li>정직원 직급 구성(선임점장·점장·GM·매니저·캡틴·헤드·HIT 등)</li><li>직급별 인원수</li><li>직급별 월임금</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트급여 — 메이트 파트</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>메이트 시급 × 근무시간</li><li>주휴수당</li><li>연차수당</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트급여 — OT 파트</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>야간근로수당</li><li>연장근로(추가)수당</li><li>휴일·휴업수당</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">실사용시간</div>
            <div class="card-shows">정직원·메이트 실근무시간, 주간 사용, 계약 초과</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>출퇴근 시각을 시간으로 집계</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">출퇴근기록부(정직원)</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>정직원 출근~퇴근 시각</li><li>휴게시간 제외</li><li>연차 외 휴가분 보정(→헬퍼)</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">출퇴근기록부(메이트)</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>메이트 출근~퇴근 시각</li><li>주간 누적 근무시간</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트 퇴사발령</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>계약시간 기준값</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">메이트 인원 구성 (풀타임 / 중간 / 초단기)</div>
            <div class="card-shows">근무시간대별 메이트 인원과 초단기 비율</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>주당 근무시간으로 인원 분류</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">근무자관리_n주차</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>매장·사번·직위 명단</li><li>근무 구분(정규/단기 등)</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">메이트급여</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>개인별 주당 근무시간</li><li>풀타임 / 중간 / 초단기(주 15h 미만) 판정</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">메이트 계약현황 (홀 / 홀·주방 / 주방)</div>
            <div class="card-shows">파트별 메이트 계약시간과 홀 비율</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>파트 구분별 계약시간 집계</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">근무자관리_n주차</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>파트 구분(홀 / 홀·주방 / 주방)</li><li>파트별 인원·계약시간</li><li>홀 비율 산출</li></ul>
    </div></div>
        </div></div></section><section class="group"><h2>시간대별 · 정직원 · 비교</h2><div class="cards">
        <div class="card">
          <div class="card-h">
            <div class="card-name">시간대별 인원 세팅</div>
            <div class="card-shows">오픈~마감 5구간 일평균 셋팅인원 · 피어 대비</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>출퇴근 시각을 5개 시간대에 배분 후 일평균화</div>
          <div class="srcgrid"><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">출퇴근기록부(정직원)</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>정직원 근무시각을 시간대에 배분</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">출퇴근기록부(메이트)</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>메이트 근무시각을 시간대에 배분</li><li>오픈·런치피크·스윙·디너피크·마감 구간별 인시(man-hour)</li><li>일평균 셋팅인원 환산</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">정직원 현황</div>
            <div class="card-shows">직위별 정직원 인원·월임금·정직원 비율/지수</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>백판 명부 + 통상임금 + 특이사항 반영</div>
          <div class="srcgrid"><div class="srcbox staff">
      <div class="srcbox-h"><span class="src-name">백판 (정직원급여)</span><span class="src-edit">매장 정직원 탭에서 수정</span></div>
      <ul class="src-items"><li>직위별 정직원 명부</li><li>직위별 월임금</li><li>정직원 지수·비율</li></ul>
    </div><div class="srcbox ref">
      <div class="srcbox-h"><span class="src-name">통상임금표</span><span class="src-edit">기준정보(통상임금 등)</span></div>
      <ul class="src-items"><li>직위별 월임금 단가</li></ul>
    </div><div class="srcbox staff">
      <div class="srcbox-h"><span class="src-name">정직원 특이사항</span><span class="src-edit">매장 정직원 탭에서 수정</span></div>
      <ul class="src-items"><li>육아휴직·리프레쉬·병가 등</li><li>인건비 제외/보정 반영</li></ul>
    </div></div>
        </div>
        <div class="card">
          <div class="card-h">
            <div class="card-name">피어 비교 · 제안</div>
            <div class="card-shows">우수 피어 매장과 지표 비교, 규칙 기반 제안</div>
          </div>
          <div class="calc-row"><span class="calc-label">계산식</span>위 지표 전부를 피어그룹 내에서 대조</div>
          <div class="srcgrid"><div class="srcbox mon">
      <div class="srcbox-h"><span class="src-name">피어그룹분류</span><span class="src-edit">모니터링실 보정</span></div>
      <ul class="src-items"><li>피어그룹 배정</li><li>그룹 내 인건비율 하위(우수) 매장 선정</li></ul>
    </div><div class="srcbox up">
      <div class="srcbox-h"><span class="src-name">고객평점</span><span class="src-edit">모니터링실 · 차E 업로드</span></div>
      <ul class="src-items"><li>매장 고객평점(운영 품질 확인용)</li></ul>
    </div><div class="srcbox derive">
      <div class="srcbox-h"><span class="src-name">위 모든 지표</span><span class="src-edit">다른 카드에서 자동 계산</span></div>
      <ul class="src-items"><li>정직원율·메이트율·가산율·초단기비율 대조</li></ul>
    </div></div>
        </div></div></section>

<footer>애슐리 인건비관리 · 2026년 7월 1주차 기준 · 각 박스 안의 항목이 그 칸의 수치에 합산·반영되는 세부 구성요소입니다.</footer>
</div></body></html>`;

function downloadGuide() {
  try {
    const blob = new Blob([GUIDE_HTML], { type: "text/html;charset=utf-8" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = "데이터계보_설명서.html";
    document.body.appendChild(a); a.click();
    document.body.removeChild(a); setTimeout(() => URL.revokeObjectURL(url), 1000);
  } catch (e) { window.open().document.write(GUIDE_HTML); }
}

/* ───────── helpers ───────── */
const fmtPct = (v, d = 1) => (v == null ? "–" : (v * 100).toFixed(d) + "%");
const fmtPp = (v, d = 1) => (v == null ? "–" : (v > 0 ? "+" : "") + (v * 100).toFixed(d) + "%p");
const fmtWon = (v) => (v == null ? "–" : v >= 1e7 ? (v / 1e8).toFixed(2) + "억" : Math.round(v / 1e4).toLocaleString() + "만");
const fmtWon2 = (v) => (v == null ? "–" : Math.round(v).toLocaleString() + "원");
const fmtNum = (v, d = 0) => (v == null ? "–" : Number(v).toLocaleString(undefined, { minimumFractionDigits: d, maximumFractionDigits: d }));
const fmtSign = (v, d = 1) => (v == null ? "–" : (v > 0 ? "+" : "") + Number(v).toFixed(d));

let STORES = DATA.stores;
let WAGES = DATA.wages;

// 주차 ↔ 실제 날짜 (6월4주차=6/22~28, 7월1주차=6/29~7/5)
const WK4 = "6/29~7/5";
const WK5 = "7/6~7/12";
let CUR_WEEK_LABEL = `2026년 7월 1주차 (6/29~7/5)`;
let WEEK_RANGE = "6/29~7/5"; // 표에 쓰는 날짜범위 (initData에서 갱신)
STORES.forEach((s) => { s.cell = s.senior; });
// 업로드 반영을 위한 절대금액 초기화 (비율 × 매출 → 원 단위 보관)
STORES.forEach((s) => {
  s._payStaff = (s.ls || 0) * (s.s || 0);
  s._payMate = (s.lm || 0) * (s.s || 0);
  s._payOT = (s.lo || 0) * (s.s || 0);
});
// 절대금액·매출로부터 비율/갭 재계산 (파일 업로드 후 호출)
function recomputeMetrics() {
  STORES.forEach((s) => {
    if (!s.s) return;
    const sv = s.s; // 매출 파일 자체가 이미 VAT 제외 금액이라 추가 보정 없음
    s.ls = s._payStaff / sv; s.lm = s._payMate / sv; s.lo = s._payOT / sv;
    s.lt = s.ls + s.lm; // 메이트율(_payMate)에 가산 포함 → lo는 표시용(중복합산 안 함)
  });
  // gap: 매장진단(diagnose)의 피어평균과 완전히 같은 정의를 사용 — 자기 매장 제외 + 표본 3개↑면 최대·최소 제외
  // (peersOf/avg는 아래에서 정의되지만 이 함수는 데이터 로드 후 나중에 "호출"되므로 참조 시점엔 이미 초기화되어 있음)
  STORES.forEach((s) => {
    if (s.lt == null) { s.gap = null; return; }
    if (s.pg === "신규매장" || s.pg === "노출제외") { s.gap = null; return; } // 신규매장·노출제외만 비교 제외(199매장은 199매장끼리 비교)
    const peerAvgLt = avg([s, ...peersOf(s)], "lt"); // diagnose()의 pAvg.lt와 동일한 함수·동일한 대상(본인 포함)
    s.gap = peerAvgLt == null ? null : s.lt - peerAvgLt;
  });
}
const getCells = () => [...new Set(STORES.map((s) => s.cell).filter(Boolean))].sort();
const orderStores = (list) => [...list].sort((a, b) => {
  const pgo = (s) => typeof s.pg === "number" ? s.pg : (s.pg === "199매장" ? 998 : 999);
  return pgo(a) - pgo(b) || (b.s ?? -1) - (a.s ?? -1);
});

// ===== 피어 로직 A타입(v49~v80): 매출 기준 동적 비교군 (그룹 번호 없음) =====
// ※ v81에서 "±10% 전부 + 3개이하면 위아래 2개씩"으로 바꿨다가, v81대비 롤백 요청으로 v80 방식으로 되돌림.
// 신규매장·노출제외는 비교 대상에서 완전히 빠짐(자기 자신이든, 후보로서든).
// 199매장은 유사매출대 방식이 아니라 "199매장끼리"만 서로 비교(매출 근접도·10% 제한 없음).
// 일반매장: 매출로 줄세운 뒤 자기 위치에서 매출 낮은 쪽 최대 3개 · 높은 쪽 최대 3개를 뽑고,
// 그중 자기 매출과 10% 이내인 것만 채택. 단, 10% 이내가 하나도 없으면(외곽 매장 등) 아예 비교대상이
// 없어지는 걸 막기 위해 그 6개 후보 중 매출 차이가 가장 작은 1개로 폴백(항상 최소 1개는 보장).
// pool은 보통 STORES(클라이언트) 또는 pipeBuild() 내부의 이번 주 stores 배열(관리자 업로드 미리보기) — 둘 다 이 함수 하나로 통일.
function peersBySalesBand(store, pool) {
  if (!store || !store.s) return [];
  if (store.pg === "신규매장" || store.pg === "노출제외") return [];
  if (store.pg === "199매장") return pool.filter((s) => s.c !== store.c && s.s > 0 && s.pg === "199매장");
  const eligible = pool.filter((s) => s.c !== store.c && s.s > 0 && s.pg !== "신규매장" && s.pg !== "199매장" && s.pg !== "노출제외").sort((a, b) => a.s - b.s);
  let pos = eligible.findIndex((s) => s.s > store.s);
  if (pos === -1) pos = eligible.length;
  const below = eligible.slice(Math.max(0, pos - 3), pos).reverse(); // 매출 낮은 쪽, 가까운 순으로 최대 3개
  const above = eligible.slice(pos, pos + 3); // 매출 높은 쪽, 가까운 순으로 최대 3개
  const candidates = [...below, ...above];
  const within10 = candidates.filter((s) => Math.abs(s.s - store.s) / store.s <= 0.1);
  if (within10.length) return within10;
  if (!candidates.length) return [];
  // 폴백: ±10% 이내가 없으면 후보 중 매출 차이가 가장 작은 1개만(항상 비교대상 1개는 보장)
  const nearest = candidates.reduce((best, s) => (Math.abs(s.s - store.s) < Math.abs(best.s - store.s) ? s : best), candidates[0]);
  return [nearest];
}
const peersOf = (store) => peersBySalesBand(store, STORES);
// 피어 평균(공통): 표본이 3개 이상이면 최대·최소를 제외한 트림평균, 2개 이하면 있는 값 그대로 평균
// 종합현황의 gap과 매장진단의 pAvg/슬롯비교/가산율비교가 전부 이 한 함수를 거치므로 두 화면의 "피어평균"이 항상 같은 값으로 나옴
const avg = (arr, key) => {
  const v = arr.map((x) => x[key]).filter((x) => x != null).sort((a, b) => a - b);
  const t = v.length > 2 ? v.slice(1, -1) : v;
  return t.length ? t.reduce((a, b) => a + b, 0) / t.length : null;
};

/* ───────── diagnose ───────── */
// 원인 진단 문구에서 시간·인원 차이를 강조
function hiTime(s) {
  let t = String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
  // 더 씀(+) = 빨강, 아낌(−) = 파랑
  t = t.replace(/(\+\d[\d,]*\.?\d*\s*시간)/g, '<b class="hi-over">$1</b>');
  t = t.replace(/(−\d[\d,]*\.?\d*\s*시간|-\d[\d,]*\.?\d*\s*시간)/g, '<b class="hi-under">$1</b>');
  t = t.replace(/(\+\d+(?:\.\d+)?명)/g, '<b class="hi-over">$1</b>');
  t = t.replace(/(−\d+(?:\.\d+)?명|-\d+(?:\.\d+)?명)/g, '<b class="hi-under">$1</b>');
  t = t.replace(/(\+\d[\d,]*\.?\d*%p)/g, '<b class="hi-over">$1</b>');
  t = t.replace(/(−\d[\d,]*\.?\d*%p|-\d[\d,]*\.?\d*%p)/g, '<b class="hi-under">$1</b>');
  t = t.replace(/(\d{1,3}(?:,\d{3})+원)/g, '<b class="hi-won">$1</b>');
  return t;
}

function diagnose(store) {
  if (store.pg === "신규매장")
    return { excluded: true, reason: "신규 오픈 매장은 안정화 전까지 유사매출대 비교에서 제외됩니다." };
  if (store.pg === "노출제외")
    return { excluded: true, reason: "노출제외로 지정된 매장입니다." };
  const peers = peersOf(store);
  if (peers.length < 1) return { excluded: true, reason: store.pg === "199매장" ? "비교할 다른 199매장이 없습니다." : "비교 대상 매장이 없습니다." };
  const thin = peers.length < 2;
  const withSelf = [store, ...peers]; // 평균에는 본인도 포함(요청사항) — 존재/thin 판정과 "잘하는 인근매장" 목록은 peers(본인 제외)로 유지
  const pAvg = {}; ["lt", "ls", "lm", "lo", "ur", "mateWage", "staffRatio", "lmHallBase", "lmKitBase", "useHall", "useKit", "ctHall", "ctKit"].forEach((k) => (pAvg[k] = avg(withSelf, k)));
  const diffs = { ls: store.ls - pAvg.ls, lm: store.lm - pAvg.lm, lo: store.lo - pAvg.lo };
  const slotPeerSet = store.slots.map((sl, idx) => avg(withSelf.map((p) => ({ v: p.slots[idx]?.set })), "v"));
  const best = [...peers].filter((p) => p.lt != null).sort((a, b) => a.lt - b.lt).slice(0, 2);

  const tips = [];
  const md = store.mateDetail || { 홀: {}, 주방: {} };
  const asum = (k) => ((md.홀 || {})[k] || 0) + ((md.주방 || {})[k] || 0);
  const H = (n) => `${fmtNum(n, 0)}시간`;
  const dH = (n) => (n >= 0 ? `+${fmtNum(n, 0)}` : `${fmtNum(n, 0)}`) + "시간";
  // 홀/주방 슬롯 피크 과다 (해당 파트)
  const partOver = (part) => {
    const mine = (store.slotHK && store.slotHK.wd && store.slotHK.wd[part]) || [];
    return mine.map((v, i) => {
      const p = avg(withSelf.map((q) => ({ v: (q.slotHK && q.slotHK.wd && q.slotHK.wd[part]) ? q.slotHK.wd[part][i] : 0 })), "v");
      return { k: store.slots[i] ? store.slots[i].k : "", t: store.slots[i] ? store.slots[i].t : "", mine: v, peer: p, d: v - p };
    }).filter((x) => x.d > 0.3).sort((a, b) => b.d - a.d);
  };

  // 정직원 (인원 중심 유지 — ④ 제외로 홀/주방·시간 세분화는 안 함)
  if (diffs.ls > 0.002) {
    const facts = [`정직원 인건비율 ${fmtPct(store.ls)} (유사매출대 평균 ${fmtPct(pAvg.ls)}, ${fmtPp(diffs.ls)})`];
    if (store.staff) facts.push(`정직원 ${store.staff.length}명`);
    tips.push({ tag: "정직원", level: diffs.ls > 0.006 ? "high" : "mid", head: `정직원 인건비율이 유사매출대 매장보다 ${fmtPp(diffs.ls)} 높습니다`, facts, alts: ["직급 배치 재검토(고직급 편중 여부)", "인접 매장 헬퍼 파견 / 수급 조정"] });
  }

  // 메이트 — 홀/주방 각각 원인(율 + 실사용시간 + 계약시간 + 초단기 + 피크과다)
  const matePart = (part, lmKey, useKey, ctKey) => {
    const myRate = store[lmKey] || 0, pRate = pAvg[lmKey] || 0;
    const use = store[useKey], puse = pAvg[useKey], ct = store[ctKey], pct = pAvg[ctKey];
    const rateDiff = myRate - pRate, useDiff = use != null && puse != null ? use - puse : null;
    const overPre = partOver(part);
    const isOver = rateDiff > 0.0015 || (useDiff != null && useDiff > 15) || (overPre.length && overPre[0].d >= 1);
    const isGood = rateDiff < -0.0015 || (useDiff != null && useDiff < -15);
    // 항상 표시 (사용시간이 가장 직관적) — 문제면 red, 아꼈으면 blue
    const facts = [];
    facts.push(`${part} 메이트 인건비율 ${fmtPct(myRate)} (유사매출대 ${fmtPct(pRate)}, ${fmtPp(rateDiff)})`);
    if (use != null && puse != null) facts.push(`${part} 실사용시간 ${H(use)} (유사매출대 ${H(puse)}, ${dH(useDiff)})`);
    if (ct != null && pct != null) facts.push(`${part} 계약시간 ${H(ct)} (유사매출대 ${H(pct)}, ${dH(ct - pct)})`);
    overPre.slice(0, 3).forEach((o, i) => {
      facts.push(`${i === 0 && o.d >= 1 ? "▶ " : ""}${part} ${o.k}(${o.t}) ${fmtNum(o.mine, 1)}명 — 유사매출대 ${fmtNum(o.peer, 1)}명, ${o.d >= 0 ? "+" : ""}${o.d.toFixed(1)}명`);
    });
    const alts = [];
    if (isOver && overPre.length) alts.push(`${part} ${overPre[0].k} 시간대 인원이 유사매출대 매장보다 ${overPre[0].d.toFixed(1)}명 많음 → 해당 시간대 배치·계약시간 점검`);
    if (useDiff != null && useDiff > 15) alts.push(`${part} 실사용시간이 유사매출대 매장보다 ${fmtNum(useDiff, 0)}시간 많음 → 피크 외 시간대 인원 점검`);
    if (ct != null && pct != null && ct - pct > 15) alts.push(`${part} 계약시간이 유사매출대 매장보다 ${fmtNum(ct - pct, 0)}시간 많음 → 계약시간 조정 검토`);
    const level = isOver ? (rateDiff > 0.006 ? "high" : "mid") : (isGood ? "good" : "flat");
    const head = isOver ? (overPre.length && overPre[0].d >= 1 ? `${part} 메이트 — ${overPre[0].k} 시간대 인원이 유사매출대 매장보다 많습니다` : `${part} 메이트 — 유사매출대 매장보다 더 씀`) : (isGood ? `${part} 메이트 — 유사매출대 매장보다 아낌 👍` : `${part} 메이트 — 유사매출대 매장과 비슷`);
    tips.push({ tag: `메이트·${part}`, level, head, facts, alts });
  };
  matePart("홀", "lmHallBase", "useHall", "ctHall");
  matePart("주방", "lmKitBase", "useKit", "ctKit");
  // 메이트 전체가 높은데 홀/주방 개별로 안 잡힌 경우 폴백
  if (diffs.lm > 0.002 && !tips.some((t) => t.tag.startsWith("메이트"))) {
    const facts = [`메이트 인건비율 ${fmtPct(store.lm)} (유사매출대 ${fmtPct(pAvg.lm)}, ${fmtPp(diffs.lm)})`];
    if (store.useMate != null) facts.push(`실사용시간 ${H(store.useMate)}`);
    if (store.ctMate != null) facts.push(`계약시간 ${H(store.ctMate)}`);
    if (store.ur != null) facts.push(`초단기 메이트 비율 ${fmtPct(store.ur, 0)} (유사매출대 ${fmtPct(pAvg.ur, 0)})`);
    tips.push({ tag: "메이트", level: diffs.lm > 0.006 ? "high" : "mid", head: "메이트 인건비 원인", facts, alts: [] });
  }

  // 가산 세분화 (야간/휴일/추가 각각 율 + 피어)
  const allowRate = (k) => (asum(k)) / (store.s || 1);
  const peerAllowRate = (k) => avg(withSelf.map((p) => { const m = p.mateDetail || { 홀: {}, 주방: {} }; return { v: (((m.홀 || {})[k] || 0) + ((m.주방 || {})[k] || 0)) / (p.s || 1) }; }), "v");
  if (diffs.lo > 0.0004 || asum("추가") > 0 || asum("야간") > 0 || asum("휴일") > 0) {
    const totalOT = asum("야간") + asum("휴일") + asum("추가");
    const htime = (k) => ((md.홀 || {})[k] || 0) + ((md.주방 || {})[k] || 0);
    const facts = [`가산수당 합계 ${fmtWon2(totalOT)} · 가산율 ${fmtPct(store.lo)} (유사매출대 ${fmtPct(pAvg.lo)})`];
    [["야간", "야간시간", "야간근무"], ["휴일", "휴일시간", "휴일근무"], ["추가", "추가시간", "추가근무"]].forEach(([amtK, timeK, lab]) => {
      if (asum(amtK) > 0) facts.push(`${lab} ${H(htime(timeK))} · ${fmtWon2(asum(amtK))} (율 ${fmtPct(allowRate(amtK))})`);
    });
    tips.push({ tag: "가산", level: diffs.lo > 0.0008 ? "high" : "mid", head: diffs.lo > 0.0004 ? `가산수당이 유사매출대 매장보다 ${fmtPp(diffs.lo)} 높습니다` : "가산수당이 발생했습니다", facts, alts: [asum("추가") > 0 && "추가근무 시간 발생 → 출근시간 조정(예: 30분) 검토", asum("야간") > 0 && "야간근무 시간 발생 → 요일별 분산 검토", asum("휴일") > 0 && "휴일근무 시간 발생 → 휴일 스케줄 점검"].filter(Boolean) });
  }

  if (!tips.length && store.gap != null && store.gap <= 0) {
    tips.push({ tag: "양호", level: "good", head: "유사매출대 평균보다 인건비율이 낮습니다", facts: [`인건비율 ${fmtPct(store.lt)} (유사매출대 평균 ${fmtPct(pAvg.lt)}, ${fmtPp(store.gap)})`], alts: [] });
  }
  return { excluded: false, thin, peers, pAvg, diffs, best, slotPeerSet, tips };
}

/* ───────── atoms ───────── */
const GapBar = ({ gap, max = 0.02, w: width = 120 }) => {
  const w = gap == null ? 0 : Math.min(Math.abs(gap) / max, 1) * 50;
  const over = gap > 0;
  return (
    <span className="gapbar" style={{ width }} title={fmtPp(gap)}>
      <span className="gapbar-axis" />
      <span className={"gapbar-fill " + (over ? "over" : "under")} style={over ? { left: "50%", width: w + "%" } : { right: "50%", width: w + "%" }} />
    </span>
  );
};
const Stat = ({ label, value, sub, tone, tip }) => (
  <div className={"stat" + (tone ? " tone-" + tone : "") + (tip ? " has-gaptip" : "")}>
    <div className="stat-label">{label}{tip && <span className="th-info">ⓘ</span>}</div>
    <div className="stat-value">{value}</div>
    {sub && <div className="stat-sub">{sub}</div>}
    {tip && <div className="gaptip"><div className="gaptip-title">유사매출대란</div><div className="gaptip-desc">{tip}</div></div>}
  </div>
);

/* ───────── 1. 종합 현황 ───────── */
function Dashboard({ goDetail }) {
  const [cell, setCell] = useState("전체");
  const [grp, setGrp] = useState("전체");
  const [sortKey, setSortKey] = useState("peer");
  const [q, setQ] = useState("");
  const [redOnly, setRedOnly] = useState(false);
  const [capturing, setCapturing] = useState(false);
  const captureRef = React.useRef(null);
  const [period, setPeriod] = useState("주간"); // 주간 | 월간(=직전4주, 달력 월과 무관하게 게시 순서 기준 최근 4개)
  const TRAIL_N = 4;
  // 월간(직전4주) 집계: 게시 순서 기준 최근 TRAIL_N개 주차(제외 표시된 주는 건너뛰고 그 앞주까지 당겨서 채움) 매출가중 평균
  const periodStores = useMemo(() => {
    if (period !== "월간") return null;
    const idxs = []; let totalInPeriod = 0;
    for (let i = WEEK_LABELS.length - 1; i >= 0 && idxs.length < TRAIL_N; i--) {
      if (!WEEK_LABELS[i]) continue;
      totalInPeriod++;
      if (!EXCLUDED_WEEKS.has(WEEK_LABELS[i])) idxs.push(i);
    }
    idxs.reverse();
    const list = STORES.map((s) => {
      let sumPay = 0, sumSales = 0, sumOT = 0, weeksCount = 0, hoursSum = 0, hoursWeeks = 0;
      idxs.forEach((i) => {
        const lt = s.trend && s.trend.lt && s.trend.lt[i], sl = s.trend && s.trend.sales && s.trend.sales[i];
        const lo = s.trend && s.trend.lo && s.trend.lo[i], wk = s.trend && s.trend.wkTot && s.trend.wkTot[i];
        if (lt != null && sl != null) { sumPay += lt * sl; sumSales += sl; weeksCount++; }
        if (lo != null && sl != null) sumOT += lo * sl;
        if (wk != null) { hoursSum += wk; hoursWeeks++; } // hoursWeeks=0이면 "값이 0"이 아니라 "데이터 자체가 없음"
      });
      return { ...s, lt: sumSales > 0 ? sumPay / sumSales : null, lo: sumSales > 0 ? sumOT / sumSales : null, s: sumSales || null, _periodWeeks: weeksCount, _periodHours: hoursWeeks > 0 ? hoursSum : null, _periodHoursWeeks: hoursWeeks };
    });
    // 월간(직전4주) 기준 유사매출대 gap도 같은 방식(본인 포함 트림평균)으로 재계산
    list.forEach((ps) => {
      if (ps.pg === "신규매장" || ps.pg === "노출제외" || !ps.s || ps.lt == null) { ps.gap = null; return; }
      const peers = peersBySalesBand(ps, list);
      const pAvg = avg([ps, ...peers], "lt");
      ps.gap = pAvg == null ? null : ps.lt - pAvg;
    });
    const labelsUsed = idxs.map((i) => WEEK_LABELS[i]);
    return { list, weekCount: idxs.length, totalWeeksInPeriod: totalInPeriod, labelsUsed };
  }, [period]);

  // 화면 캡처: 캡처 대상(captureRef) DOM의 실제 렌더링 크기를 그대로 이미지로 뜸
  // → 필터로 표시되는 매장 수(행)가 줄어들면 캡처 높이도 자동으로 같이 줄어듦(고정 크기 아님)
  const captureScreen = async () => {
    const root = captureRef.current;
    if (!root || capturing) return;
    setCapturing(true);
    const wrap = root.querySelector(".tablewrap");
    const prevStyle = wrap ? { overflow: wrap.style.overflow, width: wrap.style.width, maxWidth: wrap.style.maxWidth } : null;
    try {
      const { default: html2canvas } = await import("html2canvas");
      if (wrap) {
        // 표가 가로 스크롤 중이면(컬럼이 화면보다 넓으면) 잘리지 않도록 캡처 직전에 실제 전체 폭만큼 펼침
        wrap.style.overflow = "visible";
        wrap.style.width = wrap.scrollWidth + "px";
        wrap.style.maxWidth = "none";
        await new Promise((r) => requestAnimationFrame(r));
      }
      const canvas = await html2canvas(root, {
        backgroundColor: "#F3F5F2",
        scale: Math.min(2, (typeof window !== "undefined" && window.devicePixelRatio) || 1.5),
        useCORS: true,
        ignoreElements: (el) => el.classList && el.classList.contains("no-capture"),
      });
      const url = canvas.toDataURL("image/png");
      const a = document.createElement("a");
      a.href = url;
      a.download = `애슐리_종합현황_${(CUR_WEEK_LABEL || "").replace(/[^\w가-힣]/g, "_")}.png`;
      document.body.appendChild(a); a.click(); a.remove();
    } catch (e) {
      alert("화면 캡처에 실패했습니다: " + e.message);
    } finally {
      if (wrap && prevStyle) { wrap.style.overflow = prevStyle.overflow; wrap.style.width = prevStyle.width; wrap.style.maxWidth = prevStyle.maxWidth; }
      setCapturing(false);
    }
  };

  const pgOrder = (v) => (typeof v === "number" ? 0 : v === "199매장" ? 900 : 901); // A타입: 그룹번호는 정렬에 안 씀(특수분류만 뒤로)
  const source = period !== "주간" && periodStores ? periodStores.list : STORES;
  const rows = useMemo(() => {
    let r = source.filter((s) => s.s != null && s.pg !== "노출제외");
    if (redOnly) r = r.filter((s) => s.gap != null && s.gap >= 0.01);
    if (cell !== "전체") r = r.filter((s) => s.cell === cell);
    if (grp !== "전체") r = r.filter((s) => (grp === "일반" ? typeof s.pg === "number" : String(s.pg) === grp));
    if (q) r = r.filter((s) => (s.n || "").includes(q));
    if (sortKey === "peer")
      // A타입: 신규매장·199매장·노출제외만 뒤로, 나머지는 매출 높은 순(=실제 피어 비교 순서와 동일하게 나열)
      return [...r].sort((a, b) => pgOrder(a.pg) - pgOrder(b.pg) || (b.s ?? -1) - (a.s ?? -1));
    if (redOnly) return [...r].sort((a, b) => (b.gap ?? -9) - (a.gap ?? -9)); // 빨간불은 gap 큰 순
    return [...r].sort((a, b) => (b[sortKey] ?? -9) - (a[sortKey] ?? -9));
  }, [source, cell, grp, sortKey, q, redOnly]);

  const VIS = source.filter((s) => s.pg !== "노출제외");
  const totSales = VIS.reduce((a, s) => a + (s.s || 0), 0);
  const buRate = VIS.reduce((a, s) => a + (s.lt || 0) * (s.s || 0), 0) / totSales;
  const band12 = VIS.filter((s) => s.gap != null && s.gap >= 0.01 && s.gap < 0.02).length;
  const band2 = VIS.filter((s) => s.gap != null && s.gap >= 0.02).length;

  const COLS = [
    { h: "점장", get: (s) => s.mg || <span className="dim">–</span>, cls: "cell" },
    { h: "셀구분", get: (s) => s.cell || <span className="dim">–</span>, cls: "cell" },
    { h: "주간\n인건비율", get: (s) => fmtPct(s.lt), cls: "num strong", tone: (s) => (s.gap > 0 ? "over" : "under"), tip: `인건비율 합 = 정직원율 + 메이트율 (${WEEK_RANGE})` },
    { h: "gap\n(유사매출대 대비)", get: (s) => <span><GapBar gap={s.gap} /><span className={"gapnum " + (s.gap > 0 ? "c-over" : "c-under")}>{fmtPp(s.gap, 1)}</span></span>, cls: "gapcol", tip: "우리 인건비율 − 매출 유사 매장 평균(위·아래 최대 3개씩, ±10% 이내, 최대·최소 제외). +면 유사매출대 평균보다 높음(불리)." },
    { h: "주간\n매출", get: (s) => fmtWon(s.s), cls: "num", tip: `${WEEK_RANGE} 매출` },
    { h: "주간\n사용시간\n(정직원+메이트)", get: (s) => fmtNum(s.wkTot, 1), cls: "num", tip: `${WEEK_RANGE} 정직원+메이트 근무시간(h)` },
    { h: "주간\n사용시간\n(메이트)", get: (s) => fmtNum(s.wkMate, 1), cls: "num", tip: `${WEEK_RANGE} 메이트 근무시간(h)` },
    { h: "주간\n메이트\n인건비율", get: (s) => fmtPct((s.lm || 0) + (s.lo || 0), 2), cls: "num", tip: `메이트 인건비 전체(기본+가산) ÷ 매출 (${WEEK_RANGE})` },
    { h: "주간\n가산율", get: (s) => fmtPct(s.lo), cls: "num", tip: `메이트 가산수당(야간+추가+휴일) ÷ 매출 (${WEEK_RANGE})` },
    { h: "이번 주\n메이트\n계약시간", get: (s) => fmtNum(s.ct4, 1), cls: "num", tip: `이번 주(${WEEK_RANGE}) 메이트 계약시간(h) · 저번 주에 올린 근무자관리 파일 기준` },
    { h: "이번 주\n추가근무\n(메이트)", get: (s) => fmtNum(s.mateExtraH, 1), cls: "num", tip: "이번 주 메이트 추가근무시간(h) · 메이트급여(주간) 파일 '추가' 항목 기준. 지난주 계약시간과 무관하게 이번 주 데이터만으로 계산." },
    { h: "다음 주\n메이트\n계약시간", get: (s) => fmtNum(s.ct5, 1), cls: "num", tip: "다음 주 메이트 계약시간(h) · 이번 주에 올린 근무자관리 파일 기준(이번 주 월요일 업로드 = 다음 월~일 스케줄)" },
    { h: "메이트\n평균시급", get: (s) => s.realWage ? fmtNum(s.realWage) + "원" : "–", cls: "num", tip: "실지급 총급여(주휴·연차·야간·추가·휴일·휴업 포함) ÷ 정상근무시간" },
  ];

  // 모바일(터치 화면)엔 :hover가 없어서 ⓘ 팝업이 안 열림 → 탭으로 열고 닫기
  const handleTapPopup = (e) => {
    const trigger = e.target.closest(".has-gaptip, .has-mtip, .has-staffpop2, .has-wagepop");
    document.querySelectorAll(".tap-open").forEach((el) => { if (el !== trigger) el.classList.remove("tap-open"); });
    if (trigger) trigger.classList.toggle("tap-open");
  };

  return (
    <div ref={captureRef} onClick={handleTapPopup}>
      <div className="statgrid three">
        <Stat label="애슐리 주간 인건비율" value={fmtPct(buRate)} />
        <Stat label="유사매출대 대비 +1~2%p 초과" value={band12 + "개"} tone="warn" tip="매출 순위 기준 위·아래 최대 3개씩(자기 매출 대비 ±10% 이내)인 매장들의 평균 인건비율보다 1~2%p 더 높은 매장 수입니다." />
        <Stat label="유사매출대 대비 +2%p 이상 초과" value={band2 + "개"} tone="over" tip="매출 순위 기준 위·아래 최대 3개씩(자기 매출 대비 ±10% 이내)인 매장들의 평균 인건비율보다 2%p 이상 더 높은 매장 수입니다." />
      </div>

      <div className="toolbar" style={{ marginBottom: 8 }}>
        <div className="periodtoggle">
          <button className={period === "주간" ? "on" : ""} onClick={() => setPeriod("주간")}>주간</button>
          <button className={period === "월간" ? "on" : ""} onClick={() => setPeriod("월간")}>월간(직전4주)</button>
        </div>
        {period === "월간" && periodStores && periodStores.weekCount < TRAIL_N && (
          <span className="dimtxt" style={{ fontSize: 12, color: "var(--warn, #b45309)" }}>
            ⚠ 직전 {TRAIL_N}주 중 {periodStores.weekCount}주만 집계됨 ({periodStores.totalWeeksInPeriod - periodStores.weekCount}주는 "주차 관리"에서 제외 표시됨)
            {periodStores.labelsUsed.length > 0 && ` — 집계 대상: ${periodStores.labelsUsed.join(", ")}`}
          </span>
        )}
      </div>
      <div className="toolbar">
        <select value={cell} onChange={(e) => setCell(e.target.value)}><option>전체</option>{getCells().map((c) => <option key={c}>{c}</option>)}</select>
        <select value={grp} onChange={(e) => setGrp(e.target.value)}><option>전체</option><option value="일반">일반매장</option><option value="신규매장">신규매장</option><option value="199매장">199매장</option><option value="노출제외">노출제외</option></select>
        <select value={sortKey} onChange={(e) => setSortKey(e.target.value)}><option value="peer">매출 순위(비교군 기준)</option><option value="gap">gap 큰 순</option><option value="lt">인건비율 높은 순</option><option value="s">매출 높은 순</option>{period === "주간" && <option value="mateExtraH">추가근무 큰 순</option>}</select>
        <input placeholder="매장 검색" value={q} onChange={(e) => setQ(e.target.value)} />
        <button className={"redfilter" + (redOnly ? " on" : "")} onClick={() => setRedOnly((v) => !v)} title="유사매출대 매장 평균보다 인건비율 +1%p 이상 높은 매장만">🔴 빨간불만</button>
        <button className="ghost no-capture" disabled={capturing} onClick={captureScreen} title="현재 필터가 적용된 화면 그대로를 이미지 파일로 저장합니다">{capturing ? "캡처 중…" : "📸 화면 캡처"}</button>
        <span className="count">{rows.length}개</span>
      </div>
      {period !== "주간" ? (
        <div className="tablewrap">
          <table className="wide">
            <thead><tr><th className="stick">매장</th><th>셀</th><th>점장</th><th className="num">월 매출</th><th className="num">월 인건비율</th><th className="gapcol">gap(유사매출대 대비)</th><th className="num">월 가산율</th><th className="num">월 사용시간</th></tr></thead>
            <tbody>
              {(() => {
                const catIdx = {};
                return rows.map((s) => {
                  const cat = typeof s.pg === "number" ? "일반" : s.pg;
                  const i = (catIdx[cat] = (catIdx[cat] ?? -1) + 1);
                  const stripe = Math.floor(i / 5) % 2 === 1;
                  return (
                    <tr key={s.c} onClick={() => goDetail(s.c)} className={stripe ? "pg-stripe" : ""}>
                      <td className="stick"><span className="storename">{s.n}</span><span className="storecode">{s.c}</span></td>
                      <td>{s.cell || "–"}</td>
                      <td>{s.mg || "–"}</td>
                      <td className="num">{s.s ? fmtWon(s.s) : "–"}</td>
                      <td className="num strong">{fmtPct(s.lt)}</td>
                      <td className="gapcol"><GapBar gap={s.gap} /><span className={"gapnum " + (s.gap > 0 ? "c-over" : "c-under")}>{fmtPp(s.gap)}</span></td>
                      <td className="num">{fmtPct(s.lo)}</td>
                      <td className="num">{s._periodHours != null ? `${fmtNum(s._periodHours, 1)}${s._periodHoursWeeks < TRAIL_N ? ` (${s._periodHoursWeeks}주치)` : ""}` : "–"}</td>
                    </tr>
                  );
                });
              })()}
            </tbody>
          </table>
        </div>
      ) : (
      <div className="tablewrap">
        <table className="wide">
          <thead><tr><th className="stick">매장</th>{COLS.map((c) => (
            <th key={c.h} className={(c.cls && c.cls.includes("num") ? "num" : "") + (c.cls && c.cls.includes("refstart") ? " refstart" : "") + (c.cls && c.cls.includes("soft") ? " soft" : "") + (c.tip ? " has-tip" : "")}>
              {c.h.split("\n").map((l, i) => <div key={i}>{l}</div>)}
              {c.tip && <span className="th-info">ⓘ</span>}
              {c.tip && <span className="th-pop">{c.tip}</span>}
            </th>
          ))}</tr></thead>
          <tbody>
            {(() => {
              const catIdx = {}; // 일반/199매장/신규매장 구간별로 각각 0부터 세서 5개 단위로 음영(구간 바뀌면 리셋)
              return rows.map((s) => {
                const cat = typeof s.pg === "number" ? "일반" : s.pg;
                const i = (catIdx[cat] = (catIdx[cat] ?? -1) + 1);
                const stripe = Math.floor(i / 5) % 2 === 1;
                return (
                  <tr key={s.c} onClick={() => goDetail(s.c)} className={stripe ? "pg-stripe" : ""}>
                    <td className="stick"><span className="storename">{s.n}</span><span className="storecode">{s.c}</span></td>
                    {COLS.map((c) => <td key={c.h} className={c.cls + (c.tone && c.tone(s) ? " c-" + c.tone(s) : "")}>{c.get(s)}</td>)}
                  </tr>
                );
              });
            })()}
          </tbody>
        </table>
      </div>
      )}
    </div>
  );
}

/* ───────── 2. 매장 진단 ───────── */
/* ───────── 주차별 추이 (지표 드롭다운) ───────── */
let WEEK_LABELS = ["", "", "6월 4주차", "7월 1주차"]; // 매주 게시(publish)할 때마다 한 칸씩 늘어남(initData에서 갱신)
let EXCLUDED_WEEKS = new Set(); // 월별 집계에서 제외할 주차 라벨(오류로 판단된 주차) — Supabase week_flags 테이블에서 로드
function setExcludedWeeks(labels) { EXCLUDED_WEEKS = new Set(labels || []); }
const TREND_METRICS = [
  { k: "lt", label: "인건비율", mode: "line", fmt: (v) => fmtPct(v), real: false, invert: true },
  { k: "lo", label: "가산율", mode: "line", fmt: (v) => fmtPct(v), real: false, invert: true },
  { k: "wage", label: "메이트 평균시급", mode: "wagebar", fmt: (v) => fmtNum(v) + "원" },
  { k: "rt", label: "고객평점", mode: "ratingbar", fmt: (v) => (v == null ? "–" : v.toFixed(2)) },
];
const WAGE_PARTS = [
  { k: "base", label: "기본시급", color: "var(--cobalt)" },
  { k: "weekly", label: "주휴수당", color: "var(--under)" },
  { k: "extra", label: "추가근무", color: "var(--warn)" },
  { k: "night", label: "야간", color: "#7B61A8" },
  { k: "etc", label: "기타", color: "#9AA4AD" },
];
function TrendChart({ store }) {
  const [mk, setMk] = useState("lt");
  const metric = TREND_METRICS.find((m) => m.k === mk);
  const pgKey = (store.pg !== "신규매장" && store.pg !== "노출제외" && store.s) ? store.c : null; // A타입: 그룹번호 대신 매장코드로 키(199매장은 199매장끼리 비교되므로 여기서 제외 안 함)

  const Picker = (
    <select value={mk} onChange={(e) => setMk(e.target.value)}>
      {TREND_METRICS.map((m) => <option key={m.k} value={m.k}>{m.label}</option>)}
    </select>
  );

  // ── 메이트 평균시급: 구성요소 누적 막대 (우리 vs 피어) ──
  if (metric.mode === "wagebar") {
    const wb = store.wageBreak;
    const wagePeers = peersOf(store).filter((p) => p.wageBreak);
    const pw = wagePeers.length ? (() => { const withSelf = [store, ...wagePeers]; const pick = (k) => avg(withSelf.map((p) => ({ v: p.wageBreak && p.wageBreak[k] })), "v"); return { base: pick("base") || 0, weekly: pick("weekly") || 0, extra: pick("extra") || 0, night: pick("night") || 0, etc: pick("etc") || 0, total: pick("total") || 0 }; })() : null;
    const peerName = "매출 유사 매장 평균";
    if (!wb) return <div className="trendbox"><div className="trend-top">{Picker}</div><div className="notice">시급 데이터가 없는 매장입니다.</div></div>;
    const rows = [{ name: store.n, tag: "우리 매장", wb, me: true }];
    if (pw) rows.push({ name: peerName, tag: "유사매출대 평균", wb: pw, me: false });
    const mx = Math.max(...rows.map((r) => r.wb.total), 1);
    return (
      <div className="trendbox">
        <div className="trend-top">{Picker}<div className="trend-now"><span className="trend-now-val">{fmtNum(wb.total)}원</span><span className="trend-now-d">{store.n} · 근로시간당(주휴·수당 포함)</span></div><span className="trend-badge real">실측</span></div>
        <div className="wagebars">
          {rows.map((r, ri) => (
            <div key={ri} className="wagebar-row">
              <div className="wagebar-name"><span className={"barlabel-tag " + (r.me ? "me" : "peer")}>{r.tag}</span>{r.name}</div>
              <div className="wagebar-track">
                {WAGE_PARTS.map((p) => r.wb[p.k] > 0 ? <div key={p.k} className="wagebar-seg" style={{ width: (r.wb[p.k] / mx) * 100 + "%", background: p.color }} title={`${p.label}: ${fmtNum(r.wb[p.k])}원`} /> : null)}
                <span className="wagebar-total">{fmtNum(r.wb.total)}원</span>
              </div>
            </div>
          ))}
        </div>
        <div className="wage-legend">
          {WAGE_PARTS.map((p) => <span key={p.k} className="wl-item"><span className="wl-box" style={{ background: p.color }} />{p.label}</span>)}
        </div>
        <div className="wage-appendix">
          <table className="mini">
            <thead><tr><th>구성요소</th><th className="num">{store.n}</th>{pw && <th className="num">{peerName}</th>}{pw && <th className="num">차이</th>}</tr></thead>
            <tbody>
              {WAGE_PARTS.map((p) => (
                <tr key={p.k}><td><span className="wl-box" style={{ background: p.color }} /> {p.label}</td><td className="num">{fmtNum(wb[p.k])}원</td>{pw && <td className="num soft">{fmtNum(pw[p.k])}원</td>}{pw && <td className={"num " + (wb[p.k] - pw[p.k] > 0 ? "c-over" : "c-under")}>{fmtSign(wb[p.k] - pw[p.k], 0)}원</td>}</tr>
              ))}
              <tr className="sumline"><td>평균 시급</td><td className="num strong">{fmtNum(wb.total)}원</td>{pw && <td className="num soft">{fmtNum(pw.total)}원</td>}{pw && <td className={"num " + (wb.total - pw.total > 0 ? "c-over" : "c-under")}>{fmtSign(wb.total - pw.total, 0)}원</td>}</tr>
            </tbody>
          </table>
        </div>
        <div className="trend-note">막대 위에 커서를 올리면 각 구성요소 금액이 표시됩니다. 아래 표에 우리 매장·유사매출대 평균·차이를 정리했습니다. 연차수당은 이번 급여주기 전원 0원이라 제외됩니다. 주차별 시급 추이는 매주 데이터가 쌓이면 자동 표시됩니다.</div>
      </div>
    );
  }

  // ── 고객평점: 막대 (우리 vs 피어) ──
  if (metric.mode === "ratingbar") {
    const mine = store.rt;
    const ratingPeers = peersOf(store).filter((p) => p.rt != null);
    const pr = (mine != null && ratingPeers.length) ? avg([store, ...ratingPeers].map((p) => ({ v: p.rt })), "v") : null;
    if (mine == null) return <div className="trendbox"><div className="trend-top">{Picker}</div><div className="notice">고객평점 데이터가 없는 매장입니다.</div></div>;
    const rows = [{ name: store.n, tag: "우리 매장", v: mine, me: true }];
    if (pr != null) rows.push({ name: "매출 유사 매장 평균", tag: "유사매출대 평균", v: pr, me: false });
    const better = pr != null && mine >= pr;
    return (
      <div className="trendbox">
        <div className="trend-top">{Picker}<div className="trend-now"><span className="trend-now-val">★ {mine.toFixed(2)}</span>{pr != null && <span className={"trend-now-d " + (better ? "c-under" : "c-over")}>{store.n} · 유사매출대 평균 {pr.toFixed(2)} 대비 {better ? "높음" : "낮음"}</span>}</div><span className="trend-badge real">실측</span></div>
        <div className="ratingbars">
          {rows.map((r, ri) => (
            <div key={ri} className="rating-row2">
              <div className="rating-name"><span className={"barlabel-tag " + (r.me ? "me" : "peer")}>{r.tag}</span>{r.name}</div>
              <div className="rating-track"><div className={"rating-fill " + (r.me ? (better ? "good" : "bad") : "peer")} style={{ width: (r.v / 5) * 100 + "%" }} /></div>
              <div className="rating-val">★ {r.v.toFixed(2)}</div>
            </div>
          ))}
        </div>
        <div className="trend-note">5점 만점 고객평점. 위가 우리 매장, 아래가 유사매출대 평균. 주차별 평점 추이는 매주 데이터가 쌓이면 자동 표시됩니다.</div>
      </div>
    );
  }

  // ── 라인 추이 (인건비율 / 가산율) ──
  const t = store.trend;
  const rawData = t && t[mk];
  if (!rawData) return <div className="trendbox"><div className="trend-top">{Picker}</div><div className="notice">주차별 데이터가 없는 매장입니다.</div></div>;
  const peerFullRaw = (pgKey && DATA.peerTrend && DATA.peerTrend[pgKey] && DATA.peerTrend[pgKey][mk]) || null;
  // "주차 관리"에서 제외 표시한 주차는 잘못된 데이터라는 뜻이므로, 값뿐 아니라 x축 라벨도 통째로 뺌(null 처리만 하면
  // 점은 안 보여도 라벨은 남아서 빈 자리가 생김 — 아예 배열에서 제외해 나머지 주차들이 자연스럽게 이어지게 함)
  // labelOffset: 이 매장의 추이 배열이 전체 WEEK_LABELS보다 짧을 수 있어서(예: 신규매장이라 뒤늦게 쌓이기 시작한 경우),
  // 항상 "맨 뒤(최신)"부터 맞춰 정렬한다 — 앞에서부터 맞추면 최신 주차 라벨이 밀려서 엉뚱하게 붙는 문제가 있었음.
  const labelOffset = WEEK_LABELS.length - rawData.length;
  const validIdx = []; rawData.forEach((v, i) => { const li = i + labelOffset; if (li >= 0 && !EXCLUDED_WEEKS.has(WEEK_LABELS[li])) validIdx.push(i); });
  const data = validIdx.map((i) => rawData[i]);
  const dataLabels = validIdx.map((i) => WEEK_LABELS[i + labelOffset]);
  const peerFull = peerFullRaw ? validIdx.map((i) => peerFullRaw[i]) : null;

  // 전주 대비·피어 대비는 항상 "쌓인 전체 데이터의 마지막 값" 기준 (그래프 표시 구간과 무관하게 최신 주차 기준)
  const last = data[data.length - 1];
  const prev = data.length > 1 ? data[data.length - 2] : null;
  const delta = prev != null && last != null ? last - prev : null;
  const worse = delta != null && (metric.invert ? delta > 0 : delta < 0);
  const peerLast = peerFull ? peerFull[peerFull.length - 1] : null;
  const vsPeer = peerLast != null && last != null ? last - peerLast : null;
  const abovePeer = vsPeer != null && (metric.invert ? vsPeer > 0 : vsPeer < 0);

  // 그래프에는 최근 N주만 표시(주차가 계속 쌓이면 x축이 빽빽해지므로) — 원본 데이터는 계속 보관됨
  const MAX_SHOW = 12;
  const startIdx = Math.max(0, data.length - MAX_SHOW);
  const shown = data.slice(startIdx);
  const shownPeer = peerFull ? peerFull.slice(startIdx) : null;
  const shownLabels = dataLabels.slice(startIdx, data.length);
  const allVals = [...shown, ...(shownPeer || [])].filter((v) => v != null);
  const minV = allVals.length ? Math.min(...allVals) : 0, maxV = allVals.length ? Math.max(...allVals) : 1;
  const pad = (maxV - minV) * 0.2 || Math.abs(maxV) * 0.1 || 1;
  const loY = minV - pad, hiY = maxV + pad;
  const W = 640, H = 210, PL = 30, PR = 30, PT = 16, PB = 40;
  const x = (i) => (shown.length > 1 ? PL + (i / (shown.length - 1)) * (W - PL - PR) : W / 2);
  const y = (v) => PT + (1 - (v - loY) / (hiY - loY)) * (H - PT - PB);
  const line = (arr) => arr.map((v, i) => (v == null ? "" : (i && arr[i - 1] != null ? "L" : "M") + x(i).toFixed(1) + " " + y(v).toFixed(1))).join(" ");
  const pts = shown.map((v, i) => [x(i), y(v)]);

  return (
    <div className="trendbox">
      <div className="trend-top">
        {Picker}
        <div className="trend-now">
          <span className="trend-now-val">{metric.fmt(last)}</span>
          <span className={"trend-now-d " + (delta == null ? "" : delta === 0 ? "" : worse ? "c-over" : "c-under")}>
            {delta == null ? "이전 주차 데이터 없음" : delta === 0 ? "전주와 동일" : `전주 대비 ${metric.fmt(Math.abs(delta))} ${worse ? "악화" : "개선"}`}
          </span>
        </div>
        {vsPeer != null && (
          <div className={"trend-vspeer " + (abovePeer ? "bad" : "good")}>
            유사매출대 평균 대비 {metric.fmt(Math.abs(vsPeer))} {abovePeer ? "높음" : "낮음"}
          </div>
        )}
        <span className="trend-badge">추정</span>
      </div>
      <svg viewBox={`0 0 ${W} ${H}`} className="trend-svg" preserveAspectRatio="none">
        <line x1={PL} y1={H - PB} x2={W - PR} y2={H - PB} stroke="var(--line)" />
        {shownPeer && <path d={line(shownPeer)} fill="none" stroke="var(--muted)" strokeWidth="2" strokeDasharray="5 4" />}
        <path d={line(shown)} fill="none" stroke={worse ? "var(--over)" : "var(--under)"} strokeWidth="2.5" />
        {shownPeer && shownPeer.map((v, i) => v == null ? null : <circle key={"p" + i} cx={x(i)} cy={y(v)} r="3" fill="var(--muted)" />)}
        {pts.map((p, i) => (
          <g key={i}>
            <circle cx={p[0]} cy={p[1]} r={i === pts.length - 1 ? 5 : 3.5} fill={worse ? "var(--over)" : "var(--under)"} />
            <text x={p[0]} y={p[1] - 10} textAnchor="middle" className="trend-val">{metric.fmt(shown[i])}</text>
            <text x={p[0]} y={H - 20} textAnchor="middle" className="trend-xlabel">{(shownLabels[i] || "").split("(")[0]}</text>
          </g>
        ))}
      </svg>
      <div className="trend-legend">
        <span className="lg-item"><span className="lg-line" style={{ background: worse ? "var(--over)" : "var(--under)" }} />우리 매장</span>
        <span className="lg-item"><span className="lg-line dash" />매출 유사 매장 평균</span>
      </div>
      <div className="trend-note">
        점선은 유사매출대 평균. 매주 업로드가 쌓일수록 추세선이 길어집니다{data.length > MAX_SHOW ? " (전체 주차 데이터는 계속 보관되고, 그래프에는 최근 12주만 표시됩니다)" : ""}.
      </div>
    </div>
  );
}

function Detail({ code, setCode, goStaff, refresh }) {
  const topSalesStore = [...STORES].sort((a, b) => (b.s ?? -1) - (a.s ?? -1))[0];
  const store = STORES.find((s) => s.c === code) || topSalesStore;
  const [, localBump] = React.useReducer((x) => x + 1, 0);
  const [dayType, setDayType] = useState("평일");
  const [slotPart, setSlotPart] = useState("종합");
  const [showReport, setShowReport] = useState(false);
  const [allowView, setAllowView] = useState("금액"); // 수당 상세: 금액/시간
  const [cmpView, setCmpView] = useState("율");       // 1:1 메이트: 율/시간
  const [storeQuery, setStoreQuery] = useState(""); // 매장 검색(그룹이 없어져서 이름으로 빠르게 찾기용)
  const captureRef = React.useRef(null);
  const [capturingDetail, setCapturingDetail] = useState(false);
  const captureDetail = async () => {
    const root = captureRef.current;
    if (!root || capturingDetail) return;
    setCapturingDetail(true);
    try {
      const { default: html2canvas } = await import("html2canvas");
      const canvas = await html2canvas(root, {
        backgroundColor: "#F3F5F2",
        scale: Math.min(2, (typeof window !== "undefined" && window.devicePixelRatio) || 1.5),
        useCORS: true,
        ignoreElements: (el) => el.classList && el.classList.contains("no-capture"),
      });
      const url = canvas.toDataURL("image/png");
      const a = document.createElement("a");
      a.href = url;
      a.download = `애슐리_매장진단_${store.n}_${(CUR_WEEK_LABEL || "").replace(/[^\w가-힣]/g, "_")}.png`;
      document.body.appendChild(a); a.click(); a.remove();
    } catch (e) {
      alert("화면 캡처에 실패했습니다: " + e.message);
    }
    setCapturingDetail(false);
  };
  // 1:1 비교 대상 — 기본값은 같은 피어그룹 내 매출이 가장 비슷한 매장
  const closestPeer = useMemo(() => {
    const pool = peersOf(store);
    if (!pool.length) return null;
    return [...pool].sort((a, b) => Math.abs((a.s || 0) - (store.s || 0)) - Math.abs((b.s || 0) - (store.s || 0)))[0];
  }, [store]);
  const [cmpCode, setCmpCode] = useState(closestPeer ? closestPeer.c : null);
  React.useEffect(() => { setCmpCode(closestPeer ? closestPeer.c : null); }, [closestPeer]);
  const cmp = STORES.find((s) => s.c === cmpCode);
  const dx = useMemo(() => diagnose(store), [store, store.pg]);

  // 모바일(터치 화면)엔 :hover가 없어서 ⓘ 팝업이 안 열림 → 탭으로 열고 닫기(같은 곳 다시 탭하면 닫힘, 다른 곳 탭하면 자동으로 닫힘)
  const handleTapPopup = (e) => {
    const trigger = e.target.closest(".has-gaptip, .has-mtip, .has-staffpop2, .has-wagepop");
    document.querySelectorAll(".tap-open").forEach((el) => { if (el !== trigger) el.classList.remove("tap-open"); });
    if (trigger) trigger.classList.toggle("tap-open");
  };

  return (
    <div onClick={handleTapPopup}>
      <div className="toolbar">
        <input placeholder="매장 검색" value={storeQuery} onChange={(e) => setStoreQuery(e.target.value)} style={{ width: 110 }} />
        <select value={store.c} onChange={(e) => { setCode(e.target.value); setStoreQuery(""); }}>
          {orderStores(STORES).filter((s) => s.c === store.c || !storeQuery || (s.n || "").includes(storeQuery) || s.c.toLowerCase().includes(storeQuery.toLowerCase())).map((s) => <option key={s.c} value={s.c}>{typeof s.pg === "number" ? "" : `[${s.pg}] `}{s.n}</option>)}
        </select>
        <span className="pill">{store.cell}</span><span className="pill">점장 {store.mg}</span>
        <span className="pill">{typeof store.pg === "number" ? "일반매장(매출 유사 비교)" : store.pg}</span>
        {/* 설명서(GUIDE_HTML)가 그룹번호 방식(유사매출대 전환 이전) 기준 내용이라 화면에서 숨김 — 내용 갱신 후 아래 주석만 풀면 재노출됨
        <button className="ghost guide-btn" onClick={downloadGuide} title="각 칸의 수치가 어떤 데이터에서 나오는지 정리한 설명서">↓ 설명서 다운로드</button>
        */}
        <button className="ghost guide-btn no-capture" onClick={captureDetail} disabled={capturingDetail} title="지금 보이는 매장진단 화면 전체를 이미지로 저장합니다">{capturingDetail ? "캡처 중…" : "📸 화면 캡처"}</button>
      </div>
      <div ref={captureRef}>

      <div className="statgrid">
        {(() => {
          const cnt = store._helperCnt || 0;
          const netWon = store._helperNetWon || 0;
          const deltaPp = store.s ? -netWon / store.s : 0; // 이 매장에 실제로 반영된 보정의 순효과(%p) — +면 인건비 늘어남, -면 줄어듦
          const before = store.lt - deltaPp; // 보정이 없었다면 어땠을지(역산)
          return (
            <div className={"stat " + (store.gap > 0 ? "over" : "under") + (cnt ? " has-gaptip" : "")}>
              <div className="stat-label">주간 인건비율{cnt ? <span className="th-info">ⓘ</span> : null}</div>
              <div className={"stat-value " + (store.gap > 0 ? "c-over" : "c-under")}>{fmtPct(store.lt)}</div>
              {cnt > 0 && (
                <div className="gaptip">
                  <div className="gaptip-title">직접 보정 반영</div>
                  <div className="gaptip-desc">인건비 직접 보정 {cnt}건 반영 시</div>
                  <div className="gaptip-line"><span>보정 전</span><b>{fmtPct(before)}</b></div>
                  <div className="gaptip-line"><span>보정 효과</span><b className={deltaPp > 0 ? "c-over" : "c-under"}>{fmtPp(deltaPp)}</b></div>
                  <div className="gaptip-line total"><span>보정 반영 누적(현재)</span><b>{fmtPct(store.lt)}</b></div>
                </div>
              )}
            </div>
          );
        })()}
        <div className={"stat has-gaptip " + (dx.excluded ? "" : store.gap > 0 ? "over" : "under")}>
          <div className="stat-label">유사매출대 대비 gap <span className="th-info">ⓘ</span></div>
          <div className={"stat-value " + (dx.excluded ? "" : store.gap > 0 ? "c-over" : "c-under")}>{dx.excluded ? "제외" : fmtPp(store.gap)}</div>
          <div className="stat-sub">{typeof store.pg === "number" ? "매출 유사 매장과 비교" : store.pg === "199매장" ? "199매장끼리 비교" : String(store.pg)}</div>
          {!dx.excluded && dx.peers.length > 0 && (
            <div className="peerlist-inline">유사매출대: {[...dx.peers].sort((a, b) => (b.s ?? -1) - (a.s ?? -1)).map((p) => p.n).join(", ")}</div>
          )}
          {dx.excluded ? (
            <div className="gaptip">
              <div className="gaptip-title">{store.pg === "신규매장" ? "신규 오픈 매장" : store.pg === "노출제외" ? "노출제외 매장" : "비교 대상 없음"}</div>
              {store.newUntil ? (
                <div className="gaptip-desc">오픈 3개월 이내라 유사매출대 비교에서 제외됩니다.<br /><b>{store.newUntil}까지 신규</b> · 이후 자동으로 일반 매장으로 전환되어 비교에 포함됩니다.</div>
              ) : (
                <div className="gaptip-desc">{dx.reason}</div>
              )}
            </div>
          ) : (
            <div className="gaptip">
              <div className="gaptip-title">어디서 차이가 나는가</div>
              <div className="gaptip-desc">매출 유사 매장(±10%, 위·아래 최대 3개씩) · {dx.peers.length}개 매장 평균 대비 구성별 차이</div>
              {[{ label: "정직원", d: dx.diffs.ls }, { label: "메이트", d: dx.diffs.lm }, { label: "가산", d: dx.diffs.lo }].map((x) => (
                <div key={x.label} className="gaptip-row"><span className="gaptip-label">{x.label}</span><GapBar gap={x.d} max={0.012} width={"100%"} /><span className={"gapnum " + (x.d > 0 ? "c-over" : "c-under")}>{fmtPp(x.d)}</span></div>
              ))}
            </div>
          )}
        </div>
        <Stat label="이번 주 추가근무(메이트)" value={store.mateExtraH != null ? fmtNum(store.mateExtraH, 1) : "–"} sub={`정상 ${fmtNum(store.mateNormH, 0)}h 중 추가근무 ${fmtNum(store.mateExtraH, 0)}h`} tone={store.mateExtraH > 0 ? "over" : undefined} />
        <div className={"stat rating " + (store.rt >= 4.5 ? "good" : store.rt < 4 ? "bad" : "")}>
          <div className="stat-label">고객평점 <span className="rating-hint">인건비와 함께 보기</span></div>
          <div className="stat-value">{store.rt ?? "–"}<span className="rating-star">★</span></div>
          <div className="stat-sub">{store.rt >= 4.5 ? "품질 양호 — 무리한 감축 주의" : store.rt < 4 ? "평점 낮음 — 인력 감축 신중히" : "평균권"}</div>
        </div>
      </div>

      <h3 className="sectionhead">시간대별 인원 세팅 <span>출퇴근기록부 실측 · 평일/주말 · 종합/홀/주방</span></h3>
      <div className="daytoggle">
        <button className={dayType === "평일" ? "on" : ""} onClick={() => setDayType("평일")}>평일</button>
        <button className={dayType === "주말" ? "on" : ""} onClick={() => setDayType("주말")}>주말</button>
        <span className="toggle-sep" />
        {["종합", "홀", "주방"].map((p) => <button key={p} className={slotPart === p ? "on" : ""} onClick={() => setSlotPart(p)}>{p}</button>)}
      </div>
      <div className="slotwrap">
        {(() => {
          const wk = dayType === "평일" ? "wd" : "we";
          const hk = store.slotHK ? store.slotHK[wk] : null;
          const real = hk ? hk[slotPart] : (dayType === "평일" ? store.slotWD : store.slotWE);
          const slotPeers = peersOf(store);
          const phk = slotPeers.length ? (() => {
            const withSelf = [store, ...slotPeers];
            const out = { wd: {}, we: {} };
            ["wd", "we"].forEach((k) => { out[k] = {}; ["종합", "홀", "주방"].forEach((part) => { out[k][part] = [0, 1, 2, 3, 4].map((idx) => avg(withSelf.map((p) => ({ v: (p.slotHK && p.slotHK[k] && p.slotHK[k][part]) ? p.slotHK[k][part][idx] : null })), "v")); }); });
            return out;
          })() : null;
          const peerArr = phk ? phk[wk][slotPart] : null;
          if (!real) return <div className="notice">이 매장은 출퇴근 실측 데이터가 없습니다.</div>;
          const mx = Math.max(...real, ...(peerArr || [0]), 1);
          return store.slots.map((sl, i) => {
            const ourSet = real[i];
            const peer = peerArr ? peerArr[i] : null;
            const over = peer != null && ourSet > peer + 0.1;
            const vs = peer != null ? ourSet - peer : null;
            return (
              <div key={sl.k} className="slot">
                <div className="slot-head"><span className="slot-name">{sl.k}</span><span className="slot-time">{sl.t}</span></div>
                <div className="slot-bars">
                  <div className="slot-row"><span className="slot-tag">우리</span><span className="bartrack"><span className={"bar " + (over ? "over" : "ok")} style={{ width: (ourSet / mx) * 100 + "%" }} /></span><b>{fmtNum(ourSet, 1)}명</b></div>
                  {peer != null && <div className="slot-row"><span className="slot-tag muted">유사매출대</span><span className="bartrack"><span className="bar peer" style={{ width: (peer / mx) * 100 + "%" }} /></span><span className="muted">{fmtNum(peer, 1)}명</span></div>}
                </div>
                {vs != null && <div className={"slot-foot " + (vs > 0 ? "c-over" : "c-under")}>유사매출대 대비 {fmtSign(vs)}명</div>}
              </div>
            );
          });
        })()}
      </div>

      {!dx.excluded && (
        <>
          <h3 className="sectionhead">원인 진단 <span>유사매출대 대비 · 시간·계약 중심 (액션은 매장 상황에 맞게 판단)</span></h3>
          <div className="tips">
            {dx.tips.map((t, i) => (
              <div key={i} className={"tip level-" + t.level}>
                <span className="tip-tag">{t.tag}</span>
                <div className="tip-main">
                  <div className="tip-head">{t.head}</div>
                  {t.facts && <ul className="tip-facts">{t.facts.map((f, j) => <li key={j} dangerouslySetInnerHTML={{ __html: hiTime(f) }} />)}</ul>}
                  {t.alts && t.alts.length > 0 && (
                    <div className="tip-alts">
                      <div className="tip-alts-label">참고 제안</div>
                      <ul>{t.alts.map((a, j) => <li key={j}>{a}</li>)}</ul>
                    </div>
                  )}
                </div>
              </div>
            ))}
          </div>
        </>
      )}

      <h3 className="sectionhead">평일 · 주말 매출 <span>일평균 실측 · 유사매출대 비교</span></h3>
      <div className="wewrap">
        {(() => {
          const wd = store.wdAvg || 0, we = store.weAvg || 0;
          const wePeers = peersOf(store);
          const pw = wePeers.length ? { wd: avg([store, ...wePeers].map((p) => ({ v: p.wdAvg })), "v") || 0, we: avg([store, ...wePeers].map((p) => ({ v: p.weAvg })), "v") || 0 } : null;
          const pwd = pw ? pw.wd : 0, pwe = pw ? pw.we : 0;
          const mx = Math.max(wd, we, pwd, pwe, 1);
          return (
            <>
              <div className="we-row"><span className="we-tag">평일</span><span className="bartrack big"><span className="bar ok" style={{ width: (wd / mx) * 100 + "%" }} /></span><b>{fmtWon(wd)}</b></div>
              {pw && <div className="we-row peer"><span className="we-tag muted">평일(유사매출대)</span><span className="bartrack big"><span className="bar peer" style={{ width: (pwd / mx) * 100 + "%" }} /></span><span className="muted">{fmtWon(pwd)}</span></div>}
              <div className="we-row"><span className="we-tag">주말</span><span className="bartrack big"><span className="bar over" style={{ width: (we / mx) * 100 + "%" }} /></span><b>{fmtWon(we)}</b></div>
              {pw && <div className="we-row peer"><span className="we-tag muted">주말(유사매출대)</span><span className="bartrack big"><span className="bar peer" style={{ width: (pwe / mx) * 100 + "%" }} /></span><span className="muted">{fmtWon(pwe)}</span></div>}
            </>
          );
        })()}
      </div>

      <h3 className="sectionhead">세부 항목 <span>종합판 기준 구성 내역</span></h3>
      <div className="detailgrid">
        <div className="dbox wide-dbox">
          <div className="dbox-head">
            <h4>인건비율 구성 (홀 / 주방 메이트)</h4>
            <button className="ghost tiny report-btn" onClick={() => setShowReport(true)}>영업지표 한판조회 →</button>
          </div>
          {(() => {
            const md = store.mateDetail || { 홀: {}, 주방: {} };
            const sum = (k) => (md.홀[k] || 0) + (md.주방[k] || 0);
            const baseAmt = { 정상: sum("정상"), 주휴: sum("주휴"), 연차: sum("연차"), 휴일: sum("휴일"), 휴업: sum("휴업") };
            const otAmt = { 야간: sum("야간"), 추가: sum("추가") };
            const realAmt = { ...baseAmt, ...otAmt };
            const fmtA = (o) => Object.entries(o).filter(([, v]) => v > 0).map(([k, v]) => `${k} ${fmtNum(v)}원`).join(" · ") || "0원";
            return (
              <div className="hktable-wrap">
              <table className="hktable">
                <thead><tr><th>구성</th><th className="num">홀 메이트</th><th className="num">주방 메이트</th><th className="num">합계</th></tr></thead>
                <tbody>
                  <tr className="has-staffpop2">
                    <td>정직원 <span className="th-info">ⓘ</span>
                      <div className="staffpop2">
                        <div className="staffpop-title">정직원 {store.staff.length}명 · 백판 월임금 ÷ 4.345</div>
                        <div className="chips">
                          {(() => { const order = ["선임점장", "점장", "GM", "GMIT", "매니저", "MIT", "캡틴", "CIT", "헤드", "HIT"]; const cnt = {}; store.staff.forEach((p) => (cnt[p.pos] = (cnt[p.pos] || 0) + 1)); const keys = Object.keys(cnt).sort((a, b) => (order.indexOf(a) + 99 * (order.indexOf(a) < 0)) - (order.indexOf(b) + 99 * (order.indexOf(b) < 0))); return keys.map((k) => <span key={k} className="chip"><b>{k}</b>{cnt[k]}</span>); })()}
                        </div>
                      </div>
                    </td>
                    <td className="num dim">–</td><td className="num dim">–</td><td className="num">{fmtPct(store.ls)}</td>
                  </tr>
                  <tr className="has-mtip">
                    <td>메이트 <span className="dimtxt">(총급여·가산 포함)</span> <span className="th-info">ⓘ</span>
                      <div className="mtip"><div className="mtip-t">메이트 = 정상+주휴+연차+휴업+야간+추가+휴일 (실지급 총액)</div><div className="mtip-b">{fmtA(realAmt)}</div></div>
                    </td>
                    <td className="num">{fmtPct(store.lmHallBase)}</td><td className="num">{fmtPct(store.lmKitBase)}</td><td className="num">{fmtPct(store.lm)}</td>
                  </tr>
                  <tr className="subline has-mtip">
                    <td><span className="dimtxt">└ 이 중 가산 (야간+추가+휴일)</span> <span className="th-info">ⓘ</span>
                      <div className="mtip"><div className="mtip-t">가산은 위 메이트 총급여에 포함됨 (참고용 분해)</div><div className="mtip-b">{fmtA(otAmt)}</div></div>
                    </td>
                    <td className="num dim">{fmtPct(store.loHall)}</td><td className="num dim">{fmtPct(store.loKit)}</td><td className="num dim">{fmtPct(store.lo)}</td>
                  </tr>
                  <tr className="sumrow"><td>총합 인건비율</td><td className="num dim">–</td><td className="num dim">–</td><td className={"num strong " + (store.gap > 0 ? "c-over" : "c-under")}>{fmtPct(store.lt)}</td></tr>
                </tbody>
              </table>
              </div>
            );
          })()}
          {(() => {
            const md = store.mateDetail || { 홀: {}, 주방: {} };
            const H = md.홀 || {}, K = md.주방 || {};
            const s = store.s || 1;
            const fmtMin = (m) => (m == null ? "–" : fmtNum(m) + "분");
            const toMin = (h) => (h == null ? null : Math.round(h * 60));
            // 정상근무 시간: 급여파일의 '정상시간' 컬럼은 비어있는 경우가 많아 대신 출퇴근기록부 실사용시간(useHall/useKit)에서
            // 가산시간(야간+추가+휴일)을 뺀 값으로 계산 — 실측 출퇴근 데이터 기준이라 더 신뢰도가 높음
            const otHallH = (H.야간시간 || 0) + (H.추가시간 || 0) + (H.휴일시간 || 0);
            const otKitH = (K.야간시간 || 0) + (K.추가시간 || 0) + (K.휴일시간 || 0);
            const normHallH = store.useHall != null ? Math.max(0, store.useHall - otHallH) : (H.정상시간 || null);
            const normKitH = store.useKit != null ? Math.max(0, store.useKit - otKitH) : (K.정상시간 || null);
            const ROWS = [
              { k: "정상근무", grp: "기본", hm: toMin(normHallH), km: toMin(normKitH), hw: H.정상, kw: K.정상 },
              { k: "주휴수당", grp: "기본", hm: null, km: null, hw: H.주휴, kw: K.주휴 },
              { k: "연차", grp: "기본", hm: null, km: null, hw: H.연차, kw: K.연차 },
              { k: "휴업근무", grp: "기본", hm: null, km: null, hw: H.휴업, kw: K.휴업 },
              { k: "야간근무", grp: "가산", hm: toMin(H.야간시간), km: toMin(K.야간시간), hw: H.야간, kw: K.야간 },
              { k: "추가근무", grp: "가산", hm: toMin(H.추가시간), km: toMin(K.추가시간), hw: H.추가, kw: K.추가 },
              { k: "휴일근무", grp: "가산", hm: toMin(H.휴일시간), km: toMin(K.휴일시간), hw: H.휴일, kw: K.휴일 },
            ];
            const byTime = allowView === "시간";
            return (
              <div className="allowbox">
                <div className="allow-head">
                  <h4 className="allow-h">메이트 수당 상세 <span>홀/주방 · {WEEK_RANGE}</span></h4>
                  <select className="viewsel" value={allowView} onChange={(e) => setAllowView(e.target.value)}><option value="금액">금액</option><option value="시간">시간</option></select>
                </div>
                <div className="hktable-wrap">
                <table className="hktable allowtable">
                  <thead><tr><th>수당</th><th className="num">홀 {byTime ? "시간" : "금액"}</th><th className="num">주방 {byTime ? "시간" : "금액"}</th><th className="num">{byTime ? "합계 시간" : "비율(합)"}</th></tr></thead>
                  <tbody>
                    {ROWS.map((r) => {
                      const tot = (r.hw || 0) + (r.kw || 0);
                      const totMin = (r.hm || 0) + (r.km || 0);
                      return (
                        <tr key={r.k} className={r.grp === "가산" ? "ot-row" : ""}>
                          <td>{r.k}{r.grp === "가산" && <span className="ot-badge">가산</span>}</td>
                          {byTime ? <><td className="num">{fmtMin(r.hm)}</td><td className="num">{fmtMin(r.km)}</td><td className="num strong">{r.hm == null && r.km == null ? "–" : fmtMin(totMin)}</td></>
                                  : <><td className="num">{fmtNum(r.hw || 0)}</td><td className="num">{fmtNum(r.kw || 0)}</td><td className="num strong">{fmtPct(tot / s)}</td></>}
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
                </div>
              </div>
            );
          })()}
        </div>
        <div className="dbox">
          <h4>실사용시간</h4>
          <dl>
            <div className="sumline"><dt>사용 합계</dt><dd>{fmtNum((store.useStaff || 0) + (store.mateNormH || 0), 0)}h</dd></div>
            <div><dt>정직원</dt><dd>{fmtNum(store.useStaff, 0)}h</dd></div>
            <div><dt>메이트</dt><dd>{fmtNum(store.mateNormH, 0)}h</dd></div>
            <div><dt>(메이트)계약시간</dt><dd>{store.ct4 != null ? fmtNum(store.ct4, 0) + "h" : "– (신규매장, 비교 불가)"}</dd></div>
            <div><dt>추가근무(메이트)</dt><dd className={store.mateExtraH > 0 ? "c-over" : "c-under"}>{store.mateExtraH != null ? fmtNum(store.mateExtraH, 1) + "h" : "–"}</dd></div>
          </dl>
        </div>
        <div className="dbox">
          <h4>메이트 인원 구성</h4>
          <dl>
            <div><dt>풀타임</dt><dd>{fmtNum(store.ft)}명</dd></div>
            <div><dt>중간</dt><dd>{fmtNum(store.md)}명</dd></div>
            <div><dt>초단기</dt><dd>{fmtNum(store.us)}명</dd></div>
            <div className="sumline"><dt>초단기 비율</dt><dd>{fmtPct(store.ur, 0)}</dd></div>
            {store.wageBreak ? (
              <div className="has-wagepop"><dt>평균 시급 <span className="th-info">ⓘ</span>
                <div className="wagepop">
                  <div className="wagepop-title">시급 구성 <span>근로시간당 · 메이트급여 기준</span></div>
                  {[{ k: "base", l: "기본시급(정상급여)" }, { k: "weekly", l: "주휴수당" }, { k: "extra", l: "추가근무수당" }, { k: "night", l: "야간수당" }, { k: "etc", l: "기타(휴일·휴업 등)" }].map((x) => (
                    store.wageBreak[x.k] > 0 ? <div key={x.k} className="wagepop-row"><span>{x.l}</span><b>{fmtNum(store.wageBreak[x.k])}원</b></div> : null
                  ))}
                  <div className="wagepop-row total"><span>평균 시급</span><b>{fmtNum(store.wageBreak.total)}원</b></div>
                  <div className="wagepop-note">연차수당은 이번 급여주기 전원 0원이라 제외됩니다.</div>
                </div>
              </dt><dd>{fmtNum(store.wageBreak.total)}원</dd></div>
            ) : (
              <div><dt>평균 시급</dt><dd>{store.realWage ? fmtNum(store.realWage) + "원" : "–"}</dd></div>
            )}
          </dl>
        </div>
        <div className="dbox">
          <h4>메이트 계약현황</h4>
          <dl>
            <div><dt>홀</dt><dd>{fmtNum(store.ctHall)}h</dd></div>
            <div><dt>주방</dt><dd>{fmtNum(store.ctKit)}h</dd></div>
            <div className="sumline"><dt>합계 · 홀비율</dt><dd>{fmtNum(store.ctSum)}h · {fmtPct(store.ctHallR, 0)}</dd></div>
          </dl>
        </div>
      </div>

      {dx.excluded ? <div className="notice">{dx.reason}</div> : (
        <>
          <h3 className="sectionhead">1:1 매장 비교 <span>비슷한 매장과 나란히 · 유사매출대 평균은 보조지표</span></h3>
          <div className="cmp-toolbar">
            <span className="cmp-label">비교 대상</span>
            <select value={cmpCode || ""} onChange={(e) => setCmpCode(e.target.value)}>
              {orderStores(peersOf(store)).map((s) => <option key={s.c} value={s.c}>{s.n} · 주 {fmtWon(s.s)}</option>)}
            </select>
            <span className="cmp-label" style={{ marginLeft: "12px" }}>메이트</span>
            <select value={cmpView} onChange={(e) => setCmpView(e.target.value)}><option value="율">율</option><option value="시간">시간</option></select>
          </div>
          <div className="tablewrap">
            <table className="compare cmp1v1">
              <thead><tr><th>지표</th><th className="num me">{store.n}<span className="coltag me">내 매장</span></th><th className="num">{cmp ? cmp.n : "–"}<span className="coltag">비교 대상</span></th><th className="num diff-col">차이<span className="coltag">내−비교</span></th><th className="num soft">유사매출대 평균<span className="coltag soft">보조</span></th></tr></thead>
              <tbody>
                {(() => {
                  const pAvg = dx.excluded ? {} : dx.pAvg;
                  const md = (s) => s && s.mateDetail ? s.mateDetail : { 홀: {}, 주방: {} };
                  const hallReal = (s) => s ? (s.lmHallBase || 0) + (s.loHall || 0) : null;
                  const kitReal = (s) => s ? (s.lmKitBase || 0) + (s.loKit || 0) : null;
                  const mateMin = (s, part) => { const d = md(s)[part] || {}; return (d.정상분 || 0) + (d.야간분 || 0) + (d.추가분 || 0) + (d.휴일분 || 0) + (d.휴업분 || 0) + (d.연차분 || 0); };
                  const pAvgFn = (fn) => { if (dx.excluded) return null; const v = dx.peers.map(fn).filter((x) => x != null); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
                  const byTime = cmpView === "시간";
                  const mateRows = byTime
                    ? [{ l: "메이트 시간 (홀)", rv: mateMin(store, "홀"), cv: cmp && mateMin(cmp, "홀"), pv: pAvgFn((s) => mateMin(s, "홀")), unit: "min", higherBad: true },
                       { l: "메이트 시간 (주방)", rv: mateMin(store, "주방"), cv: cmp && mateMin(cmp, "주방"), pv: pAvgFn((s) => mateMin(s, "주방")), unit: "min", higherBad: true }]
                    : [{ l: "메이트율 (홀)", rv: hallReal(store), cv: cmp && hallReal(cmp), pv: pAvgFn(hallReal), unit: "pct", higherBad: true },
                       { l: "메이트율 (주방)", rv: kitReal(store), cv: cmp && kitReal(cmp), pv: pAvgFn(kitReal), unit: "pct", higherBad: true }];
                  const R = [
                    { l: "인건비율 총합", rv: store.lt, cv: cmp && cmp.lt, pv: pAvg.lt, unit: "pct", higherBad: true },
                    { l: "정직원율", rv: store.ls, cv: cmp && cmp.ls, pv: pAvg.ls, unit: "pct", higherBad: true },
                    ...mateRows,
                    { l: "가산율", rv: store.lo, cv: cmp && cmp.lo, pv: pAvg.lo, unit: "pct", higherBad: true },
                    { l: "초단기 비율", rv: store.ur, cv: cmp && cmp.ur, pv: pAvg.ur, unit: "rate0", higherBad: false },
                    { l: "평균 시급", rv: store.realWage, cv: cmp && cmp.realWage, pv: pAvgFn((s) => s.realWage), unit: "won", higherBad: true },
                    { l: "고객평점 ★", rv: store.rt, cv: cmp && cmp.rt, pv: pAvgFn((s) => s.rt), unit: "pt", higherBad: false, rating: true },
                  ];
                  const fmtV = (v, u) => v == null ? "–" : u === "pct" ? fmtPct(v) : u === "rate0" ? fmtPct(v, 0) : u === "min" ? fmtNum(v) + "분" : u === "won" ? fmtNum(v) + "원" : u === "pt" ? v.toFixed(1) : fmtNum(v);
                  const fmtD = (d, u) => u === "pct" || u === "rate0" ? fmtPp(d) : u === "min" ? fmtSign(Math.round(d)) + "분" : u === "won" ? fmtSign(Math.round(d)) + "원" : u === "pt" ? (d > 0 ? "+" : "") + d.toFixed(1) : fmtSign(d);
                  return R.map((r) => {
                    const hasD = r.rv != null && r.cv != null && r.cv !== false;
                    const diff = hasD ? r.rv - r.cv : null;
                    const worse = diff != null && (r.higherBad ? diff > 0 : diff < 0);
                    const near = diff != null && Math.abs(diff) < (r.unit === "pct" ? 0.0005 : r.unit === "min" ? 30 : r.unit === "won" ? 100 : r.unit === "pt" ? 0.05 : 0.005);
                    return (
                      <tr key={r.l} className={r.rating ? "rating-row" : ""}>
                        <td>{r.l}</td>
                        <td className="num me strong">{fmtV(r.rv, r.unit)}</td>
                        <td className="num">{fmtV(r.cv === false ? null : r.cv, r.unit)}</td>
                        <td className={"num diff-col " + (diff == null || near ? "" : worse ? "c-over" : "c-under")}>{diff == null ? "–" : near ? "≈" : fmtD(diff, r.unit)}</td>
                        <td className="num soft">{fmtV(r.pv, r.unit)}</td>
                      </tr>
                    );
                  });
                })()}
              </tbody>
            </table>
          </div>
        </>
      )}

      <h3 className="sectionhead">주차별 추이 · 유사매출대 비교 <span>지표를 바꿔가며 확인 — 인건비율·가산율·메이트 평균시급·고객평점</span></h3>
      <TrendChart store={store} />
      </div>
      {showReport && <ReportModal store={store} onClose={() => setShowReport(false)} />}
    </div>
  );
}

/* ───────── 영업지표 한판조회 (피어그룹 비교) ───────── */
function ReportModal({ store, onClose }) {
  const pgKey = (store.pg !== "신규매장" && store.pg !== "노출제외" && store.s) ? store.c : null; // A타입: 그룹번호 대신 매장코드로 키(199매장은 199매장끼리 비교되므로 여기서 제외 안 함)
  const peers = pgKey ? orderStores(peersOf(store).concat([store])) : [store];
  const pgLabel = typeof store.pg === "number" ? "매출 유사 매장(±10%)" : String(store.pg);
  const COLS = [
    { h: "매장", get: (s) => <span className={s.c === store.c ? "me-row" : ""}>{s.n}{s.c === store.c && <span className="me-dot">우리</span>}</span>, cls: "stickc" },
    { h: "주간매출", get: (s) => fmtWon(s.s), cls: "num" },
    { h: "인건비율", get: (s) => fmtPct(s.lt), cls: "num strong", tone: (s) => s.gap > 0 ? "over" : "under" },
    { h: "전주→금주", get: (s) => { const t = s.trend && s.trend.lt; const prev = t && t[t.length - 2], cur = t && t[t.length - 1]; return prev != null && cur != null ? <span className={cur > prev ? "c-over" : "c-under"}>{fmtPct(prev, 1)}→{fmtPct(cur, 1)}</span> : "–"; }, cls: "num small" },
    { h: "정직원율", get: (s) => fmtPct(s.ls), cls: "num" },
    { h: "메이트 합", get: (s) => fmtPct((s.lm || 0) + (s.lo || 0), 2), cls: "num" },
    { h: "메이트율(홀)", get: (s) => fmtPct((s.lmHallBase || 0) + (s.loHall || 0), 2), cls: "num" },
    { h: "메이트율(주방)", get: (s) => fmtPct((s.lmKitBase || 0) + (s.loKit || 0), 2), cls: "num" },
    { h: "메이트 합(원)", get: (s) => fmtNum((s._payMate || 0) + (s._payOT || 0)), cls: "num refstart" },
    { h: "메이트 홀(원)", get: (s) => fmtNum(s.mHallReal || 0), cls: "num" },
    { h: "메이트 주방(원)", get: (s) => fmtNum(s.mKitReal || 0), cls: "num" },
    { h: "주휴수당", get: (s) => fmtNum(((s.mateDetail || {}).홀 || {}).주휴 + ((s.mateDetail || {}).주방 || {}).주휴 || 0), cls: "num" },
    { h: "연차수당", get: (s) => fmtNum(((s.mateDetail || {}).홀 || {}).연차 + ((s.mateDetail || {}).주방 || {}).연차 || 0), cls: "num" },
    { h: "기타수당(야간+추가+휴일+휴업)", get: (s) => fmtNum(s._payOT || 0), cls: "num" },
  ];
  return (
    <div className="modal-overlay" onClick={onClose}>
      <div className="modal-panel" onClick={(e) => e.stopPropagation()}>
        <div className="modal-head">
          <div><h3>영업지표 한판조회 <span className="modal-sub">{pgLabel} · {peers.length}개 매장 · {CUR_WEEK_LABEL || WEEK_RANGE}</span></h3></div>
          <button className="modal-x" onClick={onClose}>✕</button>
        </div>
        <div className="modal-body">
          <table className="reporttable">
            <thead><tr>{COLS.map((c) => <th key={c.h} className={c.cls}>{c.h}</th>)}</tr></thead>
            <tbody>
              {peers.map((s) => (
                <tr key={s.c} className={s.c === store.c ? "me-tr" : ""}>
                  {COLS.map((c) => <td key={c.h} className={(c.cls || "") + (c.tone && c.tone(s) ? " c-" + c.tone(s) : "")}>{c.get(s)}</td>)}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ───────── 3. 인건비 직접 보정 ───────── */
const KINDS = ["헬퍼", "HIT", "연차", "교육", "기타"];
// 원본 헬퍼 시트 305건 — 현재 지표에 이미 반영된 상태(기본 반영)
let HELPERS = [];
function buildHelpers() {
  HELPERS = (DATA.helpers || []).map((e, i) => ({
    id: "base-" + i, kind: e.kind, from: e.frm, to: e.to, pos: e.pos, days: e.days,
    half: e.half, name: e.name, memo: e.memo, amt: (e.staffAmt || 0) + (e.mateAmt || 0),
    baseline: true,
  }));
}
let CUR_WEEK_LO = null, CUR_WEEK_HI = null; // 이번 주 월~일 ISO 날짜(헬퍼보정 주차매칭용) — 라벨("8월 4주차")이 아니라 실제 날짜로 비교해야 정확함
function initData(json) {
  DATA = json;
  STORES = DATA.stores || [];
  WAGES = DATA.wages || [];
  if (json._label) CUR_WEEK_LABEL = json._label;
  // 날짜범위: _weekRange 우선, 없으면 라벨의 (...) 추출
  if (json._weekRange) WEEK_RANGE = json._weekRange;
  else if (json._label) { const m = String(json._label).match(/\(([^)]+)\)/); if (m) WEEK_RANGE = m[1]; }
  CUR_WEEK_LO = json._weekLo || null; CUR_WEEK_HI = json._weekHi || null;
  if (Array.isArray(json._weekLabels) && json._weekLabels.length) WEEK_LABELS = json._weekLabels; // 없으면(구버전 스냅샷) 기존 라벨 유지
  buildHelpers();
}
// 점장명/셀 정보(store_directory) 반영 — 매장코드로 매칭, 매주 올리는 파이프라인과 무관하게 독립적으로 갱신됨
// entries: [{code, mg, cell}] · 매칭 안 되는 매장코드는 조용히 무시(경고는 호출한 쪽에서 modal로 표시)
function applyStoreDirectory(entries) {
  const norm = (v) => String(v || "").trim().toUpperCase();
  let matched = 0; const unmatched = [];
  (entries || []).forEach((e) => {
    const key = norm(e.code);
    const s = STORES.find((x) => norm(x.c) === key);
    if (s) { if (e.mg) s.mg = e.mg; if (e.cell) s.cell = e.cell; matched++; }
    else unmatched.push(e.code);
  });
  return { matched, unmatched };
}
// ===== 헬퍼(인건비 직접 보정) 반영 — "매번 처음부터 다시 계산" 방식 =====
// _payStaffBase/_payMateBase(파일 기준 순수값, pipeBuild가 계산·저장. 절대 안 건드림) 위에,
// "지금 이 순간 승인(반영)된 헬퍼 중 이번 주(월~일)와 겹치는 것들의 효과"를 매번 새로 계산해서 대입한다.
// +=/-= 로 누적하지 않고 매번 base에서 다시 시작하므로, 몇 번을 호출하든(새로고침·재승인·재계산 몇 번이든) 항상 같은 결과가 나온다.
function recomputeHelperEffects(liveHelpers, curLo, curHi) {
  const staffDelta = {}, mateDelta = {}, netWon = {}, cnt = {};
  (liveHelpers || []).forEach((h) => {
    const segs = weekDaySegments(h.start, h.end); // 저장된 weeks 칼럼 안 믿고, 원본 날짜에서 매번 다시 계산(포맷/의미 불일치 문제 원천 차단)
    const totalDays = segs.reduce((a, w) => a + (w.days || 0), 0);
    const wk = curLo && curHi ? segs.find((w) => w.lo === curLo && w.hi === curHi) : null;
    if (!totalDays || !wk) return; // 이번 주와 안 겹치는 헬퍼는 이번 주 계산에서 완전히 제외
    const amt = h.amt * ((wk.days || 0) / totalDays);
    const bucket = (h.pos === "메이트" || h.kind === "퇴사연차제거" || h.kind === "교육" || h.kind === "연차") ? mateDelta : staffDelta;
    if (h.store_from) { bucket[h.store_from] = (bucket[h.store_from] || 0) - amt; netWon[h.store_from] = (netWon[h.store_from] || 0) + amt; cnt[h.store_from] = (cnt[h.store_from] || 0) + 1; }
    if (h.store_to) { bucket[h.store_to] = (bucket[h.store_to] || 0) + amt; netWon[h.store_to] = (netWon[h.store_to] || 0) - amt; cnt[h.store_to] = (cnt[h.store_to] || 0) + 1; }
  });
  STORES.forEach((s) => {
    s._payStaff = (s._payStaffBase ?? s._payStaff ?? 0) + (staffDelta[s.c] || 0);
    s._payMate = (s._payMateBase ?? s._payMate ?? 0) + (mateDelta[s.c] || 0);
    s._helperNetWon = netWon[s.c] || 0;
    s._helperCnt = cnt[s.c] || 0;
  });
  recomputeMetrics();
}
// 위 재계산에 필요한 "지금 승인(반영)된 헬퍼" 목록을 DB에서 가져와서 바로 재계산까지 실행
async function refreshHelperEffects() {
  try {
    const { data } = await supabase.from("helpers").select("store_from,store_to,amt,status,kind,pos,start_date,end_date").eq("status", "반영");
    const mapped = (data || []).map((h) => ({ ...h, start: h.start_date, end: h.end_date }));
    recomputeHelperEffects(mapped, CUR_WEEK_LO, CUR_WEEK_HI);
  } catch (e) { console.warn("헬퍼 효과 재계산 실패", e); }
}

// 날짜 범위를 주(월~일) 단위로 쪼개기 — 주차별 트래킹용
function splitByWeek(startStr, endStr) {
  if (!startStr || !endStr) return [];
  const s = new Date(startStr + "T00:00:00"), e = new Date(endStr + "T00:00:00");
  if (isNaN(s) || isNaN(e) || e < s) return [];
  const monday = (d) => { const x = new Date(d); const wd = (x.getDay() + 6) % 7; x.setDate(x.getDate() - wd); return x; };
  const fmt = (d) => `${d.getMonth() + 1}/${d.getDate()}`;
  const out = []; let cur = new Date(s);
  while (cur <= e) {
    const wkStart = monday(cur);
    const wkEnd = new Date(wkStart); wkEnd.setDate(wkStart.getDate() + 6);
    const segStart = cur < wkStart ? wkStart : cur;
    const segEnd = e < wkEnd ? e : wkEnd;
    const days = Math.round((segEnd - segStart) / 86400000) + 1;
    out.push({ label: `${fmt(wkStart)}~${fmt(wkEnd)}`, days, from: fmt(segStart), to: fmt(segEnd) });
    cur = new Date(wkEnd); cur.setDate(cur.getDate() + 1);
  }
  return out;
}
// 매칭 전용: 원본 시작~종료일(ISO)에서 "그 주(월~일)의 정확한 ISO 경계"와 그 주에 걸친 일수를 다시 계산.
// splitByWeek()의 from/to는 "8/17"처럼 연도 없는 화면표시용 포맷+실제 겹치는 날짜라 이번 주(CUR_WEEK_LO/HI, ISO)와
// 매칭이 안 됨 — 저장된 weeks 칼럼에 의존하지 않고 원본 날짜에서 매번 다시 계산하면 예전에 저장된 데이터도 그대로 맞음.
function weekDaySegments(startStr, endStr) {
  if (!startStr) return [];
  const s = new Date(startStr + "T00:00:00"), e = new Date((endStr || startStr) + "T00:00:00");
  if (isNaN(s) || isNaN(e) || e < s) return [];
  const monday = (d) => { const x = new Date(d); const wd = (x.getDay() + 6) % 7; x.setDate(x.getDate() - wd); return x; };
  const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  const out = []; let cur = new Date(s);
  while (cur <= e) {
    const wkStart = monday(cur);
    const wkEnd = new Date(wkStart); wkEnd.setDate(wkStart.getDate() + 6);
    const segStart = cur < wkStart ? wkStart : cur;
    const segEnd = e < wkEnd ? e : wkEnd;
    const days = Math.round((segEnd - segStart) / 86400000) + 1;
    out.push({ lo: iso(wkStart), hi: iso(wkEnd), days });
    cur = new Date(wkEnd); cur.setDate(cur.getDate() + 1);
  }
  return out;
}

function Helper({ refresh, myCode }) {
  const isAdmin = !myCode; // 매장 연결 없는 계정 = 관리자
  const [, localBump] = React.useReducer((x) => x + 1, 0);
  const [f, setF] = useState({ kind: "헬퍼", emp: "정직원", from: myCode || STORES[0].c, to: STORES[1].c, pos: "점장", wageH: "", hours: "", start: "", end: "", half: 1, name: "", memo: "" });
  const [fq, setFq] = useState(""); const [fk, setFk] = useState("전체"); const [fw, setFw] = useState("전체");
  const [fs, setFs] = useState("전체"); // 상태 필터: 전체/미확정/등록실패/반영됨
  const [editingId, setEditingId] = useState(null); // 수정 중인 기존 항목의 dbId(있으면 저장 시 기존 행을 지우고 새로 등록)
  const [selected, setSelected] = useState({}); // 일괄확정 체크된 항목: dbId → true
  const [bulkApproving, setBulkApproving] = useState(false);
  const [dbRows, setDbRows] = useState([]);   // DB에 저장된 보정
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const isMate = f.emp === "메이트";
  const wage = WAGES.find((w) => w.pos === f.pos);
  const weeks = splitByWeek(f.start, f.end);
  const days = weeks.reduce((a, w) => a + w.days, 0);
  const mateAmt = Math.round((Number(f.wageH) || 0) * (Number(f.hours) || 0) * f.half);
  const amt = isMate ? mateAmt : (wage ? Math.round(wage.daily * days * f.half) : 0);
  const validEntry = isMate ? (Number(f.wageH) > 0 && Number(f.hours) > 0) : (!!wage && days > 0);
  const fromStore = STORES.find((s) => s.c === f.from);
  const set = (k) => (e) => setF({ ...f, [k]: k === "half" ? Number(e.target.value) : e.target.value });

  // DB → 화면용 행 매핑
  const mapRow = (r) => ({
    id: "db-" + r.id, dbId: r.id, kind: r.kind, from: r.store_from, to: r.store_to,
    pos: r.pos, days: r.days, half: r.half, name: r.person, memo: r.memo,
    start: r.start_date, end: r.end_date, weeks: r.weeks || [], amt: r.amt,
    status: r.status, needAdmin: r.need_admin, baseline: false, createdAt: r.created_at,
    by: r.created_by_email || null,
  });
  async function loadDb() {
    setLoading(true);
    const { data, error } = await supabase.from("helpers").select("*").order("created_at", { ascending: false });
    if (!error && data) setDbRows(data.map(mapRow));
    setLoading(false);
  }
  React.useEffect(() => { loadDb(); }, []);
  // helpers 테이블에 삽입/수정 — created_by_email 컬럼이 아직 없으면(마이그레이션 전) 그 필드만 빼고 재시도
  const insertHelperRows = async (rowsToInsert) => {
    let { error } = await supabase.from("helpers").insert(rowsToInsert);
    if (error && /created_by_email/i.test(error.message)) {
      const stripped = rowsToInsert.map(({ created_by_email, ...rest }) => rest);
      ({ error } = await supabase.from("helpers").insert(stripped));
    }
    return { error };
  };
  const updateHelperRow = async (id, row) => {
    let { error } = await supabase.from("helpers").update(row).eq("id", id);
    if (error && /created_by_email/i.test(error.message)) {
      const { created_by_email, ...stripped } = row;
      ({ error } = await supabase.from("helpers").update(stripped).eq("id", id));
    }
    return { error };
  };

  // 일괄 업로드 (본사/셀 관리자만)
  const [bulkMsg, setBulkMsg] = useState("");
  const canBulk = AUTH.role === "hq" || AUTH.role === "cell";
  // 날짜·숫자 값을 형태와 무관하게 정규화 — 엑셀에서 열었다 저장하면 "2026-07-05" 텍스트가
  // 실제 날짜셀(Date 객체)이나 엑셀 시리얼번호로, 금액이 "150,000" 같은 천단위 구분기호 포함 텍스트로 바뀌는 경우가 흔함.
  const normDate = (v) => {
    if (v == null || v === "") return null;
    if (v instanceof Date) { const y = v.getFullYear(), m = String(v.getMonth() + 1).padStart(2, "0"), d = String(v.getDate()).padStart(2, "0"); return `${y}-${m}-${d}`; }
    if (typeof v === "number") {
      try { const d = XLSX.SSF.parse_date_code(v); if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`; } catch (e) { /* fallthrough */ }
      const dt = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      return isNaN(dt) ? null : dt.toISOString().slice(0, 10);
    }
    const s = String(v).trim();
    const m = s.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return `${m[1]}-${String(+m[2]).padStart(2, "0")}-${String(+m[3]).padStart(2, "0")}`;
    return s || null;
  };
  const normNum = (v) => { if (v == null || v === "") return 0; if (typeof v === "number") return v; return Number(String(v).replace(/[,원\s]/g, "")) || 0; };
  const onBulkUpload = async (file) => {
    if (!file) return;
    setBulkMsg("파일 읽는 중…");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
      if (!rows.length) { setBulkMsg("빈 파일입니다."); return; }
      // 헤더 이름으로 컬럼 위치를 찾음(고정 위치 아님) — "⬇ 보정내역 다운로드" 파일을 그대로 다시 올려도,
      // "일괄 보정 양식"(출발매장코드 등) 파일을 올려도 둘 다 인식되도록 두 양식의 헤더명을 모두 지원.
      const header = (rows[0] || []).map((h) => String(h || "").trim());
      const find = (...names) => { for (const n of names) { const i = header.indexOf(n); if (i >= 0) return i; } return -1; };
      const ci = {
        kind: find("구분"),
        from: find("출발매장코드", "출발매장"),
        to: find("도착매장코드(헬퍼만)", "도착매장"),
        emp: find("직원구분"), // "일괄 보정 양식"에만 있는 컬럼
        pos: find("직급(정직원)", "직급"),
        name: find("인원명"),
        start: find("시작일(YYYY-MM-DD)", "시작일"),
        end: find("종료일(YYYY-MM-DD)", "종료일"),
        half: find("반영비율(1 또는 0.5)"), // "일괄 보정 양식"에만 있는 컬럼
        amt: find("금액(메이트/퇴사연차)", "금액"),
        memo: find("메모"),
        status: find("상태"), // "보정내역 다운로드" 파일에만 있는 컬럼 — 있으면 이미 등록된 기존 기록이라는 뜻
      };
      if (ci.kind === -1 || ci.from === -1) { setBulkMsg("헤더를 인식하지 못했습니다. '구분', '출발매장코드(또는 출발매장)' 컬럼이 있는지 확인해 주세요."); return; }
      const get = (r, key) => (ci[key] >= 0 ? r[ci[key]] : null);
      // 중복 체크: 구분(연차/헬퍼 등)은 안 보고, 차감매장+이름이 같고 기간이 겹치기만 해도 중복으로 봄(정확히 같은 날짜일 필요 없음)
      const overlaps = (s1, e1, s2, e2) => s1 <= (e2 || s2) && s2 <= (e1 || s1);
      const dupGroupKey = (from, name) => `${from}|${String(name || "").trim()}`;
      const existingByGroup = {};
      dbRows.filter((e) => !e.baseline && e.name && e.start).forEach((e) => {
        const k = dupGroupKey(e.from, e.name);
        (existingByGroup[k] = existingByGroup[k] || []).push({ start: e.start, end: e.end || e.start });
      });
      const batchByGroup = {}; // 같은 파일 안에서의 중복도 같은 방식으로 체크
      const isDup = (from, name, start, end) => {
        const k = dupGroupKey(from, name);
        const list = [...(existingByGroup[k] || []), ...(batchByGroup[k] || [])];
        return list.some((r) => overlaps(start, end, r.start, r.end));
      };
      const { data: sess } = await supabase.auth.getUser();
      const out = []; let skippedDone = 0, failedCount = 0, dupCount = 0;
      for (let i = 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || get(r, "kind") == null) continue;
        const status = get(r, "status");
        if (status != null && status !== "승인대기") { skippedDone++; continue; } // 이미 반영/처리된 과거 기록은 재업로드해도 건너뜀(중복 등록 방지)
        const kind = String(get(r, "kind")).trim();
        const fromRaw = get(r, "from"), toRaw = get(r, "to");
        const from = resolveStoreCode(fromRaw);
        if (!from) { failedCount++; continue; } // 출발매장 자체를 못 찾으면 저장할 방법이 없어 건너뜀(DB 제약)
        const to = kind === "헬퍼" ? resolveStoreCode(toRaw) : null;
        const posRaw = get(r, "pos");
        // "직원구분" 컬럼이 없는 파일(보정내역 다운로드)은 직급값이 정직원 직급표에 있는지로 정직원/메이트를 판정
        const emp = ci.emp >= 0 ? String(get(r, "emp") || "").trim() : (WAGES.some((w) => w.pos === String(posRaw || "").trim()) ? "정직원" : "메이트");
        const name = get(r, "name"), start = normDate(get(r, "start")), end = normDate(get(r, "end")), memo = get(r, "memo");
        // 매장+이름이 같고 기간이 겹치면(매장에서 개별로도 신청하고 본사가 일괄로도 올리는 경우 등) 중복으로 보고 건너뜀
        if (name && start) {
          if (isDup(from, name, start, end)) { dupCount++; continue; }
          const k = dupGroupKey(from, name);
          (batchByGroup[k] = batchByGroup[k] || []).push({ start, end: end || start });
        }
        const halfRaw = ci.half >= 0 ? get(r, "half") : 1; // 다운로드 파일엔 반영비율 컬럼이 없으니 기본 100%
        const half = normNum(halfRaw) === 0.5 ? 0.5 : 1;
        const errReasons = [];
        if (kind === "헬퍼" && !to) errReasons.push(`도착매장 '${toRaw}' 인식 불가`);
        let amt = 0, days = 0, weeks = [];
        if (emp === "정직원") {
          const wage = WAGES.find((w) => w.pos === String(posRaw).trim());
          if (!wage) errReasons.push(`직급 '${posRaw}' 인식 불가`);
          weeks = splitByWeek(start, end); days = weeks.reduce((a, w) => a + w.days, 0);
          if (!days) errReasons.push(`시작/종료일 필요(읽은 값: ${get(r, "start")} ~ ${get(r, "end")})`);
          amt = wage && days ? Math.round(wage.daily * days * half) : 0;
        } else { // 메이트 / 퇴사연차제거 → 금액 직접
          amt = Math.round(normNum(get(r, "amt")));
          if (!amt) errReasons.push("금액 필요");
          weeks = start ? splitByWeek(start, end || start) : [];
        }
        const failed = errReasons.length > 0;
        if (failed) failedCount++;
        out.push({
          kind, store_from: from, store_to: to,
          pos: emp === "정직원" ? String(posRaw || "").trim() : "메이트",
          days, half, person: name || null,
          memo: failed ? `[${i + 1}행 업로드 오류: ${errReasons.join(", ")}]${memo ? " " + memo : ""}` : (memo || null),
          start_date: start || null, end_date: end || null, weeks, amt,
          status: failed ? "등록실패" : "승인대기", need_admin: !(kind === "헬퍼" && to),
          week: (weeks[0] && weeks[0].label) || CUR_WEEK_LABEL, created_by: sess?.user?.id || null, created_by_email: AUTH.email || null,
        });
      }
      if (!out.length) { setBulkMsg(`반영할 행 없음.${skippedDone ? ` (이미 처리된 ${skippedDone}건은 건너뜀)` : ""}${failedCount ? ` (출발매장 인식 불가 ${failedCount}건 건너뜀)` : ""}${dupCount ? ` (이미 등록된 것과 중복 ${dupCount}건 건너뜀)` : ""}`); return; }
      const { error } = await insertHelperRows(out);
      if (error) { setBulkMsg("저장 실패: " + error.message); return; }
      const okCount = out.length - out.filter((o) => o.status === "등록실패").length;
      setBulkMsg(`${okCount}건 등록 완료 (승인대기)${failedCount ? ` · ${failedCount}건은 정보 부족으로 "등록실패" 상태로 남겨뒀습니다(아래 목록에서 "수정"으로 고쳐서 다시 올리세요)` : ""}${skippedDone ? ` · 이미 처리된 ${skippedDone}건 건너뜀` : ""}${dupCount ? ` · 같은 매장·이름에 기간이 겹치는 기존 신청이 있어 ${dupCount}건 건너뜀` : ""}`);
      await loadDb(); localBump();
    } catch (e) { setBulkMsg("처리 실패: " + e.message); }
  };

  const add = async () => {
    const crossStore = f.kind === "헬퍼" && f.to;
    // 같은 매장·이름에 기간이 겹치는 신청이 이미 있으면(매장에서 개별 신청 + 본사 일괄 등록이 겹치는 경우 등) 확인 후 진행 — 구분(연차/헬퍼 등)은 안 따짐
    if (!editingId && f.name && f.start) {
      const overlaps = (s1, e1, s2, e2) => s1 <= (e2 || s2) && s2 <= (e1 || s1);
      const dup = dbRows.find((e) => !e.baseline && e.from === f.from && String(e.name || "").trim() === String(f.name).trim() && e.start && overlaps(f.start, f.end, e.start, e.end));
      if (dup && typeof window !== "undefined" && !window.confirm(`같은 매장·이름에 기간이 겹치는 신청이 이미 있습니다(${dup.start}~${dup.end || dup.start}, 상태: ${dup.status}). 그래도 등록할까요?`)) return;
    }
    setSaving(true);
    const { data: sess } = await supabase.auth.getUser();
    const mateNote = isMate ? `[시급 ${fmtNum(Number(f.wageH) || 0)}원 × ${f.hours}시간]${f.memo ? " " + f.memo : ""}` : f.memo;
    const row = {
      kind: f.kind, store_from: f.from, store_to: f.kind === "헬퍼" ? (f.to || null) : null,
      pos: isMate ? "메이트" : f.pos, days: isMate ? 0 : days, half: f.half, person: f.name, memo: mateNote,
      start_date: f.start || null, end_date: f.end || null, weeks, amt,
      status: "승인대기", need_admin: !crossStore, week: (weeks[0] && weeks[0].label) || null,
      created_by: sess?.user?.id || null, created_by_email: AUTH.email || null,
    };
    const { error } = editingId ? await updateHelperRow(editingId, row) : await insertHelperRows([row]);
    setSaving(false);
    if (error) { alert("저장 실패: " + error.message); return; }
    setF({ ...f, name: "", memo: "" });
    setEditingId(null);
    await loadDb();
  };
  // "수정" — 기존 항목을 위 입력 폼으로 불러옴(등록실패 건은 오류 안내 문구를 메모에서 떼어냄, 메이트 건은 메모에 적어둔 시급·시간을 다시 파싱)
  const startEdit = (e) => {
    setEditingId(e.dbId);
    const cleanMemo = (e.memo || "").replace(/^\[\d+행 업로드 오류:[^\]]*\]\s*/, "");
    const mateM = cleanMemo.match(/^\[시급 ([\d,]+)원 × ([\d.]+)시간\]\s*/);
    setF({
      kind: e.kind, emp: e.pos === "메이트" ? "메이트" : "정직원",
      from: e.from, to: e.to || "", pos: e.pos && e.pos !== "메이트" ? e.pos : "점장",
      wageH: mateM ? mateM[1].replace(/,/g, "") : "", hours: mateM ? mateM[2] : "",
      start: e.start || "", end: e.end || "", half: e.half || 1, name: e.name || "",
      memo: mateM ? cleanMemo.slice(mateM[0].length) : cleanMemo,
    });
    if (typeof window !== "undefined") window.scrollTo({ top: 0, behavior: "smooth" });
  };
  const cancelEdit = () => { setEditingId(null); setF({ ...f, name: "", memo: "" }); };
  const approve = async (entry) => {
    const { error } = await supabase.from("helpers").update({ status: "반영" }).eq("id", entry.dbId);
    if (error) { alert("승인 실패: " + error.message); return; }
    await refreshHelperEffects(); // "지금 승인된 헬퍼 전체" 기준으로 처음부터 다시 계산(누적 아님)
    await loadDb();
    refresh(); localBump();
  };
  const undo = async (entry) => {
    const { error } = await supabase.from("helpers").delete().eq("id", entry.dbId);
    if (error) { alert("취소 실패: " + error.message); return; }
    await refreshHelperEffects(); // 지운 뒤 "지금 승인된 헬퍼 전체" 기준으로 다시 계산 → 삭제된 건 자동으로 빠짐
    await loadDb();
    refresh(); localBump();
  };
  // 이 계정이 이 건을 확정/승인할 수 있는가
  // 승인 권한: 본사(hq)=전체, 셀관리자(cell)=본인 셀 매장만
  const storeCell = (code) => (STORES.find((s) => s.c === code) || {}).cell;
  const canApprove = (e) => {
    if (e.status !== "승인대기") return false;
    if (AUTH.role === "hq") return true;
    if (AUTH.role === "cell" && AUTH.cell) return storeCell(e.from) === AUTH.cell || (e.to && storeCell(e.to) === AUTH.cell);
    return false;
  };
  // 이 건을 지금 계정이 수정할 수 있는가 — 등록실패 건은 누구든(재제출 필요), 승인대기 건은 승인권자 본인이거나 본인이 신청한 건만
  const canEdit = (e) => e.status === "등록실패" || (e.status === "승인대기" && (canApprove(e) || e.by === AUTH.email));
  const bulkCancel = async () => {
    const ids = Object.keys(selected).filter((id) => selected[id]);
    if (!ids.length) return;
    if (typeof window !== "undefined" && !window.confirm(`선택한 ${ids.length}건을 취소(삭제)할까요? 되돌릴 수 없습니다.`)) return;
    setBulkApproving(true);
    const targets = dbRows.filter((e) => ids.includes(String(e.dbId)) && !e.baseline);
    for (const e of targets) { await supabase.from("helpers").delete().eq("id", e.dbId); }
    await refreshHelperEffects();
    setSelected({});
    setBulkApproving(false);
    await loadDb();
    refresh(); localBump();
  };
  // 체크한 항목 일괄확정 — 승인 권한 있는 건만 체크 대상이 되므로 개별 승인 버튼을 100번 누를 필요가 없음
  const toggleSelect = (id) => setSelected((s) => { const n = { ...s }; if (n[id]) delete n[id]; else n[id] = true; return n; });
  const bulkApprove = async () => {
    const ids = Object.keys(selected).filter((id) => selected[id]);
    if (!ids.length) return;
    setBulkApproving(true);
    const targets = dbRows.filter((e) => ids.includes(String(e.dbId)) && canApprove(e));
    for (const e of targets) {
      await supabase.from("helpers").update({ status: "반영" }).eq("id", e.dbId);
    }
    await refreshHelperEffects(); // 전부 업데이트 끝난 뒤 딱 한 번만 처음부터 다시 계산
    setSelected({});
    setBulkApproving(false);
    await loadDb();
    refresh(); localBump();
  };

  // 6월 등 스냅샷 내장(baseline)은 노출 안 함 — DB 저장분만 표시
  const ALL = dbRows;
  // 기존 등록내역 전체에서 같은 매장+이름인데 기간이 겹치는 쌍을 찾음(등록 시점과 무관하게, 이미 들어가있는 것들끼리도 검사)
  // dupKeepIds: 완전중복 그룹(같은 매장·이름·시작~종료일)에서 "남겨둘 1건"(가장 먼저 등록된 것) — 취소 후보 아님
  // dupExtraIds: 그 그룹에서 남겨둔 1건을 뺀 나머지(진짜 취소해야 할 여분) — 전체선택·일괄취소는 이것만 대상으로 함
  // dupOverlapIds: 기간이 겹치기만 하고 날짜 자체는 다른 건(예: 3개월 HIT 하나 + 그 기간 중 개별 연차 여러 건처럼
  // 정상적인 조합일 수 있음) — 사람이 직접 봐야 하므로 일괄취소 대상에서 계속 제외
  const [dupExtraIds, dupKeepIds, dupOverlapIds] = React.useMemo(() => {
    const overlaps = (s1, e1, s2, e2) => s1 <= (e2 || s2) && s2 <= (e1 || s1);
    const groups = {}, exactGroups = {};
    ALL.forEach((e) => {
      if (e.baseline || !e.name || !e.start) return;
      const k = `${e.from}|${String(e.name).trim()}`;
      (groups[k] = groups[k] || []).push(e);
      const ek = `${k}|${e.start}|${e.end || e.start}`;
      (exactGroups[ek] = exactGroups[ek] || []).push(e);
    });
    const extra = new Set(), keep = new Set(), overlap = new Set();
    Object.values(exactGroups).forEach((list) => {
      if (list.length < 2) return;
      const sorted = [...list].sort((a, b) => (a.dbId || 0) - (b.dbId || 0)); // 먼저 등록된 것을 원본으로 보고 유지
      keep.add(sorted[0].id);
      sorted.slice(1).forEach((e) => extra.add(e.id));
    });
    Object.values(groups).forEach((list) => {
      for (let i = 0; i < list.length; i++) {
        for (let j = i + 1; j < list.length; j++) {
          const a = list[i], b = list[j];
          if (!overlaps(a.start, a.end, b.start, b.end)) continue;
          const isExact = a.start === b.start && (a.end || a.start) === (b.end || b.start);
          if (!isExact) { overlap.add(a.id); overlap.add(b.id); }
        }
      }
    });
    return [extra, keep, overlap];
  }, [ALL]);
  const dupIds = useMemo(() => new Set([...dupExtraIds, ...dupKeepIds, ...dupOverlapIds]), [dupExtraIds, dupKeepIds, dupOverlapIds]); // 표시(필터·건수)는 전부 포함
  // 주차 드롭다운 목록 (DB분의 weeks에서 추출)
  const weekOptions = React.useMemo(() => {
    const set = new Set();
    dbRows.forEach((e) => (e.weeks || []).forEach((w) => set.add(w.label)));
    return [...set].sort();
  }, [dbRows]);
  const filtered = ALL.filter((e) => {
    if (fk !== "전체" && e.kind !== fk) return false;
    if (fw !== "전체" && !(e.weeks || []).some((w) => w.label === fw)) return false;
    if (fs === "미확정" && e.status !== "승인대기") return false;
    if (fs === "등록실패" && e.status !== "등록실패") return false;
    if (fs === "반영됨" && e.status !== "반영") return false;
    if (fs === "중복의심" && !dupIds.has(e.id)) return false;
    if (fq) {
      const fn = STORES.find((s) => s.c === e.from)?.n || e.from || "";
      const tn = STORES.find((s) => s.c === e.to)?.n || e.to || "";
      return (fn + tn + (e.name || "") + (e.memo || "") + (e.pos || "")).includes(fq);
    }
    return true;
  });
  const shown = fq || fk !== "전체" || fw !== "전체" || fs !== "전체" ? filtered : filtered.slice(0, 60);
  const selectableShown = shown.filter((e) => canApprove(e) || (fs === "중복의심" && dupExtraIds.has(e.id)));
  const allSelected = selectableShown.length > 0 && selectableShown.every((e) => selected[e.dbId]);
  const toggleSelectAll = () => {
    const n = { ...selected };
    if (allSelected) selectableShown.forEach((e) => delete n[e.dbId]);
    else selectableShown.forEach((e) => (n[e.dbId] = true));
    setSelected(n);
  };

  return (
    <div className="helperlayout">
      <div>
        <h3 className="sectionhead">보정 입력 <span>{editingId ? "기존 항목 수정 중 — 고친 뒤 저장하면 이 항목이 갱신됩니다" : "등록 즉시 종합 현황·매장 진단 지표에 반영"}</span></h3>
        <div className="form">
          <label>구분<select value={f.kind} onChange={set("kind")}>{KINDS.map((k) => <option key={k}>{k}</option>)}</select></label>
          <label>직원구분<select value={f.emp} onChange={set("emp")}><option value="정직원">정직원</option><option value="메이트">메이트</option></select></label>
          <label>차감 매장{myCode ? <span className="pill" style={{ marginLeft: 6 }}>{fromStore ? fromStore.n : myCode}</span> : <select value={f.from} onChange={set("from")}>{orderStores(STORES).map((s) => <option key={s.c} value={s.c}>{s.n}</option>)}</select>}</label>
          {f.kind === "헬퍼" && <label>추가 매장 (파견처)<select value={f.to || ""} onChange={set("to")}><option value="">없음 (파견처 없음)</option>{orderStores(STORES).map((s) => <option key={s.c} value={s.c}>{s.n}</option>)}</select></label>}
          {isMate ? (
            <div className="formrow">
              <label>시급(원)<input type="number" min="0" step="100" value={f.wageH} onChange={set("wageH")} placeholder="예: 11000" /></label>
              <label>근무시간(h)<input type="number" min="0" step="0.5" value={f.hours} onChange={set("hours")} placeholder="예: 24" /></label>
              <label>보정<select value={f.half} onChange={set("half")}><option value={1}>1.0</option><option value={0.5}>0.5</option></select></label>
            </div>
          ) : (
            <div className="formrow">
              <label>직급<select value={f.pos} onChange={set("pos")}>{WAGES.map((w) => <option key={w.pos}>{w.pos}</option>)}</select></label>
              <label>보정<select value={f.half} onChange={set("half")}><option value={1}>1.0</option><option value={0.5}>0.5</option></select></label>
            </div>
          )}
          <div className="formrow">
            <label>시작일<input type="date" value={f.start} onChange={set("start")} /></label>
            <label>종료일<input type="date" value={f.end} onChange={set("end")} min={f.start || undefined} /></label>
          </div>
          {!isMate && weeks.length > 0 && (
            <div className="weekbox">
              <div className="weekbox-h">주차별 분리 <span>총 {days}일 (참고용)</span></div>
              <table className="weektable"><tbody>
                {weeks.map((w, i) => <tr key={i}><td>{w.label}</td><td className="wt-sub">{w.from}~{w.to}</td><td className="num">{w.days}일</td></tr>)}
              </tbody></table>
            </div>
          )}
          <div className="formrow">
            <label>이름<input value={f.name} onChange={set("name")} placeholder="예: 이예슬" /></label>
            <label>내역<input value={f.memo} onChange={set("memo")} placeholder="예: 오픈 매장 헬퍼 파견" /></label>
          </div>
          <div className="amountline"><span>반영 대상</span><span className="calcnote">{isMate ? (f.wageH && f.hours ? `시급 ${fmtNum(Number(f.wageH))}원 · ${f.hours}시간 × ${f.half}` : "") : (wage && days ? `${wage.pos} · ${days}일 × ${f.half}${weeks.length > 1 ? ` · ${weeks.length}개 주차` : ""}` : "")}</span></div>
          <p className="hint mini approve-note">{f.kind === "헬퍼" && f.to ? "타 매장 파견은 파견처 승인 후 반영됩니다." : "파견처 없는 보정은 관리자 확정 후 반영됩니다."}</p>
          <div className="formrow">
            <button className="primary" onClick={add} disabled={!f.name || !validEntry || saving}>{saving ? "저장 중…" : editingId ? "수정 내용 저장" : "등록 → 승인 요청"}</button>
            {editingId && <button className="ghost" onClick={cancelEdit}>수정 취소</button>}
          </div>
        </div>
      </div>
      <aside className="sidecard">
        <h4>차감 매장 현황 (실시간)</h4>
        <div className="side-store">{fromStore.n}</div>
        <dl>
          <div><dt>주간 인건비율</dt><dd className={fromStore.gap > 0 ? "c-over" : "c-under"}>{fmtPct(fromStore.lt)}</dd></div>
          <div><dt>유사매출대 대비 gap</dt><dd className={fromStore.gap > 0 ? "c-over" : "c-under"}>{fmtPp(fromStore.gap)}</dd></div>
          <div><dt>추가근무(메이트, 시간)</dt><dd>{fmtNum(fromStore.mateExtraH, 1)}h</dd></div>
          <div><dt>초단기 비율</dt><dd>{fmtPct(fromStore.ur, 0)}</dd></div>
        </dl>
      </aside>

      <div className="fullrow">
        <h3 className="sectionhead">등록내역 <span>{loading ? "불러오는 중…" : `${filtered.length}건 · 클라우드 저장`}</span></h3>
        {bulkMsg && <div className="pipe-msg">{bulkMsg}</div>}
        <div className="toolbar">
          <button className="dl-btn" onClick={() => exportHelpers(dbRows)}>⬇ 보정내역 다운로드</button>
          <button className="dl-btn" onClick={exportHelperTemplate}>⬇ 일괄보정 양식</button>
          {canBulk && <label className="dl-btn up-btn">⬆ 일괄 업로드<input type="file" accept=".xlsx,.xls" style={{ display: "none" }} onChange={(e) => { onBulkUpload(e.target.files[0]); e.target.value = ""; }} /></label>}
          <select value={fk} onChange={(e) => setFk(e.target.value)}><option>전체</option>{KINDS.map((k) => <option key={k}>{k}</option>)}</select>
          <select value={fw} onChange={(e) => setFw(e.target.value)}><option value="전체">전체 주차</option>{weekOptions.map((w) => <option key={w} value={w}>{w}</option>)}</select>
          <select value={fs} onChange={(e) => setFs(e.target.value)}><option value="전체">전체 상태</option><option value="미확정">미확정만</option><option value="등록실패">등록실패만</option><option value="반영됨">반영됨만</option><option value="중복의심">⚠ 중복의심만</option></select>
          {dupExtraIds.size > 0 && <span className="pill warn" style={{ cursor: "pointer" }} onClick={() => setFs("중복의심")}>⚠ 중복 취소대상 {dupExtraIds.size}건(1건씩 유지)</span>}
          {dupOverlapIds.size > 0 && <span className="pill" style={{ cursor: "pointer", background: "#FFF4E5", color: "#B45309" }} onClick={() => setFs("중복의심")}>기간겹침(확인필요) {dupOverlapIds.size}건</span>}
          <input placeholder="매장명·이름·내역·직급 검색" value={fq} onChange={(e) => setFq(e.target.value)} />
          {Object.keys(selected).length > 0 && (
            fs === "중복의심"
              ? <button className="primary danger" onClick={bulkCancel} disabled={bulkApproving} style={{ marginLeft: "auto" }}>{bulkApproving ? "처리 중…" : `선택 ${Object.keys(selected).length}건 취소`}</button>
              : <button className="primary" onClick={bulkApprove} disabled={bulkApproving} style={{ marginLeft: "auto" }}>{bulkApproving ? "처리 중…" : `선택 ${Object.keys(selected).length}건 일괄확정`}</button>
          )}
        </div>
        <div className="tablewrap">
          <table>
            <thead><tr><th style={{ width: 28 }}><input type="checkbox" checked={allSelected} onChange={toggleSelectAll} title="승인 가능한 건 전체 선택" /></th><th>상태</th><th>구분</th><th>차감 매장</th><th>추가 매장</th><th>직급</th><th>이름</th><th>기간 · 주차</th><th className="num">일수</th><th>내역</th><th>신청자</th><th></th></tr></thead>
            <tbody>
              {shown.map((e) => (
                <tr key={e.id} className={e.status === "등록실패" ? "row-fail" : dupIds.has(e.id) ? "row-dupe" : ""}>
                  <td>{(canApprove(e) || (fs === "중복의심" && !e.baseline)) && <input type="checkbox" checked={!!selected[e.dbId]} onChange={() => toggleSelect(e.dbId)} />}</td>
                  <td>{e.baseline ? <span className="pill dim">기본 반영</span> : e.status === "등록실패" ? <span className="pill warn">등록실패</span> : e.status === "승인대기" ? <span className="pill wait">{e.needAdmin ? "관리자 확정 대기" : "파견처 승인 대기"}</span> : <span className="pill live">반영됨</span>}{dupExtraIds.has(e.id) && <span className="pill warn" title="같은 매장·이름·기간(시작~종료일까지 완전히 동일)이 또 있어서, 그중 이 건은 취소 후보로 잡혔습니다(가장 먼저 등록된 1건만 남김)" style={{ marginLeft: 4 }}>⚠ 중복(취소대상)</span>}{dupKeepIds.has(e.id) && <span className="pill" title="같은 매장·이름·기간의 중복 중, 가장 먼저 등록된 이 건은 유지 대상으로 남겨뒀습니다" style={{ marginLeft: 4, background: "#E8F5E9", color: "#2E7D32" }}>중복 중 유지</span>}{dupOverlapIds.has(e.id) && <span className="pill" title="같은 매장·이름에 기간이 겹치는 다른 신청이 있습니다(날짜는 다름 — 장기건+개별연차처럼 정상 조합일 수 있어 확인 필요)" style={{ marginLeft: 4, background: "#FFF4E5", color: "#B45309" }}>기간겹침(확인필요)</span>}</td>
                  <td><span className="pill">{e.kind}</span></td>
                  <td>{STORES.find((s) => s.c === e.from)?.n || e.from || "–"}</td>
                  <td>{e.to ? (STORES.find((s) => s.c === e.to)?.n || e.to) : "–"}</td>
                  <td>{e.pos}</td><td>{e.kind === "기타" ? <span className="blind">🔒 비공개</span> : e.name}</td>
                  <td className="wt-cell">{e.start && e.end ? <><span className="wt-range">{e.start}~{e.end}</span>{e.weeks && e.weeks.length > 0 && <span className="wt-weeks">{e.weeks.map((w) => `${w.label}(${w.days}일)`).join(", ")}</span>}</> : "–"}</td>
                  <td className="num">{e.days ? fmtNum(e.days, 0) : "–"}</td>
                  <td className="memo">{e.kind === "기타" ? <span className="blind">🔒 개인정보 보호 (본인만 열람)</span> : e.memo}</td>
                  <td className="dimtxt" style={{ fontSize: 12 }}>{e.by || "–"}</td>
                  <td style={{ display: "flex", gap: 4 }}>
                    {canApprove(e) && <button className="ghost tiny approve-btn" onClick={() => approve(e)}>{e.needAdmin ? "확정" : "승인"}</button>}
                    {canEdit(e) && <button className="ghost tiny" onClick={() => startEdit(e)}>수정</button>}
                    {!e.baseline && e.status !== "승인대기" && e.status !== "등록실패" && <button className="ghost tiny" onClick={() => undo(e)}>취소</button>}
                    {e.status === "등록실패" && <button className="ghost tiny" onClick={() => undo(e)}>삭제</button>}
                  </td>
                </tr>
              ))}
              {!shown.length && <tr><td colSpan={12} className="empty">검색 결과가 없습니다.</td></tr>}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

/* ───────── 4. 매장 정직원 ───────── */
const STAFF_KINDS = ["육아휴직", "리프레쉬 휴가", "출산휴가", "병가", "무급휴직", "파견(전출)", "기타"];
const POS_ORDER = ["선임점장", "점장", "GM", "GMIT", "매니저", "MIT", "캡틴", "CIT", "헤드", "HIT"];
const STAFF_SEED = [
  { store: "AL413", pos: "매니저", name: "김세연", kind: "육아휴직", from: "2026-06-01", to: "2027-05-31", memo: "1년 육아휴직, 인건비 제외" },
  { store: "AL132", pos: "헤드", name: "박도현", kind: "리프레쉬 휴가", from: "2026-06-16", to: "2026-06-20", memo: "5일 리프레쉬" },
];
function StaffTab({ code: initCode, onBack }) {
  const [code, setCode] = useState(initCode || STORES.find((s) => s.staff.length)?.c || STORES[0].c);
  const [notes, setNotes] = useState(STAFF_SEED);
  const [f, setF] = useState({ pos: "매니저", name: "", kind: "육아휴직", from: "", to: "", memo: "" });
  const store = STORES.find((s) => s.c === code);
  const roster = [...store.staff].sort((a, b) => {
    const ia = POS_ORDER.indexOf(a.pos), ib = POS_ORDER.indexOf(b.pos);
    return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib);
  });
  const totalPay = store.staff.reduce((a, p) => a + (p.pay || 0), 0);
  const storeNotes = notes.filter((n) => n.store === code);
  const set = (k) => (e) => setF({ ...f, [k]: e.target.value });
  const add = () => { setNotes([{ ...f, store: code }, ...notes]); setF({ ...f, name: "", from: "", to: "", memo: "" }); };

  return (
    <div>
      <div className="toolbar">
        <button className="ghost" onClick={onBack}>← 매장 진단</button>
        <select value={code} onChange={(e) => setCode(e.target.value)}>
          {[...STORES].sort((a, b) => a.n.localeCompare(b.n, "ko")).map((s) => <option key={s.c} value={s.c}>{s.n}</option>)}
        </select>
        <span className="pill">{store.cell}</span><span className="pill">점장 {store.mg}</span>
        <span className="count">정직원 {store.staff.length}명 · 메이트 제외</span>
      </div>

      <div className="statgrid">
        <Stat label="정직원 인원" value={store.staff.length + "명"} sub="백판 기준 · 메이트 제외" />
      </div>

      <div className="staffcols">
        <div>
          <h3 className="sectionhead">정직원 명부 <span>백판에서 불러옴 · {roster.length}명</span></h3>
          <div className="tablewrap">
            <table>
              <thead><tr><th>직위</th><th>특이사항</th></tr></thead>
              <tbody>
                {roster.map((p, i) => {
                  const note = storeNotes.find((n) => n.pos === p.pos);
                  return (
                    <tr key={i}>
                      <td><span className="chip"><b>{p.pos}</b></span></td>
                      <td>{note ? <span className="pill warn">{note.kind}</span> : <span className="muted">–</span>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        <div>
          <h3 className="sectionhead">특이사항 입력 <span>육아휴직·리프레쉬 등</span></h3>
          <div className="form">
            <div className="formrow two">
              <label>직위<select value={f.pos} onChange={set("pos")}>{POS_ORDER.map((p) => <option key={p}>{p}</option>)}</select></label>
              <label>이름<input value={f.name} onChange={set("name")} placeholder="예: 이예슬" /></label>
            </div>
            <label>구분<select value={f.kind} onChange={set("kind")}>{STAFF_KINDS.map((k) => <option key={k}>{k}</option>)}</select></label>
            <div className="formrow two">
              <label>시작일<input type="date" value={f.from} onChange={set("from")} /></label>
              <label>종료일<input type="date" value={f.to} onChange={set("to")} /></label>
            </div>
            <label>내역<input value={f.memo} onChange={set("memo")} placeholder="예: 1년 육아휴직, 인건비 제외" /></label>
            <button className="primary" onClick={add} disabled={!f.name}>등록</button>
          </div>
        </div>
      </div>

      <h3 className="sectionhead">등록된 특이사항 <span>{store.n} · {storeNotes.length}건</span></h3>
      <div className="tablewrap">
        <table>
          <thead><tr><th>구분</th><th>직위</th><th>이름</th><th>기간</th><th>내역</th></tr></thead>
          <tbody>
            {storeNotes.map((n, i) => (
              <tr key={i}>
                <td><span className="pill warn">{n.kind}</span></td>
                <td>{n.pos}</td><td>{n.name}</td>
                <td className="muted">{n.from || "–"} ~ {n.to || "–"}</td>
                <td className="memo">{n.memo}</td>
              </tr>
            ))}
            {!storeNotes.length && <tr><td colSpan={5} className="empty">등록된 특이사항이 없습니다.</td></tr>}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ───────── 5. 차주 시뮬레이션 ───────── */
const WK = 4.345; // 월→주 환산
const J_WEEKLY = 0.15; // 주휴 프리미엄 가정(초단기 비율 변화 효과 계수)
function SimTab() {
  const [code, setCode] = useState([...STORES].sort((a, b) => (b.s ?? -1) - (a.s ?? -1))[0].c);
  const [storeQuery, setStoreQuery] = useState(""); // 매장 검색(그룹이 없어져서 이름으로 빠르게 찾기용)
  const store = STORES.find((s) => s.c === code);
  const salesTrend = store.trend && store.trend.sales;
  const salesNow = Math.round((salesTrend && salesTrend[salesTrend.length - 1]) || (store.ps || store.s) / WK); // 금주 실매출(주간)
  const W0 = Math.round((store.ps || store.s) / WK / 10000) * 10000; // 차주 예상매출 기본값(주간·만원단위)

  const mateWageH = store.realWage || store.mateWage || 11000; // 메이트 평균시급(실지급 총액 기준)
  const baseMateH0 = Math.round(store.wkMate || (store.hm ? store.hm / WK : 0)); // 현재 주간 메이트 근무시간
  const totalMateCnt0 = Math.round((store.ft || 0) + (store.md || 0) + (store.us || 0)); // 메이트 총원
  const ultraCnt0 = Math.round(store.us || 0); // 현재 초단기 인원

  // 기준(현재) 주간 원 단위 — 금주 실매출 기준, ratio0 == lt
  const staff0 = store.ls * salesNow;
  const mate0 = store.lm * salesNow;
  const ot0 = store.lo * salesNow;
  const total0 = staff0 + mate0 + ot0;
  const ratio0 = total0 / salesNow;
  const otHours0 = mateWageH ? ot0 / (mateWageH * 1.5) : 0;

  // ── 레버: 실물 단위 ──
  const [sales, setSales] = useState(W0);            // 차주 예상매출(원)
  const [mateH, setMateH] = useState(baseMateH0);    // 차주 메이트 근무시간(h/주)
  const [ultraNum, setUltraNum] = useState(ultraCnt0);   // 초단기 인원(분자)
  const [mateDen, setMateDen] = useState(totalMateCnt0); // 메이트 총원(분모)
  const [dOtH, setdOtH] = useState(0);               // OT 시간 ±h(주간)
  const [holidayWon, setHolidayWon] = useState(0);   // 공휴일 예상 인건비(만원) — 버전업 예정, 결과 미반영
  const [covers, setCovers] = useState([]);          // 연차·교육 보장 항목 [{pos, days}]

  React.useEffect(() => {
    setSales(W0); setMateH(baseMateH0); setUltraNum(ultraCnt0); setMateDen(totalMateCnt0); setdOtH(0); setCovers([]);
  }, [code]);

  // 현재 초단기 비율 대비 조정 후 비율
  const ur0 = totalMateCnt0 ? ultraCnt0 / totalMateCnt0 : 0;
  const ur1 = mateDen ? ultraNum / mateDen : 0;
  const dUltraRatio = ur1 - ur0;

  // 차주 조정 적용
  const mateHmult = baseMateH0 ? mateH / baseMateH0 : 1;             // 시간 배수
  const mate1 = mate0 * mateHmult * (1 - J_WEEKLY * dUltraRatio);    // 시간·초단기 반영
  const otHours1 = otHours0 + dOtH;
  const ot1 = Math.max(0, otHours1) * mateWageH * 1.5;
  // 연차·교육 보장 공제 = Σ(직급 일급 × 일수)
  const coverCredit = covers.reduce((a, c) => {
    const w = WAGES.find((x) => x.pos === c.pos);
    return a + (w ? w.daily * (c.days || 0) : 0);
  }, 0);
  const staff1 = staff0; // 정직원 급여는 보장 유지(고정)
  const total1 = Math.max(0, staff1 + mate1 + ot1 - coverCredit);
  const ratio1 = sales > 0 ? total1 / sales : 0;

  const dRatio = ratio1 - ratio0;
  const peerAvg = store.gap != null ? store.lt - store.gap : null;
  const newGap = peerAvg != null ? ratio1 - peerAvg : null;

  // ── 홀/주방 분해 (현재/차주) ──
  const mateReal0 = mate0 + ot0;
  const hallShare = (store.lm + store.lo) > 0 ? ((store.lmHallBase || 0) + (store.loHall || 0)) / (store.lm + store.lo) : 0.3;
  const mateReal1 = mate1 + ot1;
  const hall0 = mateReal0 * hallShare, kit0 = mateReal0 * (1 - hallShare);
  const hall1 = mateReal1 * hallShare, kit1 = mateReal1 * (1 - hallShare);

  // ── 최적(피어 기반 가이드) — 로직은 추후 버전업 예정 ──
  const matePeers = peersOf(store);
  const pm = matePeers.length ? (() => { const withSelf = [store, ...matePeers]; return { lmHallBase: avg(withSelf.map((p) => ({ v: p.lmHallBase })), "v"), loHall: avg(withSelf.map((p) => ({ v: p.loHall })), "v"), lmKitBase: avg(withSelf.map((p) => ({ v: p.lmKitBase })), "v"), loKit: avg(withSelf.map((p) => ({ v: p.loKit })), "v") }; })() : null;
  const optHall = pm ? ((pm.lmHallBase || 0) + (pm.loHall || 0)) * sales : null;
  const optKit = pm ? ((pm.lmKitBase || 0) + (pm.loKit || 0)) * sales : null;
  const optStaff = staff0;
  const optTotal = pm ? optStaff + optHall + optKit : null;
  const optRatio = pm && sales > 0 ? optTotal / sales : null;

  // 목표: 피어 대비 gap ≤ +1%p
  const GOAL = 0.01;
  const goalMet = newGap != null && newGap <= GOAL;
  const toGoalWon = peerAvg != null ? total1 - (peerAvg + GOAL) * sales : null;
  const toGoalPp = newGap != null ? newGap - GOAL : null;

  const dx = diagnose(store);
  const topDriver = !dx.excluded ? [
    { key: "메이트", d: dx.diffs.lm, lever: "메이트 근무시간·초단기 인원" },
    { key: "정직원", d: dx.diffs.ls, lever: "연차·교육 보장 커버" },
    { key: "가산", d: dx.diffs.lo, lever: "가산 시간" },
  ].filter((x) => x.d > 0).sort((a, b) => b.d - a.d)[0] : null;

  const rows = [
    { k: "정직원 (보장·고정)", w0: staff0, w1: staff1, wo: optStaff },
    { k: "메이트 홀 (기본+가산)", w0: hall0, w1: hall1, wo: optHall },
    { k: "메이트 주방 (기본+가산)", w0: kit0, w1: kit1, wo: optKit },
    { k: "연차·교육 보장 공제", w0: 0, w1: -coverCredit, wo: 0 },
  ];

  return (
    <div>
      <div className="toolbar">
        <input placeholder="매장 검색" value={storeQuery} onChange={(e) => setStoreQuery(e.target.value)} style={{ width: 110 }} />
        <select value={code} onChange={(e) => { setCode(e.target.value); setStoreQuery(""); }}>
          {orderStores(STORES).filter((s) => s.c === code || !storeQuery || (s.n || "").includes(storeQuery) || s.c.toLowerCase().includes(storeQuery.toLowerCase())).map((s) => <option key={s.c} value={s.c}>{typeof s.pg === "number" ? "" : `[${s.pg}] `}{s.n}</option>)}
        </select>
        <span className="pill">{store.cell}</span><span className="pill">점장 {store.mg}</span>
        <span className="tip-pop-wrap">
          <button className="ghost tiny tip-pop-btn">매장진단 제안 보기</button>
          <div className="tip-pop">
            {(() => {
              const dx = diagnose(store);
              if (dx.excluded) return <div className="tip-pop-empty">{dx.reason}</div>;
              if (!dx.tips.length) return <div className="tip-pop-empty">특이 제안 없음</div>;
              return dx.tips.map((t, i) => (
                <div key={i} className="tip-pop-item">
                  <div className="tip-pop-head"><span className={"tip-tag lv-" + t.level}>{t.tag}</span>{t.head}</div>
                  {t.alts && <ul className="tip-pop-alts">{t.alts.map((a, j) => <li key={j}>{a}</li>)}</ul>}
                </div>
              ));
            })()}
          </div>
        </span>
        <button className="ghost" style={{ marginLeft: "auto" }} onClick={() => { setSales(W0); setMateH(baseMateH0); setUltraNum(ultraCnt0); setMateDen(totalMateCnt0); setdOtH(0); setCovers([]); }}>초기화</button>
      </div>

      <div className="simhead">
        <div className="simhead-box base">
          <div className="simhead-label">현재 인건비율</div>
          <div className="simhead-val">{fmtPct(ratio0)}</div>
          <div className="simhead-sub">금주 매출 {fmtWon(salesNow)} 기준</div>
        </div>
        <div className="simhead-arrow">→</div>
        <div className={"simhead-box proj " + (dRatio > 0 ? "up" : dRatio < 0 ? "down" : "")}>
          <div className="simhead-label">차주 예상 인건비율</div>
          <div className="simhead-val">{fmtPct(ratio1)}</div>
          <div className="simhead-sub">차주 예상매출 {fmtWon(sales)} 기준 · {dRatio === 0 ? "현재와 동일" : `현재 대비 ${fmtPp(dRatio)}`}</div>
        </div>
        <div className={"simhead-box gap " + (goalMet ? "goal-met" : "")}>
          <div className="simhead-label">유사매출대 대비 gap (예상)</div>
          <div className={"simhead-val " + (newGap > 0 ? "c-over" : "c-under")}>{fmtPp(newGap)}</div>
          <div className="simhead-sub">{goalMet ? "✓ 목표(+1%p 이하) 달성" : toGoalPp != null ? `목표까지 −${(toGoalPp * 100).toFixed(1)}%p` : `현재 ${fmtPp(store.gap)}`}</div>
        </div>
      </div>

      {topDriver && !goalMet && (
        <div className="drivertip">
          <span className="drivertip-tag">가장 큰 격차</span>
          <span><b>{topDriver.key}</b> 파트가 유사매출대 평균 대비 <b>{fmtPp(topDriver.d)}</b>로 가장 벌어져 있습니다. 목표(gap +1%p 이하) 도달을 위해 <b>{topDriver.lever}</b> 레버를 우선 조정해 보세요. {toGoalWon != null && toGoalWon > 0 ? `주간 인건비를 약 ${fmtWon(toGoalWon)} 더 줄이면 목표에 도달합니다.` : ""}</span>
        </div>
      )}

      <div className="simcols">
        <div>
          <h3 className="sectionhead">차주 조정 레버 <span>현재 수치 기준 · ± 조정</span></h3>
          <div className="levers">
            <div className="lever">
              <div className="lever-h"><span className="lever-label">차주 예상매출</span>
                <span className="lever-input"><input type="number" step="1" value={Math.round(sales / 10000)} onChange={(e) => setSales((Number(e.target.value) || 0) * 10000)} /><span className="unit">만원</span></span>
              </div>
              <input type="range" min={Math.round(W0 * 0.7)} max={Math.round(W0 * 1.3)} step={10000} value={Math.min(Math.max(sales, Math.round(W0 * 0.7)), Math.round(W0 * 1.3))} onChange={(e) => setSales(Number(e.target.value))} />
            </div>
            <div className="lever">
              <div className="lever-h"><span className="lever-label">메이트 근무시간</span>
                <span className="lever-input"><input type="number" step="1" value={mateH} onChange={(e) => setMateH(Math.max(0, Number(e.target.value) || 0))} /><span className="unit">h/주</span></span>
              </div>
              <input type="range" min={Math.round(baseMateH0 * 0.6)} max={Math.round(baseMateH0 * 1.4) || 100} step={1} value={Math.min(Math.max(mateH, Math.round(baseMateH0 * 0.6)), Math.round(baseMateH0 * 1.4) || 100)} onChange={(e) => setMateH(Number(e.target.value))} />
              <div className="lever-hint">현재 주 {fmtNum(baseMateH0, 0)}h · 실제 근무시간을 직접 입력</div>
              {mateH !== baseMateH0 && <div className="lever-real">➜ 현재 {fmtNum(baseMateH0, 0)}h → <b>{fmtNum(mateH, 0)}h</b> ({mateH > baseMateH0 ? "+" : ""}{fmtNum(mateH - baseMateH0, 0)}h {mateH > baseMateH0 ? "증가" : "단축"}){totalMateCnt0 ? ` · 1인당 약 ${fmtNum(Math.abs(mateH - baseMateH0) / totalMateCnt0, 1)}h/주` : ""}</div>}
            </div>
            <div className="lever">
              <div className="lever-h"><span className="lever-label">초단기 메이트 (분자 / 분모)</span>
                <span className="lever-input">
                  <input type="number" step="1" min="0" value={ultraNum} onChange={(e) => setUltraNum(Math.max(0, Number(e.target.value) || 0))} />
                  <span className="unit">/</span>
                  <input type="number" step="1" min="1" value={mateDen} onChange={(e) => setMateDen(Math.max(1, Number(e.target.value) || 1))} />
                  <span className="unit">명</span>
                </span>
              </div>
              <div className="lever-hint">현재 {ultraCnt0}명 / {totalMateCnt0}명 = {fmtPct(ur0, 0)} · 분자=초단기 인원, 분모=메이트 총원</div>
              {(ultraNum !== ultraCnt0 || mateDen !== totalMateCnt0) && <div className="lever-real">➜ 초단기 비율 {fmtPct(ur0, 0)} → <b>{fmtPct(ur1, 0)}</b> ({ultraNum}명 / {mateDen}명){ur1 < ur0 ? " · 주휴수당 부담 증가" : ur1 > ur0 ? " · 주휴수당 부담 감소" : ""}</div>}
            </div>
            <div className="lever">
              <div className="lever-h"><span className="lever-label">가산 시간(야간+추가+휴일)</span>
                <span className="lever-input"><input type="number" step="0.5" value={dOtH} onChange={(e) => setdOtH(Number(e.target.value) || 0)} /><span className="unit">h/주</span></span>
              </div>
              <input type="range" min={-Math.ceil(otHours0)} max={40} step={0.5} value={Math.min(Math.max(dOtH, -Math.ceil(otHours0)), 40)} onChange={(e) => setdOtH(Number(e.target.value))} />
              <div className="lever-hint">현재 주 {fmtNum(otHours0, 1)}h 가산 ({fmtPct(store.lo, 1)}) · 조정 후 {fmtNum(Math.max(0, otHours1), 1)}h</div>
              {dOtH !== 0 && <div className="lever-real">➜ 주간 OT(야간·추가) <b>{dOtH > 0 ? "+" : ""}{fmtNum(dOtH, 1)}시간</b> {dOtH > 0 ? "증가" : "감축"} → {fmtNum(otHours0, 1)}h → {fmtNum(Math.max(0, otHours1), 1)}h</div>}
            </div>
            <div className="lever holiday-lever">
              <div className="lever-h"><span className="lever-label">공휴일 주간 예상 인건비</span>
                <span className="ver-memo">⚙ 버전업 예정</span>
              </div>
              <span className="lever-input" style={{ marginTop: "6px" }}><input type="number" step="1" value={holidayWon} onChange={(e) => setHolidayWon(Math.max(0, Number(e.target.value) || 0))} /><span className="unit">만원</span></span>
              <div className="lever-hint">공휴일이 낀 주의 추가 인건비(휴일근로·가산 등)를 예상 반영하는 칸입니다. <b>계산 로직은 검토 중</b>이라 지금은 입력만 받고 결과에는 반영하지 않습니다.</div>
              <div className="lever-real dimtxt">※ 공휴일 인건비 산식 확정 후 최적/차주 인건비율에 반영 예정입니다.</div>
            </div>
            <div className="lever cover-lever">
              <div className="lever-h"><span className="lever-label">연차·교육 보장 커버</span>
                <span className="cover-total">{covers.length ? `${covers.length}건 반영` : "항목 추가"}</span>
              </div>
              <div className="lever-hint">직급·일수를 입력하면 통상임금 기준으로 자동 공제됩니다. 정직원 급여는 보장 유지, 매장 인건비에서 공제됩니다.</div>
              <div className="coverlist">
                {covers.map((c, i) => (
                  <div key={i} className="coverrow">
                    <select value={c.pos} onChange={(e) => setCovers(covers.map((x, j) => j === i ? { ...x, pos: e.target.value } : x))}>
                      {WAGES.map((w) => <option key={w.pos}>{w.pos}</option>)}
                    </select>
                    <input type="number" min="0" step="1" value={c.days} onChange={(e) => setCovers(covers.map((x, j) => j === i ? { ...x, days: Math.max(0, Number(e.target.value) || 0) } : x))} />
                    <span className="cover-unit">일</span>
                    <button className="ghost tiny" onClick={() => setCovers(covers.filter((_, j) => j !== i))}>삭제</button>
                  </div>
                ))}
                <button className="ghost cover-add" onClick={() => setCovers([...covers, { pos: "매니저", days: 1 }])}>+ 보장 항목 추가</button>
              </div>
            </div>
          </div>
        </div>

        <div>
          <h3 className="sectionhead">인건비 구성 변화 <span>주간 · 원 · 현재 / 차주 / 최적(유사매출대 기반)</span></h3>
          <div className="tablewrap">
            <table>
              <thead><tr><th>구성</th><th className="num">현재</th><th className="num">차주</th><th className="num opt-col">최적</th></tr></thead>
              <tbody>
                {rows.map((r) => {
                  const d = r.w1 - r.w0;
                  return (
                    <tr key={r.k}>
                      <td>{r.k}</td>
                      <td className="num">{Math.round(r.w0).toLocaleString()}</td>
                      <td className="num strong">{Math.round(r.w1).toLocaleString()}<span className={"cell-d " + (d > 0 ? "c-over" : d < 0 ? "c-under" : "")}>{d === 0 ? "" : (d > 0 ? " +" : " ") + Math.round(d).toLocaleString()}</span></td>
                      <td className="num opt-col">{r.wo == null ? "–" : Math.round(r.wo).toLocaleString()}</td>
                    </tr>
                  );
                })}
                <tr className="sumrow">
                  <td>총 인건비 (주간)</td>
                  <td className="num">{Math.round(total0).toLocaleString()}</td>
                  <td className="num strong">{Math.round(total1).toLocaleString()}</td>
                  <td className="num opt-col">{optTotal == null ? "–" : Math.round(optTotal).toLocaleString()}</td>
                </tr>
                <tr className="sumrow">
                  <td>인건비율</td>
                  <td className="num">{fmtPct(ratio0)}</td>
                  <td className={"num strong " + (dRatio > 0 ? "c-over" : "c-under")}>{fmtPct(ratio1)}</td>
                  <td className="num opt-col">{optRatio == null ? "–" : fmtPct(optRatio)}</td>
                </tr>
              </tbody>
            </table>
          </div>
          <div className="simbars">
            <div className="simbar-row"><span className="simbar-tag">현재</span><span className="bartrack big"><span className="bar ok" style={{ width: Math.min(ratio0 / 0.1, 1) * 100 + "%" }} /></span><b>{fmtPct(ratio0)}</b></div>
            <div className="simbar-row"><span className="simbar-tag">차주</span><span className="bartrack big"><span className={"bar " + (dRatio > 0 ? "over" : "ok")} style={{ width: Math.min(ratio1 / 0.1, 1) * 100 + "%" }} /></span><b>{fmtPct(ratio1)}</b></div>
            {optRatio != null && <div className="simbar-row"><span className="simbar-tag opt">최적</span><span className="bartrack big"><span className="bar opt" style={{ width: Math.min(optRatio / 0.1, 1) * 100 + "%" }} /></span><b>{fmtPct(optRatio)}</b></div>}
          </div>
          {optRatio != null && (
            <div className="optguide">
              <div className="optguide-h"><span className="optguide-tag">최적 가이드</span><span className="ver-memo">⚙ 인건비율 가이드 버전업 예정</span></div>
              <div className="optguide-b">
                같은 유사매출대의 <b>금주 메이트율</b>(홀 {fmtPct(pm.lmHallBase + pm.loHall, 1)} · 주방 {fmtPct(pm.lmKitBase + pm.loKit, 1)})을 차주 예상매출 {fmtWon(sales)}에 적용한 목표치입니다. 최적 인건비율 <b>{fmtPct(optRatio)}</b> · 우리 차주 예상 {fmtPct(ratio1)}과 차이 <b className={ratio1 > optRatio ? "c-over" : "c-under"}>{fmtPp(ratio1 - optRatio)}</b>.
                <br /><span className="dimtxt">※ 이 최적 로직(유사매출대 차주 예상매출 · 유사매출대 금주 메이트율 기반)은 추후 변경될 예정입니다.</span>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* ───────── 피어그룹분류 편집 (관리자) ───────── */
function PeerClassEditor({ refresh }) {
  const [, bump] = React.useReducer((x) => x + 1, 0);
  const [q, setQ] = useState("");
  const [gf, setGf] = useState("전체");
  const [draft, setDraft] = useState({});   // 변경 대기: code → 새 분류
  const [saving, setSaving] = useState(false);
  const [msg, setMsg] = useState("");
  const [dirBusy, setDirBusy] = useState(false);
  const [dirMsg, setDirMsg] = useState("");
  // 매장정보(분류+점장+셀) 통합 양식 업로드 — 직접 만든 고정 헤더라 항상 정확히 읽힘
  const onStoreInfoFile = async (file) => {
    if (!file) return;
    setDirBusy(true); setDirMsg("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
      const headerIdx = rows.findIndex((r) => Array.isArray(r) && r.some((c) => String(c || "").trim().startsWith("매장코드")));
      if (headerIdx === -1) throw new Error("이 양식이 아닌 것 같습니다. '⬇ 매장정보 다운로드'로 받은 파일을 그대로 사용해 주세요.");
      const header = rows[headerIdx].map((h) => String(h || "").trim());
      const idx = (prefix) => header.findIndex((h) => h.startsWith(prefix));
      const ci = { code: idx("매장코드"), pg: idx("분류"), mg: idx("점장명"), cell: idx("셀") };
      const norm = (v) => String(v || "").trim().toUpperCase();
      const entries = [];
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || !r[ci.code]) continue;
        const pgRaw = ci.pg >= 0 ? String(r[ci.pg] || "").trim() : "일반";
        const pg = (pgRaw === "신규매장" || pgRaw === "199매장" || pgRaw === "노출제외") ? pgRaw : 1; // 정해진 3개 외엔 전부 '일반' 처리
        entries.push({ code: String(r[ci.code]).trim(), pg, mg: ci.mg >= 0 && r[ci.mg] != null ? String(r[ci.mg]).trim() : null, cell: ci.cell >= 0 && r[ci.cell] != null ? String(r[ci.cell]).trim() : null });
      }
      if (!entries.length) throw new Error("인식된 매장 행이 없습니다.");
      let matched = 0; const unmatched = [];
      entries.forEach((e) => {
        const key = norm(e.code);
        const s = STORES.find((x) => norm(x.c) === key);
        if (s) { s.pg = e.pg; if (e.mg) s.mg = e.mg; if (e.cell) s.cell = e.cell; matched++; }
        else unmatched.push(e.code);
      });
      recomputeMetrics();
      const full = {}; STORES.forEach((s) => (full[s.c] = s.pg));
      const pgErr = (await supabase.from("peer_config").upsert({ week: DATA._week || "current", data: full, updated_at: new Date().toISOString() })).error;
      const dirPayload = entries.map((e) => ({ code: e.code, mg: e.mg, cell: e.cell, updated_at: new Date().toISOString() }));
      const dirErr = (await supabase.from("store_directory").upsert(dirPayload, { onConflict: "code" })).error;
      if (pgErr || dirErr) setDirMsg(`화면엔 반영됐지만 저장 중 일부 실패했습니다: ${[pgErr?.message, dirErr?.message].filter(Boolean).join(" / ")}`);
      else setDirMsg(`반영 완료 — ${matched}개 매장 매칭${unmatched.length ? `, ${unmatched.length}개 매장코드를 못 찾았습니다(${unmatched.slice(0, 5).join(", ")}${unmatched.length > 5 ? " 등" : ""})` : ""}`);
      setDraft({}); bump(); refresh && refresh();
    } catch (e) {
      setDirMsg("읽기 실패: " + e.message);
    }
    setDirBusy(false);
  };
  // 점장명/셀 정보 업로드(HR 제공 연락처 파일 그대로) — 매장코드로 매칭. 주차별 파이프라인과 무관, 인사이동 있을 때만 다시 올리면 됨
  const onDirFile = async (file) => {
    if (!file) return;
    setDirBusy(true); setDirMsg("");
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
      const headerIdx = rows.findIndex((r) => Array.isArray(r) && r.some((c) => String(c || "").trim() === "매장코드"));
      if (headerIdx === -1) throw new Error("헤더(매장코드 컬럼)를 찾지 못했습니다. 양식을 확인해 주세요.");
      const header = rows[headerIdx];
      const idx = (name) => header.findIndex((h) => String(h || "").trim() === name);
      const ci = { mg: idx("이름"), code: idx("매장코드"), cell: idx("선임셀") };
      const entries = [];
      for (let i = headerIdx + 1; i < rows.length; i++) {
        const r = rows[i]; if (!r || !r[ci.code]) continue;
        entries.push({ code: String(r[ci.code]).trim(), mg: ci.mg >= 0 && r[ci.mg] != null ? String(r[ci.mg]).trim() : null, cell: ci.cell >= 0 && r[ci.cell] != null ? String(r[ci.cell]).trim() : null });
      }
      if (!entries.length) throw new Error("인식된 매장 행이 없습니다.");
      const { matched, unmatched } = applyStoreDirectory(entries);
      recomputeMetrics();
      try {
        const payload = entries.map((e) => ({ code: e.code, mg: e.mg, cell: e.cell, updated_at: new Date().toISOString() }));
        const { error } = await supabase.from("store_directory").upsert(payload, { onConflict: "code" });
        if (error) throw error;
        setDirMsg(`반영 완료 — ${matched}개 매장 매칭${unmatched.length ? `, ${unmatched.length}개 매장코드를 못 찾았습니다(${unmatched.slice(0, 5).join(", ")}${unmatched.length > 5 ? " 등" : ""})` : ""}`);
      } catch (e) {
        setDirMsg(`화면엔 반영됐지만 저장은 실패했습니다(새로고침하면 사라짐): ${e.message}`);
      }
      bump(); refresh && refresh();
    } catch (e) {
      setDirMsg("읽기 실패: " + e.message);
    }
    setDirBusy(false);
  };
  // DATA.peerClass는 원본 시드 데이터에 박혀있던 정적 목록이라, 나중에 자동 추가된 신규매장이 절대 반영되지 않는 문제가 있었음
  // → 매번 현재 STORES에서 직접 만들어서, 새로 추가된 매장도 바로 이 목록에 뜨도록 수정
  const cls = STORES.map((s) => ({ c: s.c, sales: s.s }));
  const byCode = {}; STORES.forEach((s) => (byCode[s.c] = s));
  const label = (pg) => (typeof pg === "number" ? "일반" : pg);
  const rows = cls
    .map((r) => ({ ...r, store: byCode[r.c] }))
    .filter((r) => r.store)
    .filter((r) => (q ? (r.store.n || "").includes(q) || r.c.includes(q) : true))
    .filter((r) => (gf === "전체" ? true : label(curPg(r)) === gf))
    .sort((a, b) => (b.sales ?? -1) - (a.sales ?? -1))
    .map((r, i) => ({ ...r, rankNow: i + 1 }));   // 매출순 실시간 순위

  function curPg(r) { return draft[r.c] !== undefined ? draft[r.c] : r.store.pg; }
  const stage = (code, v) => {
    const val = (v === "신규매장" || v === "199매장" || v === "노출제외") ? v : 1; // "일반"은 그냥 숫자 1(비교 로직상 어떤 숫자든 동일 취급)
    setDraft((d) => ({ ...d, [code]: val }));
  };
  const dirtyCount = Object.keys(draft).length;
  const [rgFrom, setRgFrom] = useState(""); const [rgTo, setRgTo] = useState(""); const [rgTarget, setRgTarget] = useState("일반");

  const apply = async () => {
    if (!dirtyCount) return;
    setSaving(true); setMsg("");
    // 1) 로컬 반영 + 재계산
    Object.entries(draft).forEach(([code, pg]) => { const s = STORES.find((x) => x.c === code); if (s) s.pg = pg; });
    recomputeMetrics();
    // 2) DB 저장 (전 매장 현재 분류를 한 번에)
    const full = {}; STORES.forEach((s) => (full[s.c] = s.pg));
    try {
      const week = (DATA._week) || "current";
      const { error } = await supabase.from("peer_config").upsert({ week, data: full, updated_at: new Date().toISOString() });
      if (error) throw error;
      setMsg(`반영 완료 · ${dirtyCount}개 매장 변경됨`);
      setDraft({});
    } catch (e) { setMsg("저장 실패(로컬만 반영됨): " + e.message); }
    setSaving(false);
    bump(); refresh && refresh();
  };
  const reset = () => { setDraft({}); setMsg(""); bump(); };

  return (
    <div>
      <h3 className="sectionhead">매장 분류 <span>신규매장 · 199매장 · 노출제외 지정 · 수정 후 "반영 요청"을 눌러야 확정·저장됩니다</span></h3>
      <div className="pipe-msg" style={{ marginBottom: 12 }}>
        유사매출대 비교(매장진단의 gap·제안)는 이제 매출이 비슷한 매장을 자동으로 찾아서 비교합니다(위·아래 최대 3개씩, ±10% 이내). 그룹 번호를 따로 지정할 필요가 없어졌고, 여기서는 <b>비교 대상에서 빼야 할 매장(신규매장/199매장/노출제외)</b>만 지정합니다.
      </div>
      <div className="pipe-msg" style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <span><b>매장정보(분류·점장·셀) 일괄 수정</b> — 아래에서 현재 값을 다운로드해 엑셀로 고친 뒤 그대로 다시 올리면, 이 표의 "구간 일괄 지정"이나 개별 드롭다운 없이도 한 번에 반영됩니다. 이 양식은 직접 만든 고정 형식이라 항상 정확히 인식됩니다.</span>
        <button className="dl-btn" onClick={exportStoreInfoTemplate}>⬇ 매장정보 다운로드</button>
        <label className="dl-btn up-btn">{dirBusy ? "처리 중…" : "⬆ 매장정보 업로드"}<input type="file" accept=".xlsx,.xls" style={{ display: "none" }} disabled={dirBusy} onChange={(e) => { onStoreInfoFile(e.target.files[0]); e.target.value = ""; }} /></label>
      </div>
      <div className="pipe-msg" style={{ marginBottom: 12, display: "flex", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <span><b>(보조) 본사 점장 연락처 파일 업로드</b> — 위 양식이 아니라 HR에서 받는 원본 연락처 파일을 그대로 올리고 싶을 때만 쓰세요. 점장명·셀만 갱신되고 분류는 안 바뀝니다.</span>
        <label className="dl-btn up-btn">{dirBusy ? "처리 중…" : "⬆ 점장/셀 파일 업로드(HR 원본)"}<input type="file" accept=".xlsx,.xls" style={{ display: "none" }} disabled={dirBusy} onChange={(e) => { onDirFile(e.target.files[0]); e.target.value = ""; }} /></label>
        {dirMsg && <span>{dirMsg}</span>}
      </div>
      <div className="toolbar">
        <select value={gf} onChange={(e) => setGf(e.target.value)}>
          <option>전체</option><option value="일반">일반</option>
          <option value="신규매장">신규매장</option><option value="199매장">199매장</option><option value="노출제외">노출제외</option>
        </select>
        <input placeholder="매장 검색" value={q} onChange={(e) => setQ(e.target.value)} />
        <span className="count">{rows.length}개</span>
        <span style={{ marginLeft: "auto", display: "flex", gap: "8px", alignItems: "center" }}>
          {msg && <span className="peer-msg">{msg}</span>}
          {dirtyCount > 0 && <span className="dirty-badge">변경 {dirtyCount}건 대기</span>}
          <button className="ghost" onClick={reset} disabled={!dirtyCount || saving}>되돌리기</button>
          <button className="primary" onClick={apply} disabled={!dirtyCount || saving}>{saving ? "반영 중…" : "반영 요청"}</button>
        </span>
      </div>
      <div className="rangebar">
        <span className="rangebar-label">매출순위 구간 일괄 지정</span>
        <input className="rg-num" type="number" min="1" value={rgFrom} onChange={(e) => setRgFrom(e.target.value)} placeholder="시작" />
        <span>~</span>
        <input className="rg-num" type="number" min="1" value={rgTo} onChange={(e) => setRgTo(e.target.value)} placeholder="끝" />
        <span>위 →</span>
        <select value={rgTarget} onChange={(e) => setRgTarget(e.target.value)}>
          <option value="일반">일반</option><option value="신규매장">신규매장</option><option value="199매장">199매장</option><option value="노출제외">노출제외</option>
        </select>
        <button className="ghost" onClick={() => {
          const a = parseInt(rgFrom), b = parseInt(rgTo);
          if (!a || !b) { setMsg("시작·끝 순위를 입력하세요"); return; }
          const lo = Math.min(a, b), hi = Math.max(a, b);
          const val = rgTarget === "일반" ? 1 : rgTarget;
          const nd = { ...draft }; rows.forEach((r) => { if (r.rankNow >= lo && r.rankNow <= hi) nd[r.c] = val; });
          setDraft(nd); setMsg(`${lo}~${hi}위 → ${rgTarget} 로 지정(대기)`);
        }}>구간 적용</button>
      </div>
      <div className="tablewrap">
        <table>
          <thead><tr><th className="num">매출순위</th><th>매장</th><th>점장</th><th>셀</th><th className="num">매출총액</th><th>분류 (수정)</th></tr></thead>
          <tbody>
            {rows.map((r) => {
              const pg = curPg(r); const dirty = draft[r.c] !== undefined;
              return (
                <tr key={r.c} className={dirty ? "row-dirty" : ""}>
                  <td className="num">{r.rankNow}</td>
                  <td><span className="storename">{r.store.n}</span><span className="storecode">{r.c}</span></td>
                  <td>{r.store.mg || <span className="dim">–</span>}</td>
                  <td>{r.store.cell || <span className="dim">–</span>}</td>
                  <td className="num">{r.sales ? fmtWon(r.sales) : "–"}</td>
                  <td>
                    <select className="pgselect" value={label(pg)} onChange={(e) => stage(r.c, e.target.value)}>
                      <option value="일반">일반</option><option value="신규매장">신규매장</option><option value="199매장">199매장</option><option value="노출제외">노출제외</option>
                    </select>
                    {dirty && <span className="dirty-dot" title="반영 대기">●</span>}
                    {typeof pg !== "number" && <span className="pill warn" style={{ marginLeft: 6 }}>{pg}</span>}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ───────── 주차 관리 (이상 주차 제외) ───────── */
function WeekFlagsEditor({ refresh }) {
  const [, bump] = React.useReducer((x) => x + 1, 0);
  const [busy, setBusy] = useState(null); // 처리 중인 주차 라벨
  const [msg, setMsg] = useState("");
  // WEEK_LABELS는 앞에 빈 문자열(패딩)이 섞여 있을 수 있음 — 실제 라벨만, 최근 순으로
  const weeks = WEEK_LABELS.map((label, idx) => ({ label, idx })).filter((w) => w.label).reverse();

  const toggle = async (label) => {
    setBusy(label); setMsg("");
    const nowExcluded = !EXCLUDED_WEEKS.has(label);
    try {
      const { error } = await supabase.from("week_flags").upsert({ week_label: label, excluded: nowExcluded, updated_at: new Date().toISOString() }, { onConflict: "week_label" });
      if (error) throw error;
      const next = new Set(EXCLUDED_WEEKS);
      if (nowExcluded) next.add(label); else next.delete(label);
      setExcludedWeeks([...next]);
      setMsg(`${label} — ${nowExcluded ? "제외로 표시했습니다" : "제외 해제했습니다"}`);
      bump(); refresh && refresh();
    } catch (e) {
      setMsg("저장 실패: " + e.message);
    }
    setBusy(null);
  };

  return (
    <div>
      <h3 className="sectionhead">주차 관리 <span>오류가 있었던 주차를 "제외"로 표시하면 월별 보기 집계에서 빠집니다</span></h3>
      <div className="pipe-msg" style={{ marginBottom: 12 }}>
        <b>이 목록은 주차별 추이(트렌드) 데이터를 직접 수정·삭제하는 기능이 아닙니다.</b> 이미 쌓인 숫자 자체를 지우면 그 뒤에 이어진 모든 주차의 이력이 다시 계산돼야 해서 위험합니다. 대신 잘못된 주차를 "제외"로 표시하면, 저장된 값은 그대로 두되 <b>주차별 추이 그래프와 월간·직전4주 집계 양쪽 모두에서 안 보이고 계산에서도 빠집니다.</b> 나중에 "관리자 탭 → 데이터 적재"에서 같은 년/월/주차를 골라 올바른 파일로 다시 계산·반영하면(재게시), 그 주차의 제외 표시가 <b>자동으로 해제</b>되고 새 값이 다시 그래프·집계에 포함됩니다.
      </div>
      <div className="tablewrap">
        <table>
          <thead><tr><th>주차</th><th>상태</th><th></th></tr></thead>
          <tbody>
            {weeks.length === 0 && <tr><td colSpan={3} className="empty">아직 쌓인 주차가 없습니다.</td></tr>}
            {weeks.map((w) => {
              const excluded = EXCLUDED_WEEKS.has(w.label);
              return (
                <tr key={w.idx} className={excluded ? "row-fail" : ""}>
                  <td>{w.label}</td>
                  <td>{excluded ? <span className="pill warn">제외됨</span> : <span className="dimtxt">정상 반영</span>}</td>
                  <td><button className="ghost tiny" disabled={busy === w.label} onClick={() => toggle(w.label)}>{busy === w.label ? "처리 중…" : excluded ? "제외 해제" : "이 주차 제외"}</button></td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      {msg && <div className="pipe-msg" style={{ marginTop: 10 }}>{msg}</div>}
    </div>
  );
}

/* ───────── 비밀번호 변경 (관리자) ───────── */
function PwChange() {
  const [cur, setCur] = useState("");
  const [nw, setNw] = useState("");
  const [nw2, setNw2] = useState("");
  const [msg, setMsg] = useState(null);
  const submit = () => {
    if (cur !== ADMIN_PW) { setMsg({ ok: false, t: "현재 비밀번호가 올바르지 않습니다." }); return; }
    if (nw.length < 4) { setMsg({ ok: false, t: "새 비밀번호는 4자 이상이어야 합니다." }); return; }
    if (nw !== nw2) { setMsg({ ok: false, t: "새 비밀번호 확인이 일치하지 않습니다." }); return; }
    ADMIN_PW = nw;
    setCur(""); setNw(""); setNw2("");
    setMsg({ ok: true, t: "비밀번호가 변경되었습니다. 다음 로그인부터 적용됩니다." });
  };
  return (
    <div>
      <h3 className="sectionhead">비밀번호 변경 <span>관리자 접근 비밀번호를 변경합니다</span></h3>
      <div className="pwform">
        <label>현재 비밀번호<input type="password" value={cur} onChange={(e) => { setCur(e.target.value); setMsg(null); }} /></label>
        <label>새 비밀번호<input type="password" value={nw} onChange={(e) => { setNw(e.target.value); setMsg(null); }} /></label>
        <label>새 비밀번호 확인<input type="password" value={nw2} onChange={(e) => { setNw2(e.target.value); setMsg(null); }} onKeyDown={(e) => e.key === "Enter" && submit()} /></label>
        {msg && <div className={msg.ok ? "pwok" : "pwerr"}>{msg.t}</div>}
        <button className="primary" onClick={submit}>변경</button>
        </div>
    </div>
  );
}

/* ───────── 아웃풋: 엑셀 다운로드 ───────── */
function downloadXlsx(sheets, filename) {
  const wb = XLSX.utils.book_new();
  sheets.forEach(({ name, aoa }) => {
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
  });
  XLSX.writeFile(wb, filename);
}
const pctv = (v) => (v == null ? "" : +(v * 100).toFixed(2));

// ① 종합판
function exportSummary() {
  const head = ["매장코드", "매장명", "셀", "분류", "주간매출", "인건비율(%)", "정직원율(%)", "메이트율(%)", "가산율(%)", "정직원수", "정직원급여", "메이트인건비", "가산수당", "실사용시간", "계약시간", "홀실사용", "주방실사용", "홀계약", "주방계약", "평균시급", "평점", "유사매출대대비gap(%p)"];
  const aoa = [head];
  STORES.forEach((s) => {
    aoa.push([s.c, s.n, s.cell, s.pg, s.s || 0, pctv(s.lt), pctv(s.ls), pctv(s.lm), pctv(s.lo), (s.staff || []).length, s._payStaff || 0, s._payMate || 0, s._payOT || 0, s.useMate || 0, s.ctMate || 0, s.useHall || 0, s.useKit || 0, s.ctHall || 0, s.ctKit || 0, s.realWage || 0, s.rt || "", pctv(s.gap)]);
  });
  downloadXlsx([{ name: "종합판", aoa }], `애슐리_종합판_${(CUR_WEEK_LABEL || "").replace(/[^\w가-힣]/g, "_")}.xlsx`);
}

// ② 매장별 진단
function exportDiagnosis() {
  const head = ["매장코드", "매장명", "셀", "분류", "인건비율(%)", "유사매출대gap(%p)", "진단항목", "심각도", "원인/근거"];
  const aoa = [head];
  STORES.forEach((s) => {
    const dx = diagnose(s);
    if (dx.excluded) { aoa.push([s.c, s.n, s.cell, s.pg, pctv(s.lt), "", "제외", "", dx.reason]); return; }
    if (!dx.tips.length) { aoa.push([s.c, s.n, s.cell, s.pg, pctv(s.lt), pctv(s.gap), "-", "-", "특이사항 없음"]); return; }
    dx.tips.forEach((t) => {
      aoa.push([s.c, s.n, s.cell, s.pg, pctv(s.lt), pctv(s.gap), t.head, t.level === "high" ? "높음" : t.level === "good" ? "양호" : "보통", (t.facts || []).join(" / ")]);
    });
  });
  downloadXlsx([{ name: "매장별진단", aoa }], `애슐리_매장별진단_${(CUR_WEEK_LABEL || "").replace(/[^\w가-힣]/g, "_")}.xlsx`);
}

// ③ 인건비 보정내역
function exportHelpers(dbRows) {
  const head = ["구분", "출발매장", "도착매장", "직급", "인원명", "시작일", "종료일", "금액", "상태", "메모", "등록일"];
  const aoa = [head];
  const codeName = (c) => { const s = STORES.find((x) => x.c === c); return s ? s.n : (c || ""); };
  (dbRows || []).forEach((r) => {
    aoa.push([r.kind, codeName(r.from), r.to ? codeName(r.to) : "", r.pos || "", r.name || "", r.start || "", r.end || "", r.amt || 0, r.status || "", r.memo || "", (r.createdAt || "").slice(0, 10)]);
  });
  downloadXlsx([{ name: "보정내역", aoa }], `애슐리_인건비보정내역_${(CUR_WEEK_LABEL || "").replace(/[^\w가-힣]/g, "_")}.xlsx`);
}

// 매장정보(분류·점장·셀) 통합 양식 다운로드 — 현재 값을 미리 채워서 내려주고, 고친 뒤 그대로 재업로드하면 됨(다른 파일 형식 추측 안 해도 되게 직접 만든 양식)
function exportStoreInfoTemplate() {
  const head = ["매장코드", "매장명(참고용, 안 읽음)", "분류(일반/신규매장/199매장/노출제외)", "점장명", "셀"];
  const aoa = [head];
  [...STORES].sort((a, b) => a.c.localeCompare(b.c)).forEach((s) => {
    aoa.push([s.c, s.n, typeof s.pg === "number" ? "일반" : s.pg, s.mg || "", s.cell || ""]);
  });
  const notes = [
    ["작성 안내"],
    ["1. 매장코드 칸은 절대 지우거나 바꾸지 마세요(이 값으로 매장을 찾습니다). 이 칸이 비어있는 행은 통째로 무시됩니다."],
    ["2. 매장명은 참고용입니다. 이 칸을 고쳐도 반영되지 않습니다."],
    ["3. 분류는 반드시 '일반' / '신규매장' / '199매장' / '노출제외' 중 하나로만 입력하세요. 다른 글자를 쓰면 '일반'으로 처리됩니다."],
    ["4. 점장명·셀은 자유롭게 고치면 됩니다. 비워두면 기존 값이 지워집니다."],
    ["5. 새 매장을 추가하고 싶으면 맨 아래에 새 행으로 매장코드부터 채워서 추가하세요."],
  ];
  downloadXlsx([{ name: "매장정보", aoa }, { name: "작성안내", aoa: notes }], `애슐리_매장정보_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

// 일괄 보정 양식 다운로드
function exportHelperTemplate() {
  const head = ["구분", "출발매장코드", "도착매장코드(헬퍼만)", "직원구분", "직급(정직원)", "인원명", "시작일(YYYY-MM-DD)", "종료일(YYYY-MM-DD)", "반영비율(1 또는 0.5)", "금액(메이트/퇴사연차)", "메모"];
  const ex = [
    ["교육", "AL132", "", "메이트", "", "홍길동", "", "", "1", "150000", "7월 신입 교육"],
    ["연차", "AL392", "", "메이트", "", "김철수", "", "", "1", "80000", "연차수당 보정"],
    ["헬퍼", "AL132", "AL392", "정직원", "캡틴", "이영희", "2026-07-13", "2026-07-15", "1", "", "3일 파견"],
    ["퇴사연차제거", "AL132", "", "메이트", "", "박퇴사", "", "", "1", "120000", "15일 퇴사 잔여연차 제거"],
  ];
  const notes = [
    ["작성 안내"],
    ["1. 구분: 헬퍼 / 교육 / 연차 / 퇴사연차제거 중 하나"],
    ["2. 출발매장코드: 인건비가 발생/차감되는 매장 (예: AL132)"],
    ["3. 도착매장코드: '헬퍼'(파견)일 때만 입력. 파견 받는 매장"],
    ["4. 직원구분: 정직원 또는 메이트"],
    ["5. 정직원: 직급+시작일+종료일 입력 → 금액 자동계산 (금액칸 비움)"],
    ["   직급: 점장/GM/GMIT/매니저/MIT/캡틴/CIT/헤드/HIT/선임점장"],
    ["6. 메이트/퇴사연차제거: 금액 직접 입력 (원 단위 숫자)"],
    ["7. 퇴사연차제거: 15일 포함 주간, 퇴사자에게 발생한 연차수당을 금액에 입력 → 그만큼 차감"],
    ["8. 반영비율: 하루 종일=1, 반일=0.5"],
    ["9. 업로드 후 관리자(본사/셀) 승인 시 지표에 반영됩니다"],
  ];
  downloadXlsx([{ name: "보정입력", aoa: [head, ...ex] }, { name: "작성안내", aoa: notes }], "애슐리_인건비보정_양식.xlsx");
}

// 매장 코드/명 → 코드
function resolveStoreCode(v) {
  if (!v) return null;
  const s = String(v).trim();
  if (STORES.find((x) => x.c === s)) return s;
  const byName = STORES.find((x) => x.n === s || x.n.replace(/\s/g, "") === s.replace(/\s/g, ""));
  return byName ? byName.c : null;
}

// 매출류 파일의 월을 요일 라벨로 특정 (여러 컬럼 교집합)
function pipeSalesMonth(rows, year) {
  const KO = ["일", "월", "화", "수", "목", "금", "토"];
  const hdr = rows[1] || [];
  let cands = null;
  for (let i = 5; i < hdr.length; i++) {
    const m = String(hdr[i] || "").match(/^(\d+)\((.)\)/);
    if (!m) continue;
    const day = +m[1], w = m[2];
    const ok = [];
    for (let mo = 1; mo <= 12; mo++) { const d = new Date(year, mo - 1, day); if (d.getMonth() === mo - 1 && KO[d.getDay()] === w) ok.push(mo); }
    cands = cands === null ? ok : cands.filter((x) => ok.includes(x));
    if (cands.length === 1) break;
  }
  return cands && cands.length ? cands[0] : null;
}
// 매출류 파일(들) → {isoDate: {code: 매출}}
function pipeSalesByDate(fileList, year) {
  const map = {};
  fileList.forEach((rows) => {
    const mo = pipeSalesMonth(rows, year); if (!mo) return;
    const hdr = rows[1] || []; const dayCols = [];
    for (let i = 5; i < hdr.length; i++) { const m = String(hdr[i] || "").match(/^(\d+)\(/); if (m) dayCols.push({ i, day: +m[1] }); }
    for (let r = 2; r < rows.length; r++) {
      const row = rows[r]; const code = row && row[1]; if (!code) continue;
      dayCols.forEach(({ i, day }) => {
        const v = pN(row[i]); if (!v) return;
        const iso = `${year}-${String(mo).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
        (map[iso] = map[iso] || {})[code] = v;
      });
    }
  });
  return map;
}

/* ───────── 데이터 검증(정의서) ───────── */
function VerifyTab() {
  const rows = STORES.filter((s) => (s.pg !== "신규매장" || s.s > 0) && s.pg !== "노출제외").map((s) => {
    const mate = (s._payMate || 0) + (s._payOT || 0);
    const flags = [];
    if (!s.s) flags.push("매출 없음");
    if (s.s && !mate) flags.push("메이트급여 없음");
    if (s.s && !s._payStaff) flags.push("정직원급여 없음");
    if (s.s && !s.useMate) flags.push("실사용시간 없음");
    if (s.s && !s.ctMate) flags.push("계약시간 없음");
    if (s.s && !s.rt) flags.push("평점 없음");
    if (s.s && (!s.slotWD || s.slotWD.every((v) => !v))) flags.push("시간대 인원 없음");
    return { s, mate, flags };
  }).sort((a, b) => (b.flags.length - a.flags.length) || (b.s.s || 0) - (a.s.s || 0));

  const T = {
    n: rows.length,
    sales: rows.reduce((a, r) => a + (r.s.s || 0), 0),
    mate: rows.reduce((a, r) => a + r.mate, 0),
    staff: rows.reduce((a, r) => a + (r.s._payStaff || 0), 0),
    use: rows.reduce((a, r) => a + (r.s.useMate || 0), 0),
    ct: rows.reduce((a, r) => a + (r.s.ctMate || 0), 0),
    bad: rows.filter((r) => r.flags.length).length,
  };
  const totLt = T.sales ? (T.mate + T.staff) / T.sales : 0;

  return (
    <div>
      <h3 className="sectionhead">데이터 검증 <span>이번 주차에 적재된 값 — 매장 대조용</span></h3>
      <div className="dl-bar">
        <button className="dl-btn" onClick={exportSummary}>⬇ 종합판 다운로드</button>
        <button className="dl-btn" onClick={exportDiagnosis}>⬇ 매장별 진단 다운로드</button>
      </div>
      <div className="vf-cards">
        <div className="vf-card"><div className="vf-k">매장 수</div><div className="vf-v">{T.n}</div></div>
        <div className="vf-card"><div className="vf-k">주간매출 합</div><div className="vf-v">{fmtWon(T.sales)}</div></div>
        <div className="vf-card"><div className="vf-k">메이트 인건비 합</div><div className="vf-v">{fmtWon(T.mate)}</div></div>
        <div className="vf-card"><div className="vf-k">정직원 인건비 합</div><div className="vf-v">{fmtWon(T.staff)}</div></div>
        <div className="vf-card"><div className="vf-k">전체 인건비율</div><div className="vf-v">{fmtPct(totLt)}</div></div>
        <div className="vf-card"><div className="vf-k">메이트 실사용 / 계약</div><div className="vf-v">{fmtNum(T.use, 0)} / {fmtNum(T.ct, 0)}h</div></div>
        <div className={"vf-card" + (T.bad ? " bad" : " ok")}><div className="vf-k">이상 매장</div><div className="vf-v">{T.bad}곳</div></div>
      </div>

      <div className="tablewrap">
        <table className="vf-table">
          <thead><tr>
            <th>매장</th><th>셀</th><th className="num">주간매출</th><th className="num">정직원(명/원)</th>
            <th className="num">메이트 인건비</th><th className="num">실사용h</th><th className="num">계약h</th>
            <th className="num">인건비율</th><th className="num">평점</th><th>이상</th>
          </tr></thead>
          <tbody>
            {rows.map(({ s, mate, flags }) => (
              <tr key={s.c} className={flags.length ? "vf-bad" : ""}>
                <td className="vf-name">{s.n}</td>
                <td className="vf-cell">{s.cell}</td>
                <td className="num">{fmtWon(s.s)}</td>
                <td className="num">{(s.staff || []).length}명 / {fmtWon(s._payStaff)}</td>
                <td className="num">{fmtWon(mate)}</td>
                <td className="num">{fmtNum(s.useMate, 1)}</td>
                <td className="num">{fmtNum(s.ctMate, 1)}</td>
                <td className="num strong">{fmtPct(s.lt)}</td>
                <td className="num">{s.rt || "–"}</td>
                <td className="vf-flag">{flags.length ? flags.join(", ") : ""}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* ───────── 파이프라인: 원본 파일 → 스냅샷 계산 ───────── */
const PIPE_PAY = { "점장": 4500000, "GM": 4000000, "GMIT": 3500000, "매니저": 3170000, "MIT": 3170000, "캡틴": 2920000, "CIT": 2700000, "헤드": 2700000, "HIT": 2700000, "선임점장": 2350000 };
const PIPE_WK = 4.345;
const pN = (v) => { const n = parseFloat(v); return isFinite(n) ? n : 0; };

function pipeDetect(rows) {
  const h0 = (rows[0] || []).map((x) => String(x || "")).join("|");
  const h1 = (rows[1] || []).map((x) => String(x || "")).join("|");
  if (h0.includes("직원구분") && h0.includes("직위")) return "근무자";
  if (h0.includes("고객 평점") || h0.includes("고객평점")) return "평점";
  if (h0.includes("사용자아이디") && h0.includes("출근시간기록")) return "출퇴근";
  if (h0.includes("총 급여") || h1.includes("주휴수당") || h0.includes("주휴수당")) return "메이트";
  if (h1.includes("브랜드명") && h1.includes("영업일수")) return "매출류";
  return "알수없음";
}
// 매출류에서 목표주(월~일) 컬럼 자동탐지: 데이터 있는 마지막 일요일 기준 직전 7일
function pipeWeekCols(rows) {
  const hdr = rows[1] || [];
  const cols = []; // {i, day, wdayKo}
  for (let i = 5; i < hdr.length; i++) {
    const m = String(hdr[i] || "").match(/^(\d+)\((.)\)/);
    if (m) cols.push({ i, day: +m[1], w: m[2] });
  }
  // 데이터 있는 마지막 컬럼 찾기
  let last = -1;
  for (const c of cols) {
    let has = 0;
    for (let r = 2; r < Math.min(rows.length, 40); r++) if (pN((rows[r] || [])[c.i]) > 0) has++;
    if (has >= 3) last = cols.indexOf(c);
  }
  if (last < 0) return cols.slice(-7).map((c) => c.i);
  // last가 일요일이 아니면, 마지막 완주(월~일) 블록으로 back-align
  let end = last;
  while (end >= 0 && cols[end].w !== "일") end--;
  if (end < 6) end = last; // 못 찾으면 그냥 마지막 7개
  return cols.slice(Math.max(0, end - 6), end + 1).map((c) => c.i);
}

// ===== 파싱 공용화 1~3단계 =====
// pipeBuild()와 PartialRepublish()에 완전히 똑같은 파싱 로직이 두 번 복사돼 있던 것을 함수 3개로 뽑음.
// ※ 지금 이 3개 함수는 아직 아무도 호출하지 않음 — pipeBuild/PartialRepublish는 기존 인라인 코드로 계속 동작(4~5단계에서 교체 예정).
// 세 함수는 순서대로 의존한다: 출퇴근 파싱(3번)은 근무자 파싱(1번)의 lookupCode/uidInfo와 메이트급여 파싱(2번)의 uid2job/uid2jobMix가 있어야 동작함.

// 1) 근무자관리 파일 파싱 — 이름↔코드 매핑, 정직원 급여 합계, 메이트 계약시간(홀/주방), UID 정보, 매장코드→이름
function pipeParseWorker(근무자) {
  const norm = (x) => String(x || "").replace(/[\s()·・.\-_]/g, "").toLowerCase();
  const name2code = {}; for (let i = 1; i < 근무자.length; i++) { const r = 근무자[i]; if (r && r[3] && r[4]) { const nm = String(r[4]).trim(); name2code[nm] = r[3]; name2code[norm(nm)] = r[3]; } }
  const lookupCode = (nm) => name2code[String(nm).trim()] || name2code[norm(nm)] || null;
  const staff = {}, staffPay = {}, ctMate = {}, ctHK = {};
  for (let i = 1; i < 근무자.length; i++) {
    const r = 근무자[i]; if (!r || !r[3]) continue;
    const code = r[3], emp = r[8], pos = String(r[9] || "").trim(), work = r[12];
    if (emp === "정직원" && work === "재직") { (staff[code] = staff[code] || []).push(pos); staffPay[code] = (staffPay[code] || 0) + (PIPE_PAY[pos] || 0); }
    if (emp === "메이트" && work === "재직") {
      const ctv = pN(r[18]); ctMate[code] = (ctMate[code] || 0) + ctv;
      const jb = String(r[11]); ctHK[code] = ctHK[code] || { 홀: 0, 주방: 0 };
      if (jb === "홀") ctHK[code].홀 += ctv; else if (jb === "주방") ctHK[code].주방 += ctv; else { ctHK[code].홀 += ctv / 2; ctHK[code].주방 += ctv / 2; }
    }
  }
  const uidInfo = {}; for (let i = 1; i < 근무자.length; i++) { const r = 근무자[i]; if (r && r[5] != null) uidInfo[Math.round(pN(r[5]))] = { code: r[3], emp: r[8], job: String(r[11]) }; }
  const codeName = {}; for (let i = 1; i < 근무자.length; i++) { const r = 근무자[i]; if (r && r[3] && r[4]) codeName[r[3]] = String(r[4]).trim(); }
  return { name2code, lookupCode, staff, staffPay, ctMate, ctHK, uidInfo, codeName };
}

// 2) 메이트급여(주간) 파일 파싱 — UID→홀/주방 구분, 매장별 정상/야간/추가/휴일/휴업/연차/주휴 금액·시간(혼합근무는 절반씩)
function pipeParseMatePay(메이트) {
  const uid2job = {}, uid2jobMix = {}; // uid2job: 시간대별 인원세팅용 · uid2jobMix: 실사용시간 홀/주방 분배용(혼합근무 절반씩)
  for (let i = 2; i < 메이트.length; i++) { const r = 메이트[i]; if (r && r[3] != null) { const uid = Math.round(pN(r[3])); const jv = String(r[7]); uid2job[uid] = jv === "홀" ? "홀" : "주방"; uid2jobMix[uid] = jv === "홀" ? "홀" : jv === "주방" ? "주방" : "혼합"; } }
  const mate = {};
  for (let i = 2; i < 메이트.length; i++) {
    const r = 메이트[i]; if (!r || !r[0]) continue; const c = r[0];
    const jv = String(r[7]); const job = jv === "홀" ? "홀" : jv === "주방" ? "주방" : "혼합"; // 홀/주방이 아닌 값(혼합근무)은 절반씩 반영
    mate[c] = mate[c] || { 홀: {}, 주방: {} };
    const put = (k, v) => { if (job === "혼합") { mate[c].홀[k] = (mate[c].홀[k] || 0) + v / 2; mate[c].주방[k] = (mate[c].주방[k] || 0) + v / 2; } else { mate[c][job][k] = (mate[c][job][k] || 0) + v; } };
    const add = (k, ci) => put(k, pN(r[ci]));
    add("정상", 12); add("야간", 14); add("휴일", 16); add("추가", 18); add("휴업", 20); add("연차", 22); add("주휴", 23); add("총", 24);
    put("정상시간", pN(r[11]) / 60); put("야간시간", pN(r[13]) / 60); put("휴일시간", pN(r[15]) / 60); put("추가시간", pN(r[17]) / 60);
  }
  return { uid2job, uid2jobMix, mate };
}

// 3) 출퇴근기록부 파일 파싱 — 실사용시간(홀/주방/정직원), 시간대별 인원(cnt/days/slotHK)
// ctx: 1·2단계 결과 중 필요한 것만 { lookupCode, uidInfo, uid2job, uid2jobMix }
function pipeParseAttendance(출퇴근, lo, hi, ctx) {
  const { lookupCode, uidInfo, uid2job, uid2jobMix } = ctx;
  const SLOTS = [600, 750, 930, 1110, 1290];
  const tmin = (v) => { if (typeof v === "string" && v.length >= 16) { const h = +v.slice(11, 13), m = +v.slice(14, 16); return isFinite(h) ? h * 60 + m : null; } if (v instanceof Date) return v.getHours() * 60 + v.getMinutes(); return null; };
  const cnt = {}, days = {};
  const useHK = {}; // 홀/주방 실사용시간(메이트)
  const useStaff = {}; // 정직원 실사용시간
  const MATE_TITLE = /메이트|트레이너|파트타임|파트장/;
  for (let i = 1; i < 출퇴근.length; i++) {
    const r = 출퇴근[i]; if (!r) continue;
    const uid = r[0], nm = r[2], grade = String(r[3] || ""), day = r[4] ? String(r[4]).slice(0, 10) : null, ci = tmin(r[6]), co = tmin(r[8]);
    if (!day || day < lo || day > hi) continue;
    const code = nm ? lookupCode(nm) : null;
    if (!code || ci == null || co == null || co <= ci) continue;
    const wd = (new Date(day + "T00:00:00").getDay());
    const wk = (wd === 0 || wd === 6) ? "we" : "wd";
    days[code] = days[code] || { wd: new Set(), we: new Set() };
    days[code][wk].add(day);
    const info = uid != null ? uidInfo[Math.round(pN(uid))] : null;
    const isMate = MATE_TITLE.test(grade);
    const key = isMate ? (uid2job[Math.round(pN(uid))] || "정") : "정";
    cnt[code] = cnt[code] || { wd: Array.from({ length: 5 }, () => ({})), we: Array.from({ length: 5 }, () => ({})) };
    SLOTS.forEach((mid, si) => { if (ci <= mid && mid <= co) { const o = cnt[code][wk][si]; o[key] = (o[key] || 0) + 1; } });
    if (isMate) {
      const jvMix = uid2jobMix[Math.round(pN(uid))];
      useHK[code] = useHK[code] || { 홀: 0, 주방: 0 };
      if (jvMix === "혼합") { useHK[code].홀 += pN(r[10]) / 2; useHK[code].주방 += pN(r[10]) / 2; }
      else { const jb = jvMix || ((info && info.job === "홀") ? "홀" : "주방"); useHK[code][jb] += pN(r[10]); }
    } else {
      useStaff[code] = (useStaff[code] || 0) + pN(r[10]);
    }
  }
  const slotHK = (code) => { const o = { wd: { 종합: [], 홀: [], 주방: [] }, we: { 종합: [], 홀: [], 주방: [] } }; for (const wk of ["wd", "we"]) { const nd = (days[code] && days[code][wk].size) || 1; for (let si = 0; si < 5; si++) { const c = (cnt[code] && cnt[code][wk][si]) || {}; const h = (c.홀 || 0) / nd, k = (c.주방 || 0) / nd, st = (c.정 || 0) / nd; o[wk].홀.push(+h.toFixed(2)); o[wk].주방.push(+k.toFixed(2)); o[wk].종합.push(+(h + k + st).toFixed(2)); } } return o; };
  return { useHK, useStaff, cnt, days, slotHK };
}

// files: {근무자, 메이트, 매출, 예상, 출퇴근, 평점} 각 rows배열
// "2026-07-2" 같은 주차번호 → "7월 2주차" 같은 짧은 라벨(추이 그래프 x축 표시용)
function shortWeekLabel(weekStr) {
  const m = String(weekStr || "").trim().match(/(\d{4})-(\d{1,2})-(\d{1,2})$/); // 끝에서 매칭 → 앞에 숫자가 잘못 더 붙어도(예: "22026-08-1") 복구됨
  if (!m) return String(weekStr || "").trim() || "신규주차";
  return `${+m[2]}월 ${+m[3]}주차`;
}
function pipeBuild(files, prevStores, weekStr) {
  const { 근무자, 메이트, 출퇴근, 평점 } = files;
  const 매출Arr = Array.isArray(files.매출) ? files.매출 : [files.매출];
  const 예상Arr = Array.isArray(files.예상) ? files.예상 : [files.예상];
  const { lookupCode, staff, staffPay, ctMate, ctHK, uidInfo, codeName } = pipeParseWorker(근무자);
  const { uid2job, uid2jobMix, mate } = pipeParseMatePay(메이트);
  // 고객평점 매칭: 첫 컬럼이 매장코드(AL077 등)면 코드로, 아니면 매장명으로(공백·괄호·기호 무시하고 느슨하게) 매칭
  const normName = (v) => String(v || "").replace(/[\s()·・.\-_]/g, "");
  const code2rt = {}, name2rt = {};
  const codeSet = new Set(STORES.map((s) => s.c));
  for (let i = 1; i < 평점.length; i++) {
    const r = 평점[i]; if (!r || !r[0]) continue;
    const key = String(r[0]).trim();
    if (codeSet.has(key)) code2rt[key] = pN(r[1]);
    else name2rt[normName(key)] = pN(r[1]);
  }
  const rtOf = (code, name) => code2rt[code] ?? name2rt[normName(name)] ?? null;
  // 목표주 날짜범위: 출퇴근에서 자동(가장 최근 월~일). 간단히 데이터의 최대 일자 기준 직전 7일.
  let maxDay = ""; for (let i = 1; i < 출퇴근.length; i++) { const d = 출퇴근[i] && 출퇴근[i][4]; if (d) { const s = String(d).slice(0, 10); if (s > maxDay) maxDay = s; } }
  const end = new Date(maxDay + "T00:00:00"); const endW = (end.getDay() + 6) % 7; if (endW !== 6) end.setDate(end.getDate() - (endW + 1)); // 직전 일요일
  const start = new Date(end); start.setDate(end.getDate() - 6);
  const iso = (d) => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`; // 로컬 날짜(시간대 밀림 방지)
  const lo = iso(start), hi = iso(end);
  const year = +lo.slice(0, 4);
  // 매출: 날짜 기반 합산 (월 경계 주는 두 파일에 걸쳐 합산)
  const salesByDate = pipeSalesByDate(매출Arr, year);
  const sales = {};
  Object.keys(salesByDate).forEach((iso2) => { if (iso2 >= lo && iso2 <= hi) { const day = salesByDate[iso2]; Object.keys(day).forEach((c) => sales[c] = (sales[c] || 0) + day[c]); } });
  // 평일/주말 일평균 매출 — 지금까지 이 계산이 파이프라인에 없어서(정적 시드 데이터만 있었음) 신규매장은 항상 0만으로 표시됐음
  const wdSum = {}, wdCnt = {}, weSum = {}, weCnt = {};
  Object.keys(salesByDate).forEach((iso2) => {
    if (iso2 < lo || iso2 > hi) return;
    const dow = new Date(iso2 + "T00:00:00").getDay(); // 0=일,6=토
    const bucket = (dow === 0 || dow === 6) ? "we" : "wd";
    const day = salesByDate[iso2];
    Object.keys(day).forEach((c) => {
      if (bucket === "we") { weSum[c] = (weSum[c] || 0) + day[c]; weCnt[c] = (weCnt[c] || 0) + 1; }
      else { wdSum[c] = (wdSum[c] || 0) + day[c]; wdCnt[c] = (wdCnt[c] || 0) + 1; }
    });
  });
  const wdAvgOf = (code) => wdCnt[code] ? Math.round(wdSum[code] / wdCnt[code]) : 0;
  const weAvgOf = (code) => weCnt[code] ? Math.round(weSum[code] / weCnt[code]) : 0;
  // 예상: 목표주 월요일(=lo)이 속한 달의 예상 파일 선택
  const targetMonth = +lo.slice(5, 7);
  let 예상 = 예상Arr[0];
  if (예상Arr.length > 1) { const m = 예상Arr.find((rows) => pipeSalesMonth(rows, year) === targetMonth); if (m) 예상 = m; }
  const fcast = {}; for (let i = 2; i < (예상 ? 예상.length : 0); i++) { const r = 예상[i]; if (r && r[1]) fcast[r[1]] = pN(r[3]); }
  const { useHK, useStaff, slotHK } = pipeParseAttendance(출퇴근, lo, hi, { lookupCode, uidInfo, uid2job, uid2jobMix });

  const stores = []; const warnings = [];
  const prev = {}; prevStores.forEach((s) => prev[s.c] = s);
  // 그 주 매출이 잡힌(>0) 매장 = 실제 영업 매장. 기존 목록에 없으면 신규매장으로 추가
  Object.keys(sales).forEach((code) => {
    if (sales[code] > 0 && !prev[code]) { prev[code] = { c: code, n: codeName[code] || code, cell: "", pg: "신규매장", mg: "", staff: [], slots: (prevStores[0] && prevStores[0].slots) ? prevStores[0].slots.map((sl) => ({ k: sl.k, t: sl.t, set: 0 })) : [], newStore: true }; warnings.push(`신규매장 추가: ${codeName[code] || code} (${code}) — 셀 지정 필요`); }
  });
  for (const code in prev) { const st0 = prev[code]; const s = pN(sales[code]); if (s <= 0) warnings.push(`${st0.n || code}(${code}): 이번 주 매출이 0원입니다 — 인건비율 등 매출 기반 지표는 계산하지 않고 비워둡니다(–). 매출 파일을 확인하세요.`); const dH = (mate[code] && mate[code].홀) || {}, dK = (mate[code] && mate[code].주방) || {}; const baseH = (dH.정상 || 0) + (dH.주휴 || 0) + (dH.연차 || 0) + (dH.휴업 || 0), otH = (dH.야간 || 0) + (dH.추가 || 0) + (dH.휴일 || 0); const baseK = (dK.정상 || 0) + (dK.주휴 || 0) + (dK.연차 || 0) + (dK.휴업 || 0), otK = (dK.야간 || 0) + (dK.추가 || 0) + (dK.휴일 || 0); const payOT = otH + otK, payStaffW = (staffPay[code] || 0) / PIPE_WK; const sv = s; const totMate = (dH.총 || 0) + (dK.총 || 0); const ls = sv > 0 ? payStaffW / sv : null, lm = sv > 0 ? totMate / sv : null, lo2 = sv > 0 ? payOT / sv : null, lt = (ls != null && lm != null) ? ls + lm : null; const normH = (dH.정상시간 || 0) + (dK.정상시간 || 0); const realWage = normH ? Math.round(totMate / normH) : 0; const hk = slotHK(code);
    // 이번 주 계약시간 = 저번 주에 올렸던 근무자관리 파일 값(그때는 "다음 주"로 표시됐던 값) — 저번 주 월요일부터 이번 주 일요일까지 실제로 적용된 계약
    // 비교할 과거(저번 주 ct5)가 없는 신규매장 등은 억지로 이번 주 값을 채우지 않고 비워둠(–)
    // 다음 주 계약시간 = 이번에 올리는 근무자관리 파일 값(이번 주 월요일 업로드 = 다음 월~일 스케줄)
    const ct4v = st0.ct5 != null ? +(st0.ct5).toFixed(1) : null;
    const ct5v = +(ctMate[code] || 0).toFixed(1);
    const md = (d) => { const o = {}; ["정상", "야간", "휴일", "추가", "휴업", "연차", "주휴"].forEach((k) => o[k] = Math.round(d[k] || 0)); o.정상시간 = +(d.정상시간 || 0).toFixed(1); o.야간시간 = +(d.야간시간 || 0).toFixed(1); o.추가시간 = +(d.추가시간 || 0).toFixed(1); o.휴일시간 = +(d.휴일시간 || 0).toFixed(1); return o; };
    // 평균 시급 구성(시급 단위 분해) — 정상/야간/추가/주휴/기타(휴일+휴업)를 정상시간으로 나눔. 연차는 급여주기상 0원이라 제외.
    const wageBreak = normH > 0 ? (() => {
      const sum = (k) => (dH[k] || 0) + (dK[k] || 0);
      const base = Math.round(sum("정상") / normH), weekly = Math.round(sum("주휴") / normH), extra = Math.round(sum("추가") / normH), night = Math.round(sum("야간") / normH), etc = Math.round((sum("휴일") + sum("휴업")) / normH);
      return { base, weekly, extra, night, etc, total: realWage };
    })() : null;
    if (lt > 0.5) warnings.push(`${st0.n}: 인건비율 ${(lt * 100).toFixed(0)}% (비정상)`); stores.push({ ...st0, s: Math.round(s), ps: Math.round(fcast[code] || st0.ps || 0), _payStaff: Math.round(payStaffW), _payMate: Math.round(totMate), _payOT: Math.round(payOT), _payStaffBase: Math.round(payStaffW), _payMateBase: Math.round(totMate), _helperCnt: 0, _helperNetWon: 0, ls: ls != null ? +ls.toFixed(5) : null, lm: lm != null ? +lm.toFixed(5) : null, lo: lo2 != null ? +lo2.toFixed(5) : null, lt: lt != null ? +lt.toFixed(5) : null, lmHallBase: sv > 0 ? +((dH.총 || 0) / sv).toFixed(5) : null, lmKitBase: sv > 0 ? +((dK.총 || 0) / sv).toFixed(5) : null, loHall: sv > 0 ? +(otH / sv).toFixed(5) : null, loKit: sv > 0 ? +(otK / sv).toFixed(5) : null, realWage, wageBreak, mateExtraH: +((dH.추가시간 || 0) + (dK.추가시간 || 0)).toFixed(1), mateNormH: +normH.toFixed(1), staff: (staff[code] || []).map((p) => ({ pos: p })), mateDetail: { 홀: md(dH), 주방: md(dK) }, rt: rtOf(code, st0.n) ?? st0.rt ?? 0, slotHK: hk, slotWD: hk.wd.종합, slotWE: hk.we.종합, useHall: +((useHK[code] && useHK[code].홀) || 0).toFixed(1), useKit: +((useHK[code] && useHK[code].주방) || 0).toFixed(1), useMate: +(((useHK[code] && useHK[code].홀) || 0) + ((useHK[code] && useHK[code].주방) || 0)).toFixed(1), ctHall: +((ctHK[code] && ctHK[code].홀) || 0).toFixed(1), ctKit: +((ctHK[code] && ctHK[code].주방) || 0).toFixed(1), ctSum: +(((ctHK[code] && ctHK[code].홀) || 0) + ((ctHK[code] && ctHK[code].주방) || 0)).toFixed(1), ctHallR: (((ctHK[code] && ctHK[code].홀) || 0) + ((ctHK[code] && ctHK[code].주방) || 0)) > 0 ? +(((ctHK[code] && ctHK[code].홀) || 0) / (((ctHK[code] && ctHK[code].홀) || 0) + ((ctHK[code] && ctHK[code].주방) || 0))).toFixed(4) : null, ctMate: +(ctMate[code] || 0).toFixed(1), wkMate: +(((useHK[code] && useHK[code].홀) || 0) + ((useHK[code] && useHK[code].주방) || 0)).toFixed(1), wkTot: +(((useHK[code] && useHK[code].홀) || 0) + ((useHK[code] && useHK[code].주방) || 0) + (useStaff[code] || 0)).toFixed(1), useStaff: +((useStaff[code] || 0)).toFixed(1), ct4: ct4v, ct5: ct5v, wdAvg: wdAvgOf(code), weAvg: weAvgOf(code) }); }
  // 피어 gap — A타입: 클라이언트 peersOf()와 동일한 peersBySalesBand() 함수 사용(그룹번호 없음, 199매장은 199매장끼리)
  // ※ 이 값은 게시 후 클라이언트에서 recomputeMetrics()가 다시 계산해 덮어쓰지만, 미리보기 단계에서도 같은 숫자가 보이도록 규칙을 맞춰둠
  const trim = (v) => { v = v.slice().sort((a, b) => a - b); if (v.length > 2) v = v.slice(1, -1); return v.length ? v.reduce((a, b) => a + b, 0) / v.length : null; };
  stores.forEach((s) => {
    if (s.pg === "신규매장" || s.pg === "노출제외" || !s.s) { s.gap = null; return; }
    const peers = [s, ...peersBySalesBand(s, stores)];
    const pAvg = trim(peers.map((p) => p.lt));
    s.gap = pAvg == null ? null : +(s.lt - pAvg).toFixed(5);
  });

  // 주차별 추이 스택 적재 — 매장별 trend.lt/lo/sales, 그리고 "매장별" 피어 평균 추이(peerTrend, 매장코드로 키 — 그룹번호가 없어졌으므로)
  // (여기서는 result만 만들 뿐, 실제 전역 WEEK_LABELS/DATA.peerTrend에 반영은 publish() 때 initData()가 함 — "계산" 미리보기만으론 안 늘어남)
  // ※ peerTrend는 주차 이력이 필요해서 여기서 계속 누적 계산함. peerWE/peerWage/peerRating/peerSlotHK/peerMate는
  //   과거 이력이 필요 없는 "이번 주 스냅샷"이라 화면에서 그때그때 즉시 계산하도록 바꿨고(republish 안 기다려도 됨), 여기선 뺐음.
  const weekLabel = shortWeekLabel(weekStr);
  const existingIdx = WEEK_LABELS.indexOf(weekLabel); // 이미 게시된 적 있는 주차면 그 인덱스 자리를 "덮어쓰기"(수정), 없으면 새 주차로 "추가"
  const prevWeekCount = WEEK_LABELS.length; // 이번 주 반영 전까지 쌓여있던 주차 수
  const pushWeek = (arr, value) => {
    const base = Array.isArray(arr) ? arr.slice() : [];
    if (existingIdx >= 0) { while (base.length <= existingIdx) base.push(null); base[existingIdx] = value; return base; } // 같은 주차 재게시 = 그 자리 값만 교체(수정), 새 칸 안 늘어남
    while (base.length < prevWeekCount) base.push(null); // 신규매장 등 과거 데이터 없는 경우 null로 패딩해 인덱스를 맞춤
    base.push(value);
    return base;
  };
  stores.forEach((s) => {
    const t0 = s.trend || {};
    s.trend = {
      ...t0,
      lt: pushWeek(t0.lt, s.lt != null ? s.lt : null),
      lo: pushWeek(t0.lo, s.lo != null ? s.lo : null),
      lm: pushWeek(t0.lm, s.lm != null ? s.lm : null),
      sales: pushWeek(t0.sales, s.s != null ? s.s : null),
      wkTot: pushWeek(t0.wkTot, s.wkTot != null ? s.wkTot : null),
      wkMate: pushWeek(t0.wkMate, s.wkMate != null ? s.wkMate : null),
    };
  });
  const peerTrend = {};
  stores.forEach((s) => {
    if (s.pg === "신규매장" || s.pg === "노출제외" || !s.s) return; // 비교 대상 없음 → peerTrend 갱신 안 함(과거 값이 있었다면 그대로 유지)
    const peers = [s, ...peersBySalesBand(s, stores)];
    const pt0 = DATA.peerTrend && DATA.peerTrend[s.c] || {};
    peerTrend[s.c] = { ...pt0, lt: pushWeek(pt0.lt, trim(peers.map((p) => p.lt))), lo: pushWeek(pt0.lo, trim(peers.map((p) => p.lo))) };
  });
  Object.keys(DATA.peerTrend || {}).forEach((code) => { if (!peerTrend[code]) peerTrend[code] = DATA.peerTrend[code]; }); // 이번 주 비교대상 없던 매장의 과거 이력 보존

  return { stores, warnings, weekRange: `${lo} ~ ${hi}`, lo, hi, salesCount: Object.keys(sales).length, weekLabel, isCorrection: existingIdx >= 0, peerTrend };
}

let PIPE_FILES = {}; // 파이프라인 업로드 파일 유지 (탭 이동해도 안 사라짐)
let PIPE_WEEK = "";
function PipelineTab({ refresh }) {
  const [week, setWeek] = useState(PIPE_WEEK);
  const parseWeek = (w) => { const m = String(w || "").match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/); const now = new Date(); return m ? { y: m[1], mo: String(+m[2]), wk: String(+m[3]) } : { y: String(now.getFullYear()), mo: String(now.getMonth() + 1), wk: "1" }; };
  const initW = parseWeek(week);
  const [wYear, setWYear] = useState(initW.y);
  const [wMonth, setWMonth] = useState(initW.mo);
  const [wWeekNo, setWWeekNo] = useState(initW.wk);
  React.useEffect(() => { setWeek(`${wYear}-${String(wMonth).padStart(2, "0")}-${wWeekNo}`); }, [wYear, wMonth, wWeekNo]);
  const [files, setFilesRaw] = useState(PIPE_FILES);   // type → {name, rows}
  const setFiles = (nf) => { PIPE_FILES = nf; setFilesRaw(nf); };
  React.useEffect(() => { PIPE_WEEK = week; }, [week]);
  const [result, setResult] = useState(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const NEED = [["근무자", "매장근무자조회"], ["메이트", "메이트급여(주간)"], ["매출", "일자별매출"], ["예상", "일자별매출(예상)"], ["출퇴근", "출퇴근기록부"], ["평점", "고객평점"]];

  const onFiles = async (fileList) => {
    const nf = { ...files };
    for (const file of Array.from(fileList)) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
        let t = pipeDetect(rows);
        if (t === "매출류") { // 실적/예상 구분: 파일명 우선, 없으면 영업일수
          const fn = String(file.name);
          if (/예상|forecast|plan/i.test(fn)) t = "예상";
          else if (/실적|actual/.test(fn)) t = "매출";
          else { const dcol = rows.slice(2, 30).map((r) => pN(r && r[4])); const avg = dcol.reduce((a, b) => a + b, 0) / (dcol.length || 1); t = avg >= 20 ? "예상" : "매출"; }
        }
        if (t === "알수없음") { setMsg(`인식 실패: ${file.name}`); continue; }
        if (t === "매출" || t === "예상") { const arr = Array.isArray(nf[t]) ? nf[t] : []; if (!arr.some((x) => x.name === file.name)) arr.push({ name: file.name, rows }); nf[t] = arr.slice(-2); } // 최대 2개(월경계)
        else nf[t] = { name: file.name, rows };
      } catch (e) { setMsg(`읽기 실패: ${file.name}`); }
    }
    setFiles(nf); setResult(null);
  };

  // 항목별 수동 업로드 (자동 판별 무시, 지정 칸에 강제 배치)
  const assignFile = async (type, file) => {
    if (!file) return;
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
      const nf = { ...files };
      if (type === "매출" || type === "예상") { const arr = Array.isArray(nf[type]) ? nf[type] : []; if (!arr.some((x) => x.name === file.name)) arr.push({ name: file.name, rows }); nf[type] = arr.slice(-2); }
      else nf[type] = { name: file.name, rows };
      setFiles(nf); setResult(null); setMsg(`${file.name} → ${type} 수동 지정됨`);
    } catch (e) { setMsg(`읽기 실패: ${file.name}`); }
  };

  const compute = () => {
    const missing = NEED.filter(([k]) => { const v = files[k]; return Array.isArray(v) ? !v.length : !v; });
    if (missing.length) { setMsg("부족한 파일: " + missing.map((m) => m[1]).join(", ")); return; }
    try {
      const r = pipeBuild({ 근무자: files.근무자.rows, 메이트: files.메이트.rows, 매출: files.매출.map((f) => f.rows), 예상: files.예상.map((f) => f.rows), 출퇴근: files.출퇴근.rows, 평점: files.평점.rows }, STORES, week.trim());
      setResult(r); setMsg("");
    } catch (e) { setMsg("계산 오류: " + e.message); }
  };

  const publish = async () => {
    if (!result) return;
    if (!week.trim()) { setMsg("주차를 입력하세요 (예: 2026-07-2)"); return; }
    setBusy(true); setMsg("");
    // pipeBuild가 이미 "새 주차인지 / 기존 주차 재게시(수정)인지" 판단해뒀으므로 그대로 사용(따로 다시 계산 안 함 — 계산 시점과 어긋나는 것 방지)
    const weekLabel = result.weekLabel;
    const newWeekLabels = result.isCorrection ? WEEK_LABELS : [...WEEK_LABELS, weekLabel]; // 기존 주차 재게시면 라벨 배열은 그대로(중복 추가 안 함)
    const out = { ...DATA, stores: result.stores, peerTrend: result.peerTrend, _week: week.trim(), _weekRange: result.weekRange.replace(/-/g, "/").replace(/\d{4}\//g, ""), _weekLo: result.lo, _weekHi: result.hi, _label: `${week.trim()} 주차 (${result.weekRange})`, _weekLabels: newWeekLabels };
    const blob = new Blob([JSON.stringify(out)], { type: "application/json" });
    const path = `ashley_week_${week.trim()}.json`;
    try {
      const { error } = await supabase.storage.from("snapshots").upload(path, blob, { upsert: true, contentType: "application/json" });
      if (error) throw error;
      // 기존 주차를 재게시(수정)한 거라면, 그 주차가 "제외" 표시돼 있었을 수 있으니 자동으로 해제(고친 데이터니 다시 계산에 포함되게)
      if (result.isCorrection && EXCLUDED_WEEKS.has(weekLabel)) {
        try { await supabase.from("week_flags").delete().eq("week_label", weekLabel); } catch (e) { /* 무시 - 다음 로드 때 다시 시도됨 */ }
        const next = new Set(EXCLUDED_WEEKS); next.delete(weekLabel); setExcludedWeeks([...next]);
      }
      setMsg(`저장 완료: ${path} — 대시보드에 반영하려면 DATA_URL을 이 파일로 바꾸거나 새로고침하세요.`);
      // 즉시 로컬 반영
      initData(out); recomputeMetrics(); refresh && refresh();
    } catch (e) { setMsg("저장 실패: " + e.message); }
    setBusy(false);
  };

  return (
    <div>
      <h3 className="sectionhead">데이터 적재 <span>차E 원본 파일 6종을 올리면 자동 계산 → 검증 → 저장</span></h3>
      <div className="pipe-week">
        <label>주차
          <span style={{ display: "inline-flex", gap: 6, verticalAlign: "middle" }}>
            <select value={wYear} onChange={(e) => setWYear(e.target.value)}>{[wYear - 1, +wYear, +wYear + 1, +wYear + 2].filter((v, i, a) => a.indexOf(v) === i).sort().map((y) => <option key={y} value={y}>{y}년</option>)}</select>
            <select value={wMonth} onChange={(e) => setWMonth(e.target.value)}>{Array.from({ length: 12 }, (_, i) => i + 1).map((m) => <option key={m} value={m}>{m}월</option>)}</select>
            <select value={wWeekNo} onChange={(e) => setWWeekNo(e.target.value)}>{[1, 2, 3, 4, 5].map((w) => <option key={w} value={w}>{w}주차</option>)}</select>
          </span>
          <span className="dimtxt" style={{ marginLeft: 8, fontSize: 12 }}>→ 그래프에 "{shortWeekLabel(week)}"로 표시됩니다</span>
        </label>
      </div>
      <label className="dropzone">
        <input type="file" multiple accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => { onFiles(e.target.files); e.target.value = ""; }} />
        <div className="dropzone-big">파일 6종 선택 또는 끌어다 놓기</div>
        <div className="dropzone-sub">매장근무자·메이트급여(주간)·일자별매출·일자별매출(예상)·출퇴근기록부·고객평점 (종류 자동 인식) · 분류가 틀리면 아래 각 항목의 ＋ 로 직접 지정</div>
      </label>
      <div className="pipe-files">
        {NEED.map(([k, label]) => {
          const v = files[k]; const has = Array.isArray(v) ? v.length > 0 : !!v;
          const nameTxt = Array.isArray(v) ? (v.length ? v.map((x) => x.name).join(", ") : "미업로드") : (v ? v.name : "미업로드");
          const multi = k === "매출" || k === "예상";
          return (
            <div key={k} className={"pipe-file" + (has ? " ok" : "")}>
              <span className="pf-check">{has ? "✓" : "○"}</span>
              <span className="pf-label">{label}{multi ? " (월경계 시 2개)" : ""}</span>
              <span className="pf-name">{nameTxt}{Array.isArray(v) && v.length > 1 ? ` · ${v.length}개` : ""}</span>
              <label className="pf-manual" title="이 항목으로 직접 올리기">＋<input type="file" accept=".xlsx,.xls,.csv" style={{ display: "none" }} onChange={(e) => { assignFile(k, e.target.files[0]); e.target.value = ""; }} /></label>
              {has && <button className="pf-remove" title="내리기" onClick={() => { const nf = { ...files }; delete nf[k]; setFiles(nf); setResult(null); setMsg(""); }}>✕</button>}
            </div>
          );
        })}
      </div>
      {Object.keys(files).length > 0 && <button className="pf-clearall" onClick={() => { setFiles({}); setResult(null); setMsg(""); }}>전체 비우기</button>}
      <div className="pipe-actions">
        <button className="primary" onClick={compute} disabled={NEED.some(([k]) => { const v = files[k]; return Array.isArray(v) ? !v.length : !v; })}>계산 · 검증</button>
        {result && (
          <span className={result.isCorrection ? "pill warn" : "pill"} style={{ marginRight: 8 }}>
            {result.isCorrection ? `"${result.weekLabel}" 기존 주차 수정 — 반영하면 그 주차 값을 덮어씁니다` : `"${result.weekLabel}" 새 주차로 추가됩니다`}
          </span>
        )}
        {result && <button className="primary" onClick={publish} disabled={busy}>{busy ? "저장 중…" : "반영 (Storage 저장)"}</button>}
      </div>
      {msg && <div className="pipe-msg">{msg}</div>}
      {result && (
        <div className="pipe-report">
          <h4>검증 리포트</h4>
          <div className="pipe-report-row">기간(자동 인식): {result.weekRange} · 매출 매장 {result.salesCount}개</div>
          <table className="pipe-table">
            <thead><tr><th>매장</th><th className="num">주간매출</th><th className="num">인건비율</th><th className="num">정직원</th><th className="num">메이트</th><th className="num">가산</th></tr></thead>
            <tbody>
              {["AL132", "AL392"].concat(result.stores.slice(0, 6).map((s) => s.c)).filter((v, i, a) => a.indexOf(v) === i).slice(0, 8).map((c) => {
                const s = result.stores.find((x) => x.c === c); if (!s) return null;
                return <tr key={c}><td>{s.n}</td><td className="num">{fmtWon(s.s)}</td><td className="num strong">{fmtPct(s.lt)}</td><td className="num">{fmtPct(s.ls)}</td><td className="num">{fmtPct(s.lm)}</td><td className="num">{fmtPct(s.lo)}</td></tr>;
              })}
            </tbody>
          </table>
          {result.warnings.length > 0 && <div className="pipe-warn">⚠ {result.warnings.slice(0, 5).join(" / ")}</div>}
          {!result.warnings.length && <div className="pipe-ok">이상값 없음 · 검증 통과</div>}
        </div>
      )}
      <PartialRepublish refresh={refresh} />
    </div>
  );
}

/* ───────── 부분 재게시: 과거에 이미 게시된 주차 중 일부 항목만 다시 계산해서 그 자리를 덮어씀 ─────────
   - "매출만": 매출 파일 하나만 있으면 됨. 저장돼있던 그 주의 급여 금액(비율 역산)에 새 매출을 대입해 인건비율만 다시 계산.
   - "근무 데이터": 정직원/메이트 급여, 가산, 사용시간 등을 다시 계산. 단, 출퇴근기록부는 그 자체로는 어느 매장 소속인지
     알 수 없어(이름으로 매장코드를 찾음) 근무자관리·메이트급여 파일이 같이 있어야 함(3개). 매출은 저장된 값을 그대로 씀. */
function PartialRepublish({ refresh }) {
  const [open, setOpen] = useState(false);
  const [mode, setMode] = useState("sales"); // sales | work
  const now = new Date();
  const [wYear, setWYear] = useState(String(now.getFullYear()));
  const [wMonth, setWMonth] = useState(String(now.getMonth() + 1));
  const [wWeekNo, setWWeekNo] = useState("1");
  const [monday, setMonday] = useState(""); // 이 주차의 월요일 날짜(직접 지정) — 파일에서 며칠치를 뽑을지 결정
  const [salesFile, setSalesFile] = useState(null);
  const [workFiles, setWorkFiles] = useState({}); // {근무자, 메이트, 출퇴근}
  const [preview, setPreview] = useState(null);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");

  const targetLabel = `${+wMonth}월 ${wWeekNo}주차`;
  const targetIdx = WEEK_LABELS.indexOf(targetLabel);

  const readXlsx = async (file) => {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    return XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: null, raw: true });
  };

  const buildPreview = async () => {
    setMsg(""); setPreview(null);
    if (targetIdx === -1) { setMsg(`"${targetLabel}"는 아직 게시된 적이 없는 주차입니다. 이 기능은 기존 주차 일부만 고치는 용도라, 먼저 "데이터 적재"로 그 주차를 한 번 게시해야 사용할 수 있습니다.`); return; }
    if (!monday) { setMsg("이 주차의 월요일 날짜를 선택해주세요(파일에서 어느 7일치를 뽑을지 결정하는 기준입니다)."); return; }
    const lo = monday; const hiDate = new Date(monday + "T00:00:00"); hiDate.setDate(hiDate.getDate() + 6);
    const hi = `${hiDate.getFullYear()}-${String(hiDate.getMonth() + 1).padStart(2, "0")}-${String(hiDate.getDate()).padStart(2, "0")}`;
    setBusy(true);
    try {
      if (mode === "sales") {
        if (!salesFile) { setMsg("매출 파일을 올려주세요."); setBusy(false); return; }
        const rows = await readXlsx(salesFile);
        const salesByDate = pipeSalesByDate([rows], +wYear);
        const sales = {};
        Object.keys(salesByDate).forEach((iso) => { if (iso >= lo && iso <= hi) { const day = salesByDate[iso]; Object.keys(day).forEach((c) => (sales[c] = (sales[c] || 0) + day[c])); } });
        if (!Object.keys(sales).length) { setMsg(`파일에서 ${lo}~${hi} 기간 매출을 못 찾았습니다. 월요일 날짜나 파일을 확인해주세요.`); setBusy(false); return; }
        const updates = []; const skipped = [];
        STORES.forEach((s) => {
          const newSales = sales[s.c]; if (newSales == null) return;
          const t = s.trend || {};
          const oldSales = t.sales && t.sales[targetIdx], ltOld = t.lt && t.lt[targetIdx];
          if (oldSales == null || !oldSales || ltOld == null) { skipped.push(`${s.n}: 이 주차의 기존 매출·인건비율 이력이 없어 역산 불가`); return; }
          const lmOld = t.lm && t.lm[targetIdx], loOld = t.lo && t.lo[targetIdx];
          const oldSv = oldSales, newSv = newSales;
          const totalAmt = ltOld * oldSv;
          const mateAmt = lmOld != null ? lmOld * oldSv : null;
          const otAmt = loOld != null ? loOld * oldSv : null;
          const newLt = newSv > 0 ? totalAmt / newSv : null;
          const newLm = (mateAmt != null && newSv > 0) ? mateAmt / newSv : lmOld;
          const newLo = (otAmt != null && newSv > 0) ? otAmt / newSv : loOld;
          const staffAmt = mateAmt != null ? totalAmt - mateAmt : null;
          updates.push({ code: s.c, name: s.n, oldSales, newSales, oldLt: ltOld, newLt, newLm, newLo, newPayStaffBase: staffAmt != null ? Math.round(staffAmt) : null, newPayMateBase: mateAmt != null ? Math.round(mateAmt) : null });
        });
        setPreview({ kind: "sales", lo, hi, updates, skipped });
      } else {
        const need = ["근무자", "메이트", "출퇴근"];
        if (need.some((k) => !workFiles[k])) { setMsg("근무자관리·메이트급여(주간)·출퇴근기록부 3개 파일이 모두 필요합니다."); setBusy(false); return; }
        const 근무자 = await readXlsx(workFiles.근무자), 메이트 = await readXlsx(workFiles.메이트), 출퇴근 = await readXlsx(workFiles.출퇴근);
        const { lookupCode, staffPay, uidInfo } = pipeParseWorker(근무자);
        const { uid2job, uid2jobMix, mate } = pipeParseMatePay(메이트);
        const { useHK, useStaff } = pipeParseAttendance(출퇴근, lo, hi, { lookupCode, uidInfo, uid2job, uid2jobMix });
        const updates = []; const skipped = [];
        const codesFound = new Set([...Object.keys(mate), ...Object.keys(staffPay), ...Object.keys(useHK), ...Object.keys(useStaff)]);
        STORES.forEach((s) => {
          if (!codesFound.has(s.c)) return;
          const t = s.trend || {};
          const oldSales = t.sales && t.sales[targetIdx];
          if (oldSales == null || !oldSales) { skipped.push(`${s.n}: 이 주차의 매출 이력이 없어 인건비율 계산 불가(매출은 그대로 두고 반영 안 함)`); return; }
          const dH = (mate[s.c] && mate[s.c].홀) || {}, dK = (mate[s.c] && mate[s.c].주방) || {};
          const otH = (dH.야간 || 0) + (dH.추가 || 0) + (dH.휴일 || 0), otK = (dK.야간 || 0) + (dK.추가 || 0) + (dK.휴일 || 0);
          const payOT = otH + otK, payStaffW = (staffPay[s.c] || 0) / PIPE_WK, sv = oldSales;
          const totMate = (dH.총 || 0) + (dK.총 || 0);
          const newLs = payStaffW / sv, newLm = totMate / sv, newLo = payOT / sv, newLt = newLs + newLm;
          const normH = (dH.정상시간 || 0) + (dK.정상시간 || 0);
          const extraH = +((dH.추가시간 || 0) + (dK.추가시간 || 0)).toFixed(1);
          const wkTot = +(((useHK[s.c] && useHK[s.c].홀) || 0) + ((useHK[s.c] && useHK[s.c].주방) || 0) + (useStaff[s.c] || 0)).toFixed(1);
          const newUseStaff = +((useStaff[s.c] || 0)).toFixed(1);
          updates.push({ code: s.c, name: s.n, oldLt: t.lt && t.lt[targetIdx], newLt, newLm, newLo, oldWkTot: t.wkTot && t.wkTot[targetIdx], newWkTot: wkTot, newMateNormH: +normH.toFixed(1), newUseStaff, newMateExtraH: extraH, newPayStaffBase: Math.round(payStaffW), newPayMateBase: Math.round(totMate) });
        });
        setPreview({ kind: "work", lo, hi, updates, skipped });
      }
    } catch (e) { setMsg("처리 실패: " + e.message); }
    setBusy(false);
  };

  const apply = async () => {
    if (!preview) return;
    setBusy(true);
    const isLatest = targetIdx === WEEK_LABELS.length - 1;
    preview.updates.forEach((u) => {
      const s = STORES.find((x) => x.c === u.code); if (!s || !s.trend) return;
      if (preview.kind === "sales") {
        if (s.trend.sales) s.trend.sales[targetIdx] = u.newSales;
        if (s.trend.lt) s.trend.lt[targetIdx] = u.newLt;
        if (s.trend.lm) s.trend.lm[targetIdx] = u.newLm;
        if (s.trend.lo) s.trend.lo[targetIdx] = u.newLo;
        if (isLatest) { s.s = Math.round(u.newSales); if (u.newPayStaffBase != null) s._payStaffBase = u.newPayStaffBase; if (u.newPayMateBase != null) s._payMateBase = u.newPayMateBase; }
      } else {
        if (s.trend.lt) s.trend.lt[targetIdx] = u.newLt;
        if (s.trend.lm) s.trend.lm[targetIdx] = u.newLm;
        if (s.trend.lo) s.trend.lo[targetIdx] = u.newLo;
        if (s.trend.wkTot) s.trend.wkTot[targetIdx] = u.newWkTot;
        if (isLatest) { s._payStaffBase = u.newPayStaffBase; s._payMateBase = u.newPayMateBase; s.wkTot = u.newWkTot; s.mateNormH = u.newMateNormH; s.useStaff = u.newUseStaff; s.mateExtraH = u.newMateExtraH; }
      }
    });
    await refreshHelperEffects(); // base가 바뀌었으니, "지금 승인된 헬퍼" 기준으로 _payStaff/_payMate(및 lt)를 다시 계산(recomputeMetrics 포함)
    try {
      // base(_payStaffBase/_payMateBase)와 라이브 헬퍼효과가 분리돼 있어서(base/live 재설계),
      // STORES를 그냥 그대로 저장해도 안전합니다 — 다음 로드 때 refreshHelperEffects()가 항상 처음부터 다시 계산해주므로,
      // 저장 시점에 헬퍼효과가 얼마나 들어있었는지는 중요하지 않습니다(예전엔 이걸 되돌렸다 다시 적용하는 복잡한 춤이 필요했음).
      const out = { ...DATA, stores: STORES };
      const blob = new Blob([JSON.stringify(out)], { type: "application/json" });
      const path = `ashley_week_${DATA._week || "current"}.json`;
      const { error } = await supabase.storage.from("snapshots").upload(path, blob, { upsert: true, contentType: "application/json" });
      if (error) throw error;
      setMsg(`"${targetLabel}" 부분 반영 완료 — ${preview.updates.length}개 매장 갱신`);
      initData(out);
      await refreshHelperEffects(); // initData가 STORES를 새로 불러왔으니, 화면용으로 라이브 헬퍼효과를 다시 계산
      refresh && refresh();
      setPreview(null);
    } catch (e) { setMsg("저장 실패: " + e.message); }
    setBusy(false);
  };

  return (
    <div className="partial-republish">
      <button className="ghost" onClick={() => setOpen((v) => !v)}>{open ? "▾" : "▸"} 부분 재게시 (과거 주차 일부 항목만 다시 계산)</button>
      {open && (
        <div className="pipe-msg" style={{ marginTop: 10 }}>
          <p className="hint mini">이미 게시된 과거 주차 중, 특정 항목(매출 또는 근무 데이터)만 다시 계산해서 그 주차 자리를 덮어씁니다. 다른 항목은 그대로 유지됩니다. 6개 파일을 다 모을 필요 없이, 해당하는 파일만 있으면 됩니다.</p>
          <div className="formrow">
            <label>수정할 주차
              <span style={{ display: "inline-flex", gap: 6, verticalAlign: "middle" }}>
                <select value={wYear} onChange={(e) => setWYear(e.target.value)}>{[+wYear - 1, +wYear, +wYear + 1].map((y) => <option key={y} value={y}>{y}년</option>)}</select>
                <select value={wMonth} onChange={(e) => setWMonth(e.target.value)}>{Array.from({ length: 12 }, (_, i) => i + 1).map((m) => <option key={m} value={m}>{m}월</option>)}</select>
                <select value={wWeekNo} onChange={(e) => setWWeekNo(e.target.value)}>{[1, 2, 3, 4, 5].map((w) => <option key={w} value={w}>{w}주차</option>)}</select>
              </span>
            </label>
            <label>이 주차의 월요일 날짜<input type="date" value={monday} onChange={(e) => setMonday(e.target.value)} /></label>
          </div>
          <span className="dimtxt" style={{ fontSize: 12 }}>{targetIdx === -1 ? `⚠ "${targetLabel}"는 게시된 적 없는 주차입니다` : `"${targetLabel}" 자리를 덮어씁니다`}</span>
          <div className="daytoggle" style={{ marginTop: 10 }}>
            <button className={mode === "sales" ? "on" : ""} onClick={() => { setMode("sales"); setPreview(null); }}>매출만</button>
            <button className={mode === "work" ? "on" : ""} onClick={() => { setMode("work"); setPreview(null); }}>근무 데이터(정직원·메이트·사용시간)</button>
          </div>
          {mode === "sales" ? (
            <div className="formrow" style={{ marginTop: 10 }}>
              <label>일자별매출 파일<input type="file" accept=".xlsx,.xls" onChange={(e) => setSalesFile(e.target.files[0])} /></label>
            </div>
          ) : (
            <div className="formrow" style={{ marginTop: 10, flexWrap: "wrap" }}>
              <label>매장근무자조회<input type="file" accept=".xlsx,.xls" onChange={(e) => setWorkFiles({ ...workFiles, 근무자: e.target.files[0] })} /></label>
              <label>메이트급여(주간)<input type="file" accept=".xlsx,.xls" onChange={(e) => setWorkFiles({ ...workFiles, 메이트: e.target.files[0] })} /></label>
              <label>출퇴근기록부<input type="file" accept=".xlsx,.xls" onChange={(e) => setWorkFiles({ ...workFiles, 출퇴근: e.target.files[0] })} /></label>
            </div>
          )}
          <div className="formrow" style={{ marginTop: 10 }}>
            <button className="ghost" disabled={busy} onClick={buildPreview}>{busy ? "처리 중…" : "미리보기 계산"}</button>
            {preview && <button className="primary" disabled={busy} onClick={apply}>{busy ? "저장 중…" : `이 내용으로 반영 (${preview.updates.length}개 매장)`}</button>}
          </div>
          {msg && <div className="pipe-msg">{msg}</div>}
          {preview && (
            <div className="tablewrap" style={{ marginTop: 10 }}>
              <table>
                <thead><tr><th>매장</th>{preview.kind === "sales" ? <><th className="num">기존 매출</th><th className="num">새 매출</th></> : null}<th className="num">기존 인건비율</th><th className="num">새 인건비율</th>{preview.kind === "work" && <><th className="num">기존 사용시간</th><th className="num">새 사용시간</th></>}</tr></thead>
                <tbody>
                  {preview.updates.slice(0, 30).map((u) => (
                    <tr key={u.code}>
                      <td>{u.name}</td>
                      {preview.kind === "sales" ? <><td className="num">{fmtWon(u.oldSales)}</td><td className="num">{fmtWon(u.newSales)}</td></> : null}
                      <td className="num">{fmtPct(u.oldLt)}</td><td className="num strong">{fmtPct(u.newLt)}</td>
                      {preview.kind === "work" && <><td className="num">{u.oldWkTot ?? "–"}</td><td className="num strong">{u.newWkTot}</td></>}
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.updates.length > 30 && <p className="hint mini">그 외 {preview.updates.length - 30}개 매장 더 있음(반영에는 전부 포함됩니다)</p>}
              {preview.skipped.length > 0 && <div className="pipe-warn">⚠ {preview.skipped.slice(0, 5).join(" / ")}</div>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ───────── 6. 관리자 (데이터 적재 — 엑셀 자동 인식·반영) ───────── */
// 파일 시그니처 감지: 상단 5행에서 헤더 키워드 조합으로 파일 종류 판별
function detectFileType(rows) {
  const flat = rows.slice(0, 6).flat().map((v) => String(v ?? ""));
  const has = (kw) => flat.some((v) => v.includes(kw));
  if (has("매장코드") && has("주휴수당") && has("총 급여")) return "matePay";
  if (has("매장코드") && has("합계") && has("영업일수")) return "sales";
  if (has("매장코드") && has("출근") && has("퇴근")) return "attendance";
  if (has("사번") && has("발령")) return "resign";
  if (has("긍정") || has("백분위")) return "rating";
  return null;
}

function parseSales(rows) {
  // 헤더행: '매장코드'…'합계'…'영업일수'
  const hIdx = rows.findIndex((r) => r.some((v) => String(v ?? "").includes("매장코드")));
  const header = rows[hIdx].map((v) => String(v ?? ""));
  const cCode = header.findIndex((v) => v.includes("매장코드"));
  const cSum = header.findIndex((v) => v.includes("합계"));
  let applied = 0, skipped = 0;
  for (let i = hIdx + 1; i < rows.length; i++) {
    const code = String(rows[i][cCode] ?? "");
    if (!code.startsWith("AL")) continue;
    const st = STORES.find((s) => s.c === code);
    const val = Number(rows[i][cSum]);
    if (st && isFinite(val) && val > 0) { st.s = val; applied++; } else skipped++;
  }
  return { applied, skipped, msg: `누적매출 갱신 ${applied}개 매장` };
}

function parseMatePay(rows) {
  // row0 헤더 고정 레이아웃: c0 매장코드, c12 정상수당, c14 야간, c16 휴일, c18 추가, c20 휴업, c22 연차, c23 주휴
  const agg = {};
  let people = 0;
  for (let i = 2; i < rows.length; i++) {
    const code = String(rows[i][0] ?? "");
    if (!code.startsWith("AL")) continue;
    const n = (j) => { const v = Number(rows[i][j]); return isFinite(v) ? v : 0; };
    const a = (agg[code] = agg[code] || { mate: 0, ot: 0 });
    a.mate += n(12) + n(22) + n(23);      // 정상 + 연차 + 주휴
    a.ot += n(14) + n(18) + n(16) + n(20); // 야간 + 추가 + 휴일 + 휴업
    people++;
  }
  let applied = 0;
  Object.entries(agg).forEach(([code, a]) => {
    const st = STORES.find((s) => s.c === code);
    if (st) { st._payMate = a.mate; st._payOT = a.ot; applied++; }
  });
  return { applied, skipped: 0, msg: `메이트급여 반영 ${applied}개 매장 · ${people.toLocaleString()}명` };
}

const ADMIN_FILES = [
  { key: "sales", name: "일자별매출", ready: true, checks: ["매장코드 매핑", "합계·영업일수 검증"] },
  { key: "matePay", name: "메이트급여 (+시급산출용)", ready: true, checks: ["매장별 정상·주휴·연차 / OT 집계", "재입사자 시간 제거(예정)"] },
  { key: "attendance", name: "출퇴근기록부 (정직원/메이트)", ready: false, checks: ["시간대별 세팅 재계산", "연차 외 휴가 보정"] },
  { key: "roster", name: "근무자관리 (주차 스냅샷)", ready: false, checks: ["입·퇴사 diff", "인원구성 갱신"] },
  { key: "resign", name: "메이트 퇴사발령", ready: false, checks: ["계약시간 자동 차감"] },
  { key: "rating", name: "고객평점", ready: false, checks: ["매장명 → 코드 매핑"] },
];

/* ───────── 권한 관리 (approvers 테이블 CRUD) ───────── */
const roleLabel = (r) => r === "hq" ? "본사 관리자 · 관리자 탭 접근 + 전체 보정 승인" : r === "cell" ? "셀 승인자 · 해당 셀 보정 승인" : r === "store" ? "일반매장 · 접속 + 인건비 보정 신청" : "미승인 · 접속 불가";
function AccessTab() {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [msg, setMsg] = useState("");
  const [saving, setSaving] = useState(false);
  const [codeSupported, setCodeSupported] = useState(true); // approvers 테이블에 code 컬럼이 있는지
  const [f, setF] = useState({ email: "", role: "hq", cell: getCells()[0] || "", code: (STORES[0] && STORES[0].c) || "" });

  const load = async () => {
    setLoading(true);
    try {
      // code 컬럼이 아직 없는(마이그레이션 전) 테이블도 지원 — 우선 시도 후 실패하면 role,cell만으로 폴백
      let { data, error } = await supabase.from("approvers").select("email,role,cell,code").order("email");
      if (error) {
        setCodeSupported(false);
        ({ data, error } = await supabase.from("approvers").select("email,role,cell").order("email"));
      }
      if (error) throw error;
      setRows(data || []);
    } catch (e) { setMsg("목록을 불러오지 못했습니다: " + e.message + " (approvers 테이블·권한 설정을 확인하세요)"); }
    setLoading(false);
  };
  React.useEffect(() => { load(); }, []);

  const addApprover = async () => {
    const email = f.email.trim().toLowerCase();
    if (!email || !email.includes("@")) { setMsg("이메일을 정확히 입력하세요."); return; }
    setSaving(true); setMsg("");
    try {
      const payload = { email, role: f.role, cell: f.role === "cell" ? f.cell : null };
      if (codeSupported) payload.code = f.role === "store" ? f.code : null;
      const { error } = await supabase.from("approvers").upsert(payload, { onConflict: "email" });
      if (error) throw error;
      setMsg(`${email} 권한 부여 완료 — ${roleLabel(f.role)}`);
      setF({ ...f, email: "" });
      load();
    } catch (e) {
      const hint = /column .*code.* does not exist/i.test(e.message) ? " — approvers 테이블에 code 컬럼을 추가해야 일반매장 권한을 저장할 수 있습니다 (Supabase SQL: alter table approvers add column code text;)" : "";
      setMsg("저장 실패: " + e.message + hint);
    }
    setSaving(false);
  };

  const removeApprover = async (email) => {
    if (email.toLowerCase() === MASTER_EMAIL.toLowerCase()) { setMsg("마스터 계정은 권한을 회수할 수 없습니다."); return; }
    if (typeof window !== "undefined" && !window.confirm(`${email} 님의 권한을 회수할까요?\n(로그인 계정 자체는 삭제되지 않고, 이 사람이 갖고 있던 접속·승인 권한만 없어집니다. 다음 로그인부터 접속이 막힙니다)`)) return;
    setSaving(true); setMsg("");
    try {
      const { error } = await supabase.from("approvers").delete().eq("email", email);
      if (error) throw error;
      setMsg(`${email} 권한 회수 완료`);
      load();
    } catch (e) { setMsg("삭제 실패: " + e.message); }
    setSaving(false);
  };

  return (
    <div>
      <h3 className="sectionhead">권한 관리 <span>이메일 계정에 접속·관리자 탭·보정 승인 권한을 부여/회수합니다</span></h3>
      <div className="pipe-msg" style={{ marginBottom: 12 }}>
        <b>이 목록에 없는(미승인) 이메일은 로그인 자체가 막힙니다.</b> 아래 3단계 중 하나를 반드시 부여해야 접속할 수 있습니다.<br />
        <b>① 일반매장(store)</b> — 접속 가능 + 자기 매장 인건비 보정 신청. 지정한 매장으로 고정됩니다.<br />
        <b>② 셀 승인자(cell)</b> — 관리자 탭엔 여전히 비밀번호 필요, 지정한 셀 소속 매장의 보정 요청만 승인.<br />
        <b>③ 본사 관리자(hq)</b> — 관리자 탭 비밀번호 없이 바로 접속 + 전체 보정 승인.
      </div>
      {!codeSupported && <div className="pwerr" style={{ marginBottom: 10 }}>approvers 테이블에 <code>code</code> 컬럼이 없어 "일반매장" 권한은 저장되지 않습니다. Supabase에서 <code>alter table approvers add column code text;</code> 실행 후 새로고침 해주세요.</div>}
      <div className="formrow" style={{ display: "flex", gap: 8, flexWrap: "wrap", marginBottom: 12 }}>
        <input placeholder="이메일 (로그인 계정과 동일해야 함)" value={f.email} onChange={(e) => setF({ ...f, email: e.target.value })} style={{ flex: "1 1 240px" }} />
        <select value={f.role} onChange={(e) => setF({ ...f, role: e.target.value })}>
          <option value="hq">본사 관리자(hq)</option>
          <option value="cell">셀 승인자(cell)</option>
          <option value="store">일반매장(store)</option>
        </select>
        {f.role === "cell" && (
          <select value={f.cell} onChange={(e) => setF({ ...f, cell: e.target.value })}>
            {getCells().map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        )}
        {f.role === "store" && (
          <select value={f.code} onChange={(e) => setF({ ...f, code: e.target.value })}>
            {orderStores(STORES).map((s) => <option key={s.c} value={s.c}>{s.n}</option>)}
          </select>
        )}
        <button className="primary" disabled={saving} onClick={addApprover}>{saving ? "처리 중…" : "권한 부여"}</button>
      </div>
      {msg && <div className="pipe-msg">{msg}</div>}
      <div className="tablewrap">
        <table>
          <thead><tr><th>이메일</th><th>권한</th><th>셀 / 매장</th><th></th></tr></thead>
          <tbody>
            {loading ? <tr><td colSpan="4">불러오는 중…</td></tr> : rows.length === 0 ? <tr><td colSpan="4">부여된 권한이 없습니다.</td></tr> : rows.map((r) => {
              const isMaster = r.email.toLowerCase() === MASTER_EMAIL.toLowerCase();
              const storeName = r.code ? (STORES.find((s) => s.c === r.code)?.n || r.code) : null;
              return (
                <tr key={r.email}>
                  <td>{r.email}{isMaster && <span className="pgtag" style={{ marginLeft: 6 }}>마스터</span>}</td>
                  <td>{roleLabel(isMaster ? "hq" : r.role)}</td>
                  <td>{r.cell || storeName || "–"}</td>
                  <td>{isMaster ? <span className="dimtxt">회수 불가</span> : <button className="ghost tiny" onClick={() => removeApprover(r.email)}>권한 회수</button>}</td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
      <div className="dimtxt" style={{ marginTop: 10, fontSize: 12 }}>
        ※ 이 목록은 Supabase <code>approvers</code> 테이블을 그대로 조회·수정합니다. 로그인 계정(이메일·비밀번호) 자체는 Supabase Auth에서 별도로 만들어야 하며, 여기서는 "이미 있는 로그인 계정에 어떤 권한을 줄지"만 관리합니다. 저장이 안 되면 Supabase 쪽에서 approvers 테이블에 대한 쓰기 권한(RLS 정책)이 막혀 있는지 확인이 필요합니다.
      </div>
    </div>
  );
}

let ADMIN_PW = "ashley1!"; // 데모 기본값 · 관리자 탭에서 변경 가능(세션 한정) · 실서비스는 서버 인증

function AdminTab({ refresh }) {
  const isHq = AUTH.role === "hq"; // 권한 관리에서 hq를 부여받은 계정은 비밀번호 없이 통과
  const [authed, setAuthed] = useState(isHq);
  const [pw, setPw] = useState("");
  const [err, setErr] = useState(false);
  const [uploads, setUploads] = useState({}); // key -> [{name, at}]
  const [rejects, setRejects] = useState([]); // 인식 실패 [{name, at, reason}]
  const [adminView, setAdminView] = useState("data");
  const tryLogin = () => { if (pw === ADMIN_PW) { setAuthed(true); setErr(false); } else { setErr(true); } };
  const now = () => new Date().toLocaleString("ko-KR", { year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" });
  const addUpload = (key, name) => setUploads((u) => ({ ...u, [key]: [{ name, at: now() }, ...(u[key] || [])] }));
  const addReject = (name, reason) => setRejects((r) => [{ name, at: now(), reason }, ...r]);

  const onFiles = async (fileList) => {
    for (const file of Array.from(fileList)) {
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array" });
        let handled = false;
        for (const name of wb.SheetNames) {
          const rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1 });
          const t = detectFileType(rows);
          if (t === "sales") { parseSales(rows); recomputeMetrics(); refresh(); addUpload("sales", file.name); handled = true; break; }
          if (t === "matePay") { parseMatePay(rows); recomputeMetrics(); refresh(); addUpload("matePay", file.name); handled = true; break; }
          if (t === "attendance") { addUpload("attendance", file.name); handled = true; break; }
          if (t === "resign") { addUpload("resign", file.name); handled = true; break; }
          if (t === "rating") { addUpload("rating", file.name); handled = true; break; }
        }
        if (!handled) addReject(file.name, "형식을 인식하지 못했습니다");
      } catch (e) {
        addReject(file.name, "읽기 실패: " + e.message);
      }
    }
  };

  if (!authed) {
    return (
      <div className="pwgate">
        <div className="pwcard">
          <div className="pwlock">🔒</div>
          <h3>관리자 전용</h3>
          <p>데이터 적재는 관리자만 접근할 수 있습니다. 비밀번호를 입력하세요.<br /><span style={{ fontSize: 12, color: "#68737E" }}>※ "권한 관리"에서 본사 관리자(hq)로 등록된 이메일로 로그인하면 이 화면 없이 바로 들어옵니다.</span></p>
          <input type="password" value={pw} autoFocus placeholder="비밀번호"
            onChange={(e) => { setPw(e.target.value); setErr(false); }}
            onKeyDown={(e) => e.key === "Enter" && tryLogin()} />
          {err && <div className="pwerr">비밀번호가 올바르지 않습니다.</div>}
          <button className="primary" onClick={tryLogin}>입력</button>
        </div>
      </div>
    );
  }

  return (
    <div>
      <div className="datasource">
        <span className="ds-label">현재 데이터</span>
        <span className="ds-week">{CUR_WEEK_LABEL}</span>
        <span className="ds-file">{DATA_SOURCE.name || "폴백(기본 파일)"}</span>
        {DATA_SOURCE.loadedAt && <span className="ds-at">업로드 {DATA_SOURCE.loadedAt}</span>}
      </div>
      <div className="subnav">
        <button className={adminView === "data" ? "on" : ""} onClick={() => setAdminView("data")}>데이터 적재</button>
        <button className={adminView === "verify" ? "on" : ""} onClick={() => setAdminView("verify")}>데이터 검증</button>
        <button className={adminView === "peer" ? "on" : ""} onClick={() => setAdminView("peer")}>매장 분류</button>
        <button className={adminView === "weeks" ? "on" : ""} onClick={() => setAdminView("weeks")}>주차 관리</button>
        <button className={adminView === "access" ? "on" : ""} onClick={() => setAdminView("access")}>권한 관리</button>
        <button className={adminView === "pw" ? "on" : ""} onClick={() => setAdminView("pw")}>비밀번호 변경</button>
      </div>
      {adminView === "pw" ? <PwChange /> : adminView === "peer" ? <PeerClassEditor refresh={refresh} /> : adminView === "weeks" ? <WeekFlagsEditor refresh={refresh} /> : adminView === "verify" ? <VerifyTab /> : adminView === "access" ? <AccessTab /> : (
      <PipelineTab refresh={refresh} />
      )}
    </div>
  );
}

/* ───────── shell ───────── */
const NAV = [{ id: "dash", label: "종합 현황" }, { id: "detail", label: "매장 진단" }, { id: "schedule", label: "스케줄 진단" }, { id: "helper", label: "인건비 직접 보정" }, { id: "sim", label: "차주 시뮬레이션" }, { id: "upload", label: "관리자" }];
function AppInner({ myCode }) {
  const [view, setView] = useState("dash");
  const [code, setCode] = useState(myCode || null);
  const [, bump] = React.useReducer((x) => x + 1, 0);
  const goDetail = (c) => { setCode(c); setView("detail"); };
  const navActive = (id) => view === id || (id === "detail" && view === "staff");
  return (
    <div className="app">
      <style>{CSS}</style>
      <header>
        <div className="brand"><span className="brand-mark">A</span><div><div className="brand-name">애슐리 인건비 모니터</div><div className="brand-sub">{CUR_WEEK_LABEL}</div></div></div>
        <nav>{NAV.map((n) => <button key={n.id} className={navActive(n.id) ? "on" : ""} onClick={() => setView(n.id)}>{n.label}</button>)}</nav>
      </header>
      <main className={view === "schedule" ? "schedule-main" : ""}>
        {view === "dash" && <Dashboard goDetail={goDetail} />}
        {view === "schedule" && <ScheduleDiagnosisFrame />}
        {view === "detail" && <Detail code={code || myCode} setCode={setCode} goStaff={() => setView("staff")} refresh={bump} />}
        {view === "staff" && <StaffTab code={code} onBack={() => setView("detail")} />}
        {view === "sim" && <SimTab />}
        {view === "helper" && <Helper refresh={bump} myCode={myCode} />}
        {view === "upload" && <AdminTab refresh={bump} />}
      </main>
    </div>
  );
}

const CSS = `
:root{--paper:#F3F5F2;--surface:#FFFFFF;--ink:#182027;--muted:#68737E;--line:#E1E5DE;--over:#C2402A;--over-soft:#F9E9E5;--under:#1E7A5B;--under-soft:#E5F2EC;--warn:#B5751A;--warn-soft:#FBF0DC;--cobalt:#24478F;--cobalt-soft:#E8EEF9;}
*{box-sizing:border-box;margin:0}
.app{min-height:100vh;background:var(--paper);color:var(--ink);font-family:-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Pretendard","Malgun Gothic",sans-serif;font-size:14px;line-height:1.5}
header{display:flex;justify-content:space-between;align-items:center;gap:16px;flex-wrap:wrap;padding:14px 24px;background:var(--surface);border-bottom:1px solid var(--line);position:sticky;top:0;z-index:20}
.brand{display:flex;gap:12px;align-items:center}
.brand-mark{width:36px;height:36px;border-radius:9px;background:var(--ink);color:#fff;display:flex;align-items:center;justify-content:center;font-weight:800;font-size:18px}
.brand-name{font-weight:800;letter-spacing:-.02em;font-size:16px}.brand-sub{color:var(--muted);font-size:12px}
nav{display:flex;gap:4px}
nav button{border:0;background:transparent;padding:8px 14px;border-radius:8px;font-size:14px;font-weight:600;color:var(--muted);cursor:pointer}
nav button:hover{background:var(--paper)}nav button.on{background:var(--ink);color:#fff}
nav button:focus-visible,button.primary:focus-visible{outline:2px solid var(--cobalt);outline-offset:2px}
main{max-width:1240px;margin:0 auto;padding:24px}
main.schedule-main{max-width:none;margin:0;padding:0;background:#F6F7FB}
.schedule-diagnosis-shell{width:100%;height:calc(100vh - 66px);min-height:760px;background:#F6F7FB}
.schedule-diagnosis-frame{display:block;width:100%;height:100%;border:0;background:#F6F7FB}
.statgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:12px;margin-bottom:20px}
.statgrid.three{grid-template-columns:repeat(3,1fr)}
@media(max-width:720px){.statgrid.three{grid-template-columns:1fr}}
.stat{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.stat-label{font-size:12px;color:var(--muted);margin-bottom:4px}
.stat-value{font-size:24px;font-weight:800;letter-spacing:-.02em;font-variant-numeric:tabular-nums}
.stat-sub{font-size:12px;color:var(--muted);margin-top:2px}
.tone-over .stat-value{color:var(--over)}.tone-under .stat-value{color:var(--under)}.tone-warn .stat-value{color:var(--warn)}
.toolbar{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:14px;align-items:center}
select,input{border:1px solid var(--line);background:var(--surface);border-radius:8px;padding:8px 10px;font-size:13px;font-family:inherit;color:var(--ink)}
select:focus,input:focus{outline:2px solid var(--cobalt);outline-offset:0;border-color:var(--cobalt)}
.count{color:var(--muted);font-size:12px;margin-left:auto}
.pill{background:var(--cobalt-soft);color:var(--cobalt);border-radius:99px;padding:4px 10px;font-size:12px;font-weight:600;white-space:nowrap}
.tablewrap{background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:auto}
table{width:100%;border-collapse:collapse;font-size:13px}
th{position:sticky;top:0;background:var(--surface);text-align:left;font-size:11px;color:var(--muted);font-weight:600;padding:9px 12px;border-bottom:1px solid var(--line);white-space:nowrap;z-index:2}
th div{line-height:1.25}
td{padding:9px 12px;border-bottom:1px solid var(--line);white-space:nowrap;vertical-align:middle}
tbody tr:last-child td{border-bottom:0}
tbody tr{cursor:pointer}tbody tr:hover{background:#FAFBF8}
.helperlayout tbody tr,.compare tbody tr{cursor:default}
.wide{font-size:11.5px}
.wide th{padding:6px 7px;font-size:10px}
.wide td{padding:6px 7px}
.wide td.num,.wide th.num{font-variant-numeric:tabular-nums}
table.wide th.stick,table.wide td.stick{position:sticky;left:0;background:var(--surface);z-index:3;box-shadow:1px 0 0 var(--line)}
table.wide tbody tr:hover td.stick{background:#FAFBF8}
table.wide tbody tr.pg-stripe td.stick{background:#F7F3EA}
table.wide tbody tr.pg-stripe:hover td.stick{background:#F1EBDD}
th.stick{z-index:5}
td.soft,.wide td.soft{color:var(--muted);font-weight:400}
tr.pg-stripe td{background:#F7F3EA}
tr.pg-stripe:hover td{background:#F1EBDD}
.wide th.refstart,.wide td.refstart{border-left:2px solid var(--line)}
.wide td.soft{background:#FBFAF7}
tr.pg-stripe td.soft{background:#F5F0E4}
th.has-tip{cursor:help;position:sticky;top:0}
.th-info{color:var(--cobalt);font-size:10px;margin-left:3px;opacity:.7}
.th-pop{display:none;position:absolute;top:100%;left:0;z-index:30;width:220px;white-space:normal;background:var(--ink);color:#fff;font-size:11px;font-weight:400;line-height:1.5;padding:9px 11px;border-radius:8px;box-shadow:0 6px 18px rgba(0,0,0,.22);margin-top:2px}
th.has-tip:hover .th-pop{display:block}
.num{text-align:right;font-variant-numeric:tabular-nums}
.strong{font-weight:700}.c-over{color:var(--over)}.c-under{color:var(--under)}
.storename{font-weight:600;display:block}.storecode{font-size:11px;color:var(--muted)}
.cell,.band{color:var(--muted);font-size:12px}
.blank .blankcell{color:#B4BBB2;font-size:11px;font-style:italic}
.gapcol{min-width:190px}
.gapbar{position:relative;display:inline-block;vertical-align:middle;height:12px;background:#F0F2EE;border-radius:6px;overflow:hidden;margin-right:8px}
.gapbar-axis{position:absolute;left:50%;top:0;bottom:0;width:1px;background:#C6CCC2}
.gapbar-fill{position:absolute;top:2px;bottom:2px;border-radius:4px}
.gapbar-fill.over{background:var(--over)}.gapbar-fill.under{background:var(--under)}
.gapnum{font-size:12px;font-variant-numeric:tabular-nums;font-weight:600}
.hint{color:var(--muted);font-size:12px;margin-top:10px}
.sectionhead{font-size:15px;font-weight:800;letter-spacing:-.01em;margin:26px 0 12px}
.sectionhead span{font-weight:400;font-size:12px;color:var(--muted);margin-left:8px}
.notice{background:var(--surface);border:1px dashed var(--line);border-radius:12px;padding:20px;color:var(--muted)}
.detailgrid{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:12px}
.dbox{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.dbox h4{font-size:12px;color:var(--muted);margin-bottom:10px;font-weight:700}
.dbox dl{display:grid;gap:7px}
.dbox dl>div{display:flex;justify-content:space-between;align-items:baseline}
.dbox dl .staffpop{display:none}
.dbox dl .staffpop .chips{display:flex}
.dbox dl .has-staffpop:hover .staffpop{display:block}
.dbox dl .wagepop{display:none}
.dbox dl .has-wagepop:hover .wagepop,.dbox dl .has-wagepop.tap-open .wagepop{display:block}
.has-wagepop{position:relative}
.has-wagepop dt{cursor:help}
.has-wagepop .th-info{font-size:10px;color:var(--cobalt);opacity:.7;margin-left:2px}
.wagepop{position:absolute;bottom:100%;left:0;z-index:40;width:min(250px, calc(100vw - 52px));background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.18);padding:13px 14px;margin-bottom:4px}
.wagepop-title{font-weight:800;font-size:13px;margin-bottom:9px}
.wagepop-title span{font-weight:400;font-size:11px;color:var(--muted)}
.wagepop-row{display:flex;justify-content:space-between;align-items:center;font-size:12px;padding:4px 0}
.wagepop-row b{font-variant-numeric:tabular-nums}
.wagepop-row.total{border-top:1px solid var(--line);margin-top:4px;padding-top:7px;font-weight:800}
.wagepop-note{font-size:10px;color:var(--muted);margin-top:8px;line-height:1.5}
.dbox dt{color:var(--muted);font-size:13px}.dbox dd{font-weight:700;font-variant-numeric:tabular-nums}
.dbox .sumline{border-top:1px solid var(--line);padding-top:7px;margin-top:1px}
.slotwrap{display:grid;grid-template-columns:repeat(5,1fr);gap:10px}
@media(max-width:860px){.slotwrap{grid-template-columns:repeat(2,1fr)}}
.slot{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:12px}
.slot-head{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:10px}
.slot-name{font-weight:800;font-size:14px}.slot-time{font-size:11px;color:var(--muted)}
.slot-bars{display:grid;gap:5px}
.slot-row{display:grid;grid-template-columns:32px 1fr auto;align-items:center;gap:6px;font-size:12px}
.slot-tag{color:var(--ink);font-weight:600}.slot-tag.muted,.slot-row .muted{color:var(--muted);font-weight:400}
.bartrack{background:#F0F2EE;border-radius:4px;height:9px;overflow:hidden}
.bar{display:block;height:100%;border-radius:4px}
.bar.ok{background:var(--under)}.bar.over{background:var(--over)}.bar.peer{background:#C6CCC2}
.slot-row b{font-variant-numeric:tabular-nums;font-size:12px}
.slot-foot{margin-top:9px;font-size:11px;font-weight:600;font-variant-numeric:tabular-nums}
.pgcol{text-align:center}
.pgtag{display:inline-flex;min-width:22px;justify-content:center;background:var(--paper);border:1px solid var(--line);border-radius:6px;padding:2px 6px;font-size:12px;font-weight:700}
.staffwrap{display:grid;grid-template-columns:1fr 220px;gap:12px}
@media(max-width:720px){.staffwrap{grid-template-columns:1fr}}
.staffroster,.staffidx{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.staff-top{display:flex;justify-content:space-between;align-items:center;margin-bottom:12px;gap:10px}
.staff-total{font-size:22px;font-weight:800;letter-spacing:-.02em}
.staff-sub{color:var(--muted);font-size:12px;margin-left:8px}
button.ghost{border:1px solid var(--line);background:var(--surface);color:var(--cobalt);border-radius:8px;padding:7px 11px;font-size:12px;font-weight:600;cursor:pointer;white-space:nowrap}
button.ghost:hover{background:var(--cobalt-soft)}
button.ghost:focus-visible{outline:2px solid var(--cobalt);outline-offset:2px}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chip{background:var(--paper);border:1px solid var(--line);border-radius:8px;padding:5px 9px;font-size:12px;display:inline-flex;gap:5px;align-items:baseline}
.chip b{font-weight:700}
.staffidx dl{display:grid;gap:8px}
.staffidx dl div{display:flex;justify-content:space-between;align-items:baseline}
.staffidx dt{color:var(--muted);font-size:13px}.staffidx dd{font-weight:700;font-variant-numeric:tabular-nums}
.staffidx .sumline{border-top:1px solid var(--line);padding-top:8px}
.staffcols{display:grid;grid-template-columns:1fr 320px;gap:20px;align-items:start}
@media(max-width:860px){.staffcols{grid-template-columns:1fr}}
.formrow.two{grid-template-columns:1fr 1fr}
.pill.warn{background:var(--warn-soft);color:var(--warn)}
.peerset{display:inline-flex;align-items:center;gap:5px;background:var(--cobalt-soft);color:var(--cobalt);border-radius:99px;padding:2px 6px 2px 10px;font-size:12px;font-weight:600}
.peerset select{border:1px solid var(--cobalt);background:#fff;color:var(--cobalt);border-radius:99px;padding:3px 8px;font-size:12px;font-weight:700}
.muted{color:var(--muted)}
.guide-btn{margin-left:auto;font-weight:700}
.simhead{display:flex;align-items:stretch;gap:12px;margin-bottom:22px;flex-wrap:wrap}
.simhead-box{flex:1;min-width:180px;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px}
.simhead-box.proj{border-width:2px}
.simhead-box.proj.up{border-color:var(--over);background:var(--over-soft)}
.simhead-box.proj.down{border-color:var(--under);background:var(--under-soft)}
.simhead-arrow{display:flex;align-items:center;color:var(--muted);font-size:22px}
.simhead-label{font-size:12px;color:var(--muted);margin-bottom:4px}
.simhead-val{font-size:30px;font-weight:800;letter-spacing:-.02em;font-variant-numeric:tabular-nums}
.simhead-sub{font-size:12px;color:var(--muted);margin-top:3px}
@media(max-width:720px){.simhead-arrow{display:none}}
.simcols{display:grid;grid-template-columns:1fr 1fr;gap:20px;align-items:start}
@media(max-width:860px){.simcols{grid-template-columns:1fr}}
.levers{display:grid;gap:14px;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px}
.lever-h{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:6px}
.lever-label{font-weight:600;font-size:13px}
.lever-val{font-weight:800;font-variant-numeric:tabular-nums;font-size:14px;color:var(--cobalt)}
.lever input[type=range]{width:100%;accent-color:var(--cobalt);cursor:pointer}
.lever-hint{font-size:11px;color:var(--muted);margin-top:4px}
.lever-real{font-size:12px;color:var(--cobalt);font-weight:600;margin-top:5px;background:var(--cobalt-soft);border-radius:7px;padding:6px 9px}
.lever-real b{font-weight:800}
.lever-input{display:inline-flex;align-items:center;gap:3px}
.lever-input input{width:96px;text-align:right;padding:4px 7px;font-size:13px;font-weight:700;font-variant-numeric:tabular-nums;border:1px solid var(--line);border-radius:6px;color:var(--cobalt)}
.lever-input input:focus{outline:2px solid var(--cobalt);border-color:var(--cobalt)}
.lever-input .unit{font-size:11px;color:var(--muted);font-weight:600}
.cover-lever .lever-h{margin-bottom:2px}
.cover-total{font-size:13px;font-weight:800;color:var(--cobalt);font-variant-numeric:tabular-nums}
.coverlist{display:grid;gap:8px;margin-top:8px}
.coverrow{display:flex;align-items:center;gap:6px;flex-wrap:wrap}
.coverrow select{border:1px solid var(--line);border-radius:6px;padding:5px 7px;font-size:12px}
.coverrow input{width:56px;text-align:right;border:1px solid var(--line);border-radius:6px;padding:5px 7px;font-size:12px;font-variant-numeric:tabular-nums}
.cover-unit{font-size:11px;color:var(--muted)}
.cover-amt{font-size:11px;color:#3D4650;font-variant-numeric:tabular-nums;margin-left:auto;white-space:nowrap}
.cover-add{border:1px dashed var(--cobalt);color:var(--cobalt);border-radius:8px;padding:7px;font-size:12px;font-weight:600;cursor:pointer;background:transparent;width:100%}
.cover-add:hover{background:var(--cobalt-soft)}
.sumrow td{border-top:2px solid var(--line);font-weight:700}
.simbars{display:grid;gap:8px;margin-top:14px;background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.simbar-row{display:grid;grid-template-columns:36px 1fr auto;align-items:center;gap:10px;font-size:13px}
.simbar-tag{color:var(--muted);font-weight:600}
.bartrack.big{height:14px}
.simbar-row b{font-variant-numeric:tabular-nums}
.simhead-box.goal-met{border-color:var(--under);background:var(--under-soft)}
.simhead-box.goal-met .simhead-sub{color:var(--under);font-weight:700}
.drivertip{display:flex;gap:12px;align-items:flex-start;background:var(--warn-soft);border:1px solid var(--warn);border-radius:12px;padding:12px 14px;margin-bottom:20px;font-size:13px;color:#5A4212}
.drivertip-tag{background:var(--warn);color:#fff;border-radius:6px;padding:3px 9px;font-size:11px;font-weight:700;white-space:nowrap}
.drivertip b{font-weight:800}
.dropzone{display:block;background:var(--surface);border:2px dashed var(--cobalt);border-radius:14px;padding:26px 20px;text-align:center;cursor:pointer;transition:background .15s}
.dropzone:hover{background:var(--cobalt-soft)}
.dropzone-big{font-weight:800;font-size:16px;color:var(--cobalt)}
.dropzone-sub{font-size:12px;color:var(--muted);margin-top:5px}
.filecard.static{cursor:default}
.filecard.ready{border-color:var(--under)}
.filecard.ready .filecard-status{color:var(--under);font-weight:700}
.loglist{background:var(--surface);border:1px solid var(--line);border-radius:12px;overflow:hidden}
.logrow{display:flex;gap:12px;padding:10px 14px;border-bottom:1px solid var(--line);font-size:13px;align-items:baseline}
.logrow:last-child{border-bottom:0}
.logrow.ok{border-left:4px solid var(--under)}
.logrow.warn{border-left:4px solid var(--warn)}
.logrow.err{border-left:4px solid var(--over)}
.log-at{color:var(--muted);font-size:11px;font-variant-numeric:tabular-nums;white-space:nowrap}
.pwgate{display:flex;justify-content:center;padding:48px 16px}
.pwcard{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:32px 28px;max-width:360px;width:100%;text-align:center}
.pwlock{font-size:32px;margin-bottom:8px}
.pwcard h3{font-size:18px;font-weight:800;margin-bottom:6px}
.pwcard p{font-size:13px;color:var(--muted);margin-bottom:16px}
.pwcard input{width:100%;padding:11px 12px;border:1px solid var(--line);border-radius:9px;font-size:14px;text-align:center;margin-bottom:10px}
.pwcard input:focus{outline:2px solid var(--cobalt);border-color:var(--cobalt)}
.pwcard button.primary{width:100%}
.pwerr{color:var(--over);font-size:12px;font-weight:600;margin-bottom:8px}
.pwhint{font-size:11px;color:var(--muted);margin-top:12px}
.wide-dbox{grid-column:1/-1}
.dbox-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:8px}
.dbox-head h4{margin:0}
.report-btn{white-space:nowrap}
.hktable{width:100%;border-collapse:collapse;font-size:13px}
.hktable th,.hktable td{padding:7px 10px;border-bottom:1px solid var(--line);text-align:left}
.hktable th.num,.hktable td.num{text-align:right;font-variant-numeric:tabular-nums}
.hktable td.dim{color:var(--muted-2,#B9C0C7)}
.hktable .dimtxt{font-size:11px;color:var(--muted);font-weight:400}
.hktable tr.subline td{background:#F7FAF9;font-weight:600}
.hktable tr.sumrow td{border-top:2px solid var(--ink);border-bottom:none;font-weight:800;padding-top:9px}
.hktable td .th-info{color:var(--cobalt);font-size:11px;cursor:help}
.hint.mini{font-size:11px;margin-top:8px}
tr.has-mtip,tr.has-staffpop2{position:relative}
.mtip,.staffpop2{display:none;position:absolute;left:10px;top:100%;z-index:30;background:var(--ink);color:#fff;padding:9px 11px;border-radius:8px;font-size:12px;min-width:240px;max-width:min(320px, calc(100vw - 52px));box-shadow:0 8px 24px rgba(0,0,0,.22)}
tr.has-mtip:hover .mtip,tr.has-staffpop2:hover .staffpop2,tr.has-mtip.tap-open .mtip,tr.has-staffpop2.tap-open .staffpop2{display:block}
.mtip-t{font-weight:700;margin-bottom:3px}.mtip-b{color:#D4DAE0;line-height:1.5}
.staffpop2{background:var(--surface);color:var(--ink);border:1px solid var(--line);min-width:220px}
.staffpop2 .staffpop-title{font-weight:700;font-size:12px;margin-bottom:6px}
.staffpop2 .chips{display:flex;flex-wrap:wrap;gap:4px;margin-bottom:8px}
.modal-overlay{position:fixed;inset:0;background:rgba(20,28,26,.5);display:flex;align-items:center;justify-content:center;z-index:100;padding:20px}
.modal-panel{background:var(--paper);border-radius:14px;max-width:1100px;width:100%;max-height:88vh;display:flex;flex-direction:column;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal-head{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;border-bottom:1px solid var(--line)}
.modal-head h3{margin:0;font-size:16px}
.modal-sub{font-size:12px;color:var(--muted);font-weight:400;margin-left:8px}
.modal-x{background:none;border:none;font-size:18px;cursor:pointer;color:var(--muted);padding:4px 8px}
.modal-body{overflow:auto;padding:16px 20px}
.reporttable{width:100%;border-collapse:collapse;font-size:12px;white-space:nowrap}
.reporttable th,.reporttable td{padding:6px 9px;border-bottom:1px solid var(--line);text-align:left}
.reporttable th.num,.reporttable td.num{text-align:right;font-variant-numeric:tabular-nums}
.reporttable th.stickc,.reporttable td.stickc{position:sticky;left:0;background:var(--paper);z-index:2}
.reporttable td.refstart,.reporttable th.refstart{border-left:2px solid var(--line)}
.reporttable td.small{font-size:11px}
.reporttable tr.me-tr td{background:var(--cobalt-soft)}
.reporttable tr.me-tr td.stickc{background:#DCE6F5}
.me-dot{font-size:9px;font-weight:700;color:#fff;background:var(--cobalt);border-radius:4px;padding:1px 4px;margin-left:5px}
.opt-col{color:var(--cobalt);background:var(--cobalt-soft)}
.cell-d{font-size:11px;margin-left:6px;font-weight:600}
.ver-memo{font-size:11px;font-weight:700;color:var(--warn);background:var(--warn-soft);border-radius:6px;padding:2px 8px;white-space:nowrap}
.optguide{margin-top:12px;background:var(--cobalt-soft);border:1px solid var(--cobalt);border-radius:10px;padding:12px 14px}
.optguide-h{display:flex;align-items:center;justify-content:space-between;gap:8px;margin-bottom:6px}
.optguide-tag{font-size:12px;font-weight:800;color:var(--cobalt)}
.optguide-b{font-size:13px;line-height:1.6;color:var(--ink)}
.simbar-tag.opt{background:var(--cobalt);color:#fff}
.bar.opt{background:var(--cobalt)}
.holiday-lever .lever-real.dimtxt{color:var(--muted)}
.dimtxt{color:var(--muted);font-size:11px}
.mystore-bar{display:flex;align-items:center;gap:10px;background:var(--cobalt-soft);border:1px solid var(--cobalt);border-radius:10px;padding:10px 14px;margin-bottom:14px}
.mystore-label{font-size:12px;font-weight:800;color:var(--cobalt)}
.mystore-hint{font-size:12px;color:var(--muted)}
tr.mine-row td{background:#F0F5FC}
.mine-tag{font-size:9px;font-weight:700;color:#fff;background:var(--cobalt);border-radius:4px;padding:1px 5px;margin-left:5px}
.blind{color:var(--muted);font-style:italic;font-size:12px}
.pill.wait{background:var(--warn-soft);color:var(--warn-ink,#9A6A1E)}
.weekbox{background:var(--cobalt-soft);border:1px solid var(--cobalt);border-radius:9px;padding:9px 11px;margin:8px 0}
.weekbox-h{font-size:12px;font-weight:700;color:var(--cobalt);margin-bottom:5px}
.weekbox-h span{font-weight:400;color:var(--muted);margin-left:6px}
.weektable{width:100%;border-collapse:collapse;font-size:12px}
.weektable td{padding:3px 4px;border-bottom:1px solid #DCE6F5}
.weektable td.wt-sub{color:var(--muted);font-size:11px}
.weektable td.num{text-align:right;font-weight:600}
.wt-cell{font-size:11px}
.wt-range{display:block;font-weight:600}
.wt-weeks{display:block;color:var(--muted);font-size:10px;margin-top:1px}
.row-dirty td{background:#FFF7E8}
.row-fail td{background:#FDECEA}
.row-dupe td{background:#FFF4E5}
.dirty-dot{color:var(--warn);margin-left:6px;font-size:10px}
.dirty-badge{font-size:11px;font-weight:700;color:var(--warn-ink,#9A6A1E);background:var(--warn-soft);border-radius:6px;padding:2px 8px}
.peer-msg{font-size:12px;color:var(--under);font-weight:600}
.fc-input{width:110px;padding:5px 8px;border:1px solid var(--line);border-radius:7px;font-size:13px;text-align:right;font-variant-numeric:tabular-nums}
.pipe-week{margin:8px 0}
.pipe-week input{margin-left:8px;padding:7px 10px;border:1px solid var(--line);border-radius:8px;font-size:14px}
.pipe-files{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin:12px 0}
.pipe-file{display:flex;align-items:center;gap:8px;padding:8px 10px;border:1px solid var(--line);border-radius:8px;font-size:13px;background:#FBFAF7}
.pipe-file.ok{border-color:var(--under);background:var(--under-bg,#EAF5F0)}
.pf-check{font-weight:700;color:var(--muted)}.pipe-file.ok .pf-check{color:var(--under)}
.pf-label{font-weight:600;min-width:130px}
.pf-name{color:var(--muted);font-size:12px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;flex:1;min-width:0}
.pf-remove{margin-left:auto;border:none;background:transparent;color:var(--muted);cursor:pointer;font-size:14px;padding:2px 6px;border-radius:6px;flex-shrink:0}
.pf-remove:hover{background:#F0DCDC;color:var(--over)}
.pf-manual{margin-left:auto;flex-shrink:0;cursor:pointer;color:var(--cobalt,#24478F);font-weight:800;font-size:15px;padding:2px 7px;border:1px solid var(--cobalt,#24478F);border-radius:6px;line-height:1}
.pf-manual:hover{background:var(--cobalt,#24478F);color:#fff}
.pf-manual + .pf-remove{margin-left:6px}
.pf-clearall{margin:2px 0 10px;border:1px solid var(--line);background:#fff;color:var(--muted);border-radius:7px;padding:5px 12px;font-size:12px;cursor:pointer}
.datasource{display:flex;align-items:center;gap:10px;flex-wrap:wrap;padding:9px 13px;background:var(--cobalt-soft,#EAF0FA);border:1px solid var(--cobalt,#24478F);border-radius:9px;margin-bottom:12px;font-size:12px}
.ds-label{font-weight:700;color:var(--cobalt,#24478F)}
.ds-week{font-weight:600}
.ds-file{color:var(--muted);font-family:monospace}
.ds-at{color:var(--muted);margin-left:auto}
.pipe-actions{display:flex;gap:8px;margin:12px 0}
.pipe-msg{padding:8px 12px;background:#FCF4E4;border-radius:8px;font-size:13px;margin:8px 0}
.pipe-report{margin-top:14px;border:1px solid var(--line);border-radius:10px;padding:14px}
.pipe-report h4{margin:0 0 8px}
.pipe-report-row{font-size:13px;color:var(--muted);margin-bottom:8px}
.pipe-table{width:100%;border-collapse:collapse;font-size:13px}
.pipe-table th,.pipe-table td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left}
.pipe-table th.num,.pipe-table td.num{text-align:right;font-variant-numeric:tabular-nums}
.pipe-warn{margin-top:10px;padding:8px 12px;background:#FBE7E4;color:var(--over);border-radius:8px;font-size:13px}
.pipe-ok{margin-top:10px;padding:8px 12px;background:var(--under-bg,#EAF5F0);color:var(--under);border-radius:8px;font-size:13px;font-weight:600}
.approve-btn{color:var(--under);border-color:var(--under)}
.allowbox{margin-top:16px;border-top:1px dashed var(--line);padding-top:12px}
.allow-h{margin:0 0 8px}.allow-h span{font-size:11px;color:var(--muted);font-weight:400}
.allowtable tr.ot-row td{background:#FBF6EE}
.ot-badge{font-size:9px;font-weight:700;color:#fff;background:var(--warn);border-radius:4px;padding:1px 5px;margin-left:6px}
.schedbox{margin-top:12px;background:#FBF6EE;border:1px solid var(--warn);border-radius:10px;padding:12px 14px}
.schedbox-h{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px}
.schedbox-tag{font-size:12px;font-weight:800;color:var(--warn-ink,#9A6A1E)}
.schedlist{margin:0;padding-left:18px;font-size:13px;line-height:1.6}
.schedlist li{margin-bottom:4px}
.cmp-toolbar{display:flex;align-items:center;gap:10px;margin-bottom:10px;flex-wrap:wrap}
.cmp-label{font-size:12px;font-weight:700;color:var(--muted)}
.cmp-hint{font-size:11px;color:var(--muted)}
.cmp1v1 .coltag{display:block;font-size:9px;font-weight:700;margin-top:2px;color:var(--muted)}
.cmp1v1 .coltag.me{color:var(--cobalt)}
.cmp1v1 td.soft,.cmp1v1 th.soft{color:var(--muted)}
.cmp1v1 .diff-col{background:#FAFAF8;font-weight:600}
.cmp1v1 th.diff-col{background:none}
.allow-head{display:flex;align-items:center;justify-content:space-between;gap:10px;margin-bottom:6px}
.allow-head .allow-h{margin:0}
.viewsel,.cmp-toolbar select{padding:4px 8px;border:1px solid var(--line);border-radius:7px;font-size:13px}
.toggle-sep{display:inline-block;width:1px;height:20px;background:var(--line);margin:0 6px;vertical-align:middle}
.tip-pop-wrap{position:relative;display:inline-block}
.tip-pop-btn{color:var(--cobalt);border-color:var(--cobalt)}
.tip-pop{display:none;position:absolute;left:0;top:100%;margin-top:6px;z-index:40;background:var(--surface);border:1px solid var(--line);border-radius:10px;padding:12px;width:360px;max-height:420px;overflow:auto;box-shadow:0 12px 32px rgba(0,0,0,.18)}
.tip-pop-wrap:hover .tip-pop{display:block}
.tip-pop-item{padding:8px 0;border-bottom:1px solid var(--line)}
.tip-pop-item:last-child{border-bottom:none}
.tip-pop-head{font-size:13px;font-weight:700;margin-bottom:5px;display:flex;align-items:center;gap:6px}
.tip-pop-head .tip-tag{font-size:10px;padding:1px 6px;border-radius:5px;color:#fff}
.tip-tag.lv-high{background:var(--over)}.tip-tag.lv-mid{background:var(--warn)}.tip-tag.lv-good{background:var(--under)}
.tip-pop-alts{margin:0;padding-left:18px;font-size:12px;line-height:1.6;color:var(--muted)}
.tip-pop-empty{font-size:13px;color:var(--muted);padding:6px}
.pwform{max-width:360px;display:grid;gap:12px}
.pwform label{display:grid;gap:5px;font-size:13px;font-weight:600;color:var(--muted)}
.pwform input{padding:10px 12px;border:1px solid var(--line);border-radius:9px;font-size:14px}
.pwform input:focus{outline:2px solid var(--cobalt);border-color:var(--cobalt)}
.pwok{color:var(--under);font-size:12px;font-weight:600}
.uploadgrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px}
.uploadcard{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px;display:flex;flex-direction:column}
.uploadcard.ready{border-color:var(--under)}
.uploadcard-head{display:flex;justify-content:space-between;align-items:baseline;gap:8px;margin-bottom:10px}
.uploadcard-name{font-weight:800;font-size:14px}
.uploadcard-status{font-size:11px;color:var(--muted);white-space:nowrap}
.uploadcard.ready .uploadcard-status{color:var(--under);font-weight:700}
.uploadlist{display:grid;gap:5px;min-height:44px}
.upload-empty{color:#B4BBB2;font-size:12px;font-style:italic;padding:10px 0}
.upload-row{display:flex;gap:10px;align-items:baseline;font-size:12px;padding:5px 8px;border-radius:7px;background:var(--paper)}
.upload-row.latest{background:var(--under-soft)}
.upload-at{color:var(--muted);font-variant-numeric:tabular-nums;white-space:nowrap;font-size:11px}
.upload-file{font-weight:600;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.uploadcard-foot{margin-top:8px;padding-top:8px;border-top:1px dashed var(--line);font-size:11px;color:var(--muted)}
.subnav{display:flex;gap:4px;margin-bottom:18px;border-bottom:1px solid var(--line)}
.subnav button{border:0;background:transparent;padding:9px 16px;font-size:14px;font-weight:700;color:var(--muted);cursor:pointer;border-bottom:2px solid transparent;margin-bottom:-1px}
.subnav button:hover{color:var(--ink)}
.subnav button.on{color:var(--ink);border-bottom-color:var(--ink)}
.pgselect{border:1px solid var(--cobalt);background:#fff;color:var(--cobalt);border-radius:7px;padding:4px 8px;font-size:13px;font-weight:700}
.pill.dim{background:var(--paper);color:var(--muted)}
.pill.live{background:var(--under-soft);color:var(--under)}
button.ghost.tiny{padding:4px 8px;font-size:11px}
.adjline{margin-top:9px;padding-top:8px;border-top:1px dashed var(--line);font-size:11px;color:var(--cobalt);font-weight:600}
.weeknote{background:var(--cobalt-soft);color:var(--cobalt);border-radius:10px;padding:9px 14px;font-size:13px;margin-bottom:16px}
.weeknote b{font-weight:800}
.daytoggle{display:inline-flex;gap:2px;background:var(--paper);border:1px solid var(--line);border-radius:9px;padding:3px;margin-bottom:12px}
.daytoggle button{border:0;background:transparent;padding:6px 18px;font-size:13px;font-weight:700;color:var(--muted);border-radius:7px;cursor:pointer}
.daytoggle button.on{background:#fff;color:var(--ink);box-shadow:0 1px 3px rgba(0,0,0,.12)}
.periodtoggle{display:inline-flex;gap:2px;background:var(--paper);border:1px solid var(--line);border-radius:9px;padding:3px}
.periodtoggle button{border:0;background:transparent;padding:6px 16px;font-size:13px;font-weight:700;color:var(--muted);border-radius:7px;cursor:pointer}
.periodtoggle button.on{background:#fff;color:var(--ink);box-shadow:0 1px 3px rgba(0,0,0,.12)}
.adjsum{margin-top:6px}
.adjsum .sumline dt{font-weight:700}
.stat.has-gaptip{position:relative;cursor:help}
.stat.has-gaptip .th-info{font-size:10px;color:var(--cobalt);opacity:.7;margin-left:2px}
.gaptip{display:none;position:absolute;top:100%;left:0;z-index:40;width:min(300px, calc(100vw - 32px));background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.18);padding:14px 15px;margin-top:6px;text-align:left}
.stat.has-gaptip:hover .gaptip,.stat.has-gaptip.tap-open .gaptip{display:block}
.gaptip-title{font-weight:800;font-size:13px;margin-bottom:3px}
.gaptip-desc{font-size:11px;color:var(--muted);margin-bottom:10px;line-height:1.5}
.gaptip-row{display:flex;align-items:center;gap:8px;margin-bottom:7px}
.peerlist-inline{font-size:11.5px;color:var(--muted);margin-top:4px;line-height:1.5}
.gaptip-label{font-size:12px;font-weight:600;width:46px;flex-shrink:0}
.gaptip-row .gapnum{font-size:12px;width:52px;text-align:right;flex-shrink:0}
.gaptip-line{display:flex;justify-content:space-between;align-items:center;font-size:12px;padding:4px 0}
.gaptip-line.total{border-top:1px solid var(--line);margin-top:4px;padding-top:7px;font-weight:800}
.gaptip-line b{font-variant-numeric:tabular-nums}
.has-staffpop{position:relative}
.has-staffpop dt{cursor:help}
.has-staffpop .th-info{font-size:10px;color:var(--cobalt);opacity:.7;margin-left:2px}
.staffpop{position:absolute;top:100%;left:0;z-index:40;width:280px;background:#fff;border:1px solid var(--line);border-radius:12px;box-shadow:0 10px 30px rgba(0,0,0,.18);padding:13px 14px}
.staffpop-title{font-weight:800;font-size:13px;margin-bottom:9px}
.staffpop-title span{font-weight:400;font-size:11px;color:var(--muted)}
.staffpop-btn{margin-top:10px}
.stat.rating.good{border-color:var(--under);background:var(--under-soft)}
.stat.rating.bad{border-color:var(--over);background:var(--over-soft)}
.rating-hint{font-weight:400;font-size:10px;color:var(--muted);margin-left:4px}
.rating-star{color:#E0A93B;font-size:16px;margin-left:3px}
.trendbox{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.trend-top{display:flex;align-items:center;gap:12px;margin-bottom:8px;flex-wrap:wrap}
.trend-now{display:flex;flex-direction:column;line-height:1.2}
.trend-now-val{font-size:20px;font-weight:800;font-variant-numeric:tabular-nums}
.trend-now-d{font-size:12px;font-weight:600}
.trend-badge{margin-left:auto;font-size:11px;font-weight:700;color:var(--warn);background:var(--warn-soft);border-radius:6px;padding:3px 8px}
.wagebars{display:grid;gap:12px;margin:14px 0 10px}
.wagebar-row{display:grid;gap:5px}
.wagebar-name{font-size:13px;font-weight:700;display:flex;align-items:center;gap:6px}
.me-tag{font-size:10px;font-weight:700;color:#fff;background:var(--cobalt);border-radius:5px;padding:1px 6px}
.wagebar-track{position:relative;display:flex;height:26px;border-radius:6px;overflow:hidden;background:var(--paper)}
.wagebar-seg{height:100%}
.wagebar-total{position:absolute;right:8px;top:50%;transform:translateY(-50%);font-size:12px;font-weight:800;color:var(--ink);font-variant-numeric:tabular-nums;text-shadow:0 0 4px #fff,0 0 4px #fff}
.wage-legend{display:flex;flex-wrap:wrap;gap:10px;margin:8px 0}
.wl-item{display:flex;align-items:center;gap:5px;font-size:11px;color:var(--muted)}
.wl-box{width:11px;height:11px;border-radius:3px}
.ratingbars{display:grid;gap:10px;margin:14px 0 10px}
.rating-row2{display:grid;grid-template-columns:130px 1fr 60px;align-items:center;gap:10px}
.rating-name{font-size:13px;font-weight:700;display:flex;align-items:center;gap:6px}
.rating-track{height:20px;background:var(--paper);border-radius:6px;overflow:hidden}
.rating-fill{height:100%;border-radius:6px}
.rating-fill.good{background:var(--under)}.rating-fill.bad{background:var(--over)}.rating-fill.peer{background:var(--muted)}
.rating-val{font-size:13px;font-weight:800;font-variant-numeric:tabular-nums}
.trend-badge.real{color:var(--under);background:var(--under-soft)}
.trend-svg{width:100%;height:auto;display:block}
.trend-val{font-size:11px;fill:var(--ink);font-weight:700}
.trend-xlabel{font-size:11px;fill:var(--muted)}
.trend-note{font-size:11px;color:var(--muted);margin-top:8px}
.trend-vspeer{font-size:12px;font-weight:700;border-radius:6px;padding:3px 9px}
.trend-vspeer.bad{color:var(--over);background:var(--over-soft)}
.trend-vspeer.good{color:var(--under);background:var(--under-soft)}
.trend-legend{display:flex;gap:16px;margin-top:6px;font-size:11px;color:var(--muted)}
.lg-item{display:inline-flex;align-items:center;gap:6px}
.lg-line{width:16px;height:3px;border-radius:2px;display:inline-block}
.lg-line.dash{background:repeating-linear-gradient(90deg,var(--muted) 0 4px,transparent 4px 7px)}
.we-row.peer{margin-top:-4px}
.bar.peer{background:#C6CCC2}
.wewrap{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px;display:grid;gap:9px}
.we-row{display:grid;grid-template-columns:44px 1fr auto;align-items:center;gap:10px;font-size:13px}
.we-tag{font-weight:700;color:var(--muted)}
.we-row b{font-variant-numeric:tabular-nums}
.we-note{font-size:12px;color:#3D4650;margin-top:2px}.we-note b{font-weight:800}
tr.rating-row{background:var(--warn-soft)}
.decomp{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px 16px;display:grid;gap:10px}
.decomp-row{display:grid;grid-template-columns:80px 1fr 70px;align-items:center;gap:10px}
.decomp-label{font-size:13px;font-weight:600}
.tips{display:grid;gap:10px}
.tip{display:flex;gap:12px;background:var(--surface);border:1px solid var(--line);border-left:4px solid var(--muted);border-radius:10px;padding:12px 14px}
.tip.level-high{border-left-color:var(--over)}.tip.level-mid{border-left-color:var(--warn)}.tip.level-good{border-left-color:var(--under)}.tip.level-flat{border-left-color:var(--line,#D8DCD6)}
.tip-tag{align-self:flex-start;font-size:11px;font-weight:700;background:var(--paper);border-radius:6px;padding:3px 8px;white-space:nowrap}
.tip-head{font-weight:700;margin-bottom:2px}.tip-body{font-size:13px;color:#3D4650}
.tip-main{flex:1}
.tip-facts{margin:4px 0 8px;padding-left:16px;font-size:12px;color:#4A5560}
.tip-facts li{margin-bottom:2px;list-style:disc}
.tip-alts-label{font-size:11px;font-weight:700;color:var(--muted);margin-bottom:3px}
.hi-over{color:var(--over,#C2402A);font-weight:800;font-size:1.05em;background:#FBEAE6;padding:0 3px;border-radius:3px}
.hi-under{color:var(--cobalt,#24478F);font-weight:800;font-size:1.05em;background:#E8F0FB;padding:0 3px;border-radius:3px}
.hi-cnt{color:var(--cobalt,#24478F);font-weight:800}
.hi-won{color:#182027;font-weight:800}
.rangebar{display:flex;align-items:center;gap:8px;flex-wrap:wrap;padding:10px 12px;background:#F5F7FB;border:1px solid var(--cobalt,#24478F);border-radius:9px;margin-bottom:10px;font-size:13px}
.rangebar-label{font-weight:700;color:var(--cobalt,#24478F)}
.rg-num{width:70px;padding:5px 8px;border:1px solid var(--line);border-radius:7px;text-align:center}
.rec-pg{font-weight:600;color:var(--muted)}
.rec-diff{color:var(--over,#C2402A);font-weight:800;background:#FBEAE6;padding:1px 6px;border-radius:5px}
.redfilter{padding:6px 12px;border:1px solid var(--line);background:#fff;border-radius:8px;font-size:13px;cursor:pointer;font-weight:600}
.redfilter.on{background:var(--over,#C2402A);color:#fff;border-color:var(--over,#C2402A)}
.vf-cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:8px;margin:10px 0 16px}
.vf-card{padding:10px 12px;border:1px solid var(--line);border-radius:9px;background:#fff}
.vf-card.ok{border-color:var(--under,#1E7A5B);background:#EAF5F0}
.vf-card.bad{border-color:var(--over,#C2402A);background:#FBEAE6}
.vf-k{font-size:11px;color:var(--muted);margin-bottom:3px}
.vf-v{font-size:17px;font-weight:800;font-variant-numeric:tabular-nums}
.vf-card.bad .vf-v{color:var(--over,#C2402A)}
.vf-card.ok .vf-v{color:var(--under,#1E7A5B)}
.vf-table{width:100%;border-collapse:collapse;font-size:12.5px}
.vf-table th,.vf-table td{padding:6px 8px;border-bottom:1px solid var(--line);white-space:nowrap}
.vf-table th{text-align:left;background:#FBFAF7;position:sticky;top:0}
.vf-table td.num,.vf-table th.num{text-align:right;font-variant-numeric:tabular-nums}
.vf-name{font-weight:600}
.vf-cell{color:var(--muted);font-size:11.5px}
.vf-bad{background:#FDF0ED}
.vf-bad .vf-name{color:var(--over,#C2402A);font-weight:800}
.vf-flag{color:var(--over,#C2402A);font-weight:700;font-size:11.5px;white-space:normal}
.dl-bar{display:flex;gap:8px;margin:8px 0 14px;flex-wrap:wrap}
.dl-btn{padding:8px 14px;background:var(--cobalt,#24478F);color:#fff;border:none;border-radius:8px;font-size:13px;font-weight:600;cursor:pointer}
.dl-btn:hover{opacity:.9}
.up-btn{display:inline-flex;align-items:center;cursor:pointer;background:var(--under,#1E7A5B)}
.tip-alts ul{padding-left:16px;font-size:13px}
.tip-alts li{margin-bottom:3px;list-style:"– ";color:#3D4650}
.barlabel-tag{font-size:10px;font-weight:700;color:#fff;border-radius:5px;padding:1px 6px;margin-right:6px}
.barlabel-tag.me{background:var(--cobalt)}.barlabel-tag.peer{background:var(--muted)}
.wage-appendix{margin:10px 0 4px}
table.mini{width:100%;border-collapse:collapse;font-size:12px}
table.mini th,table.mini td{padding:5px 8px;text-align:left;border-bottom:1px solid var(--line)}
table.mini th.num,table.mini td.num{text-align:right;font-variant-numeric:tabular-nums}
table.mini td.soft{color:var(--muted)}
table.mini .sumline td{border-top:1px solid var(--ink);border-bottom:none;font-weight:700}
table.mini .wl-box{display:inline-block;vertical-align:middle;margin-right:4px}
.compare .me{background:var(--cobalt-soft)}
.helperlayout{display:grid;grid-template-columns:1fr 280px;gap:20px;align-items:start}
.helperlayout .fullrow{grid-column:1/-1}
@media(max-width:860px){.helperlayout{grid-template-columns:1fr}}
.form{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px;display:grid;gap:12px}
.form label{display:grid;gap:5px;font-size:12px;font-weight:600;color:var(--muted)}
.formrow{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px}
@media(max-width:640px){.formrow{grid-template-columns:1fr}}
.amountline{display:flex;align-items:baseline;gap:10px;background:var(--paper);border-radius:8px;padding:10px 12px;font-size:13px;flex-wrap:wrap}
.amountline strong{font-size:18px;font-variant-numeric:tabular-nums}.calcnote{color:var(--muted);font-size:12px}
button.primary{border:0;background:var(--ink);color:#fff;border-radius:8px;padding:10px;font-size:14px;font-weight:700;cursor:pointer}
button.primary.danger{background:#D14343}
button.primary:disabled{opacity:.4;cursor:not-allowed}
.memo{white-space:normal;min-width:150px;color:var(--muted);font-size:12px}
.empty{text-align:center;color:var(--muted);padding:24px}
.sidecard{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:16px;position:sticky;top:76px}
.sidecard h4{font-size:12px;color:var(--muted);margin-bottom:6px}
.side-store{font-weight:800;font-size:16px;letter-spacing:-.01em;margin-bottom:10px}
.sidecard dl{display:grid;gap:8px;margin-bottom:12px}
.sidecard dl div{display:flex;justify-content:space-between}
.sidecard dt{color:var(--muted);font-size:13px}.sidecard dd{font-weight:700;font-variant-numeric:tabular-nums}
.filegrid{display:grid;grid-template-columns:repeat(auto-fill,minmax(240px,1fr));gap:12px}
.filecard{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:14px;cursor:pointer;transition:border-color .15s}
.filecard:hover{border-color:var(--cobalt)}
.filecard.ok{border-color:var(--under);background:var(--under-soft)}
.filecard-head{display:flex;justify-content:space-between;margin-bottom:8px}
.filecard-name{font-weight:700}.filecard-status{font-size:11px;color:var(--muted)}
.filecard.ok .filecard-status{color:var(--under);font-weight:700}
.filecard ul{list-style:none;padding:0;display:grid;gap:3px;font-size:12px;color:var(--muted)}
@media(prefers-reduced-motion:reduce){*{transition:none!important}}

/* ── 모바일(≤720px) 대응 — 매장 진단 화면 중심 ── */
.hktable-wrap{overflow-x:auto;-webkit-overflow-scrolling:touch}
@media(max-width:720px){
  main{padding:14px 12px}
  header{padding:10px 14px;gap:10px}
  .brand-name{font-size:14px}
  .brand-sub{font-size:11px}
  nav{overflow-x:auto;-webkit-overflow-scrolling:touch;flex-wrap:nowrap;max-width:100%;padding-bottom:2px}
  nav::-webkit-scrollbar{display:none}
  nav button{flex:0 0 auto;padding:8px 12px;font-size:13px;white-space:nowrap}
  .toolbar{gap:6px}
  .toolbar select,.toolbar input,.toolbar button{min-height:40px;font-size:14px}
  .toolbar select,.toolbar input{flex:1 1 140px}
  .statgrid{grid-template-columns:1fr 1fr;gap:8px}
  .statgrid.three{grid-template-columns:1fr}
  .stat{padding:12px}
  .stat-value{font-size:20px}
  .slotwrap{grid-template-columns:repeat(auto-fit,minmax(96px,1fr));gap:8px}
  .detailgrid{grid-template-columns:1fr}
  .wide-dbox{grid-column:auto}
  .hktable th,.hktable td{padding:6px 8px;font-size:12px}
  .cmp-toolbar{flex-wrap:wrap;gap:6px}
  .cmp-toolbar select{min-height:38px}
  .sectionhead{font-size:14px}
  .daytoggle{overflow-x:auto;-webkit-overflow-scrolling:touch;max-width:100%}
  .daytoggle button{white-space:nowrap}
  .pwgate .pwcard{width:100%;max-width:320px}
  .modal-panel{max-height:94vh;border-radius:10px}
  .modal-body{padding:12px 14px}
}
`;

/* ───────── 라이브 로더: Supabase에서 데이터 fetch ───────── */
export default function LaborApp({ myCode: myCodeProp } = {}) {
  const myCode = myCodeProp || (AUTH.role === "store" ? AUTH.code : null); // 일반매장 계정은 자기 매장에 고정
  const [state, setState] = useState("loading"); // loading | ready | error
  const [, bump] = React.useReducer((x) => x + 1, 0);
  React.useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const ctrl = new AbortController();
        const to = setTimeout(() => ctrl.abort(), 12000);
        const url = await latestSnapshotUrl();
        const r = await fetch(url + "?t=" + Date.now(), { cache: "no-store", signal: ctrl.signal });
        clearTimeout(to);
        if (!r.ok) throw new Error("HTTP " + r.status);
        const json = await r.json();
        if (!alive) return;
        initData(json);
        recomputeMetrics();
        setState("ready");
        // 아래 4가지 백그라운드 보정(점장/셀 정보, 매장 분류, 주차 제외 목록, 승인된 헬퍼 보정)을 한꺼번에 기다렸다가
        // 전부 끝난 뒤 딱 한 번만 재계산·재렌더링합니다. 예전엔 4개가 각자 따로 recomputeMetrics()+bump()를 불러서,
        // 네트워크 응답이 도착하는 타이밍이 제각각이라 화면이 여러 번 다시 그려졌습니다 — 특히 "매장 분류"가 늦게
        // 도착하면 유사매출대 비교 대상 자체가 중간에 바뀌면서 인건비율·gap이 눈에 띄게 오르락내리락해 보였습니다.
        (async () => {
          let changed = false;
          await Promise.allSettled([
            (async () => {
              try {
                const { data } = await supabase.from("store_directory").select("code,mg,cell");
                if (alive && data && data.length) { applyStoreDirectory(data); changed = true; }
              } catch (e) { console.warn("점장/셀 정보 반영 건너뜀", e); }
            })(),
            (async () => {
              try {
                const { data } = await supabase.from("peer_config").select("data").order("updated_at", { ascending: false }).limit(1).maybeSingle();
                if (alive && data && data.data) {
                  Object.entries(data.data).forEach(([code, pg]) => { const s = STORES.find((x) => x.c === code); if (s) s.pg = pg; });
                  changed = true;
                }
              } catch (e) { console.warn("매장 분류 반영 건너뜀", e); }
            })(),
            (async () => {
              try {
                const { data } = await supabase.from("week_flags").select("week_label").eq("excluded", true);
                if (alive && data) { setExcludedWeeks(data.map((r) => r.week_label)); changed = true; }
              } catch (e) { console.warn("주차 제외 목록 반영 건너뜀", e); }
            })(),
            (async () => {
              try {
                await refreshHelperEffects(); // "지금 승인된 헬퍼 전체" 기준으로 매번 처음부터 계산 → 몇 번 실행되든 항상 같은 결과(중복방지 가드 불필요)
                if (alive) changed = true;
              } catch (e) { console.warn("보정 반영 건너뜀", e); }
            })(),
          ]);
          if (alive && changed) { recomputeMetrics(); bump(); }
        })();
      } catch (e) { if (alive) { console.error(e); setState("error"); } }
    })();
    return () => { alive = false; };
  }, []);
  if (state === "loading") return (
    <div style={{ minHeight: "60vh", display: "flex", alignItems: "center", justifyContent: "center", fontFamily: "system-ui", color: "#68737E" }}>
      데이터를 불러오는 중…
    </div>
  );
  if (state === "error") return (
    <div style={{ minHeight: "60vh", display: "flex", flexDirection: "column", gap: "8px", alignItems: "center", justifyContent: "center", fontFamily: "system-ui", color: "#C2402A" }}>
      <div>데이터를 불러오지 못했습니다.</div>
      <div style={{ fontSize: "13px", color: "#68737E" }}>네트워크·주소·공개 설정을 확인해 주세요.</div>
      <button onClick={() => location.reload()} style={{ marginTop: "8px", padding: "6px 14px", borderRadius: "8px", border: "1px solid #E1E5DE", background: "#fff", cursor: "pointer" }}>다시 시도</button>
    </div>
  );
  return <AppInner myCode={myCode} />;
}
