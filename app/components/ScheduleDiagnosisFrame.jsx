"use client";
import React from "react";

export default function ScheduleDiagnosisFrame() {
  return (
    <section className="schedule-diagnosis-shell" aria-label="스케줄 진단">
      <iframe
        className="schedule-diagnosis-frame"
        src="/schedule-diagnosis.html"
        title="애슐리 스케줄 진단"
        loading="eager"
      />
    </section>
  );
}
