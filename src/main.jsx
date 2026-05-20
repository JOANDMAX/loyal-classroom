import { useState, useCallback } from "react";
import * as XLSX from "xlsx";

const B = { pri: "#00565E", priL: "#007A85", priD: "#003D43", acc: "#F0A500", ok: "#27AE60", bg: "#F7FAFA", brd: "#E2EDED", t1: "#1A2F33", t2: "#5A7A7E", t3: "#8FA5A8" };

const STEPS = [
  { key: "L", label: "Listen", desc: "Active listening and paraphrasing", icon: "👂" },
  { key: "O", label: "Offer Empathy", desc: "Emotional Connection", icon: "💛" },
  { key: "Y", label: "You Own the Problem", desc: "Taking Ownership & Expectation Management", icon: "🤝" },
  { key: "A", label: "Ask if the passenger is happy", desc: "Follow-up & Emotional Repair", icon: "✅" },
  { key: "L2", label: "Let the passenger know you appreciate the feedback", desc: "Positive ending · Provide reassurance · Show appreciation", icon: "🙏" },
];

const SCENARIO = {
  context: "Due to an aircraft change, Mr. and Mrs. Smith's pre-selected seats are no longer available. They have now been seated apart on their 10th anniversary trip. With the flight nearly full, they appear disappointed and are seeking your assistance.",
  passenger: "Mr & Mrs Smith",
  quote: `"Please, is there anything you can do? It's our 10th anniversary, and we've been looking forward to this trip. We want to sit together — it would mean so much. Can you help us?"`,
};

const DEMO = [
  { email: "alice.wong@cathay.com", answers: { L: "I understand that you and Mrs. Smith were supposed to sit together for your 10th anniversary, and now your seats have been changed. That must be really frustrating and disappointing, especially for such a meaningful trip.", O: "I can truly imagine how disappointing this must feel — your 10th anniversary is such a special milestone, and being seated apart is not the experience you planned for.", Y: "Let me personally take care of this for you. I'll check a few options right now: first, I can look for any adjacent seats that might have opened up; second, I can ask other passengers if anyone would be willing to swap. I'll have an update for you within the next 5 minutes.", A: "Mr. and Mrs. Smith, I'm happy we found seats together for you. I just wanted to check — is there anything else I can help with to make sure your anniversary trip starts off right?", L2: "Thank you so much for your patience while I sorted this out. You bringing this to my attention actually gave us the opportunity to ensure your anniversary journey is as special as it should be. Happy anniversary to you both!" }, scores: { L: { score: 3, why: "Covers all facts AND acknowledges emotional impact.", suggest: "Strong. Consider paraphrasing their request to show active listening." }, O: { score: 3, why: "Strong emotional labeling with 'disappointing' and 'special milestone'.", suggest: "Could also acknowledge the anticipation of planning this trip." }, Y: { score: 3, why: "Personal ownership, multiple options, clear 5-minute timeframe.", suggest: "Consider explicitly asking 'Which option would you prefer?'" }, A: { score: 3, why: "Proactive follow-up, opens door for additional needs, ties to anniversary.", suggest: "Could also observe their expressions to confirm mood improvement." }, L2: { score: 3, why: "Reframes inconvenience as opportunity. Adds anniversary wishes.", suggest: "Excellent closure setting a lasting positive impression." } } },
  { email: "ben.chan@cathay.com", answers: { L: "I see you want to sit together. Let me check.", O: "I'm sorry for the inconvenience caused by the seat change.", Y: "I'll move one of you to row 35. That should work.", A: "Is that okay for you?", L2: "Thank you for your feedback." }, scores: { L: { score: 1, why: "Misses anniversary, no emotional acknowledgment. Too brief.", suggest: "Try: 'I understand you're celebrating your 10th anniversary, and being seated apart must be really disappointing.'" }, O: { score: 1, why: "Forbidden phrase 'sorry for the inconvenience'. No emotional depth.", suggest: "Try: 'I can imagine how frustrating this must feel, especially when you've been looking forward to celebrating together.'" }, Y: { score: 1, why: "Single solution, no reassurance, no preference asked.", suggest: "Try: 'Let me personally look into this. I can check for adjacent seats or ask passengers to swap. I'll update you within 5 minutes.'" }, A: { score: 2, why: "Functional closed question. No emotional repair.", suggest: "Try: 'I'm glad we arranged seats together. How are you feeling? Anything else for your anniversary trip?'" }, L2: { score: 1, why: "Generic. No reframing, no personal touch.", suggest: "Try: 'Thank you for bringing this to our attention — it gave us a chance to make your anniversary special.'" } } },
  { email: "clara.liu@cathay.com", answers: { L: "I understand your seats were changed because of the aircraft swap, and you'd like to sit together for your anniversary. I can see this isn't what you expected.", O: "I'm sorry this happened. A 10th anniversary is a big deal and you shouldn't have to worry about seats on your special day.", Y: "Let me see what I can do. I'll check if there are two seats together anywhere and get back to you shortly.", A: "I've found two seats together in row 22. Does that work for you both?", L2: "Thanks for letting me know. I'm glad we could sort it out. Have a lovely anniversary trip!" }, scores: { L: { score: 2, why: "Covers facts but doesn't name the disappointment.", suggest: "Try: 'I understand this must feel disappointing, especially when you've planned this anniversary trip together.'" }, O: { score: 2, why: "Acknowledges occasion but lacks emotional labeling.", suggest: "Try: 'I can imagine how disappointing it must feel to not be sitting together for such a special milestone.'" }, Y: { score: 2, why: "One approach, no timeframe. 'Shortly' is vague.", suggest: "Try: 'Within 5 minutes, I'll check adjacent seats and whether nearby passengers might swap.'" }, A: { score: 2, why: "Confirms solution but doesn't check emotional state.", suggest: "Add: 'Is there anything else to make your anniversary celebration even better?'" }, L2: { score: 2, why: "Sincere but doesn't reframe or show value of feedback.", suggest: "Try: 'Thank you for your patience — it gave us the chance to make sure your anniversary starts right.'" } } },
];

const Stars = ({ count, sz = 16 }) => <span style={{ letterSpacing: 2 }}>{[1,2,3].map(i => <span key={i} style={{ fontSize: sz, opacity: i <= count ? 1 : 0.18 }}>⭐</span>)}</span>;

function wrapText(text, maxChars) {
  const words = text.split(" ");
  const lines = []; let cur = "";
  words.forEach(w => { if ((cur + " " + w).trim().length > maxChars) { if (cur) lines.push(cur); cur = w; } else { cur = (cur + " " + w).trim(); } });
  if (cur) lines.push(cur);
  return lines.length ? lines : [""];
}

function generatePDF(student) {
  const total = STEPS.reduce((s, st) => s + (student.scores?.[st.key]?.score || 0), 0);
  const w = 595; const m = 40;
  let y = 0; const pages = []; let page = [];
  const txt = (t, x, _y, o = {}) => page.push({ T: 1, t, x, y: _y, ...o });
  const rect = (x, _y, rw, rh, c) => page.push({ T: 0, x, y: _y, w: rw, h: rh, c });
  const np = () => { pages.push(page); page = []; y = m; };

  rect(0, 0, w, 70, B.pri);
  txt("Your LOYAL Report Card", m, 42, { s: 20, b: 1, c: "#fff" });
  y = 85;
  txt(`Email: ${student.email}`, m, y, { s: 10, c: B.t2 });
  txt(`Total Score: ${total} / 15`, w - m - 110, y, { s: 10, b: 1, c: B.pri });
  y = 110;

  STEPS.forEach(st => {
    const r = student.scores?.[st.key] || {};
    if (y + 180 > 800) np();
    rect(m, y, w - m * 2, 24, B.pri);
    txt(`${st.key === "L2" ? "L" : st.key} — ${st.label}   ${"⭐".repeat(r.score || 0)}`, m + 10, y + 16, { s: 10, b: 1, c: "#fff" });
    y += 32;
    [
      { lbl: "YOUR ANSWER", val: student.answers[st.key], lc: B.pri },
      { lbl: "FEEDBACK", val: r.why, lc: B.acc },
      { lbl: "SUGGESTION", val: r.suggest, lc: B.ok },
    ].forEach(f => {
      txt(f.lbl, m, y, { s: 8, b: 1, c: f.lc }); y += 12;
      wrapText(f.val || "N/A", 75).forEach(l => { txt(l, m + 4, y, { s: 9, c: B.t2 }); y += 13; });
      y += 4;
    });
    y += 12;
  });
  pages.push(page);

  const svgs = pages.map(cmds => {
    let s = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="842" viewBox="0 0 ${w} 842"><rect width="${w}" height="842" fill="#fff"/>`;
    cmds.forEach(c => {
      if (c.T === 0) s += `<rect x="${c.x}" y="${c.y}" width="${c.w}" height="${c.h}" fill="${c.c}"/>`;
      else { const esc = c.t.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"); s += `<text x="${c.x}" y="${c.y}" font-family="Arial,sans-serif" font-size="${c.s||11}" font-weight="${c.b?"bold":"normal"}" fill="${c.c||B.t1}">${esc}</text>`; }
    });
    return s + "</svg>";
  });
  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>LOYAL Report - ${student.email}</title><style>@media print{@page{margin:0}body{margin:0}}body{margin:20px;font-family:Arial,sans-serif}svg{display:block;margin:0 auto 20px;box-shadow:0 1px 4px rgba(0,0,0,.1)}</style></head><body>${svgs.join("")}<script>window.onload=function(){window.print()}<\/script></body></html>`;
  const blob = new Blob([html], { type: "text/html" });
  window.open(URL.createObjectURL(blob), "_blank");
}

function exportExcel(students) {
  const rows = students.map(s => {
    const row = { Email: s.email };
    STEPS.forEach(st => {
      const k = st.key;
      row[`${k}_Answer`] = s.answers?.[k] || "";
      row[`${k}_Score`] = s.scores?.[k]?.score || "";
      row[`${k}_Feedback`] = s.scores?.[k]?.why || "";
      row[`${k}_Suggestion`] = s.scores?.[k]?.suggest || "";
    });
    row.Total = STEPS.reduce((sum, st) => sum + (s.scores?.[st.key]?.score || 0), 0);
    return row;
  });
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "LOYAL Results");
  XLSX.writeFile(wb, `LOYAL_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

/* ═══ Role Select ═══ */
function RoleSelect({ onSelect }) {
  return (
    <div style={{ minHeight: "100vh", background: `linear-gradient(155deg, ${B.priD} 0%, ${B.pri} 50%, ${B.priL} 100%)`, display: "flex", alignItems: "center", justifyContent: "center", padding: 20, fontFamily: "'DM Sans','Helvetica Neue',sans-serif" }}>
      <div style={{ textAlign: "center", maxWidth: 440 }}>
        <div style={{ fontSize: 13, letterSpacing: 6, color: "rgba(255,255,255,0.45)", textTransform: "uppercase", marginBottom: 12 }}>Cathay Academy</div>
        <h1 style={{ fontSize: 44, fontWeight: 800, color: "#fff", margin: "0 0 8px", letterSpacing: -1 }}>LOYAL</h1>
        <p style={{ fontSize: 15, color: "rgba(255,255,255,0.6)", margin: "0 0 48px" }}>Classroom Assessment System</p>
        <div style={{ display: "flex", gap: 16, justifyContent: "center", flexWrap: "wrap" }}>
          {[
            { role: "teacher", label: "Facilitator", icon: "🎓", sub: "Dashboard & scoring" },
            { role: "student", label: "Participant", icon: "✈️", sub: "Submit responses" },
          ].map(r => (
            <button key={r.role} onClick={() => onSelect(r.role)} style={{
              background: "rgba(255,255,255,0.08)", backdropFilter: "blur(10px)",
              border: "1px solid rgba(255,255,255,0.15)", borderRadius: 16,
              padding: "28px 32px", cursor: "pointer", transition: "all 0.25s", minWidth: 170, textAlign: "center",
            }}
            onMouseEnter={e => { e.currentTarget.style.background = "rgba(255,255,255,0.18)"; e.currentTarget.style.transform = "translateY(-2px)"; }}
            onMouseLeave={e => { e.currentTarget.style.background = "rgba(255,255,255,0.08)"; e.currentTarget.style.transform = "translateY(0)"; }}>
              <div style={{ fontSize: 36, marginBottom: 10 }}>{r.icon}</div>
              <div style={{ fontSize: 15, fontWeight: 700, color: "#fff" }}>{r.label}</div>
              <div style={{ fontSize: 11, color: "rgba(255,255,255,0.45)", marginTop: 6 }}>{r.sub}</div>
            </button>
          ))}
        </div>
      </div>
    </div>
  );
}

/* ═══ Class Summary ═══ */
function ClassSummary({ students }) {
  const scored = students.filter(s => s.scores);
  if (!scored.length) return null;
  const avgT = (scored.reduce((s, st) => s + STEPS.reduce((ss, step) => ss + (st.scores[step.key]?.score || 0), 0), 0) / scored.length).toFixed(1);
  const avgs = STEPS.map(st => ({ ...st, avg: +(scored.reduce((s, stu) => s + (stu.scores[st.key]?.score || 0), 0) / scored.length).toFixed(1) }));
  const sorted = [...avgs].sort((a, b) => a.avg - b.avg);

  return (
    <div style={{ margin: "16px 20px", padding: "16px 18px", background: "#fff", borderRadius: 14, border: `1px solid ${B.brd}` }}>
      <div style={{ fontSize: 13, fontWeight: 700, color: B.pri, marginBottom: 12 }}>📊 Class Summary</div>
      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", marginBottom: 12 }}>
        <div style={{ background: `${B.pri}08`, borderRadius: 10, padding: "8px 14px", flex: 1, minWidth: 80 }}>
          <div style={{ fontSize: 18, fontWeight: 800, color: B.pri }}>{scored.length}</div>
          <div style={{ fontSize: 10, color: B.t3 }}>Submitted</div>
        </div>
        <div style={{ background: `${B.pri}08`, borderRadius: 10, padding: "8px 14px", flex: 1, minWidth: 80 }}>
          <div style={{ fontSize: 18, fontWeight: 800, color: B.pri }}>{avgT}</div>
          <div style={{ fontSize: 10, color: B.t3 }}>Avg /15</div>
        </div>
      </div>
      <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
        {avgs.map(s => (
          <div key={s.key} style={{ flex: 1, minWidth: 50, textAlign: "center", padding: "6px 4px", borderRadius: 8, background: s.key === sorted[0].key ? "#FFF0F0" : s.key === sorted[sorted.length - 1].key ? "#F0FFF4" : `${B.pri}05` }}>
            <div style={{ fontSize: 14, fontWeight: 700, color: s.key === sorted[0].key ? "#E74C3C" : s.key === sorted[sorted.length - 1].key ? B.ok : B.pri }}>{s.avg}</div>
            <div style={{ fontSize: 9, color: B.t3 }}>{s.key === "L2" ? "L₂" : s.key}</div>
          </div>
        ))}
      </div>
      <div style={{ marginTop: 10, fontSize: 11, color: B.t2, lineHeight: 1.6 }}>
        💪 Strongest: <strong style={{ color: B.ok }}>{sorted[sorted.length-1].key === "L2" ? "L₂" : sorted[sorted.length-1].key} — {sorted[sorted.length-1].label}</strong><br/>
        📌 Needs work: <strong style={{ color: "#E74C3C" }}>{sorted[0].key === "L2" ? "L₂" : sorted[0].key} — {sorted[0].label}</strong>
      </div>
    </div>
  );
}

/* ═══ Report Card ═══ */
function ReportCard({ student }) {
  const total = STEPS.reduce((s, st) => s + (student.scores?.[st.key]?.score || 0), 0);
  return (
    <div style={{ padding: 20 }}>
      <div style={{ maxWidth: 640, margin: "0 auto" }}>
        <div style={{ background: `linear-gradient(135deg, ${B.priD}, ${B.pri})`, borderRadius: "16px 16px 0 0", padding: "26px 24px" }}>
          <h1 style={{ fontSize: 21, fontWeight: 800, color: "#fff", margin: 0 }}>Your LOYAL Report Card</h1>
        </div>
        <div style={{ background: "#E6F4F5", padding: "12px 24px", display: "flex", justifyContent: "space-between", flexWrap: "wrap", gap: 8, borderBottom: `2px solid ${B.brd}` }}>
          <div><div style={{ fontSize: 9, color: B.pri, fontWeight: 700, textTransform: "uppercase" }}>Email</div><div style={{ fontSize: 13, fontWeight: 600, color: B.t1 }}>{student.email}</div></div>
          <div><div style={{ fontSize: 9, color: B.pri, fontWeight: 700, textTransform: "uppercase" }}>Total</div><span style={{ background: B.pri, color: "#fff", padding: "3px 14px", borderRadius: 20, fontSize: 13, fontWeight: 700 }}>{total}/15</span></div>
        </div>
        <div style={{ background: "#fff", padding: "18px 24px 10px", borderRadius: "0 0 16px 16px", border: `1px solid ${B.brd}`, borderTop: "none" }}>
          {STEPS.map(st => {
            const r = student.scores?.[st.key] || {};
            return (
              <div key={st.key} style={{ border: `1px solid ${B.brd}`, borderRadius: 12, marginBottom: 14, overflow: "hidden" }}>
                <div style={{ background: B.pri, padding: "9px 16px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <span style={{ color: "#fff", fontWeight: 700, fontSize: 12 }}>{st.key === "L2" ? "L" : st.key} — {st.label}</span>
                  <Stars count={r.score || 0} sz={13} />
                </div>
                <div style={{ padding: "12px 16px" }}>
                  {[
                    { lbl: "YOUR ANSWER", val: student.answers[st.key], bg: "#F4F9FA", bc: B.pri },
                    { lbl: "FEEDBACK", val: r.why, bg: "#FFFBF0", bc: B.acc },
                    { lbl: "SUGGESTION", val: r.suggest, bg: "#F0FFF4", bc: B.ok },
                  ].map(f => (
                    <div key={f.lbl}>
                      <div style={{ fontSize: 9, color: B.pri, fontWeight: 700, marginBottom: 3, textTransform: "uppercase" }}>{f.lbl}</div>
                      <div style={{ fontSize: 12.5, lineHeight: 1.6, padding: "7px 11px", background: f.bg, borderLeft: `3px solid ${f.bc}`, borderRadius: "0 6px 6px 0", marginBottom: 9 }}>{f.val || "N/A"}</div>
                    </div>
                  ))}
                </div>
              </div>
            );
          })}
          <button onClick={() => generatePDF(student)} style={{ width: "100%", padding: 13, background: B.pri, color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit", marginBottom: 14 }}>
            📄 Download / Print Report Card
          </button>
        </div>
      </div>
    </div>
  );
}

/* ═══ Student View ═══ */
function StudentView({ session, onSubmit, onBack }) {
  const [email, setEmail] = useState("");
  const [started, setStarted] = useState(false);
  const [step, setStep] = useState(0);
  const [answers, setAnswers] = useState({});
  const [submitted, setSubmitted] = useState(false);

  const me = session.students.find(s => s.email === email);

  if (me?.approved) {
    return (
      <div style={{ minHeight: "100vh", background: B.bg, fontFamily: "'DM Sans',sans-serif" }}>
        <div style={{ padding: "10px 16px", borderBottom: `1px solid ${B.brd}`, background: "rgba(255,255,255,0.92)" }}>
          <button onClick={onBack} style={{ background: "none", border: "none", fontSize: 13, fontWeight: 700, color: B.pri, cursor: "pointer", fontFamily: "inherit" }}>← Back to Home</button>
        </div>
        <ReportCard student={me} />
        <ClassSummary students={session.students.filter(s => s.approved)} />
      </div>
    );
  }

  if (submitted) {
    return (
      <div style={{ minHeight: "100vh", background: B.bg, fontFamily: "'DM Sans',sans-serif" }}>
        <div style={{ padding: "10px 16px", borderBottom: `1px solid ${B.brd}`, background: "rgba(255,255,255,0.92)" }}>
          <button onClick={onBack} style={{ background: "none", border: "none", fontSize: 13, fontWeight: 700, color: B.pri, cursor: "pointer", fontFamily: "inherit" }}>← Back to Home</button>
        </div>
        <div style={{ padding: "70px 20px", textAlign: "center" }}>
          <div style={{ fontSize: 56, marginBottom: 16 }}>✅</div>
          <h2 style={{ fontSize: 22, color: B.pri, margin: "0 0 12px" }}>Submitted!</h2>
          <p style={{ fontSize: 14, color: B.t2, lineHeight: 1.7, maxWidth: 340, margin: "0 auto" }}>
            Your facilitator is reviewing your responses.<br/>Your report card will appear here once approved.
          </p>
          <div style={{ marginTop: 28, padding: "14px 24px", background: `${B.pri}08`, borderRadius: 12, fontSize: 13, color: B.pri, display: "inline-block", lineHeight: 1.6 }}>
            🔄 <strong>Refresh this page</strong> after your facilitator announces results are ready.
          </div>
        </div>
      </div>
    );
  }

  if (!started) {
    return (
      <div style={{ minHeight: "100vh", background: B.bg, fontFamily: "'DM Sans',sans-serif" }}>
        <div style={{ padding: "10px 16px", borderBottom: `1px solid ${B.brd}`, background: "rgba(255,255,255,0.92)" }}>
          <button onClick={onBack} style={{ background: "none", border: "none", fontSize: 13, fontWeight: 700, color: B.pri, cursor: "pointer", fontFamily: "inherit" }}>← Back to Home</button>
        </div>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "center", padding: 20, minHeight: "70vh" }}>
          <div style={{ maxWidth: 400, width: "100%" }}>
            <div style={{ background: B.pri, borderRadius: "16px 16px 0 0", padding: "24px 28px" }}>
              <div style={{ fontSize: 11, letterSpacing: 4, color: "rgba(255,255,255,0.45)", textTransform: "uppercase" }}>LOYAL Assessment</div>
              <h2 style={{ fontSize: 20, color: "#fff", margin: "8px 0 0", fontWeight: 700 }}>Enter Your Email</h2>
            </div>
            <div style={{ background: "#fff", borderRadius: "0 0 16px 16px", padding: 28, border: `1px solid ${B.brd}`, borderTop: "none" }}>
              <input type="email" placeholder="your.name@cathay.com" value={email}
                onChange={e => setEmail(e.target.value)} onKeyDown={e => e.key === "Enter" && email.includes("@") && setStarted(true)}
                style={{ width: "100%", padding: "14px 16px", fontSize: 15, border: `1.5px solid ${B.brd}`, borderRadius: 10, outline: "none", boxSizing: "border-box", fontFamily: "inherit" }}
                onFocus={e => e.target.style.borderColor = B.pri} onBlur={e => e.target.style.borderColor = B.brd} />
              <button onClick={() => email.includes("@") && setStarted(true)} disabled={!email.includes("@")}
                style={{ width: "100%", marginTop: 16, padding: 14, background: email.includes("@") ? B.pri : B.brd, color: "#fff", border: "none", borderRadius: 10, fontSize: 15, fontWeight: 600, cursor: email.includes("@") ? "pointer" : "default", fontFamily: "inherit" }}>
                Start →
              </button>
            </div>
          </div>
        </div>
      </div>
    );
  }

  const cur = STEPS[step];
  const progress = ((step + 1) / STEPS.length) * 100;
  const allFilled = STEPS.every(s => (answers[s.key] || "").trim().length > 0);

  return (
    <div style={{ minHeight: "100vh", background: B.bg, fontFamily: "'DM Sans',sans-serif", paddingBottom: 80 }}>
      <div style={{ position: "sticky", top: 0, zIndex: 10, background: B.bg }}>
        <div style={{ height: 3, background: B.brd }}><div style={{ height: 3, background: B.pri, width: `${progress}%`, transition: "width 0.4s" }} /></div>
        <div style={{ padding: "10px 20px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <button onClick={onBack} style={{ background: "none", border: "none", fontSize: 12, fontWeight: 700, color: B.pri, cursor: "pointer", fontFamily: "inherit" }}>← Exit</button>
          <div style={{ display: "flex", gap: 5 }}>
            {STEPS.map((s, i) => (
              <div key={s.key} onClick={() => setStep(i)} style={{
                width: 26, height: 26, borderRadius: 7,
                background: i === step ? B.pri : (answers[s.key] || "").trim() ? `${B.pri}18` : "transparent",
                color: i === step ? "#fff" : (answers[s.key] || "").trim() ? B.pri : B.t3,
                display: "flex", alignItems: "center", justifyContent: "center",
                fontSize: 10, fontWeight: 700, cursor: "pointer",
                border: `1.5px solid ${i === step ? B.pri : (answers[s.key] || "").trim() ? B.pri : B.brd}`,
              }}>{s.key === "L2" ? "L₂" : s.key}</div>
            ))}
          </div>
        </div>
      </div>
      <div style={{ padding: "4px 20px 12px" }}>
        <div style={{ background: `linear-gradient(135deg, ${B.priD}, ${B.pri})`, borderRadius: 14, padding: "16px 18px", color: "#fff" }}>
          <div style={{ fontSize: 10, letterSpacing: 3, opacity: 0.45, textTransform: "uppercase", marginBottom: 6 }}>Scenario</div>
          <div style={{ fontSize: 12.5, lineHeight: 1.6, opacity: 0.85 }}>{SCENARIO.context}</div>
          <div style={{ marginTop: 10, padding: "10px 14px", background: "rgba(255,255,255,0.1)", borderRadius: 10, borderLeft: "3px solid rgba(255,255,255,0.35)" }}>
            <div style={{ fontSize: 10, fontWeight: 700, opacity: 0.55, marginBottom: 3 }}>{SCENARIO.passenger}</div>
            <div style={{ fontSize: 12.5, fontStyle: "italic", lineHeight: 1.5, opacity: 0.9 }}>{SCENARIO.quote}</div>
          </div>
        </div>
      </div>
      <div style={{ padding: "4px 20px" }}>
        <div style={{ background: "#fff", borderRadius: 14, border: `1px solid ${B.brd}`, overflow: "hidden" }}>
          <div style={{ padding: "14px 18px", borderBottom: `1px solid ${B.brd}`, display: "flex", alignItems: "center", gap: 10 }}>
            <span style={{ fontSize: 22 }}>{cur.icon}</span>
            <div>
              <div style={{ fontSize: 15, fontWeight: 700, color: B.t1 }}>{cur.key === "L2" ? "L" : cur.key} — {cur.label}</div>
              <div style={{ fontSize: 11, color: B.t3 }}>{cur.desc}</div>
            </div>
          </div>
          <div style={{ padding: 18 }}>
            <textarea placeholder={`What would you say to ${SCENARIO.passenger}?`}
              value={answers[cur.key] || ""} onChange={e => setAnswers({ ...answers, [cur.key]: e.target.value })}
              style={{ width: "100%", minHeight: 130, padding: 14, fontSize: 14, lineHeight: 1.6, border: `1.5px solid ${B.brd}`, borderRadius: 10, outline: "none", resize: "vertical", fontFamily: "inherit", boxSizing: "border-box" }}
              onFocus={e => e.target.style.borderColor = B.pri} onBlur={e => e.target.style.borderColor = B.brd} />
          </div>
        </div>
      </div>
      <div style={{ position: "fixed", bottom: 0, left: 0, right: 0, background: "#fff", borderTop: `1px solid ${B.brd}`, padding: "10px 20px", display: "flex", gap: 10 }}>
        <button onClick={() => step > 0 && setStep(step - 1)} disabled={step === 0}
          style={{ flex: 1, padding: 13, background: "transparent", color: step === 0 ? B.t3 : B.pri, border: `1.5px solid ${step === 0 ? B.brd : B.pri}`, borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: step === 0 ? "default" : "pointer", fontFamily: "inherit" }}>← Back</button>
        {step < STEPS.length - 1 ? (
          <button onClick={() => setStep(step + 1)} style={{ flex: 2, padding: 13, background: B.pri, color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>Next →</button>
        ) : (
          <button onClick={() => { onSubmit({ email, answers, timestamp: new Date().toISOString() }); setSubmitted(true); }} disabled={!allFilled}
            style={{ flex: 2, padding: 13, background: allFilled ? B.ok : B.brd, color: "#fff", border: "none", borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: allFilled ? "pointer" : "default", fontFamily: "inherit" }}>Submit All ✓</button>
        )}
      </div>
    </div>
  );
}

/* ═══ Teacher View ═══ */
function TeacherView({ session, onUpdateStudent, onApprove, onApproveAll, onLoadDemo, onBack }) {
  const [expanded, setExpanded] = useState(null);
  const students = session.students || [];
  const pending = students.filter(s => !s.approved && s.scores);
  const approved = students.filter(s => s.approved);

  return (
    <div style={{ minHeight: "100vh", background: B.bg, fontFamily: "'DM Sans',sans-serif" }}>
      <div style={{ background: `linear-gradient(135deg, ${B.priD}, ${B.pri})`, padding: "22px 22px 18px" }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
          <div style={{ fontSize: 10, letterSpacing: 4, color: "rgba(255,255,255,0.35)", textTransform: "uppercase" }}>Facilitator Dashboard</div>
          <button onClick={onBack} style={{ background: "rgba(255,255,255,0.12)", border: "1px solid rgba(255,255,255,0.2)", borderRadius: 8, padding: "5px 12px", fontSize: 11, fontWeight: 600, color: "#fff", cursor: "pointer", fontFamily: "inherit" }}>← Home</button>
        </div>
        <h1 style={{ fontSize: 20, fontWeight: 800, color: "#fff", margin: "0 0 14px" }}>LOYAL — Live Session</h1>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {[
            { label: "Submitted", count: students.length, color: "#fff" },
            { label: "AI Scored", count: pending.length + approved.length, color: B.acc },
            { label: "Approved", count: approved.length, color: "#4ADE80" },
          ].map(s => (
            <div key={s.label} style={{ background: "rgba(255,255,255,0.08)", borderRadius: 10, padding: "9px 14px", minWidth: 72 }}>
              <div style={{ fontSize: 20, fontWeight: 800, color: s.color }}>{s.count}</div>
              <div style={{ fontSize: 10, color: "rgba(255,255,255,0.45)" }}>{s.label}</div>
            </div>
          ))}
        </div>
      </div>

      <div style={{ padding: "14px 22px", display: "flex", gap: 8, flexWrap: "wrap" }}>
        {students.length === 0 && (
          <button onClick={onLoadDemo} style={{ padding: "10px 18px", background: `${B.pri}10`, color: B.pri, border: `1.5px solid ${B.pri}30`, borderRadius: 10, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>📋 Load 3 Demo Students</button>
        )}
        {pending.length > 0 && (
          <button onClick={onApproveAll} style={{ padding: "10px 18px", background: B.ok, color: "#fff", border: "none", borderRadius: 10, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>✓ Approve All ({pending.length})</button>
        )}
        {students.length > 0 && (
          <button onClick={() => exportExcel(students)} style={{ padding: "10px 18px", background: "#1D6F42", color: "#fff", border: "none", borderRadius: 10, fontSize: 12, fontWeight: 600, cursor: "pointer", fontFamily: "inherit" }}>📊 Export Excel</button>
        )}
      </div>

      <div style={{ padding: "0 22px 24px" }}>
        {students.length === 0 ? (
          <div style={{ textAlign: "center", padding: "40px 20px", color: B.t3 }}>
            <div style={{ fontSize: 44, marginBottom: 10 }}>📱</div>
            <div style={{ fontSize: 14, fontWeight: 600 }}>Waiting for participants...</div>
            <div style={{ fontSize: 12, marginTop: 6, lineHeight: 1.6 }}>Share the link with your class, or tap "Load 3 Demo Students" to preview.</div>
          </div>
        ) : (
          students.map((student, idx) => {
            const total = STEPS.reduce((s, st) => s + (student.scores?.[st.key]?.score || 0), 0);
            const isOpen = expanded === idx;
            return (
              <div key={idx} style={{ background: "#fff", borderRadius: 12, border: `1px solid ${student.approved ? B.ok + "40" : B.brd}`, marginBottom: 10, overflow: "hidden" }}>
                <div onClick={() => setExpanded(isOpen ? null : idx)} style={{ padding: "13px 16px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer" }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <div style={{ width: 34, height: 34, borderRadius: 9, background: student.approved ? `${B.ok}12` : `${B.pri}0A`, display: "flex", alignItems: "center", justifyContent: "center", fontSize: 15 }}>
                      {student.approved ? "✅" : "📝"}
                    </div>
                    <div>
                      <div style={{ fontSize: 13, fontWeight: 600, color: B.t1 }}>{student.email}</div>
                      <div style={{ fontSize: 10, color: B.t3 }}>{student.approved ? "Approved" : "Pending review"}</div>
                    </div>
                  </div>
                  <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
                    <span style={{ background: B.pri, color: "#fff", padding: "2px 10px", borderRadius: 20, fontSize: 12, fontWeight: 700 }}>{total}/15</span>
                    <span style={{ fontSize: 11, color: B.t3, transform: isOpen ? "rotate(180deg)" : "rotate(0)", transition: "transform 0.2s", display: "inline-block" }}>▼</span>
                  </div>
                </div>
                {isOpen && (
                  <div style={{ padding: "0 16px 16px", borderTop: `1px solid ${B.brd}` }}>
                    {STEPS.map(st => {
                      const r = student.scores?.[st.key] || {};
                      return (
                        <div key={st.key} style={{ padding: "12px 0", borderBottom: `1px solid ${B.brd}` }}>
                          <div style={{ fontSize: 12, fontWeight: 700, color: B.pri, marginBottom: 6 }}>{st.icon} {st.key === "L2" ? "L" : st.key} — {st.label}</div>
                          <div style={{ fontSize: 12.5, lineHeight: 1.6, padding: "7px 11px", background: "#F4F9FA", borderLeft: `3px solid ${B.pri}`, borderRadius: "0 6px 6px 0", marginBottom: 8 }}>
                            {student.answers[st.key] || "N/A"}
                          </div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                            <span style={{ fontSize: 11, color: B.t3, fontWeight: 600 }}>Score:</span>
                            {[1, 2, 3].map(v => (
                              <button key={v} onClick={e => { e.stopPropagation(); onUpdateStudent(idx, st.key, "score", v); }}
                                style={{ width: 32, height: 32, borderRadius: 8, border: `1.5px solid ${r.score === v ? B.pri : B.brd}`, background: r.score === v ? B.pri : "#fff", color: r.score === v ? "#fff" : B.t2, fontSize: 13, fontWeight: 700, cursor: "pointer", fontFamily: "inherit" }}>{v}</button>
                            ))}
                            <Stars count={r.score || 0} sz={13} />
                          </div>
                          <div style={{ fontSize: 10, color: B.acc, fontWeight: 700, marginBottom: 3 }}>FEEDBACK</div>
                          <textarea value={r.why || ""} onChange={e => { e.stopPropagation(); onUpdateStudent(idx, st.key, "why", e.target.value); }}
                            style={{ width: "100%", minHeight: 44, padding: 8, fontSize: 12, lineHeight: 1.5, border: `1px solid ${B.brd}`, borderRadius: 8, outline: "none", resize: "vertical", fontFamily: "inherit", boxSizing: "border-box" }}
                            onFocus={e => e.target.style.borderColor = B.acc} onBlur={e => e.target.style.borderColor = B.brd} />
                          <div style={{ fontSize: 10, color: B.ok, fontWeight: 700, marginBottom: 3, marginTop: 6 }}>SUGGESTION</div>
                          <textarea value={r.suggest || ""} onChange={e => { e.stopPropagation(); onUpdateStudent(idx, st.key, "suggest", e.target.value); }}
                            style={{ width: "100%", minHeight: 44, padding: 8, fontSize: 12, lineHeight: 1.5, border: `1px solid ${B.brd}`, borderRadius: 8, outline: "none", resize: "vertical", fontFamily: "inherit", boxSizing: "border-box" }}
                            onFocus={e => e.target.style.borderColor = B.ok} onBlur={e => e.target.style.borderColor = B.brd} />
                        </div>
                      );
                    })}
                    {!student.approved && (
                      <button onClick={e => { e.stopPropagation(); onApprove(idx); }} style={{
                        width: "100%", marginTop: 12, padding: 12, background: B.ok, color: "#fff",
                        border: "none", borderRadius: 10, fontSize: 13, fontWeight: 600, cursor: "pointer", fontFamily: "inherit",
                      }}>✓ Approve & Release Report Card</button>
                    )}
                  </div>
                )}
              </div>
            );
          })
        )}
      </div>
    </div>
  );
}

/* ═══ Main ═══ */
export default function LOYALApp() {
  const [role, setRole] = useState(null);
  const [session, setSession] = useState({ students: [] });

  const handleSubmit = useCallback(sub => {
    const scores = {};
    STEPS.forEach(s => {
      const len = (sub.answers[s.key] || "").trim().length;
      const score = len > 80 ? 3 : len > 30 ? 2 : 1;
      scores[s.key] = {
        score,
        why: score === 3 ? "Strong response demonstrating key competencies." : score === 2 ? "Adequate but could add emotional labeling." : "Key elements missing.",
        suggest: score === 3 ? "Vary phrasing across scenarios." : score === 2 ? "Add emotional labeling and timeframes." : "Acknowledge all facts AND emotional state.",
      };
    });
    setSession(prev => ({ ...prev, students: [...prev.students, { email: sub.email, answers: sub.answers, scores, approved: false }] }));
  }, []);

  const handleUpdate = useCallback((idx, stepKey, field, value) => {
    setSession(prev => {
      const students = [...prev.students];
      const s = { ...students[idx], scores: { ...students[idx].scores } };
      s.scores[stepKey] = { ...s.scores[stepKey], [field]: value };
      students[idx] = s;
      return { ...prev, students };
    });
  }, []);

  const handleApprove = useCallback(idx => {
    setSession(prev => { const s = [...prev.students]; s[idx] = { ...s[idx], approved: true }; return { ...prev, students: s }; });
  }, []);

  const handleApproveAll = useCallback(() => {
    setSession(prev => ({ ...prev, students: prev.students.map(s => s.scores ? { ...s, approved: true } : s) }));
  }, []);

  const handleLoadDemo = useCallback(() => {
    setSession(prev => ({ ...prev, students: [...prev.students, ...DEMO.map(d => ({ ...d, approved: false }))] }));
  }, []);

  if (!role) return <RoleSelect onSelect={setRole} />;

  return role === "teacher" ? (
    <TeacherView session={session} onUpdateStudent={handleUpdate} onApprove={handleApprove} onApproveAll={handleApproveAll} onLoadDemo={handleLoadDemo} onBack={() => setRole(null)} />
  ) : (
    <StudentView session={session} onSubmit={handleSubmit} onBack={() => setRole(null)} />
  );
}
