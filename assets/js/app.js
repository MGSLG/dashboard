/* global XLSX */
(() => {
  const XLSX_PATH = "https://mgslg.github.io/dashboard/assets/data/trainings.xlsx";
  const SHEET_NAME = "Trainings";

  // --- UI Elements ---
  const calendarGrid = document.getElementById("calendarGrid");
  const monthLabel = document.getElementById("monthLabel");
  const prevMonthBtn = document.getElementById("prevMonthBtn");
  const nextMonthBtn = document.getElementById("nextMonthBtn");
  const enrollList = document.getElementById("enrollList");
  const upcomingList = document.getElementById("upcomingList");
  const workshopStrip = document.getElementById("workshopStrip");
  const reloadBtn = document.getElementById("reloadBtn");

  const menuBtn = document.getElementById("menuBtn");
  const nav = document.getElementById("nav");

  const modalBackdrop = document.getElementById("modalBackdrop");
  const closeModalBtn = document.getElementById("closeModalBtn");
  const modalTitle = document.getElementById("modalTitle");
  const modalSubtitle = document.getElementById("modalSubtitle");
  const modalMeta = document.getElementById("modalMeta");
  const modalGallery = document.getElementById("modalGallery");
  const modalActions = document.getElementById("modalActions");
  const modalDesc = document.getElementById("modalDesc");

  // --- State ---
  let rows = [];
  let currentMonth = new Date();
  currentMonth.setDate(1);

  // --- Helpers ---
  const clamp = (n, min, max) => Math.max(min, Math.min(max, n));

  function parseISODate(dateStr) {
    // expects YYYY-MM-DD
    const [y, m, d] = String(dateStr).split("-").map(Number);
    if (!y || !m || !d) return null;
    const dt = new Date(Date.UTC(y, m - 1, d));
    // display in local, but keep consistent
    return new Date(dt.getUTCFullYear(), dt.getUTCMonth(), dt.getUTCDate());
  }

  function formatDateShort(dt) {
    return dt.toLocaleDateString(undefined, { month: "short", day: "numeric" });
  }

  function formatMonthTitle(dt) {
    return dt.toLocaleDateString(undefined, { month: "long", year: "numeric" });
  }

  function normalizeNumber(v) {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }

  function safeStr(v) {
    return (v === null || v === undefined) ? "" : String(v).trim();
  }

  function splitCSVPaths(v) {
    const s = safeStr(v);
    if (!s) return [];
    return s.split(",").map(x => x.trim()).filter(Boolean);
  }

  function tagColor(workshopName, programName) {
    const t = (workshopName + " " + programName).toLowerCase();
    if (t.includes("safety")) return "event-orange";
    if (t.includes("skill") || t.includes("ict") || t.includes("it")) return "event-green";
    return "event-blue";
  }

  // --- Data Loading (Excel) ---
  async function loadExcel() {
    const res = await fetch(XLSX_PATH, { cache: "no-store" });
    if (!res.ok) {
      throw new Error(`Could not load ${XLSX_PATH}. Make sure it exists and GitHub Pages is serving it.`);
    }
    const arrayBuffer = await res.arrayBuffer();
    const wb = XLSX.read(arrayBuffer, { type: "array" });
    const ws = wb.Sheets[SHEET_NAME] || wb.Sheets[wb.SheetNames[0]];
    const json = XLSX.utils.sheet_to_json(ws, { defval: "" });

    // Normalize rows
    rows = json
      .map(r => {
        const dt = parseISODate(r.Date);
        return {
          date: dt,
          dateRaw: safeStr(r.Date),
          startTime: safeStr(r.StartTime),
          endTime: safeStr(r.EndTime),
          workshopName: safeStr(r.WorkshopName),
          programName: safeStr(r.ProgramName),
          location: safeStr(r.Location),
          enrolled: normalizeNumber(r.Enrolled),
          expected: normalizeNumber(r.Expected),
          description: safeStr(r.Description),
          thumb: safeStr(r.Thumb),
          gallery: splitCSVPaths(r.Gallery),
          manual: safeStr(r.Manual),
          video: safeStr(r.Video),
        };
      })
      .filter(r => r.date && r.workshopName);

    // Sort by date ascending
    rows.sort((a, b) => a.date - b.date);
  }

  // --- Calendar Rendering ---
  function renderCalendar() {
    calendarGrid.innerHTML = "";

    // Day headers
    const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    dayNames.forEach(d => {
      const el = document.createElement("div");
      el.className = "day-head";
      el.textContent = d;
      calendarGrid.appendChild(el);
    });

    monthLabel.textContent = formatMonthTitle(currentMonth);

    const year = currentMonth.getFullYear();
    const month = currentMonth.getMonth();

    const firstDay = new Date(year, month, 1);
    const startWeekday = firstDay.getDay(); // 0..6
    const daysInMonth = new Date(year, month + 1, 0).getDate();

    // We’ll render 6 weeks for stable layout
    const totalCells = 6 * 7;
    const startDate = new Date(year, month, 1 - startWeekday);

    for (let i = 0; i < totalCells; i++) {
      const cellDate = new Date(startDate);
      cellDate.setDate(startDate.getDate() + i);

      const cell = document.createElement("div");
      cell.className = "day-cell";

      // Muted for other months
      if (cellDate.getMonth() !== month) cell.classList.add("muted");

      const num = document.createElement("div");
      num.className = "day-num";
      num.textContent = String(cellDate.getDate());
      cell.appendChild(num);

      // Events on this date
      const events = rows.filter(r =>
        r.date.getFullYear() === cellDate.getFullYear() &&
        r.date.getMonth() === cellDate.getMonth() &&
        r.date.getDate() === cellDate.getDate()
      );

      // Show up to 2 pills, then “+N”
      events.slice(0, 2).forEach(ev => {
        const pill = document.createElement("div");
        pill.className = `event-pill ${tagColor(ev.workshopName, ev.programName)}`;
        pill.textContent = ev.workshopName;
        pill.title = ev.workshopName;
        pill.addEventListener("click", (e) => {
          e.stopPropagation();
          openWorkshopModal(ev);
        });
        cell.appendChild(pill);
      });

      if (events.length > 2) {
        const more = document.createElement("div");
        more.className = "event-pill event-blue";
        more.textContent = `+${events.length - 2} more`;
        more.addEventListener("click", (e) => {
          e.stopPropagation();
          openDayListModal(cellDate, events);
        });
        cell.appendChild(more);
      }

      // Tap cell: show day events
      cell.addEventListener("click", () => {
        if (events.length) openDayListModal(cellDate, events);
      });

      calendarGrid.appendChild(cell);
    }
  }

  // --- Enrollment Rendering ---
  function renderEnrollments() {
    enrollList.innerHTML = "";

    const now = new Date();
    const upcoming = rows.filter(r => r.date >= new Date(now.getFullYear(), now.getMonth(), now.getDate()));
    const top = upcoming.slice(0, 6); // show top 6

    if (!top.length) {
      enrollList.innerHTML = `<div class="hint">No upcoming trainings found.</div>`;
      return;
    }

    top.forEach(r => {
      const pct = r.expected > 0 ? clamp((r.enrolled / r.expected) * 100, 0, 100) : 0;
      const color = tagColor(r.workshopName, r.programName)
        .replace("event-", ""); // blue/orange/green

      const item = document.createElement("div");
      item.className = "enroll-item";

      const topRow = document.createElement("div");
      topRow.className = "enroll-top";
      topRow.innerHTML = `
        <div class="enroll-name">${escapeHtml(r.workshopName)}</div>
        <div class="enroll-meta">${r.enrolled} / ${r.expected} (${Math.round(pct)}%)</div>
      `;

      const bar = document.createElement("div");
      bar.className = "bar";
      const fill = document.createElement("div");
      fill.style.width = `${pct}%`;
      fill.style.background = colorToHex(color);
      bar.appendChild(fill);

      item.appendChild(topRow);
      item.appendChild(bar);

      item.addEventListener("click", () => openWorkshopModal(r));
      enrollList.appendChild(item);
    });
  }

  function colorToHex(colorName) {
    if (colorName === "orange") return "#f97316";
    if (colorName === "green") return "#16a34a";
    return "#2563eb";
  }

  // --- Upcoming List Rendering ---
  function renderUpcoming() {
    upcomingList.innerHTML = "";

    const now = new Date();
    const upcoming = rows
      .filter(r => r.date >= new Date(now.getFullYear(), now.getMonth(), now.getDate()))
      .slice(0, 8);

    if (!upcoming.length) {
      upcomingList.innerHTML = `<div class="hint">No upcoming trainings found.</div>`;
      return;
    }

    upcoming.forEach(r => {
      const el = document.createElement("div");
      el.className = "up-item";
      el.innerHTML = `
        <div class="up-date">${escapeHtml(formatDateShort(r.date))}</div>
        <div class="up-name">${escapeHtml(r.workshopName)}</div>
        <div class="up-sub">${escapeHtml(`${r.startTime}–${r.endTime}`)} • ${escapeHtml(r.location || "TBA")}</div>
      `;
      el.addEventListener("click", () => openWorkshopModal(r));
      upcomingList.appendChild(el);
    });
  }

  // --- Workshop Cards (Scrollable) ---
  function renderWorkshopCards() {
    workshopStrip.innerHTML = "";

    // Group by workshopName (if multiple dates exist, pick next occurrence)
    const now = new Date();
    const grouped = new Map();
    rows.forEach(r => {
      const key = r.workshopName.toLowerCase();
      const existing = grouped.get(key);
      if (!existing) grouped.set(key, r);
      else {
        // prefer upcoming nearest, else earliest
        const existingScore = scoreForNearest(existing, now);
        const rScore = scoreForNearest(r, now);
        if (rScore < existingScore) grouped.set(key, r);
      }
    });

    const list = Array.from(grouped.values())
      .sort((a, b) => scoreForNearest(a, now) - scoreForNearest(b, now));

    list.forEach(r => {
      const card = document.createElement("div");
      card.className = "workshop-card";

      const thumb = document.createElement("div");
      thumb.className = "workshop-thumb";
      if (r.thumb) thumb.style.backgroundImage = `url('${r.thumb}')`;

      const body = document.createElement("div");
      body.className = "workshop-body";
      body.innerHTML = `
        <div class="workshop-name">${escapeHtml(r.workshopName)}</div>
        <div class="workshop-meta">
          <span>${escapeHtml(r.programName || "Programme")}</span>
          <span>${escapeHtml(`${r.enrolled}/${r.expected}`)}</span>
        </div>
      `;

      card.appendChild(thumb);
      card.appendChild(body);
      card.addEventListener("click", () => openWorkshopModal(r));
      workshopStrip.appendChild(card);
    });
  }

  function scoreForNearest(r, now) {
    const d = new Date(r.date);
    const delta = d - now;
    // Upcoming gets priority; past pushed far
    if (delta >= 0) return delta;
    return Math.abs(delta) + 365 * 24 * 60 * 60 * 1000; // push past behind
  }

  // --- Modal ---
  function openWorkshopModal(r) {
    modalTitle.textContent = r.workshopName;
    modalSubtitle.textContent = `${r.programName || "Programme"} • ${formatDateShort(r.date)} • ${r.startTime || "TBA"}–${r.endTime || "TBA"}`;

    modalMeta.innerHTML = "";
    const pct = r.expected > 0 ? clamp((r.enrolled / r.expected) * 100, 0, 100) : 0;

    addChip(`Location: ${r.location || "TBA"}`);
    addChip(`Enrollments: ${r.enrolled}/${r.expected} (${Math.round(pct)}%)`);

    modalGallery.innerHTML = "";
    const images = r.gallery.length ? r.gallery : (r.thumb ? [r.thumb] : []);
    if (images.length) {
      images.forEach(src => {
        const img = document.createElement("img");
        img.src = src;
        img.alt = r.workshopName;
        modalGallery.appendChild(img);
      });
    } else {
      modalGallery.innerHTML = `<div class="hint">No gallery images linked for this workshop.</div>`;
    }

    modalActions.innerHTML = "";
    if (r.manual) modalActions.appendChild(makeLinkBtn("Manual (PDF)", r.manual));
    if (r.video) modalActions.appendChild(makeLinkBtn("Video", r.video));
    if (!r.manual && !r.video) {
      modalActions.innerHTML = `<div class="hint">No manual/video links provided yet.</div>`;
    }

    modalDesc.textContent = r.description || "No description provided.";

    modalBackdrop.classList.add("open");
    modalBackdrop.setAttribute("aria-hidden", "false");
  }

  function openDayListModal(dateObj, events) {
    // Simple: show first event details, and list other events as clickable chips
    const title = dateObj.toLocaleDateString(undefined, { weekday: "long", month: "long", day: "numeric", year: "numeric" });
    modalTitle.textContent = `Trainings on ${title}`;
    modalSubtitle.textContent = `${events.length} event(s)`;

    modalMeta.innerHTML = "";
    modalGallery.innerHTML = "";
    modalActions.innerHTML = "";
    modalDesc.innerHTML = "";

    const listWrap = document.createElement("div");
    listWrap.style.display = "flex";
    listWrap.style.flexDirection = "column";
    listWrap.style.gap = "10px";

    events.forEach(ev => {
      const btn = document.createElement("button");
      btn.className = "btn";
      btn.style.textAlign = "left";
      btn.innerHTML = `<strong>${escapeHtml(ev.workshopName)}</strong><br><span class="hint">${escapeHtml(ev.startTime)}–${escapeHtml(ev.endTime)} • ${escapeHtml(ev.location || "TBA")}</span>`;
      btn.addEventListener("click", () => openWorkshopModal(ev));
      listWrap.appendChild(btn);
    });

    modalDesc.appendChild(listWrap);

    modalBackdrop.classList.add("open");
    modalBackdrop.setAttribute("aria-hidden", "false");
  }

  function closeModal() {
    modalBackdrop.classList.remove("open");
    modalBackdrop.setAttribute("aria-hidden", "true");
  }

  function addChip(text) {
    const el = document.createElement("div");
    el.className = "chip";
    el.textContent = text;
    modalMeta.appendChild(el);
  }

  function makeLinkBtn(label, href) {
    const a = document.createElement("a");
    a.className = "link-btn";
    a.href = href;
    a.target = "_blank";
    a.rel = "noopener";
    a.textContent = label;
    return a;
  }

  // --- Security: basic escaping for HTML injection from Excel ---
  function escapeHtml(str) {
    return String(str)
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  // --- Events ---
  prevMonthBtn.addEventListener("click", () => {
    currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() - 1, 1);
    renderAll();
  });

  nextMonthBtn.addEventListener("click", () => {
    currentMonth = new Date(currentMonth.getFullYear(), currentMonth.getMonth() + 1, 1);
    renderAll();
  });

  reloadBtn.addEventListener("click", async () => {
    await init();
  });

  menuBtn.addEventListener("click", () => {
    nav.classList.toggle("open");
  });

  closeModalBtn.addEventListener("click", closeModal);
  modalBackdrop.addEventListener("click", (e) => {
    if (e.target === modalBackdrop) closeModal();
  });
  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeModal();
  });

  function renderAll() {
    renderCalendar();
    renderEnrollments();
    renderUpcoming();
    renderWorkshopCards();
  }

  async function init() {
    try {
      await loadExcel();
      // Default calendar month to current month for first render
      const now = new Date();
      currentMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      renderAll();
    } catch (err) {
      console.error(err);
      calendarGrid.innerHTML = `<div class="hint">Failed to load Excel data. Check console and verify <code>${XLSX_PATH}</code> exists.</div>`;
      enrollList.innerHTML = `<div class="hint">No data loaded.</div>`;
      upcomingList.innerHTML = `<div class="hint">No data loaded.</div>`;
      workshopStrip.innerHTML = `<div class="hint">No data loaded.</div>`;
    }
  }

  init();
})();
