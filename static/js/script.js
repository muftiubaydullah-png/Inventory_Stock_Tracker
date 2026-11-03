/* Inventory Stock Tracker - front-end behaviour (plain JavaScript, no libraries) */

/* ---------- Chart helpers (drawn with SVG / CSS, so they work offline) ---------- */
const STATUS_COLORS = {
  "In Stock": "#15803d", "Low Stock": "#d97706",
  "Reorder Required": "#ea580c", "Out of Stock": "#dc2626"
};

function formatINR(v) {
  if (v >= 1e7) return "₹" + (v / 1e7).toFixed(2) + " Cr";
  if (v >= 1e5) return "₹" + (v / 1e5).toFixed(2) + " L";
  return "₹" + Math.round(v).toLocaleString("en-IN");
}

/* Donut chart + legend. data = [{label, value}] */
function drawDonut(chartId, legendId, data) {
  const el = document.getElementById(chartId);
  const legend = document.getElementById(legendId);
  if (!el) return;
  const total = data.reduce((s, d) => s + d.value, 0);
  const r = 70, c = 2 * Math.PI * r;
  let offset = 0;
  let rings = "";
  data.forEach(d => {
    if (!total || !d.value) return;
    const len = (d.value / total) * c;
    rings += `<circle cx="100" cy="100" r="${r}" fill="none" stroke="${STATUS_COLORS[d.label] || "#64748b"}"
      stroke-width="28" stroke-dasharray="${len} ${c - len}" stroke-dashoffset="${-offset}"
      transform="rotate(-90 100 100)"><title>${d.label}: ${d.value}</title></circle>`;
    offset += len;
  });
  el.innerHTML = `<svg viewBox="0 0 200 200" width="100%" height="100%">
      <circle cx="100" cy="100" r="${r}" fill="none" stroke="#eef2f4" stroke-width="28"/>${rings}
      <text x="100" y="98" text-anchor="middle" class="donut-total">${total.toLocaleString("en-IN")}</text>
      <text x="100" y="118" text-anchor="middle" class="donut-label">items</text></svg>`;
  if (legend) {
    legend.innerHTML = data.map(d => {
      const pct = total ? ((d.value / total) * 100).toFixed(1) : "0.0";
      return `<li><span class="swatch" style="background:${STATUS_COLORS[d.label] || "#64748b"}"></span>
        <span class="lg-name">${d.label}</span><span class="lg-val">${d.value.toLocaleString("en-IN")} <em>${pct}%</em></span></li>`;
    }).join("");
  }
}

/* Horizontal bar chart. data = [{label, value}] */
function drawBars(chartId, data, opts) {
  const el = document.getElementById(chartId);
  if (!el) return;
  const max = Math.max(...data.map(d => d.value), 1);
  const fmt = (opts && opts.money) ? formatINR : v => v.toLocaleString("en-IN");
  el.innerHTML = data.map(d => `
    <div class="hbar">
      <span class="hbar-label" title="${d.label}">${d.label}</span>
      <span class="hbar-track"><span class="hbar-fill" style="width:${(d.value / max * 100).toFixed(1)}%"></span></span>
      <span class="hbar-value">${fmt(d.value)}</span>
    </div>`).join("");
}

/* ---------- Everything else runs once the page is ready ---------- */
document.addEventListener("DOMContentLoaded", () => {

  /* Mobile sidebar */
  const sidebar = document.getElementById("sidebar");
  const scrim = document.getElementById("scrim");
  const menuBtn = document.getElementById("menuBtn");
  const toggleMenu = open => {
    sidebar.classList.toggle("open", open);
    scrim.classList.toggle("show", open);
  };
  if (menuBtn) menuBtn.addEventListener("click", () => toggleMenu(!sidebar.classList.contains("open")));
  if (scrim) scrim.addEventListener("click", () => toggleMenu(false));

  /* Flash messages: close button, and success messages fade away */
  document.querySelectorAll(".flash").forEach(f => {
    f.querySelector(".flash-close").addEventListener("click", () => f.remove());
    if (f.classList.contains("flash-success")) setTimeout(() => f.remove(), 6000);
  });

  /* Filter drop-downs submit the form straight away */
  document.querySelectorAll(".js-autosubmit").forEach(sel =>
    sel.addEventListener("change", () => sel.form.submit()));

  /* Delete confirmation dialog */
  const modal = document.getElementById("deleteModal");
  let pendingForm = null;
  const closeModal = () => { modal.hidden = true; pendingForm = null; };
  document.querySelectorAll(".js-delete-form").forEach(form => {
    form.addEventListener("submit", e => {
      e.preventDefault();
      pendingForm = form;
      document.getElementById("deleteName").textContent = form.dataset.name;
      document.getElementById("deleteId").textContent = form.dataset.id;
      modal.hidden = false;
      document.getElementById("deleteCancel").focus();
    });
  });
  if (modal) {
    document.getElementById("deleteCancel").addEventListener("click", closeModal);
    document.getElementById("deleteConfirm").addEventListener("click", () => {
      if (pendingForm) pendingForm.submit();
    });
    modal.addEventListener("click", e => { if (e.target === modal) closeModal(); });
    document.addEventListener("keydown", e => { if (e.key === "Escape" && !modal.hidden) closeModal(); });
  }

  /* Add Item & Edit Item forms: live stock value calculation + double submit prevention */
  const itemForm = document.getElementById("addForm") || document.getElementById("editForm");
  if (itemForm) {
    const qty = itemForm.querySelector("#qty");
    const cost = itemForm.querySelector("#cost");
    const preview = document.getElementById("valuePreview");
    const update = () => {
      if (!qty || !cost || !preview) return;
      const q = parseFloat(qty.value);
      const c = parseFloat(cost.value);
      const valid = !isNaN(q) && q >= 0 && !isNaN(c) && c >= 0;
      const v = valid ? q * c : 0;
      preview.value = "₹" + v.toLocaleString("en-IN", {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
      });
    };
    if (qty && cost) {
      qty.addEventListener("input", update);
      cost.addEventListener("input", update);
      qty.addEventListener("change", update);
      cost.addEventListener("change", update);
      update();
    }
    itemForm.addEventListener("submit", () => {
      if (itemForm.checkValidity()) {
        const b = itemForm.querySelector("button[type=submit]");
        if (b) setTimeout(() => { b.disabled = true; }, 0);
      }
    });

    const optAccordion = document.getElementById("optionalAccordion");
    if (optAccordion && optAccordion.querySelector(".has-error")) {
      optAccordion.open = true;
    }
  }

  /* Update Stock: live suggestions and result preview */
  const search = document.getElementById("itemSearch");
  const list = document.getElementById("itemResults");
  if (search && list) {
    let timer = null;
    const hide = () => { list.hidden = true; list.innerHTML = ""; };
    search.addEventListener("input", () => {
      clearTimeout(timer);
      const q = search.value.trim();
      if (q.length < 2) return hide();
      timer = setTimeout(async () => {
        try {
          const res = await fetch("/api/items/search?q=" + encodeURIComponent(q));
          const items = await res.json();
          if (!items.length) { list.innerHTML = '<li class="none">No matches</li>'; list.hidden = false; return; }
          list.innerHTML = items.map(i => `<li><a href="/update-stock?item_id=${encodeURIComponent(i.item_id)}">
              <span class="mono">${i.item_id}</span> <strong>${i.name}</strong>
              <span class="muted small">${i.brand} · ${i.qty} ${i.unit}</span></a></li>`).join("");
          list.hidden = false;
        } catch (err) { hide(); }
      }, 200);
    });
    document.addEventListener("click", e => { if (!list.contains(e.target) && e.target !== search) hide(); });
  }

  const stockForm = document.getElementById("stockForm");
  if (stockForm) {
    const q = stockForm.querySelector("#quantity");
    const out = document.getElementById("stockPreview");
    const submit = document.getElementById("stockSubmit");
    const current = parseInt(q.dataset.current, 10);
    const refresh = () => {
      const amount = parseInt(q.value, 10);
      const isOut = stockForm.querySelector("#act_out").checked;
      submit.disabled = false;
      out.className = "preview";
      if (!amount || amount < 1) { out.textContent = "Enter a quantity to preview the new stock level."; return; }
      if (isOut && amount > current) {
        out.textContent = `Not enough stock: only ${current} available.`;
        out.className = "preview bad";
        submit.disabled = true;
        return;
      }
      const next = isOut ? current - amount : current + amount;
      out.textContent = `Stock will change from ${current} to ${next}.`;
    };
    q.addEventListener("input", refresh);
    stockForm.querySelectorAll("input[name=action]").forEach(r => r.addEventListener("change", refresh));
    stockForm.addEventListener("submit", () => setTimeout(() => { submit.disabled = true; }, 0));
    refresh();
  }
});
