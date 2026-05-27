/* ── Admin Panel JS ─────────────────────────────────────────── */
"use strict";

const API = "";

// ─── Auth guard ──────────────────────────────────────────────
async function checkAdminAccess() {
  const token = localStorage.getItem("token");
  if (!token) return false;
  try {
    const res = await fetch(`${API}/api/auth/profile`, {
      headers: { Authorization: `Bearer ${token}` },
    });
    if (!res.ok) return false;
    const profile = await res.json();
    if (profile.role !== "admin") return false;
    document.getElementById("adminName").textContent =
      profile.name || profile.email || "Admin";
    return true;
  } catch {
    return false;
  }
}

// ─── Helpers ─────────────────────────────────────────────────
function authHeaders() {
  return {
    Authorization: `Bearer ${localStorage.getItem("token")}`,
    "Content-Type": "application/json",
  };
}

function fmt(dt) {
  if (!dt) return "—";
  return new Date(dt).toLocaleString("vi-VN", {
    dateStyle: "short",
    timeStyle: "short",
  });
}

function esc(str) {
  return String(str ?? "").replace(
    /[&<>"']/g,
    (m) =>
      ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[
        m
      ],
  );
}

function badge(text, cls) {
  return `<span class="badge badge-${esc(cls)}">${esc(text)}</span>`;
}

function planBadge(plan) {
  const p = (plan || "free").toLowerCase();
  return badge(p.toUpperCase(), p);
}

function roleBadge(role) {
  const r = (role || "user").toLowerCase();
  return badge(r === "admin" ? "Admin" : "User", r);
}

function fmtVnd(amount) {
  return Number(amount || 0).toLocaleString("vi-VN") + " ₫";
}

let toastTimer;
function showToast(msg, type = "success") {
  const el = document.getElementById("adminToast");
  el.textContent = msg;
  el.className = "admin-toast show" + (type === "error" ? " error" : "");
  clearTimeout(toastTimer);
  toastTimer = setTimeout(() => {
    el.className = "admin-toast";
  }, 3000);
}

function closeModal(id) {
  document.getElementById(id).style.display = "none";
}

function openModal(id) {
  document.getElementById(id).style.display = "flex";
}

function buildPagination(containerId, currentPage, totalPages, loadFn) {
  const el = document.getElementById(containerId);
  if (totalPages <= 1) {
    el.innerHTML = "";
    return;
  }
  let html = "";
  const start = Math.max(1, currentPage - 2);
  const end = Math.min(totalPages, currentPage + 2);
  if (start > 1)
    html += `<button class="page-btn" onclick="${loadFn}(1)">1</button>`;
  if (start > 2)
    html += `<span style="color:var(--muted);padding:0 4px">…</span>`;
  for (let i = start; i <= end; i++) {
    html += `<button class="page-btn${i === currentPage ? " active" : ""}" onclick="${loadFn}(${i})">${i}</button>`;
  }
  if (end < totalPages - 1)
    html += `<span style="color:var(--muted);padding:0 4px">…</span>`;
  if (end < totalPages)
    html += `<button class="page-btn" onclick="${loadFn}(${totalPages})">${totalPages}</button>`;
  el.innerHTML = html;
}

// ─── Stats ───────────────────────────────────────────────────
async function loadStats() {
  try {
    const res = await fetch(`${API}/api/admin/stats`, {
      headers: authHeaders(),
    });
    if (!res.ok) return;
    const d = await res.json();
    document.getElementById("stat-users").textContent = d.total_users ?? "—";
    document.getElementById("stat-translations").textContent =
      d.total_translations ?? "—";
    document.getElementById("stat-payments").textContent =
      d.total_payments ?? "—";
    document.getElementById("stat-revenue").textContent = fmtVnd(
      d.total_revenue_vnd ?? 0,
    );
    document.getElementById("stat-admins").textContent = d.admin_count ?? "—";
    document.getElementById("stat-free").textContent =
      d.plan_distribution?.free ?? "—";
    document.getElementById("stat-pro").textContent =
      d.plan_distribution?.pro ?? "—";
    document.getElementById("stat-promax").textContent =
      d.plan_distribution?.promax ?? "—";
  } catch (e) {
    console.error(e);
  }
}

// ─── Users ───────────────────────────────────────────────────
let _usersPage = 1;
async function loadUsers(page = 1) {
  _usersPage = page;
  const q = (document.getElementById("userSearch")?.value || "").trim();
  const url = `${API}/api/admin/users?page=${page}&per_page=15${q ? "&q=" + encodeURIComponent(q) : ""}`;
  try {
    const res = await fetch(url, { headers: authHeaders() });
    const d = await res.json();
    const tbody = document.getElementById("usersBody");
    tbody.innerHTML =
      (d.users || [])
        .map(
          (u) => `
      <tr>
        <td>${esc(u.id)}</td>
        <td>${esc(u.name || "—")}</td>
        <td>${esc(u.email)}</td>
        <td>${planBadge(u.plan)}</td>
        <td>${roleBadge(u.role)}</td>
        <td>${Number(u.token_balance || 0).toLocaleString()}</td>
        <td>${fmt(u.created_at)}</td>
        <td>
          <button class="btn-icon" title="Xem" onclick="viewUser(${u.id})"><i class="fas fa-eye"></i></button>
          ${
            u.role !== "admin"
              ? `<button class="btn-icon" title="Cấp Admin" onclick="grantAdmin(${u.id})"><i class="fas fa-user-shield"></i></button>`
              : `<button class="btn-icon danger" title="Thu hồi Admin" onclick="revokeAdmin(${u.id})"><i class="fas fa-user-minus"></i></button>`
          }
          <button class="btn-icon danger" title="Xóa" onclick="deleteUser(${u.id}, '${esc(u.email)}')"><i class="fas fa-trash"></i></button>
        </td>
      </tr>`,
        )
        .join("") ||
      '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination("usersPagination", page, d.pages || 1, "loadUsers");
  } catch (e) {
    console.error(e);
    showToast("Lỗi tải danh sách người dùng", "error");
  }
}

async function viewUser(id) {
  try {
    const res = await fetch(`${API}/api/admin/users/${id}`, {
      headers: authHeaders(),
    });
    const u = await res.json();
    document.getElementById("modalUserTitle").textContent =
      `Người dùng #${u.id}`;
    document.getElementById("modalUserBody").innerHTML = `
      <div class="detail-row"><span class="detail-label">ID</span><span class="detail-value">${esc(u.id)}</span></div>
      <div class="detail-row"><span class="detail-label">Email</span><span class="detail-value">${esc(u.email)}</span></div>
      <div class="detail-row"><span class="detail-label">Tên</span><span class="detail-value">${esc(u.name || "—")}</span></div>
      <div class="detail-row"><span class="detail-label">Plan</span><span class="detail-value">${planBadge(u.plan)}</span></div>
      <div class="detail-row"><span class="detail-label">Role</span><span class="detail-value">${roleBadge(u.role)}</span></div>
      <div class="detail-row"><span class="detail-label">Tokens</span><span class="detail-value">${Number(u.token_balance || 0).toLocaleString()}</span></div>
      <div class="detail-row"><span class="detail-label">Google ID</span><span class="detail-value">${esc(u.google_id || "Không có (email/pass)")}</span></div>
      <div class="detail-row"><span class="detail-label">Ngày tạo</span><span class="detail-value">${fmt(u.created_at)}</span></div>
    `;
    document.getElementById("modalUserFooter").innerHTML = `
      <button class="btn-icon" onclick="closeModal('userModal')">Đóng</button>
      ${
        u.role !== "admin"
          ? `<button class="btn-accent" onclick="grantAdmin(${u.id}); closeModal('userModal')"><i class="fas fa-user-shield"></i> Cấp Admin</button>`
          : `<button class="btn-icon danger" onclick="revokeAdmin(${u.id}); closeModal('userModal')"><i class="fas fa-user-minus"></i> Thu hồi Admin</button>`
      }
    `;
    openModal("userModal");
  } catch (e) {
    showToast("Lỗi tải thông tin", "error");
  }
}

async function grantAdmin(id) {
  if (!confirm("Cấp quyền Admin cho người dùng này?")) return;
  const res = await fetch(`${API}/api/admin/users/${id}/grant-admin`, {
    method: "POST",
    headers: authHeaders(),
  });
  if (res.ok) {
    showToast("Đã cấp quyền Admin");
    loadUsers(_usersPage);
  } else showToast("Lỗi", "error");
}

async function revokeAdmin(id) {
  if (!confirm("Thu hồi quyền Admin?")) return;
  const res = await fetch(`${API}/api/admin/users/${id}/revoke-admin`, {
    method: "POST",
    headers: authHeaders(),
  });
  if (res.ok) {
    showToast("Đã thu hồi quyền Admin");
    loadUsers(_usersPage);
  } else showToast("Lỗi", "error");
}

async function deleteUser(id, email) {
  if (!confirm(`Xóa người dùng "${email}"? Hành động này không thể hoàn tác.`))
    return;
  const res = await fetch(`${API}/api/admin/users/${id}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  if (res.ok) {
    showToast("Đã xóa người dùng");
    loadUsers(_usersPage);
  } else {
    const d = await res.json();
    showToast(d.error || "Lỗi xóa", "error");
  }
}

// ─── Translations ─────────────────────────────────────────────
let _transPage = 1;
async function loadTranslations(page = 1) {
  _transPage = page;
  try {
    const res = await fetch(
      `${API}/api/admin/translations?page=${page}&per_page=15`,
      { headers: authHeaders() },
    );
    const d = await res.json();
    const tbody = document.getElementById("translationsBody");
    tbody.innerHTML =
      (d.translations || [])
        .map(
          (t) => `
      <tr>
        <td>${esc(t.id)}</td>
        <td>${esc(t.user_id)}</td>
        <td>${esc(t.source_lang || "—")}</td>
        <td>${esc(t.target_lang || "—")}</td>
        <td title="${esc(t.original_text)}">${esc((t.original_text || "—").substring(0, 60))}…</td>
        <td>${fmt(t.created_at)}</td>
        <td>
          <button class="btn-icon danger" title="Xóa" onclick="deleteTranslation(${t.id})"><i class="fas fa-trash"></i></button>
        </td>
      </tr>`,
        )
        .join("") ||
      '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination(
      "translationsPagination",
      page,
      d.pages || 1,
      "loadTranslations",
    );
  } catch (e) {
    showToast("Lỗi tải bản dịch", "error");
  }
}

async function deleteTranslation(id) {
  if (!confirm("Xóa bản dịch này?")) return;
  const res = await fetch(`${API}/api/admin/translations/${id}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  if (res.ok) {
    showToast("Đã xóa");
    loadTranslations(_transPage);
  } else showToast("Lỗi", "error");
}

// ─── Payments ─────────────────────────────────────────────────
let _payPage = 1;
async function loadPayments(page = 1) {
  _payPage = page;
  try {
    const res = await fetch(
      `${API}/api/admin/payments?page=${page}&per_page=15`,
      { headers: authHeaders() },
    );
    const d = await res.json();
    const tbody = document.getElementById("paymentsBody");
    tbody.innerHTML =
      (d.payments || [])
        .map((p) => {
          const status = p.status || "pending";
          return `<tr>
        <td>${esc(p.id)}</td>
        <td>${esc(p.user_id)}</td>
        <td>${planBadge(p.plan_type || p.plan || "")}</td>
        <td>${fmtVnd(p.amount)}</td>
        <td>${badge(status, status)}</td>
        <td>${esc(p.sepay_transaction_id || "—")}</td>
        <td>${fmt(p.created_at)}</td>
        <td>
          ${status !== "completed" ? `<button class="btn-icon" title="Đánh dấu hoàn thành" onclick="markPayment(${p.id},'completed')"><i class="fas fa-check"></i></button>` : ""}
          ${status !== "failed" ? `<button class="btn-icon danger" title="Đánh dấu thất bại" onclick="markPayment(${p.id},'failed')"><i class="fas fa-times"></i></button>` : ""}
        </td>
      </tr>`;
        })
        .join("") ||
      '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination("paymentsPagination", page, d.pages || 1, "loadPayments");
  } catch (e) {
    showToast("Lỗi tải thanh toán", "error");
  }
}

async function markPayment(id, status) {
  const res = await fetch(`${API}/api/admin/payments/${id}`, {
    method: "PATCH",
    headers: authHeaders(),
    body: JSON.stringify({ status }),
  });
  if (res.ok) {
    showToast(`Đã cập nhật trạng thái: ${status}`);
    loadPayments(_payPage);
  } else showToast("Lỗi cập nhật", "error");
}

// ─── Contacts ─────────────────────────────────────────────────
let _contactPage = 1;
async function loadContacts(page = 1) {
  _contactPage = page;
  try {
    const res = await fetch(
      `${API}/api/admin/contacts?page=${page}&per_page=15`,
      { headers: authHeaders() },
    );
    const d = await res.json();
    const tbody = document.getElementById("contactsBody");
    tbody.innerHTML =
      (d.contacts || [])
        .map((c) => {
          const status = c.status || "unread";
          return `<tr>
        <td>${esc(c.id)}</td>
        <td>${esc((c.first_name || "") + " " + (c.last_name || ""))}</td>
        <td>${esc(c.email)}</td>
        <td>${esc(c.subject || "—")}</td>
        <td title="${esc(c.message)}">${esc((c.message || "").substring(0, 50))}…</td>
        <td>${badge(status, status)}</td>
        <td>${fmt(c.created_at)}</td>
        <td>
          <button class="btn-icon" title="Xem" onclick="viewContact(${c.id})"><i class="fas fa-eye"></i></button>
          <button class="btn-icon danger" title="Xóa" onclick="deleteContact(${c.id})"><i class="fas fa-trash"></i></button>
        </td>
      </tr>`;
        })
        .join("") ||
      '<tr><td colspan="8" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination("contactsPagination", page, d.pages || 1, "loadContacts");
  } catch (e) {
    showToast("Lỗi tải liên hệ", "error");
  }
}

async function viewContact(id) {
  try {
    const res = await fetch(`${API}/api/admin/contacts/${id}`, {
      headers: authHeaders(),
    });
    if (!res.ok) {
      // Fallback: fetch list and find by id
      const listRes = await fetch(
        `${API}/api/admin/contacts?page=1&per_page=200`,
        { headers: authHeaders() },
      );
      const d = await listRes.json();
      const found = (d.contacts || []).find((x) => x.id === id);
      if (!found) {
        showToast("Không tìm thấy", "error");
        return;
      }
      _renderContactModal(found);
    } else {
      const c = await res.json();
      _renderContactModal(c);
    }
  } catch (e) {
    showToast("Lỗi", "error");
  }
}

function _renderContactModal(c) {
  document.getElementById("modalContactBody").innerHTML = `
    <div class="detail-row"><span class="detail-label">Người gửi</span><span class="detail-value">${esc((c.first_name || "") + " " + (c.last_name || ""))}</span></div>
    <div class="detail-row"><span class="detail-label">Email</span><span class="detail-value">${esc(c.email)}</span></div>
    <div class="detail-row"><span class="detail-label">Chủ đề</span><span class="detail-value">${esc(c.subject || "—")}</span></div>
    <div class="detail-row"><span class="detail-label">Trạng thái</span><span class="detail-value">${badge(c.status || "unread", c.status || "unread")}</span></div>
    <div class="detail-row"><span class="detail-label">Thời gian</span><span class="detail-value">${fmt(c.created_at)}</span></div>
    <div class="detail-row"><span class="detail-label">Nội dung</span><span class="detail-value" style="white-space:pre-wrap">${esc(c.message || "—")}</span></div>
    ${
      c.admin_reply
        ? `
    <div class="detail-row" style="margin-top:12px">
      <span class="detail-label" style="color:var(--accent)">Đã trả lời</span>
      <span class="detail-value" style="color:var(--accent);white-space:pre-wrap">${esc(c.admin_reply)}</span>
    </div>`
        : ""
    }
    <div style="margin-top:16px">
      <label style="color:var(--muted);font-size:0.8rem;font-weight:600;display:block;margin-bottom:6px">
        ${c.admin_reply ? "✏️ Chỉnh sửa / gửi lại phản hồi:" : "✉️ Trả lời người dùng:"}
      </label>
      <textarea id="replyText" rows="5" placeholder="Nhập nội dung phản hồi…"
        style="width:100%;background:#0b1a24;border:1px solid var(--border);color:var(--text);
               padding:10px 14px;border-radius:var(--radius);font-size:0.9rem;resize:vertical;
               outline:none;font-family:inherit"
        onfocus="this.style.borderColor='var(--accent)'" onblur="this.style.borderColor='var(--border)'"
      >${esc(c.admin_reply || "")}</textarea>
    </div>
  `;
  document.getElementById("modalContactFooter").innerHTML = `
    <button class="btn-icon" onclick="closeModal('contactModal')">Đóng</button>
    <button class="btn-accent" onclick="sendReply(${c.id})">
      <i class="fas fa-paper-plane"></i> Gửi phản hồi
    </button>
  `;
  openModal("contactModal");
  // Mark as read silently
  if (c.status === "unread") markContact(c.id, "read");
}

async function sendReply(id) {
  const text = (document.getElementById("replyText")?.value || "").trim();
  if (!text) {
    showToast("Vui lòng nhập nội dung phản hồi", "error");
    return;
  }

  const btn = document.querySelector("#modalContactFooter .btn-accent");
  if (btn) {
    btn.disabled = true;
    btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Đang gửi…';
  }

  try {
    const res = await fetch(`${API}/api/admin/contacts/${id}/reply`, {
      method: "POST",
      headers: authHeaders(),
      body: JSON.stringify({ reply: text }),
    });
    const d = await res.json();
    if (!res.ok) {
      showToast(d.error || "Lỗi gửi phản hồi", "error");
      return;
    }
    showToast("Đã gửi phản hồi qua email!");
    closeModal("contactModal");
    loadContacts(_contactPage);
  } catch (e) {
    showToast("Lỗi kết nối", "error");
  } finally {
    if (btn) {
      btn.disabled = false;
      btn.innerHTML = '<i class="fas fa-paper-plane"></i> Gửi phản hồi';
    }
  }
}

async function markContact(id, status) {
  await fetch(`${API}/api/admin/contacts/${id}`, {
    method: "PATCH",
    headers: authHeaders(),
    body: JSON.stringify({ status }),
  });
  loadContacts(_contactPage);
}

async function deleteContact(id) {
  if (!confirm("Xóa tin nhắn này?")) return;
  const res = await fetch(`${API}/api/admin/contacts/${id}`, {
    method: "DELETE",
    headers: authHeaders(),
  });
  if (res.ok) {
    showToast("Đã xóa");
    loadContacts(_contactPage);
  } else showToast("Lỗi", "error");
}

// ─── Newsletter ───────────────────────────────────────────────
let _newsPage = 1;
async function loadNewsletter(page = 1) {
  _newsPage = page;
  try {
    const res = await fetch(
      `${API}/api/admin/newsletter?page=${page}&per_page=20`,
      { headers: authHeaders() },
    );
    const d = await res.json();
    const tbody = document.getElementById("newsletterBody");
    tbody.innerHTML =
      (d.subscribers || [])
        .map(
          (s) => `
      <tr>
        <td>${esc(s.id)}</td>
        <td>${esc(s.email)}</td>
        <td>${badge(s.status || "active", s.status || "active")}</td>
        <td>${fmt(s.created_at)}</td>
      </tr>`,
        )
        .join("") ||
      '<tr><td colspan="4" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination(
      "newsletterPagination",
      page,
      d.pages || 1,
      "loadNewsletter",
    );
  } catch (e) {
    showToast("Lỗi", "error");
  }
}

// ─── Audit ────────────────────────────────────────────────────
let _auditPage = 1;
async function loadAudit(page = 1) {
  _auditPage = page;
  try {
    const res = await fetch(
      `${API}/api/admin/audit-log?page=${page}&per_page=20`,
      { headers: authHeaders() },
    );
    const d = await res.json();
    const tbody = document.getElementById("auditBody");
    tbody.innerHTML =
      (d.logs || [])
        .map((l) => {
          const ACTION_LABELS = {
            grant_admin: "🛡️ Cấp admin",
            revoke_admin: "🔓 Thu hồi admin",
            update_user: "✏️ Sửa user",
            delete_user: "🗑️ Xóa user",
            reply_contact: "💬 Trả lời LH",
            delete_contact: "🗑️ Xóa LH",
            update_payment: "💳 Sửa TT",
            delete_translation: "🗑️ Xóa BD",
          };
          const label = ACTION_LABELS[l.action] || esc(l.action || "—");
          let detail = "";
          try {
            const obj =
              typeof l.detail === "string"
                ? JSON.parse(l.detail)
                : l.detail || {};
            detail = Object.entries(obj)
              .map(([k, v]) => `${k}: ${v}`)
              .join(", ");
          } catch (e) {
            detail = l.detail || "";
          }
          return `
      <tr>
        <td>${esc(l.id)}</td>
        <td style="color:var(--accent);font-size:0.82rem">${esc(l.admin_email || l.admin_id)}</td>
        <td>${label}</td>
        <td>${esc(l.target_type || "—")} ${l.target_id ? "#" + l.target_id : ""}</td>
        <td style="font-size:0.78rem;color:var(--muted)" title="${esc(detail)}">${esc(detail.length > 60 ? detail.slice(0, 60) + "…" : detail)}</td>
        <td>${esc(l.ip_address || "—")}</td>
        <td>${fmt(l.created_at)}</td>
      </tr>`;
        })
        .join("") ||
      '<tr><td colspan="7" style="text-align:center;color:var(--muted);padding:30px">Không có dữ liệu</td></tr>';
    buildPagination("auditPagination", page, d.pages || 1, "loadAudit");
  } catch (e) {
    showToast("Lỗi", "error");
  }
}

// ─── Tab switching ────────────────────────────────────────────
const TAB_LOADERS = {
  dashboard: () => loadStats(),
  users: () => loadUsers(1),
  translations: () => loadTranslations(1),
  payments: () => loadPayments(1),
  contacts: () => loadContacts(1),
  newsletter: () => loadNewsletter(1),
  audit: () => loadAudit(1),
};

const TAB_TITLES = {
  dashboard: "Tổng quan",
  users: "Người dùng",
  translations: "Bản dịch",
  payments: "Thanh toán",
  contacts: "Liên hệ",
  newsletter: "Newsletter",
  audit: "Audit Log",
};

function switchTab(tabName) {
  document
    .querySelectorAll(".admin-tab")
    .forEach((el) => el.classList.remove("active"));
  document
    .querySelectorAll(".sidebar-item")
    .forEach((el) => el.classList.remove("active"));
  const tab = document.getElementById(`tab-${tabName}`);
  if (tab) tab.classList.add("active");
  const link = document.querySelector(`.sidebar-item[data-tab="${tabName}"]`);
  if (link) link.classList.add("active");
  document.getElementById("topbarTitle").textContent =
    TAB_TITLES[tabName] || tabName;
  TAB_LOADERS[tabName]?.();
}

// ─── Init ─────────────────────────────────────────────────────
document.addEventListener("DOMContentLoaded", async () => {
  const loading = document.getElementById("adminLoading");
  const denied = document.getElementById("adminDenied");

  const ok = await checkAdminAccess();
  loading.style.display = "none";

  if (!ok) {
    denied.style.display = "block";
    return;
  }

  // Sidebar links
  document.querySelectorAll(".sidebar-item[data-tab]").forEach((item) => {
    item.addEventListener("click", (e) => {
      e.preventDefault();
      switchTab(item.dataset.tab);
      // close sidebar on mobile
      document.getElementById("sidebar").classList.remove("open");
    });
  });

  // Mobile sidebar toggle
  document.getElementById("sidebarToggle")?.addEventListener("click", () => {
    document.getElementById("sidebar").classList.toggle("open");
  });

  // Close modal on backdrop click
  document.querySelectorAll(".admin-modal-backdrop").forEach((backdrop) => {
    backdrop.addEventListener("click", (e) => {
      if (e.target === backdrop) backdrop.style.display = "none";
    });
  });

  // Enter key on user search
  document.getElementById("userSearch")?.addEventListener("keydown", (e) => {
    if (e.key === "Enter") loadUsers(1);
  });

  // Load default tab
  switchTab("dashboard");
});
