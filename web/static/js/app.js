let allInvoices = [];
let syncInterval = null;

document.addEventListener("DOMContentLoaded", function () {
    loadMailboxes();
    loadData();
});

function formatCurrency(val) {
    if (!val) return "0 ₫";
    return Number(val).toLocaleString("vi-VN") + " ₫";
}

function formatVolume(val) {
    if (!val) return "0 L";
    return Number(val).toLocaleString("vi-VN", { minimumFractionDigits: 2, maximumFractionDigits: 3 }) + " L";
}

function loadMailboxes() {
    fetch("/api/mailboxes")
        .then(res => res.json())
        .then(data => {
            if (data.mailboxes && data.mailboxes.length > 0) {
                const sel = document.getElementById("selectMailbox");
                sel.innerHTML = "";
                data.mailboxes.forEach(mb => {
                    const opt = document.createElement("option");
                    opt.value = mb;
                    opt.textContent = mb;
                    if (mb.toLowerCase().includes("vantai")) {
                        opt.selected = true;
                    }
                    sel.appendChild(opt);
                });
            }
        })
        .catch(err => console.log("Không tải được danh sách mailbox:", err));
}

function loadData() {
    fetch("/api/status")
        .then(res => res.json())
        .then(data => {
            document.getElementById("kpiCount").textContent = data.total_count || 0;
            document.getElementById("kpiAmount").textContent = formatCurrency(data.total_amount);
            document.getElementById("kpiVolume").textContent = formatVolume(data.total_volume);

            allInvoices = data.invoices || [];
            renderTable(allInvoices);

            if (data.is_syncing && !syncInterval) {
                startLogPolling();
            }
        })
        .catch(err => {
            console.error("Lỗi khi tải dữ liệu:", err);
            document.getElementById("invoiceTableBody").innerHTML = `
                <tr><td colspan="12" class="text-center text-danger py-4">Lỗi kết nối máy chủ dữ liệu!</td></tr>
            `;
        });
}

function renderTable(invoices) {
    const tbody = document.getElementById("invoiceTableBody");
    if (!invoices || invoices.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="12" class="text-center py-5 text-muted">
                    <i class="fa-solid fa-folder-open fs-2 mb-2 text-secondary"></i>
                    <div>Chưa có dữ liệu hóa đơn nào trong hệ thống.</div>
                    <div class="small">Nhấn <b>"Bắt Đầu Quét & Tải Hóa Đơn"</b> để thu thập tự động từ Outlook.</div>
                </td>
            </tr>
        `;
        return;
    }

    let html = "";
    invoices.forEach((inv, idx) => {
        const plateHtml = inv.bien_so 
            ? `<span class="plate-badge">${inv.bien_so}</span>` 
            : `<span class="text-muted small">Chưa rõ</span>`;

        const actionButtons = inv.pdf_filename
            ? `<button class="btn btn-sm btn-outline-primary py-0 px-2" onclick="viewPdf('${inv.pdf_filename}', '${inv.so_hd}')">
                 <i class="fa-regular fa-file-pdf me-1"></i>Xem PDF
               </button>`
            : `<span class="text-muted small">Không có PDF</span>`;

        html += `
            <tr>
                <td class="text-center text-muted fw-bold">${idx + 1}</td>
                <td><span class="small fw-semibold">${inv.ngay_hd || ''}</span></td>
                <td><span class="badge bg-light text-dark border">${inv.ky_hieu || ''}-${inv.so_hd || ''}</span></td>
                <td>${plateHtml}</td>
                <td><span class="fuel-badge small">${inv.ten_hang || 'Xăng / Dầu'}</span></td>
                <td class="text-end fw-semibold">${formatVolume(inv.so_luong)}</td>
                <td class="text-end text-muted small">${inv.don_gia ? Number(inv.don_gia).toLocaleString('vi-VN') : ''}</td>
                <td class="text-end text-muted small">${inv.thanh_tien ? Number(inv.thanh_tien).toLocaleString('vi-VN') : ''}</td>
                <td class="text-end fw-bold text-success">${formatCurrency(inv.tong_tien)}</td>
                <td><div class="text-truncate small" style="max-width: 220px;" title="${inv.dia_chi_ban || inv.ben_ban}">${inv.dia_chi_ban || inv.ben_ban}</div></td>
                <td class="text-center"><code class="small text-danger">${inv.fkey || ''}</code></td>
                <td class="text-center">${actionButtons}</td>
            </tr>
        `;
    });
    tbody.innerHTML = html;
}

function filterTable() {
    const q = (document.getElementById("searchInput").value || "").toLowerCase().trim();
    if (!q) {
        renderTable(allInvoices);
        return;
    }

    const filtered = allInvoices.filter(inv => {
        return (inv.bien_so && inv.bien_so.toLowerCase().includes(q)) ||
               (inv.so_hd && inv.so_hd.toLowerCase().includes(q)) ||
               (inv.ky_hieu && inv.ky_hieu.toLowerCase().includes(q)) ||
               (inv.ten_hang && inv.ten_hang.toLowerCase().includes(q)) ||
               (inv.dia_chi_ban && inv.dia_chi_ban.toLowerCase().includes(q)) ||
               (inv.ben_ban && inv.ben_ban.toLowerCase().includes(q)) ||
               (inv.ngay_hd && inv.ngay_hd.toLowerCase().includes(q)) ||
               (inv.fkey && inv.fkey.toLowerCase().includes(q));
    });
    renderTable(filtered);
}

function startSync() {
    const mailbox = document.getElementById("selectMailbox").value;
    const limit = document.getElementById("selectLimit").value;
    const btn = document.getElementById("btnSync");

    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin me-1"></i> Đang Xử Lý...`;

    const container = document.getElementById("syncProgressContainer");
    container.classList.remove("d-none");
    document.getElementById("logBox").innerHTML = "<div class='text-warning'>Khởi động tiến trình đồng bộ...</div>";

    fetch("/api/sync", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ mailbox: mailbox, limit: limit })
    })
    .then(res => res.json())
    .then(data => {
        if (data.status === "started") {
            startLogPolling();
        } else {
            alert(data.message || "Không thể khởi động đồng bộ!");
            btn.disabled = false;
            btn.innerHTML = `<i class="fa-solid fa-cloud-arrow-down me-1"></i> Bắt Đầu Quét & Tải Hóa Đơn`;
        }
    })
    .catch(err => {
        alert("Lỗi kết nối tới máy chủ: " + err);
        btn.disabled = false;
        btn.innerHTML = `<i class="fa-solid fa-cloud-arrow-down me-1"></i> Bắt Đầu Quét & Tải Hóa Đơn`;
    });
}

function startLogPolling() {
    if (syncInterval) clearInterval(syncInterval);

    syncInterval = setInterval(() => {
        fetch("/api/logs")
            .then(res => res.json())
            .then(data => {
                const logBox = document.getElementById("logBox");
                if (data.logs && data.logs.length > 0) {
                    logBox.innerHTML = data.logs.map(l => `<div>${escapeHtml(l)}</div>`).join("");
                    logBox.scrollTop = logBox.scrollHeight;
                }

                if (!data.is_running) {
                    clearInterval(syncInterval);
                    syncInterval = null;

                    const btn = document.getElementById("btnSync");
                    btn.disabled = false;
                    btn.innerHTML = `<i class="fa-solid fa-cloud-arrow-down me-1"></i> Bắt Đầu Quét & Tải Hóa Đơn`;

                    document.getElementById("syncStatusText").innerHTML = `
                        <i class="fa-solid fa-check text-success me-1"></i> Đồng bộ hoàn tất!
                    `;
                    document.getElementById("syncProgressBar").classList.remove("progress-bar-animated");

                    loadData();
                }
            })
            .catch(err => console.log("Lỗi polling log:", err));
    }, 1500);
}

function toggleLogBox() {
    const box = document.getElementById("logBox");
    box.classList.toggle("d-none");
}

function viewPdf(filename, soHd) {
    document.getElementById("pdfModalTitle").textContent = `Hóa Đơn Số: ${soHd} - ${filename}`;
    const frame = document.getElementById("pdfFrame");
    frame.src = `/api/pdf/${encodeURIComponent(filename)}`;
    const modal = new bootstrap.Modal(document.getElementById("pdfModal"));
    modal.show();
}

function openExcel() {
    fetch("/api/open-excel", { method: "POST" })
        .then(res => res.json())
        .then(data => {
            if (data.status !== "success") {
                alert(data.message || "Không thể mở file Excel!");
            }
        })
        .catch(err => alert("Lỗi khi mở Excel: " + err));
}

function openPdfFolder() {
    fetch("/api/open-folder", { method: "POST" })
        .then(res => res.json())
        .then(data => {
            if (data.status !== "success") {
                alert(data.message || "Không thể mở thư mục PDF!");
            }
        })
        .catch(err => alert("Lỗi khi mở thư mục: " + err));
}

function escapeHtml(text) {
    const div = document.createElement("div");
    div.textContent = text;
    return div.innerHTML;
}
