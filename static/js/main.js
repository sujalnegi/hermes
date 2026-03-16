/* ═══════════════════════════════════
   Hermes — Main JavaScript
   ═══════════════════════════════════ */

document.addEventListener("DOMContentLoaded", () => {

    // ── Navbar scroll ──
    const navbar = document.getElementById("navbar");
    if (navbar) {
        window.addEventListener("scroll", () => {
            navbar.classList.toggle("scrolled", window.scrollY > 40);
        });
    }

    // ── Drop zone ──
    const dropZone   = document.getElementById("drop-zone");
    const fileInput  = document.getElementById("eml-input");
    const fileList   = document.getElementById("file-list");
    const fileCount  = document.getElementById("file-count");
    const fileNames  = document.getElementById("file-names");
    const submitBtn  = document.getElementById("submit-btn");
    const inner      = document.getElementById("drop-zone-inner");

    if (dropZone && fileInput) {
        dropZone.addEventListener("click", () => fileInput.click());

        dropZone.addEventListener("dragover", (e) => {
            e.preventDefault();
            dropZone.classList.add("drag-over");
        });
        dropZone.addEventListener("dragleave", () => dropZone.classList.remove("drag-over"));

        dropZone.addEventListener("drop", (e) => {
            e.preventDefault();
            dropZone.classList.remove("drag-over");
            if (e.dataTransfer.files.length) {
                fileInput.files = e.dataTransfer.files;
                showFiles(e.dataTransfer.files);
            }
        });

        fileInput.addEventListener("change", () => showFiles(fileInput.files));
    }

    function showFiles(files) {
        const emls = [...files].filter(f => f.name.toLowerCase().endsWith(".eml"));
        if (!emls.length) {
            fileList.style.display = "none";
            if (inner) inner.style.display = "flex";
            dropZone.classList.remove("has-files");
            submitBtn.disabled = true;
            return;
        }
        if (inner) inner.style.display = "none";
        fileList.style.display = "block";
        dropZone.classList.add("has-files");
        submitBtn.disabled = false;
        fileCount.textContent = `${emls.length} file${emls.length > 1 ? "s" : ""} selected`;
        fileNames.innerHTML = "";
        emls.forEach(f => {
            const li = document.createElement("li");
            li.textContent = f.name;
            fileNames.appendChild(li);
        });
    }

    // ── Form loading state ──
    const uploadForm = document.getElementById("upload-form");
    if (uploadForm) {
        uploadForm.addEventListener("submit", () => {
            if (submitBtn && !submitBtn.disabled) {
                submitBtn.classList.add("loading");
                submitBtn.disabled = true;
            }
        });
    }

    const dataBtn = document.getElementById("data-btn");
    if (dataBtn) {
        dataBtn.closest("form").addEventListener("submit", () => {
            dataBtn.classList.add("loading");
            dataBtn.disabled = true;
        });
    }

    // ═══════════════════════════════════
    //  Results page — Checkboxes & Filters
    // ═══════════════════════════════════

    const rowCheckboxes   = document.querySelectorAll(".row-checkbox");
    const headerCheck     = document.getElementById("header-check");
    const selectAllBtn    = document.getElementById("select-all-btn");
    const deselectAllBtn  = document.getElementById("deselect-all-btn");
    const selectionCount  = document.getElementById("selection-count");
    const exportForm      = document.getElementById("export-form");
    const exportBtn       = document.getElementById("export-btn");

    function updateSelectionCount() {
        const checked = document.querySelectorAll(".row-checkbox:checked").length;
        if (selectionCount) {
            selectionCount.textContent = `${checked} selected for AI response`;
        }
        // Update header checkbox state
        if (headerCheck) {
            const visibleChecks = getVisibleCheckboxes();
            const visibleChecked = visibleChecks.filter(cb => cb.checked).length;
            headerCheck.checked = visibleChecks.length > 0 && visibleChecked === visibleChecks.length;
            headerCheck.indeterminate = visibleChecked > 0 && visibleChecked < visibleChecks.length;
        }
    }

    function getVisibleCheckboxes() {
        return [...rowCheckboxes].filter(cb => {
            const row = cb.closest("tr");
            return row && !row.classList.contains("hidden");
        });
    }

    // Individual checkbox change
    rowCheckboxes.forEach(cb => {
        cb.addEventListener("change", () => {
            const row = cb.closest("tr");
            row.classList.toggle("checked", cb.checked);
            updateSelectionCount();
        });
    });

    // Header checkbox — toggle visible rows
    if (headerCheck) {
        headerCheck.addEventListener("change", () => {
            const visible = getVisibleCheckboxes();
            visible.forEach(cb => {
                cb.checked = headerCheck.checked;
                cb.closest("tr").classList.toggle("checked", cb.checked);
            });
            updateSelectionCount();
        });
    }

    // Select All button
    if (selectAllBtn) {
        selectAllBtn.addEventListener("click", () => {
            const visible = getVisibleCheckboxes();
            visible.forEach(cb => {
                cb.checked = true;
                cb.closest("tr").classList.add("checked");
            });
            updateSelectionCount();
        });
    }

    // Deselect All button
    if (deselectAllBtn) {
        deselectAllBtn.addEventListener("click", () => {
            const visible = getVisibleCheckboxes();
            visible.forEach(cb => {
                cb.checked = false;
                cb.closest("tr").classList.remove("checked");
            });
            updateSelectionCount();
        });
    }

    // Guidelines PDF Upload
    const addGuidelinesBtn = document.getElementById("add-guidelines-btn");
    const guidelinesPdfInput = document.getElementById("guidelines-pdf");
    const guidelinesLabel = document.getElementById("guidelines-label");

    if (addGuidelinesBtn && guidelinesPdfInput) {
        addGuidelinesBtn.addEventListener("click", () => {
            guidelinesPdfInput.click();
        });

        guidelinesPdfInput.addEventListener("change", (e) => {
            if (e.target.files && e.target.files.length > 0) {
                const file = e.target.files[0];
                guidelinesLabel.textContent = file.name;
            } else {
                guidelinesLabel.textContent = "Add Guidelines";
            }
        });
    }

    // Export form — loading state
    if (exportForm) {
        exportForm.addEventListener("submit", () => {
            if (exportBtn) {
                exportBtn.classList.add("loading");
                exportBtn.disabled = true;
                // Re-enable after a timeout (in case download starts and page doesn't navigate)
                setTimeout(() => {
                    exportBtn.classList.remove("loading");
                    exportBtn.disabled = false;
                }, 30000);
            }
        });
    }

    // ── Filter tabs ──
    const filterBar = document.getElementById("filter-bar");
    if (filterBar) {
        const btns = filterBar.querySelectorAll(".filter-btn");
        const rows = document.querySelectorAll(".result-row");
        btns.forEach(btn => {
            btn.addEventListener("click", () => {
                btns.forEach(b => b.classList.remove("active"));
                btn.classList.add("active");
                const f = btn.dataset.filter;
                rows.forEach(r => {
                    r.classList.toggle("hidden", f !== "all" && r.dataset.senderType !== f);
                });
                // Update header checkbox after filter change
                updateSelectionCount();
            });
        });
    }

    // ── Smooth anchor scroll ──
    document.querySelectorAll('a[href^="#"]').forEach(a => {
        a.addEventListener("click", e => {
            e.preventDefault();
            const t = document.querySelector(a.getAttribute("href"));
            if (t) t.scrollIntoView({ behavior: "smooth", block: "start" });
        });
    });

    // ── Intersection Observer for fade-in ──
    const obs = new IntersectionObserver(entries => {
        entries.forEach(e => {
            if (e.isIntersecting) {
                e.target.style.opacity = "1";
                e.target.style.transform = "translateY(0)";
                obs.unobserve(e.target);
            }
        });
    }, { threshold: 0.1, rootMargin: "0px 0px -40px 0px" });

    document.querySelectorAll(".step-card, .card, .stat-card").forEach(el => {
        el.style.opacity = "0";
        el.style.transform = "translateY(18px)";
        el.style.transition = "opacity .55s ease-out, transform .55s ease-out";
        obs.observe(el);
    });

    // Initialize selection count
    updateSelectionCount();
});
