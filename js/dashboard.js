// dashboard.js - Core Dashboard Application Logic

class DashboardApp {
  constructor() {
    this.data = null;
    this.projects = [];
    this.filteredProjects = [];
    this.selectedProjectForModal = null;
    this.comparisonIds = new Set();
    
    // Filters & Sorting state
    this.searchQuery = "";
    this.selectedCategory = "ALL";
    this.selectedAgency = "ALL";
    this.budgetFilter = "ALL";
    this.sortField = "Y27";
    this.sortOrder = "desc";
    this.currentViewMode = "table"; // 'table' or 'card'
    this.activeTab = "overview"; // 'overview', 'categories', 'projects', 'comparison'
    this.groupBy = "Category"; // 'Category' or 'Agency'

    // Chart instances
    this.charts = {
      categoryDonut: null,
      topProjectsBar: null,
      categoryCombo: null,
      agencyBar: null
    };

    this.init();
  }

  init() {
    // 1. Theme initialization
    this.initTheme();

    // 2. Load default data
    if (window.DEFAULT_DATA) {
      this.loadData(window.DEFAULT_DATA);
    }

    // 3. Bind UI event listeners
    this.bindEvents();

    // 4. Initial Render
    this.render();
  }

  initTheme() {
    const savedTheme = localStorage.getItem("dashboard-theme") || "light";
    document.documentElement.setAttribute("data-theme", savedTheme);
    const themeBtn = document.getElementById("themeToggleBtn");
    if (themeBtn) {
      themeBtn.innerHTML = savedTheme === "dark" ? "☀️ 라이트모드" : "🌙 다크모드";
    }
  }

  toggleTheme() {
    const current = document.documentElement.getAttribute("data-theme") || "light";
    const next = current === "dark" ? "light" : "dark";
    document.documentElement.setAttribute("data-theme", next);
    localStorage.setItem("dashboard-theme", next);
    const themeBtn = document.getElementById("themeToggleBtn");
    if (themeBtn) {
      themeBtn.innerHTML = next === "dark" ? "☀️ 라이트모드" : "🌙 다크모드";
    }
    // Update chart colors
    this.renderCharts();
  }

  loadData(dataset) {
    this.data = dataset;
    this.projects = JSON.parse(JSON.stringify(dataset.projects || []));
    
    // Ensure all projects have consistent properties
    this.projects.forEach((p, idx) => {
      if (!p.id) p.id = p.No || (idx + 1);
      if (!p.Agency) p.Agency = ExcelParser.inferAgency(p.Dept);
      if (!p.Category) p.Category = ExcelParser.inferCategory(p.Name, p.Dept, p.Agency);
      p.Y27 = Number(p.Y27) || 0;
      p.Y26 = Number(p.Y26) || 0;
      p.Diff = Number(p.Diff) !== undefined ? Number(p.Diff) : (p.Y27 - p.Y26);
    });

    // Populate filter dropdowns
    this.populateFilterOptions();

    // Select default 2 projects for comparison
    this.comparisonIds.clear();
    if (this.projects.length >= 2) {
      this.comparisonIds.add(this.projects[0].id);
      this.comparisonIds.add(this.projects[1].id);
    }

    // Update dataset banner
    this.updateDataBanner();

    // Apply filters & render
    this.applyFilters();
  }

  updateDataBanner() {
    const bannerTitle = document.getElementById("datasetTitle");
    const bannerSubtitle = document.getElementById("datasetSubtitle");
    if (bannerTitle && this.data.summary) {
      bannerTitle.textContent = this.data.summary.title || "국방예산 신규사업 현황";
    }
    if (bannerSubtitle && this.data.summary) {
      bannerSubtitle.textContent = this.data.summary.subtitle || `총 ${this.projects.length}개 사업 분석`;
    }
  }

  populateFilterOptions() {
    const catSelect = document.getElementById("categoryFilter");
    const agencySelect = document.getElementById("agencyFilter");

    if (catSelect) {
      const categories = [...new Set(this.projects.map(p => p.Category).filter(Boolean))].sort();
      catSelect.innerHTML = `<option value="ALL">전체 분야 (${categories.length}개 분야)</option>` +
        categories.map(c => `<option value="${c}">${c}</option>`).join("");
      catSelect.value = this.selectedCategory;
    }

    if (agencySelect) {
      const agencies = [...new Set(this.projects.map(p => p.Agency).filter(Boolean))].sort();
      agencySelect.innerHTML = `<option value="ALL">전체 주관기관 (${agencies.length}개 기관)</option>` +
        agencies.map(a => `<option value="${a}">${a}</option>`).join("");
      agencySelect.value = this.selectedAgency;
    }
  }

  bindEvents() {
    // Theme Toggle
    const themeBtn = document.getElementById("themeToggleBtn");
    if (themeBtn) {
      themeBtn.addEventListener("click", () => this.toggleTheme());
    }

    // Tab Navigation
    document.querySelectorAll(".tab-btn").forEach(btn => {
      btn.addEventListener("click", (e) => {
        const tab = e.currentTarget.getAttribute("data-tab");
        this.switchTab(tab);
      });
    });

    // File Upload (Input & Drag and Drop)
    const fileInput = document.getElementById("fileInput");
    const dropZone = document.getElementById("dropZone");
    const openUploadBtn = document.getElementById("openUploadBtn");

    if (openUploadBtn && fileInput) {
      openUploadBtn.addEventListener("click", () => fileInput.click());
    }

    if (fileInput) {
      fileInput.addEventListener("change", (e) => {
        if (e.target.files && e.target.files[0]) {
          this.handleFile(e.target.files[0]);
        }
      });
    }

    if (dropZone) {
      ["dragenter", "dragover"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.add("dragover");
        });
      });

      ["dragleave", "drop"].forEach(eventName => {
        dropZone.addEventListener(eventName, (e) => {
          e.preventDefault();
          e.stopPropagation();
          dropZone.classList.remove("dragover");
        });
      });

      dropZone.addEventListener("drop", (e) => {
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
          this.handleFile(e.dataTransfer.files[0]);
        }
      });
    }

    // Reset to default sample
    const resetSampleBtn = document.getElementById("resetSampleBtn");
    if (resetSampleBtn) {
      resetSampleBtn.addEventListener("click", () => {
        if (window.DEFAULT_DATA) {
          this.loadData(window.DEFAULT_DATA);
          this.showToast("기본 샘플 데이터(2027 국방예산 신규사업 12건)로 복원되었습니다.");
        }
      });
    }

    // Search and Filters
    const searchInput = document.getElementById("searchInput");
    if (searchInput) {
      searchInput.addEventListener("input", (e) => {
        this.searchQuery = e.target.value.trim().toLowerCase();
        this.applyFilters();
      });
    }

    const catSelect = document.getElementById("categoryFilter");
    if (catSelect) {
      catSelect.addEventListener("change", (e) => {
        this.selectedCategory = e.target.value;
        this.applyFilters();
      });
    }

    const agencySelect = document.getElementById("agencyFilter");
    if (agencySelect) {
      agencySelect.addEventListener("change", (e) => {
        this.selectedAgency = e.target.value;
        this.applyFilters();
      });
    }

    const budgetSelect = document.getElementById("budgetFilter");
    if (budgetSelect) {
      budgetSelect.addEventListener("change", (e) => {
        this.budgetFilter = e.target.value;
        this.applyFilters();
      });
    }

    const sortSelect = document.getElementById("sortSelect");
    if (sortSelect) {
      sortSelect.addEventListener("change", (e) => {
        const [field, order] = e.target.value.split("-");
        this.sortField = field;
        this.sortOrder = order;
        this.applyFilters();
      });
    }

    // View Mode Toggle (Table / Card)
    const viewTableBtn = document.getElementById("viewTableBtn");
    const viewCardBtn = document.getElementById("viewCardBtn");
    if (viewTableBtn && viewCardBtn) {
      viewTableBtn.addEventListener("click", () => {
        this.currentViewMode = "table";
        viewTableBtn.classList.add("active");
        viewCardBtn.classList.remove("active");
        this.renderProjectsView();
      });
      viewCardBtn.addEventListener("click", () => {
        this.currentViewMode = "card";
        viewCardBtn.classList.add("active");
        viewTableBtn.classList.remove("active");
        this.renderProjectsView();
      });
    }

    // Group By Toggle in Category Tab
    const groupToggleBtns = document.querySelectorAll(".group-toggle-btn");
    groupToggleBtns.forEach(btn => {
      btn.addEventListener("click", (e) => {
        groupToggleBtns.forEach(b => b.classList.remove("active"));
        e.currentTarget.classList.add("active");
        this.groupBy = e.currentTarget.getAttribute("data-group");
        this.renderCategoryAnalysis();
      });
    });

    // Modal Close
    const closeModalBtn = document.getElementById("closeModalBtn");
    const modalBackdrop = document.getElementById("detailModal");
    if (closeModalBtn) {
      closeModalBtn.addEventListener("click", () => this.closeModal());
    }
    if (modalBackdrop) {
      modalBackdrop.addEventListener("click", (e) => {
        if (e.target === modalBackdrop) this.closeModal();
      });
    }
    window.addEventListener("keydown", (e) => {
      if (e.key === "Escape") this.closeModal();
    });

    // Export Buttons
    const exportExcelBtn = document.getElementById("exportExcelBtn");
    if (exportExcelBtn) {
      exportExcelBtn.addEventListener("click", () => this.exportToExcel());
    }

    const exportCsvBtn = document.getElementById("exportCsvBtn");
    if (exportCsvBtn) {
      exportCsvBtn.addEventListener("click", () => this.exportToCsv());
    }

    const printBtn = document.getElementById("printBtn");
    if (printBtn) {
      printBtn.addEventListener("click", () => window.print());
    }

    // Sample Download button
    const downloadSampleBtn = document.getElementById("downloadSampleBtn");
    if (downloadSampleBtn) {
      downloadSampleBtn.addEventListener("click", () => this.downloadSampleExcel());
    }
  }

  async handleFile(file) {
    try {
      this.showToast(`'${file.name}' 파일을 분석하는 중입니다...`, 2000);
      const parsedData = await ExcelParser.parseFile(file);
      this.loadData(parsedData);
      this.showToast(`✅ ${parsedData.projects.length}개 사업 데이터를 성공적으로 불러왔습니다!`, 3000);
    } catch (err) {
      alert("파일 파싱 실패: " + err.message);
      console.error(err);
    }
  }

  switchTab(tabId) {
    this.activeTab = tabId;
    document.querySelectorAll(".tab-btn").forEach(btn => {
      btn.classList.toggle("active", btn.getAttribute("data-tab") === tabId);
    });
    document.querySelectorAll(".tab-content").forEach(content => {
      content.classList.toggle("active", content.id === `tab-${tabId}`);
    });

    // Trigger tab specific renders
    if (tabId === "overview") {
      this.renderCharts();
    } else if (tabId === "categories") {
      this.renderCategoryAnalysis();
    } else if (tabId === "projects") {
      this.renderProjectsView();
    } else if (tabId === "comparison") {
      this.renderComparisonView();
    }
  }

  applyFilters() {
    let list = [...this.projects];

    // 1. Text Search
    if (this.searchQuery) {
      list = list.filter(p => {
        const combined = `${p.Name} ${p.Dept} ${p.Agency} ${p.Category} ${p.Contact} ${p.Note} ${p.OrigNo}`.toLowerCase();
        return combined.includes(this.searchQuery);
      });
    }

    // 2. Category Filter
    if (this.selectedCategory !== "ALL") {
      list = list.filter(p => p.Category === this.selectedCategory);
    }

    // 3. Agency Filter
    if (this.selectedAgency !== "ALL") {
      list = list.filter(p => p.Agency === this.selectedAgency);
    }

    // 4. Budget Range Filter
    if (this.budgetFilter !== "ALL") {
      if (this.budgetFilter === "OVER_500") {
        list = list.filter(p => p.Y27 >= 500);
      } else if (this.budgetFilter === "100_500") {
        list = list.filter(p => p.Y27 >= 100 && p.Y27 < 500);
      } else if (this.budgetFilter === "10_100") {
        list = list.filter(p => p.Y27 >= 10 && p.Y27 < 100);
      } else if (this.budgetFilter === "UNDER_10") {
        list = list.filter(p => p.Y27 > 0 && p.Y27 < 10);
      } else if (this.budgetFilter === "ZERO") {
        list = list.filter(p => p.Y27 === 0);
      }
    }

    // 5. Sorting
    list.sort((a, b) => {
      let valA = a[this.sortField];
      let valB = b[this.sortField];

      if (typeof valA === "string") valA = valA.toLowerCase();
      if (typeof valB === "string") valB = valB.toLowerCase();

      if (valA < valB) return this.sortOrder === "asc" ? -1 : 1;
      if (valA > valB) return this.sortOrder === "asc" ? 1 : -1;
      return 0;
    });

    this.filteredProjects = list;

    // Update UI counters & badges
    this.updateKPIs();
    this.render();
  }

  updateKPIs() {
    const list = this.filteredProjects;
    const countEl = document.getElementById("kpiTotalCount");
    const budgetEl = document.getElementById("kpiTotalBudget");
    const maxProjectEl = document.getElementById("kpiMaxProject");
    const maxBudgetEl = document.getElementById("kpiMaxBudget");
    const agencyCountEl = document.getElementById("kpiAgencyCount");
    const avgBudgetEl = document.getElementById("kpiAvgBudget");
    const filteredCountBadge = document.getElementById("filteredCountBadge");

    const totalCount = list.length;
    const totalBudget = list.reduce((sum, p) => sum + (p.Diff > 0 ? p.Diff : p.Y27), 0);
    const validBudgetProjects = list.filter(p => p.Y27 > 0);
    const avgBudget = validBudgetProjects.length > 0 ? totalBudget / validBudgetProjects.length : 0;

    // Max project
    let maxProject = null;
    if (list.length > 0) {
      maxProject = list.reduce((prev, current) => ((current.Y27 || 0) > (prev.Y27 || 0)) ? current : prev, list[0]);
    }

    // Agencies & depts
    const uniqueAgencies = new Set(list.map(p => p.Agency).filter(Boolean));
    const uniqueDepts = new Set(list.map(p => p.Dept).filter(Boolean));

    if (countEl) countEl.textContent = `${totalCount.toLocaleString()}건`;
    if (budgetEl) budgetEl.textContent = `${totalBudget.toLocaleString(undefined, { minimumFractionDigits: 1, maximumFractionDigits: 2 })}억원`;
    if (avgBudgetEl) avgBudgetEl.textContent = `${avgBudget.toFixed(1)}억원`;
    if (agencyCountEl) agencyCountEl.textContent = `${uniqueAgencies.size}개 기관 (${uniqueDepts.size}개 부서)`;
    
    if (maxProjectEl && maxProject && maxProject.Y27 > 0) {
      maxProjectEl.textContent = maxProject.Name;
      if (maxBudgetEl) {
        const share = totalBudget > 0 ? ((maxProject.Y27 / totalBudget) * 100).toFixed(1) : 0;
        maxBudgetEl.textContent = `${maxProject.Y27.toLocaleString()}억원 (비중 ${share}%)`;
      }
    } else {
      if (maxProjectEl) maxProjectEl.textContent = "-";
      if (maxBudgetEl) maxBudgetEl.textContent = "-";
    }

    if (filteredCountBadge) {
      filteredCountBadge.textContent = `현재 조회: ${totalCount}건 / 전체: ${this.projects.length}건`;
    }
  }

  render() {
    this.renderCharts();
    this.renderCategoryAnalysis();
    this.renderProjectsView();
    this.renderComparisonView();
  }

  // ==========================================
  // Overview Tab Charts
  // ==========================================
  renderCharts() {
    this.renderCategoryDonut();
    this.renderTopProjectsBar();
  }

  renderCategoryDonut() {
    const canvas = document.getElementById("categoryDonutChart");
    if (!canvas) return;

    // Destroy existing
    if (this.charts.categoryDonut) {
      this.charts.categoryDonut.destroy();
    }

    // Group by Category
    const groupTotals = {};
    this.filteredProjects.forEach(p => {
      const cat = p.Category || "기타";
      groupTotals[cat] = (groupTotals[cat] || 0) + (p.Y27 || 0);
    });

    const labels = Object.keys(groupTotals);
    const data = Object.values(groupTotals);
    const total = data.reduce((a, b) => a + b, 0);

    const colors = [
      "#2563eb", "#3b82f6", "#06b6d4", "#10b981", 
      "#8b5cf6", "#f59e0b", "#ec4899", "#64748b"
    ];

    const ctx = canvas.getContext("2d");
    this.charts.categoryDonut = new Chart(ctx, {
      type: "doughnut",
      data: {
        labels: labels,
        datasets: [{
          data: data,
          backgroundColor: colors.slice(0, labels.length),
          borderWidth: 2,
          borderColor: document.documentElement.getAttribute("data-theme") === "dark" ? "#1e293b" : "#ffffff",
          hoverOffset: 6
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: "right",
            labels: {
              boxWidth: 14,
              font: { size: 12, family: "Pretendard, sans-serif" },
              generateLabels: (chart) => {
                const datasets = chart.data.datasets;
                return chart.data.labels.map((label, i) => {
                  const val = datasets[0].data[i];
                  const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                  return {
                    text: `${label} (${val.toFixed(1)}억 / ${pct}%)`,
                    fillStyle: datasets[0].backgroundColor[i],
                    hidden: isNaN(datasets[0].data[i]),
                    index: i
                  };
                });
              }
            }
          },
          tooltip: {
            callbacks: {
              label: (context) => {
                const val = context.raw || 0;
                const pct = total > 0 ? ((val / total) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${val.toLocaleString()}억원 (${pct}%)`;
              }
            }
          }
        },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const index = elements[0].index;
            const clickedCategory = labels[index];
            this.selectedCategory = clickedCategory;
            const catSelect = document.getElementById("categoryFilter");
            if (catSelect) catSelect.value = clickedCategory;
            this.applyFilters();
            this.showToast(`'${clickedCategory}' 분야로 필터링되었습니다.`);
          }
        }
      }
    });
  }

  renderTopProjectsBar() {
    const canvas = document.getElementById("topProjectsBarChart");
    if (!canvas) return;

    if (this.charts.topProjectsBar) {
      this.charts.topProjectsBar.destroy();
    }

    // Top 8 projects by Y27
    const topProjects = [...this.filteredProjects]
      .sort((a, b) => b.Y27 - a.Y27)
      .slice(0, 8);

    const labels = topProjects.map(p => {
      const name = p.Name;
      return name.length > 18 ? name.substring(0, 18) + "..." : name;
    });
    const data = topProjects.map(p => p.Y27);

    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const ctx = canvas.getContext("2d");
    this.charts.topProjectsBar = new Chart(ctx, {
      type: "bar",
      data: {
        labels: labels,
        datasets: [{
          label: "'27년 정부안 신규 예산 (억원)",
          data: data,
          backgroundColor: "#1e40af",
          hoverBackgroundColor: "#2563eb",
          borderRadius: 6,
          barPercentage: 0.65
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              title: (items) => {
                const idx = items[0].dataIndex;
                return topProjects[idx].Name;
              },
              afterTitle: (items) => {
                const idx = items[0].dataIndex;
                return `주관: ${topProjects[idx].Dept}`;
              },
              label: (context) => ` '27년 예산: ${context.raw.toLocaleString()}억원`
            }
          }
        },
        scales: {
          x: {
            grid: { color: isDark ? "#334155" : "#e2e8f0" },
            ticks: {
              font: { family: "Pretendard, sans-serif" },
              callback: (val) => val + "억"
            }
          },
          y: {
            grid: { display: false },
            ticks: {
              font: { size: 11, family: "Pretendard, sans-serif" }
            }
          }
        },
        onClick: (evt, elements) => {
          if (elements && elements.length > 0) {
            const index = elements[0].index;
            const project = topProjects[index];
            this.openDetailModal(project);
          }
        }
      }
    });
  }

  // ==========================================
  // Category Analysis Tab
  // ==========================================
  renderCategoryAnalysis() {
    const groupField = this.groupBy; // 'Category' or 'Agency'
    const groups = {};

    this.filteredProjects.forEach(p => {
      const key = p[groupField] || "기타";
      if (!groups[key]) {
        groups[key] = {
          name: key,
          count: 0,
          totalBudget: 0,
          projects: []
        };
      }
      groups[key].count++;
      groups[key].totalBudget += (p.Y27 || 0);
      groups[key].projects.push(p);
    });

    const groupList = Object.values(groups).sort((a, b) => b.totalBudget - a.totalBudget);
    const grandTotal = groupList.reduce((sum, g) => sum + g.totalBudget, 0);

    // Render Combo Chart (Budget + Count)
    this.renderCategoryComboChart(groupList);

    // Render Category Cards Grid
    const gridContainer = document.getElementById("categoryCardsGrid");
    if (!gridContainer) return;

    gridContainer.innerHTML = groupList.map((g, idx) => {
      const share = grandTotal > 0 ? ((g.totalBudget / grandTotal) * 100).toFixed(1) : 0;
      const sortedProjs = [...g.projects].sort((a, b) => b.Y27 - a.Y27);

      return `
        <div class="category-card">
          <div class="category-card-header">
            <div>
              <span class="category-rank">#${idx + 1}</span>
              <h3 class="category-title">${g.name}</h3>
            </div>
            <div class="category-badge">${share}% 점유</div>
          </div>
          <div class="category-metrics">
            <div class="category-metric-item">
              <span class="label">총 예산</span>
              <span class="value highlight">${g.totalBudget.toLocaleString()}억원</span>
            </div>
            <div class="category-metric-item">
              <span class="label">신규 사업 수</span>
              <span class="value">${g.count}건</span>
            </div>
            <div class="category-metric-item">
              <span class="label">평균 사업비</span>
              <span class="value">${(g.totalBudget / g.count).toFixed(1)}억원</span>
            </div>
          </div>
          <div class="category-project-list">
            <div class="list-title">주요 신규사업 목록</div>
            ${sortedProjs.map(p => `
              <div class="category-project-item" onclick="app.openDetailModalById(${p.id})">
                <span class="p-name">${p.Name}</span>
                <span class="p-budget">${p.Y27 > 0 ? p.Y27.toLocaleString() + '억' : '제도·시범'}</span>
              </div>
            `).join("")}
          </div>
        </div>
      `;
    }).join("");
  }

  renderCategoryComboChart(groupList) {
    const canvas = document.getElementById("categoryComboChart");
    if (!canvas) return;

    if (this.charts.categoryCombo) {
      this.charts.categoryCombo.destroy();
    }

    const labels = groupList.map(g => g.name);
    const budgets = groupList.map(g => g.totalBudget);
    const counts = groupList.map(g => g.count);

    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const ctx = canvas.getContext("2d");

    this.charts.categoryCombo = new Chart(ctx, {
      data: {
        labels: labels,
        datasets: [
          {
            type: "bar",
            label: "신규 예산액 (억원)",
            data: budgets,
            backgroundColor: "rgba(37, 99, 235, 0.8)",
            borderRadius: 6,
            yAxisID: "yBudget",
            order: 2
          },
          {
            type: "line",
            label: "사업 건수 (건)",
            data: counts,
            borderColor: "#ef4444",
            backgroundColor: "#ef4444",
            borderWidth: 3,
            pointRadius: 5,
            pointHoverRadius: 7,
            yAxisID: "yCount",
            order: 1
          }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: {
            grid: { display: false },
            ticks: { font: { family: "Pretendard, sans-serif" } }
          },
          yBudget: {
            type: "linear",
            position: "left",
            title: { display: true, text: "예산 규모 (억원)" },
            grid: { color: isDark ? "#334155" : "#e2e8f0" }
          },
          yCount: {
            type: "linear",
            position: "right",
            title: { display: true, text: "사업 수 (건)" },
            grid: { display: false },
            ticks: { stepSize: 1, precision: 0 }
          }
        }
      }
    });
  }

  // ==========================================
  // Projects View Tab (Table / Card)
  // ==========================================
  renderProjectsView() {
    const tableContainer = document.getElementById("projectsTableView");
    const cardContainer = document.getElementById("projectsCardView");
    const emptyState = document.getElementById("projectsEmptyState");

    if (this.filteredProjects.length === 0) {
      if (tableContainer) tableContainer.style.display = "none";
      if (cardContainer) cardContainer.style.display = "none";
      if (emptyState) emptyState.style.display = "block";
      return;
    }

    if (emptyState) emptyState.style.display = "none";

    if (this.currentViewMode === "table") {
      if (tableContainer) tableContainer.style.display = "block";
      if (cardContainer) cardContainer.style.display = "none";
      this.renderTableContent();
    } else {
      if (tableContainer) tableContainer.style.display = "none";
      if (cardContainer) cardContainer.style.display = "grid";
      this.renderCardContent();
    }
  }

  renderTableContent() {
    const tbody = document.getElementById("projectsTableBody");
    if (!tbody) return;

    tbody.innerHTML = this.filteredProjects.map(p => {
      const isSelected = this.comparisonIds.has(p.id);
      return `
        <tr class="project-table-row ${isSelected ? 'row-selected' : ''}" onclick="app.openDetailModalById(${p.id})">
          <td class="col-center" onclick="event.stopPropagation()">
            <input type="checkbox" ${isSelected ? 'checked' : ''} onchange="app.toggleComparison(${p.id})">
          </td>
          <td class="col-center text-muted">${p.No}</td>
          <td class="col-center"><span class="badge-orig-no">${p.OrigNo || '-'}</span></td>
          <td>
            <div class="fw-semibold text-primary">${p.Name}</div>
            <div class="sub-text text-truncate">${p.Note ? p.Note.substring(0, 45) + '...' : ''}</div>
          </td>
          <td>
            <span class="badge-dept">${p.Dept}</span>
          </td>
          <td class="col-center">
            <span class="badge-category">${p.Category}</span>
          </td>
          <td class="col-right text-muted">${p.Y26 > 0 ? p.Y26.toLocaleString() : '-'}</td>
          <td class="col-right fw-bold text-accent">${p.Y27 > 0 ? p.Y27.toLocaleString() : '-'}</td>
          <td class="col-right fw-bold text-success">${p.Diff > 0 ? '+' + p.Diff.toLocaleString() : (p.Diff === 0 && p.Y27 === 0 ? '-' : p.Diff.toLocaleString())}</td>
          <td class="col-center text-muted">${p.Unit || '억원'}</td>
          <td class="col-center" onclick="event.stopPropagation()">
            <button class="btn btn-sm btn-outline" onclick="app.openDetailModalById(${p.id})">상세보기</button>
          </td>
        </tr>
      `;
    }).join("");
  }

  renderCardContent() {
    const cardGrid = document.getElementById("projectsCardView");
    if (!cardGrid) return;

    cardGrid.innerHTML = this.filteredProjects.map(p => {
      const isSelected = this.comparisonIds.has(p.id);
      return `
        <div class="project-card ${isSelected ? 'card-selected' : ''}" onclick="app.openDetailModalById(${p.id})">
          <div class="project-card-header">
            <div class="d-flex align-items-center gap-2">
              <span class="badge-orig-no">${p.OrigNo || ('No.' + p.No)}</span>
              <span class="badge-category">${p.Category}</span>
            </div>
            <div onclick="event.stopPropagation()">
              <label class="compare-checkbox-label">
                <input type="checkbox" ${isSelected ? 'checked' : ''} onchange="app.toggleComparison(${p.id})"> 비교
              </label>
            </div>
          </div>
          <h4 class="project-card-title">${p.Name}</h4>
          <div class="project-card-dept">
            <span class="icon">🏢</span> ${p.Dept}
          </div>
          <div class="project-card-note">${p.Note || '상세 추진 내용 참조'}</div>
          <div class="project-card-footer">
            <div class="budget-block">
              <span class="budget-label">'27 신규예산</span>
              <span class="budget-value">${p.Y27 > 0 ? p.Y27.toLocaleString() + ' ' + (p.Unit || '억원') : '제도/시범'}</span>
            </div>
            <button class="btn btn-sm btn-primary" onclick="event.stopPropagation(); app.openDetailModalById(${p.id})">상세보기</button>
          </div>
        </div>
      `;
    }).join("");
  }

  // ==========================================
  // Comparison Tab
  // ==========================================
  toggleComparison(id) {
    if (this.comparisonIds.has(id)) {
      this.comparisonIds.delete(id);
    } else {
      if (this.comparisonIds.size >= 4) {
        this.showToast("사업 비교는 최대 4개까지 가능합니다.");
        return;
      }
      this.comparisonIds.add(id);
    }
    this.renderProjectsView();
    this.renderComparisonView();
  }

  renderComparisonView() {
    const container = document.getElementById("comparisonContent");
    const countBadge = document.getElementById("comparisonCountBadge");
    if (countBadge) {
      countBadge.textContent = `${this.comparisonIds.size}개 사업 선택됨`;
    }

    if (!container) return;

    const selectedProjects = this.projects.filter(p => this.comparisonIds.has(p.id));

    if (selectedProjects.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <div class="empty-icon">⚖️</div>
          <h3>비교할 사업을 선택해주세요</h3>
          <p>사업 목록에서 [비교] 체크박스를 클릭하여 2~4개 사업을 선택하시면 나란히 비교 분석할 수 있습니다.</p>
          <button class="btn btn-primary" onclick="app.switchTab('projects')">사업 목록으로 이동</button>
        </div>
      `;
      return;
    }

    const fields = [
      { label: "원사업번호", key: "OrigNo" },
      { label: "주관 기관", key: "Agency" },
      { label: "주관 부서", key: "Dept" },
      { label: "분야", key: "Category" },
      { label: "담당자", key: "Contact" },
      { label: "'26년 예산", key: "Y26", format: v => v > 0 ? v.toLocaleString() + " 억원" : "-" },
      { label: "'27년 정부안", key: "Y27", highlight: true, format: v => v > 0 ? v.toLocaleString() + " 억원" : "제도/시범" },
      { label: "신규 증감액", key: "Diff", format: v => v > 0 ? "+" + v.toLocaleString() + " 억원" : "-" },
      { label: "핵심 추진목표", key: "Note" },
      { label: "지원 대상", key: "Target" },
      { label: "추진 근거", key: "Basis" },
      { label: "세부 내용", key: "Detail" },
      { label: "기대 효과", key: "Effect" }
    ];

    container.innerHTML = `
      <div class="comparison-table-wrapper">
        <table class="comparison-table">
          <thead>
            <tr>
              <th style="width: 140px;">비교 항목</th>
              ${selectedProjects.map(p => `
                <th>
                  <div class="d-flex justify-content-between align-items-center">
                    <span class="compare-col-title">${p.Name}</span>
                    <button class="btn-remove-compare" onclick="app.toggleComparison(${p.id})">✕</button>
                  </div>
                </th>
              `).join("")}
            </tr>
          </thead>
          <tbody>
            ${fields.map(f => `
              <tr>
                <td class="compare-row-header">${f.label}</td>
                ${selectedProjects.map(p => {
                  let val = p[f.key] || "-";
                  if (f.format && p[f.key] !== undefined) {
                    val = f.format(p[f.key]);
                  }
                  return `<td class="${f.highlight ? 'compare-highlight' : ''}">${val}</td>`;
                }).join("")}
              </tr>
            `).join("")}
          </tbody>
        </table>
      </div>
    `;
  }

  // ==========================================
  // Modal Handling
  // ==========================================
  openDetailModalById(id) {
    const project = this.projects.find(p => p.id === id);
    if (project) {
      this.openDetailModal(project);
    }
  }

  openDetailModal(p) {
    this.selectedProjectForModal = p;

    document.getElementById("modalOrigNo").textContent = p.OrigNo || `연번 ${p.No}`;
    document.getElementById("modalCategory").textContent = p.Category || "신규사업";
    document.getElementById("modalTitle").textContent = p.Name;
    document.getElementById("modalDept").textContent = p.Dept;
    document.getElementById("modalContact").textContent = p.Contact || "-";

    document.getElementById("modalY26").textContent = p.Y26 > 0 ? `${p.Y26.toLocaleString()} ${p.Unit || '억원'}` : "-";
    document.getElementById("modalY27").textContent = p.Y27 > 0 ? `${p.Y27.toLocaleString()} ${p.Unit || '억원'}` : "제도/시범사업";
    document.getElementById("modalDiff").textContent = p.Diff > 0 ? `+${p.Diff.toLocaleString()} ${p.Unit || '억원'}` : (p.Diff === 0 && p.Y27 === 0 ? "-" : `${p.Diff.toLocaleString()} ${p.Unit || '억원'}`);
    document.getElementById("modalUnit").textContent = p.Unit || "억원";

    document.getElementById("modalNote").textContent = p.Note || "등록된 핵심 사업 내용이 없습니다.";
    document.getElementById("modalTarget").textContent = p.Target || "정보 없음";
    document.getElementById("modalBasis").textContent = p.Basis || "정보 없음";
    document.getElementById("modalDetail").textContent = p.Detail || "정보 없음";
    document.getElementById("modalEffect").textContent = p.Effect || "정보 없음";

    const modal = document.getElementById("detailModal");
    if (modal) {
      modal.classList.add("active");
      document.body.style.overflow = "hidden";
    }
  }

  closeModal() {
    const modal = document.getElementById("detailModal");
    if (modal) {
      modal.classList.remove("active");
      document.body.style.overflow = "";
    }
  }

  // ==========================================
  // Export & Download
  // ==========================================
  exportToExcel() {
    const rows = this.filteredProjects.map(p => ({
      "연번": p.No,
      "원사업번호": p.OrigNo,
      "신규 사업명": p.Name,
      "주관 기관": p.Agency,
      "담당 부서": p.Dept,
      "분야": p.Category,
      "담당자 연락처": p.Contact,
      "'26년 예산": p.Y26,
      "'27년 정부안(신규)": p.Y27,
      "신규 증감액": p.Diff,
      "단위": p.Unit,
      "핵심 신규사업 내용 및 추진목표": p.Note,
      "지원 대상 및 수혜 규모": p.Target,
      "추진 근거 및 법령": p.Basis,
      "세부 추진 내용 및 도입 일정": p.Detail,
      "기대 효과 및 정책 의의": p.Effect
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "신규사업 현황");

    const fileName = `신규사업_대시보드_추출자료_${new Date().toISOString().slice(0, 10)}.xlsx`;
    XLSX.writeFile(workbook, fileName);
    this.showToast(`✅ '${fileName}' 파일로 내보내기가 완료되었습니다.`);
  }

  exportToCsv() {
    const rows = this.filteredProjects.map(p => ({
      "연번": p.No,
      "원사업번호": p.OrigNo,
      "신규 사업명": p.Name,
      "주관 기관": p.Agency,
      "담당 부서": p.Dept,
      "분야": p.Category,
      "'27년 정부안(신규)": p.Y27,
      "신규 증감액": p.Diff,
      "단위": p.Unit,
      "핵심 신규사업 내용": p.Note
    }));

    const worksheet = XLSX.utils.json_to_sheet(rows);
    const csvContent = XLSX.utils.sheet_to_csv(worksheet);

    const blob = new Blob(["\uFEFF" + csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    const url = URL.createObjectURL(blob);
    link.setAttribute("href", url);
    link.setAttribute("download", `신규사업_현황_${new Date().toISOString().slice(0, 10)}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    this.showToast("✅ CSV 파일 다운로드가 완료되었습니다.");
  }

  downloadSampleExcel() {
    this.exportToExcel();
  }

  showToast(message, duration = 3000) {
    let toast = document.getElementById("appToast");
    if (!toast) {
      toast = document.createElement("div");
      toast.id = "appToast";
      toast.className = "toast";
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.classList.add("show");
    setTimeout(() => {
      toast.classList.remove("show");
    }, duration);
  }
}

// Global initialization
window.addEventListener("DOMContentLoaded", () => {
  window.app = new DashboardApp();
});
