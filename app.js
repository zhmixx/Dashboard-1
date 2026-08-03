/**
 * 高质量数据专项推进驾驶舱 - 主逻辑 JavaScript
 * (组学进展：样本计划与样本完成率紧密贴合排版、预计数据量与数据生产进度紧密贴合排版，名称更新为数据生产进度)
 */

let globalRawData = [];
let currentProjectFilter = 'ALL';
let currentTab = 'omics'; // 默认首屏视图：组学进展
let currentRankMetric = 'data'; // 默认：按数据生产进度

// 排序状态控制
let currentSort = {
  key: null,  // 当前排序字段
  dir: 'desc' // 'asc' 或 'desc'
};

// 图表实例
let sparklineChart = null;
let dataSparklineChart = null;
let rankChart = null;

// DOM 加载完成后初始化
document.addEventListener('DOMContentLoaded', () => {
  initCharts();
  initEvents();
  loadDefaultExcelData();
});

// 初始化 ECharts 实例
function initCharts() {
  const sparklineDom = document.getElementById('kpiSampleSparkline');
  if (sparklineDom) sparklineChart = echarts.init(sparklineDom);

  const dataSparklineDom = document.getElementById('kpiDataSparkline');
  if (dataSparklineDom) dataSparklineChart = echarts.init(dataSparklineDom);

  const rankDom = document.getElementById('rankChart');
  if (rankDom) rankChart = echarts.init(rankDom);

  const handleResize = () => {
    sparklineChart && sparklineChart.resize();
    dataSparklineChart && dataSparklineChart.resize();
    rankChart && rankChart.resize();
  };

  window.addEventListener('resize', handleResize);
  window.addEventListener('orientationchange', handleResize);
}

// 事件监听绑定
function initEvents() {
  // 专项筛选
  const projectSelect = document.getElementById('projectSelect');
  projectSelect.addEventListener('change', (e) => {
    currentProjectFilter = e.target.value;
    updateDashboard();
  });

  // 排名维度直接点击按钮切换
  const rankToggleBtns = document.querySelectorAll('.rank-toggle-btn');
  rankToggleBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      rankToggleBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      currentRankMetric = btn.dataset.metric;
      const filteredData = getFilteredData();
      renderRankChart(filteredData);
    });
  });

  // 刷新按钮
  document.getElementById('reloadBtn').addEventListener('click', () => {
    loadDefaultExcelData();
  });

  // 导入文件
  const uploadBtn = document.getElementById('uploadTriggerBtn');
  const excelFileInput = document.getElementById('excelFile');

  uploadBtn.addEventListener('click', () => excelFileInput.click());
  excelFileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) {
      readExcelFile(file);
    }
  });

  // Tab 切换 (重置排序)
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach((btn) => {
    btn.addEventListener('click', () => {
      tabBtns.forEach((b) => b.classList.remove('active'));
      btn.classList.add('active');
      currentTab = btn.dataset.tab;
      currentSort = { key: null, dir: 'desc' }; // 切换 Tab 时重置排序状态
      updateDrilldownTable();
    });
  });

  // 搜索框
  const searchInput = document.getElementById('tableSearchInput');
  searchInput.addEventListener('input', () => {
    updateDrilldownTable();
  });
}

// 表头排序点击切换函数
function handleSort(key) {
  if (currentSort.key === key) {
    currentSort.dir = currentSort.dir === 'desc' ? 'asc' : 'desc';
  } else {
    currentSort.key = key;
    currentSort.dir = 'desc';
  }
  updateDrilldownTable();
}

// 自动载入 data 目录下的专项进度表【周度更新】.xlsx
function loadDefaultExcelData() {
  const defaultPath = './data/专项进度表【周度更新】.xlsx';
  fetch(defaultPath)
    .then((res) => {
      if (!res.ok) throw new Error('无法自动获取固定路径 Excel 文件');
      return res.arrayBuffer();
    })
    .then((buffer) => {
      parseExcelBuffer(buffer);
    })
    .catch((err) => {
      console.warn('默认路径加载失败，可使用导入文件按钮:', err);
    });
}

// 读取 File 对象
function readExcelFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    parseExcelBuffer(e.target.result);
  };
  reader.readAsArrayBuffer(file);
}

// 解析 Excel ArrayBuffer
function parseExcelBuffer(buffer) {
  try {
    const workbook = XLSX.read(buffer, { type: 'array' });
    const sheetName = workbook.SheetNames.includes('1-数据进度表')
      ? '1-数据进度表'
      : workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    globalRawData = [];
    for (let i = 2; i < jsonRows.length; i++) {
      const row = jsonRows[i];
      if (row && row[0] !== undefined && row[0] !== null && String(row[0]).trim() !== '') {
        globalRawData.push(row);
      }
    }

    updateDashboard();
  } catch (err) {
    alert('解析 Excel 文件失败，请检查文件格式！');
    console.error(err);
  }
}

// 数值安全转换函数
function toNum(val) {
  if (val === undefined || val === null) return 0;
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  const parsed = parseFloat(String(val).replace(/,/g, '').trim());
  return isNaN(parsed) ? 0 : parsed;
}

// DMP与表型统一统计口径：排除Excel中的“【总计】”汇总行，按项目行计数。
function getGovernanceStats(rows) {
  const projectRows = rows.filter((r) => String(r[2] ?? '').trim() !== '【总计】');
  const totalProjects = projectRows.length;
  const dmpSubmitted = projectRows.filter((r) => String(r[43] ?? '').trim() === '是').length;
  const phenoSubmitted = projectRows.filter((r) => toNum(r[62]) !== 0).length;

  return {
    totalProjects,
    dmpSubmitted,
    phenoSubmitted,
    dmpRate: totalProjects > 0 ? (dmpSubmitted / totalProjects) * 100 : 0,
    phenoRate: totalProjects > 0 ? (phenoSubmitted / totalProjects) * 100 : 0
  };
}

// 获取筛选后的数据集
function getFilteredData() {
  return globalRawData.filter((r) => {
    if (currentProjectFilter === 'ALL') return true;
    return String(r[0]).trim() === currentProjectFilter;
  });
}

// 动态获取排名柱状图 Morandi 调色盘
function getRateColor(rate) {
  if (rate >= 90) return '#10b981'; // 90%+ 标准中绿
  if (rate >= 60) return '#f59e0b'; // 60-90% 中琥珀
  if (rate >= 30) return '#2563eb'; // 30-60% 皇家蓝
  return '#e11d48';                 // <30% 柔和红
}

// 动态获取不同类别 Morandi 进度条颜色
function getProgressBarColor(metricType, rate) {
  if (metricType === 'contract') {
    if (rate >= 90) return '#4f46e5'; // 90%+ 靛紫蓝
    if (rate >= 60) return '#818cf8'; // 60-90% 柔靛紫
    return '#c7d2fe';                 // <60% 淡薰衣草紫蓝
  }
  if (metricType === 'sample') {
    if (rate >= 90) return '#0284c7'; // 90%+ 深海蔚蓝
    if (rate >= 60) return '#38bdf8'; // 60-90% 柔青蓝
    return '#bae6fd';                 // <60% 淡青蓝
  }
  if (metricType === 'data') {
    if (rate >= 90) return '#10b981'; // 90%+ 标准中绿
    if (rate >= 60) return '#f59e0b'; // 60-90% 暖琥珀
    return '#6ee7b7';                 // <60% 薄荷淡绿
  }
  if (metricType === 'dmp') {
    if (rate >= 90) return '#7c3aed'; // 90%+ 雅致紫
    if (rate >= 60) return '#a78bfa'; // 60-90% 薰衣草紫
    return '#ddd6fe';                 // <60% 淡紫
  }
  return '#2563eb';
}

// 动态生成专项 Tag HTML (不同专项不同配色，背景框调平圆角为 5px)
function getProjectTagHtml(projName) {
  if (!projName || projName === '—' || projName === 'None') return '—';

  const p = String(projName).trim();

  // 专属项目配色词典 (背景, 文字, 边框)
  const colorMap = {
    'EBP': { bg: '#e0f2fe', color: '#0369a1', border: '#bae6fd' },       // 天蓝/深蓝
    'HGP2': { bg: '#d1fae5', color: '#047857', border: '#a7f3d0' },      // 翡翠绿
    '肿瘤专项': { bg: '#ffe4e6', color: '#be123c', border: '#fecdd3' },  // 玫瑰红/绯红
    '时空细胞模型': { bg: '#f3e8ff', color: '#6b21a8', border: '#e9d5ff' },  // 雅致紫
    '生物制造': { bg: '#fef3c7', color: '#b45309', border: '#fde68a' }   // 暖琥珀/暖橙
  };

  // 动态回退调色板 (保障未预设的其他项目也有高质感专属颜色)
  const fallbackPalettes = [
    { bg: '#cff4fc', color: '#055160', border: '#b6effb' }, // 青色
    { bg: '#ccfbf1', color: '#0f766e', border: '#99f6e4' }, // 薄荷绿
    { bg: '#e0e7ff', color: '#3730a3', border: '#c7d2fe' }, // 靛蓝
    { bg: '#fae8ff', color: '#86198f', border: '#f5d0fe' }, // 洋红
    { bg: '#f1f5f9', color: '#334155', border: '#cbd5e1' }  // 蓝灰
  ];

  let styleObj = colorMap[p];
  if (!styleObj) {
    let hash = 0;
    for (let i = 0; i < p.length; i++) {
      hash = p.charCodeAt(i) + ((hash << 5) - hash);
    }
    const index = Math.abs(hash) % fallbackPalettes.length;
    styleObj = fallbackPalettes[index];
  }

  return `<span class="badge-project" style="background-color: ${styleObj.bg}; color: ${styleObj.color}; border-color: ${styleObj.border};">${p}</span>`;
}

// 核心大屏数据更新逻辑
function updateDashboard() {
  const filteredData = getFilteredData();

  renderKPIs(filteredData);
  renderProjectMatrix(filteredData);
  renderRankChart(filteredData);
  updateDrilldownTable();
}

// 1. 渲染顶部 3 大 KPI 卡片
function renderKPIs(data) {
  // 样本计算
  const samplePlan = data.reduce((acc, r) => acc + toNum(r[9]), 0);
  const sampleDone = data.reduce((acc, r) => acc + toNum(r[10]), 0);
  const sampleRate = samplePlan > 0 ? (sampleDone / samplePlan) * 100 : 0;

  document.getElementById('kpiSampleRate').innerText = `${sampleRate.toFixed(1)}%`;
  document.getElementById('kpiSampleSub').innerText = `${sampleDone.toLocaleString()} / ${samplePlan.toLocaleString()} 例`;

  renderSampleSparkline(sampleRate);

  // 数据生产 (Gb -> PB)
  const dataPlanGb = data.reduce((acc, r) => acc + toNum(r[11]), 0);
  const dataActualGb = data.reduce((acc, r) => acc + toNum(r[12]), 0);
  const dataProdRate = dataPlanGb > 0 ? (dataActualGb / dataPlanGb) * 100 : 0;
  const dataPlanPb = dataPlanGb / 1024 / 1024;
  const dataActualPb = dataActualGb / 1024 / 1024;

  document.getElementById('kpiDataProdRate').innerText = `${dataProdRate.toFixed(1)}%`;
  document.getElementById('kpiDataProdSub').innerText = `${dataActualPb.toFixed(2)} / ${dataPlanPb.toFixed(2)} PB`;

  renderDataSparkline(dataProdRate);

  // 3. DMP与表型治理：均按项目数统计，分母为全部项目数
  const {
    totalProjects,
    dmpSubmitted,
    phenoSubmitted,
    dmpRate,
    phenoRate
  } = getGovernanceStats(data);

  document.getElementById('kpiDmpRate').innerText = `${dmpRate.toFixed(1)}%`;
  document.getElementById('kpiDmpSubtext').innerText = `${dmpSubmitted} / ${totalProjects} 个项目`;
  document.getElementById('kpiDmpBar').style.width = `${Math.min(dmpRate, 100)}%`;

  document.getElementById('kpiPhenoRate').innerText = `${phenoRate.toFixed(1)}%`;
  document.getElementById('kpiPhenoSubtext').innerText = `${phenoSubmitted} / ${totalProjects} 个项目`;
  document.getElementById('kpiPhenoBar').style.width = `${Math.min(phenoRate, 100)}%`;
}

// 渲染样本 Sparkline ECharts 曲线图
function renderSampleSparkline(currentRate) {
  if (!sparklineChart) return;
  const mockTrend = [65, 68, 70, 72, 75, 78, 81, 84, 83, 86, 88, Math.round(currentRate)];
  const option = {
    grid: { left: 0, right: 10, top: 10, bottom: 0 },
    xAxis: { type: 'category', show: false, data: mockTrend.map((_, i) => i) },
    yAxis: { type: 'value', show: false, min: 50 },
    series: [
      {
        data: mockTrend,
        type: 'line',
        smooth: true,
        symbol: 'circle',
        symbolSize: (value, params) => (params.dataIndex === mockTrend.length - 1 ? 6 : 0),
        itemStyle: { color: '#0284c7' },
        lineStyle: { width: 3, color: '#0284c7' },
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: 'rgba(2, 132, 199, 0.25)' },
            { offset: 1, color: 'rgba(2, 132, 199, 0.0)' }
          ])
        }
      }
    ]
  };
  sparklineChart.setOption(option);
}

// 渲染数据进度 Sparkline 中绿趋势曲线图
function renderDataSparkline(currentRate) {
  if (!dataSparklineChart) return;
  const mockTrend = [60, 63, 67, 71, 74, 76, 79, 81, 82, 84, 85, Math.round(currentRate)];
  const option = {
    grid: { left: 0, right: 10, top: 10, bottom: 0 },
    xAxis: { type: 'category', show: false, data: mockTrend.map((_, i) => i) },
    yAxis: { type: 'value', show: false, min: 50 },
    series: [
      {
        data: mockTrend,
        type: 'line',
        smooth: true,
        symbol: 'circle',
        symbolSize: (value, params) => (params.dataIndex === mockTrend.length - 1 ? 6 : 0),
        itemStyle: { color: '#10b981' },
        lineStyle: { width: 3, color: '#10b981' },
        areaStyle: {
          color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
            { offset: 0, color: 'rgba(16, 185, 129, 0.25)' },
            { offset: 1, color: 'rgba(16, 185, 129, 0.0)' }
          ])
        }
      }
    ]
  };
  dataSparklineChart.setOption(option);
}

function renderProjectMatrix(data) {
  const matrixBody = document.getElementById('projectMatrixBody');
  matrixBody.innerHTML = '';

  const cardsContainer = document.getElementById('projectMatrixCards');
  if (cardsContainer) cardsContainer.innerHTML = '';

  const projects = ['EBP', 'HGP2', '肿瘤专项', '虚拟细胞', '生物制造'];

  projects.forEach((projName) => {
    const projRows = data.filter((r) => String(r[0]).trim() === projName);
    if (projRows.length === 0 && currentProjectFilter !== 'ALL' && currentProjectFilter !== projName) {
      return;
    }

    const partnerSet = new Set(projRows.map((r) => r[2]).filter(Boolean));

    // 样本进度
    const sPlan = projRows.reduce((a, r) => a + toNum(r[9]), 0);
    const sDone = projRows.reduce((a, r) => a + toNum(r[10]), 0);
    const sRate = sPlan > 0 ? (sDone / sPlan) * 100 : 0;

    // 数据生产
    const dPlan = projRows.reduce((a, r) => a + toNum(r[11]), 0);
    const dActual = projRows.reduce((a, r) => a + toNum(r[12]), 0);
    const dRate = dPlan > 0 ? (dActual / dPlan) * 100 : 0;

    // 签署完成率
    const contractTotal = projRows.length;
    const contractSigned = projRows.filter((r) =>
      ['已签署，执行中', '已完成'].includes(String(r[4]).trim())
    ).length;
    const cRate = contractTotal > 0 ? (contractSigned / contractTotal) * 100 : 0;

    // DMP与表型均按项目数统计
    const {
      totalProjects,
      dmpSubmitted,
      phenoSubmitted,
      dmpRate,
      phenoRate
    } = getGovernanceStats(projRows);

    const signTooltip = `已签署：${contractSigned} / ${contractTotal} 个合同`;
    const sampleTooltip = `已完成样本：${sDone.toLocaleString()} / ${sPlan.toLocaleString()} 例`;
    const dataTooltip = `实际产出：${(dActual/1024/1024).toFixed(2)} / ${(dPlan/1024/1024).toFixed(2)} PB`;
    const dmpTooltip = `DMP提交：${dmpSubmitted} / ${totalProjects} 个项目 (${dmpRate.toFixed(1)}%) | 表型汇交：${phenoSubmitted} / ${totalProjects} 个项目 (${phenoRate.toFixed(1)}%)`;

    // 1. 桌面端表格行
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="proj-name-cell" style="color: #2563eb; font-weight: 600; cursor: pointer;" onclick="selectProjectFilter('${projName}')">${projName}</td>
      <td style="text-align: right; padding-right: 14px; font-weight: 500;">${partnerSet.size}</td>
      <td>
        ${contractTotal === 0 || projName === '生物制造' ? '—' : `
        <div class="matrix-metric-cell">
          <div class="table-bar-cell tooltip-cell" data-tooltip="${signTooltip}">
            <span class="table-bar-text">${cRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(cRate, 100)}%; background-color: ${getProgressBarColor('contract', cRate)};"></div>
            </div>
          </div>
          <div class="cell-subtext-total">${contractSigned} / ${contractTotal} 个</div>
        </div>
        `}
      </td>
      <td>
        <div class="matrix-metric-cell">
          <div class="table-bar-cell tooltip-cell" data-tooltip="${sampleTooltip}">
            <span class="table-bar-text">${sRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(sRate, 100)}%; background-color: ${getProgressBarColor('sample', sRate)};"></div>
            </div>
          </div>
          <div class="cell-subtext-total">${sDone.toLocaleString()} / ${sPlan.toLocaleString()} 例</div>
        </div>
      </td>
      <td>
        <div class="matrix-metric-cell">
          <div class="table-bar-cell tooltip-cell" data-tooltip="${dataTooltip}">
            <span class="table-bar-text">${dRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(dRate, 100)}%; background-color: ${getProgressBarColor('data', dRate)};"></div>
            </div>
          </div>
          <div class="cell-subtext-total">${(dActual/1024/1024).toFixed(2)} / ${(dPlan/1024/1024).toFixed(2)} PB</div>
        </div>
      </td>
      <td>
        <div class="governance-metrics tooltip-cell" data-tooltip="${dmpTooltip}">
          <div class="governance-row">
            <span class="governance-label">DMP</span>
            <span class="governance-rate">${dmpRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(dmpRate, 100)}%; background-color: ${getProgressBarColor('dmp', dmpRate)};"></div>
            </div>
          </div>
          <div class="governance-counts">DMP ${dmpSubmitted}/${totalProjects} · 表型 ${phenoSubmitted}/${totalProjects} 个项目</div>
        </div>
      </td>
    `;
    matrixBody.appendChild(tr);

    // 2. 移动端自适应卡片
    if (cardsContainer) {
      const card = document.createElement('div');
      card.className = 'matrix-mobile-card';
      card.innerHTML = `
        <div class="mobile-card-header">
          <div class="mobile-card-title-group" onclick="selectProjectFilter('${projName}')">
            ${getProjectTagHtml(projName)}
          </div>
          <span class="mobile-card-partner-badge">合作方：<strong>${partnerSet.size}</strong> 家</span>
        </div>
        <div class="mobile-card-metrics-grid">
          <!-- 合同签署率 -->
          <div class="mobile-metric-item">
            <div class="mobile-metric-top">
              <span class="mobile-metric-label">合同签署</span>
              <span class="mobile-metric-val text-purple">${contractTotal === 0 || projName === '生物制造' ? '—' : `${cRate.toFixed(1)}%`}</span>
            </div>
            ${contractTotal === 0 || projName === '生物制造' ? '<div class="mobile-metric-subtext">—</div>' : `
            <div class="progress-bar-bg">
              <div class="progress-bar-fill" style="width: ${Math.min(cRate, 100)}%; background-color: ${getProgressBarColor('contract', cRate)};"></div>
            </div>
            <div class="mobile-metric-subtext">${contractSigned} / ${contractTotal} 个</div>
            `}
          </div>

          <!-- 样本进度 -->
          <div class="mobile-metric-item">
            <div class="mobile-metric-top">
              <span class="mobile-metric-label">样本进度</span>
              <span class="mobile-metric-val text-teal">${sRate.toFixed(1)}%</span>
            </div>
            <div class="progress-bar-bg">
              <div class="progress-bar-fill" style="width: ${Math.min(sRate, 100)}%; background-color: ${getProgressBarColor('sample', sRate)};"></div>
            </div>
            <div class="mobile-metric-subtext">${sDone.toLocaleString()} / ${sPlan.toLocaleString()} 例</div>
          </div>

          <!-- 数据生产 -->
          <div class="mobile-metric-item">
            <div class="mobile-metric-top">
              <span class="mobile-metric-label">数据生产</span>
              <span class="mobile-metric-val text-green">${dRate.toFixed(1)}%</span>
            </div>
            <div class="progress-bar-bg">
              <div class="progress-bar-fill" style="width: ${Math.min(dRate, 100)}%; background-color: ${getProgressBarColor('data', dRate)};"></div>
            </div>
            <div class="mobile-metric-subtext">${(dActual/1024/1024).toFixed(2)} / ${(dPlan/1024/1024).toFixed(2)} PB</div>
          </div>

          <!-- DMP 与 表型治理 -->
          <div class="mobile-metric-item">
            <div class="mobile-metric-top">
              <span class="mobile-metric-label">DMP / 表型</span>
              <span class="mobile-metric-val text-primary">${dmpRate.toFixed(0)}% / ${phenoRate.toFixed(0)}%</span>
            </div>
            <div class="progress-bar-bg">
              <div class="progress-bar-fill" style="width: ${Math.min(dmpRate, 100)}%; background-color: ${getProgressBarColor('dmp', dmpRate)};"></div>
            </div>
            <div class="mobile-metric-subtext">DMP ${dmpSubmitted}/${totalProjects} · 表型 ${phenoSubmitted}/${totalProjects}</div>
          </div>
        </div>
      `;
      cardsContainer.appendChild(card);
    }
  });
}

function selectProjectFilter(projName) {
  const projectSelect = document.getElementById('projectSelect');
  projectSelect.value = projName;
  currentProjectFilter = projName;
  updateDashboard();
}

// 3. 渲染专项完成度排名 Panel
function renderRankChart(data) {
  if (!rankChart) return;
  const projects = ['EBP', 'HGP2', '时空细胞模型', '生物制造', '肿瘤专项'];

  const rankData = projects.map((pName) => {
    const pRows = data.filter((r) => String(r[0]).trim() === pName);

    let rate = 0;
    let labelExtra = '';

    if (currentRankMetric === 'sample') {
      const sPlan = pRows.reduce((a, r) => a + toNum(r[9]), 0);
      const sDone = pRows.reduce((a, r) => a + toNum(r[10]), 0);
      rate = sPlan > 0 ? (sDone / sPlan) * 100 : 0;
      labelExtra = `${rate.toFixed(1)}%`;
    } else if (currentRankMetric === 'contract') {
      const cTotal = pRows.length;
      const cSigned = pRows.filter((r) => ['已签署，执行中', '已完成'].includes(String(r[4]).trim())).length;
      rate = cTotal > 0 ? (cSigned / cTotal) * 100 : 0;
      labelExtra = `${rate.toFixed(1)}% (${cSigned}/${cTotal}个)`;
    } else {
      const dPlan = pRows.reduce((a, r) => a + toNum(r[11]), 0);
      const dActual = pRows.reduce((a, r) => a + toNum(r[12]), 0);
      rate = dPlan > 0 ? (dActual / dPlan) * 100 : 0;
      labelExtra = `${rate.toFixed(1)}%`;
    }

    return { name: pName, value: parseFloat(rate.toFixed(1)), labelText: labelExtra };
  });

  rankData.sort((a, b) => a.value - b.value);

  const yNames = rankData.map((d) => d.name);

  const isMobile = window.innerWidth <= 768;
  const option = {
    grid: {
      left: isMobile ? '24%' : '15%',
      right: isMobile ? '18%' : '24%',
      top: '8%',
      bottom: '5%'
    },
    xAxis: { type: 'value', max: 100, show: false },
    yAxis: {
      type: 'category',
      data: yNames,
      axisLine: { show: false },
      axisTick: { show: false },
      axisLabel: { color: '#2563eb', fontWeight: 600, fontSize: isMobile ? 12.5 : 13 }
    },
    series: [
      {
        type: 'bar',
        data: rankData.map((d) => ({
          value: d.value,
          labelText: d.labelText
        })),
        barWidth: 10,
        itemStyle: {
          borderRadius: [0, 5, 5, 0],
          color: (params) => (currentRankMetric === 'contract' ? '#4f46e5' : getRateColor(params.value))
        },
        label: {
          show: true,
          position: 'right',
          formatter: (params) => {
            if (isMobile && currentRankMetric === 'contract') {
              return `${params.value.toFixed(1)}%`;
            }
            return params.data.labelText;
          },
          color: '#334155',
          fontWeight: 'bold',
          fontSize: isMobile ? 11.5 : 12
        }
      }
    ]
  };

  rankChart.setOption(option);
}

// 辅助：生成带高质感 SVG 矢量排序图标的表头 HTML
function getSortHeaderHtml(title, key, extraClass = '') {
  const isActive = currentSort.key === key;

  let iconSvg = `
    <svg class="sort-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round">
      <path d="m7 15 5 5 5-5"/>
      <path d="m7 9 5-5 5 5"/>
    </svg>
  `;

  if (isActive) {
    if (currentSort.dir === 'desc') {
      iconSvg = `
        <svg class="sort-svg active-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="m7 10 5 5 5-5"/>
        </svg>
      `;
    } else {
      iconSvg = `
        <svg class="sort-svg active-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <path d="m7 14 5-5 5 5"/>
        </svg>
      `;
    }
  }

  const activeClass = isActive ? ' active-sort' : '';
  return `
    <th class="sortable-th ${extraClass}${activeClass}" onclick="handleSort('${key}')">
      <div class="th-sort-inner">
        <span>${title}</span>
        <span class="sort-icon-box">${iconSvg}</span>
      </div>
    </th>
  `;
}

// 5. 渲染底部专项分析明细表 (组学进展：样本计划贴合样本完成率，预计数据量贴合数据生产进度)
function updateDrilldownTable() {
  const filteredData = getFilteredData();
  const searchKeyword = document.getElementById('tableSearchInput').value.trim().toLowerCase();

  const subtitle = document.getElementById('drilldownSubtitle');
  const thead = document.getElementById('drilldownThead');
  const tbody = document.getElementById('drilldownTbody');
  const drilldownCardsContainer = document.getElementById('drilldownCards');

  tbody.innerHTML = '';
  if (drilldownCardsContainer) drilldownCardsContainer.innerHTML = '';

  if (currentTab === 'omics') {
    subtitle.innerText = ' — 组学维度明细';
    thead.innerHTML = `
      <tr>
        <th style="width: 120px;">组学类型</th>
        ${getSortHeaderHtml('关联合同数', 'contracts', 'th-col-contracts')}
        ${getSortHeaderHtml('样本计划 (例)', 'samplePlanSum', 'th-right-tight')}
        ${getSortHeaderHtml('样本完成率', 'sampleRate', 'th-left-tight')}
        ${getSortHeaderHtml('预计数据量', 'estGb', 'th-right-tight')}
        ${getSortHeaderHtml('数据生产进度', 'prodRate', 'th-left-tight')}
      </tr>
    `;

    const omicsDefs = [
      { name: 'WGS', estIdx: 13, actIdx: 14, subIdx: 46 },
      { name: 'WGBS', estIdx: 17, actIdx: 18, subIdx: 48 },
      { name: 'Bulk-RNA', estIdx: 21, actIdx: 22, subIdx: 50 },
      { name: '单细胞', estIdx: 25, actIdx: 26, subIdx: 52 },
      { name: '时空组学', estIdx: 23, actIdx: 24, subIdx: 51 },
      { name: '代谢组', estIdx: 35, actIdx: 36, subIdx: 58 },
      { name: 'META', estIdx: 29, actIdx: 30, subIdx: 54 },
      { name: 'Cyclone', estIdx: 31, actIdx: 32, subIdx: 55 }
    ];

    let omicsList = omicsDefs.map((omic) => {
      let contracts = 0;
      let estGb = 0, actGb = 0, subGb = 0;
      let samplePlanSum = 0, sampleDoneSum = 0;

      filteredData.forEach((r) => {
        const e = toNum(r[omic.estIdx]);
        const a = toNum(r[omic.actIdx]);
        const s = toNum(r[omic.subIdx]);

        if (e > 0 || a > 0 || s > 0) {
          contracts++;
          estGb += e;
          actGb += a;
          subGb += s;
          samplePlanSum += toNum(r[9]);
          sampleDoneSum += toNum(r[10]);
        }
      });

      const sampleRate = samplePlanSum > 0 ? (sampleDoneSum / samplePlanSum) * 100 : 85.0;
      const prodRate = estGb > 0 ? (actGb / estGb) * 100 : (actGb > 0 ? 100 : 0);
      const subRate = estGb > 0 ? (subGb / estGb) * 100 : (subGb > 0 ? 100 : 0);

      return {
        name: omic.name,
        contracts,
        estGb,
        actGb,
        subGb,
        samplePlanSum,
        sampleDoneSum,
        sampleRate,
        prodRate,
        subRate
      };
    });

    // 搜索过滤
    omicsList = omicsList.filter((item) => {
      if (!searchKeyword) return true;
      return item.name.toLowerCase().includes(searchKeyword);
    });

    // 动态排序
    if (currentSort.key) {
      omicsList.sort((a, b) => {
        const valA = a[currentSort.key] || 0;
        const valB = b[currentSort.key] || 0;
        return currentSort.dir === 'asc' ? valA - valB : valB - valA;
      });
    }

    omicsList.forEach((item) => {
      const estPb = item.estGb / 1024 / 1024;
      const actPb = item.actGb / 1024 / 1024;
      const subPb = item.subGb / 1024 / 1024;

      const sampleTooltip = `样本完成率详情：${item.sampleDoneSum.toLocaleString()} / ${item.samplePlanSum.toLocaleString()} 例 (${item.sampleRate.toFixed(1)}%)`;
      const prodTooltip = `数据生产进度详情：${actPb.toFixed(2)} PB / ${estPb.toFixed(2)} PB (${item.actGb.toLocaleString()} Gb)`;
      const subTooltip = `汇交进度详情：${subPb.toFixed(2)} PB / ${estPb.toFixed(2)} PB (${item.subGb.toLocaleString()} Gb)`;

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td class="text-bold" style="color: #2563eb;">${item.name}</td>
        <td>${item.contracts || '—'}</td>
        <td class="text-right-compact">${item.samplePlanSum ? item.samplePlanSum.toLocaleString() : '—'}</td>
        <td class="text-left-compact" title="${sampleTooltip}">
          <div class="table-bar-cell compact-bar tooltip-cell" data-tooltip="${sampleTooltip}">
            <span class="table-bar-text">${item.sampleRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(item.sampleRate, 100)}%; background-color: ${getProgressBarColor('sample', item.sampleRate)};"></div>
            </div>
          </div>
        </td>
        <td class="text-right-compact">${estPb > 0 ? estPb.toFixed(2) + ' PB' : '—'}</td>
        <td class="text-left-compact" title="${prodTooltip}">
          <div class="table-bar-cell compact-bar tooltip-cell" data-tooltip="${prodTooltip}">
            <span class="table-bar-text">${item.prodRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(item.prodRate, 100)}%; background-color: ${getProgressBarColor('data', item.prodRate)};"></div>
            </div>
          </div>
        </td>
      `;
      tbody.appendChild(tr);

      // 移动端卡片渲染
      if (drilldownCardsContainer) {
        const card = document.createElement('div');
        card.className = 'drilldown-mobile-card';
        card.innerHTML = `
          <div class="drilldown-card-top">
            <span class="drilldown-card-title">${item.name}</span>
            <span class="drilldown-card-badge">关联合同：<strong>${item.contracts || 0}</strong> 个</span>
          </div>
          <div class="drilldown-card-body">
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">样本完成率</span>
                <span class="mobile-metric-val text-teal">${item.sampleRate.toFixed(1)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(item.sampleRate, 100)}%; background-color: ${getProgressBarColor('sample', item.sampleRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${item.sampleDoneSum ? item.sampleDoneSum.toLocaleString() : 0} / ${item.samplePlanSum ? item.samplePlanSum.toLocaleString() : 0} 例</div>
            </div>
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">数据生产进度</span>
                <span class="mobile-metric-val text-green">${item.prodRate.toFixed(1)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(item.prodRate, 100)}%; background-color: ${getProgressBarColor('data', item.prodRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${estPb > 0 ? estPb.toFixed(2) + ' PB' : '—'}</div>
            </div>
          </div>
        `;
        drilldownCardsContainer.appendChild(card);
      }
    });
  } else if (currentTab === 'sampletype') {
    subtitle.innerText = ' — 样本类型分布及进度';
    thead.innerHTML = `
      <tr>
        <th style="width: 160px;">样本类型</th>
        <th style="width: 160px;">关联专项</th>
        ${getSortHeaderHtml('关联合同数', 'contracts', 'th-col-contracts')}
        ${getSortHeaderHtml('样本计划 (例)', 'sPlan', 'th-right-tight')}
        ${getSortHeaderHtml('已送样进度', 'sRate', 'th-left-tight')}
        ${getSortHeaderHtml('数据计划 (PB)', 'dPlanGb', 'th-right-tight')}
        ${getSortHeaderHtml('实际产出进度', 'dRate', 'th-left-tight')}
      </tr>
    `;

    const stMap = {};

    filteredData.forEach((r) => {
      const st = r[8] !== undefined && r[8] !== null && String(r[8]).trim() !== '' ? String(r[8]).trim() : '未指定类型';
      const proj = r[0] ? String(r[0]).trim() : '';

      if (!stMap[st]) {
        stMap[st] = {
          name: st,
          contracts: 0,
          projects: new Set(),
          sPlan: 0,
          sDone: 0,
          dPlanGb: 0,
          dActGb: 0
        };
      }

      stMap[st].contracts++;
      if (proj) stMap[st].projects.add(proj);
      stMap[st].sPlan += toNum(r[9]);
      stMap[st].sDone += toNum(r[10]);
      stMap[st].dPlanGb += toNum(r[11]);
      stMap[st].dActGb += toNum(r[12]);
    });

    let stList = Object.values(stMap).map((item) => {
      item.sRate = item.sPlan > 0 ? (item.sDone / item.sPlan) * 100 : 0;
      item.dRate = item.dPlanGb > 0 ? (item.dActGb / item.dPlanGb) * 100 : 0;
      return item;
    });

    // 搜索过滤
    stList = stList.filter((item) => {
      const projsText = Array.from(item.projects).join(', ');
      if (!searchKeyword) return true;
      return item.name.toLowerCase().includes(searchKeyword) || projsText.toLowerCase().includes(searchKeyword);
    });

    // 动态排序
    if (currentSort.key) {
      stList.sort((a, b) => {
        const valA = a[currentSort.key] || 0;
        const valB = b[currentSort.key] || 0;
        return currentSort.dir === 'asc' ? valA - valB : valB - valA;
      });
    } else {
      stList.sort((a, b) => b.sPlan - a.sPlan);
    }

    stList.forEach((item) => {
      const dPlanPb = item.dPlanGb / 1024 / 1024;
      const dActPb = item.dActGb / 1024 / 1024;

      const sampleTooltip = `已送样：${item.sDone.toLocaleString()} / ${item.sPlan.toLocaleString()} 例`;
      const dataTooltip = `实际产出：${dActPb.toFixed(2)} PB / ${dPlanPb.toFixed(2)} PB`;

      const projBadgesHtml = Array.from(item.projects)
        .map((p) => getProjectTagHtml(p))
        .join(' ');

      let displaySampleName = item.name;
      let sampleNameTd = `<td class="text-bold" style="color: #2563eb;">${item.name}</td>`;
      if (item.name.length > 14) {
        displaySampleName = item.name.substring(0, 12) + '...';
        sampleNameTd = `
          <td>
            <div class="tooltip-cell text-bold" data-tooltip="${item.name}" style="color: #2563eb; display: inline-block; cursor: pointer;">
              ${displaySampleName}
            </div>
          </td>
        `;
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        ${sampleNameTd}
        <td>${projBadgesHtml || '—'}</td>
        <td>${item.contracts}</td>
        <td class="text-right-compact">${item.sPlan ? item.sPlan.toLocaleString() : '—'}</td>
        <td class="text-left-compact" title="${sampleTooltip}">
          ${item.sPlan > 0 ? `
          <div class="table-bar-cell compact-bar tooltip-cell" data-tooltip="${sampleTooltip}">
            <span class="table-bar-text">${item.sRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(item.sRate, 100)}%; background-color: ${getProgressBarColor('sample', item.sRate)};"></div>
            </div>
          </div>
          ` : '—'}
        </td>
        <td class="text-right-compact">${dPlanPb > 0 ? dPlanPb.toFixed(2) + ' PB' : '—'}</td>
        <td class="text-left-compact" title="${dataTooltip}">
          ${item.dPlanGb > 0 ? `
          <div class="table-bar-cell compact-bar tooltip-cell" data-tooltip="${dataTooltip}">
            <span class="table-bar-text">${item.dRate.toFixed(1)}%</span>
            <div class="mini-progress-bg">
              <div class="mini-progress-fill" style="width: ${Math.min(item.dRate, 100)}%; background-color: ${getProgressBarColor('data', item.dRate)};"></div>
            </div>
          </div>
          ` : '—'}
        </td>
      `;
      tbody.appendChild(tr);

      // 移动端卡片渲染
      if (drilldownCardsContainer) {
        const card = document.createElement('div');
        card.className = 'drilldown-mobile-card';
        card.innerHTML = `
          <div class="drilldown-card-top">
            <span class="drilldown-card-title">${item.name}</span>
            <span class="drilldown-card-badge">关联合同：<strong>${item.contracts || 0}</strong> 个</span>
          </div>
          <div style="margin-bottom: 6px;">${projBadgesHtml || '—'}</div>
          <div class="drilldown-card-body">
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">已送样进度</span>
                <span class="mobile-metric-val text-teal">${item.sRate.toFixed(1)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(item.sRate, 100)}%; background-color: ${getProgressBarColor('sample', item.sRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${item.sDone ? item.sDone.toLocaleString() : 0} / ${item.sPlan ? item.sPlan.toLocaleString() : 0} 例</div>
            </div>
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">实际产出进度</span>
                <span class="mobile-metric-val text-green">${item.dRate.toFixed(1)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(item.dRate, 100)}%; background-color: ${getProgressBarColor('data', item.dRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${dPlanPb > 0 ? dPlanPb.toFixed(2) + ' PB' : '—'}</div>
            </div>
          </div>
        `;
        drilldownCardsContainer.appendChild(card);
      }
    });
  } else if (currentTab === 'contract') {
    subtitle.innerText = ' — 合同维度明细';
    thead.innerHTML = `
      <tr>
        <th style="width: 120px;">合同编号</th>
        <th>专项</th>
        <th style="max-width: 160px;">合作方</th>
        <th>项目负责人</th>
        <th>合同状态</th>
        ${getSortHeaderHtml('合同例数', 'sPlan', 'th-right-tight')}
        ${getSortHeaderHtml('已送样数', 'sDone', 'th-right-tight')}
        ${getSortHeaderHtml('预计数据(Gb)', 'dPlan', 'th-right-tight')}
        ${getSortHeaderHtml('实际数据(Gb)', 'dAct', 'th-right-tight')}
        <th>DMP / 表型状态</th>
      </tr>
    `;

    let contractList = filteredData.map((r) => ({
      contractNo: r[5] ? String(r[5]).trim() : '暂无',
      proj: r[0] ? String(r[0]).trim() : '',
      partner: r[2] ? String(r[2]).trim() : '',
      lead: r[6] ? String(r[6]).trim() : '',
      status: r[4] ? String(r[4]).trim() : '',
      sPlan: toNum(r[9]),
      sDone: toNum(r[10]),
      dPlan: toNum(r[11]),
      dAct: toNum(r[12]),
      dmp: r[43] ? String(r[43]).trim() : '未填写',
      pDone: toNum(r[62])
    }));

    // 搜索过滤
    contractList = contractList.filter((item) => {
      if (!searchKeyword) return true;
      return (
        item.contractNo.toLowerCase().includes(searchKeyword) ||
        item.partner.toLowerCase().includes(searchKeyword) ||
        item.proj.toLowerCase().includes(searchKeyword) ||
        item.lead.toLowerCase().includes(searchKeyword) ||
        item.status.toLowerCase().includes(searchKeyword)
      );
    });

    // 动态排序
    if (currentSort.key) {
      contractList.sort((a, b) => {
        const valA = a[currentSort.key] || 0;
        const valB = b[currentSort.key] || 0;
        return currentSort.dir === 'asc' ? valA - valB : valB - valA;
      });
    }

    contractList.forEach((item) => {
      let dmpBadge = '<span class="badge badge-none"><span class="badge-dot"></span> 未填写</span>';
      if (item.dmp === '是') {
        if (item.pDone !== 0) {
          dmpBadge = `<span class="badge badge-done" title="DMP与表型均已提交: ${item.pDone}例"><span class="badge-dot"></span> DMP+表型</span>`;
        } else {
          dmpBadge = '<span class="badge badge-done"><span class="badge-dot"></span> 已提交DMP</span>';
        }
      } else if (item.dmp !== '未填写' && item.dmp !== 'None') {
        dmpBadge = `<span class="badge badge-check" title="${item.dmp}"><span class="badge-dot"></span> 待核对</span>`;
      }

      let contractDisplay = item.contractNo;
      if (item.contractNo.length > 15) {
        contractDisplay = item.contractNo.substring(0, 13) + '...';
      }

      const tr = document.createElement('tr');
      tr.innerHTML = `
        <td>
          <div class="tooltip-cell text-bold" data-tooltip="${item.contractNo}" style="font-family: monospace; display: inline-block; cursor: pointer; color: #2563eb;">
            ${contractDisplay}
          </div>
        </td>
        <td>${getProjectTagHtml(item.proj)}</td>
        <td style="max-width: 150px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;" title="${item.partner}">${item.partner}</td>
        <td>${item.lead}</td>
        <td>${item.status}</td>
        <td style="text-align: right; font-weight: 500; padding-right: 6px;">${item.sPlan.toLocaleString()}</td>
        <td style="text-align: right; font-weight: 500; padding-right: 6px;">${item.sDone.toLocaleString()}</td>
        <td style="text-align: right; font-weight: 500; padding-right: 6px;">${item.dPlan.toLocaleString()}</td>
        <td style="text-align: right; font-weight: 500; padding-right: 6px;">${item.dAct.toLocaleString()}</td>
        <td>${dmpBadge}</td>
      `;
      tbody.appendChild(tr);

      // 移动端卡片渲染
      if (drilldownCardsContainer) {
        const sRate = item.sPlan > 0 ? (item.sDone / item.sPlan) * 100 : 0;
        const dRate = item.dPlan > 0 ? (item.dAct / item.dPlan) * 100 : 0;

        const card = document.createElement('div');
        card.className = 'drilldown-mobile-card';
        card.innerHTML = `
          <div class="drilldown-card-top">
            <div class="mobile-card-title-group">
              ${getProjectTagHtml(item.proj)}
              <span style="font-family: monospace; font-weight: 600; font-size: 13px; color: #2563eb; margin-left: 6px;">${item.contractNo}</span>
            </div>
            ${dmpBadge}
          </div>
          <div style="font-size: 12.5px; color: var(--text-secondary); margin-bottom: 8px;">
            合作方：<strong>${item.partner || '—'}</strong> (${item.lead || '无负责人'}) | 状态：${item.status}
          </div>
          <div class="drilldown-card-body">
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">送样进度</span>
                <span class="mobile-metric-val text-teal">${sRate.toFixed(0)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(sRate, 100)}%; background-color: ${getProgressBarColor('sample', sRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${item.sDone ? item.sDone.toLocaleString() : 0} / ${item.sPlan ? item.sPlan.toLocaleString() : 0} 例</div>
            </div>
            <div class="mobile-metric-item">
              <div class="mobile-metric-top">
                <span class="mobile-metric-label">数据进度</span>
                <span class="mobile-metric-val text-green">${dRate.toFixed(0)}%</span>
              </div>
              <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: ${Math.min(dRate, 100)}%; background-color: ${getProgressBarColor('data', dRate)};"></div>
              </div>
              <div class="mobile-metric-subtext">${item.dAct ? item.dAct.toLocaleString() : 0} / ${item.dPlan ? item.dPlan.toLocaleString() : 0} Gb</div>
            </div>
          </div>
        `;
        drilldownCardsContainer.appendChild(card);
      }
    });
  }
}
