/**
 * 原煤指标 - 主应用逻辑（支持多工作簿）
 */

// 应用状态
const state = {
  rows: [],
  currentWorkbookId: null,
  workbooks: [],
  storageKey: 'coal_metrics_workbooks_v2',
  storageAvailable: false
};

/**
 * 检测 LocalStorage 是否可用
 */
function checkStorageAvailability() {
  try {
    const testKey = '__storage_test__';
    localStorage.setItem(testKey, 'test');
    localStorage.removeItem(testKey);
    state.storageAvailable = true;
    return true;
  } catch (e) {
    state.storageAvailable = false;
    console.error('LocalStorage 不可用:', e);
    showToast('警告：浏览器存储功能不可用，数据将无法保存', 'warning');
    return false;
  }
}

// DOM 元素
const elements = {
  tableBody: null,
  selectAll: null,
  btnAdd: null,
  btnDelete: null,
  btnSave: null,
  btnSaveAs: null,
  btnExport: null,
  btnReset: null,
  btnNewWorkbook: null,
  btnToggleSidebar: null,
  workbookList: null,
  workbookTitle: null,
  saveDialog: null,
  workbookNameInput: null,
  btnConfirmSave: null,
  btnCancelSave: null,
  sidebar: null,
  toast: null
};

/**
 * 初始化应用
 */
function init() {
  // 检测 LocalStorage 是否可用
  checkStorageAvailability();
  
  // 获取 DOM 元素
  elements.tableBody = document.getElementById('tableBody');
  elements.selectAll = document.getElementById('selectAll');
  elements.btnAdd = document.getElementById('btnAdd');
  elements.btnDelete = document.getElementById('btnDelete');
  elements.btnSave = document.getElementById('btnSave');
  elements.btnSaveAs = document.getElementById('btnSaveAs');
  elements.btnExport = document.getElementById('btnExport');
  elements.btnReset = document.getElementById('btnReset');
  elements.btnNewWorkbook = document.getElementById('btnNewWorkbook');
  elements.btnToggleSidebar = document.getElementById('btnToggleSidebar');
  elements.workbookList = document.getElementById('workbookList');
  elements.workbookTitle = document.getElementById('workbookTitle');
  elements.saveDialog = document.getElementById('saveDialog');
  elements.workbookNameInput = document.getElementById('workbookNameInput');
  elements.btnConfirmSave = document.getElementById('btnConfirmSave');
  elements.btnCancelSave = document.getElementById('btnCancelSave');
  elements.sidebar = document.querySelector('.sidebar');
  elements.toast = document.getElementById('toast');

  // 绑定事件
  elements.btnAdd.addEventListener('click', addRow);
  elements.btnDelete.addEventListener('click', removeSelected);
  elements.btnSave.addEventListener('click', saveCurrentWorkbook);
  elements.btnSaveAs.addEventListener('click', saveAsNewWorkbook);
  elements.btnExport.addEventListener('click', exportData);
  elements.btnReset.addEventListener('click', resetData);
  elements.btnNewWorkbook.addEventListener('click', createNewWorkbook);
  elements.selectAll.addEventListener('change', toggleSelectAll);
  elements.btnConfirmSave.addEventListener('click', confirmSaveWorkbook);
  elements.btnCancelSave.addEventListener('click', closeSaveDialog);
  
  if (elements.btnToggleSidebar) {
    elements.btnToggleSidebar.addEventListener('click', toggleSidebar);
  }

  // 显示存储状态
  if (!state.storageAvailable) {
    showToast('注意：数据保存功能不可用，请使用导出CSV备份数据', 'warning');
  }

  // 加载工作簿列表
  loadWorkbooks();

  // 加载或创建默认工作簿
  if (state.workbooks.length === 0) {
    createDefaultWorkbook();
  } else {
    // 加载第一个工作簿
    loadWorkbook(state.workbooks[0].id);
  }

  // 渲染工作簿列表
  renderWorkbookList();
}

/**
 * 创建空行对象
 */
function createEmptyRow(index) {
  return {
    index: index,
    date: '',
    unit: '',
    inbound: '',
    ash_total: '',
    ash_internal: '',
    sulfur: '',
    volatile: '',
    caking: '',
    middlings: '',
    gangue: '',
    selected: false
  };
}

/**
 * 初始化 6 行空数据
 */
function initEmptyRows() {
  state.rows = [];
  for (let i = 1; i <= 6; i++) {
    state.rows.push(createEmptyRow(i));
  }
}

/**
 * 加载所有工作簿
 */
function loadWorkbooks() {
  if (!state.storageAvailable) {
    console.warn('LocalStorage 不可用，无法加载工作簿');
    return;
  }
  
  try {
    const saved = localStorage.getItem(state.storageKey);
    if (saved) {
      state.workbooks = JSON.parse(saved);
      console.log('成功加载工作簿:', state.workbooks.length, '个');
    } else {
      console.log('没有找到已保存的工作簿');
    }
  } catch (e) {
    console.error('加载工作簿列表失败:', e);
    showToast('加载工作簿列表失败: ' + e.message, 'error');
  }
}

/**
 * 保存所有工作簿到本地存储
 */
function saveWorkbooks() {
  if (!state.storageAvailable) {
    showToast('存储功能不可用，无法保存数据。请使用导出CSV功能备份', 'error');
    return false;
  }
  
  try {
    const dataStr = JSON.stringify(state.workbooks);
    localStorage.setItem(state.storageKey, dataStr);
    console.log('成功保存工作簿:', state.workbooks.length, '个');
    return true;
  } catch (e) {
    console.error('保存工作簿列表失败:', e);
    if (e.name === 'QuotaExceededError') {
      showToast('存储空间不足，请删除一些工作簿或导出数据后清理', 'error');
    } else {
      showToast('保存失败: ' + e.message, 'error');
    }
    return false;
  }
}

/**
 * 创建默认工作簿
 */
function createDefaultWorkbook() {
  const workbook = {
    id: generateId(),
    name: '默认工作簿',
    rows: [],
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
  
  // 初始化 6 行空数据
  for (let i = 1; i <= 6; i++) {
    workbook.rows.push(createEmptyRow(i));
  }
  
  state.workbooks.push(workbook);
  saveWorkbooks();
  loadWorkbook(workbook.id);
}

/**
 * 创建新工作簿
 */
function createNewWorkbook() {
  elements.workbookNameInput.value = '';
  elements.saveDialog.classList.add('show');
  elements.workbookNameInput.focus();
  
  // 标记为新建模式
  elements.saveDialog.dataset.mode = 'new';
}

/**
 * 另存为新工作簿
 */
function saveAsNewWorkbook() {
  elements.workbookNameInput.value = '';
  elements.saveDialog.classList.add('show');
  elements.workbookNameInput.focus();
  
  // 标记为另存为模式
  elements.saveDialog.dataset.mode = 'saveas';
}

/**
 * 确认保存工作簿
 */
function confirmSaveWorkbook() {
  const name = elements.workbookNameInput.value.trim();
  
  if (!name) {
    showToast('请输入工作簿名称', 'warning');
    return;
  }
  
  const mode = elements.saveDialog.dataset.mode;
  
  if (mode === 'new') {
    // 创建新工作簿
    const workbook = {
      id: generateId(),
      name: name,
      rows: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    
    // 初始化 6 行空数据
    for (let i = 1; i <= 6; i++) {
      workbook.rows.push(createEmptyRow(i));
    }
    
    state.workbooks.push(workbook);
    saveWorkbooks();
    loadWorkbook(workbook.id);
    renderWorkbookList();
    showToast('新工作簿创建成功', 'success');
  } else if (mode === 'saveas') {
    // 另存为新工作簿
    const workbook = {
      id: generateId(),
      name: name,
      rows: JSON.parse(JSON.stringify(state.rows)), // 深拷贝当前数据
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    
    state.workbooks.push(workbook);
    saveWorkbooks();
    loadWorkbook(workbook.id);
    renderWorkbookList();
    showToast('工作簿另存为成功', 'success');
  }
  
  closeSaveDialog();
}

/**
 * 关闭保存对话框
 */
function closeSaveDialog() {
  elements.saveDialog.classList.remove('show');
  elements.workbookNameInput.value = '';
}

/**
 * 保存当前工作簿
 */
function saveCurrentWorkbook() {
  if (!state.currentWorkbookId) {
    showToast('没有打开的工作簿', 'warning');
    return;
  }
  
  const workbook = state.workbooks.find(wb => wb.id === state.currentWorkbookId);
  if (!workbook) {
    showToast('工作簿不存在', 'error');
    return;
  }
  
  workbook.rows = JSON.parse(JSON.stringify(state.rows));
  workbook.updatedAt = new Date().toISOString();
  
  if (saveWorkbooks()) {
    showToast('保存成功', 'success');
  }
}

/**
 * 加载工作簿
 */
function loadWorkbook(workbookId) {
  const workbook = state.workbooks.find(wb => wb.id === workbookId);
  if (!workbook) {
    showToast('工作簿不存在', 'error');
    return;
  }
  
  state.currentWorkbookId = workbookId;
  state.rows = JSON.parse(JSON.stringify(workbook.rows));
  
  // 更新标题
  elements.workbookTitle.textContent = workbook.name;
  
  // 渲染表格
  renderTable();
  
  // 更新工作簿列表选中状态
  renderWorkbookList();
}

/**
 * 渲染工作簿列表
 */
function renderWorkbookList() {
  elements.workbookList.innerHTML = '';
  
  state.workbooks.forEach(workbook => {
    const item = document.createElement('div');
    item.className = 'workbook-item';
    if (workbook.id === state.currentWorkbookId) {
      item.classList.add('active');
    }
    
    const nameSpan = document.createElement('span');
    nameSpan.className = 'workbook-name';
    nameSpan.textContent = workbook.name;
    nameSpan.title = workbook.name;
    
    const actions = document.createElement('div');
    actions.className = 'workbook-actions';
    
    const deleteBtn = document.createElement('button');
    deleteBtn.className = 'btn-delete-workbook';
    deleteBtn.textContent = '×';
    deleteBtn.title = '删除工作簿';
    deleteBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      deleteWorkbook(workbook.id);
    });
    
    actions.appendChild(deleteBtn);
    item.appendChild(nameSpan);
    item.appendChild(actions);
    
    item.addEventListener('click', () => {
      loadWorkbook(workbook.id);
      // 移动端关闭侧边栏
      if (window.innerWidth <= 768) {
        elements.sidebar.classList.remove('show');
      }
    });
    
    elements.workbookList.appendChild(item);
  });
}

/**
 * 删除工作簿
 */
function deleteWorkbook(workbookId) {
  if (state.workbooks.length === 1) {
    showToast('至少需要保留一个工作簿', 'warning');
    return;
  }
  
  const workbook = state.workbooks.find(wb => wb.id === workbookId);
  if (!workbook) return;
  
  if (!confirm(`确定要删除工作簿"${workbook.name}"吗？`)) {
    return;
  }
  
  state.workbooks = state.workbooks.filter(wb => wb.id !== workbookId);
  saveWorkbooks();
  
  // 如果删除的是当前工作簿，加载第一个工作簿
  if (state.currentWorkbookId === workbookId) {
    if (state.workbooks.length > 0) {
      loadWorkbook(state.workbooks[0].id);
    }
  }
  
  renderWorkbookList();
  showToast('工作簿已删除', 'success');
}

/**
 * 切换侧边栏（移动端）
 */
function toggleSidebar() {
  elements.sidebar.classList.toggle('show');
}

/**
 * 生成唯一 ID
 */
function generateId() {
  return 'wb_' + Date.now() + '_' + Math.random().toString(36).substr(2, 9);
}

/**
 * 从本地存储加载数据（旧版本兼容）
 */
function loadFromStorage() {
  try {
    const saved = localStorage.getItem('coal_metrics_v1');
    if (saved) {
      const data = JSON.parse(saved);
      if (data.rows && Array.isArray(data.rows)) {
        state.rows = data.rows;
      }
    }
  } catch (e) {
    console.error('加载数据失败:', e);
  }
}

/**
 * 渲染表格
 */
function renderTable() {
  // 清空表格
  elements.tableBody.innerHTML = '';

  // 渲染每一行
  state.rows.forEach((row, index) => {
    const tr = document.createElement('tr');
    tr.dataset.index = index;

    // 复选框
    const tdCheckbox = document.createElement('td');
    tdCheckbox.className = 'checkbox-col';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    checkbox.checked = row.selected || false;
    checkbox.addEventListener('change', (e) => {
      state.rows[index].selected = e.target.checked;
      updateSelectAllState();
    });
    tdCheckbox.appendChild(checkbox);
    tr.appendChild(tdCheckbox);

    // 序号
    const tdIndex = document.createElement('td');
    tdIndex.textContent = row.index;
    tdIndex.style.textAlign = 'center';
    tr.appendChild(tdIndex);

    // 日期
    const tdDate = document.createElement('td');
    const inputDate = document.createElement('input');
    inputDate.type = 'date';
    inputDate.value = row.date || '';
    inputDate.setAttribute('aria-label', '日期');
    inputDate.addEventListener('change', (e) => {
      state.rows[index].date = e.target.value;
    });
    tdDate.appendChild(inputDate);
    tr.appendChild(tdDate);

    // 单位
    const tdUnit = document.createElement('td');
    const inputUnit = document.createElement('input');
    inputUnit.type = 'text';
    inputUnit.value = row.unit || '';
    inputUnit.setAttribute('aria-label', '单位');
    inputUnit.addEventListener('input', (e) => {
      state.rows[index].unit = e.target.value;
    });
    tdUnit.appendChild(inputUnit);
    tr.appendChild(tdUnit);

    // 数值列
    const numericFields = [
      { key: 'inbound', label: '入库' },
      { key: 'ash_total', label: '全灰' },
      { key: 'ash_internal', label: '内灰' },
      { key: 'sulfur', label: '硫' },
      { key: 'volatile', label: '挥发' },
      { key: 'caking', label: '粘结' },
      { key: 'middlings', label: '中煤' },
      { key: 'gangue', label: '矸石' }
    ];

    numericFields.forEach(field => {
      const td = document.createElement('td');
      td.className = 'num-col';
      const input = document.createElement('input');
      input.type = 'text';
      input.value = row[field.key] || '';
      input.setAttribute('aria-label', field.label);
      input.addEventListener('input', (e) => {
        const value = e.target.value;
        
        // 验证输入
        if (value !== '' && !isValidNumber(value)) {
          input.classList.add('invalid');
          showToast(`${field.label}列请输入有效数字`, 'warning');
        } else {
          input.classList.remove('invalid');
          state.rows[index][field.key] = value;
          recalculate();
        }
      });
      td.appendChild(input);
      tr.appendChild(td);
    });

    // 右侧三个合计列（显示计算值：本行数值 × 入库）
    const sumFields = [
      { key: 'ash_internal', label: '内灰合计' },
      { key: 'volatile', label: '挥发合计' },
      { key: 'caking', label: '粘结合计' }
    ];
    
    sumFields.forEach(field => {
      const td = document.createElement('td');
      td.className = 'num-col sum-col';
      td.dataset.rowIndex = index;
      td.dataset.field = field.key;
      tr.appendChild(td);
    });

    elements.tableBody.appendChild(tr);
  });

  // 重新计算（包括每行的合计列）
  recalculate();
  updateSelectAllState();
}

/**
 * 重新计算合计和平均
 */
function recalculate() {
  const numericFields = [
    'inbound', 'ash_total', 'ash_internal', 'sulfur',
    'volatile', 'caking', 'middlings', 'gangue'
  ];

  // 计算各列合计
  const totals = {};
  numericFields.forEach(field => {
    const values = state.rows.map(row => row[field]);
    totals[field] = toInteger(safeSum(values));
  });

  // 计算每行的右侧三列（本行数值 × 入库）并累加到合计
  let ash_internal_sum_total = 0;
  let volatile_sum_total = 0;
  let caking_sum_total = 0;

  state.rows.forEach((row, index) => {
    const inbound = parseNumber(row.inbound);
    const ash_internal = parseNumber(row.ash_internal);
    const volatile = parseNumber(row.volatile);
    const caking = parseNumber(row.caking);

    // 计算每行的合计列值
    const ash_internal_sum = ash_internal * inbound;
    const volatile_sum = volatile * inbound;
    const caking_sum = caking * inbound;

    // 累加到总合计
    ash_internal_sum_total += ash_internal_sum;
    volatile_sum_total += volatile_sum;
    caking_sum_total += caking_sum;

    // 更新每行的合计列显示
    const ashCell = document.querySelector(`td[data-row-index="${index}"][data-field="ash_internal"]`);
    const volatileCell = document.querySelector(`td[data-row-index="${index}"][data-field="volatile"]`);
    const cakingCell = document.querySelector(`td[data-row-index="${index}"][data-field="caking"]`);

    if (ashCell) ashCell.textContent = toFixed2(ash_internal_sum);
    if (volatileCell) volatileCell.textContent = toFixed2(volatile_sum);
    if (cakingCell) cakingCell.textContent = toFixed2(caking_sum);
  });

  // 右侧三列的表尾合计
  totals.ash_internal_sum = toInteger(ash_internal_sum_total);
  totals.volatile_sum = toInteger(volatile_sum_total);
  totals.caking_sum = toInteger(caking_sum_total);

  // 更新合计行
  numericFields.forEach(field => {
    const el = document.getElementById(`total-${field}`);
    if (el) el.textContent = totals[field];
  });
  document.getElementById('total-ash_internal_sum').textContent = totals.ash_internal_sum;
  document.getElementById('total-volatile_sum').textContent = totals.volatile_sum;
  document.getElementById('total-caking_sum').textContent = totals.caking_sum;

  // 计算平均（仅内灰、挥发、粘结）
  const avgFields = ['ash_internal', 'volatile', 'caking'];
  const averages = {};
  
  avgFields.forEach(field => {
    const values = state.rows.map(row => row[field]);
    const avg = safeAvg(values);
    averages[field] = toFixed2(avg);
  });

  // 更新平均行（左侧列）
  document.getElementById('avg-ash_internal').textContent = averages.ash_internal;
  document.getElementById('avg-volatile').textContent = averages.volatile;
  document.getElementById('avg-caking').textContent = averages.caking;

  // 计算右侧三列合计的平均值（合计值 / 入库总和）
  const inboundSum = totals.inbound;
  
  if (inboundSum > 0) {
    averages.ash_internal_sum = toFixed2(totals.ash_internal_sum / inboundSum);
    averages.volatile_sum = toFixed2(totals.volatile_sum / inboundSum);
    averages.caking_sum = toFixed2(totals.caking_sum / inboundSum);
  } else {
    averages.ash_internal_sum = '0.00';
    averages.volatile_sum = '0.00';
    averages.caking_sum = '0.00';
  }

  // 更新平均行（右侧合计列）
  document.getElementById('avg-ash_internal_sum').textContent = averages.ash_internal_sum;
  document.getElementById('avg-volatile_sum').textContent = averages.volatile_sum;
  document.getElementById('avg-caking_sum').textContent = averages.caking_sum;

  // 保存到全局以便导出使用
  state.totals = totals;
  state.averages = averages;
}

/**
 * 新增一行
 */
function addRow() {
  const newIndex = state.rows.length + 1;
  state.rows.push(createEmptyRow(newIndex));
  renderTable();
  showToast('已添加新行', 'success');
}

/**
 * 删除选中行
 */
function removeSelected() {
  const selectedCount = state.rows.filter(row => row.selected).length;
  
  if (selectedCount === 0) {
    showToast('请先选择要删除的行', 'warning');
    return;
  }

  if (!confirm(`确定要删除选中的 ${selectedCount} 行吗？`)) {
    return;
  }

  // 过滤掉选中的行
  state.rows = state.rows.filter(row => !row.selected);

  // 重新编号
  state.rows.forEach((row, index) => {
    row.index = index + 1;
  });

  renderTable();
  showToast(`已删除 ${selectedCount} 行`, 'success');
}

/**
 * 导出 CSV
 */
function exportData() {
  if (state.rows.length === 0) {
    showToast('没有数据可导出', 'warning');
    return;
  }

  try {
    const csvContent = buildCSV(state.rows, state.totals || {}, state.averages || {});
    const filename = `原煤指标_${new Date().toISOString().slice(0, 10)}.csv`;
    downloadCSV(csvContent, filename);
    showToast('导出成功', 'success');
  } catch (e) {
    console.error('导出失败:', e);
    showToast('导出失败', 'error');
  }
}

/**
 * 重置数据
 */
function resetData() {
  if (!confirm('确定要重置当前工作簿的所有数据吗？此操作不可恢复。')) {
    return;
  }

  try {
    initEmptyRows();
    renderTable();
    showToast('已重置为初始状态', 'success');
  } catch (e) {
    console.error('重置失败:', e);
    showToast('重置失败', 'error');
  }
}

/**
 * 全选/取消全选
 */
function toggleSelectAll() {
  const checked = elements.selectAll.checked;
  state.rows.forEach(row => {
    row.selected = checked;
  });
  renderTable();
}

/**
 * 更新全选复选框状态
 */
function updateSelectAllState() {
  const allSelected = state.rows.length > 0 && state.rows.every(row => row.selected);
  const someSelected = state.rows.some(row => row.selected);
  
  elements.selectAll.checked = allSelected;
  elements.selectAll.indeterminate = someSelected && !allSelected;
}

/**
 * 显示提示信息
 */
function showToast(message, type = 'info') {
  elements.toast.textContent = message;
  elements.toast.className = `toast ${type} show`;
  
  setTimeout(() => {
    elements.toast.classList.remove('show');
  }, 3000);
}

// 页面加载完成后初始化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', init);
} else {
  init();
}
