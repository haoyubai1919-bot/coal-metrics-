/**
 * CSV 导出工具
 */

/**
 * 转义 CSV 单元格内容
 */
function escapeCSVCell(value) {
  if (value === null || value === undefined) {
    return '';
  }
  const str = String(value);
  // 如果包含逗号、引号或换行符，需要用引号包裹并转义引号
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * 构建 CSV 字符串
 */
function buildCSV(rows, totals, averages) {
  const headers = [
    '序号', '日期', '单位', '入库', '全灰', '内灰', '硫', 
    '挥发', '粘结', '中煤', '矸石', '内灰合计', '挥发合计', '粘结合计'
  ];

  // 添加 UTF-8 BOM 以便 Excel 正确识别中文
  let csv = '\uFEFF';
  
  // 表头
  csv += headers.map(escapeCSVCell).join(',') + '\n';

  // 数据行
  rows.forEach(row => {
    const inbound = parseFloat(row.inbound) || 0;
    const ash_internal = parseFloat(row.ash_internal) || 0;
    const volatile = parseFloat(row.volatile) || 0;
    const caking = parseFloat(row.caking) || 0;

    // 计算每行的合计列
    const ash_internal_sum = (ash_internal * inbound).toFixed(2);
    const volatile_sum = (volatile * inbound).toFixed(2);
    const caking_sum = (caking * inbound).toFixed(2);

    const line = [
      row.index,
      row.date || '',
      row.unit || '',
      row.inbound || '',
      row.ash_total || '',
      row.ash_internal || '',
      row.sulfur || '',
      row.volatile || '',
      row.caking || '',
      row.middlings || '',
      row.gangue || '',
      ash_internal_sum,
      volatile_sum,
      caking_sum
    ];
    csv += line.map(escapeCSVCell).join(',') + '\n';
  });

  // 合计行
  const totalLine = [
    '合计',
    '',
    '',
    totals.inbound || 0,
    totals.ash_total || 0,
    totals.ash_internal || 0,
    totals.sulfur || 0,
    totals.volatile || 0,
    totals.caking || 0,
    totals.middlings || 0,
    totals.gangue || 0,
    totals.ash_internal_sum || 0,
    totals.volatile_sum || 0,
    totals.caking_sum || 0
  ];
  csv += totalLine.map(escapeCSVCell).join(',') + '\n';

  // 平均行
  const avgLine = [
    '平均',
    '—',
    '—',
    '—',
    '—',
    averages.ash_internal || '0.00',
    '—',
    averages.volatile || '0.00',
    averages.caking || '0.00',
    '—',
    '—',
    averages.ash_internal_sum || '0.00',
    averages.volatile_sum || '0.00',
    averages.caking_sum || '0.00'
  ];
  csv += avgLine.map(escapeCSVCell).join(',') + '\n';

  return csv;
}

/**
 * 触发 CSV 文件下载
 */
function downloadCSV(csvContent, filename) {
  const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
  const link = document.createElement('a');
  
  if (navigator.msSaveBlob) {
    // IE 10+
    navigator.msSaveBlob(blob, filename);
  } else {
    const url = URL.createObjectURL(blob);
    link.href = url;
    link.download = filename;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  }
}
