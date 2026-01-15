/**
 * 数学工具函数
 */

/**
 * 解析数字 - 将输入转为数值，空/非法返回 0
 */
function parseNumber(value) {
  if (value === null || value === undefined || value === '') {
    return 0;
  }
  const num = parseFloat(value);
  return isNaN(num) ? 0 : num;
}

/**
 * 验证是否为有效数字
 */
function isValidNumber(value) {
  if (value === null || value === undefined || value === '') {
    return true; // 空值视为有效（按 0 处理）
  }
  const str = String(value).trim();
  if (str === '') return true;
  // 允许数字、小数点、负号
  return /^-?\d*\.?\d+$/.test(str);
}

/**
 * 安全求和 - 将空字符串和非数值视为 0
 */
function safeSum(values) {
  return values.reduce((sum, val) => {
    return sum + parseNumber(val);
  }, 0);
}

/**
 * 安全求平均 - 只计算有效数值的平均
 */
function safeAvg(values) {
  const validValues = values.filter(val => {
    if (val === null || val === undefined || val === '') {
      return false;
    }
    const num = parseFloat(val);
    return !isNaN(num);
  }).map(val => parseFloat(val));

  if (validValues.length === 0) {
    return 0;
  }

  const sum = validValues.reduce((acc, val) => acc + val, 0);
  return sum / validValues.length;
}

/**
 * 格式化为两位小数
 */
function toFixed2(num) {
  return Number(num).toFixed(2);
}

/**
 * 格式化为整数
 */
function toInteger(num) {
  return Math.round(Number(num));
}
