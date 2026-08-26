const test = require('node:test');
const assert = require('node:assert/strict');
const { clean, money, invoiceFileName } = require('../electron/invoice-rules.cjs');

test('文件名清理 Windows 非法字符并保留结构', () => {
  const name = invoiceFileName({ total_amount: 19.9, product_name: '螺丝/刀:套装', status_name: '未报销' }, { name: '项目*一' }, 3);
  assert.equal(name, '项目_一-003-19.90-螺丝_刀_套装-未报销.pdf');
});

test('金额保留两位小数', () => {
  assert.equal(money('67.805'), 67.81);
  assert.equal(clean('  待 报销  '), '待 报销');
});
