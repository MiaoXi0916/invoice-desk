function clean(value, fallback = '') {
  return String(value || fallback).replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/\s+/g, ' ').trim();
}

function money(value) {
  return Number(Number(value || 0).toFixed(2));
}

function invoiceFileName(invoice, project, sequence) {
  return `${clean(project?.name, '待报销')}-${String(sequence || 1).padStart(3, '0')}-${money(invoice.total_amount).toFixed(2)}-${clean(invoice.product_name, '待人工补充')}-${clean(invoice.status_name, '未报销')}.pdf`;
}

module.exports = { clean, money, invoiceFileName };
