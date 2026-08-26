const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('invoiceApi', {
  bootstrap: () => ipcRenderer.invoke('bootstrap'),
  scanDirectory: () => ipcRenderer.invoke('scan-directory'),
  chooseDirectory: () => ipcRenderer.invoke('choose-directory'),
  chooseFiles: () => ipcRenderer.invoke('choose-files'),
  importFiles: (paths, projectId) => ipcRenderer.invoke('import-files', paths, projectId),
  updateInvoice: (id, patch) => ipcRenderer.invoke('update-invoice', id, patch),
  reparseInvoice: (id) => ipcRenderer.invoke('reparse-invoice', id),
  deleteInvoice: (id, trash) => ipcRenderer.invoke('delete-invoice', id, trash),
  createProject: (name, note) => ipcRenderer.invoke('create-project', name, note),
  updateProject: (id, patch) => ipcRenderer.invoke('update-project', id, patch),
  archiveProject: (id, archived) => ipcRenderer.invoke('archive-project', id, archived),
  deleteProject: (id) => ipcRenderer.invoke('delete-project', id),
  linkInvoice: (projectId, invoiceId, primary) => ipcRenderer.invoke('link-invoice', projectId, invoiceId, primary),
  unlinkInvoice: (projectId, invoiceId) => ipcRenderer.invoke('unlink-invoice', projectId, invoiceId),
  reconcileFiles: () => ipcRenderer.invoke('reconcile-files'),
  exportProject: (projectId) => ipcRenderer.invoke('export-project', projectId),
  exportAllInvoices: () => ipcRenderer.invoke('export-all-invoices'),
  openInvoice: (path) => ipcRenderer.invoke('open-invoice', path),
  addStatus: (name) => ipcRenderer.invoke('add-status', name),
  addTag: (name) => ipcRenderer.invoke('add-tag', name),
  addCategory: (name) => ipcRenderer.invoke('add-category', name)
  ,setTheme: (theme) => ipcRenderer.invoke('set-theme', theme)
});
