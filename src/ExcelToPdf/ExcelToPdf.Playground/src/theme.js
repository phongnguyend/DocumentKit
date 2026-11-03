export const themeKey = 'playground-theme';
export function readTheme() {
  try {
    const value = (localStorage.getItem(themeKey) ?? localStorage.getItem('pdfobjectkit-theme'));
    return value === 'light' || value === 'dark' ? value : 'system';
  } catch {
    return 'system';
  }
}
export function applyTheme(theme) {
  document.documentElement.dataset.theme = theme === 'system' ? (matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light') : theme;
}
export function saveTheme(theme) {
  try {
    localStorage.setItem(themeKey, theme);
  } catch {
    // Theme switching still works when browser storage is unavailable.
  }
  applyTheme(theme);
  window.dispatchEvent(new CustomEvent('playground-theme-change', { detail: theme }));
}
