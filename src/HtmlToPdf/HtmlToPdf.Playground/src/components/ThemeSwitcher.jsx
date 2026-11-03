import { useEffect, useState } from 'react';
function Icon({ children }) {
  return <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">{children}</svg>;
}
function Sun() {
  return <Icon><circle cx="12" cy="12" r="4" /><path d="M12 2v2m0 16v2M2 12h2m16 0h2M4.93 4.93l1.42 1.42m11.3 11.3 1.42 1.42M4.93 19.07l1.42-1.42m11.3-11.3 1.42-1.42" /></Icon>;
}
function Moon() {
  return <Icon><path d="M20.985 12.486a9 9 0 1 1-9.473-9.472c.405-.022.617.46.402.803a6.25 6.25 0 0 0 8.268 8.268c.344-.215.825-.004.803.401" /></Icon>;
}
function Monitor() {
  return <Icon><rect x="2" y="3" width="20" height="14" rx="2" /><path d="M8 21h8m-4-4v4" /></Icon>;
}
import { applyTheme, readTheme, saveTheme, themeKey } from '../theme';

export function ThemeSwitcher() {
  const [theme, setTheme] = useState(readTheme);
  useEffect(() => {
    const media = matchMedia('(prefers-color-scheme: dark)');
    const update = () => applyTheme(theme);
    const sync = (event) => {
      if (event.key === themeKey || event.key === 'pdfobjectkit-theme' || event.key === null) {
        setTheme(readTheme());
      }
    };
    const localSync = (event) => setTheme(event.detail);
    window.addEventListener('playground-theme-change', localSync);
    update();
    media.addEventListener('change', update);
    window.addEventListener('storage', sync);
    return () => {
      media.removeEventListener('change', update);
      window.removeEventListener('playground-theme-change', localSync);
      window.removeEventListener('storage', sync);
    };
  }, [theme]);
  return <div className="theme-switcher" role="group" aria-label="Color theme">
    {([{ value: 'light', label: 'Light theme', Icon: Sun }, { value: 'dark', label: 'Dark theme', Icon: Moon }, { value: 'system', label: 'System theme', Icon: Monitor }]).map(({ value, label, Icon }) => <button key={value} type="button" title={label} aria-label={label} aria-pressed={theme === value} onClick={() => {
      setTheme(value);
      saveTheme(value);
    }}><Icon size={17} aria-hidden="true" /></button>)}
  </div>;
}
