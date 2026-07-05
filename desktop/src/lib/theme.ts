import { useEffect, useState } from 'react';
import { useJarvisStore } from '../stores/jarvisStore';

// User-facing appearance preference. "auto" follows the OS setting.
export type ThemePreference = 'dark' | 'light' | 'auto';
// The concrete theme actually painted after resolving "auto".
export type ResolvedTheme = 'dark' | 'light';

const DARK_QUERY = '(prefers-color-scheme: dark)';

function darkMedia(): MediaQueryList | null {
  if (typeof window === 'undefined' || typeof window.matchMedia !== 'function') {
    return null;
  }
  return window.matchMedia(DARK_QUERY);
}

// Dark is this app's default identity, so treat an unknown OS preference as dark.
export function systemPrefersDark(): boolean {
  return darkMedia()?.matches ?? true;
}

export function resolveTheme(preference: ThemePreference): ResolvedTheme {
  if (preference === 'auto') {
    return systemPrefersDark() ? 'dark' : 'light';
  }
  return preference;
}

// Stamp the resolved theme onto <html>. Tailwind's `dark:` variant is redefined
// in index.css to key off this attribute, and `color-scheme` makes native
// controls (selects, scrollbars) match.
export function applyResolvedTheme(resolved: ResolvedTheme): void {
  if (typeof document === 'undefined') return;
  const root = document.documentElement;
  root.dataset.theme = resolved;
  root.style.colorScheme = resolved;
}

// Keep <html data-theme> in sync with the store preference, and — while on
// "auto" — with live OS theme changes. Mount once near the app root.
export function useThemeSync(): void {
  const preference = useJarvisStore((state) => state.theme);

  useEffect(() => {
    applyResolvedTheme(resolveTheme(preference));

    if (preference !== 'auto') return;
    const media = darkMedia();
    if (!media) return;

    const onChange = () => applyResolvedTheme(resolveTheme('auto'));
    media.addEventListener('change', onChange);
    return () => media.removeEventListener('change', onChange);
  }, [preference]);
}

// Read the concrete resolved theme for components that must branch inline styles
// (e.g. the canvas-like liquid-glass avatar) rather than use `dark:` classes.
export function useResolvedTheme(): ResolvedTheme {
  const preference = useJarvisStore((state) => state.theme);
  const [systemDark, setSystemDark] = useState(systemPrefersDark);

  useEffect(() => {
    if (preference !== 'auto') return;
    const media = darkMedia();
    if (!media) return;

    const onChange = () => setSystemDark(media.matches);
    setSystemDark(media.matches);
    media.addEventListener('change', onChange);
    return () => media.removeEventListener('change', onChange);
  }, [preference]);

  if (preference === 'auto') {
    return systemDark ? 'dark' : 'light';
  }
  return preference;
}
