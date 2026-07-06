import { useEffect, useState } from 'react';
import { useJarvisStore } from '../stores/jarvisStore';

export type ThemePreference = 'dark' | 'light' | 'auto';
export type ResolvedTheme = 'dark' | 'light';

function systemTheme(): ResolvedTheme {
  if (typeof window === 'undefined' || !window.matchMedia) return 'dark';
  return window.matchMedia('(prefers-color-scheme: light)').matches ? 'light' : 'dark';
}

/** Resolves the user's theme preference ('auto' follows the OS setting) to an actual light/dark value. */
export function useResolvedTheme(): ResolvedTheme {
  const theme = useJarvisStore((state) => state.theme);
  const [resolved, setResolved] = useState<ResolvedTheme>(() => (theme === 'auto' ? systemTheme() : theme));

  useEffect(() => {
    if (theme !== 'auto') {
      setResolved(theme);
      return;
    }

    setResolved(systemTheme());
    if (typeof window === 'undefined' || !window.matchMedia) return;

    const query = window.matchMedia('(prefers-color-scheme: light)');
    const onChange = () => setResolved(systemTheme());
    query.addEventListener('change', onChange);
    return () => query.removeEventListener('change', onChange);
  }, [theme]);

  return resolved;
}

/** Stamps the resolved theme onto <html data-theme="..."> so CSS (the
 * @custom-variant dark selector in index.css, and any [data-theme] rules)
 * reacts to the user's in-app choice instead of only prefers-color-scheme. */
export function useApplyTheme(): void {
  const resolved = useResolvedTheme();

  useEffect(() => {
    if (typeof document === 'undefined') return;
    document.documentElement.setAttribute('data-theme', resolved);
  }, [resolved]);
}
