import React from 'react';
import ReactDOM from 'react-dom/client';
import App from './App';
import './index.css';
import './styles/animations.css';
import { useJarvisStore } from './stores/jarvisStore';
import { applyResolvedTheme, resolveTheme } from './lib/theme';

// Stamp the persisted theme onto <html> before first paint so there's no flash
// of the wrong appearance. The store rehydrates from localStorage synchronously.
applyResolvedTheme(resolveTheme(useJarvisStore.getState().theme));

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
