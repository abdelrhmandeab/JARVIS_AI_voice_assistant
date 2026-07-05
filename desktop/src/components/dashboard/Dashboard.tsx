import type { ReactNode } from 'react';
import type { UICommand, FeatureFlags } from '../../protocol';
import { backToOverlay, closeApp } from '../../lib/app';
import { GradientBackground } from '../GradientBackground';
import { PromptInput } from '../overlay/PromptInput';
import { useJarvisStore, type AvatarDirection, type UiLanguage } from '../../stores/jarvisStore';
import type { ThemePreference } from '../../lib/theme';
import { Chip, PanelLabel } from '../ui/Chip';
import { Select, type SelectOption } from './Select';
import { Toggle } from './Toggle';

interface DashboardProps {
  send: (cmd: UICommand) => void;
}

type DashboardOption<T extends string> = {
  label: string;
  value: T;
};

const avatarOptions: Array<DashboardOption<AvatarDirection>> = [
  { label: 'Aurora', value: 'aurora' },
  { label: 'Glyph', value: 'glyph' },
  { label: 'Glass AI', value: 'glassai' },
  { label: 'Companion', value: 'companion' },
];

const languageOptions: Array<DashboardOption<UiLanguage>> = [
  { label: 'English', value: 'en' },
  { label: 'Arabic', value: 'ar' },
  { label: 'Auto', value: 'auto' },
];

const themeOptions: Array<DashboardOption<ThemePreference>> = [
  { label: 'Dark', value: 'dark' },
  { label: 'Light', value: 'light' },
  { label: 'Auto', value: 'auto' },
];

const personaOptions: SelectOption[] = [
  { label: 'Friendly', value: 'friendly' },
  { label: 'Formal', value: 'formal' },
  { label: 'Casual', value: 'casual' },
  { label: 'Professional', value: 'professional' },
  { label: 'Brief', value: 'brief' },
];

const modelOptions: SelectOption[] = [
  { label: 'Auto', value: 'auto' },
  { label: 'qwen3:0.6b', value: 'qwen3:0.6b' },
  { label: 'qwen3:1.7b', value: 'qwen3:1.7b' },
  { label: 'qwen3:4b', value: 'qwen3:4b' },
  { label: 'qwen3:8b', value: 'qwen3:8b' },
];

// NUMERIC_PARSING_ENABLED is intentionally omitted: the engine declares it in
// FEATURE_FLAGS but nothing in core/nlp/os_control reads it, so a toggle for
// it would be a dead control.
const featureFlagLabels: Partial<Record<keyof FeatureFlags, string>> = {
  AUTO_APP_DISCOVERY_ENABLED: 'Auto app discovery',
  MEDIA_DIRECT_DISPATCH_ENABLED: 'Media direct dispatch',
  SYSTEM_VOLUME_CONTROL: 'System volume control',
};

function ChipGroup<T extends string>({
  value,
  options,
  onChange,
}: {
  value: T;
  options: Array<DashboardOption<T>>;
  onChange: (value: T) => void;
}) {
  return (
    <div className="flex flex-wrap gap-2">
      {options.map((option) => (
        <Chip
          key={option.value}
          active={option.value === value}
          onClick={() => onChange(option.value)}
          className="text-[11px]"
        >
          {option.label}
        </Chip>
      ))}
    </div>
  );
}

function Section({
  title,
  className = '',
  children,
}: {
  title: string;
  className?: string;
  children: ReactNode;
}) {
  return (
    <section
      className={`rounded-md border border-black/[0.08] bg-white/70 p-3 text-slate-800 shadow-2xl shadow-black/10 backdrop-blur dark:border-white/10 dark:bg-black/70 dark:text-white dark:shadow-black/35 ${className}`}
    >
      <PanelLabel>{title}</PanelLabel>
      <div className="grid gap-3">{children}</div>
    </section>
  );
}

export function Dashboard({ send }: DashboardProps) {
  const config = useJarvisStore((state) => state.config);
  const avatarDirection = useJarvisStore((state) => state.avatarDirection);
  const uiLanguage = useJarvisStore((state) => state.uiLanguage);
  const muted = useJarvisStore((state) => state.muted);
  const textPromptEnabled = useJarvisStore((state) => state.textPromptEnabled);
  const theme = useJarvisStore((state) => state.theme);
  const connectionStatus = useJarvisStore((state) => state.connectionStatus);
  const setAvatarDirection = useJarvisStore((state) => state.setAvatarDirection);
  const setUiLanguage = useJarvisStore((state) => state.setUiLanguage);
  const setMuted = useJarvisStore((state) => state.setMuted);
  const setTextPromptEnabled = useJarvisStore((state) => state.setTextPromptEnabled);
  const setTheme = useJarvisStore((state) => state.setTheme);
  const setFeatureFlagLocal = useJarvisStore((state) => state.setFeatureFlagLocal);
  const setConfigValueLocal = useJarvisStore((state) => state.setConfigValueLocal);

  const hasConfig = config !== null;

  const handleLanguageChange = (language: UiLanguage) => {
    setUiLanguage(language);
    send({ type: 'setting_update', key: 'JARVIS_STT_LANGUAGE_HINT', value: language });
  };

  const handleMutedChange = (nextMuted: boolean) => {
    setMuted(nextMuted);
    send({ type: 'mute_toggle', muted: nextMuted });
  };

  return (
    <div className="frameless-scroll relative h-screen overflow-y-auto px-4 py-6 text-slate-800 dark:text-white sm:px-6 lg:px-8">
      {/* Frosted base carrying the same pink / blue / amber gradient glow in both
          themes — a pale surface in light mode, near-black in dark. */}
      <div
        aria-hidden="true"
        className="pointer-events-none fixed inset-0 z-0 bg-[#E9EDF6]/92 backdrop-blur-xl dark:bg-[#0A0A0F]/95"
      >
        <div className="absolute inset-0 opacity-40 dark:opacity-35">
          <GradientBackground
            containerClassName="h-full w-full"
            gradientColors={['rgb(255, 100, 150)', 'rgb(100, 150, 255)', 'rgb(255, 200, 100)']}
          />
        </div>
      </div>
      <div className="relative z-10 mx-auto grid w-full max-w-5xl gap-5 p-4 text-sm sm:p-5">
        {/* Frameless window: this header doubles as the drag handle. */}
        <header
          data-tauri-drag-region
          className="flex flex-col gap-3 border-b border-black/10 pb-4 dark:border-white/10 sm:flex-row sm:items-center sm:justify-between"
        >
          <div data-tauri-drag-region>
            <p className="text-xs font-semibold uppercase tracking-[0.18em] text-cyan-700/80 dark:text-cyan-50/70">
              Control Center
            </p>
            <h1 className="mt-1 text-2xl font-semibold tracking-normal text-slate-900 dark:text-white">
              <span className="font-jarvis tracking-wide">JARVIS</span>
            </h1>
          </div>

          <div className="flex flex-wrap gap-2">
            <Chip onClick={() => send({ type: 'config_request' })} className="h-9 px-3 text-sm font-medium">
              Refresh
            </Chip>
            <Chip
              active
              onClick={() => {
                void backToOverlay();
              }}
              className="h-9 px-3 text-sm font-medium"
            >
              Hide
            </Chip>
            <button
              type="button"
              aria-label="Close Jarvis"
              title="Close Jarvis"
              onClick={() => {
                void closeApp().catch((error: unknown) => console.error('Failed to close app.', error));
              }}
              className="grid h-9 w-9 place-items-center rounded border border-red-400/30 bg-red-400/15 text-red-600 transition-opacity hover:opacity-90 dark:border-red-300/25 dark:bg-red-400/12 dark:text-red-100"
            >
              <svg width="16" height="16" viewBox="0 0 16 16" fill="none" aria-hidden="true">
                <path d="M4 4 L12 12 M12 4 L4 12" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" />
              </svg>
            </button>
          </div>
        </header>

        {!hasConfig ? (
          <div className="rounded-md border border-cyan-500/25 bg-cyan-400/10 p-3 text-[12px] text-cyan-800 dark:border-cyan-200/20 dark:bg-cyan-200/10 dark:text-cyan-50/80">
            Engine config has not arrived yet. Use Refresh to request the current values from the bridge.
          </div>
        ) : null}

        <div className="grid gap-3 lg:grid-cols-2">
          <Section title="Text Prompt" className="lg:col-span-2">
            <Toggle label="Enable text prompt" checked={textPromptEnabled} onChange={setTextPromptEnabled} />
            {textPromptEnabled ? <PromptInput send={send} /> : null}
          </Section>

          <Section title="Avatar">
            <ChipGroup value={avatarDirection} options={avatarOptions} onChange={setAvatarDirection} />
          </Section>

          <Section title="Voice Persona">
            <Select
              label="Persona"
              value={config?.persona ?? 'friendly'}
              options={personaOptions}
              disabled={!hasConfig}
              onChange={(persona) => {
                setConfigValueLocal('persona', persona);
                send({ type: 'setting_update', key: 'JARVIS_PERSONA', value: persona });
              }}
            />
          </Section>

          <Section title="Language">
            <ChipGroup value={uiLanguage} options={languageOptions} onChange={handleLanguageChange} />
          </Section>

          <Section title="Appearance">
            <ChipGroup value={theme} options={themeOptions} onChange={setTheme} />
            <p className="text-[11px] text-slate-500 dark:text-white/45">
              Auto follows your system light/dark setting.
            </p>
          </Section>

          <Section title="Model">
            <Select
              label="LLM model"
              value={config?.model ?? 'auto'}
              options={modelOptions}
              disabled={!hasConfig}
              onChange={(model) => {
                setConfigValueLocal('model', model);
                send({ type: 'setting_update', key: 'JARVIS_LLM_MODEL', value: model });
              }}
            />
          </Section>

          <Section title="Feature Flags">
            {(
              Object.entries(featureFlagLabels) as Array<[keyof FeatureFlags, string]>
            ).map(([flag, label]) => (
              <Toggle
                key={flag}
                label={label}
                checked={config?.feature_flags[flag] ?? false}
                disabled={!hasConfig}
                onChange={(enabled) => {
                  setFeatureFlagLocal(flag, enabled);
                  send({ type: 'feature_flag', flag, enabled });
                }}
              />
            ))}
          </Section>

          <Section title="Audio">
            <Toggle label="Mute microphone and speech" checked={muted} onChange={handleMutedChange} />
          </Section>

          <Section title="Status">
            <dl className="grid gap-3 text-sm">
              <div className="flex items-center justify-between gap-4">
                <dt className="text-slate-500 dark:text-white/58">Connection</dt>
                <dd className="rounded border border-black/10 bg-black/[0.04] px-2 py-1 font-medium capitalize text-slate-700 dark:border-white/10 dark:bg-white/5 dark:text-white/80">
                  {connectionStatus}
                </dd>
              </div>
              <div className="flex items-center justify-between gap-4">
                <dt className="text-slate-500 dark:text-white/58">Current model</dt>
                <dd className="rounded border border-black/10 bg-black/[0.04] px-2 py-1 font-medium text-slate-700 dark:border-white/10 dark:bg-white/5 dark:text-white/80">
                  {config?.model ?? 'Unknown'}
                </dd>
              </div>
            </dl>
          </Section>
        </div>
      </div>
    </div>
  );
}
