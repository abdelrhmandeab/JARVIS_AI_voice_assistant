import { useEffect, useState, type ReactNode } from 'react';
import type { UICommand, FeatureFlags } from '../../protocol';
import { backToOverlay, closeApp } from '../../lib/app';
import { GradientBackground } from '../GradientBackground';
import { PromptInput } from '../overlay/PromptInput';
import {
  useJarvisStore,
  type AvatarDirection,
  type ConnectionStatus,
  type NotificationTone,
  type UiLanguage,
  type VoiceGender,
} from '../../stores/jarvisStore';
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
  // "Jarvis" is the Companion avatar (the default); value stays 'companion' so
  // existing persisted selections keep working.
  { label: 'Jarvis', value: 'companion' },
  { label: 'Aurora', value: 'aurora' },
  { label: 'Glyph', value: 'glyph' },
  { label: 'Glass AI', value: 'glassai' },
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

const voiceOptions: Array<DashboardOption<VoiceGender>> = [
  { label: 'Male', value: 'male' },
  { label: 'Female', value: 'female' },
];

// Mirrors core.persona.PERSONA_PROFILES exactly so every engine persona is
// selectable and config.persona always matches an option.
const personaOptions: SelectOption[] = [
  { label: 'Assistant', value: 'assistant' },
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

// Only the flags the engine actually reads are exposed (NUMERIC_PARSING_ENABLED
// is declared but never consumed, so it stays out of the UI). Each description
// reflects what toggling the flag does in os_control.
const featureFlags: Array<{ flag: keyof FeatureFlags; label: string; description: string }> = [
  {
    flag: 'AUTO_APP_DISCOVERY_ENABLED',
    label: 'Auto app discovery',
    description:
      'Scans your installed applications so Jarvis can find and open them by name. When off, only apps in the built-in catalog can be launched.',
  },
  {
    flag: 'MEDIA_DIRECT_DISPATCH_ENABLED',
    label: 'Media direct dispatch',
    description:
      'Sends play/pause, next, and previous straight to the OS as native media keys for instant control. When off, those media commands are skipped.',
  },
  {
    flag: 'SYSTEM_VOLUME_CONTROL',
    label: 'System volume control',
    description:
      'Lets Jarvis raise, lower, mute, and set your system volume by voice. When off, volume commands are ignored.',
  },
];

// Tone → banner styles for the dismissible notification area.
const NOTICE_TONES: Record<NotificationTone, string> = {
  info: 'border-cyan-500/25 bg-cyan-400/10 text-cyan-800 dark:border-cyan-200/20 dark:bg-cyan-200/10 dark:text-cyan-50/85',
  error: 'border-red-500/30 bg-red-400/12 text-red-700 dark:border-red-300/25 dark:bg-red-400/12 dark:text-red-100',
  success:
    'border-emerald-500/30 bg-emerald-400/12 text-emerald-700 dark:border-emerald-300/25 dark:bg-emerald-400/12 dark:text-emerald-100',
};

// Connection state → user-facing online/offline presentation.
const CONNECTION_META: Record<ConnectionStatus, { label: string; dot: string; text: string }> = {
  connected: { label: 'Online', dot: 'bg-emerald-500', text: 'text-emerald-700 dark:text-emerald-300' },
  connecting: { label: 'Connecting', dot: 'bg-amber-500', text: 'text-amber-700 dark:text-amber-300' },
  disconnected: { label: 'Offline', dot: 'bg-red-500', text: 'text-red-600 dark:text-red-300' },
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
      className={`rounded-md border border-black/[0.08] bg-white/55 p-3 text-slate-800 shadow-2xl shadow-black/10 backdrop-blur-md dark:border-white/10 dark:bg-black/50 dark:text-white dark:shadow-black/35 ${className}`}
    >
      <PanelLabel>{title}</PanelLabel>
      <div className="grid gap-3">{children}</div>
    </section>
  );
}

// A single dismissible notification row with a small close button.
function Notice({
  tone,
  onDismiss,
  children,
}: {
  tone: NotificationTone;
  onDismiss: () => void;
  children: ReactNode;
}) {
  return (
    <div className={`flex items-start gap-3 rounded-md border p-3 text-[12px] ${NOTICE_TONES[tone]}`}>
      <span className="min-w-0 flex-1 break-words">{children}</span>
      <button
        type="button"
        onClick={onDismiss}
        aria-label="Dismiss notification"
        title="Dismiss"
        className="grid h-5 w-5 shrink-0 place-items-center rounded opacity-70 transition-opacity hover:opacity-100"
      >
        <svg width="12" height="12" viewBox="0 0 16 16" fill="none" aria-hidden="true">
          <path d="M4 4 L12 12 M12 4 L4 12" stroke="currentColor" strokeWidth="1.7" strokeLinecap="round" />
        </svg>
      </button>
    </div>
  );
}

export function Dashboard({ send }: DashboardProps) {
  const config = useJarvisStore((state) => state.config);
  const avatarDirection = useJarvisStore((state) => state.avatarDirection);
  const uiLanguage = useJarvisStore((state) => state.uiLanguage);
  const muted = useJarvisStore((state) => state.muted);
  const textPromptEnabled = useJarvisStore((state) => state.textPromptEnabled);
  const theme = useJarvisStore((state) => state.theme);
  const voiceGender = useJarvisStore((state) => state.voiceGender);
  const connectionStatus = useJarvisStore((state) => state.connectionStatus);
  const notifications = useJarvisStore((state) => state.notifications);
  const setAvatarDirection = useJarvisStore((state) => state.setAvatarDirection);
  const setUiLanguage = useJarvisStore((state) => state.setUiLanguage);
  const setMuted = useJarvisStore((state) => state.setMuted);
  const setTextPromptEnabled = useJarvisStore((state) => state.setTextPromptEnabled);
  const setTheme = useJarvisStore((state) => state.setTheme);
  const setVoiceGender = useJarvisStore((state) => state.setVoiceGender);
  const setFeatureFlagLocal = useJarvisStore((state) => state.setFeatureFlagLocal);
  const setConfigValueLocal = useJarvisStore((state) => state.setConfigValueLocal);
  const dismissNotification = useJarvisStore((state) => state.dismissNotification);

  const hasConfig = config !== null;

  // Locally dismissible "config not arrived" hint, shown as a notification.
  const [configHintDismissed, setConfigHintDismissed] = useState(false);

  // Mock-only model picker: lets the user browse the available models without
  // reconfiguring the engine (no setting_update is sent). It tracks the engine's
  // reported model so it stays believable, but changing it is purely cosmetic.
  const [previewModel, setPreviewModel] = useState<string>(config?.model ?? 'auto');
  useEffect(() => {
    if (config?.model) setPreviewModel(config.model);
  }, [config?.model]);

  const connection = CONNECTION_META[connectionStatus];

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
        className="pointer-events-none fixed inset-0 z-0 bg-[#E9EDF6]/62 backdrop-blur-xl dark:bg-[#0A0A0F]/64"
      >
        <div className="absolute inset-0 opacity-40 dark:opacity-35">
          <GradientBackground
            containerClassName="h-full w-full"
            gradientColors={['rgb(255, 100, 150)', 'rgb(100, 150, 255)', 'rgb(255, 200, 100)']}
          />
        </div>
      </div>
      {/* Body content defaults to semibold; headline elements (h1, the eyebrow,
          and each Section's PanelLabel) already set their own font-semibold, so
          they're unaffected. */}
      <div className="relative z-10 mx-auto grid w-full max-w-5xl gap-5 p-4 text-sm font-semibold sm:p-5">
        {/* Frameless window: this header doubles as the drag handle. */}
        <header
          data-tauri-drag-region
          className="flex flex-col gap-3 border-b border-black/10 pb-4 dark:border-white/10 sm:flex-row sm:items-center sm:justify-between"
        >
          <div data-tauri-drag-region>
            <p className="font-jarvis text-xs font-semibold uppercase tracking-[0.18em] text-cyan-700/80 dark:text-cyan-50/70">
              Control Center
            </p>
            <h1 className="mt-1 text-2xl font-semibold tracking-normal text-slate-900 dark:text-white">
              <span className="font-jarvis tracking-wide">JARVIS</span>
            </h1>
          </div>

          <div className="flex flex-wrap gap-2">
            <Chip onClick={() => send({ type: 'config_request' })} className="h-9 px-3 text-sm font-semibold">
              Refresh
            </Chip>
            <Chip
              active
              onClick={() => {
                void backToOverlay();
              }}
              className="h-9 px-3 text-sm font-semibold"
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

        {notifications.length > 0 || (!hasConfig && !configHintDismissed) ? (
          <div className="grid gap-2">
            {!hasConfig && !configHintDismissed ? (
              <Notice tone="info" onDismiss={() => setConfigHintDismissed(true)}>
                Engine config has not arrived yet. Use Refresh to request the current values from the bridge.
              </Notice>
            ) : null}
            {notifications.map((item) => (
              <Notice key={item.id} tone={item.tone} onDismiss={() => dismissNotification(item.id)}>
                {item.message}
              </Notice>
            ))}
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
            <div className="grid gap-2">
              <span className="font-semibold text-slate-600 dark:text-white/72">Voice</span>
              {/* UI-only for now: remembers the preference; the engine will consume
                  it once TTS voice selection is implemented. */}
              <ChipGroup value={voiceGender} options={voiceOptions} onChange={setVoiceGender} />
              <p className="text-[11px] font-normal text-slate-500 dark:text-white/45">
                Male/female voice — engine support coming soon.
              </p>
            </div>
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
            <Select label="LLM model" value={previewModel} options={modelOptions} onChange={setPreviewModel} />
            <p className="text-[11px] text-slate-500 dark:text-white/45">
              Preview only — selecting a model here doesn't change the running engine yet.
            </p>
          </Section>

          <Section title="Features">
            {featureFlags.map(({ flag, label, description }) => (
              <div key={flag} className="grid gap-1.5">
                <Toggle
                  label={label}
                  checked={config?.feature_flags[flag] ?? false}
                  disabled={!hasConfig}
                  onChange={(enabled) => {
                    setFeatureFlagLocal(flag, enabled);
                    send({ type: 'feature_flag', flag, enabled });
                  }}
                />
                <p className="px-1 text-[11px] leading-relaxed text-slate-500 dark:text-white/45">{description}</p>
              </div>
            ))}
          </Section>

          <Section title="Audio">
            <Toggle label="Mute microphone and speech" checked={muted} onChange={handleMutedChange} />
          </Section>

          <Section title="Status">
            <dl className="grid gap-3 text-sm">
              <div className="flex items-center justify-between gap-4">
                <dt className="text-slate-500 dark:text-white/58">Connection</dt>
                <dd
                  className={`flex items-center gap-2 rounded border border-black/10 bg-black/[0.04] px-2 py-1 font-semibold dark:border-white/10 dark:bg-white/5 ${connection.text}`}
                >
                  <span className={`h-2 w-2 rounded-full ${connection.dot}`} aria-hidden="true" />
                  {connection.label}
                </dd>
              </div>
              <div className="flex items-center justify-between gap-4">
                <dt className="text-slate-500 dark:text-white/58">Current model</dt>
                <dd className="rounded border border-black/10 bg-black/[0.04] px-2 py-1 font-semibold text-slate-700 dark:border-white/10 dark:bg-white/5 dark:text-white/80">
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
