import { create } from 'zustand';
import { createJSONStorage, persist } from 'zustand/middleware';
import type { ConfigValues, DialogueState, EngineEvent, FeatureFlags, Language, PinResultStatus } from '../protocol';
import type { ThemePreference } from '../lib/theme';

export type ConnectionStatus = 'connecting' | 'connected' | 'disconnected';
export type AvatarDirection = 'aurora' | 'glyph' | 'glassai' | 'companion';
export type AppView = 'overlay' | 'dashboard';
export type UiLanguage = Language | 'auto';
export type VoiceGender = 'male' | 'female';

export interface PinRequiredState {
  description: string;
  attemptsRemaining: number;
  expiresInSeconds: number;
  receivedAt: number;
}

export interface PinResultState {
  status: PinResultStatus;
  message: string;
  attemptsRemaining: number;
  receivedAt: number;
}

export type NotificationTone = 'info' | 'error' | 'success';

export interface AppNotification {
  id: string;
  message: string;
  tone: NotificationTone;
}

let notificationCounter = 0;
function makeNotificationId(): string {
  notificationCounter += 1;
  return `n${Date.now()}-${notificationCounter}`;
}

// Append a user-facing notification, skipping an identical back-to-back
// duplicate and keeping only the most recent few so the banner can't grow
// without bound.
const MAX_NOTIFICATIONS = 5;
function appendNotification(
  list: AppNotification[],
  message: string,
  tone: NotificationTone,
): AppNotification[] {
  const text = message.trim();
  if (!text) return list;
  const last = list[list.length - 1];
  if (last && last.message === text && last.tone === tone) return list;
  return [...list, { id: makeNotificationId(), message: text, tone }].slice(-MAX_NOTIFICATIONS);
}

interface JarvisState {
  connectionStatus: ConnectionStatus;
  dialogueState: DialogueState;
  config: ConfigValues | null;
  appView: AppView;
  uiLanguage: UiLanguage;
  amplitude: number;
  muted: boolean;
  partialTranscript: string;
  finalTranscript: string;
  transcriptLanguage: Language | null;
  response: string;
  responseLanguage: Language | null;
  stages: Array<{ name: string; duration_ms: number }>;
  doctor: { ok: boolean; checks: Array<{ name: string; ok: boolean; details: string }> } | null;
  avatarDirection: AvatarDirection;
  previewDialogueState: DialogueState | null;
  textPromptEnabled: boolean;
  theme: ThemePreference;
  voiceGender: VoiceGender;
  notifications: AppNotification[];
  pinRequired: PinRequiredState | null;
  pinResult: PinResultState | null;
  dispatch: (event: EngineEvent) => void;
  notify: (message: string, tone?: NotificationTone) => void;
  dismissNotification: (id: string) => void;
  setConnectionStatus: (status: ConnectionStatus) => void;
  setMuted: (muted: boolean) => void;
  setAvatarDirection: (direction: AvatarDirection) => void;
  setAppView: (view: AppView) => void;
  setUiLanguage: (language: UiLanguage) => void;
  setTextPromptEnabled: (enabled: boolean) => void;
  setTheme: (theme: ThemePreference) => void;
  setVoiceGender: (voice: VoiceGender) => void;
  setFeatureFlagLocal: (flag: keyof FeatureFlags, enabled: boolean) => void;
  setConfigValueLocal: <K extends keyof ConfigValues>(key: K, value: ConfigValues[K]) => void;
  previewState: (state: DialogueState | null) => void;
  setPreviewState: (state: DialogueState | null) => void;
  reset: () => void;
  lastError: string | null;
}

const initialState = {
  connectionStatus: 'disconnected' as ConnectionStatus,
  dialogueState: 'idle' as DialogueState,
  config: null,
  appView: 'overlay' as AppView,
  uiLanguage: 'auto' as UiLanguage,
  amplitude: 0,
  muted: false,
  partialTranscript: '',
  finalTranscript: '',
  transcriptLanguage: null,
  response: '',
  responseLanguage: null,
  stages: [],
  doctor: null,
  avatarDirection: 'companion' as AvatarDirection,
  previewDialogueState: null,
  textPromptEnabled: true,
  theme: 'dark' as ThemePreference,
  // Jarvis is canonically a male voice; engine wiring lands later.
  voiceGender: 'male' as VoiceGender,
  notifications: [] as AppNotification[],
  pinRequired: null,
  pinResult: null,
  lastError: null,
};

export const useJarvisStore = create<JarvisState>()(
  persist(
    (set) => ({
      ...initialState,
      dispatch: (event) => {
        switch (event.type) {
          case 'state_changed':
            set({
              dialogueState: event.state,
              ...(event.state === 'idle'
                ? {
                    partialTranscript: '',
                    finalTranscript: '',
                    response: '',
                  }
                : {}),
              // pin_required only fires while genuinely PIN-pending; any other
              // state (including a different reason to be back in 'confirming',
              // e.g. slot-filling clarification) means it's stale.
              ...(event.state !== 'confirming' ? { pinRequired: null } : {}),
            });
            break;
          case 'partial_transcript':
            set({
              partialTranscript: event.text,
              transcriptLanguage: event.language,
            });
            break;
          case 'final_transcript':
            set({
              finalTranscript: event.text,
              transcriptLanguage: event.language,
              partialTranscript: '',
            });
            break;
          case 'response':
            set({
              response: event.text,
              responseLanguage: event.language,
            });
            break;
          case 'amplitude':
            set({ amplitude: event.level });
            break;
          case 'metrics':
            set({
              stages: event.stages,
              doctor: event.doctor,
            });
            break;
          case 'error':
            set((state) => ({
              lastError: event.message,
              notifications: appendNotification(state.notifications, event.message, 'error'),
            }));
            break;
          case 'config':
            set({ config: event.values });
            break;
          case 'pin_required':
            set({
              pinRequired: {
                description: event.description,
                attemptsRemaining: event.attempts_remaining,
                expiresInSeconds: event.expires_in_seconds,
                receivedAt: Date.now(),
              },
              pinResult: null,
            });
            break;
          case 'pin_result':
            set((state) => ({
              pinResult: {
                status: event.status,
                message: event.message,
                attemptsRemaining: event.attempts_remaining,
                receivedAt: Date.now(),
              },
              // "wrong" keeps the modal open (with the decremented count) so the
              // user can retry; every other outcome resolves the pending action.
              pinRequired:
                event.status === 'wrong' && state.pinRequired
                  ? { ...state.pinRequired, attemptsRemaining: event.attempts_remaining }
                  : null,
            }));
            break;
        }
      },
      setConnectionStatus: (status) => set({ connectionStatus: status }),
      setMuted: (muted) => set({ muted }),
      setAvatarDirection: (avatarDirection) => set({ avatarDirection }),
      setAppView: (appView) => set({ appView }),
      setUiLanguage: (uiLanguage) => set({ uiLanguage }),
      setTextPromptEnabled: (textPromptEnabled) => set({ textPromptEnabled }),
      setTheme: (theme) => set({ theme }),
      setVoiceGender: (voiceGender) => set({ voiceGender }),
      notify: (message, tone = 'info') =>
        set((state) => ({ notifications: appendNotification(state.notifications, message, tone) })),
      dismissNotification: (id) =>
        set((state) => ({ notifications: state.notifications.filter((item) => item.id !== id) })),
      setFeatureFlagLocal: (flag, enabled) =>
        set((state) => {
          if (!state.config) {
            return state;
          }

          return {
            config: {
              ...state.config,
              feature_flags: {
                ...state.config.feature_flags,
                [flag]: enabled,
              },
            },
          };
        }),
      setConfigValueLocal: (key, value) =>
        set((state) => {
          if (!state.config) {
            return state;
          }

          return {
            config: {
              ...state.config,
              [key]: value,
            },
          };
        }),
      previewState: (previewDialogueState) => set({ previewDialogueState }),
      setPreviewState: (previewDialogueState) => set({ previewDialogueState }),
      reset: () => set(initialState),
    }),
    {
      name: 'jarvis-ui',
      storage: createJSONStorage(() => localStorage),
      partialize: (state) => ({
        avatarDirection: state.avatarDirection,
        appView: state.appView,
        textPromptEnabled: state.textPromptEnabled,
        muted: state.muted,
        theme: state.theme,
        voiceGender: state.voiceGender,
      }),
    },
  ),
);

// Cross-window sync. Under Tauri the overlay and dashboard are separate OS
// windows, each with its own store instance. They share one localStorage
// (same origin), so when one window writes a persisted setting (avatar, view,
// text-prompt), the others rehydrate from it here. Also keeps browser tabs in
// sync. The storage event only fires in *other* documents, so no echo loop.
if (typeof window !== 'undefined') {
  window.addEventListener('storage', (event) => {
    if (event.key === 'jarvis-ui') {
      void useJarvisStore.persist.rehydrate();
    }
  });
}
