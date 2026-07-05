import { useEffect, useRef, useState } from 'react';
import { AnimatePresence, motion } from 'motion/react';
import type { UICommand } from '../../protocol';
import { useJarvisStore } from '../../stores/jarvisStore';

interface PinModalProps {
  send: (cmd: UICommand) => void;
}

const ARABIC_TEXT_PATTERN = /[؀-ۿ]/;
const AUTO_SUBMIT_MIN_DIGITS = 4;
const AUTO_SUBMIT_DEBOUNCE_MS = 500;
const TERMINAL_LINGER_MS = 1800;

function directionOf(text: string) {
  return ARABIC_TEXT_PATTERN.test(text) ? 'rtl' : 'ltr';
}

export function PinModal({ send }: PinModalProps) {
  const pinRequired = useJarvisStore((state) => state.pinRequired);
  const pinResult = useJarvisStore((state) => state.pinResult);
  const partialTranscript = useJarvisStore((state) => state.partialTranscript);
  const connectionStatus = useJarvisStore((state) => state.connectionStatus);

  const [digits, setDigits] = useState('');
  const [shakeKey, setShakeKey] = useState(0);
  const [secondsLeft, setSecondsLeft] = useState(0);
  const [lingering, setLingering] = useState<typeof pinResult>(null);
  const inputRef = useRef<HTMLInputElement | null>(null);
  const debounceRef = useRef<number | null>(null);
  const lingerTimerRef = useRef<number | null>(null);

  const visible = pinRequired !== null || lingering !== null;

  // Reset local entry state each time a fresh PIN request comes in.
  useEffect(() => {
    if (pinRequired) {
      setDigits('');
      setSecondsLeft(Math.max(0, Math.round(pinRequired.expiresInSeconds)));
      window.setTimeout(() => inputRef.current?.focus(), 0);
    }
  }, [pinRequired?.receivedAt]);

  // Client-side countdown, ticking down from the value the engine sent.
  useEffect(() => {
    if (!pinRequired) return;
    const interval = window.setInterval(() => {
      setSecondsLeft((seconds) => Math.max(0, seconds - 1));
    }, 1000);
    return () => window.clearInterval(interval);
  }, [pinRequired?.receivedAt]);

  // React to attempt results: shake + refocus on a wrong guess; for a
  // terminal outcome (executed/locked/no_pending) linger briefly so the user
  // sees the final message before the modal disappears.
  useEffect(() => {
    if (!pinResult) return;

    if (lingerTimerRef.current !== null) {
      window.clearTimeout(lingerTimerRef.current);
      lingerTimerRef.current = null;
    }

    if (pinResult.status === 'wrong') {
      setDigits('');
      setShakeKey((key) => key + 1);
      window.setTimeout(() => inputRef.current?.focus(), 0);
      return;
    }

    setLingering(pinResult);
    lingerTimerRef.current = window.setTimeout(() => {
      setLingering(null);
      lingerTimerRef.current = null;
    }, TERMINAL_LINGER_MS);
  }, [pinResult?.receivedAt]);

  useEffect(() => {
    return () => {
      if (debounceRef.current !== null) window.clearTimeout(debounceRef.current);
      if (lingerTimerRef.current !== null) window.clearTimeout(lingerTimerRef.current);
    };
  }, []);

  const disabled = connectionStatus !== 'connected' || !pinRequired;
  const description = pinRequired?.description?.trim() || 'Confirm this action';
  const expired = pinRequired !== null && secondsLeft <= 0;

  const submit = (value: string) => {
    const trimmed = value.trim();
    if (!trimmed || disabled) return;
    if (debounceRef.current !== null) {
      window.clearTimeout(debounceRef.current);
      debounceRef.current = null;
    }
    send({ type: 'pin_attempt', pin: trimmed });
  };

  const handleChange = (raw: string) => {
    const next = raw.replace(/\D/g, '').slice(0, 12);
    setDigits(next);

    if (debounceRef.current !== null) {
      window.clearTimeout(debounceRef.current);
      debounceRef.current = null;
    }
    if (next.length >= AUTO_SUBMIT_MIN_DIGITS) {
      debounceRef.current = window.setTimeout(() => submit(next), AUTO_SUBMIT_DEBOUNCE_MS);
    }
  };

  const handleCancel = () => {
    send({ type: 'text_command', text: 'cancel' });
  };

  return (
    <AnimatePresence>
      {visible ? (
        <motion.div
          key="pin-modal"
          role="dialog"
          aria-modal="true"
          aria-label="PIN confirmation"
          initial={{ opacity: 0, y: 12, scale: 0.97 }}
          animate={{ opacity: 1, y: 0, scale: 1 }}
          exit={{ opacity: 0, y: 8, scale: 0.97 }}
          transition={{ duration: 0.22, ease: [0.16, 1, 0.3, 1] }}
          className="pointer-events-auto z-20 w-[260px] rounded-md border border-black/10 bg-white/85 p-4 text-slate-800 shadow-2xl shadow-black/15 backdrop-blur-xl dark:border-white/10 dark:bg-black/80 dark:text-white dark:shadow-black/40"
        >
          {lingering ? (
          <div className="grid gap-2 text-center">
            <p
              dir={directionOf(lingering.message)}
              className={`text-sm font-medium ${
                lingering.status === 'executed'
                  ? 'text-cyan-700 dark:text-cyan-100'
                  : 'text-red-600 dark:text-red-200'
              }`}
            >
              {lingering.message}
            </p>
          </div>
        ) : (
          <motion.div
            key={shakeKey}
            animate={shakeKey ? { x: [0, -7, 7, -5, 5, 0] } : undefined}
            transition={{ duration: 0.35 }}
            className="grid gap-3"
          >
            <div>
              <p className="text-[11px] font-semibold uppercase tracking-[0.14em] text-cyan-700/80 dark:text-cyan-50/70">
                PIN required
              </p>
              <p dir={directionOf(description)} className="mt-1 text-sm text-slate-700 dark:text-white/85">
                {description}
              </p>
            </div>

            {expired ? (
              <p className="text-xs text-red-600 dark:text-red-200">
                This PIN request expired. Repeat your command to try again.
              </p>
            ) : (
              <>
                <input
                  ref={inputRef}
                  type="password"
                  inputMode="numeric"
                  autoComplete="off"
                  value={digits}
                  disabled={disabled}
                  onChange={(event) => handleChange(event.target.value)}
                  onKeyDown={(event) => {
                    if (event.key === 'Enter') submit(digits);
                    if (event.key === 'Escape') handleCancel();
                  }}
                  placeholder="Enter PIN"
                  aria-label="PIN"
                  className="h-11 w-full rounded border border-black/10 bg-black/[0.04] px-3 text-center text-lg tracking-[0.4em] text-slate-900 outline-none transition focus:border-cyan-500/60 focus:bg-black/[0.06] disabled:opacity-50 dark:border-white/10 dark:bg-white/[0.07] dark:text-white dark:focus:border-[#8EEBFF]/70 dark:focus:bg-white/[0.1]"
                />

                <div className="flex items-center justify-between text-[11px] text-slate-500 dark:text-white/50">
                  <span>{pinRequired ? `${pinRequired.attemptsRemaining} attempt(s) left` : ''}</span>
                  <span>{secondsLeft}s</span>
                </div>

                {partialTranscript ? (
                  <p dir={directionOf(partialTranscript)} className="text-[11px] text-slate-500 dark:text-white/45">
                    Hearing: {partialTranscript}
                  </p>
                ) : (
                  <p className="text-[11px] text-slate-400 dark:text-white/35">Type it, or just say your PIN.</p>
                )}
              </>
            )}

            <div className="flex gap-2">
              <button
                type="button"
                onClick={handleCancel}
                className="h-9 flex-1 rounded border border-black/10 bg-black/[0.04] text-xs font-medium text-slate-600 transition hover:bg-black/[0.07] dark:border-white/10 dark:bg-white/[0.05] dark:text-white/70 dark:hover:bg-white/[0.09]"
              >
                Cancel
              </button>
              {!expired ? (
                <button
                  type="button"
                  disabled={disabled || !digits}
                  onClick={() => submit(digits)}
                  className="h-9 flex-1 rounded border border-cyan-500/40 bg-cyan-400/20 text-xs font-medium text-cyan-700 transition hover:bg-cyan-400/30 disabled:cursor-not-allowed disabled:opacity-40 dark:border-[#8EEBFF]/30 dark:bg-[#8EEBFF]/12 dark:text-[#DDFBFF] dark:hover:bg-[#8EEBFF]/18"
                >
                  Confirm
                </button>
              ) : null}
            </div>
          </motion.div>
          )}
        </motion.div>
      ) : null}
    </AnimatePresence>
  );
}
