import { FormEvent, useState } from 'react';
import type { Language, UICommand } from '../../protocol';

interface PromptInputProps {
  send: (cmd: UICommand) => void;
  // 'glass' (default) is the always-dark HUD look used by the floating
  // Overlay window, which never reacts to the theme setting. 'surface' is
  // theme-aware, for embedding in the Dashboard where the surrounding panel
  // already flips between light and dark.
  variant?: 'glass' | 'surface';
}

const ARABIC_TEXT_PATTERN = /[\u0600-\u06FF]/;

function detectLanguage(text: string): Language {
  return ARABIC_TEXT_PATTERN.test(text) ? 'ar' : 'en';
}

export function PromptInput({ send, variant = 'glass' }: PromptInputProps) {
  const [text, setText] = useState('');

  const handleSubmit = (event: FormEvent<HTMLFormElement>) => {
    event.preventDefault();

    const trimmedText = text.trim();
    if (!trimmedText) {
      return;
    }

    send({
      type: 'text_command',
      text: trimmedText,
      language: detectLanguage(trimmedText),
    });
    setText('');
  };

  const inputClassName =
    variant === 'surface'
      ? 'h-10 min-w-0 flex-1 rounded border border-black/10 bg-black/[0.03] px-3 text-sm text-slate-800 outline-none transition placeholder:text-slate-500 focus:border-[#0F8FB8]/70 focus:bg-black/[0.05] dark:border-white/10 dark:bg-white/[0.07] dark:text-white dark:placeholder:text-white/50 dark:focus:border-[#8EEBFF]/70 dark:focus:bg-white/[0.1]'
      : 'h-10 min-w-0 flex-1 rounded border border-white/10 bg-white/[0.07] px-3 text-sm text-white outline-none transition focus:border-[#8EEBFF]/70 focus:bg-white/[0.1]';

  return (
    <form onSubmit={handleSubmit} className="flex min-w-0 flex-1 items-center gap-2">
      <input
        type="text"
        value={text}
        onChange={(event) => setText(event.target.value)}
        placeholder="Type a prompt"
        className={inputClassName}
      />
      <button
        type="submit"
        aria-label="Send"
        title="Send"
        className="grid h-10 w-10 shrink-0 place-items-center rounded border border-[#8EEBFF]/40 bg-[#8EEBFF]/15 text-cyan-700 transition hover:bg-[#8EEBFF]/25 dark:border-[#8EEBFF]/30 dark:bg-[#8EEBFF]/12 dark:text-[#DDFBFF] dark:hover:bg-[#8EEBFF]/18"
      >
        <svg width="17" height="17" viewBox="0 0 24 24" fill="none" aria-hidden="true">
          <path
            d="M22 2 L11 13 M22 2 L15 22 L11 13 L2 9 Z"
            stroke="currentColor"
            strokeWidth="1.8"
            strokeLinecap="round"
            strokeLinejoin="round"
          />
        </svg>
      </button>
    </form>
  );
}
