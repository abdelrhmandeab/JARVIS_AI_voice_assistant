export interface SelectOption {
  label: string;
  value: string;
}

interface SelectProps {
  label: string;
  value: string;
  options: SelectOption[];
  onChange: (value: string) => void;
  disabled?: boolean;
}

export function Select({ label, value, options, onChange, disabled = false }: SelectProps) {
  return (
    <label className="grid gap-2 text-sm">
      <span className="font-medium text-slate-600 dark:text-white/72">{label}</span>
      <select
        value={value}
        disabled={disabled}
        onChange={(event) => onChange(event.target.value)}
        className="h-11 rounded border border-black/[0.08] bg-black/[0.04] px-3 text-slate-800 outline-none transition focus:border-[#0F8FB8]/65 disabled:cursor-not-allowed disabled:opacity-45 dark:border-white/10 dark:bg-[#111118] dark:text-white dark:focus:border-[#8EEBFF]/65"
      >
        {options.map((option) => (
          <option key={option.value} value={option.value}>
            {option.label}
          </option>
        ))}
      </select>
    </label>
  );
}
