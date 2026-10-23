// A few 16px stroke icons, drawn on a 16-unit grid. Inline so the app has
// no icon dependency.
const paths = {
  'arrow-right': 'M3 8h10M9 4l4 4-4 4',
  up: 'M4 10l4-4 4 4',
  down: 'M4 6l4 4 4-4',
  sun: 'M8 5.5a2.5 2.5 0 1 0 0 5 2.5 2.5 0 0 0 0-5zM8 1.5v1.5M8 13v1.5M1.5 8H3M13 8h1.5M3.4 3.4l1 1M11.6 11.6l1 1M3.4 12.6l1-1M11.6 4.4l1-1',
  moon: 'M13.5 9.5A5.5 5.5 0 0 1 6.5 2.5a5.5 5.5 0 1 0 7 7z',
  monitor: 'M2 3h12v8H2zM6 14h4M8 11v3',
  unified: 'M2.5 3h11M2.5 6.5h11M2.5 10h11M2.5 13.5h11',
  split: 'M2 3h12v10H2zM8 3v10',
  sliders: 'M2.5 4.5h6M11.5 4.5h2M2.5 11.5h2M7.5 11.5h6M10 3v3M6 10v3',
  download: 'M8 2v8M4.5 6.5L8 10l3.5-3.5M2.5 13.5h11',
  panel: 'M2 3h12v10H2zM10 3v10',
  upload: 'M8 10V2M4.5 5.5L8 2l3.5 3.5M2.5 10v3.5h11V10',
  file: 'M4 1.5h5l3 3v10H4zM9 1.5v3h3',
  lock: 'M4 7h8v6.5H4zM5.5 7V5a2.5 2.5 0 0 1 5 0v2',
} as const;

export type IconName = keyof typeof paths;

export function Icon({ name, className }: { name: IconName; className?: string }) {
  return (
    <svg className={`icon ${className ?? ''}`} width="16" height="16" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round" aria-hidden>
      <path d={paths[name]} />
    </svg>
  );
}
