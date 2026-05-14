/** @type {import('tailwindcss').Config} */
export default {
  content: ['./index.html', './src/**/*.{js,ts,jsx,tsx}'],
  theme: {
    extend: {
      colors: {
        vois: {
          50:      '#fff5f5',
          100:     '#ffe0e0',
          200:     '#fecaca',
          300:     '#fca5a5',
          400:     '#f87171',
          500:     '#ef4444',
          600:     '#e60000',   // VOIS / Vodafone Red
          700:     '#cc0000',
          800:     '#1a1a1a',   // Near-black
          900:     '#111111',
          950:     '#0a0a0a',
          pink:    '#e60000',
          accent:  '#ff3333',
        },
      },
      fontFamily: {
        sans: ['Inter', 'Calibri', 'system-ui', 'sans-serif'],
      },
      boxShadow: {
        card:       '0 2px 12px 0 rgba(230,0,0,0.10)',
        'card-hover': '0 6px 24px 0 rgba(230,0,0,0.18)',
      },
    },
  },
  plugins: [],
};
