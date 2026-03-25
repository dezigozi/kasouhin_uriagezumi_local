/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ['./src/**/*.{js,jsx}', './public/index.html'],
  theme: {
    extend: {
      fontFamily: {
        sans: ['"Noto Sans JP"', '"Segoe UI"', 'sans-serif'],
        mono: ['"JetBrains Mono"', '"Consolas"', 'monospace'],
      },
    },
  },
  plugins: [],
};
