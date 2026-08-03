/**
 * Eduport brand theme.
 *
 * The colour values were recovered from the last compiled stylesheet, so they
 * match what the app shipped with — `brand` is the Eduport orange, `accent` the
 * blue used for links and secondary actions.
 *
 * @type {import('tailwindcss').Config}
 */
module.exports = {
  content: [
    "./app/**/*.{js,jsx}",
    "./components/**/*.{js,jsx}",
    "./state/**/*.{js,jsx}",
    "./lib/**/*.{js,jsx}",
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          50: "#fff4ee",
          100: "#ffe4d5",
          400: "#f5854e",
          500: "#f26a2c",
          600: "#e15718",
          700: "#bb4512",
        },
        accent: {
          500: "#2e7cf6",
        },
      },
      // Wired to the next/font variables declared in app/layout.jsx.
      fontFamily: {
        sans: ["var(--font-inter)", "ui-sans-serif", "system-ui", "sans-serif"],
        serif: ["var(--font-lora)", "ui-serif", "Georgia", "serif"],
        hand: ["var(--font-caveat)", "ui-serif", "cursive"],
      },
    },
  },
  plugins: [],
};
