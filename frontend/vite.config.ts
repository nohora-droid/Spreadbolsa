import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'
import tailwindcss from '@tailwindcss/vite'

// https://vite.dev/config/
export default defineConfig({
  plugins: [react(), tailwindcss()],

  // ── Configuración de build para producción (Vercel) ───────────────────────
  build: {
    // Directorio de salida estándar — Vercel lo detecta automáticamente
    outDir: 'dist',

    // Sin sourcemaps en producción para reducir el tamaño del bundle
    sourcemap: false,

    // Umbral de advertencia de chunk en kB (recharts + xlsx hacen el bundle grande)
    chunkSizeWarningLimit: 1200,

    rollupOptions: {
      output: {
        // Separar vendors grandes en chunks independientes para mejor caché
        // Nota: Vite 8 (rolldown) requiere manualChunks como función, no objeto
        manualChunks(id: string) {
          if (id.includes('node_modules/react') || id.includes('node_modules/react-dom')) {
            return 'vendor-react'
          }
          if (id.includes('node_modules/recharts')) {
            return 'vendor-recharts'
          }
          if (id.includes('node_modules/xlsx')) {
            return 'vendor-xlsx'
          }
        },
      },
    },
  },
})
