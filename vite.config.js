import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [react()],
  base: './', // <-- මේ පේළිය අනිවාර්යයෙන්ම තියෙන්න ඕනේ!
})
