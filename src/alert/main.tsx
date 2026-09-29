import { StrictMode } from 'react'
import { createRoot } from 'react-dom/client'
import { AlertApp } from './AlertApp'
import '../styles.css'

createRoot(document.getElementById('root')!).render(
  <StrictMode>
    <AlertApp />
  </StrictMode>
)
