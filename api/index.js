import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

// Cargar variables de entorno en desarrollo local (Vercel las lee nativamente)
dotenv.config();

const app = express();

// Configuración de CORS y lectura de JSON
app.use(cors());
app.use(express.json());

// Ruta de prueba para verificar que el backend está vivo
app.get('/api/saludo', (req, res) => {
  res.json({ 
    status: "ok", 
    message: "Servidor del Sistema de Transferencia de Conocimiento operativo (Agencia APP) 🚀" 
  });
});

// Función interna para obtener el Token de Acceso de Microsoft Graph de forma segura
async function getMicrosoftToken() {
  const url = `https://login.microsoftonline.com/${process.env.TENANT_ID}/oauth2/v2.0/token`;
  
  const params = new URLSearchParams();
  params.append('client_id', process.env.CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', process.env.CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  try {
    const response = await axios.post(url, params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return response.data.access_token;
  } catch (error) {
    console.error("Error obteniendo el token de Azure:", error.response?.data || error.message);
    throw new Error("No se pudo autenticar con Microsoft Entra ID");
  }
}

// Ruta de prueba para verificar la conexión real con Microsoft Graph
app.get('/api/test-conexion', async (req, res) => {
  try {
    const token = await getMicrosoftToken();
    res.json({ 
      conexion: "exitosa", 
      message: "El backend se autenticó correctamente con Azure Entra ID. Token generado." 
    });
  } catch (error) {
    res.status(500).json({ error: error.message });
  }
});

// Exportar la app para que Vercel la maneje como Serverless
export default app;

// Solo encender el puerto si se corre de forma local (en Vercel no es necesario)
if (process.env.PORT) {
  app.listen(process.env.PORT, () => {
    console.log(`Servidor local corriendo en el puerto ${process.env.PORT}`);
  });
}
