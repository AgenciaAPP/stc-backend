import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

app.use(cors());
app.use(express.json());

// 1. Ruta de diagnóstico para verificar que el backend sigue vivo
app.get('/api/saludo', (req, res) => {
  res.json({ 
    status: "ok", 
    message: "Servidor del Sistema de Transferencia de Conocimiento operativo (Agencia APP) 🚀" 
  });
});

// 2. Función interna para obtener el Token de Acceso de Microsoft Graph
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

// 3. Ruta para validar el estado de conexión
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

// ==========================================
// 4. NUEVA RUTA: GUARDAR EN STC_GENERAL
// ==========================================
app.post('/api/guardar-general', async (req, res) => {
  try {
    // Recibir los campos enviados desde el formulario de la Web App
    const datosFormulario = req.body;

    // Obtener el token de acceso vigente
    const token = await getMicrosoftToken();

    // Endpoint de Microsoft Graph para insertar un elemento en una lista específica
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.LIST_ID_GENERAL}/items`;

    // Estructura obligatoria que pide Microsoft Graph para las columnas de SharePoint
    const payload = {
      fields: {
        Title: datosFormulario.cedula, // Usamos la cédula como identificador principal (Title)
        NombreContratista: datosFormulario.nombreContratista,
        CorreoContratista: datosFormulario.correoContratista,
        NumeroContrato: datosFormulario.numeroContrato,
        ObjetoContrato: datosFormulario.objetoContrato,
        Supervisor: datosFormulario.supervisor,
        Dependencia: datosFormulario.dependencia,
        Estado: "PROCESO" // Todo contrato nuevo arranca en estado de Proceso de transferencia
      }
    };

    // Disparar la petición POST hacia SharePoint
    const response = await axios.post(graphUrl, payload, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

    // Responder a la Web App con el ID único que SharePoint le asignó al registro
    res.json({
      status: "success",
      message: "Registro creado exitosamente en STC_General",
      sharepoint_id: response.data.id
    });

  } catch (error) {
    console.error("Error al guardar en STC_General:", error.response?.data || error.message);
    res.status(500).json({
      status: "error",
      message: "No se pudo registrar la información en SharePoint",
      detalle: error.response?.data || error.message
    });
  }
});

export default app;

if (process.env.PORT) {
  app.listen(process.env.PORT, () => {
    console.log(`Servidor local corriendo en el puerto ${process.env.PORT}`);
  });
}
