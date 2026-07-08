import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

app.use(cors());
app.use(express.json());

// 1. Ruta de diagnóstico
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

// 3. Ruta para validar el estado de conexión con Azure
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

// 4. Ruta para guardar un registro nuevo en STC_General
app.post('/api/guardar-general', async (req, res) => {
  try {
    const datosFormulario = req.body;
    const token = await getMicrosoftToken();
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${process.env.SITE_ID}/lists/${process.env.LIST_ID_GENERAL}/items`;

    const payload = {
      fields: {
        Title: datosFormulario.cedula,
        NombreContratista: datosFormulario.nombreContratista,
        CorreoContratista: datosFormulario.correoContratista,
        NumeroContrato: datosFormulario.numeroContrato,
        ObjetoContrato: datosFormulario.objetoContrato,
        Supervisor: datosFormulario.supervisor,
        Dependencia: datosFormulario.dependencia,
        Estado: "PROCESO"
      }
    };

    const response = await axios.post(graphUrl, payload, {
      headers: {
        'Authorization': `Bearer ${token}`,
        'Content-Type': 'application/json'
      }
    });

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

// ==========================================
// 5. NUEVA RUTA: CONSULTAR SECOP II (CORREGIDA POR NIT)
// ==========================================
app.get('/api/buscar-secop', async (req, res) => {
  try {
    const { contrato } = req.query;

    if (!contrato) {
      return res.status(400).json({ status: "error", message: "Falta el parámetro 'contrato' en la consulta" });
    }

    // Filtramos simultáneamente por la referencia del contrato Y por el NIT numérico de la Agencia APP
    const nitAgenciaAPP = "900623766"; 
    const secopUrl = `https://www.datos.gov.co/resource/jbjy-vk9h.json?referencia_del_contrato=${encodeURIComponent(contrato)}&nit_entidad=${nitAgenciaAPP}`;

    const response = await axios.get(secopUrl);

    if (response.data.length === 0) {
      return res.status(404).json({
        status: "not_found",
        message: "No se encontró ningún contrato con esa referencia asignado a la Agencia APP en SECOP II."
      });
    }

    // Tomamos el contrato coincidente que es 100% de la Agencia APP
    const contratoData = response.data[0];

    // Mapeamos las variables con los nombres de columna confirmados
    res.json({
      status: "success",
      datos: {
        numeroContrato: contratoData.referencia_del_contrato,
        objetoContrato: contratoData.objeto_del_contrato,
        nombreContratista: contratoData.proveedor_adjudicado,
        documentoContratista: contratoData.documento_proveedor,
        valorContrato: contratoData.valor_del_contrato,
        fechaInicio: contratoData.fecha_de_inicio_del_contrato,
        fechaFin: contratoData.fecha_de_fin_del_contrato,
        supervisor: contratoData.nombre_supervisor !== "No definido" ? contratoData.nombre_supervisor : ""
      }
    });

  } catch (error) {
    console.error("Error consultando la API de SECOP II:", error.message);
    res.status(500).json({
      status: "error",
      message: "Error al conectarse con el servidor de Datos Abiertos de Colombia",
      detalle: error.message
    });
  }
});

export default app;

if (process.env.PORT) {
  app.listen(process.env.PORT, () => {
    console.log(`Servidor local corriendo en el puerto ${process.env.PORT}`);
  });
}
