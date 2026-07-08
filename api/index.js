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
// 5. NUEVA RUTA: CONSULTAR SECOP II
// ==========================================
app.get('/api/buscar-secop', async (req, res) => {
  try {
    const { contrato } = req.query;

    if (!contrato) {
      return res.status(400).json({ status: "error", message: "Falta el parámetro 'contrato' en la consulta" });
    }

    // Endpoint oficial de Datos Abiertos para SECOP II
    // Filtramos específicamente por el NIT de la Agencia APP para no traer basura de otras entidades
    const secopUrl = `https://datos.gov.co/resource/p6dx-8zbt.json?numero_del_contrato=${contrato}&nit_de_la_entidad=901035652`;

    const response = await axios.get(secopUrl);

    if (response.data.length === 0) {
      return res.status(404).json({
        status: "not_found",
        message: "No se encontró ningún contrato con ese número asignado a la Agencia APP en SECOP II."
      });
    }

    // Tomamos el primer contrato coincidente que devuelva la API del estado colombia
    const contratoData = response.data[0];

    // Mapeamos y limpiamos las variables para entregárselas masticadas al frontend
    res.json({
      status: "success",
      datos: {
        numeroContrato: contratoData.numero_del_contrato,
        objetoContrato: contratoData.objeto_del_contrato,
        nombreContratista: contratoData.nombre_del_contratista,
        documentoContratista: contratoData.documento_proveedor,
        valorContrato: contratoData.valor_del_contrato,
        fechaInicio: contratoData.fecha_de_inicio_del_contrato,
        fechaFin: contratoData.fecha_de_fin_del_contrato,
        supervisor: contratoData.nombre_de_la_dependencia // O el campo asignado al supervisor en SECOP
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
