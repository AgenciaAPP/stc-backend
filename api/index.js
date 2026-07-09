import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

app.use(cors());
app.use(express.json());

const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;

const LIST_ID_GENERAL = process.env.LIST_ID_GENERAL;
const LIST_ID_ACCIONES = process.env.LIST_ID_ACCIONES;
const LIST_ID_ASUNTOS = process.env.LIST_ID_ASUNTOS;
const LIST_ID_SISTEMAS = process.env.LIST_ID_SISTEMAS;
const LIST_ID_DIRECTORIO = process.env.LIST_ID_DIRECTORIO;

async function getMicrosoftGraphToken() {
  const url = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`;
  const params = new URLSearchParams();
  params.append('client_id', CLIENT_ID);
  params.append('scope', 'https://graph.microsoft.com/.default');
  params.append('client_secret', CLIENT_SECRET);
  params.append('grant_type', 'client_credentials');

  try {
    const response = await axios.post(url, params, {
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' }
    });
    return response.data.access_token;
  } catch (error) {
    console.error('Error de autenticación con Azure AD:', error.response?.data || error.message);
    throw new Error('No se pudo adquirir el Token de Acceso de Microsoft.');
  }
}

app.get('/', (req, res) => {
  res.send('Servidor STC de la Agencia APP operando correctamente con columnas de SharePoint homologadas.');
});

app.get('/api/status', (req, res) => {
  res.json({ status: "online", timestamp: new Date() });
});

// ==========================================
// 3. RUTA: CONSULTAR SECOP II 
// ==========================================
app.get('/api/buscar-secop', async (req, res) => {
  try {
    const { contrato } = req.query;
    if (!contrato) {
      return res.status(400).json({ success: false, message: "Falta el parámetro 'contrato' en la consulta" });
    }

    const nitAgenciaAPP = "900623766"; 
    const secopUrl = `https://www.datos.gov.co/resource/jbjy-vk9h.json?referencia_del_contrato=${encodeURIComponent(contrato)}&nit_entidad=${nitAgenciaAPP}`;
    
    const response = await axios.get(secopUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
      }
    });
    
    if (response.data && response.data.length > 0) {
      const contratoData = response.data[0];
      res.json({
        success: true,
        nombre: contratoData.proveedor_adjudicado || "No registrado",
        cedula: contratoData.documento_proveedor || "No registrado",
        objeto: contratoData.objeto_del_contrato || "No registrado",
        nombreSupervisor: contratoData.nombre_supervisor || "No registrado",
        cedulaSupervisor: contratoData.n_mero_de_documento_supervisor || "No registrado",
        fechaFirma: contratoData.fecha_de_firma || null // Captura para la columna de fecha inicio
      });
    } else {
      res.json({ success: false, message: "No se encontró ningún contrato con esa referencia asignado a la Agencia APP en SECOP II." });
    }
  } catch (error) {
    console.error("Error consultando SECOP II:", error.message);
    res.status(500).json({ success: false, message: "Error al conectarse con el servidor gubernamental." });
  }
});

// ==========================================
// 4. RUTA: PERSISTENCIA COMPLETA EN SHAREPOINT (FINALIZAR Y ENVIAR)
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;

  if (!datosGenerales || !datosGenerales.numeroContrato) {
    return res.status(400).json({ success: false, message: 'Faltan los datos contractuales mínimos.' });
  }

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    // 1. Homologación: STC_General
    const generalPayload = {
      fields: {
        Title: datosGenerales.numeroContrato, // Guarda número de contrato
        Supervisor: datosGenerales.supervisor,
        Objetocontractual: datosGenerales.objetoContrato,
        Fechadeiniciodelcontrato: datosGenerales.fechaInicio || '', 
        Dependencia: datosGenerales.dependencia,
        Contratista: datosGenerales.nombreContratista,
        Fechadediligenciamiento: datosGenerales.isFinal ? new Date().toISOString().split('T')[0] : '',
        NIT_x002f_CC: datosGenerales.cedula,
        Lineamientos: datosGenerales.lineamientos || '',
        Recomendaciones: datosGenerales.recomendacionesAcciones || '',
        CorreoContratista: datosGenerales.correoContratista,
        Estado: datosGenerales.isFinal ? 'Finalizado' : 'En diligenciamiento'
      }
    };
    await axios.post(`${graphBaseUrl}/${LIST_ID_GENERAL}/items`, generalPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    // Solo si es el envío final guardamos los multirregistros hijos usando como llaves el número de contrato o cédula
    if (datosGenerales.isFinal) {
      // 2. Homologación: STC_Acciones
      for (const item of acciones) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, {
          fields: {
            Title: item.proceso,
            Prioridad: item.prioridad,
            Productosentrega: item.productos,
            Acci_x00f3_nparalatransferenciad: item.accionConocimiento, // Nuevo campo inyectado
            escribac_x00f3_mosellev_x00f3_a: item.ejecucion,
            Fechaenqueseejecut_x00f3_laacci_: item.fecha,
            Ruta_x0028_s_x0029_dondereposa_x: item.ruta,
            Observaciones: item.obs
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }

      // 3. Homologación: STC_Asuntos
      for (const item of asuntos) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items`, {
          fields: {
            Title: item.tramite,
            Estado: item.estado,
            Entidad_x002f_Dependencia: item.entidad,
            Accionespendientesporrealizar: item.accionesPendientes,
            Fechal_x00ed_mite: item.fecha
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }

      // 4. Homologación: STC_Sistemas
      for (const item of sistemas) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items`, {
          fields: {
            Title: item.nombre,
            Usuario: item.usuario,
            Contrase_x00f1_a: item.contrasena,
            Observaciones: item.obs
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }

      // 5. Homologación: STC_Directorio
      for (const item of directorio) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items`, {
          fields: {
            Title: item.nombre,
            Tel_x00e9_fono: item.tel,
            E_x002d_Mail: item.correo,
            Tipodecontacto: item.tipo,
            Entidad_x002f_Dependencia: item.entidad,
            Recomendaciones: item.reco
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    return res.status(200).json({ success: true, message: '¡Acta procesada con éxito!' });

  } catch (error) {
    const apiErrorDetail = error.response?.data?.error || error.message;
    console.error('Error inyectando en SharePoint:', JSON.stringify(apiErrorDetail));
    return res.status(500).json({ success: false, message: 'Error en Microsoft Graph.', detail: apiErrorDetail });
  }
});

export default app;
