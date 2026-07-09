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
    console.error('Error de autenticación con Azure AD:', error.message);
    throw new Error('No se pudo adquirir el Token de Acceso de Microsoft.');
  }
}

app.get('/', (req, res) => {
  res.send('Servidor STC de la Agencia APP operando en vivo con persistencia homologada.');
});

// ==========================================
// RUTA: CONSULTAR SECOP II
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
      headers: { 'User-Agent': 'Mozilla/5.0' }
    });
    
    if (response.data && response.data.length > 0) {
      const contratoData = response.data[0];
      
      let fechaLimpia = contratoData.fecha_de_firma || null;
      if (fechaLimpia && fechaLimpia.includes('T')) {
        fechaLimpia = fechaLimpia.split('T')[0];
      }

      res.json({
        success: true,
        nombre: contratoData.proveedor_adjudicado || "No registrado",
        cedula: contratoData.documento_proveedor || "No registrado",
        objeto: contratoData.objeto_del_contrato || "No registrado",
        nombreSupervisor: contratoData.nombre_supervisor || "No registrado",
        cedulaSupervisor: contratoData.n_mero_de_documento_supervisor || "No registrado",
        fechaFirma: fechaLimpia
      });
    } else {
      res.json({ success: false, message: "No se encontró ningún contrato con esa referencia en SECOP II." });
    }
  } catch (error) {
    res.status(500).json({ success: false, message: "Error al conectarse con el SECOP II." });
  }
});

// ==========================================
// RUTA: CREACIÓN INICIAL POR TALENTO HUMANO (HABILITACIÓN)
// ==========================================
app.post('/api/habilitar-contrato', async (req, res) => {
  const { contrato, contratista, cedula, objeto, supervisor, fechaInicio } = req.body;

  if (!contrato || !cedula) {
    return res.status(400).json({ success: false, message: 'Faltan datos obligatorios.' });
  }

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    const habilitarPayload = {
      fields: {
        Title: contrato, 
        Contratista: contratista,
        NIT_x002f_CC: String(cedula).trim(),
        Objetocontractual: objeto,
        Supervisor: supervisor,
        Fechadeiniciodelcontrato: fechaInicio || '',
        Estado: 'Sin diligenciar'
      }
    };

    await axios.post(`${graphBaseUrl}/${LIST_ID_GENERAL}/items`, habilitarPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    return res.status(200).json({ success: true });
  } catch (error) {
    return res.status(500).json({ success: false, detail: error.response?.data?.error || error.message });
  }
});

// ==========================================
// RUTA: OBTENER TODOS LOS CONTRATOS EN VIVO 
// ==========================================
app.get('/api/contratos', async (req, res) => {
  try {
    const token = await getMicrosoftGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_GENERAL}/items?expand=fields`;
    
    const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
    
    const listaFormateada = response.data.value.map(item => ({
      idSharePoint: item.id,
      name: item.fields.Contratista,
      contract: item.fields.Title,
      boss: item.fields.Supervisor,
      objeto: item.fields.Objetocontractual, 
      status: item.fields.Estado ? item.fields.Estado.toUpperCase() : 'SIN DILIGENCIAR',
      cedula: item.fields.NIT_x002f_CC,
      lineamientos: item.fields.Lineamientos || '',
      recomendaciones: item.fields.Recomendaciones || '',
      dependencia: item.fields.Dependencia || '',
      correo: item.fields.CorreoContratista || '' // Propiedad homologada para consulta general
    }));

    res.json({ success: true, data: listaFormateada });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Error trayendo datos de SharePoint.' });
  }
});

// ==========================================
// RUTA: LOGIN CONTRATISTA
// ==========================================
app.get('/api/login-contratista', async (req, res) => {
  const { cedula } = req.query;
  try {
    const token = await getMicrosoftGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_GENERAL}/items?expand=fields`;
    const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
    
    const match = response.data.value.find(item => {
      const nitFila = item.fields.NIT_x002f_CC ? String(item.fields.NIT_x002f_CC).trim() : '';
      return nitFila === String(cedula).trim();
    });
    
    if (match) {
      res.json({
        success: true,
        exists: true,
        idSharePoint: match.id,
        nombre: match.fields.Contratista,
        contract: match.fields.Title,
        objeto: match.fields.Objetocontractual,
        supervisor: match.fields.Supervisor || 'No registrado',
        estado: match.fields.Estado || 'Sin diligenciar',
        correo: match.fields.CorreoContratista || '',
        dependencia: match.fields.Dependencia || '',
        lineamientos: match.fields.Lineamientos || '', 
        recomendaciones: match.fields.Recomendaciones || ''
      });
    } else {
      res.json({ success: true, exists: false });
    }
  } catch (error) {
    res.status(500).json({ success: false, detail: error.message });
  }
});

// ==========================================
// RUTA: PATCH ACTUALIZACIÓN
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { idSharePoint, datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;

  if (!idSharePoint) {
    return res.status(400).json({ success: false, message: 'Falta el identificador de registro de SharePoint.' });
  }

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    const generalPayload = {
      fields: {
        Dependencia: datosGenerales.dependencia,
        Lineamientos: datosGenerales.lineamientos || '',
        Recomendaciones: datosGenerales.recomendacionesAcciones || '',
        CorreoContratista: datosGenerales.correoContratista,
        Estado: datosGenerales.isFinal ? 'Finalizado' : 'En diligenciamiento',
        Fechadediligenciamiento: datosGenerales.isFinal ? new Date().toISOString().split('T')[0] : ''
      }
    };

    await axios.patch(`${graphBaseUrl}/${LIST_ID_GENERAL}/items/${idSharePoint}`, generalPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    if (datosGenerales.isFinal) {
      if (acciones && acciones.length > 0) {
        for (const item of acciones) {
          await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, {
            fields: {
              Title: item.proceso,
              Prioridad: item.prioridad,
              Productosentrega: item.productos,
              Acci_x00f3_nparalatransferenciad: item.accionConocimiento, 
              escribac_x00f3_mosellev_x00f3_a: item.ejecucion,
              Fechaenqueseejecut_x00f3_laacci_: item.fecha,
              Ruta_x0028_s_x0029_dondereposa_x: item.ruta,
              Observaciones: item.obs
            }
          }, { headers: { 'Authorization': `Bearer ${token}` } });
        }
      }

      if (asuntos && asuntos.length > 0) {
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
      }

      if (sistemas && sistemas.length > 0) {
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
      }

      if (directorio && directorio.length > 0) {
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
    }

    return res.status(200).json({ success: true });

  } catch (error) {
    return res.status(500).json({ success: false, detail: error.response?.data?.error || error.message });
  }
});

export default app;
