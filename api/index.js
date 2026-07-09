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
  res.send('Servidor STC de la Agencia APP operando estrictamente con la matriz homologada de SharePoint.');
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
    const response = await axios.get(secopUrl, { headers: { 'User-Agent': 'Mozilla/5.0' } });
    
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
        Supervisor: supervisor,
        Objetocontractual: objeto,
        Fechadeiniciodelcontrato: fechaInicio || '',
        Contratista: contratista,
        NIT_x002f_CC: String(cedula).trim(),
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
      contract: item.fields.Title,
      boss: item.fields.Supervisor,
      objeto: item.fields.Objetocontractual, 
      name: item.fields.Contratista,
      cedula: item.fields.NIT_x002f_CC,
      status: item.fields.Estado ? item.fields.Estado.toUpperCase() : 'SIN DILIGENCIAR',
      lineamientos: item.fields.Lineamientos || '',
      recomendaciones: item.fields.Recomendaciones || '',
      dependencia: item.fields.Dependencia || '',
      correo: item.fields.CorreoContratista || ''
    }));
    res.json({ success: true, data: listaFormateada });
  } catch (error) {
    res.status(500).json({ success: false, message: 'Error trayendo datos de SharePoint.' });
  }
});

// ==========================================
// RUTA: ENCONTRAR TABLAS HIJAS POR CÉDULA DEL CONTRATISTA
// ==========================================
app.get('/api/obtener-detalles-hijos', async (req, res) => {
  const { cedula } = req.query;
  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    const resAcciones = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    
    // Mapeamos asumiendo que un campo relacional o de control asocia el hijo al contratista (usamos la cédula del contratista)
    const filtradoAcciones = resAcciones.data.value
      .filter(item => {
        // Buscamos coincidencia con la cédula vinculadora (si guardas la cédula en una propiedad o notas de observaciones)
        return true; // Por ahora pasamos el listado para renderizar localmente
      })
      .map(item => ({
        idSharePointHijo: item.id,
        proceso: item.fields.Title || 'No registrado', 
        prioridad: item.fields.Prioridad,
        productos: item.fields.Productosentrega,
        accionConocimiento: item.fields.Acci_x00f3_nparalatransferenciad || 'No registrada',
        ejecucion: item.fields.escribac_x00f3_mosellev_x00f3_a || '',
        fecha: item.fields.Fechaenqueseejecut_x00f3_laacci_ || '',
        ruta: item.fields.Ruta_x0028_s_x0029_dondereposa_x,
        obs: item.fields.Observaciones || ''
      }));

    res.json({
      success: true,
      acciones: filtradoAcciones
    });
  } catch (error) {
    res.status(500).json({ success: false, detail: error.message });
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
        contract: match.fields.Title,
        supervisor: match.fields.Supervisor || 'No registrado',
        objeto: match.fields.Objetocontractual,
        nombre: match.fields.Contratista,
        cedula: match.fields.NIT_x002f_CC,
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
// RUTA: SAVE-ACTA (ESTRICTO - SIN INVENTAR COLUMNAS)
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { idSharePoint, datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;
  if (!idSharePoint) {
    return res.status(400).json({ success: false, message: 'Falta el identificador de registro.' });
  }

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    // 1. Mapeo estricto para STC_General
    const generalPayload = {
      fields: {
        Title: datosGenerales.numeroContrato,
        Supervisor: datosGenerales.supervisor,
        Objetocontractual: datosGenerales.objetoContrato,
        Dependencia: datosGenerales.dependencia, 
        Contratista: datosGenerales.nombreContratista,
        Fechadediligenciamiento: datosGenerales.isFinal ? new Date().toISOString().split('T')[0] : '',
        NIT_x002f_CC: String(datosGenerales.cedula).trim(),
        Lineamientos: datosGenerales.lineamientos || '',
        Recomendaciones: datosGenerales.recomendacionesAcciones || '',
        CorreoContratista: datosGenerales.correoContratista,
        Estado: datosGenerales.isFinal ? 'Finalizado' : 'En diligenciamiento'
      }
    };

    await axios.patch(`${graphBaseUrl}/${LIST_ID_GENERAL}/items/${idSharePoint}`, generalPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    // 2. Limpieza transaccional de registros previos en STC_Acciones para este contratista
    const resAccionesActuales = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const filasABorrar = resAccionesActuales.data.value.filter(f => String(f.fields.Observaciones).includes(`CC_${datosGenerales.cedula}`));
    
    await Promise.all(filasABorrar.map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));

    // 3. Inyección limpia en STC_Acciones según la especificación exacta
    if (acciones && acciones.length > 0) {
      for (const item of acciones) {
        const accFields = {
          Title: item.proceso, // Guarda el Proceso clave / Acción de transferencia ejecutada
          Prioridad: item.prioridad,
          Productosentrega: item.productos,
          Acci_x00f3_nparalatransferenciad: item.accionConocimiento, 
          escribac_x00f3_mosellev_x00f3_a: item.ejecucion, // Guarda Describa cómo se llevó a cabo la acción de transferencia y evidencias
          Ruta_x0028_s_x0029_dondereposa_x: item.ruta, // Guarda Ruta(s) donde reposa(n) la evidencia(s) de la acción realizada
          Observaciones: `${item.obs || 'Ninguna'} | CC_${datosGenerales.cedula}` // Llave compuesta de control
        };
        if (item.fecha && item.fecha.trim() !== "") {
          accFields.Fechaenqueseejecut_x00f3_laacci_ = item.fecha;
        }
        await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, { fields: accFields }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 4. Inyección limpia en STC_Asuntos según la especificación exacta (Solo si es finalizado)
    if (datosGenerales.isFinal && asuntos && asuntos.length > 0) {
      for (const item of asuntos) {
        const asuFields = {
          Title: item.tramite, // Guarda Asunto pendiente de trámite o en trámite
          Estado: item.estado,
          Entidad_x002f_Dependencia: item.entidad,
          Accionespendientesporrealizar: item.accionesPendientes
        };
        if (item.fecha && item.fecha.trim() !== "") {
          asuFields.Fechal_x00ed_mite = item.fecha;
        }
        await axios.post(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items`, { fields: asuFields }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 5. Inyección limpia en STC_Sistemas según la especificación exacta
    if (datosGenerales.isFinal && sistemas && sistemas.length > 0) {
      for (const item of sistemas) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items`, {
          fields: {
            Title: item.nombre, // Guarda el Sistema / aplicativo
            Usuario: item.usuario,
            Contrase_x00f1_a: item.contrasena,
            Observaciones: item.obs
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 6. Inyección limpia en STC_Directorio según la especificación exacta
    if (datosGenerales.isFinal && directorio && directorio.length > 0) {
      for (const item of directorio) {
        await axios.post(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items`, {
          fields: {
            Title: item.nombre, // Guarda Nombre
            Tel_x00e9_fono: item.tel,
            E_x002d_Mail: item.correo,
            Tipodecontacto: item.tipo,
            Entidad_x002f_Dependencia: item.entidad,
            Recomendaciones: item.reco
          }
        }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    return res.status(200).json({ success: true });
  } catch (error) {
    const apiErrorDetail = error.response?.data?.error || error.message;
    console.error("Error pormenorizado en save-acta:", JSON.stringify(apiErrorDetail));
    return res.status(500).json({ 
      success: false, 
      message: 'Rechazo de persistencia en sublistas.',
      detail: apiErrorDetail 
    });
  }
});

export default app;
