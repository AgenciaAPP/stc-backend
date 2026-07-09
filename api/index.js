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
  res.send('Servidor STC operando con filtrado por CedulaRelacion en todas las sublistas.');
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
// RUTA: OBTENER DETALLES HIJOS FILTRADOS POR CEDULARELACION
// ==========================================
app.get('/api/obtener-detalles-hijos', async (req, res) => {
  const { cedula } = req.query;
  if (!cedula) return res.status(400).json({ success: false, message: "Falta la cédula." });

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;
    const strCedula = String(cedula).trim();

    // 1. Cargar y filtrar Acciones
    const resAcciones = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const acciones = resAcciones.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        proceso: item.fields.Title,
        prioridad: item.fields.Prioridad,
        productos: item.fields.Productosentrega,
        accionConocimiento: item.fields.Acci_x00f3_nparalatransferenciad || 'No registrada',
        ejecucion: item.fields.escribac_x00f3_mosellev_x00f3_a || '',
        fecha: item.fields.Fechaenqueseejecut_x00f3_laacci_ || '',
        ruta: item.fields.Ruta_x0028_s_x0029_dondereposa_x,
        obs: item.fields.Observaciones || ''
      }));

    // 2. Cargar y filtrar Asuntos
    const resAsuntos = await axios.get(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const asuntos = resAsuntos.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        tramite: item.fields.Title,
        estado: item.fields.Estado,
        entidad: item.fields.Entidad_x002f_Dependencia,
        accionesPendientes: item.fields.Accionespendientesporrealizar,
        fecha: item.fields.Fechal_x00ed_mite || ''
      }));

    // 3. Cargar y filtrar Sistemas
    const resSistemas = await axios.get(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const sistemas = resSistemas.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        nombre: item.fields.Title,
        usuario: item.fields.Usuario,
        contrasena: item.fields.Contrase_x00f1_a,
        obs: item.fields.Observaciones || ''
      }));

    // 4. Cargar y filtrar Directorio
    const resDirectorio = await axios.get(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const directorio = resDirectorio.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        nombre: item.fields.Title,
        tel: item.fields.Tel_x00e9_fono,
        correo: item.fields.E_x002d_Mail,
        tipo: item.fields.Tipodecontacto,
        entidad: item.fields.Entidad_x002f_Dependencia,
        reco: item.fields.Recomendaciones || ''
      }));

    res.json({ success: true, acciones, asuntos, sistemas, directorio });
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
// RUTA: SAVE-ACTA CON PERSISTENCIA Y RELACIÓN TRANSACCIONAL
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { idSharePoint, datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;
  if (!idSharePoint) return res.status(400).json({ success: false, message: 'Falta el ID de registro.' });

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;
    const strCedula = String(datosGenerales.cedula).trim();

    // 1. Actualizar STC_General
    const generalPayload = {
      fields: {
        Title: datosGenerales.numeroContrato,
        Supervisor: datosGenerales.supervisor,
        Objetocontractual: datosGenerales.objetoContrato,
        Dependencia: datosGenerales.dependencia, 
        Contratista: datosGenerales.nombreContratista,
        Fechadediligenciamiento: datosGenerales.isFinal ? new Date().toISOString().split('T')[0] : '',
        NIT_x002f_CC: strCedula,
        Lineamientos: datosGenerales.lineamientos || '',
        Recomendaciones: datosGenerales.recomendacionesAcciones || '',
        CorreoContratista: datosGenerales.correoContratista,
        Estado: datosGenerales.isFinal ? 'Finalizado' : 'En diligenciamiento'
      }
    };
    await axios.patch(`${graphBaseUrl}/${LIST_ID_GENERAL}/items/${idSharePoint}`, generalPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    // 2. Transacción de Limpieza y Guardado para ACCIONES
    const rAcc = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rAcc.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (acciones && acciones.length > 0) {
      for (const item of acciones) {
        const f = {
          Title: item.proceso, CedulaRelacion: strCedula, Prioridad: item.prioridad, Productosentrega: item.productos,
          Acci_x00f3_nparalatransferenciad: item.accionConocimiento, escribac_x00f3_mosellev_x00f3_a: item.ejecucion, Ruta_x0028_s_x0029_dondereposa_x: item.ruta, Observaciones: item.obs
        };
        if (item.fecha && item.fecha.trim() !== "") f.Fechaenqueseejecut_x00f3_laacci_ = item.fecha;
        await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 3. Transacción de Limpieza y Guardado para ASUNTOS
    const rAsu = await axios.get(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rAsu.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (asuntos && asuntos.length > 0) {
      for (const item of asuntos) {
        const f = { Title: item.tramite, CedulaRelacion: strCedula, Estado: item.estado, Entidad_x002f_Dependencia: item.entidad, Accionespendientesporrealizar: item.accionesPendientes };
        if (item.fecha && item.fecha.trim() !== "") f.Fechal_x00ed_mite = item.fecha;
        await axios.post(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 4. Transacción de Limpieza y Guardado para SISTEMAS
    const rSis = await axios.get(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rSis.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (sistemas && sistemas.length > 0) {
      for (const item of sistemas) {
        const f = { Title: item.nombre, CedulaRelacion: strCedula, Usuario: item.usuario, Contrase_x00f1_a: item.contrasena, Observaciones: item.obs };
        await axios.post(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    // 5. Transacción de Limpieza y Guardado para DIRECTORIO
    const rDir = await axios.get(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rDir.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (directorio && directorio.length > 0) {
      for (const item of directorio) {
        const f = { Title: item.nombre, CedulaRelacion: strCedula, Tel_x00e9_fono: item.tel, E_x002d_Mail: item.correo, Tipodecontacto: item.tipo, Entidad_x002f_Dependencia: item.entidad, Recomendaciones: item.reco };
        await axios.post(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    return res.status(200).json({ success: true });
  } catch (error) {
    return res.status(500).json({ success: false, detail: error.response?.data?.error || error.message });
  }
});

export default app;
