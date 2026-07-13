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

// FUNCIÓN AUXILIAR: CALCULAR PIN DE SUPERVISOR (2 PRIMEROS Y 2 ÚLTIMOS DÍGITOS)
function calcularPinSupervisor(cedula) {
  const str = String(cedula).trim();
  if (str.length < 4) return "0000"; 
  const primerosDos = str.substring(0, 2);
  const ultimosDos = str.substring(str.length - 2);
  return `${primerosDos}${ultimosDos}`;
}

// ====================================================================================
// FUNCIÓN AUXILIAR: DISPARAR CORREO DE BIENVENIDA CON EL PIN ASIGNADO Y LINK DE INGRESO
// ====================================================================================
async function enviarCorreoBienvenidaPIN(token, emailDestino, contratista, contrato, pinGenerado) {
  const url = 'https://graph.microsoft.com/v1.0/users/lina.martinez@app.gov.co/sendMail'; 
  
  // URL de la plataforma para el redireccionamiento directo del contratista
  const urlPlataforma = 'https://stc-app.vercel.app'; 

  const mailPayload = {
    message: {
      subject: "🔒 Activación de Acceso - Sistema de Transferencia de Conocimiento",
      body: {
        contentType: "HTML",
        content: `
          <div style="font-family: 'Segoe UI', Arial, sans-serif; color: #333; max-width: 600px; margin: 0 auto; border: 1px solid #e0e0e0; border-radius: 8px; padding: 24px;">
            <h2 style="color: #0056b3; border-bottom: 2px solid #0056b3; padding-bottom: 10px; margin-top: 0;">¡Hola, ${contratista.toUpperCase()}!</h2>
            <p style="font-size: 14px; line-height: 1.6;">Te informamos que has sido habilitado por parte de la <strong>Dirección Administrativa y Financiera</strong> en la plataforma institucional de la <strong>Agencia APP</strong> para registrar y estructurar tu Acta de Transferencia de Conocimiento.</p>
            
            <div style="background-color: #f8f9fa; border-left: 4px solid #ffc107; padding: 12px; margin: 18px 0; border-radius: 4px;">
              <p style="margin: 4px 0; font-size: 14px;"><strong>Contrato Referencia:</strong> ${contrato}</p>
            </div>

            <p style="font-size: 14px; line-height: 1.6;">Para garantizar la custodia, confidencialidad y reserva de tus accesos y entregables, se ha generado de forma automática un <strong>PIN Corto de Seguridad</strong> para tu cuenta:</p>
            
            <div style="text-align: center; margin: 20px 0;">
              <div style="display: inline-block; background-color: #e2e8f0; color: #0f172a; font-size: 24px; font-weight: bold; letter-spacing: 6px; padding: 12px 32px; border-radius: 6px; border: 1px dashed #94a3b8;">
                ${pinGenerado}
              </div>
            </div>

            <!-- BOTÓN INTERACTIVO ADICIONADO: Redirección directa a la plataforma -->
            <div style="text-align: center; margin: 28px 0;">
              <a href="${urlPlataforma}" target="_blank" style="background-color: #0056b3; color: #ffffff; font-size: 14px; font-weight: bold; text-decoration: none; padding: 12px 28px; border-radius: 6px; display: inline-block; box-shadow: 0 4px 6px rgba(0,86,179,0.15);">
                🚀 Ingresar a la Plataforma STC
              </a>
            </div>

            <p style="font-size: 13px; color: #475569; line-height: 1.5;">💡 <em>Nota: Para ingresar a la plataforma, digita tu número de cédula tradicional acompañado de este código de 4 dígitos. No compartas este PIN con nadie.</em></p>
            <hr style="border: 0; border-top: 1px solid #e2e8f0; margin: 24px 0;">
            <p style="font-size: 11px; color: #94a3b8; text-align: center; margin: 0;">Transferencia de Conocimiento • Dirección Administrativa y Financiera • Agencia APP</p>
          </div>
        `
      },
      toRecipients: [
        { emailAddress: { address: emailDestino } }
      ]
    }
  };

  try {
    await axios.post(url, mailPayload, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
    console.log(`Correo de notificación STC enviado con éxito a: ${emailDestino}`);
  } catch (err) {
    console.error('Error enviando correo de notificación desde Graph API:', err.response?.data || err.message);
  }
}

app.get('/', (req, res) => {
  res.send('Servidor STC operando con filtrado dinámico de supervisores y doble factor por PIN de 4 dígitos desde el buzón institucional.');
});

// ==========================================
// RUTA: CONSULTAR SECOP II (EXTRACTOR INTEGRAL)
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
// RUTA: HABILITAR CONTRATO (GENERACIÓN DINÁMICA DE PIN Y DISPARO DE CORREO AUTOMÁTICO)
// ==========================================
app.post('/api/habilitar-contrato', async (req, res) => {
  const { contrato, contratista, cedula, objeto, supervisor, cedulaSupervisor, fechaInicio, correoNotificacion } = req.body;
  if (!contrato || !cedula) {
    return res.status(400).json({ success: false, message: 'Faltan datos obligatorios.' });
  }
  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;
    
    const pinAleatorio = String(Math.floor(1000 + Math.random() * 9000));
    const destinoMail = correoNotificacion ? String(correoNotificacion).trim() : 'correo.pendiente@agenciaapp.co';

    const habilitarPayload = {
      fields: {
        Title: contrato ? String(contrato).substring(0, 255) : '', 
        Supervisor: supervisor,
        CedulaSupervisor: String(cedulaSupervisor).trim(), 
        Objetocontractual: objeto,
        Fechadeiniciodelcontrato: fechaInicio || '',
        Contratista: contratista,
        NIT_x002f_CC: String(cedula).trim(),
        Estado: 'Sin diligenciar',
        PIN_Contratista: pinAleatorio, 
        CorreoContratista: destinoMail
      }
    };

    await axios.post(`${graphBaseUrl}/${LIST_ID_GENERAL}/items`, habilitarPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    await enviarCorreoBienvenidaPIN(token, destinoMail, contratista, contrato, pinAleatorio);

    return res.status(200).json({ success: true });
  } catch (error) {
    return res.status(500).json({ success: false, detail: error.response?.data?.error || error.message });
  }
});

// ==========================================
// RUTA: MONITOREO CON VALIDACIÓN EXCLUSIVA DE PIN PARA FUNCIONARIOS (REGLA ALGORÍTMICA 2-2)
// ==========================================
app.get('/api/contratos', async (req, res) => {
  const { queryCedula, queryPin } = req.query; 
  if (!queryCedula || !queryPin) {
    return res.status(400).json({ success: false, message: "Identificación y PIN de seguridad requeridos." });
  }

  const strCedula = String(queryCedula).trim();
  const strPin = String(queryPin).trim();

  if (strCedula !== '123') {
    const pinCalculado = calcularPinSupervisor(strCedula);
    if (strPin !== pinCalculado) {
      return res.status(401).json({ success: false, message: "PIN de seguridad inválido para este perfil de supervisión." });
    }
  } else {
    if (strPin !== '2026') {
      return res.status(401).json({ success: false, message: "Contraseña de superusuario incorrecta." });
    }
  }

  try {
    const token = await getMicrosoftGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_GENERAL}/items?expand=fields`;
    const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
    
    let itemsFiltrados = response.data.value;

    if (strCedula !== '123') {
      itemsFiltrados = itemsFiltrados.filter(item => {
        const itemCedulaSuper = item.fields.CedulaSupervisor ? String(item.fields.CedulaSupervisor).trim() : '';
        return itemCedulaSuper === strCedula;
      });

      if (itemsFiltrados.length === 0) {
        return res.status(403).json({ 
          success: false, 
          message: "Acceso denegado. No figuras como supervisor activo de ningún contrato habilitado." 
        });
      }
    }
    
    const listaFormateada = itemsFiltrados.map(item => ({
      idSharePoint: item.id,
      contract: item.fields.Title,
      boss: item.fields.Supervisor,
      cedulaSupervisor: item.fields.CedulaSupervisor || '',
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
// RUTA: OBTENER DETALLES HIJOS
// ==========================================
app.get('/api/obtener-detalles-hijos', async (req, res) => {
  const { cedula } = req.query;
  if (!cedula) return res.status(400).json({ success: false, message: "Falta la cédula." });

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;
    const strCedula = String(cedula).trim();

    const resAcciones = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const acciones = resAcciones.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => {
        let fechaLimpia = item.fields.Fechaenqueseejecut_x00f3_laacci_ || '';
        if (fechaLimpia.includes('T')) { fechaLimpia = fechaLimpia.split('T')[0]; }
        return {
          proceso: item.fields.Title,
          prioridad: item.fields.Prioridad,
          productos: item.fields.Productosentrega,
          accionConocimiento: item.fields.Acci_x00f3_nparalatransferenciad || 'No registrada',
          ejecucion: item.fields.Describac_x00f3_mosellev_x00f3_a || '',
          fecha: fechaLimpia,
          ruta: item.fields.Ruta_x0028_s_x0029_dondereposa_x || '',
          obs: item.fields.Observaciones || ''
        };
      });

    const resAsuntos = await axios.get(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const asuntos = resAsuntos.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        tramite: item.fields.Title,
        estado: item.fields.Estado,
        entidad: item.fields.Entidad_x002f_Dependencia,
        accionesPendientes: item.fields.Accionespendientesporrealizar,
        fecha: item.fields.Fechal_x00ed_mite ? item.fields.Fechal_x00ed_mite.split('T')[0] : ''
      }));

    const resSistemas = await axios.get(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    const sistemas = resSistemas.data.value
      .filter(item => String(item.fields.CedulaRelacion).trim() === strCedula)
      .map(item => ({
        nombre: item.fields.Title,
        usuario: item.fields.Usuario,
        contrasena: item.fields.Contrase_x00f1_a,
        obs: item.fields.Observaciones || ''
      }));

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
// RUTA: LOGIN CONTRATISTA CON FILTRADO Y VALIDACIÓN DOUBLE FACTOR POR PIN
// ==========================================
app.get('/api/login-contratista', async (req, res) => {
  const { cedula, pin } = req.query;
  if (!cedula || !pin) {
    return res.status(400).json({ success: false, message: "Faltan credenciales obligatorias para autenticar." });
  }
  try {
    const token = await getMicrosoftGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_GENERAL}/items?expand=fields`;
    const response = await axios.get(url, { headers: { 'Authorization': `Bearer ${token}` } });
    
    const match = response.data.value.find(item => {
      const nitFila = item.fields.NIT_x002f_CC ? String(item.fields.NIT_x002f_CC).trim() : '';
      return nitFila === String(cedula).trim();
    });
    
    if (match) {
      const dbPin = match.fields.PIN_Contratista ? String(match.fields.PIN_Contratista).trim() : '';
      
      if (dbPin !== String(pin).trim()) {
        return res.status(401).json({ success: false, incorrectPin: true, message: "PIN de seguridad incorrecto. Verifique su correo institucional." });
      }

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
// RUTA: REABRIR ACTA DESDE SUPERVISIÓN
// ==========================================
app.post('/api/reabrir-acta', async (req, res) => {
  const { idSharePoint } = req.body;
  if (!idSharePoint) return res.status(400).json({ success: false, message: 'Falta el ID del acta.' });
  try {
    const token = await getMicrosoftGraphToken();
    const url = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID_GENERAL}/items/${idSharePoint}`;
    const payload = { fields: { Estado: 'En diligenciamiento' } };
    await axios.patch(url, payload, { headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' } });
    return res.json({ success: true });
  } catch (error) { return res.status(500).json({ success: false, detail: error.message }); }
});

// ==========================================
// RUTA: SAVE-ACTA
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { idSharePoint, datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;
  if (!idSharePoint) return res.status(400).json({ success: false, message: 'Falta el ID de registro.' });

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;
    const strCedula = String(datosGenerales.cedula).trim();

    const generalPayload = {
      fields: {
        Title: datosGenerales.numeroContrato ? String(datosGenerales.numeroContrato).substring(0, 255) : '',
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

    const rAcc = await axios.get(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rAcc.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (acciones && acciones.length > 0) {
      for (const item of acciones) {
        const f = { 
          Title: item.proceso ? String(item.proceso).substring(0, 255) : '', 
          CedulaRelacion: strCedula, 
          Prioridad: item.prioridad, 
          Productosentrega: item.productos, 
          Acci_x00f3_nparalatransferenciad: item.accionConocimiento || 'No registrada', 
          Describac_x00f3_mosellev_x00f3_a: item.ejecucion, 
          Ruta_x0028_s_x0029_dondereposa_x: String(item.ruta).trim(), 
          Observaciones: item.obs 
        };
        if (item.fecha && item.fecha.trim() !== "") f.Fechaenqueseejecut_x00f3_laacci_ = item.fecha;
        await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    const rAsu = await axios.get(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rAsu.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (asuntos && asuntos.length > 0) {
      for (const item of asuntos) {
        const f = { Title: item.tramite ? String(item.tramite).substring(0, 255) : '', CedulaRelacion: strCedula, Estado: item.estado, Entidad_x002f_Dependencia: item.entidad, Accionespendientesporrealizar: item.accionesPendientes };
        if (item.fecha && item.fecha.trim() !== "") f.Fechal_x00ed_mite = item.fecha;
        await axios.post(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    const rSis = await axios.get(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rSis.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (sistemas && sistemas.length > 0) {
      for (const item of sistemas) {
        const f = { Title: item.nombre ? String(item.nombre).substring(0, 255) : '', CedulaRelacion: strCedula, Usuario: item.usuario, Contrase_x00f1_a: item.contrasena, Observaciones: item.obs };
        await axios.post(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    const rDir = await axios.get(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items?expand=fields`, { headers: { 'Authorization': `Bearer ${token}` } });
    await Promise.all(rDir.data.value.filter(f => String(f.fields.CedulaRelacion).trim() === strCedula).map(f => 
      axios.delete(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items/${f.id}`, { headers: { 'Authorization': `Bearer ${token}` } })
    ));
    if (directorio && directorio.length > 0) {
      for (const item of directorio) {
        const f = { Title: item.nombre ? String(item.nombre).substring(0, 255) : '', CedulaRelacion: strCedula, Tel_x00e9_fono: item.tel, E_x002d_Mail: item.correo, Tipodecontacto: item.tipo, Entidad_x002f_Dependencia: item.entidad, Recommendations: item.reco };
        await axios.post(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items`, { fields: f }, { headers: { 'Authorization': `Bearer ${token}` } });
      }
    }

    return res.status(200).json({ success: true });
  } catch (error) {
    const apiErrorDetail = error.response?.data?.error || error.message;
    return res.status(500).json({ success: false, detail: apiErrorDetail });
  }
});

export default app;
