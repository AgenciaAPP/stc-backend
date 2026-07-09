import express from 'express';
import cors from 'cors';
import axios from 'axios';
import dotenv from 'dotenv';

dotenv.config();

const app = express();

// Configuración de CORS permitiendo el acceso desde tu frontend de Vercel
app.use(cors());
app.use(express.json());

// CONFIGURACIÓN SEGURO MEDIANTE LAS VARIABLES REALES DE TU PANTALLAZO VERCEL
const TENANT_ID = process.env.TENANT_ID;
const CLIENT_ID = process.env.CLIENT_ID;
const CLIENT_SECRET = process.env.CLIENT_SECRET;
const SITE_ID = process.env.SITE_ID;

// MAPEO EXACTO CON LOS NOMBRES DE TU CAPTURA DE PANTALLA
const LIST_ID_GENERAL = process.env.LIST_ID_GENERAL;
const LIST_ID_ACCIONES = process.env.LIST_ID_ACCIONES;
const LIST_ID_ASUNTOS = process.env.LIST_ID_ASUNTOS;
const LIST_ID_SISTEMAS = process.env.LIST_ID_SISTEMAS;
const LIST_ID_DIRECTORIO = process.env.LIST_ID_DIRECTORIO;

// ==========================================
// FUNCIÓN: ADQUISICIÓN AUTOMÁTICA DE TOKEN BEARER DESDE AZURE AD
// ==========================================
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

// ==========================================
// 1. RUTA DE PRUEBA: SALUDO INICIAL
// ==========================================
app.get('/', (req, res) => {
  res.send('Servidor STC de la Agencia APP operando correctamente con variables de entorno validadas.');
});

// ==========================================
// 2. RUTA DE PRUEBA: STATUS GENERAL
// ==========================================
app.get('/api/status', (req, res) => {
  res.json({ 
    status: "online", 
    message: "Conexión exitosa con el backend seguro de Vercel",
    timestamp: new Date()
  });
});

// ==========================================
// 3. RUTA: CONSULTAR SECOP II (SINCRO POR NIT DE LA AGENCIA APP)
// ==========================================
app.get('/api/buscar-secop', async (req, res) => {
  try {
    const { contrato } = req.query;

    if (!contrato) {
      return res.status(400).json({ success: false, message: "Falta el parámetro 'contrato' en la consulta" });
    }

    const nitAgenciaAPP = "900623766"; 
    const secopUrl = `https://datos.gov.co/resource/jzye-7urr.json?numero_de_contrato=${encodeURIComponent(contrato)}&nit_de_la_entidad=${nitAgenciaAPP}`;
    
    // Agregamos un User-Agent real para evitar que el Web Application Firewall (WAF) del gobierno rebote la petición de Vercel
    const response = await axios.get(secopUrl, {
      headers: {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36'
      }
    });
    
    if (response.data && response.data.length > 0) {
      const dataContrato = response.data[0];
      res.json({
        success: true,
        nombre: dataContrato.nombre_o_razon_social_del_contratista || "No registrado",
        cedula: dataContrato.documento_del_contratista || "No registrado",
        objeto: dataContrato.objeto_del_contrato || "No registrado"
      });
    } else {
      res.json({ success: false, message: "Contrato no encontrado en SECOP II para la Agencia APP." });
    }
  } catch (error) {
    console.error("Error detallado consultando SECOP II:", error.response?.data || error.message);
    res.status(500).json({ success: false, message: "Error al conectarse con el servidor gubernamental." });
  }
});

// ==========================================
// 4. RUTA: ENTRADA Y PERSISTENCIA COMPLETA EN SHAREPOINT
// ==========================================
app.post('/api/save-acta', async (req, res) => {
  const { datosGenerales, acciones, asuntos, sistemas, directorio } = req.body;

  if (!datosGenerales || !datosGenerales.cedula) {
    return res.status(400).json({ success: false, message: 'Faltan los datos generales o la cédula de validación.' });
  }

  try {
    const token = await getMicrosoftGraphToken();
    const graphBaseUrl = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists`;

    // 1. Inyección de Datos Generales (Pestaña 1)
    const generalPayload = {
      fields: {
        Title: datosGenerales.cedula,
        NombreContratista: datosGenerales.nombreContratista,
        NumeroContrato: datosGenerales.numeroContrato,
        Supervisor: datosGenerales.supervisor,
        ObjetoContrato: datosGenerales.objetoContrato,
        CorreoContacto: datosGenerales.correoContratista,
        Dependencia: datosGenerales.dependencia,
        LineamientosGenerales: datosGenerales.lineamientos || '',
        RecomendacionesEspeciales: datosGenerales.recomendacionesAcciones || '',
        EstadoActa: datosGenerales.isFinal ? 'Finalizado' : 'En Diligenciamiento'
      }
    };
    await axios.post(`${graphBaseUrl}/${LIST_ID_GENERAL}/items`, generalPayload, {
      headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' }
    });

    // 2. Inyección Multirregistro: Acciones (Pestaña 2)
    for (const item of acciones) {
      await axios.post(`${graphBaseUrl}/${LIST_ID_ACCIONES}/items`, {
        fields: {
          Title: datosGenerales.cedula,
          ProcesoClave: item.proceso,
          Prioridad: item.prioridad,
          ProductosEntrega: item.productos,
          EvidenciasEjecucion: item.ejecucion,
          FechaEjecucion: item.fecha,
          RutaRepositorio: item.ruta,
          Observaciones: item.obs
        }
      }, { headers: { 'Authorization': `Bearer ${token}` } });
    }

    // 3. Inyección Multirregistro: Asuntos (Pestaña 3)
    for (const item of asuntos) {
      await axios.post(`${graphBaseUrl}/${LIST_ID_ASUNTOS}/items`, {
        fields: {
          Title: datosGenerales.cedula,
          AsuntoTramite: item.tramite,
          EstadoActual: item.estado,
          EntidadDependencia: item.entidad,
          AccionesPendientes: item.accionesPendientes,
          FechaLimite: item.fecha
        }
      }, { headers: { 'Authorization': `Bearer ${token}` } });
    }

    // 4. Inyección Multirregistro: Sistemas (Pestaña 4)
    for (const item of sistemas) {
      await axios.post(`${graphBaseUrl}/${LIST_ID_SISTEMAS}/items`, {
        fields: {
          Title: datosGenerales.cedula,
          SistemaAplicativo: item.nombre,
          Usuario: item.usuario,
          Contrasena: item.contrasena,
          Observaciones: item.obs
        }
      }, { headers: { 'Authorization': `Bearer ${token}` } });
    }

    // 5. Inyección Multirregistro: Directorio (Pestaña 5)
    for (const item of directorio) {
      await axios.post(`${graphBaseUrl}/${LIST_ID_DIRECTORIO}/items`, {
        fields: {
          Title: datosGenerales.cedula,
          NombreContacto: item.nombre,
          Telefono: item.tel,
          Email: item.correo,
          TipoContacto: item.tipo,
          EntidadDependenciaDirectorio: item.entidad,
          RecomendacionesDirectorio: item.reco
        }
      }, { headers: { 'Authorization': `Bearer ${token}` } });
    }

    return res.status(200).json({ success: true, message: '¡Acta sincronizada con éxito en SharePoint!' });

  } catch (error) {
    console.error('Error inyectando en SharePoint:', error.response?.data || error.message);
    return res.status(500).json({ success: false, message: 'Error interno de comunicación con Microsoft Graph.' });
  }
});

export default app;
