import express from 'express';
import cors from 'cors';
import axios from 'axios';

const app = express();

// Configuración de CORS permitiendo el acceso desde tu frontend de Vercel
app.use(cors());
app.use(express.json());

// ==========================================
// 1. RUTA DE PRUEBA: SALUDO INICIAL
// ==========================================
app.get('/', (req, res) => {
  res.send('Servidor STC de la Agencia APP funcionando correctamente.');
});

// ==========================================
// 2. RUTA DE PRUEBA: CONEXIÓN DE BASE DE DATOS / STATUS
// ==========================================
app.get('/api/status', (req, res) => {
  res.json({ 
    status: "online", 
    message: "Conexión exitosa con el backend en Vercel",
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
      return res.status(400).json({ status: "error", message: "Falta el parámetro 'contrato' en la consulta" });
    }

    // Filtramos simultáneamente por la referencia del contrato Y por el NIT numérico de la Agencia APP (900623766)
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

    // Mapeamos las variables con los nombres de columna confirmados del JSON
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
        supervisor: contratoData.nombre_supervisor && contratoData.nombre_supervisor !== "No definido" ? contratoData.nombre_supervisor : ""
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

// ==========================================
// 4. RUTA: GUARDAR DATOS GENERALES (MOCK DE ENVÍO A SHAREPOINT)
// ==========================================
app.post('/api/guardar-general', (req, res) => {
  try {
    const datosRecibidos = req.body;
    console.log("Datos para guardar en SharePoint:", datosRecibidos);

    // Aquí irá más adelante el código de integración con Microsoft Graph API / SharePoint
    res.json({
      status: "success",
      message: "Información contractual almacenada correctamente en el repositorio temporal de SharePoint."
    });
  } catch (error) {
    res.status(500).json({ status: "error", message: "Error al guardar en el servidor" });
  }
});

export default app;
