require('dotenv').config();
const express = require('express');
const cors = require('cors');
const XLSX = require('xlsx');
const multer = require('multer');
const { createClient } = require('@supabase/supabase-js');

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.static('public'));

const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });
const supabase = createClient(process.env.SUPABASE_URL, process.env.SUPABASE_KEY);

function excelSerialToDate(serial) {
  if (!serial || typeof serial !== 'number') return null;
  const epoch = new Date(1900, 0, 1);
  const date = new Date(epoch.getTime() + (serial - 1) * 86400000);
  return date.toISOString().split('T')[0];
}

function parseExcelART(data) {
  const empresa = data[0]?.[0] || '';
  const nit = data[1]?.[0] || '';
  const titulo = data[3]?.[0] || '';
  const fechaDesde = excelSerialToDate(data[4]?.[1]);
  const fechaHasta = excelSerialToDate(data[4]?.[3]);
  const productos = [];
  let i = 10;

  while (i < data.length) {
    const row = data[i];
    if (!row || !row[0]) { i++; continue; }
    const texto = String(row[0]).trim();
    if (texto === 'TOTAL' || texto.startsWith('***') || texto === 'TOTALES GENERALES') { i++; continue; }
    if (/^\d{6}\s/.test(texto)) {
      const parts = texto.match(/^(\d+)\s+(.*)/);
      const codigo = parts ? parts[1] : '';
      const nombre = parts ? parts[2].trim() : texto;
      let j = i + 1;
      while (j < data.length) {
        const d = data[j];
        if (d && d[1] !== undefined && d[1] !== null && d[1] !== '') {
          const cantidades = parseFloat(d[1]) || 0;
          const valores = parseFloat(d[2]) || 0;
          const porcentaje = parseFloat(d[3]) || 0;
          if (cantidades !== 0 || valores !== 0) {
            productos.push({ producto_codigo: codigo, producto_nombre: nombre, cantidades, valores, porcentaje });
          }
          break;
        }
        if (d && d[0] && String(d[0]).trim() === 'TOTAL') break;
        j++;
      }
    }
    i++;
  }

  return {
    empresa: empresa.replace(/\s+/g, ' ').trim(),
    nit: nit.replace(/\s+/g, ' ').trim(),
    titulo: titulo.replace(/\s+/g, ' ').trim(),
    fechaDesde, fechaHasta, tipo: 'ART', productos
  };
}

function parseExcelVENTAS(data) {
  const empresa = data[0]?.[0] || '';
  const nit = data[1]?.[0] || '';
  const titulo = data[3]?.[0] || '';
  const fechaDesde = excelSerialToDate(data[4]?.[1]);
  const fechaHasta = excelSerialToDate(data[4]?.[3]);
  const ventas = [];
  let zonaActual = '';
  let i = 10;

  while (i < data.length) {
    const row = data[i];
    if (!row || !row[0]) { i++; continue; }
    const texto = String(row[0]).trim();
    if (texto === 'TOTAL' || texto.startsWith('***') || texto === 'TOTALES GENERALES') { i++; continue; }
    if (texto.startsWith('ZONA ') || texto === ' SIN ZONA') { zonaActual = texto.trim(); i++; continue; }
    const clienteMatch = texto.match(/^(.+?)\s+Id:(\d+)/);
    if (clienteMatch) {
      const clienteNombre = clienteMatch[1].trim();
      const clienteId = clienteMatch[2];
      let j = i + 1;
      while (j < data.length) {
        const d = data[j];
        if (d && d[1] !== undefined && d[1] !== null && d[1] !== '') {
          const cantidades = parseFloat(d[1]) || 0;
          const valores = parseFloat(d[2]) || 0;
          const porcentaje = parseFloat(d[3]) || 0;
          if (cantidades !== 0 || valores !== 0) {
            ventas.push({ zona: zonaActual, cliente_id: clienteId, cliente_nombre: clienteNombre, cantidades, valores, porcentaje });
          }
          break;
        }
        if (d && d[0] && String(d[0]).trim() === 'TOTAL') break;
        j++;
      }
    }
    i++;
  }

  return {
    empresa: empresa.replace(/\s+/g, ' ').trim(),
    nit: nit.replace(/\s+/g, ' ').trim(),
    titulo: titulo.replace(/\s+/g, ' ').trim(),
    fechaDesde, fechaHasta, tipo: 'VENTAS', ventas
  };
}

async function saveInforme(tipo, parsed) {
  const { data: informe, error: errInf } = await supabase
    .from('informes')
    .insert({
      empresa: parsed.empresa,
      nit: parsed.nit,
      titulo: parsed.titulo,
      fecha_desde: parsed.fechaDesde,
      fecha_hasta: parsed.fechaHasta,
      tipo
    })
    .select()
    .single();

  if (errInf) throw errInf;

  const rows = tipo === 'ART' ? parsed.productos.map(p => ({ informe_id: informe.id, ...p })) :
                                     parsed.ventas.map(v => ({ informe_id: informe.id, ...v }));
  const table = tipo === 'ART' ? 'detalle_art' : 'detalle_ventas';

  if (rows.length > 0) {
    const { error: errDet } = await supabase.from(table).insert(rows);
    if (errDet) throw errDet;
  }

  return { id: informe.id, tipo, registros: rows.length };
}

app.post('/api/upload', upload.fields([
  { name: 'art', maxCount: 1 },
  { name: 'ventas', maxCount: 1 }
]), async (req, res) => {
  try {
    if (!req.files || !req.files.art || !req.files.ventas) {
      return res.status(400).json({ error: 'Debes subir ambos archivos: uno de ART y uno de VENTAS' });
    }

    const artWb = XLSX.read(req.files.art[0].buffer, { type: 'buffer' });
    const ventasWb = XLSX.read(req.files.ventas[0].buffer, { type: 'buffer' });

    const artData = parseExcelART(XLSX.utils.sheet_to_json(artWb.Sheets[artWb.SheetNames[0]], { header: 1 }));
    const ventasData = parseExcelVENTAS(XLSX.utils.sheet_to_json(ventasWb.Sheets[ventasWb.SheetNames[0]], { header: 1 }));

    const resultadoART = await saveInforme('ART', artData);
    const resultadoVENTAS = await saveInforme('VENTAS', ventasData);

    res.json({ mensaje: 'Archivos cargados correctamente', informes: [resultadoART, resultadoVENTAS] });
  } catch (error) {
    console.error('Error al subir:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/informes', async (req, res) => {
  const { data, error } = await supabase
    .from('informes')
    .select('*')
    .order('fecha_carga', { ascending: false });

  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.get('/api/informes/:id/art', async (req, res) => {
  const { data, error } = await supabase
    .from('detalle_art')
    .select('*')
    .eq('informe_id', req.params.id)
    .order('valores', { ascending: false });

  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.get('/api/informes/:id/ventas', async (req, res) => {
  const { data, error } = await supabase
    .from('detalle_ventas')
    .select('*')
    .eq('informe_id', req.params.id)
    .order('valores', { ascending: false });

  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.get('/api/exportar/:id', async (req, res) => {
  try {
    const { data: informe, error: errInf } = await supabase
      .from('informes')
      .select('*')
      .eq('id', req.params.id)
      .single();

    if (errInf) throw errInf;
    if (!informe) return res.status(404).json({ error: 'Informe no encontrado' });

    const wb = XLSX.utils.book_new();
    const headers = ['Código', 'Nombre/Cliente', 'Zona', 'Cantidades', 'Valores', '% Part'];
    const table = informe.tipo === 'ART' ? 'detalle_art' : 'detalle_ventas';
    const { data: detalle } = await supabase.from(table).select('*').eq('informe_id', informe.id);

    const rows = (detalle || []).map(d =>
      informe.tipo === 'ART'
        ? [d.producto_codigo, d.producto_nombre, '-', d.cantidades, d.valores, d.porcentaje]
        : [d.cliente_id, d.cliente_nombre, d.zona, d.cantidades, d.valores, d.porcentaje]
    );

    const wsData = [
      [informe.empresa], [informe.nit], [],
      [informe.titulo], [`Desde: ${informe.fecha_desde}`, `Hasta: ${informe.fecha_hasta}`],
      [], headers, ...rows
    ];

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    ws['!cols'] = informe.tipo === 'ART'
      ? [{ wch: 10 }, { wch: 60 }, { wch: 8 }, { wch: 12 }, { wch: 15 }, { wch: 10 }]
      : [{ wch: 15 }, { wch: 45 }, { wch: 20 }, { wch: 12 }, { wch: 15 }, { wch: 10 }];

    XLSX.utils.book_append_sheet(wb, ws, informe.tipo);

    const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    const filename = `Informe_${informe.tipo}_${informe.fecha_desde}_${informe.fecha_hasta}.xlsx`;

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
  } catch (error) {
    console.error('Error al exportar:', error);
    res.status(500).json({ error: error.message });
  }
});

app.delete('/api/informes/:id', async (req, res) => {
  const { error } = await supabase.from('informes').delete().eq('id', req.params.id);
  if (error) return res.status(500).json({ error: error.message });
  res.json({ mensaje: 'Informe eliminado' });
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
