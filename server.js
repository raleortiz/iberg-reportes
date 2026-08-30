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

const MESES_ES = {
  'ene': '01', 'feb': '02', 'mar': '03', 'abr': '04', 'may': '05', 'jun': '06',
  'jul': '07', 'ago': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dic': '12'
};

function fechaTextoToDate(valor) {
  if (!valor) return null;
  if (typeof valor === 'number') return excelSerialToDate(valor);
  const str = String(valor).trim();
  const m = str.match(/^([A-Za-z]{3})-(\d{1,2})-(\d{4})$/);
  if (m) {
    const anio = m[3];
    const mes = MESES_ES[m[1].toLowerCase()];
    const dia = String(parseInt(m[2], 10)).padStart(2, '0');
    if (mes) return `${anio}-${mes}-${dia}`;
  }
  const iso = str.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return str;
  return null;
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

function parseRecaudo(data) {
  const empresa = data[0]?.[0] || '';
  const nit = data[1]?.[0] || '';
  const titulo = data[3]?.[0] || '';
  const periodo = data[4]?.[0] || '';
  const registros = [];
  let i = 7;

  while (i < data.length) {
    const r = data[i];
    if (!r) { i++; continue; }

    const fechaRecaudo = r[1];
    const numero = r[2];
    const fechaFactura = r[3];
    const factura = r[4];
    const recaudo = r[5];
    const noExterno = r[6];
    const descuentos = r[7];
    const retencion = r[8];
    const total = r[9];
    const nitCliente = r[10];
    const nombre = r[11];

    const tieneDatos = numero !== undefined && numero !== null && String(numero).trim() !== '' &&
                       factura !== undefined && factura !== null && String(factura).trim() !== '';

    if (tieneDatos) {
      const claveUnica = `${String(numero).trim()}|${String(factura).trim()}`;
      registros.push({
        clave_unica: claveUnica,
        fecha_recaudo: fechaTextoToDate(fechaRecaudo),
        numero: String(numero).trim(),
        fecha_factura: fechaTextoToDate(fechaFactura),
        factura: String(factura).trim(),
        recaudo: parseFloat(recaudo) || 0,
        no_externo: noExterno !== undefined && noExterno !== null ? String(noExterno).trim() : '',
        descuentos: parseFloat(descuentos) || 0,
        retencion: parseFloat(retencion) || 0,
        total: parseFloat(total) || 0,
        nit: String(nitCliente || '').trim(),
        nombre_cliente: String(nombre || '').trim()
      });
    }
    i++;
  }

  return {
    empresa: empresa.replace(/\s+/g, ' ').trim(),
    nit: nit.replace(/\s+/g, ' ').trim(),
    titulo: titulo.replace(/\s+/g, ' ').trim(),
    periodo: periodo.replace(/\s+/g, ' ').trim(),
    registros
  };
}

app.post('/api/upload/recaudo', upload.single('recaudo'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No se recibio ningun archivo de recaudo' });
    }

    const wb = XLSX.read(req.file.buffer, { type: 'buffer' });
    const recaudoData = parseRecaudo(XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1 }));

    if (recaudoData.registros.length === 0) {
      return res.status(400).json({ error: 'No se encontraron registros en el archivo' });
    }

    const claves = recaudoData.registros.map(r => r.clave_unica);
    const { data: existentes, error: errExist } = await supabase
      .from('recaudo')
      .select('clave_unica, numero, factura, nit, nombre_cliente, total')
      .in('clave_unica', claves);

    if (errExist) throw errExist;

    const clavesExistentes = new Set((existentes || []).map(e => e.clave_unica));

    const repetidosEnArchivo = {};
    const vistosArchivo = new Set();
    for (const r of recaudoData.registros) {
      if (vistosArchivo.has(r.clave_unica)) {
        if (!repetidosEnArchivo[r.clave_unica]) repetidosEnArchivo[r.clave_unica] = [];
        repetidosEnArchivo[r.clave_unica].push(r);
      }
      vistosArchivo.add(r.clave_unica);
    }

    const enBase = recaudoData.registros
      .filter(r => clavesExistentes.has(r.clave_unica))
      .map(r => {
        const existente = (existentes || []).find(e => e.clave_unica === r.clave_unica);
        return { ...r, ya_en_base: true, ...(existente || {}) };
      });

    function agruparPorClave(items) {
      const grupos = {};
      for (const item of items) {
        if (!grupos[item.clave_unica]) grupos[item.clave_unica] = [];
        grupos[item.clave_unica].push(item);
      }
      return Object.values(grupos).map(grupo => ({
        clave_unica: grupo[0].clave_unica,
        numero: grupo[0].numero,
        factura: grupo[0].factura,
        registros: grupo.map(r => ({
          nit: r.nit,
          nombre_cliente: r.nombre_cliente,
          fecha_recaudo: r.fecha_recaudo,
          fecha_factura: r.fecha_factura,
          recaudo: r.recaudo,
          total: r.total,
          ya_en_base: !!r.ya_en_base
        }))
      }));
    }

    const enBaseAgrupado = agruparPorClave(enBase);
    const repetidosEnArchivoAgrupado = agruparPorClave(Object.values(repetidosEnArchivo).flat());

    const totalDuplicados =
      enBaseAgrupado.reduce((s, g) => s + g.registros.length, 0) +
      repetidosEnArchivoAgrupado.reduce((s, g) => s + g.registros.length, 0);

    if (enBaseAgrupado.length > 0 || repetidosEnArchivoAgrupado.length > 0) {
      return res.status(409).json({
        error: 'Se encontraron registros duplicados. La carga no se realizo.',
        en_base: enBaseAgrupado,
        repetidos_en_archivo: repetidosEnArchivoAgrupado,
        total_duplicados: totalDuplicados
      });
    }

    const registrosUnicos = Object.values(
      recaudoData.registros.reduce((acc, r) => { acc[r.clave_unica] = r; return acc; }, {})
    );

    const { error: errInsert } = await supabase.from('recaudo').insert(registrosUnicos);
    if (errInsert) throw errInsert;

    res.json({
      mensaje: 'Recaudos cargados correctamente',
      registros: registrosUnicos.length
    });
  } catch (error) {
    console.error('Error al cargar recaudo:', error);
    res.status(500).json({ error: error.message });
  }
});

app.get('/api/recaudos', async (req, res) => {
  const { data, error } = await supabase
    .from('recaudo')
    .select('*')
    .order('fecha_recaudo', { ascending: false });

  if (error) return res.status(500).json({ error: error.message });
  res.json(data);
});

app.get('/api/recaudos/clientes', async (req, res) => {
  const { data, error } = await supabase
    .from('recaudo')
    .select('nombre_cliente')
    .not('nombre_cliente', 'is', null)
    .neq('nombre_cliente', '');

  if (error) return res.status(500).json({ error: error.message });

  const clientes = [...new Set((data || []).map(d => d.nombre_cliente))]
    .filter(Boolean)
    .sort((a, b) => a.localeCompare(b, 'es'));

  res.json(clientes);
});

app.get('/api/recaudos/dashboard', async (req, res) => {
  try {
    const { desde, hasta, clientes, agrupar } = req.query;

    let query = supabase.from('recaudo').select('*');

    if (desde) query = query.gte('fecha_recaudo', desde);
    if (hasta) query = query.lte('fecha_recaudo', hasta);

    const { data, error } = await query;

    if (error) throw error;

    let filtrados = data || [];

    if (clientes) {
      const lista = clientes.split(',');
      filtrados = filtrados.filter(r => lista.includes(r.nombre_cliente));
    }

    if (filtrados.length === 0) {
      return res.json({
        kpis: { recaudo_total: 0, facturas: 0, clientes: 0, promedio: 0 },
        por_dia: [],
        por_cliente: []
      });
    }

    const recaudoTotal = filtrados.reduce((s, r) => s + (parseFloat(r.recaudo) || 0), 0);
    const facturas = filtrados.length;
    const clientesUnicos = new Set(filtrados.map(r => r.nombre_cliente).filter(Boolean)).size;

    const porDiaMap = {};
    const esMes = agrupar === 'mes';
    for (const r of filtrados) {
      const fecha = r.fecha_recaudo || null;
      const clave = !fecha ? 'Sin fecha' : esMes ? fecha.slice(0, 7) : fecha;
      porDiaMap[clave] = porDiaMap[clave] || { fecha: clave, recaudo: 0, facturas: 0 };
      porDiaMap[clave].recaudo += parseFloat(r.recaudo) || 0;
      porDiaMap[clave].facturas += 1;
    }
    const porDia = Object.values(porDiaMap).sort((a, b) => a.fecha.localeCompare(b.fecha));

    const porClienteMap = {};
    for (const r of filtrados) {
      const c = r.nombre_cliente || 'Sin cliente';
      porClienteMap[c] = porClienteMap[c] || { recaudo: 0, facturas: 0 };
      porClienteMap[c].recaudo += parseFloat(r.recaudo) || 0;
      porClienteMap[c].facturas += 1;
    }
    const porCliente = Object.entries(porClienteMap)
      .map(([nombre, v]) => ({ nombre, recaudo: v.recaudo, facturas: v.facturas, porcentaje: recaudoTotal ? (v.recaudo / recaudoTotal) * 100 : 0 }))
      .sort((a, b) => b.recaudo - a.recaudo);

    res.json({
      kpis: {
        recaudo_total: recaudoTotal,
        facturas,
        clientes: clientesUnicos,
        promedio: facturas ? recaudoTotal / facturas : 0
      },
      por_dia: porDia,
      por_cliente: porCliente
    });
  } catch (err) {
    console.error('Error dashboard:', err);
    res.status(500).json({ error: err.message });
  }
});

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
