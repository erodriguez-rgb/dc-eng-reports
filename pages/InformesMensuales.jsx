import { useState, useEffect } from "react";
import { base44 } from "@/api/base44Client";
import { Link } from "react-router-dom";
import { FileText, ArrowLeft, Calendar, TrendingUp, Download, Loader2 } from "lucide-react";

const RAILWAY_URL = "https://dc-eng-reports-production.up.railway.app";
const RAILWAY_TOKEN = "dceng2026secret";

function formatFecha(fecha) {
  if (!fecha) return "—";
  const [y, m, d] = fecha.split("-");
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  return `${meses[parseInt(m,10)-1]} ${d}, ${y}`;
}

function getMesTexto(mes) {
  const [anio, mesNum] = mes.split("-");
  const meses = ["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  return `${meses[parseInt(mesNum,10)-1]} ${anio}`;
}

export default function InformesMensuales() {
  const [proyectos, setProyectos] = useState([]);
  const [agencias, setAgencias] = useState([]);
  const [user, setUser] = useState(null);
  const [proyectoId, setProyectoId] = useState("");
  const [mes, setMes] = useState(new Date().toISOString().slice(0, 7));
  const [resumen, setResumen] = useState(null);
  const [loading, setLoading] = useState(false);
  const [loadingInit, setLoadingInit] = useState(true);
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState(null);

  useEffect(() => { loadData(); }, []);

  async function loadData() {
    const me = await base44.auth.me();
    setUser(me);
    const [p, a] = await Promise.all([
      base44.entities.Proyectos.list(),
      base44.entities.AgenciasMunicipios.list()
    ]);
    const accesibles = p.filter(pr => me?.role === "admin" || (pr.usuarios_asignados || []).includes(me?.email));
    setProyectos(accesibles);
    setAgencias(a);
    setLoadingInit(false);
  }

  async function generarResumen() {
    if (!proyectoId || !mes) return;
    setLoading(true);
    setResumen(null);
    const [anio, mesNum] = mes.split("-");
    const fechaInicio = `${anio}-${mesNum}-01`;
    const lastDay = new Date(parseInt(anio), parseInt(mesNum), 0).getDate();
    const fechaFin = `${anio}-${mesNum}-${lastDay.toString().padStart(2, "0")}`;

    const todos = await base44.entities.InformesDiarios.filter({ id_proyecto: proyectoId });
    const filtrados = todos.filter(i => i.fecha >= fechaInicio && i.fecha <= fechaFin && i.estatus !== "Borrador");

    const actividadesConsolidadas = filtrados.map(i => `• ${i.fecha}: ${i.actividades || "—"}`).join("\n");
    const observacionesConsolidadas = filtrados.filter(i => i.observaciones).map(i => `• ${i.fecha}: ${i.observaciones}`).join("\n");

    setResumen({
      total: filtrados.length,
      fechaInicio,
      fechaFin,
      actividadesConsolidadas,
      observacionesConsolidadas,
      informes: filtrados,
    });
    setLoading(false);
  }

  // ── EXPORTAR DOCX — llama al servidor Railway ─────────────────────────────
  async function exportarDocx() {
    if (!resumen) return;
    setExporting(true);
    setExportError(null);

    const proyectoActual = proyectos.find(p => p.id === proyectoId);
    const agenciaActual  = agencias.find(a => a.id === proyectoActual?.id_agencia);

    const payload = {
      numero_informe:      1,
      nombre_proyecto:     proyectoActual?.nombre_proyecto || "Proyecto",
      agencia:             agenciaActual?.nombre_agencia   || "",
      contratista:         proyectoActual?.contratista     || "",
      numero_contrato:     proyectoActual?.numero_contrato || "",
      inspector:           user?.full_name || user?.email  || "",
      periodo_texto:       `${formatFecha(resumen.fechaInicio)} al ${formatFecha(resumen.fechaFin)}`,
      resumen_actividades: resumen.actividadesConsolidadas || "",
      observaciones:       resumen.observacionesConsolidadas || "",
      total_informes:      resumen.total,
      actividades:         resumen.informes.map(i => ({
        fecha:        i.fecha,
        descripcion:  i.actividades || "",
        observacion:  i.observaciones || "",
      })),
      comunicaciones:         [],
      submittals_acumulativos:[],
      submittals_periodo:     [],
      certificaciones:        [],
      rfis:                   [],
    };

    try {
      const resp = await fetch(`${RAILWAY_URL}/generar-informe-mensual`, {
        method:  "POST",
        headers: {
          "Content-Type": "application/json",
          "X-API-Token":  RAILWAY_TOKEN,
        },
        body: JSON.stringify(payload),
      });

      if (!resp.ok) {
        const errText = await resp.text();
        throw new Error(`Error del servidor (${resp.status}): ${errText}`);
      }

      const blob = await resp.blob();
      const url  = URL.createObjectURL(blob);
      const a    = document.createElement("a");
      a.href     = url;
      const mesTexto = getMesTexto(mes);
      a.download = `Informe_Mensual_${proyectoActual?.nombre_proyecto || "proyecto"}_${mesTexto}.docx`
        .replace(/[/\\?%*:|"<>]/g, "-");
      a.click();
      URL.revokeObjectURL(url);

    } catch (e) {
      console.error("exportarDocx error:", e);
      setExportError(e.message);
    }

    setExporting(false);
  }

  const proyectoActual = proyectos.find(p => p.id === proyectoId);
  const agenciaActual  = agencias.find(a => a.id === proyectoActual?.id_agencia);

  if (loadingInit) return (
    <div className="flex items-center justify-center h-64">
      <div className="w-8 h-8 border-4 border-blue-200 border-t-blue-600 rounded-full animate-spin" />
    </div>
  );

  return (
    <div className="min-h-screen bg-gray-50">
      <div className="bg-white border-b border-gray-200 px-6 py-4">
        <div className="max-w-5xl mx-auto flex items-center justify-between">
          <div className="flex items-center gap-3">
            <Link to="/Dashboard" className="text-gray-400 hover:text-gray-600"><ArrowLeft className="w-5 h-5" /></Link>
            <Calendar className="w-6 h-6 text-purple-600" />
            <h1 className="text-xl font-bold text-gray-900">Informe Mensual</h1>
          </div>
        </div>
      </div>

      <div className="max-w-5xl mx-auto px-6 py-6 space-y-6">

        {/* Selector */}
        <div className="bg-white rounded-xl border border-gray-200 p-6">
          <h2 className="font-semibold text-gray-700 mb-4">Parámetros del Informe</h2>
          <div className="grid md:grid-cols-3 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Proyecto</label>
              <select
                value={proyectoId}
                onChange={e => { setProyectoId(e.target.value); setResumen(null); setExportError(null); }}
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500"
              >
                <option value="">Seleccionar...</option>
                {proyectos.map(p => <option key={p.id} value={p.id}>{p.nombre_proyecto}</option>)}
              </select>
            </div>
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-1">Mes</label>
              <input
                type="month"
                value={mes}
                onChange={e => { setMes(e.target.value); setResumen(null); setExportError(null); }}
                className="w-full border border-gray-200 rounded-lg px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-purple-500"
              />
            </div>
            <div className="flex items-end">
              <button
                onClick={generarResumen}
                disabled={!proyectoId || loading}
                className="w-full bg-purple-600 text-white py-2 rounded-lg text-sm font-medium hover:bg-purple-700 disabled:opacity-50 flex items-center justify-center gap-2"
              >
                {loading
                  ? <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin" />
                  : <TrendingUp className="w-4 h-4" />
                }
                Generar Resumen
              </button>
            </div>
          </div>
        </div>

        {/* Resumen */}
        {resumen && (
          <>
            <div className="grid md:grid-cols-3 gap-4">
              <div className="bg-white rounded-xl border border-gray-200 p-5 text-center">
                <div className="text-3xl font-bold text-purple-600 mb-1">{resumen.total}</div>
                <div className="text-sm text-gray-500">Informes Diarios</div>
              </div>
              <div className="bg-white rounded-xl border border-gray-200 p-5 text-center">
                <div className="text-lg font-bold text-gray-800 mb-1">{formatFecha(resumen.fechaInicio)}</div>
                <div className="text-sm text-gray-500">Fecha Inicio</div>
              </div>
              <div className="bg-white rounded-xl border border-gray-200 p-5 text-center">
                <div className="text-lg font-bold text-gray-800 mb-1">{formatFecha(resumen.fechaFin)}</div>
                <div className="text-sm text-gray-500">Fecha Fin</div>
              </div>
            </div>

            {agenciaActual?.logo_url && (
              <div className="bg-white rounded-xl border border-gray-200 p-4 flex items-center gap-4">
                <img src={agenciaActual.logo_url} alt={agenciaActual.nombre_agencia} className="h-12 object-contain" />
                <div>
                  <div className="font-semibold text-gray-800">{agenciaActual.nombre_agencia}</div>
                  <div className="text-sm text-gray-500">{proyectoActual?.nombre_proyecto}</div>
                </div>
              </div>
            )}

            <div className="bg-white rounded-xl border border-gray-200 p-6">
              <h3 className="font-semibold text-gray-800 mb-3">Actividades Consolidadas del Mes</h3>
              <pre className="text-sm text-gray-700 whitespace-pre-wrap bg-gray-50 rounded-lg p-4 leading-relaxed">
                {resumen.actividadesConsolidadas || "No hay actividades registradas"}
              </pre>
            </div>

            {resumen.observacionesConsolidadas && (
              <div className="bg-white rounded-xl border border-gray-200 p-6">
                <h3 className="font-semibold text-gray-800 mb-3">Observaciones del Mes</h3>
                <pre className="text-sm text-gray-700 whitespace-pre-wrap bg-gray-50 rounded-lg p-4 leading-relaxed">
                  {resumen.observacionesConsolidadas}
                </pre>
              </div>
            )}

            <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 text-sm text-blue-700">
              <strong>📧 Borrador de Email:</strong><br />
              Asunto: Resumen de progreso — {proyectoActual?.nombre_proyecto} — {getMesTexto(mes)}<br />
              Cuerpo: {resumen.total} informes diarios registrados. Actividades principales compiladas adjuntas.
            </div>

            {/* Error de exportación */}
            {exportError && (
              <div className="bg-red-50 border border-red-200 rounded-xl p-4 text-sm text-red-700">
                <strong>Error al exportar:</strong> {exportError}
              </div>
            )}

            <div className="flex gap-3">
              <button
                onClick={exportarDocx}
                disabled={exporting}
                className="flex items-center gap-2 bg-gray-800 text-white px-5 py-2.5 rounded-lg text-sm font-medium hover:bg-gray-900 disabled:opacity-50"
              >
                {exporting
                  ? <Loader2 className="w-4 h-4 animate-spin" />
                  : <Download className="w-4 h-4" />
                }
                {exporting ? "Generando DOCX…" : "Exportar .docx"}
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  );
}
