import React, { useState, useEffect, useMemo, useRef } from "react";
import {
  Plus,
  Trash2,
  Printer,
  ClipboardList,
  Wallet,
  ChevronLeft,
  Check,
  Clock,
  Pencil,
  Home,
  Search,
} from "lucide-react";
import * as XLSX from "xlsx";
import { storage } from "../lib/storage";
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
} from "recharts";


const LOGO_URI =
  "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAWgAAADCCAYAAABkFVjpAACJeklEQVR42u2dd5gkVb3+P+dU6Dh5Ni+w7JJzVhFRJCMYUERAMFwz3qvo1Z8BBUzXcBUM14SgoKgEUZScRAEByWEDS1jYvDt5plOlc35/nKqenrBpCDu7W9/nqWdCd1dXOPWe97zfBKmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWWWmqppZZaaqmlllpqqaWW2itl/x80g1Q5g5jqTAAAAABJRU5ErkJggg==";

// ---------- helpers ----------
const brl = (n) =>
  (Number.isFinite(n) ? n : 0).toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL",
  });

const uid = () => Math.random().toString(36).slice(2, 10);

// ---------- import helpers (colar OS antiga da planilha) ----------
function normStr(s) {
  return (s || "")
    .toString()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function parseBRLNumber(raw) {
  if (raw === undefined || raw === null) return 0;
  let s = raw.toString().replace(/r\$|R\$/g, "").trim();
  if (!s) return 0;
  const hasComma = s.includes(",");
  const hasDot = s.includes(".");
  if (hasComma && hasDot) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (hasComma) {
    s = s.replace(",", ".");
  }
  s = s.replace(/[^\d.-]/g, "");
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : 0;
}

function parseDateBR(raw) {
  if (!raw) return "";
  const m = raw.toString().trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return "";
  let [, d, mo, y] = m;
  if (y.length === 2) y = "20" + y;
  d = d.padStart(2, "0");
  mo = mo.padStart(2, "0");
  return `${y}-${mo}-${d}`;
}

function rowsToOS(rows) {
  const cellStr = (v) => (v === undefined || v === null ? "" : v.toString());
  rows = rows.map((row) => row.map(cellStr));
  const result = { data: "", cliente: "", carro: "", km: "", cor: "", placa: "", items: [] };

  let headerRowIdx = -1;
  let colMap = {};

  rows.forEach((row, rIdx) => {
    if (headerRowIdx !== -1) return;
    row.forEach((cell, cIdx) => {
      const n = normStr(cell);
      if (!n) return;
      const nextVal = (row.slice(cIdx + 1).find((c) => c && c.trim() !== "") || "").trim();

      if (n.includes("cliente") && !result.cliente && nextVal) result.cliente = nextVal;
      else if ((n.includes("carro") || n.includes("veiculo")) && !result.carro && nextVal)
        result.carro = nextVal;
      else if ((n === "km" || n.startsWith("km")) && !result.km && nextVal) result.km = nextVal;
      else if (n.includes("cor") && !result.cor && nextVal) result.cor = nextVal;
      else if (n.includes("placa") && !result.placa && nextVal) result.placa = nextVal;
      else if (n.includes("data") && !result.data && nextVal) result.data = parseDateBR(nextVal);

      if (n === "item" || n === "itens") {
        const rowNorm = row.map(normStr);
        if (rowNorm.some((c) => c.includes("qtd") || c.includes("quantidade"))) {
          headerRowIdx = rIdx;
          row.forEach((c, ci) => {
            const cn = normStr(c);
            if (cn === "item" || cn === "itens") colMap.item = ci;
            else if (cn.includes("descri")) colMap.descricao = ci;
            else if (cn.includes("qtd") || cn.includes("quantidade")) colMap.qtd = ci;
            else if (cn.includes("unit")) colMap.unitario = ci;
            else if (cn.includes("total")) colMap.total = ci;
          });
        }
      }
    });
  });

  if (headerRowIdx >= 0) {
    for (let r = headerRowIdx + 1; r < rows.length; r++) {
      const row = rows[r];
      const allEmpty = row.every((c) => !c || c.trim() === "");
      if (allEmpty) break;
      const itemName = colMap.item !== undefined ? (row[colMap.item] || "").trim() : "";
      const descricao = colMap.descricao !== undefined ? (row[colMap.descricao] || "").trim() : "";
      if (!itemName && !descricao) continue;
      if (normStr(itemName).includes("total") && !descricao) break;

      const qtdRaw = colMap.qtd !== undefined ? row[colMap.qtd] : "1";
      const qtd = parseBRLNumber(qtdRaw) || 1;
      const unitRaw = colMap.unitario !== undefined ? row[colMap.unitario] : "";
      const totalRaw = colMap.total !== undefined ? row[colMap.total] : "";
      let valorUnit = parseBRLNumber(unitRaw);
      if (!valorUnit && totalRaw) valorUnit = parseBRLNumber(totalRaw) / (qtd || 1);

      result.items.push({
        id: uid(),
        item: itemName || descricao,
        descricao: itemName ? descricao : "",
        categoria: "peca",
        qtd,
        custo: valorUnit || "",
        valorCliente: valorUnit || "",
      });
    }
  }

  return result;
}

function parsePastedOS(text) {
  const rows = text.split(/\r?\n/).map((line) => line.split("\t"));
  return rowsToOS(rows);
}

async function readWorkbookOS(file) {
  const buf = await file.arrayBuffer();
  const wb = XLSX.read(buf, { type: "array" });
  const results = [];
  wb.SheetNames.forEach((name) => {
    const ws = wb.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
    const parsed = rowsToOS(rows);
    const hasContent = parsed.cliente.trim() || parsed.carro.trim() || parsed.items.length > 0;
    if (hasContent) results.push(parsed);
  });
  return { results, totalSheets: wb.SheetNames.length };
}

const CATEGORIAS = [
  { value: "peca", label: "Peça" },
  { value: "servico", label: "Serviço" },
  { value: "mao_obra", label: "Mão de obra / Deslocamento" },
];

const emptyItem = (overrides = {}) => ({
  id: uid(),
  item: "",
  descricao: "",
  categoria: "peca",
  qtd: 1,
  custo: "",
  valorCliente: "",
  ...overrides,
});

const emptyLaborItem = (overrides = {}) => ({
  id: uid(),
  item: "MRM Serviço com Deslocamentos",
  descricao: "",
  categoria: "mao_obra",
  custos: [],
  maoDeObra: "",
  ...overrides,
});

const emptyDraft = () => ({
  id: uid(),
  data: new Date().toISOString().slice(0, 10),
  cliente: "",
  carro: "",
  km: "",
  cor: "",
  placa: "",
  status: "aberto", // pagamento: aberto | recebido
  workflowStatus: "andamento", // andamento | fechada
  items: [emptyItem()],
});

function computeItem(it) {
  if (it.categoria === "mao_obra" && Array.isArray(it.custos)) {
    const custoTotal = it.custos.reduce((s, c) => s + (Number(c.valor) || 0), 0);
    const maoDeObra = Number(it.maoDeObra) || 0;
    return {
      custoTotal,
      clienteTotal: custoTotal + maoDeObra,
      valorClienteUnit: custoTotal + maoDeObra,
    };
  }
  const qtd = Number(it.qtd) || 0;
  const custo = Number(it.custo) || 0;
  const valorClienteUnit =
    it.valorCliente === "" || it.valorCliente === null || it.valorCliente === undefined
      ? custo
      : Number(it.valorCliente) || 0;
  return {
    custoTotal: qtd * custo,
    clienteTotal: qtd * valorClienteUnit,
    valorClienteUnit,
  };
}

function osTotals(os) {
  return os.items.reduce(
    (acc, it) => {
      const c = computeItem(it);
      acc.custo += c.custoTotal;
      acc.cliente += c.clienteTotal;
      return acc;
    },
    { custo: 0, cliente: 0 }
  );
}

const STORAGE_KEY = "os-records-v1";
const CLIENTES_KEY = "clientes-v1";
const OUTROS_KEY = "outros-custos-v1";

function upsertCliente(list, os) {
  const cliente = (os.cliente || "").trim();
  const carro = (os.carro || "").trim();
  if (!cliente && !carro) return list;
  const key = (os.placa || "").trim().toUpperCase() || `${cliente}|${carro}`.toLowerCase();
  const entry = { id: key, cliente, carro, placa: os.placa || "", cor: os.cor || "" };
  const idx = list.findIndex((c) => c.id === key);
  if (idx >= 0) {
    const next = list.slice();
    next[idx] = entry;
    return next;
  }
  return [entry, ...list];
}

export default function App() {
  const [tab, setTab] = useState("inicio"); // inicio | servico | previa | painel
  const [records, setRecords] = useState([]);
  const [clientes, setClientes] = useState([]);
  const [outrosCustos, setOutrosCustos] = useState([]);
  const [draft, setDraft] = useState(emptyDraft());
  const [editingId, setEditingId] = useState(null);
  const [loaded, setLoaded] = useState(false);
  const [saveState, setSaveState] = useState("idle"); // idle | saving | saved | error
  const [saveError, setSaveError] = useState("");
  const autosaveTimer = useRef(null);

  useEffect(() => {
    (async () => {
      try {
        const res = await storage.get(STORAGE_KEY, false);
        if (res && res.value) {
          setRecords(JSON.parse(res.value));
        }
      } catch (e) {
        // no records yet
      }
      try {
        const resC = await storage.get(CLIENTES_KEY, false);
        if (resC && resC.value) {
          setClientes(JSON.parse(resC.value));
        }
      } catch (e) {
        // no clients yet
      }
      try {
        const resO = await storage.get(OUTROS_KEY, false);
        if (resO && resO.value) {
          setOutrosCustos(JSON.parse(resO.value));
        }
      } catch (e) {
        // no other costs yet
      } finally {
        setLoaded(true);
      }
    })();
  }, []);

  async function persist(next) {
    setRecords(next);
    setSaveState("saving");
    try {
      const res = await storage.set(STORAGE_KEY, JSON.stringify(next), false);
      setSaveState(res ? "saved" : "error");
    } catch (e) {
      setSaveState("error");
      setSaveError(e && e.message ? e.message : String(e));
    }
  }

  async function persistClientes(next) {
    setClientes(next);
    try {
      await storage.set(CLIENTES_KEY, JSON.stringify(next), false);
    } catch (e) {
      // best-effort
    }
  }

  async function persistOutros(next) {
    setOutrosCustos(next);
    try {
      await storage.set(OUTROS_KEY, JSON.stringify(next), false);
    } catch (e) {
      // best-effort
    }
  }

  function addOutroCusto(entry) {
    persistOutros([{ id: uid(), ...entry }, ...outrosCustos]);
  }

  function removeOutroCusto(id) {
    persistOutros(outrosCustos.filter((o) => o.id !== id));
  }

  function updateItem(idx, patch) {
    setDraft((d) => {
      const items = d.items.slice();
      items[idx] = { ...items[idx], ...patch };
      return { ...d, items };
    });
  }

  function addItem() {
    setDraft((d) => ({ ...d, items: [...d.items, emptyItem()] }));
  }

  function addLaborItem() {
    setDraft((d) => {
      if (d.items.some((it) => it.categoria === "mao_obra")) return d;
      return { ...d, items: [...d.items, emptyLaborItem()] };
    });
  }

  function updateLaborItem(idx, patch) {
    setDraft((d) => {
      const items = d.items.slice();
      items[idx] = { ...items[idx], ...patch };
      return { ...d, items };
    });
  }

  function addLaborCusto(idx) {
    setDraft((d) => {
      const items = d.items.slice();
      const custos = [...(items[idx].custos || []), { id: uid(), descricao: "", valor: "" }];
      items[idx] = { ...items[idx], custos };
      return { ...d, items };
    });
  }

  function updateLaborCusto(idx, custoIdx, patch) {
    setDraft((d) => {
      const items = d.items.slice();
      const custos = items[idx].custos.slice();
      custos[custoIdx] = { ...custos[custoIdx], ...patch };
      items[idx] = { ...items[idx], custos };
      return { ...d, items };
    });
  }

  function removeLaborCusto(idx, custoIdx) {
    setDraft((d) => {
      const items = d.items.slice();
      const custos = items[idx].custos.slice();
      custos.splice(custoIdx, 1);
      items[idx] = { ...items[idx], custos };
      return { ...d, items };
    });
  }

  function removeItem(idx) {
    setDraft((d) => {
      const items = d.items.slice();
      items.splice(idx, 1);
      return { ...d, items: items.length ? items : [emptyItem()] };
    });
  }

  function saveOS(finalize) {
    const cleanItems = draft.items.filter(
      (it) => it.item.trim() !== "" || it.descricao.trim() !== ""
    );
    if (draft.cliente.trim() === "" && draft.carro.trim() === "") {
      return;
    }
    if (finalize && cleanItems.length === 0) {
      return;
    }
    const nextWorkflow = finalize
      ? "fechada"
      : draft.workflowStatus === "fechada"
      ? "fechada"
      : "andamento";
    const toSave = { ...draft, items: cleanItems, workflowStatus: nextWorkflow };
    let next;
    if (editingId) {
      next = records.map((r) => (r.id === editingId ? toSave : r));
    } else {
      next = [toSave, ...records];
    }
    persist(next);
    persistClientes(upsertCliente(clientes, toSave));
    setEditingId(toSave.id);
    if (finalize) {
      setDraft(toSave);
      setTab("previa");
    }
  }

  useEffect(() => {
    if (!loaded) return;
    if (!draft.cliente.trim() && !draft.carro.trim()) return;
    if (autosaveTimer.current) clearTimeout(autosaveTimer.current);
    autosaveTimer.current = setTimeout(() => {
      saveOS(false);
    }, 900);
    return () => clearTimeout(autosaveTimer.current);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [draft, loaded]);

  function openRecord(rec) {
    setDraft(rec);
    setEditingId(rec.id);
    setTab(rec.workflowStatus === "fechada" ? "previa" : "servico");
  }

  function startNew() {
    setDraft(emptyDraft());
    setEditingId(null);
    setTab("servico");
  }

  function handleImport(parsed) {
    setDraft({
      ...emptyDraft(),
      ...parsed,
      items: parsed.items.length ? parsed.items : [emptyItem()],
      workflowStatus: "fechada",
      status: "recebido",
    });
    setEditingId(null);
    setTab("servico");
  }

  function handleBulkImport(parsedList) {
    if (!parsedList.length) return;
    const novos = parsedList.map((p) => ({
      ...emptyDraft(),
      ...p,
      items: p.items.length ? p.items : [emptyItem()],
      workflowStatus: "fechada",
      status: "recebido",
    }));
    persist([...novos, ...records]);
    let nextClientes = clientes;
    novos.forEach((os) => {
      nextClientes = upsertCliente(nextClientes, os);
    });
    persistClientes(nextClientes);
  }

  function deleteRecord(id) {
    persist(records.filter((r) => r.id !== id));
  }

  function toggleStatus(id) {
    persist(
      records.map((r) =>
        r.id === id ? { ...r, status: r.status === "recebido" ? "aberto" : "recebido" } : r
      )
    );
  }

  function toggleWorkflow(id) {
    persist(
      records.map((r) =>
        r.id === id
          ? { ...r, workflowStatus: r.workflowStatus === "fechada" ? "andamento" : "fechada" }
          : r
      )
    );
  }

  const totals = useMemo(() => osTotals(draft), [draft]);

  return (
    <div style={styles.page} className="app-shell">
      <style>{printStyles}</style>
      <Header saveState={saveState} saveError={saveError} onNew={startNew} />
      <TabBar tab={tab} setTab={setTab} />

      <div style={styles.body}>
        {tab === "inicio" && (
          <InicioView
            records={records}
            onOpen={openRecord}
            onNew={startNew}
            onImport={handleImport}
            onBulkImport={handleBulkImport}
            loaded={loaded}
          />
        )}
        {tab === "servico" && (
          <FormView
            draft={draft}
            setDraft={setDraft}
            updateItem={updateItem}
            addItem={addItem}
            addLaborItem={addLaborItem}
            updateLaborItem={updateLaborItem}
            addLaborCusto={addLaborCusto}
            updateLaborCusto={updateLaborCusto}
            removeLaborCusto={removeLaborCusto}
            removeItem={removeItem}
            totals={totals}
            onSave={saveOS}
            isEditing={!!editingId}
            clientes={clientes}
          />
        )}
        {tab === "previa" && (
          <PreviaView
            os={draft}
            totals={totals}
            onEdit={() => setTab("servico")}
          />
        )}
        {tab === "painel" && (
          <PainelView
            records={records}
            outrosCustos={outrosCustos}
            onAddOutro={addOutroCusto}
            onRemoveOutro={removeOutroCusto}
            onEdit={openRecord}
            onDelete={deleteRecord}
            onToggleStatus={toggleStatus}
            onToggleWorkflow={toggleWorkflow}
            loaded={loaded}
          />
        )}
      </div>
    </div>
  );
}

// ---------- header / tabs ----------
function Header({ saveState, saveError, onNew }) {
  return (
    <div style={styles.header} className="no-print">
      <div style={styles.logoBadge}>
        <img src={LOGO_URI} alt="MRM Personal Car" style={styles.logoBadgeImg} />
      </div>
      <div style={{ flex: 1 }}>
        <div style={styles.headerTitle}>MRM Personal Car</div>
        <div style={styles.headerSub}>
          Gestão de ordens de serviço
          {saveState === "saving" && "  ·  salvando…"}
          {saveState === "saved" && "  ·  salvo"}
          {saveState === "error" && "  ·  erro ao salvar"}
        </div>
        {saveState === "error" && saveError && (
          <div style={{ fontSize: 11, color: "#c0392b", marginTop: 2, wordBreak: "break-word" }}>
            {saveError}
          </div>
        )}
      </div>
      <button onClick={onNew} style={styles.newBtn} aria-label="Abrir novo serviço">
        <Plus size={16} color="#101113" />
      </button>
    </div>
  );
}

function TabBar({ tab, setTab }) {
  const items = [
    { key: "inicio", label: "Início", icon: Home },
    { key: "servico", label: "Serviço", icon: ClipboardList },
    { key: "previa", label: "Prévia", icon: Printer },
    { key: "painel", label: "Painel", icon: Wallet },
  ];
  return (
    <div style={styles.tabBar} className="no-print">
      {items.map(({ key, label, icon: Icon }) => (
        <button
          key={key}
          onClick={() => setTab(key)}
          style={{
            ...styles.tabBtn,
            ...(tab === key ? styles.tabBtnActive : {}),
          }}
        >
          <Icon size={15} style={{ marginRight: 6 }} />
          {label}
        </button>
      ))}
    </div>
  );
}

// ---------- Início ----------
function InicioView({ records, onOpen, onNew, onImport, onBulkImport, loaded }) {
  const [busca, setBusca] = useState("");
  const [showImport, setShowImport] = useState(false);
  const [pasteText, setPasteText] = useState("");
  const [importando, setImportando] = useState(false);
  const [resumoImport, setResumoImport] = useState("");

  const andamento = records.filter((r) => r.workflowStatus !== "fechada");

  const resultadosBusca = useMemo(() => {
    if (busca.trim().length < 2) return [];
    const q = busca.toLowerCase();
    return records
      .filter((r) =>
        `${r.cliente} ${r.carro} ${r.placa}`.toLowerCase().includes(q)
      )
      .slice(0, 8);
  }, [busca, records]);

  function analisar() {
    if (!pasteText.trim()) return;
    const parsed = parsePastedOS(pasteText);
    onImport(parsed);
    setPasteText("");
    setShowImport(false);
  }

  const fileInputRef = useRef(null);

  async function handleFile(e) {
    const file = e.target.files[0];
    if (!file) return;
    setImportando(true);
    setResumoImport("");
    try {
      const { results, totalSheets } = await readWorkbookOS(file);
      onBulkImport(results);
      setResumoImport(
        `${results.length} de ${totalSheets} aba(s) foram reconhecidas e importadas como OS fechadas. Confira e ajuste no Painel.`
      );
    } catch (err) {
      setResumoImport(
        "Não consegui ler esse arquivo. Confira se é um .xlsx exportado do Google Sheets (Arquivo → Fazer download → .xlsx)."
      );
    } finally {
      setImportando(false);
      if (fileInputRef.current) fileInputRef.current.value = "";
    }
  }

  return (
    <div>
      <button onClick={onNew} style={styles.bigNewBtn}>
        <Plus size={18} style={{ marginRight: 8 }} /> Novo serviço
      </button>

      <button
        onClick={() => setShowImport((v) => !v)}
        style={styles.importToggleBtn}
      >
        {showImport ? "Cancelar importação" : "Importar OS antigas"}
      </button>

      {showImport && (
        <div style={styles.importBox}>
          <div style={styles.importHint}>
            <strong>Arquivo completo:</strong> no Google Sheets, use Arquivo → Fazer download
            → Microsoft Excel (.xlsx) — se cada OS é uma aba, eu leio todas de uma vez e já
            importo cada uma como uma OS fechada.
          </div>
          <button
            onClick={() => fileInputRef.current?.click()}
            disabled={importando}
            style={{ ...styles.primaryBtnSmall, marginBottom: 12 }}
          >
            {importando ? "Lendo arquivo…" : "Escolher arquivo .xlsx"}
          </button>
          <input
            ref={fileInputRef}
            type="file"
            accept=".xlsx,.xls"
            onChange={handleFile}
            style={{ display: "none" }}
          />
          {resumoImport && <div style={styles.importSummary}>{resumoImport}</div>}

          <div style={styles.orDivider}>ou cole uma OS por vez</div>
          <div style={styles.importHint}>
            Abra a aba de uma OS antiga, selecione tudo (Ctrl+A na área da OS) e copie
            (Ctrl+C). Cole aqui — depois você confere e ajusta antes de salvar. Como a
            planilha antiga só tinha o valor pro cliente, o custo de cada item vem igual ao
            valor cliente; corrija se lembrar do seu custo real.
          </div>
          <textarea
            value={pasteText}
            onChange={(e) => setPasteText(e.target.value)}
            placeholder="Cole aqui o conteúdo copiado da planilha…"
            style={styles.importTextarea}
          />
          <button onClick={analisar} style={styles.primaryBtnSmall}>
            Analisar e revisar
          </button>
        </div>
      )}

      <div style={styles.searchBox}>
        <Search size={15} color="#8A8F98" />
        <input
          value={busca}
          onChange={(e) => setBusca(e.target.value)}
          placeholder="Buscar por cliente, carro ou placa"
          style={styles.searchInput}
        />
      </div>

      {busca.trim().length >= 2 ? (
        <div>
          <SectionLabel style={{ marginTop: 16 }}>Resultados</SectionLabel>
          {resultadosBusca.length === 0 && (
            <div style={styles.emptyMsg}>Nada encontrado para "{busca}".</div>
          )}
          {resultadosBusca.map((r) => (
            <ServicoCard key={r.id} r={r} onOpen={onOpen} />
          ))}
        </div>
      ) : (
        <div>
          <SectionLabel style={{ marginTop: 20 }}>Em andamento</SectionLabel>
          {!loaded && <div style={styles.emptyMsg}>Carregando…</div>}
          {loaded && andamento.length === 0 && (
            <div style={styles.emptyMsg}>
              Nenhum serviço em andamento agora. Toque em "Novo serviço" para começar.
            </div>
          )}
          {andamento.map((r) => (
            <ServicoCard key={r.id} r={r} onOpen={onOpen} />
          ))}
        </div>
      )}
    </div>
  );
}

function ServicoCard({ r, onOpen }) {
  const t = osTotals(r);
  return (
    <button onClick={() => onOpen(r)} style={styles.servicoCard}>
      <div style={{ flex: 1, minWidth: 0, textAlign: "left" }}>
        <div style={styles.recordTitle}>
          {r.cliente || "Sem cliente"} · {r.carro || "—"}
        </div>
        <div style={styles.recordSub}>
          {formatDateBR(r.data)} · custo até agora {brl(t.custo)}
          {r.placa ? ` · ${r.placa}` : ""}
        </div>
      </div>
      <span
        style={{
          ...styles.statusBtn,
          background: r.workflowStatus === "fechada" ? "#101113" : "#F0EEE7",
          color: r.workflowStatus === "fechada" ? "#F5C400" : "#5B6169",
        }}
      >
        {r.workflowStatus === "fechada" ? "Fechada" : "Em andamento"}
      </span>
    </button>
  );
}


function LaborItemCard({ it, idx, updateLaborItem, addLaborCusto, updateLaborCusto, removeLaborCusto, removeItem }) {
  const c = computeItem(it);
  const maoDeObra = Number(it.maoDeObra) || 0;
  return (
    <div style={{ ...styles.itemCard, borderColor: "#F5C400", borderWidth: 1.5 }}>
      <div style={styles.itemRowTop}>
        <input
          value={it.item}
          onChange={(e) => updateLaborItem(idx, { item: e.target.value })}
          style={{ ...styles.input, flex: 1.1, fontWeight: 600 }}
        />
        <button onClick={() => removeItem(idx)} style={styles.iconBtn} aria-label="Remover item">
          <Trash2 size={15} color="#B5651D" />
        </button>
      </div>
      <div style={styles.itemsHintRow}>
        Vá lançando cada custo de deslocamento conforme acontece (buscar o carro, levar pra
        outra oficina, devolver pro cliente etc). Todos somam neste mesmo item.
      </div>

      {(it.custos || []).map((custo, cIdx) => (
        <div key={custo.id} style={styles.custoRow}>
          <input
            value={custo.descricao}
            onChange={(e) => updateLaborCusto(idx, cIdx, { descricao: e.target.value })}
            placeholder="Ex.: Buscar carro"
            style={{ ...styles.inputSmall, flex: 1.4 }}
          />
          <input
            type="number"
            min="0"
            step="0.01"
            value={custo.valor}
            onChange={(e) => updateLaborCusto(idx, cIdx, { valor: e.target.value })}
            placeholder="0,00"
            style={{ ...styles.inputSmall, flex: 0.8 }}
          />
          <button
            onClick={() => removeLaborCusto(idx, cIdx)}
            style={styles.iconBtn}
            aria-label="Remover custo"
          >
            <Trash2 size={13} color="#B5651D" />
          </button>
        </div>
      ))}

      <button onClick={() => addLaborCusto(idx)} style={styles.addCustoBtn}>
        <Plus size={13} style={{ marginRight: 4 }} /> Adicionar custo de deslocamento
      </button>

      <div style={{ marginTop: 10 }}>
        <Field label="Sua mão de obra (o que você cobra pelo serviço)" small>
          <input
            type="number"
            min="0"
            step="0.01"
            value={it.maoDeObra}
            onChange={(e) => updateLaborItem(idx, { maoDeObra: e.target.value })}
            placeholder="0,00"
            style={styles.inputSmall}
          />
        </Field>
      </div>

      <div style={styles.itemFooter}>
        <span>Deslocamentos: {brl(c.custoTotal)}</span>
        <span>Mão de obra: {brl(maoDeObra)}</span>
        <span style={{ fontWeight: 700, color: "#101113" }}>Total: {brl(c.clienteTotal)}</span>
      </div>
    </div>
  );
}

function FormView({
  draft,
  setDraft,
  updateItem,
  addItem,
  addLaborItem,
  updateLaborItem,
  addLaborCusto,
  updateLaborCusto,
  removeLaborCusto,
  removeItem,
  totals,
  onSave,
  isEditing,
  clientes,
}) {
  const set = (patch) => setDraft((d) => ({ ...d, ...patch }));
  const lucro = totals.cliente - totals.custo;
  const [buscaCliente, setBuscaCliente] = useState("");

  const sugestoes = useMemo(() => {
    if (buscaCliente.trim().length < 2) return [];
    const q = buscaCliente.toLowerCase();
    return clientes
      .filter((c) => `${c.cliente} ${c.carro} ${c.placa}`.toLowerCase().includes(q))
      .slice(0, 4);
  }, [buscaCliente, clientes]);

  function aplicarCliente(c) {
    set({ cliente: c.cliente, carro: c.carro, placa: c.placa, cor: c.cor });
    setBuscaCliente("");
  }

  const temMaoDeObra = draft.items.some((it) => it.categoria === "mao_obra");

  return (
    <div style={styles.card}>
      <SectionLabel>{isEditing ? "Editar ordem de serviço" : "Dados do veículo"}</SectionLabel>

      {!isEditing && clientes.length > 0 && (
        <div style={{ marginBottom: 14 }}>
          <div style={styles.searchBox}>
            <Search size={14} color="#8A8F98" />
            <input
              value={buscaCliente}
              onChange={(e) => setBuscaCliente(e.target.value)}
              placeholder="Cliente já atendido? Busque pra puxar os dados"
              style={styles.searchInput}
            />
          </div>
          {sugestoes.map((c) => (
            <button key={c.id} onClick={() => aplicarCliente(c)} style={styles.sugestaoBtn}>
              <span style={{ fontWeight: 600 }}>{c.cliente}</span>
              <span style={{ color: "#8A8F98" }}>
                {" "}
                · {c.carro} {c.placa ? `· ${c.placa}` : ""}
              </span>
            </button>
          ))}
        </div>
      )}

      <div style={styles.grid2}>
        <Field label="Data">
          <input
            type="date"
            value={draft.data}
            onChange={(e) => set({ data: e.target.value })}
            style={styles.input}
          />
        </Field>
        <Field label="Cliente">
          <input
            value={draft.cliente}
            onChange={(e) => set({ cliente: e.target.value })}
            placeholder="Nome do cliente"
            style={styles.input}
          />
        </Field>
      </div>

      <div style={styles.grid2}>
        <Field label="Carro">
          <input
            value={draft.carro}
            onChange={(e) => set({ carro: e.target.value })}
            placeholder="Ex.: BMW X1 2016"
            style={styles.input}
          />
        </Field>
        <Field label="Placa">
          <input
            value={draft.placa}
            onChange={(e) => set({ placa: e.target.value.toUpperCase() })}
            placeholder="ABC1D23"
            style={styles.input}
          />
        </Field>
      </div>

      <div style={styles.grid2}>
        <Field label="KM">
          <input
            value={draft.km}
            onChange={(e) => set({ km: e.target.value })}
            placeholder="93745"
            style={styles.input}
          />
        </Field>
        <Field label="Cor">
          <input
            value={draft.cor}
            onChange={(e) => set({ cor: e.target.value })}
            placeholder="Preta"
            style={styles.input}
          />
        </Field>
      </div>

      <SectionLabel style={{ marginTop: 22 }}>Peças e serviços</SectionLabel>
      <div style={styles.itemsHintRow}>
        Deixe "Valor pro cliente" em branco quando não quiser cobrar margem — o cliente paga
        exatamente o seu custo.
      </div>

      {draft.items.map((it, idx) => {
        if (it.categoria === "mao_obra") {
          return (
            <LaborItemCard
              key={it.id}
              it={it}
              idx={idx}
              updateLaborItem={updateLaborItem}
              addLaborCusto={addLaborCusto}
              updateLaborCusto={updateLaborCusto}
              removeLaborCusto={removeLaborCusto}
              removeItem={removeItem}
            />
          );
        }
        const c = computeItem(it);
        const hasMarkup = it.valorCliente !== "" && it.valorCliente !== null;
        const margem = c.custoTotal > 0 ? ((c.clienteTotal - c.custoTotal) / c.custoTotal) * 100 : 0;
        return (
          <div key={it.id} style={styles.itemCard}>
            <div style={styles.itemRowTop}>
              <input
                value={it.item}
                onChange={(e) => updateItem(idx, { item: e.target.value })}
                placeholder="Item (ex.: Óleo, Oficina)"
                style={{ ...styles.input, flex: 1.1, fontWeight: 600 }}
              />
              <button onClick={() => removeItem(idx)} style={styles.iconBtn} aria-label="Remover item">
                <Trash2 size={15} color="#B5651D" />
              </button>
            </div>
            <select
              value={it.categoria}
              onChange={(e) => updateItem(idx, { categoria: e.target.value })}
              style={{ ...styles.input, marginTop: 6, fontSize: 12 }}
            >
              {CATEGORIAS.map((c) => (
                <option key={c.value} value={c.value}>
                  {c.label}
                </option>
              ))}
            </select>
            <input
              value={it.descricao}
              onChange={(e) => updateItem(idx, { descricao: e.target.value })}
              placeholder="Descrição / marca (ex.: Motul 5w40)"
              style={{ ...styles.input, marginTop: 6 }}
            />
            <div style={styles.itemRowBottom}>
              <Field label="Qtd" small>
                <input
                  type="number"
                  min="0"
                  value={it.qtd}
                  onChange={(e) => updateItem(idx, { qtd: e.target.value })}
                  style={styles.inputSmall}
                />
              </Field>
              <Field label="Seu custo (un.)" small>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={it.custo}
                  onChange={(e) => updateItem(idx, { custo: e.target.value })}
                  placeholder="0,00"
                  style={styles.inputSmall}
                />
              </Field>
              <Field label="Valor cliente (un.)" small>
                <input
                  type="number"
                  min="0"
                  step="0.01"
                  value={it.valorCliente}
                  onChange={(e) => updateItem(idx, { valorCliente: e.target.value })}
                  placeholder="= custo"
                  style={styles.inputSmall}
                />
              </Field>
            </div>
            <div style={styles.itemFooter}>
              <span>Custo total: {brl(c.custoTotal)}</span>
              <span>Cliente paga: {brl(c.clienteTotal)}</span>
              <span style={{ color: hasMarkup && margem > 0 ? "#1F7A4D" : "#8A8F98" }}>
                {hasMarkup ? `Margem ${margem.toFixed(0)}%` : "Sem margem"}
              </span>
            </div>
          </div>
        );
      })}

      <div style={{ display: "flex", gap: 8 }}>
        <button onClick={addItem} style={{ ...styles.addItemBtn, flex: 1 }}>
          <Plus size={15} style={{ marginRight: 6 }} /> Adicionar item
        </button>
        <button
          onClick={addLaborItem}
          disabled={temMaoDeObra}
          style={{
            ...styles.addItemBtn,
            flex: 1,
            borderColor: temMaoDeObra ? "#E4E1D8" : "#F5C400",
            color: temMaoDeObra ? "#B9BCC2" : "#101113",
            cursor: temMaoDeObra ? "default" : "pointer",
          }}
        >
          <Plus size={15} style={{ marginRight: 6 }} />
          {temMaoDeObra ? "Mão de obra já adicionada" : "Mão de obra"}
        </button>
      </div>
      <div style={styles.itemsHintRow}>
        O botão "Mão de obra" cria um único item — vá lançando cada deslocamento nele conforme
        acontece, e defina sua mão de obra quando for fechar a OS.
      </div>

      <div style={styles.summaryBar}>
        <SummaryStat label="Seu custo" value={brl(totals.custo)} />
        <SummaryStat label="Cliente paga" value={brl(totals.cliente)} />
        <SummaryStat label="Seu lucro" value={brl(lucro)} accent />
      </div>

      <div style={styles.formActions}>
        <button onClick={() => onSave(false)} style={styles.secondaryBtn}>
          Salvar custos (continuar depois)
        </button>
        <button onClick={() => onSave(true)} style={styles.primaryBtn}>
          Fechar OS e gerar pro cliente
        </button>
      </div>
    </div>
  );
}

function Field({ label, small, children }) {
  return (
    <div style={{ marginBottom: small ? 0 : 12, flex: small ? 1 : undefined }}>
      <div style={small ? styles.labelSmall : styles.label}>{label}</div>
      {children}
    </div>
  );
}

function SectionLabel({ children, style }) {
  return <div style={{ ...styles.sectionLabel, ...style }}>{children}</div>;
}

function SummaryStat({ label, value, accent }) {
  return (
    <div style={styles.summaryStat}>
      <div style={styles.summaryLabel}>{label}</div>
      <div style={{ ...styles.summaryValue, color: accent ? "#1F7A4D" : "#101113" }}>{value}</div>
    </div>
  );
}

// ---------- Prévia (client-facing OS) ----------
function buildOSText(os, totals) {
  const linhas = [];
  linhas.push("MRM PERSONAL CAR");
  linhas.push("Ordem de Serviço");
  linhas.push("");
  linhas.push(`Data: ${formatDateBR(os.data)}`);
  linhas.push(`Cliente: ${os.cliente || "-"}`);
  linhas.push(`Carro: ${os.carro || "-"}`);
  if (os.km) linhas.push(`KM: ${os.km}`);
  if (os.cor) linhas.push(`Cor: ${os.cor}`);
  if (os.placa) linhas.push(`Placa: ${os.placa}`);
  linhas.push("");
  os.items
    .filter((it) => it.item || it.descricao)
    .forEach((it) => {
      const c = computeItem(it);
      const nome = [it.item, it.descricao].filter(Boolean).join(" - ");
      linhas.push(`${nome} (${it.qtd}x) — ${brl(c.clienteTotal)}`);
    });
  linhas.push("");
  linhas.push(`TOTAL: ${brl(totals.cliente)}`);
  return linhas.join("\n");
}

// ---------- nome automático do arquivo PDF ----------
function slugFileName(os) {
  const parts = [
    "OS",
    os.cliente,
    os.carro,
    formatDateBR(os.data).replace(/\//g, "-"),
  ]
    .filter(Boolean)
    .map((p) => p.toString().trim().replace(/\s+/g, "-"));
  return parts.join("_");
}

function gerarPDF(os) {
  const tituloOriginal = document.title;
  document.title = slugFileName(os);
  const restaurar = () => {
    document.title = tituloOriginal;
    window.removeEventListener("afterprint", restaurar);
  };
  window.addEventListener("afterprint", restaurar);
  window.print();
}

function PreviaView({ os, totals, onEdit }) {
  const isEmpty = os.items.every((it) => !it.item && !it.descricao);
  const [mostrarTexto, setMostrarTexto] = useState(false);
  const textAreaRef = useRef(null);
  const texto = buildOSText(os, totals);

  function abrirTexto() {
    setMostrarTexto(true);
    setTimeout(() => {
      if (textAreaRef.current) {
        textAreaRef.current.focus();
        textAreaRef.current.select();
      }
    }, 50);
  }

  return (
    <div>
      <div className="no-print" style={styles.previaActions}>
        <button onClick={onEdit} style={styles.ghostBtn}>
          <ChevronLeft size={15} style={{ marginRight: 4 }} /> Editar
        </button>
        <button onClick={abrirTexto} style={styles.primaryBtnSmall}>
          <Printer size={15} style={{ marginRight: 6 }} /> Texto pra enviar
        </button>
        <button onClick={() => gerarPDF(os)} style={styles.primaryBtnSmall}>
          <Printer size={15} style={{ marginRight: 6 }} /> Gerar PDF
        </button>
      </div>

      {mostrarTexto && (
        <div className="no-print" style={styles.copyBox}>
          <div style={styles.importHint}>
            Toque no texto abaixo, segure pra selecionar tudo (ou use "Selecionar tudo") e
            copie — depois é só colar no WhatsApp ou Telegram.
          </div>
          <textarea
            ref={textAreaRef}
            readOnly
            value={texto}
            onFocus={(e) => e.target.select()}
            style={styles.copyTextarea}
          />
          <button onClick={() => setMostrarTexto(false)} style={styles.ghostBtn}>
            Fechar
          </button>
        </div>
      )}

      <div style={styles.osPaper} id="os-paper">
        {os.workflowStatus === "andamento" && (
          <div style={styles.draftRibbon} className="no-print">
            OS em andamento — os valores ainda podem mudar até você fechar
          </div>
        )}
        <div style={styles.osTopBar} />
        <div style={styles.osHeaderRow}>
          <div>
            <div style={styles.osCompany}>MRM Personal Car ME</div>
            <div style={styles.osCompanyDetail}>www.mrmpersonalcar.com.br</div>
            <div style={styles.osCompanyDetail}>contato@mrmpersonalcar.com.br</div>
            <div style={styles.osCompanyDetail}>Tel: +55 11 97187-8884</div>
          </div>
          <div style={styles.osLogo}>
            <img src={LOGO_URI} alt="MRM Personal Car" style={styles.osLogoImg} />
          </div>
        </div>

        <div style={styles.osTitle}>Ordem de Serviço</div>

        <div style={styles.osMetaGrid}>
          <MetaField label="Data" value={formatDateBR(os.data)} />
          <MetaField label="Cliente" value={os.cliente || "—"} />
          <MetaField label="Carro" value={os.carro || "—"} />
          <MetaField label="KM" value={os.km || "—"} />
          <MetaField label="Cor" value={os.cor || "—"} />
          <MetaField label="Placa" value={os.placa || "—"} />
        </div>

        <div style={styles.osTableHeader}>
          <span style={{ flex: 1.2 }}>Item</span>
          <span style={{ flex: 1.6 }}>Descrição</span>
          <span style={{ width: 40, textAlign: "right" }}>Qtd</span>
          <span style={{ width: 90, textAlign: "right" }}>Unitário</span>
          <span style={{ width: 90, textAlign: "right" }}>Total</span>
        </div>
        {isEmpty ? (
          <div style={styles.osEmptyRow}>Adicione itens na aba "Nova OS" para ver a prévia aqui.</div>
        ) : (
          os.items
            .filter((it) => it.item || it.descricao)
            .map((it) => {
              const c = computeItem(it);
              const qtdExibida = it.categoria === "mao_obra" ? 1 : it.qtd;
              return (
                <div key={it.id} style={styles.osTableRow}>
                  <span style={{ flex: 1.2 }}>{it.item}</span>
                  <span style={{ flex: 1.6, color: "#5B6169" }}>{it.descricao}</span>
                  <span style={{ width: 40, textAlign: "right" }}>{qtdExibida}</span>
                  <span style={{ width: 90, textAlign: "right" }}>{brl(c.valorClienteUnit)}</span>
                  <span style={{ width: 90, textAlign: "right", fontWeight: 600 }}>
                    {brl(c.clienteTotal)}
                  </span>
                </div>
              );
            })
        )}

        <div style={styles.osTotalRow}>
          <span>Total</span>
          <span>{brl(totals.cliente)}</span>
        </div>

        {os.status === "recebido" && <div style={styles.stamp}>PAGO</div>}
      </div>
    </div>
  );
}

function MetaField({ label, value }) {
  return (
    <div style={styles.metaField}>
      <div style={styles.metaLabel}>{label}</div>
      <div style={styles.metaValue}>{value}</div>
    </div>
  );
}

function formatDateBR(iso) {
  if (!iso) return "—";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

// ---------- Painel financeiro ----------
function PainelView({
  records,
  outrosCustos,
  onAddOutro,
  onRemoveOutro,
  onEdit,
  onDelete,
  onToggleStatus,
  onToggleWorkflow,
  loaded,
}) {
  const [filtro, setFiltro] = useState("");
  const [periodoInicio, setPeriodoInicio] = useState("");
  const [periodoFim, setPeriodoFim] = useState("");
  const [novoOutro, setNovoOutro] = useState({
    data: new Date().toISOString().slice(0, 10),
    descricao: "",
    valor: "",
  });

  const dentroPeriodo = (dataStr) => {
    if (!dataStr) return !periodoInicio && !periodoFim;
    if (periodoInicio && dataStr < periodoInicio) return false;
    if (periodoFim && dataStr > periodoFim) return false;
    return true;
  };

  const recordsPeriodo = useMemo(
    () => records.filter((r) => dentroPeriodo(r.data)),
    [records, periodoInicio, periodoFim]
  );

  const outrosPeriodo = useMemo(
    () => outrosCustos.filter((o) => dentroPeriodo(o.data)),
    [outrosCustos, periodoInicio, periodoFim]
  );

  const dados = useMemo(() => {
    let custoServicos = 0,
      faturado = 0,
      recebido = 0,
      aReceber = 0;
    const porCategoria = { peca: 0, servico: 0, mao_obra: 0 };
    const chart = recordsPeriodo
      .slice()
      .reverse()
      .map((r) => {
        const t = osTotals(r);
        custoServicos += t.custo;
        faturado += t.cliente;
        if (r.status === "recebido") recebido += t.cliente;
        else aReceber += t.cliente;
        r.items.forEach((it) => {
          const c = computeItem(it);
          const cat = it.categoria || "peca";
          porCategoria[cat] = (porCategoria[cat] || 0) + c.clienteTotal;
        });
        return {
          nome: r.carro ? r.carro.slice(0, 14) : r.cliente.slice(0, 14),
          Custo: Math.round(t.custo),
          Faturado: Math.round(t.cliente),
        };
      });
    const outrosTotal = outrosPeriodo.reduce((s, o) => s + (Number(o.valor) || 0), 0);
    return {
      custoServicos,
      faturado,
      outrosTotal,
      lucroLiquido: faturado - custoServicos - outrosTotal,
      recebido,
      aReceber,
      porCategoria,
      chart,
    };
  }, [recordsPeriodo, outrosPeriodo]);

  const filtrados = useMemo(() => {
    if (filtro.trim().length < 2) return recordsPeriodo;
    const q = filtro.toLowerCase();
    return recordsPeriodo.filter((r) =>
      `${r.cliente} ${r.carro} ${r.placa}`.toLowerCase().includes(q)
    );
  }, [filtro, recordsPeriodo]);

  function limparPeriodo() {
    setPeriodoInicio("");
    setPeriodoFim("");
  }

  function salvarOutro() {
    if (!novoOutro.descricao.trim() || !novoOutro.valor) return;
    onAddOutro({ ...novoOutro });
    setNovoOutro({ data: novoOutro.data, descricao: "", valor: "" });
  }

  return (
    <div>
      <SectionLabel>Período</SectionLabel>
      <div style={styles.periodoRow}>
        <input
          type="date"
          value={periodoInicio}
          onChange={(e) => setPeriodoInicio(e.target.value)}
          style={{ ...styles.input, flex: 1 }}
        />
        <span style={{ color: "#8A8F98", fontSize: 12 }}>até</span>
        <input
          type="date"
          value={periodoFim}
          onChange={(e) => setPeriodoFim(e.target.value)}
          style={{ ...styles.input, flex: 1 }}
        />
        {(periodoInicio || periodoFim) && (
          <button onClick={limparPeriodo} style={styles.iconBtn} aria-label="Limpar período">
            <Trash2 size={14} color="#5B6169" />
          </button>
        )}
      </div>

      <div style={{ ...styles.statsGrid, marginTop: 14 }}>
        <StatCard label="Faturado" value={brl(dados.faturado)} />
        <StatCard label="Custo serviços" value={brl(dados.custoServicos)} />
        <StatCard label="Outros custos" value={brl(dados.outrosTotal)} tone="#B5651D" />
        <StatCard label="Lucro líquido" value={brl(dados.lucroLiquido)} accent />
      </div>

      <div style={{ ...styles.statsGrid, marginTop: 8 }}>
        <StatCard label="Recebido" value={brl(dados.recebido)} tone="#1F7A4D" />
        <StatCard label="A receber" value={brl(dados.aReceber)} tone="#B5651D" />
      </div>

      <div style={styles.statsGrid3}>
        <StatCard label="Peças (faturado)" value={brl(dados.porCategoria.peca)} />
        <StatCard label="Serviços (faturado)" value={brl(dados.porCategoria.servico)} />
        <StatCard label="Mão de obra (faturado)" value={brl(dados.porCategoria.mao_obra)} tone="#101113" />
      </div>

      {dados.chart.length > 0 && (
        <div style={styles.chartCard}>
          <SectionLabel>Custo vs. faturado por OS</SectionLabel>
          <div style={{ width: "100%", height: 220 }}>
            <ResponsiveContainer>
              <BarChart data={dados.chart} margin={{ top: 10, right: 8, left: -16, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#E4E1D8" />
                <XAxis dataKey="nome" tick={{ fontSize: 10, fill: "#5B6169" }} />
                <YAxis tick={{ fontSize: 10, fill: "#5B6169" }} />
                <Tooltip
                  formatter={(v) => brl(v)}
                  contentStyle={{ fontSize: 12, borderRadius: 8, border: "1px solid #E4E1D8" }}
                />
                <Legend wrapperStyle={{ fontSize: 11 }} />
                <Bar dataKey="Custo" fill="#B5651D" radius={[3, 3, 0, 0]} />
                <Bar dataKey="Faturado" fill="#F5C400" radius={[3, 3, 0, 0]} />
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <SectionLabel style={{ marginTop: 22 }}>Outros custos (fora dos serviços)</SectionLabel>
      <div style={styles.itemsHintRow}>
        Combustível do dia a dia, ferramentas, aluguel, ou qualquer gasto do negócio que não
        seja de uma OS específica.
      </div>
      <div style={styles.outroForm}>
        <input
          type="date"
          value={novoOutro.data}
          onChange={(e) => setNovoOutro({ ...novoOutro, data: e.target.value })}
          style={{ ...styles.inputSmall, marginBottom: 6 }}
        />
        <input
          value={novoOutro.descricao}
          onChange={(e) => setNovoOutro({ ...novoOutro, descricao: e.target.value })}
          placeholder="Descrição (ex.: Combustível)"
          style={{ ...styles.inputSmall, marginBottom: 6 }}
        />
        <div style={{ display: "flex", gap: 6 }}>
          <input
            type="number"
            min="0"
            step="0.01"
            value={novoOutro.valor}
            onChange={(e) => setNovoOutro({ ...novoOutro, valor: e.target.value })}
            placeholder="0,00"
            style={{ ...styles.inputSmall, flex: 1 }}
          />
          <button onClick={salvarOutro} style={styles.primaryBtnSmall}>
            <Plus size={14} style={{ marginRight: 4 }} /> Adicionar
          </button>
        </div>
      </div>
      {outrosPeriodo.length === 0 && (
        <div style={styles.emptyMsg}>Nenhum outro custo lançado nesse período.</div>
      )}
      {outrosPeriodo.map((o) => (
        <div key={o.id} style={styles.recordRow}>
          <div style={styles.recordTopRow}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <div style={styles.recordTitle}>{o.descricao}</div>
              <div style={styles.recordSub}>
                {formatDateBR(o.data)} · {brl(Number(o.valor) || 0)}
              </div>
            </div>
            <button onClick={() => onRemoveOutro(o.id)} style={styles.iconBtn} aria-label="Excluir">
              <Trash2 size={14} color="#B5651D" />
            </button>
          </div>
        </div>
      ))}

      <SectionLabel style={{ marginTop: 22 }}>Histórico de OS</SectionLabel>
      <div style={{ ...styles.searchBox, marginBottom: 10 }}>
        <Search size={14} color="#8A8F98" />
        <input
          value={filtro}
          onChange={(e) => setFiltro(e.target.value)}
          placeholder="Filtrar por cliente, carro ou placa"
          style={styles.searchInput}
        />
      </div>
      {!loaded && <div style={styles.emptyMsg}>Carregando…</div>}
      {loaded && recordsPeriodo.length === 0 && (
        <div style={styles.emptyMsg}>Nenhuma OS nesse período.</div>
      )}
      {loaded && recordsPeriodo.length > 0 && filtrados.length === 0 && (
        <div style={styles.emptyMsg}>Nada encontrado para "{filtro}".</div>
      )}
      {filtrados.map((r) => {
        const t = osTotals(r);
        const lucro = t.cliente - t.custo;
        return (
          <div key={r.id} style={styles.recordRow}>
            <div style={styles.recordTopRow}>
              <div style={{ flex: 1, minWidth: 0 }}>
                <div style={styles.recordTitle}>
                  {r.cliente || "Sem cliente"} · {r.carro || "—"}
                </div>
                <div style={styles.recordSub}>
                  {formatDateBR(r.data)} · custo {brl(t.custo)} · lucro {brl(lucro)}
                </div>
              </div>
              <button onClick={() => onEdit(r)} style={styles.iconBtn} aria-label="Editar">
                <Pencil size={14} color="#5B6169" />
              </button>
              <button onClick={() => onDelete(r.id)} style={styles.iconBtn} aria-label="Excluir">
                <Trash2 size={14} color="#B5651D" />
              </button>
            </div>
            <div style={styles.recordBadgeRow}>
              <button
                onClick={() => onToggleWorkflow(r.id)}
                style={{
                  ...styles.statusBtn,
                  background: r.workflowStatus === "fechada" ? "#101113" : "#F0EEE7",
                  color: r.workflowStatus === "fechada" ? "#F5C400" : "#5B6169",
                }}
              >
                {r.workflowStatus === "fechada" ? "Fechada" : "Em andamento"}
              </button>
              <button
                onClick={() => onToggleStatus(r.id)}
                style={{
                  ...styles.statusBtn,
                  background: r.status === "recebido" ? "#EAF4EE" : "#FBF0E4",
                  color: r.status === "recebido" ? "#1F7A4D" : "#B5651D",
                }}
              >
                {r.status === "recebido" ? <Check size={12} /> : <Clock size={12} />}
                {r.status === "recebido" ? "Recebido" : "A receber"}
              </button>
            </div>
          </div>
        );
      })}
    </div>
  );
}

function StatCard({ label, value, tone, accent }) {
  return (
    <div style={styles.statCard}>
      <div style={styles.statLabel}>{label}</div>
      <div style={{ ...styles.statValue, color: tone || (accent ? "#1F7A4D" : "#101113") }}>
        {value}
      </div>
    </div>
  );
}

// ---------- styles ----------
const styles = {
  page: {
    minHeight: "100vh",
    background: "#F6F5F1",
    fontFamily:
      "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif",
    color: "#101113",
    paddingBottom: 40,
  },
  header: {
    display: "flex",
    alignItems: "center",
    gap: 10,
    padding: "18px 16px 12px",
  },
  logoBadge: {
    width: 44,
    height: 24,
    borderRadius: 4,
    background: "transparent",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
    overflow: "hidden",
  },
  logoBadgeImg: {
    width: "100%",
    height: "100%",
    objectFit: "contain",
  },
  headerTitle: { fontWeight: 700, fontSize: 15, letterSpacing: 0.2 },
  headerSub: { fontSize: 11.5, color: "#8A8F98", marginTop: 1 },
  newBtn: {
    width: 32,
    height: 32,
    borderRadius: "50%",
    border: "1px solid #DEDBD1",
    background: "#FFFFFF",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
    flexShrink: 0,
  },
  tabBar: {
    display: "flex",
    gap: 6,
    padding: "0 16px 14px",
    borderBottom: "1px solid #E4E1D8",
  },
  tabBtn: {
    display: "flex",
    alignItems: "center",
    fontSize: 12.5,
    fontWeight: 600,
    padding: "8px 12px",
    borderRadius: 8,
    border: "1px solid transparent",
    background: "transparent",
    color: "#5B6169",
    cursor: "pointer",
  },
  tabBtnActive: {
    background: "#101113",
    color: "#F5C400",
  },
  body: { padding: "16px" },
  bigNewBtn: {
    width: "100%",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "14px",
    borderRadius: 12,
    border: "none",
    background: "#F5C400",
    color: "#101113",
    fontSize: 14.5,
    fontWeight: 700,
    cursor: "pointer",
    marginBottom: 14,
  },
  importToggleBtn: {
    width: "100%",
    padding: "10px",
    borderRadius: 10,
    border: "1px dashed #C7C2B3",
    background: "transparent",
    color: "#5B6169",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
    marginBottom: 14,
  },
  importBox: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: 12,
    marginBottom: 16,
  },
  importHint: {
    fontSize: 11.5,
    color: "#5B6169",
    lineHeight: 1.5,
    marginBottom: 10,
  },
  importTextarea: {
    width: "100%",
    minHeight: 120,
    boxSizing: "border-box",
    borderRadius: 8,
    border: "1px solid #DEDBD1",
    background: "#FCFBF8",
    padding: 10,
    fontSize: 12,
    fontFamily: "monospace",
    marginBottom: 10,
    resize: "vertical",
  },
  importSummary: {
    fontSize: 12,
    color: "#1F7A4D",
    background: "#EAF4EE",
    borderRadius: 8,
    padding: "8px 10px",
    marginTop: 8,
  },
  orDivider: {
    textAlign: "center",
    fontSize: 11,
    color: "#8A8F98",
    margin: "14px 0 10px",
    fontWeight: 600,
  },
  searchBox: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: "10px 12px",
  },
  searchInput: {
    flex: 1,
    border: "none",
    outline: "none",
    fontSize: 13.5,
    background: "transparent",
    color: "#101113",
  },
  sugestaoBtn: {
    display: "block",
    width: "100%",
    textAlign: "left",
    background: "#FCFBF8",
    border: "1px solid #E4E1D8",
    borderRadius: 8,
    padding: "8px 10px",
    fontSize: 12.5,
    marginTop: 6,
    cursor: "pointer",
  },
  servicoCard: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    width: "100%",
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: "12px 10px",
    marginBottom: 8,
    cursor: "pointer",
  },
  card: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 12,
    padding: 16,
  },
  sectionLabel: {
    fontSize: 11,
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: 0.6,
    color: "#8A8F98",
    marginBottom: 10,
  },
  grid2: { display: "flex", gap: 10, marginBottom: 2 },
  label: { fontSize: 11.5, color: "#5B6169", marginBottom: 4, fontWeight: 600 },
  labelSmall: { fontSize: 10, color: "#8A8F98", marginBottom: 3, fontWeight: 600 },
  input: {
    width: "100%",
    boxSizing: "border-box",
    padding: "9px 10px",
    borderRadius: 8,
    border: "1px solid #DEDBD1",
    fontSize: 13.5,
    background: "#FCFBF8",
    color: "#101113",
    outline: "none",
  },
  inputSmall: {
    width: "100%",
    boxSizing: "border-box",
    padding: "7px 8px",
    borderRadius: 7,
    border: "1px solid #DEDBD1",
    fontSize: 12.5,
    background: "#FCFBF8",
    color: "#101113",
    outline: "none",
  },
  itemsHintRow: {
    fontSize: 11.5,
    color: "#8A8F98",
    marginBottom: 10,
    lineHeight: 1.4,
  },
  itemCard: {
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: 10,
    marginBottom: 10,
    background: "#FCFBF8",
  },
  itemRowTop: { display: "flex", gap: 8, alignItems: "center" },
  itemRowBottom: { display: "flex", gap: 8, marginTop: 8 },
  itemFooter: {
    display: "flex",
    justifyContent: "space-between",
    fontSize: 11,
    color: "#5B6169",
    marginTop: 8,
    paddingTop: 8,
    borderTop: "1px dashed #E4E1D8",
  },
  custoRow: {
    display: "flex",
    gap: 6,
    alignItems: "center",
    marginBottom: 6,
  },
  addCustoBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    padding: "8px",
    borderRadius: 8,
    border: "1px dashed #C7C2B3",
    background: "transparent",
    color: "#5B6169",
    fontSize: 11.5,
    fontWeight: 600,
    cursor: "pointer",
    marginTop: 2,
  },
  iconBtn: {
    background: "transparent",
    border: "none",
    cursor: "pointer",
    padding: 6,
    display: "flex",
    alignItems: "center",
  },
  addItemBtn: {
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    width: "100%",
    padding: "10px",
    borderRadius: 8,
    border: "1px dashed #C7C2B3",
    background: "transparent",
    color: "#5B6169",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
    marginTop: 2,
  },
  summaryBar: {
    display: "flex",
    gap: 8,
    marginTop: 18,
    marginBottom: 14,
  },
  summaryStat: {
    flex: 1,
    background: "#101113",
    borderRadius: 10,
    padding: "10px 8px",
    textAlign: "center",
  },
  summaryLabel: { fontSize: 9.5, color: "#B9BCC2", textTransform: "uppercase", letterSpacing: 0.4 },
  summaryValue: { fontSize: 13.5, fontWeight: 700, marginTop: 2, color: "#F5C400" },
  primaryBtn: {
    width: "100%",
    padding: "13px",
    borderRadius: 10,
    border: "none",
    background: "#F5C400",
    color: "#101113",
    fontSize: 14,
    fontWeight: 700,
    cursor: "pointer",
  },
  formActions: {
    display: "flex",
    flexDirection: "column",
    gap: 8,
  },
  secondaryBtn: {
    width: "100%",
    padding: "12px",
    borderRadius: 10,
    border: "1px solid #DEDBD1",
    background: "#FFFFFF",
    color: "#5B6169",
    fontSize: 13,
    fontWeight: 600,
    cursor: "pointer",
  },
  primaryBtnSmall: {
    display: "flex",
    alignItems: "center",
    padding: "8px 14px",
    borderRadius: 8,
    border: "none",
    background: "#101113",
    color: "#F5C400",
    fontSize: 12.5,
    fontWeight: 700,
    cursor: "pointer",
  },
  ghostBtn: {
    display: "flex",
    alignItems: "center",
    padding: "8px 12px",
    borderRadius: 8,
    border: "1px solid #DEDBD1",
    background: "transparent",
    color: "#5B6169",
    fontSize: 12.5,
    fontWeight: 600,
    cursor: "pointer",
  },
  previaActions: {
    display: "flex",
    justifyContent: "space-between",
    marginBottom: 12,
  },
  copyBox: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: 12,
    marginBottom: 14,
  },
  copyTextarea: {
    width: "100%",
    minHeight: 160,
    boxSizing: "border-box",
    borderRadius: 8,
    border: "1px solid #DEDBD1",
    background: "#FCFBF8",
    padding: 10,
    fontSize: 12.5,
    fontFamily: "monospace",
    marginBottom: 10,
    resize: "vertical",
  },
  osPaper: {
    position: "relative",
    background: "#FFFFFF",
    borderRadius: 4,
    boxShadow: "0 2px 14px rgba(16,17,19,0.08)",
    padding: "0 18px 20px",
    overflow: "hidden",
  },
  osTopBar: { height: 6, background: "#F5C400", margin: "0 -18px 18px" },
  draftRibbon: {
    background: "#FBF0E4",
    color: "#B5651D",
    fontSize: 11,
    fontWeight: 700,
    textAlign: "center",
    padding: "8px 10px",
    margin: "0 -18px 0",
  },
  osHeaderRow: { display: "flex", justifyContent: "space-between", alignItems: "flex-start" },
  osCompany: { fontFamily: "Georgia, 'Times New Roman', serif", fontSize: 16, fontWeight: 700 },
  osCompanyDetail: { fontSize: 10, color: "#5B6169", marginTop: 1 },
  osLogo: {
    width: 64,
    height: 36,
    borderRadius: 4,
    background: "transparent",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    flexShrink: 0,
  },
  osLogoImg: {
    width: "100%",
    height: "100%",
    objectFit: "contain",
  },
  osTitle: {
    fontFamily: "Georgia, 'Times New Roman', serif",
    fontSize: 24,
    fontWeight: 700,
    margin: "18px 0 14px",
  },
  osMetaGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: "10px 14px",
    paddingBottom: 14,
    borderBottom: "1px solid #E4E1D8",
    marginBottom: 12,
  },
  metaField: {},
  metaLabel: { fontSize: 9.5, fontWeight: 700, color: "#8A8F98", textTransform: "uppercase" },
  metaValue: { fontSize: 13, marginTop: 2 },
  osTableHeader: {
    display: "flex",
    fontSize: 10.5,
    fontWeight: 700,
    color: "#101113",
    borderBottom: "1px solid #101113",
    paddingBottom: 6,
    marginBottom: 4,
  },
  osTableRow: {
    display: "flex",
    fontSize: 12,
    padding: "6px 0",
    borderBottom: "1px solid #F0EEE7",
  },
  osEmptyRow: { fontSize: 12, color: "#8A8F98", padding: "14px 0" },
  osTotalRow: {
    display: "flex",
    justifyContent: "space-between",
    fontSize: 15,
    fontWeight: 700,
    marginTop: 10,
    paddingTop: 10,
    borderTop: "2px solid #101113",
  },
  stamp: {
    position: "absolute",
    right: 24,
    bottom: 70,
    border: "3px solid #1F7A4D",
    color: "#1F7A4D",
    fontWeight: 800,
    fontSize: 20,
    padding: "4px 14px",
    borderRadius: 6,
    transform: "rotate(-12deg)",
    opacity: 0.75,
    letterSpacing: 2,
  },
  periodoRow: {
    display: "flex",
    alignItems: "center",
    gap: 8,
  },
  outroForm: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: 10,
    marginBottom: 10,
  },
  statsGrid: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr",
    gap: 8,
  },
  statsGrid3: {
    display: "grid",
    gridTemplateColumns: "1fr 1fr 1fr",
    gap: 8,
    marginTop: 10,
  },
  statCard: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: "12px 10px",
  },
  statLabel: { fontSize: 10, color: "#8A8F98", fontWeight: 600, textTransform: "uppercase" },
  statValue: { fontSize: 16, fontWeight: 700, marginTop: 3 },
  chartCard: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 12,
    padding: 14,
    marginTop: 16,
  },
  emptyMsg: { fontSize: 12.5, color: "#8A8F98", padding: "10px 0" },
  recordRow: {
    background: "#FFFFFF",
    border: "1px solid #E4E1D8",
    borderRadius: 10,
    padding: "10px 10px",
    marginBottom: 8,
  },
  recordTopRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
  },
  recordBadgeRow: {
    display: "flex",
    gap: 6,
    marginTop: 8,
  },
  recordTitle: { fontSize: 13, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  recordSub: { fontSize: 11, color: "#8A8F98", marginTop: 2 },
  statusBtn: {
    display: "flex",
    alignItems: "center",
    gap: 4,
    fontSize: 10.5,
    fontWeight: 700,
    padding: "6px 8px",
    borderRadius: 20,
    border: "none",
    cursor: "pointer",
    whiteSpace: "nowrap",
  },
};

const printStyles = `
  @media print {
    .no-print { display: none !important; }

    /* impede o wrapper de forçar uma "tela cheia" extra em branco */
    html, body { min-height: 0 !important; height: auto !important; }
    .app-shell { min-height: 0 !important; }

    /* deixa o papel da OS fluir normalmente, sem overlay/absolute */
    #os-paper {
      position: static !important;
      box-shadow: none !important;
      max-width: 700px;
      margin: 0 auto;
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
  }
`;
