"use client";

import { ChangeEvent, useMemo, useState } from "react";

type Origin = "Funcionarios" | "FORMACION" | "ELECTRONICOS";
type Equipment = {
  origin: Origin;
  componente: string;
  marca: string;
  modelo: string;
  serial: string;
  placa: string;
  disco: string;
  procesador: string;
  ram: string;
  cuentadante: string;
  sede: string;
  hostname: string;
  dominio: string;
};

const SOURCES: Origin[] = ["Funcionarios", "FORMACION", "ELECTRONICOS"];
const EMPTY_EQUIPMENT: Equipment = {
  origin: "Funcionarios", componente: "", marca: "", modelo: "", serial: "", placa: "",
  disco: "", procesador: "", ram: "", cuentadante: "", sede: "", hostname: "", dominio: ""
};

function normalize(value: unknown) {
  return String(value ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-zA-Z0-9]/g, "").toLowerCase();
}

function cellText(value: unknown): string {
  if (value instanceof Date) return value.toLocaleDateString("es-CO");
  if (typeof value === "object" && value && "text" in value) return String((value as { text: string }).text);
  return String(value ?? "").trim();
}

function download(content: BlobPart, name: string) {
  const url = URL.createObjectURL(new Blob([content], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const link = document.createElement("a");
  link.href = url;
  link.download = name;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 1000);
}

export default function Home() {
  const [inventoryFile, setInventoryFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [equipment, setEquipment] = useState<Equipment[]>([]);
  const [plateQuery, setPlateQuery] = useState("");
  const [source, setSource] = useState<"Todas" | Origin>("Todas");
  const [matches, setMatches] = useState<Equipment[]>([]);
  const [selected, setSelected] = useState<Equipment | null>(null);
  const [message, setMessage] = useState("Cargue el inventario y la plantilla para comenzar.");
  const [busy, setBusy] = useState(false);

  const filteredCount = useMemo(() => source === "Todas" ? equipment.length : equipment.filter(item => item.origin === source).length, [equipment, source]);

  async function readInventory(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0];
    if (!file) return;
    setBusy(true);
    try {
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await file.arrayBuffer());
      const loaded: Equipment[] = [];

      SOURCES.forEach((origin) => {
        const sheet = workbook.getWorksheet(origin);
        if (!sheet) return;
        const headers = sheet.getRow(4).values as unknown[];
        const columns = new Map<string, number>();
        headers.forEach((header, index) => columns.set(normalize(cellText(header)), index));
        const value = (row: import("exceljs").Row, header: string) => cellText(row.getCell(columns.get(normalize(header)) ?? 0).value);

        for (let rowNumber = 5; rowNumber <= sheet.rowCount; rowNumber += 1) {
          const row = sheet.getRow(rowNumber);
          const placa = value(row, "Placa");
          if (!placa || normalize(placa) === "na") continue;
          loaded.push({
            origin,
            componente: value(row, "Elemento"),
            marca: value(row, "Marca"),
            modelo: value(row, "Modelo"),
            serial: value(row, "Serial"),
            placa,
            disco: value(row, "Capacidad de Disco Duro"),
            procesador: value(row, "Tipo de Procesador"),
            ram: value(row, "Capacidad de Memoria"),
            cuentadante: value(row, "Cuentadante") || value(row, "Propietario"),
            sede: value(row, "Nombre de la Sede"),
            hostname: value(row, "Nombre del CI"),
            dominio: value(row, "Dominio")
          });
        }
      });
      setInventoryFile(file);
      setEquipment(loaded);
      setMatches([]);
      setSelected(null);
      setMessage(`Inventario cargado: ${loaded.length} equipos encontrados en las tres hojas.`);
    } catch {
      setMessage("No fue posible leer el inventario. Verifique que sea un archivo .xlsx válido.");
    } finally {
      setBusy(false);
    }
  }

  function selectTemplate(event: ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0] ?? null;
    setTemplateFile(file);
    if (file) setMessage("Plantilla cargada correctamente. Ya puede buscar un equipo.");
  }

  function search() {
    const query = normalize(plateQuery);
    if (!query) return setMessage("Ingrese una Placa SENA para realizar la búsqueda.");
    const found = equipment.filter(item => (source === "Todas" || item.origin === source) && normalize(item.placa).includes(query));
    setMatches(found);
    setSelected(found.length === 1 ? found[0] : null);
    setMessage(found.length === 0 ? "No se encontró ningún equipo con esa Placa SENA." : found.length === 1 ? "Equipo encontrado. Revise los datos antes de generar el archivo." : `Se encontraron ${found.length} registros. Seleccione el correcto.`);
  }

  function updateField(field: keyof Equipment, value: string) {
    setSelected(current => current ? { ...current, [field]: value } : current);
  }

  async function generate() {
    if (!templateFile || !selected) return setMessage("Seleccione una plantilla y un equipo antes de generar el archivo.");
    setBusy(true);
    try {
      const ExcelJS = (await import("exceljs")).default;
      const workbook = new ExcelJS.Workbook();
      await workbook.xlsx.load(await templateFile.arrayBuffer());
      const sheet = workbook.worksheets[0];
      if (!sheet) throw new Error("Plantilla sin hoja");

      // Estas celdas corresponden a los rótulos dejados en la plantilla GTI-F-132.
      const mapping: Record<string, string> = {
        A23: selected.componente, B23: selected.marca, D23: selected.modelo, E23: selected.serial, G23: selected.placa,
        C17: selected.cuentadante, C19: selected.sede, J12: selected.hostname, J13: selected.dominio,
        E35: selected.disco, E36: selected.procesador, E37: selected.ram
      };
      Object.entries(mapping).forEach(([cell, value]) => { sheet.getCell(cell).value = value || ""; });
      const content = await workbook.xlsx.writeBuffer();
      download(content, `HOJA_DE_VIDA_${selected.placa.replace(/[^a-zA-Z0-9_-]/g, "_")}.xlsx`);
      setMessage("Hoja de vida generada y descargada en formato Excel editable.");
    } catch {
      setMessage("No fue posible generar el Excel. Use la plantilla original GTI-F-132 en formato .xlsx.");
    } finally {
      setBusy(false);
    }
  }

  const fields: Array<[keyof Equipment, string]> = [
    ["componente", "Componente"], ["marca", "Marca"], ["modelo", "Modelo"], ["serial", "Serial"], ["placa", "Placa SENA"],
    ["disco", "Disco duro"], ["procesador", "Procesador"], ["ram", "Memoria RAM"], ["cuentadante", "Nombre de usuario / cuentadante"],
    ["sede", "Sede"], ["hostname", "Hostname"], ["dominio", "Dominio"]
  ];

  return (
    <main>
      <header className="hero"><div><p className="eyebrow">SENA · Gestión de Tecnologías de la Información</p><h1>Hoja de Vida Equipos SENA</h1><p>Genere hojas de vida desde el inventario, usando la Placa SENA.</p></div></header>
      <section className="steps" aria-label="Proceso"><span>1. Cargar archivos</span><span>2. Buscar equipo</span><span>3. Revisar datos</span><span>4. Descargar Excel</span></section>
      <section className="grid uploads">
        <label className="file-card"><strong>Inventario SENA 2026</strong><small>Lee Funcionarios, FORMACION y ELECTRONICOS.</small><input type="file" accept=".xlsx" onChange={readInventory} />{inventoryFile && <em>{inventoryFile.name}</em>}</label>
        <label className="file-card"><strong>Plantilla GTI-F-132</strong><small>Se conservará como base del Excel generado.</small><input type="file" accept=".xlsx" onChange={selectTemplate} />{templateFile && <em>{templateFile.name}</em>}</label>
      </section>
      <section className="panel search-panel"><div><h2>Buscar equipo</h2><p>Registros disponibles: {filteredCount}</p></div><div className="search-controls"><select value={source} onChange={event => setSource(event.target.value as "Todas" | Origin)}><option>Todas</option>{SOURCES.map(item => <option key={item}>{item}</option>)}</select><input value={plateQuery} onChange={event => setPlateQuery(event.target.value)} onKeyDown={event => event.key === "Enter" && search()} placeholder="Ingrese la Placa SENA" /><button onClick={search} disabled={busy || !equipment.length}>Buscar</button></div></section>
      <p className="message" role="status">{busy ? "Procesando archivo…" : message}</p>
      {matches.length > 1 && <section className="panel"><h2>Seleccione el equipo</h2><div className="results">{matches.map((item, index) => <button className="result" key={`${item.origin}-${item.placa}-${index}`} onClick={() => { setSelected(item); setMessage("Registro seleccionado. Puede editar sus datos."); }}><strong>{item.placa}</strong><span>{item.componente} · {item.marca} {item.modelo}</span><small>{item.origin} · {item.cuentadante || "Sin cuentadante"}</small></button>)}</div></section>}
      {selected && <section className="panel details"><div className="section-heading"><div><p className="eyebrow">Origen: {selected.origin}</p><h2>Información del equipo</h2></div><button className="primary" onClick={generate} disabled={busy || !templateFile}>Generar Excel</button></div><p className="hint">Puede ajustar cualquier dato antes de descargar la hoja de vida.</p><div className="form-grid">{fields.map(([field, label]) => <label key={field}>{label}<input value={selected[field]} onChange={event => updateField(field, event.target.value)} /></label>)}</div></section>}
      <footer>Los archivos se procesan en su navegador y no se almacenan permanentemente.</footer>
    </main>
  );
}
